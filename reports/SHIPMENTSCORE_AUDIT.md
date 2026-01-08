# ShipmentsCore.js — Full Static Audit (Read-only)

**File:** [ShipmentsCore.js](ShipmentsCore.js)  
**Audit type:** End-to-end static audit (no code changes)  
**Hard constraints enforced in this audit:**
- No header rename/reorder; schema changes append-only (prefer none).
- `onEdit` handlers must remain lightweight (row-local only; no full-sheet scans; no ledger writes).
- Queue jobs must be idempotent + lock-protected; ledger writes must use deterministic Txn IDs (via [`logInventoryTxnBatch_`](InventoryCore.js) / [`_inv_makeTxnId_`](InventoryCore.js)).
- Ledger-first: `Inventory_Transactions` is truth; snapshots are derived.
- Canonical EG warehouse must be **exactly** `EG-TAN` (normalize legacy aliases).
- Any ShipmentsCore change may impact AppCore dispatch/queue; such touchpoints are flagged explicitly.

> **Notation:** Some parts of the provided file content in chat are truncated with `…`.  
> Findings marked **Hypothesis** must be verified in the live file view around the quoted snippet before implementation.

---

## A) Executive Summary

### P0 (crash / deploy-blocker / data corruption / duplicate-ledger risk)
1. **SCS-001 (Confirmed):** Undefined variable `colArrival` in [`_updateShipmentUaeEgRowTotalsAndStatus_`](ShipmentsCore.js) → runtime crash on edit.
2. **SCS-002 (Hypothesis):** Possible **syntax/parse error** in [`updateShipmentsUaeEgStatusAndTotals`](ShipmentsCore.js) due to incomplete `if (...) else` line fragment shown in excerpt.
3. **SCS-003 (Hypothesis):** Multiple `if (cond)` lines with no visible `continue/return` (e.g., in [`_getInventoryUaeInfoForSku_`](ShipmentsCore.js), [`_sueg_buildUaeBasisFromLedger_`](ShipmentsCore.js)) may be truncated artifacts; if real, would cause severe correctness break.

### P1 (high-risk correctness / idempotency / wrong-warehouse / repeated queue churn)
4. **SCS-004 (Likely):** UAE→EG sync destination uses legacy EG warehouse (`TAN-GH` / `TAN`) instead of required `EG-TAN`.
5. **SCS-005 (Confirmed cross-module behavior):** Repeated queue churn: rows with `Qty > Qty Synced` but missing readiness fields get re-flagged every minute (ShipmentsCore readiness gate vs AppCore auto-detect). **Touches AppCore.**
6. **SCS-006 (Confirmed):** Planned seeding key in [`seedShipmentsUaeEgFromInventoryUae`](ShipmentsCore.js) uses `SKU||WH` only → duplicates across variants and persistent duplicates after rebuild/seed cycles.
7. **SCS-007 (Confirmed):** [`resolveUaeWarehouseFromCourier_`](ShipmentsCore.js) defaults to `UAE-DXB`, which violates the stated canonicalization requirement (“UAE must be KOR/ATTIA”).

### P2 (performance / write amplification / observability hygiene)
8. **SCS-008 (Confirmed):** UAE→EG sync reads full Shipments_UAE_EG + full Inventory_UAE + full ledger, then writes back entire Shipments range each run → slow at scale.
9. **SCS-009 (Confirmed):** Seeding updates rows with per-row `setValues` loop → avoidable slowness and longer lock contention.
10. **SCS-010 (Confirmed):** CN→UAE rebuild helper loops row-by-row, writing entire row each time (`getRange(...).setValues([row])`) → slow for large sheets (manual path).
11. **SCS-011 (Confirmed):** QC→Inventory sync scans entire ledger each run to compute “already-synced” sets; can become a hot path as ledger grows.
12. **SCS-012 (Confirmed):** ErrorLog rate limits exist for baseline mismatch, but readiness/missing-data logging policy is inconsistent; missing row-level tags cause repeated operator confusion/noise.

---

## B) Inventory of Entrypoints (AppCore / triggers / menus)

> AppCore owns the only global triggers; ShipmentsCore provides module handlers invoked by dispatch/queue.  
> Touchpoints are via [`_dispatchOnEdit_`](AppCore.js) and [`coco_processSyncQueue`](AppCore.js) (and menu actions).

### onEdit-dispatched (must remain lightweight)
1. **[`shipmentsCnUaeOnEdit_`](ShipmentsCore.js)**  
   - **Purpose:** Recompute status/total for the edited CN→UAE row.  
   - **Reads/Writes:** Reads header map + edited row; writes that row (status/total).  
   - **Perf risk:** Medium on large paste operations (builds header map each time; row-level writes).  
   - **Constraint fit:** OK (row-local; no ledger writes).

2. **[`shipmentsUaeEgOnEdit_`](ShipmentsCore.js)**  
   - **Purpose:** When SKU changes, auto-fill from Inventory_UAE; when cost/date changes, recalc total/status.  
   - **Reads/Writes:** Reads/writes a single row; may read Inventory_UAE sheet for lookup.  
   - **Perf risk:** Medium (Inventory scan in `_getInventoryUaeInfoForSku_` appears full-scan).  
   - **Constraint fit:** Should be OK if lookup is efficient; currently lookup is full sheet scan.

3. **[`qcOnEdit_`](ShipmentsCore.js)**  
   - **Purpose:** Recalculate computed QC columns for edited rows (Qty Missing/OK/Result).  
   - **Reads/Writes:** Batch window read/write for only edited rows.  
   - **Perf risk:** Low.  
   - **Constraint fit:** Good (row-window only; no UI; no ledger writes).

### Queue/manual sync writers (must be idempotent + lock-safe; ledger-first)
4. **[`syncPurchasesToShipmentsCnUae`](ShipmentsCore.js)**  
   - **Purpose:** Upsert Shipments_CN_UAE lines from Purchases (line-level keys).  
   - **Side effects:** Writes Shipments_CN_UAE rows (updates + appends), may trigger status/totals rebuild.  
   - **Perf risk:** Medium (scans Purchases and Shipments ranges; uses batching for qty updates).

5. **[`syncQCtoInventory_UAE`](ShipmentsCore.js)**  
   - **Purpose:** Post QC_UAE inventory receipts into ledger (`Inventory_Transactions`) as `QC_UAE` IN rows.  
   - **Side effects:** Writes ledger via [`logInventoryTxnBatch_`](InventoryCore.js); optionally triggers [`inv_rebuildAllSnapshots`](InventoryCore.js).  
   - **Perf risk:** High (full QC scan + full Purchases scan + full ledger scan each run).  
   - **Idempotency basis:** skips if QC row already represented in ledger via Source ID / legacy notes scanning.

6. **[`syncShipmentsUaeEgToInventory`](ShipmentsCore.js)**  
   - **Purpose:** Post UAE→EG shipments as ledger `OUT` (UAE) + `IN` (EG) with landed cost.  
   - **Side effects:** Writes ledger via [`logInventoryTxnBatch_`](InventoryCore.js), updates `Qty Synced` in Shipments_UAE_EG, rebuilds snapshots.  
   - **Perf risk:** Very High (full sheet + full ledger scans; full-range writeback).  
   - **Idempotency basis:** dedupe sets of Txn IDs and Source IDs; persists Line IDs before posting.

### Menu/manual utilities (non-trigger; can be heavier)
7. **[`updateShipmentsCnUaeStatusAndTotals`](ShipmentsCore.js)**  
8. **[`rebuildShipmentsCnUaeStatus_`](ShipmentsCore.js)**  
9. **[`setupShipmentsCnUaeStatusValidation_`](ShipmentsCore.js)**  
10. **[`updateShipmentsUaeEgStatusAndTotals`](ShipmentsCore.js)**  
11. **[`updateAllShipmentsStatusAndTotals`](ShipmentsCore.js)**  
12. **[`backfillShipmentsUaeEgFromInventory`](ShipmentsCore.js)**  
13. **[`seedShipmentsUaeEgFromInventoryUae`](ShipmentsCore.js)**  
14. **[`qc_generateFromPurchasesPrompt`](ShipmentsCore.js)** (UI)  
15. **[`qc_generateFromPurchases_`](ShipmentsCore.js)** (heavy; writes QC sheet)  
16. **[`qc_recalcQuantitiesAndResult`](ShipmentsCore.js)** (manual recalculation window)  
17. **[`debugTestInventoryLookup_`](ShipmentsCore.js)** (debug only)  
18. **[`migrateFixShipUaeEgInLandedCostV1`](ShipmentsCore.js)** (migration; lock-protected; ledger mutation)

**AppCore touchpoints that must be considered when patching ShipmentsCore:**
- [`_dispatchOnEdit_`](AppCore.js) (calls `shipmentsCnUaeOnEdit_`, `shipmentsUaeEgOnEdit_`, `qcOnEdit_`)
- [`coco_processSyncQueue`](AppCore.js) (calls `syncPurchasesToShipmentsCnUae`, `qc_generateFromPurchases_`, `syncQCtoInventory_UAE`, `syncShipmentsUaeEgToInventory`)
- Queue auto-detect: [`coco_hasPendingShipUaeEgInventorySync_`](AppCore.js) (re-flags when `Qty > Qty Synced`)

---

## C) Crashers & Correctness (undefined vars, header guards, wrong mapping, non-deterministic dates)

### SCS-001 (P0, Confirmed) — Undefined `colArrival` crashes UAE→EG row recalc
- **Location:** [`_updateShipmentUaeEgRowTotalsAndStatus_`](ShipmentsCore.js) ~ lines **1180–1295**
- **Evidence snippet:**
  - Declares: `const colArr = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] || ...`
  - Uses: `const arr = colArrival ? row[colArrival - 1] : '';` (**colArrival undefined**)
- **Symptom:** `ReferenceError: colArrival is not defined` when onEdit triggers recalc → row totals/status not updated.
- **Root cause:** Variable name mismatch (`colArr` declared; `colArrival` referenced).
- **Minimal fix intent:** Use the declared arrival column variable consistently (`colArr`) for read.  
  **Invariant:** onEdit row updater must never throw.
- **Touchpoints:** AppCore onEdit dispatch (**Yes**, via [`_dispatchOnEdit_`](AppCore.js)).
- **Minimal repro:** Edit “Actual Arrival (EG)” or any recalc column in `Shipments_UAE_EG`.

---

### SCS-002 (P0, Hypothesis) — Potential syntax error in UAE→EG sheet-level updater
- **Location:** [`updateShipmentsUaeEgStatusAndTotals`](ShipmentsCore.js) ~ lines **364–520**
- **Evidence snippet in provided content:**  
  `if (lastRow < 2) { if (interactive && typeof safeAlert_ === 'function')      else return; }`
- **Symptom if real:** Entire project fails to load / `npm run check:syntax` fails / deployments fail.
- **Root cause:** Incomplete `if (...) else` statement.
- **Minimal fix intent:** Ensure the empty-sheet branch has valid statements (alert/log then return).  
  **Invariant:** module parses under Apps Script V8 and Node `--check`.
- **Verify:** Open file around line ~420 and confirm whether this is real code or a summarization artifact.

---

### SCS-003 (P0, Hypothesis) — “Bare if” statements may invert logic or break loops
- **Location examples:**  
  - [`_getInventoryUaeInfoForSku_`](ShipmentsCore.js) ~ **820–900**  
    `if (!rowSku || rowSku !== normalizedSku)` followed by next statement.
  - [`_sueg_buildUaeBasisFromLedger_`](ShipmentsCore.js) ~ **2700–2850**  
    `if (!sku)` etc.
- **Symptom if real:** Inventory lookups return incorrect rows; basis computation corrupt; landed cost wrong; overship checks wrong.
- **Root cause:** Missing `continue/return` after guard conditions.
- **Minimal fix intent:** Ensure guards actually skip/continue rather than conditionally executing the next statement.  
  **Invariant:** non-matching rows never contribute to lookup/basis.
- **Verify:** Confirm if the actual file has `continue;` but the excerpt omitted it.

---

### SCS-004 (P1, Likely) — EG warehouse canonicalization violates “EG-TAN only”
- **Location:** [`syncShipmentsUaeEgToInventory`](ShipmentsCore.js) ~ **2288–2860**
- **Evidence:** File comment explicitly states EG IN posts “IN في TAN-GH” (legacy). Baseline helper + prior patch fragments also reference `TAN-GH`.
- **Symptom:** Ledger `IN` rows for SHIP_UAE_EG land in `TAN-GH`/`TAN` buckets → `Inventory_EG` snapshot split/wrong.
- **Root cause:** Legacy warehouse code used for EG destination and/or incomplete normalization contract.
- **Minimal fix intent:** Force all SHIP_UAE_EG `IN` postings to warehouse **exactly `EG-TAN`** after normalization.  
  **Invariant:** no ledger SHIP_UAE_EG IN row has any other EG warehouse.
- **Touchpoints:** AppCore queue runner (**Yes**, via [`coco_processSyncQueue`](AppCore.js)).

---

### SCS-005 (P1, Confirmed cross-module behavior) — Readiness gate vs auto-detect causes repeated churn
- **Location:**  
  - [`syncShipmentsUaeEgToInventory`](ShipmentsCore.js) readiness gate (inside main loop; truncated)  
  - [`coco_hasPendingShipUaeEgInventorySync_`](AppCore.js) flags when `Qty > Qty Synced` regardless of readiness
- **Symptom:** Every minute, incomplete rows (missing Ship Date/Arrival/Warehouse) are reprocessed, skipped again, and may spam ErrorLog.
- **Root cause:** Auto-detect uses delta only; sync requires readiness fields; no persistent “blocked” state.
- **Minimal fix intent:** Align detection with readiness OR tag rows to prevent repeated log/flagging.  
  **Invariant:** incomplete rows do not re-trigger heavy sync endlessly.
- **Touchpoints:** AppCore (**Yes**).

---

### SCS-006 (P1, Confirmed) — Planned seeding key ignores Variant (duplicates)
- **Location:** [`seedShipmentsUaeEgFromInventoryUae`](ShipmentsCore.js) ~ **1017–1180**
- **Evidence snippet:**  
  - Inventory key: `invByKey.set(\`\${sku}||\${wh}\`, ...)`
  - Planning index key: `existingPlanRowByKey.set(\`\${sku}||\${wh}\`, i)`
- **Symptom:** Same SKU with multiple variants collapses into one key or creates duplicates (depending on existing data); planned rows can accumulate duplicates across cycles.
- **Root cause:** Key is `SKU||WH` only; variant is not part of identity.
- **Minimal fix intent:** Include Variant (and any other discriminator used in layouts) in the planned-row key and dedupe policy.  
  **Invariant:** at most 1 planned row per `(SKU, Variant, Warehouse)` when Shipment ID blank.
- **Touchpoints:** AppCore indirectly (seed often run after rebuild workflows) (**Possible**; treat as **Yes** for planning).

---

### SCS-007 (P1, Confirmed) — Courier→warehouse resolver defaults to `UAE-DXB` (policy violation)
- **Location:** [`resolveUaeWarehouseFromCourier_`](ShipmentsCore.js) ~ **780–818**
- **Evidence snippet:**  
  `if (!s) return 'UAE-DXB';` and final `return 'UAE-DXB';`
- **Symptom:** Rows get `UAE-DXB` as source warehouse; if canonical UAE warehouses must be KOR/ATTIA only, this leaks a disallowed code into sheet and potentially ledger pricing/outflow.
- **Root cause:** Resolver is permissive and uses `UAE-DXB` as default bucket.
- **Minimal fix intent:** Default to blank/unknown (force operator) or map only to allowed canonical set; ensure normalization collapses to allowed set if policy requires.  
  **Invariant:** no non-approved UAE warehouse reaches ledger postings.

---

### SCS-008 (P2, Confirmed) — UAE→EG sync full-range writeback and full scans
- **Location:** [`syncShipmentsUaeEgToInventory`](ShipmentsCore.js) ~ **2288–2860**
- **Evidence snippet:**  
  - Full read: `shipSh.getRange(2, 1, lastShipRow - 1, shipSh.getLastColumn()).getValues();`
  - Full write: `shipSh.getRange(2, 1, shipData.length, shipSh.getLastColumn()).setValues(shipData);`
  - Full ledger scan for dedupe sets: `ledgerSh.getRange(2, 1, lr - 1, ledgerSh.getLastColumn()).getValues();`
- **Symptom:** Slow queue runs, lock contention, increased timeout risk; partial progress leads to repeated retries.
- **Root cause:** No windowing / no “write only changed rows” strategy.
- **Minimal fix intent:** Window processing and minimal writes of changed rows/columns only.  
  **Invariant:** One queue tick completes under time budget and remains idempotent.

---

### SCS-009 (P2, Confirmed) — Seeding updates row-by-row (write amplification)
- **Location:** [`seedShipmentsUaeEgFromInventoryUae`](ShipmentsCore.js) ~ **1120–1180**
- **Evidence snippet:**  
  `for (let k = 0; k < rowsToWrite.length; k++) { shipSh.getRange(rowNumbersToWrite[k], 1, 1, shipCols).setValues([rowsToWrite[k]]); }`
- **Symptom:** Slow seeding; more lock contention; higher chance of Apps Script execution limits.
- **Root cause:** No batching for updates to existing rows.
- **Minimal fix intent:** Batch contiguous row segments into fewer `setValues` calls.  
  **Invariant:** seeding is idempotent and efficient for large snapshots.

---

### SCS-010 (P2, Confirmed) — CN→UAE rebuild is per-row full-row write
- **Location:** [`rebuildShipmentsCnUaeStatus_`](ShipmentsCore.js) ~ **333–363** and helper [`_updateShipmentCnUaeStatusForRow_`](ShipmentsCore.js) ~ **168–248**
- **Symptom:** Very slow on large sheets; unnecessary range IO.
- **Root cause:** Per-row `getRange(...).getValues()` + `setValues([row])` in a loop.
- **Minimal fix intent:** Batch compute status/total for all rows and write only those columns, like sheet-level updater does.  
  **Invariant:** rebuild runs within manual execution budget.

---

### SCS-011 (P2, Confirmed) — QC→Inventory: full ledger scan for dedupe each run
- **Location:** [`syncQCtoInventory_UAE`](ShipmentsCore.js) ~ **1987–2287**
- **Evidence snippet:**  
  `ledgerSh.getRange(2, 1, ledgerLast - 1, ledgerSh.getLastColumn()).getValues();` then scan SourceType/SourceId/Notes
- **Symptom:** Queue tick time grows with ledger size.
- **Root cause:** Deduping is done by scanning the full ledger for QC rows on every run.
- **Minimal fix intent:** Use deterministic Source IDs and rely on ledger-side Txn-ID dedupe, plus track a pointer or cache of processed QC IDs if needed.  
  **Invariant:** Re-run posts 0 new ledger rows, and idle runs are fast.

---

### SCS-012 (P2, Confirmed) — Observability gaps: rate limit exists for baseline mismatch but not systematically for “missing readiness”
- **Location:** [`syncShipmentsUaeEgToInventory`](ShipmentsCore.js) end section shows baseline mismatch rate limit key  
  `CocoERP_RL_sueg_extrasBaselineMismatch_v1`
- **Symptom:** Operator-visible “stuck” rows without stable row-level tagging; ErrorLog can be noisy depending on current readiness logging implementation.
- **Root cause:** Logging policy is uneven (some rate-limited; others unclear/truncated); row-level notes/tags not consistently applied.
- **Minimal fix intent:** Standardize:
  - One row-level tag per missing prerequisite
  - One rate-limited ErrorLog entry per `(issueType, stableKey)` per window
  **Invariant:** no per-minute spam; operator can fix from the row itself.

---

## D) Idempotency & Duplication Risks

### Ledger posting duplication risks
- **QC (`syncQCtoInventory_UAE`):**
  - Uses `sourceId: qcId` and also scans ledger to detect synced QC rows.
  - **Risk:** If `qcId` is missing or unstable and `txnDate` defaults to `new Date()`, then deterministic Txn ID changes each run (depends on `_inv_makeTxnId_` inputs), potentially allowing duplicates.  
  - **Action:** Ensure source IDs are always stable (prefer QC ID), and `txnDate` is deterministic (prefer QC Date) when constructing Txn IDs.  
  - **Status:** Partially satisfied; depends on how `if (!qcId)` is implemented (truncated).

- **UAE→EG (`syncShipmentsUaeEgToInventory`):**
  - Good pattern: “Persist newly generated Line IDs BEFORE writing ledger rows.”
  - Dedupe: builds `existingSourceIds` and `existingTxnIds` from ledger for `SourceType == SHIP_UAE_EG`.
  - **Risks:**
    1. Any use of row-number-based discriminator in Source ID (not shown; baseline helper expects `SUEG|...` stable keys) would break idempotency when rows move.
    2. If txnDate uses `new Date()` fallback anywhere in txnId path, Txn IDs may change across reruns.
  - **Status:** Likely good, but verify main loop construction (truncated).

### Seeding duplication causes
- Key currently `SKU||WH` only (confirmed). If Variant exists, duplicates persist or collide (P1).
- No automatic enforcement of “one planned row per key” beyond “first wins” map fill.

---

## E) Performance & Write Amplification (queue paths)

### Hot paths
- [`syncShipmentsUaeEgToInventory`](ShipmentsCore.js): full read Shipments + full read Inventory_UAE + full read ledger + full write Shipments.
- [`syncQCtoInventory_UAE`](ShipmentsCore.js): full read QC + Purchases + ledger.
- [`seedShipmentsUaeEgFromInventoryUae`](ShipmentsCore.js): row-by-row update loop.

### Deterministic batching/windowing strategy (plan-level, no code here)
- **Window shipments rows per queue tick:** process max N rows (e.g., 200–500) starting from a pointer stored in DocumentProperties; pointer is a scan cursor only (not identity).
- **Write-minimization:** update only columns that changed (`Qty Synced`, `Line ID`, optional autofills), grouped into contiguous ranges.
- **Ledger reads:** avoid multiple ledger scans per run; derive basis + baseline + dedupe sets from a single `getValues()` read if possible.

---

## F) Observability & Noise Policy (ErrorLog + row tags)

### Current state
- Some rate limiting exists (baseline mismatch key in UAE→EG sync).
- QC missing-warehouse handling uses per-run arrays; whether it tags rows and rate-limits per row is unclear due to truncation.

### Recommended deterministic conventions (no code here)
**Rate-limit key naming (DocumentProperties):**
- `CocoERP_RL_shipUaeEg_ready_missing_v1::<shipmentId>|<stableLineKey>`
- `CocoERP_RL_shipUaeEg_wh_missing_v1::<shipmentId>|<sku>|<whHint>`
- `CocoERP_RL_qc_missing_wh_v1::<qcId>`
- `CocoERP_RL_qc_missing_id_v1::<orderId>|<sku>|<purchLineId>`
- Keep windows explicit (e.g., 6h) and version suffix `_v1` to allow controlled migrations.

**Row-level tag policy (Notes field only; idempotent):**
- Add tokens in a consistent delimiter (e.g., `| TAG |`):
  - `Missing Arrival (EG) - sync skipped`
  - `Missing Ship Date - sync skipped`
  - `Missing Warehouse (UAE) - sync skipped`
- Tag must be token-safe (no duplication); remove token when the missing field is resolved.

---

## Appendix: Issue Index (IDs)
- P0: SCS-001, SCS-002, SCS-003
- P1: SCS-004, SCS-005, SCS-006, SCS-007
- P2: SCS-008, SCS-009, SCS-010, SCS-011, SCS-012
