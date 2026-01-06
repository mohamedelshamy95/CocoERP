# CocoERP Repo Audit (Apps Script)

## Executive Summary
- P0: 1 (cost column contract broken in Shipments_UAE_EG → wrong landed costs/ledger values).
- P1: 3 (missing contract/runbook docs; UAE warehouse normalization collapses UAE-DXB into KOR; queue auto-detect loops can spam errors when rows are incomplete).
- P2: 4 (ledger write scans entire Txn ID column each call; auto-detect scans Shipments_UAE_EG every minute; non-deterministic line IDs depend on successful writes; ErrorLog has no rate-limit while per-row loops log repeatedly).

## Contract Mismatches (file/function/line)
- `ShipmentsCore.js:updateShipmentsUaeEgStatusAndTotals` (439-445) + `_updateShipmentUaeEgRowTotalsAndStatus_` (1221-1229) + `syncShipmentsUaeEgToInventory` (2313-2327, 2574-2586): Customs/Other headers are defined as per-unit (`APP.COLS.SHIP_UAE_EG.*` and `Logistics.js` headers), but code treats them as shipment totals and divides by Qty when computing extras, yielding understated landed costs and incorrect ledger pricing.
- Contract docs absent: `SYSTEM_CONTRACT_PACK.md` and `RUNBOOK.md` are not in the repo; only `PDF/CocoERP_Roadmap.pdf` exists, so required contract precedence cannot be satisfied.

## Idempotency Risks (scenarios)
- Per-unit/total mismatch above means re-running status/totals or inventory sync recomputes a different landed cost each time depending on how the user entered Customs/Other (per-unit vs total), so ledger entries diverge on retries.
- Non-deterministic IDs: Purchases line IDs (`Purchases.js:92-123`) and Shipments_UAE_EG line IDs when Box ID is blank (`ShipmentsCore.js:2475-2489`) use `Utilities.getUuid()`. If a write fails part-way, a rerun generates different IDs, changing dedupe keys for downstream syncs.
- Ledger Txn IDs rely on `new Date()` when txnDate is omitted (`InventoryCore.js:414`), so any caller that forgets a deterministic date will generate a new Txn ID on every retry.

## Concurrency & Trigger Safety
- Queue auto-detect scans full Shipments_UAE_EG every minute when the flag is clear (`AppCore.js:1629-1666`). Missing Ship Date/Arrival/warehouse causes `syncShipmentsUaeEgToInventory` to log an error and skip the row (`ShipmentsCore.js:2458-2467, 2526-2533`), but auto-detect re-flags it on the next run, creating an infinite loop and ErrorLog noise.
- onEdit handlers remain light, but `shipmentsUaeEgOnEdit_` recomputes totals/status on every relevant edit without throttling; acceptable but note for large pastes.

## Ledger Integrity
- Ledger writes use deterministic Txn IDs (`_inv_makeTxnId_`) and additional Source ID dedupe in Shipments_UAE_EG sync, which is good. However, per-unit Customs/Other mismatch feeds incorrect `unitCostEgp` into inventory ledger (see Contract Mismatch above).
- logInventoryTxnBatch_ always reads the entire Txn ID column to dedupe (`InventoryCore.js:474-505`), even when callers already built dedupe sets; this increases timeout risk on large ledgers, jeopardizing one-minute trigger reliability.

## Normalization Issues (SKU/Warehouse/Status)
- Warehouse normalization collapses `UAE`, `UAE-DXB`, and `DUBAI` to `KOR` (`AppCore.js:584-598`), which can misroute stock/ledger entries away from Dubai rows if the sheet still uses those codes.
- Cost headers (`APP.COLS.SHIP_UAE_EG`) expect per-unit values, but normalization aliases in `APP.HEADER_ALIASES` map legacy totals to the same header; combined with current math, this silently undercharges customs/other on multi-qty rows.
- Status enums consistent across CN→UAE and UAE→EG; no drift observed.

## Performance Risks (hot paths)
- Ledger dedupe scan of all Txn IDs per call (`InventoryCore.js:474-505`) will grow O(N) and is invoked by QC/Shipments sync; could exceed the 1-minute trigger as ledger scales.
- `coco_hasPendingShipUaeEgInventorySync_` scans the entire Shipments_UAE_EG sheet whenever the flag is not already set (`AppCore.js:1629-1666`), even if there are no deltas; for large sheets this is a hot path on every queue tick.
- Shipments_UAE_EG sync writes back the full data range every run (`ShipmentsCore.js:2781-2782`), not just touched rows, which increases write volume under lock.

## ErrorLog Hygiene
- `logError_` appends immediately with no rate limit (`AppCore.js:519-536`). Combined with the auto-detect loop in Shipments_UAE_EG sync, the ErrorLog can be spammed every minute for the same missing Ship Date/Arrival/Warehouse condition. Only a few cases are rate-limited (e.g., missing QC warehouse, extras baseline), others are not.

## Test Gaps + Proposed Test Harness Functions (Apps Script)
- Add a contract test that seeds a mock Shipments_UAE_EG row with Qty>1 and per-unit Customs/Other and asserts `updateShipmentsUaeEgStatusAndTotals` writes `(ship+customs+other)*qty` and that `syncShipmentsUaeEgToInventory` posts landed cost matching that formula.
- Add a test to validate `normalizeWarehouseCode_` mapping: ensure `UAE-DXB` stays distinct from `KOR` or explicitly document the collapse.
- Add a ledger write test that calls `logInventoryTxnBatch_` twice with the same payload and a fixed `txnDate`, confirming no duplicate rows are added and runtime stays under a budget with a synthetic large Txn ID set.
- Add a queue test harness that populates Shipments_UAE_EG with a missing Arrival date and asserts the queue sets `SHIP_UAE_EG_INV_SYNC_FLAG` once and does not re-log identical errors on subsequent ticks (requires rate-limit/change in code).

## Recommended Fix Plan (ordered patchsets; minimal & safe)
1) **Cost contract patch** (see `patches/audit_patchset_suggestions.patch`): treat Customs/Other as per-unit, recompute totals as `(ship + customs + other) * qty`, and derive extrasPerUnit without dividing totals by qty to align ledger costs with headers. Regenerate totals/status before any inventory sync.
2) **Guardrails**: add rate-limit or one-time error markers for Shipments_UAE_EG rows missing Ship Date/Arrival/Warehouse to prevent auto-detect loops from spamming ErrorLog; optionally throttle the auto-detect scan with a timestamp property.
3) **Performance**: allow `logInventoryTxnBatch_` to accept precomputed Txn ID sets (from callers like Shipments sync) to skip the full-column scan when provided; consider batching writes smaller when ledger grows.
4) **Normalization**: revisit `normalizeWarehouseCode_` mapping for `UAE/UAE-DXB` to avoid forcing everything to `KOR`, or document the policy and add sheet-level validation to keep canonical codes.
