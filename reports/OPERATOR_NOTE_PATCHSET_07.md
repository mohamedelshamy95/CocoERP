# Patchset-07: Shipments_UAE_EG hardening (warehouse + readiness + extras)

Summary (behavior changes):
- UAE→EG readiness is strict: a row is ready only when `Qty > Qty Synced`, Ship Date exists, Actual Arrival exists, and Warehouse (UAE) resolves strictly to `KOR` or `ATTIA`.
- syncShipmentsUaeEgToInventory uses strict warehouse resolution; OUT posts only from `KOR`/`ATTIA`, IN posts to `EG-TAN` exactly. Blocked rows do not post and are tagged in Notes (token-safe).
- Blocked tags (idempotent, pipe-delimited):
  - Missing Ship Date: `SUEG_BLOCKED_NO_SHIPDATE_V1`
  - Missing Arrival: `SUEG_BLOCKED_NO_ARRIVAL_V1`
  - Missing/invalid Warehouse (UAE): `SUEG_BLOCKED_NO_UAE_WAREHOUSE_V1`
  Tags are removed automatically once the prerequisite is fixed.
- Extras baseline builder now skips legacy/blank source IDs instead of aborting; scan completes for all rows.
- Manual updater tolerates missing optional cost columns (status still updates); logs a rate-limited warning once per day if totals headers are missing.
- AppCore pending detector aligns with the strict readiness predicate (no queue churn on blocked rows).

Operator steps after deploy:
1) `npm run check:syntax` (done locally) and `clasp push`.
2) In `Shipments_UAE_EG`, fix any rows tagged with the blocked tags:
   - Fill Ship Date / Actual Arrival.
   - Set Warehouse (UAE) to `KOR` or `ATTIA` (no UAE-DXB/DUBAI/blank).
3) Run the queue once (`coco_processSyncQueue()`) or let the scheduled trigger run.
4) Verify:
   - `Inventory_Transactions`: UAE→EG OUT warehouses are `KOR`/`ATTIA`; IN warehouses are `EG-TAN`.
   - `Inventory_EG` grouping shows only `EG-TAN` for new IN rows.

Rollback:
- Revert `ShipmentsCore.js`, `AppCore.js`, this note, and the patch file `patches/patchset-07-ship-uae-eg-hardening.patch`, then `clasp push`.
- Tags in Notes are harmless; removing them is optional after rollback.
