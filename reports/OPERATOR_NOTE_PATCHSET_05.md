# Patchset-05 – ShipmentsCore hardening (P0/P1 fixes)

What changed:
- Fixed `colArrival` runtime error in `_updateShipmentUaeEgRowTotalsAndStatus_` by using a single `colActualArrival` alias block.
- Shipments_UAE_EG readiness: queue pending detector now requires `Qty > Qty Synced` **and** Ship Date + Actual Arrival; sync job skips non-ready rows with a single rate-limited log.
- QC_UAE sync: warehouse resolution prefers QC_UAE, falls back to Purchases line warehouse; only canonical ATTIA/KOR are accepted; missing-warehouse note tags are added/removed token-safely.
- Courier resolver no longer defaults to UAE-DXB; unknown couriers leave warehouse blank for safer downstream logic.
- Seeded planning rows: dedupe key now includes Warehouse + SKU + Variant; new helper `shipUaeEg_dedupeSeededPlannedRowsOnce_()` removes duplicate seeded planned rows.
- EG landing uses canonical `EG-TAN` through normalization; no change to ledger keys beyond normalization.

Runbook (once after deploy):
1) `clasp push`
2) Run `shipUaeEg_dedupeSeededPlannedRowsOnce_()` (manual) to drop duplicate seeded planning rows.
3) Run `inv_rebuildAllSnapshots()` to refresh Inventory_UAE/EG snapshots.
4) Rerun the queue dispatcher (e.g., `coco_processSyncQueue()` or the scheduled trigger) to process ready Shipments_UAE_EG rows.

Expected results:
- No more pending-queue churn until Ship Date and Actual Arrival are present.
- QC_UAE rows with valid warehouses post to the ledger; missing-warehouse notes stay clean (no duplicated tags).
- Shipments_UAE_EG sync skips non-ready rows without noisy ErrorLog spam; ready rows still enforce deterministic ledger posting.
- Seeded planning rows no longer accumulate duplicates on repeated seeds/rebuilds.

Rollback:
- Revert the patchset (`git checkout -- ShipmentsCore.js AppCore.js reports/OPERATOR_NOTE_PATCHSET_05.md patches/patchset-05-shipmentscore-hardening.patch` or `git restore ...`), then rerun `clasp push`.
- If the dedupe helper was executed, rerun the relevant seeds from Inventory_UAE or restore from a sheet backup if needed, then rebuild snapshots.
