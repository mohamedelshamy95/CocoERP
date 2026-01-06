# Patchset-04: EG-TAN Canonicalization + SHIP_UAE_EG Landed Cost Repair + QC Note Hygiene

## What changed
- Canonical EG warehouse is now `EG-TAN`; aliases like `TAN-GH` normalize to `EG-TAN` (snapshots group under the canonical code).
- Added migration helper `migrateFixShipUaeEgInLandedCostV1()` to fix legacy SHIP_UAE_EG IN ledger rows that missed per-unit extras (ship/customs/other). Updates IN Unit Cost/Total Cost and tags rows with `MIG04_FIX_SUEG_IN_COST_V1` for idempotency.
- QC_UAE sync now clears the stale `Missing Warehouse (UAE) - sync skipped` note tag when a valid warehouse is present.

## How to run
1) Deploy: `npm run check:syntax` (optional) → `npm run lint` (optional) → `clasp push`.
2) Run migration (Apps Script): `migrateFixShipUaeEgInLandedCostV1()`. It is lock-protected and safe to re-run; updated rows are tagged.
3) Rebuild snapshots (optional but recommended): `inv_rebuildAllSnapshots()` to ensure EG snapshot uses `EG-TAN` grouping.
4) QC note hygiene: run `syncQCtoInventory_UAE()`; rows with valid Warehouse (UAE) will have the stale missing-warehouse tag removed automatically.

## Expected results
- Inventory_Transactions IN rows for SHIP_UAE_EG have Unit Cost = OUT base cost + per-unit extras from the shipment sheet; Total Cost recalculated; notes include `MIG04_FIX_SUEG_IN_COST_V1` on touched rows.
- Inventory_EG snapshot lists EG stock under `EG-TAN` (no `TAN-GH` buckets).
- QC_UAE rows with a valid Warehouse no longer show the `Missing Warehouse (UAE) - sync skipped` note.

## Rollback
- Revert the patchset files (AppCore.js, InventoryCore.js, ShipmentsCore.js) and delete this operator note; redeploy via `clasp push`.
- Migration is tag-based; rerunning the old code leaves existing ledger data as-is. No destructive data deletion is performed.
