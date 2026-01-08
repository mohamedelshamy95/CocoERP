# Patchset-06: QC determinism + onEdit fail-safe

Summary:
- QC onEdit now exits silently when headers are missing/renamed; it no longer throws or blocks triggers.
- `qc_recalcRows_` supports a non-throw mode for triggers; manual/admin flows remain strict by default.
- `syncQCtoInventory_UAE` uses only stable identities (QC ID preferred; else Purchases Line ID via `QCL|<line>|<sku>|<variant>`) and a deterministic QC Date; rows missing required fields are skipped and tagged. No row-number IDs and no `new Date()` fallbacks.
- Token-safe Notes tags used on skips:
  - Missing ID: `QC_SYNC_SKIPPED_NO_ID_V1`
  - Missing QC Date: `QC_SYNC_SKIPPED_NO_DATE_V1`
  - Missing/invalid Warehouse (UAE): `QC_SYNC_SKIPPED_NO_WAREHOUSE_V1`

Operator steps after deploy:
1) `clasp push`
2) In `QC_UAE`, fill missing `QC ID`, `QC Date`, and canonical Warehouse (UAE) (KOR/ATTIA) for any rows tagged with:
   - `QC_SYNC_SKIPPED_NO_ID_V1`
   - `QC_SYNC_SKIPPED_NO_DATE_V1`
   - `QC_SYNC_SKIPPED_NO_WAREHOUSE_V1`
3) Run `syncQCtoInventory_UAE()` once (or let the queue trigger run).
4) Verify in `Inventory_Transactions` that QC posts appear once; reruns produce 0 new rows.

What to expect:
- Editing QC rows when headers are incomplete is a no-op (no errors).
- Rows lacking stable ID, QC Date, or canonical Warehouse are skipped and tagged; once fixed, tags are cleared and posting proceeds.
- Sorting or moving QC rows will not create duplicate ledger entries.

Rollback:
- Revert `ShipmentsCore.js`, this note, and the patch file `patches/patchset-06-qc-determinism.patch`, then `clasp push`.
- Ledger entries already written remain; if necessary, restore sheets from history and rerun sync after rollback.
