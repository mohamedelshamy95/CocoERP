Patchset-10: GRN-only Inventory Posting
---------------------------------------

Summary
- When `Enable Receipts_EG GRN Flow (v1)` is ON, Shipments_UAE_EG never writes ledger rows (OUT/IN). Rows are optionally tagged `WAITING_FOR_GRN_V1` if Qty > Qty Out Synced.
- Receipts_EG (GRN) is the sole ledger source: each GRN line posts delta = Qty Received - Qty Synced as OUT (UAE KOR/ATTIA) + IN (EG-TAN). Delta < 0 is skipped and tagged `GRN_NEEDS_REVERSAL_V1`.

Skip/validation tags (pipe-delimited)
- `GRN_SYNC_SKIPPED_NO_ID_V1`, `GRN_SYNC_SKIPPED_NO_LINEID_V1`, `GRN_SYNC_SKIPPED_NO_DATE_V1`, `GRN_SYNC_SKIPPED_NO_UAE_WAREHOUSE_V1`, `GRN_SYNC_SKIPPED_NO_SKU_V1`, `GRN_SYNC_SKIPPED_NO_QTY_V1`, `GRN_NEEDS_REVERSAL_V1`.
- Tags clear automatically when fields are fixed.

Rollout steps
1) Ensure Receipts_EG has headers: GRN ID, GRN Line ID, Receipt Date, Warehouse (EG), Warehouse (UAE), SKU, Variant / Color, Qty Received, Qty Synced (optional), Notes.
2) Set Settings → `Enable Receipts_EG GRN Flow (v1)` = 1.
3) `npm run check:syntax` then `clasp push`.
4) Enter GRN lines; on edit GRN Line ID should be populated once (UUID). Run queue or `syncReceiptsEgToInventory_EG()`; rerun posts 0.
5) Verify Shipments_UAE_EG ledger posting is disabled (no new OUT/IN) and rows with remaining Qty show `WAITING_FOR_GRN_V1`.

Verification checklist (multi-batch example 10/17/13/6/4 across KOR/ATTIA)
- Multiple GRNs covering the batches post once each; reruns post 0; sorting GRN rows causes 0 duplicates.
- GRN with Qty Received < Qty Synced tags `GRN_NEEDS_REVERSAL_V1` and posts nothing.
- Missing required fields tag appropriately and skip; fixing fields clears tags and posts once.
- Shipments_UAE_EG rows do not generate ledger rows while GRN mode is ON.

Rollback
- Set the flag to 0/blank and `clasp push`, or revert this patchset. (Shipments legacy behavior resumes when flag is OFF.)
