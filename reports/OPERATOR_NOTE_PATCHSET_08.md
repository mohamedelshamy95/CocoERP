Patchset-08: GRN-based Receipts for Egypt
----------------------------------------

What changed
- Added feature flag `Enable Receipts_EG GRN Flow (v1)` (Settings) to gate GRN-based receiving.
- New sheet `Receipts_EG` with auto defaults (GRN ID, GRN Line ID, Receipt Date, Warehouse EG).
- New sync `syncReceiptsEgToInventory_EG` posts GRN-ledger (OUT from UAE KOR/ATTIA, IN to EG-TAN) and updates Shipments_UAE_EG received qty. Idempotent by GRN Line source IDs.
- Shipments_UAE_EG sync now supports OUT-only mode when GRN flag is ON (IN leg is disabled) and uses `Qty Out Synced` to track UAE outbound posting.
- Queue integration: new flag `CocoERP_ReceiptsEgInvSyncFlag_v1` auto-detected from Receipts_EG edits/pending lines.

Enablement / rollout steps
1) `npm run check:syntax`
2) `clasp push`
3) In Settings sheet, set `Enable Receipts_EG GRN Flow (v1)` = 1 (or TRUE).
4) Ensure Shipments_UAE_EG has a `Qty Out Synced` column (header exact). If missing, add the column before enabling.
5) Run `ensureReceiptsEgSheet_()` once (creates headers) then fill GRN lines (GRN ID/Line ID auto on edit).
6) Run queue or call `syncReceiptsEgToInventory_EG()`; reruns are idempotent.

Notes/tags
- Blocking tags on Receipts_EG Notes: `GRN_SYNC_SKIPPED_NO_ID_V1`, `_NO_DATE_V1`, `_NO_SKU_V1`, `_NO_QTY_V1`, `_BAD_WAREHOUSE_V1`, `_BAD_UAE_WAREHOUSE_V1`, `_NO_COST_V1`.
- Shipments_UAE_EG uses `SUEG_WAITING_GRN_V1` when GRN mode is ON and OUT posted but IN pending.

Manual test plan
- Flag OFF: Shipments_UAE_EG sync behaves as before (OUT+IN).
- Flag ON:
  - Add GRN rows (multiple SKUs/variants), edit → GRN ID/Line ID/Date auto-filled.
  - Run `syncReceiptsEgToInventory_EG()` twice; first posts OUT+IN per line, second posts 0.
  - Move/sort GRN rows and rerun → no duplicates.
  - Leave GRN ID or Receipt Date blank → row tagged/skipped; fix and rerun clears tag and posts once.
  - Shipments_UAE_EG OUT posts when Qty > Qty Out Synced + Ship Date + UAE warehouse (KOR/ATTIA), without requiring Actual Arrival.

Rollback
- Set `Enable Receipts_EG GRN Flow (v1)` back to 0/blank.
- Revert repo changes (git checkout .) and `clasp push`.
