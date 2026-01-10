Patchset-11: Receipts_EG schema auto-ensure (self-healing headers)
-------------------------------------------------------------------

What changed
- New helper `receiptsEgEnsureSchema_()` now auto-creates `Receipts_EG` and enforces headers safely:
  - If row1 is empty → writes headers.
  - If row1 has data (no headers) → inserts a new header row above existing data.
  - If some headers are missing → appends only the missing ones at the end (no reordering, no data edits).
- Helper runs at the start of Receipts_EG onEdit and GRN sync, non-throwing, idempotent.
- Light formatting: freeze row1, bold, wrap, filter on header, date format on GRN Date, number format on Qty columns, data validation for Warehouse (UAE) to KOR/ATTIA.

Headers ensured (in order)
- GRN ID, GRN Line ID, GRN Date, Warehouse (UAE), SKU, Product Name, Variant / Color, Qty Received, Qty Synced, Notes, Shipment ID, Shipment Line ID, Sync Status, Last Synced At, Posted Txn ID.

How to verify
- If Receipts_EG row1 is blank: run `syncReceiptsEgToInventory_EG()` or edit any cell → headers appear; data remains intact.
- If row1 contains data (no headers): rerun → a header row is inserted above; data shifts down intact.
- If a single header (e.g., Qty Synced) is missing: rerun → header is appended at the end only; order of existing columns preserved.
- Reruns do not duplicate headers.

Rollback
- Revert this patchset and `clasp push`. (Headers already written are harmless.)
