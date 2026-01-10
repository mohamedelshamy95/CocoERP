Patchset-09: No-throw notReady + queue churn elimination
---------------------------------------------------------

Summary
- Shipments_UAE_EG inventory sync now skips not-ready rows (missing Ship Date / Actual Arrival when required / UAE warehouse) without throwing. Skips add token-safe tags and return stats; queue no longer logs errors for expected waiting rows.
- Readiness result object returned for observability; logging is rate-limited and informational (no ErrorLog spam).
- Tags cleared automatically when fields are fixed; GRN waiting tag remains informational.

Tags to watch (Notes, pipe-delimited)
- `SUEG_BLOCKED_NO_SHIPDATE_V1`
- `SUEG_BLOCKED_NO_ARRIVAL_V1` (only when arrival is required)
- `SUEG_BLOCKED_NO_UAE_WAREHOUSE_V1`
- `SUEG_WAITING_GRN_V1` (informational when GRN mode is ON)

Runbook
1) `npm run check:syntax`
2) `clasp push`
3) Run `syncShipmentsUaeEgToInventory()` (or queue). Not-ready rows are tagged/skipped; ready rows post once. Returned object includes postedOut/postedIn/skippedNotReady.
4) Fix tagged fields (Ship Date / Actual Arrival / Warehouse), rerun; tags clear and posting occurs once; rerun posts 0.

Rollback
- Revert this patch and `clasp push`.
- No data migration required; tags are harmless.
