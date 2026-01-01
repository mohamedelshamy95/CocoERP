# CONSISTENCY AUDIT (AppCore → Sales)

## Commands Executed
- `npm run check:syntax` ✅
- `npm run lint` ✅ (87 warnings: unused globals across modules)
- `git status -sb` ➜ dirty tree (AppCore.js, ShipmentsCore.js, .vscode/settings.json, .github/)

## Symbol Inventories
### DefinedSymbols (name → file:line)
```text
L: Purchases.js:655
ORDER_SCALAR: Purchases.js:662
R: Purchases.js:656
_applyDateFormatByHeaders_: AppCore.js:1357
_applyDateFormat_: InventoryCore.js:311
_applyDecimalFormatByHeaders_: AppCore.js:1363
_applyDecimalFormat_: InventoryCore.js:294
_applyIntFormatByHeaders_: AppCore.js:1360
_applyIntFormat_: InventoryCore.js:277
_applyNumberFormatByHeaders_: AppCore.js:1340
_canon_: AppCore.js:456
_coco_bootstrapLayoutsIfMissingHeaders_: AppCore.js:964
_coco_deleteTriggersNoLock_: AppCore.js:1246
_dispatchOnEdit_: AppCore.js:1087
_fallbackGetOrCreateSheet_: Logistics.js:211
_fillShipmentUaeEgFromInventory_: ShipmentsCore.js:897, ShipmentsCore.js:1368
_getInventoryUaeInfoForSku_: ShipmentsCore.js:815, ShipmentsCore.js:1286
_installPurchaseFormulasCore: Purchases.js:602
_inv_isEgWarehouse_: InventoryCore.js:584
_inv_isUaeWarehouse_: InventoryCore.js:578
_inv_isWhInList_: InventoryCore.js:569
_inv_makeTxnId_: InventoryCore.js:533
_inv_normWh_: InventoryCore.js:565
_setUseInstallableOnEditFlag_: AppCore.js:1081
_setupSheetWithHeaders_: Logistics.js:227
_sheetHeaderEmpty_: AppCore.js:990
_updateShipmentCnUaeStatusForRow_: ShipmentsCore.js:168
_useInstallableOnEditFlag_: AppCore.js:1072
addIfMissing_: AppCore.js:720
addNoteUnique_: Orders.js:186, Orders.js:370
applyListValidation: CatalogCore.js:107
applyPurchasesFormats_: Purchases.js:520
applyPurchasesValidations_: Purchases.js:548
assertRequiredColumns_: AppCore.js:419
backfillPurchasesSku: SkuUtils.js:358
backfillShipmentsUaeEgFromInventory: ShipmentsCore.js:987, ShipmentsCore.js:1458
buildRow_: InventoryCore.js:410
catalogLookupBySku_: Sales.js:527
catalog_applyDataValidation_: CatalogCore.js:100
catalog_syncFromInventoryEg: CatalogCore.js:147
clearInventoryTransactions_: AppCore.js:872
clearSettingsCache_: AppCore.js:762
coco_debugOrdersSyncStatus: AppCore.js:1864
coco_enqueueOrdersSyncFromPurchasesEdit_: AppCore.js:1407
coco_enqueueOrdersSync_: AppCore.js:1368
coco_enqueueQcGenFromPurchasesEdit_: AppCore.js:1476
coco_enqueueQcGenFromShipmentsCnUaeEdit_: AppCore.js:1163
coco_enqueueQcGen_: AppCore.js:1442
coco_flagQcInventorySyncFromQcEdit_: AppCore.js:1562
coco_flagShipUaeEgInventorySyncFromEdit_: AppCore.js:1592
coco_flagShipmentsCnUaeSyncFromPurchasesEdit_: AppCore.js:1527
coco_hasPendingShipUaeEgInventorySync_: AppCore.js:1629
coco_installTriggers: AppCore.js:1257
coco_onEditInstallable: AppCore.js:1233
coco_preflightAndRepair: AppCore.js:906
coco_processSyncQueue: AppCore.js:1661
coco_processSyncQueueNow: AppCore.js:1850
coco_uninstallTriggers: AppCore.js:1286
colLetter_: AppCore.js:400
debug_listTriggers: AppCore.js:1899
deepFreeze_: AppCore.js:1326
dvList_: Purchases.js:562
ensureErrorLog_: AppCore.js:511
ensureInventoryTxnHeader_: AppCore.js:856
ensureSettingsSheet_: AppCore.js:687
ensureSheetSchema_: AppCore.js:428
ensureSheet_: AppCore.js:380
getDefaultCurrency_: AppCore.js:806
getDefaultCustomsPct_: AppCore.js:826
getDefaultFxAedEgp_: AppCore.js:800
getDefaultFxRate_: AppCore.js:812
getDefaultShipUaeEgPerOrder_: AppCore.js:820
getHeaderMap_: AppCore.js:387
getSetting_: AppCore.js:794
getSettingsListByHeader_: AppCore.js:832
getSettingsMap_: AppCore.js:768
getSheet_: AppCore.js:373
getSpreadsheet_: AppCore.js:363
getUiIfAllowed_: AppCore.js:628
idx: Orders.js:142, Orders.js:329
installPurchasesFormulas: Purchases.js:597
inv_backfillMissingTxnIds: InventoryCore.js:1226
inv_fullRebuildFromLogistics: AppCore.js:882
inv_rebuildAllSnapshots: InventoryCore.js:1128
inv_repairInventoryTransactionsHeaders: InventoryCore.js:136
inv_repairInventoryTransactionsHeaders_: InventoryCore.js:155
inv_repairQcUaeLedgerWarehousesFromQcSheet: InventoryCore.js:1301
isBlank: Purchases.js:182, Purchases.js:474
isDeliveredStatus_: AppCore.js:605
logError_: AppCore.js:519
logInventoryTxnBatch_: InventoryCore.js:373
logInventoryTxn_: InventoryCore.js:354
lookupCatalog_: Sales.js:309
markOrdersSuccess_: AppCore.js:1801
normSku_: Sales.js:290
normalizeHeaders_: AppCore.js:453
normalizeSku: Sales.js:530
normalizeWarehouseCode_: AppCore.js:584
onEdit: AppCore.js:1215
onOpen: AppCore.js:1004
orders_alert_: Orders.js:39
orders_applyHeaderStyle_: Orders.js:540
orders_applyOrdersFormats_: Orders.js:572
orders_assertPurchasesHeadersForOrders_: Orders.js:604
orders_clearOrdersData_: Orders.js:564
orders_ensureSheet_: Orders.js:525
orders_parseDate_: Orders.js:633
orders_removeFilterIfAny_: Orders.js:533
orders_syncFromPurchasesByOrderIds_: Orders.js:294
orders_tryGetUi_: Orders.js:31
purchasesOnEditDefaults_: Purchases.js:135
purchasesOnEditSku_: SkuUtils.js:363
purchasesOnEdit_: Purchases.js:317
purchases_backfillDefaults_: Purchases.js:513
purchases_backfillOrderDefaults_: Purchases.js:447
purchases_backfillSkuSafe_: Purchases.js:426
purchases_clearComputedColumns_: Purchases.js:613
purchases_ensureLineIds_: Purchases.js:92
purchases_installFormulasCore_: Purchases.js:642
purchases_maybeAutoRepairFormulas_: Purchases.js:39
purchases_removeFilterIfAny_: Purchases.js:416
purchases_repairAutofill: Purchases.js:588
rebuildInventoryEGFromLedger: InventoryCore.js:799
rebuildInventoryUAEFromLedger: InventoryCore.js:590
rebuildOrdersSummary: Orders.js:112
rebuildShipmentsCnUaeStatus_: ShipmentsCore.js:333
recalcSalesRowAmounts_: Sales.js:602
repairBlankHeadersByPosition_: AppCore.js:478
resolveCol_: AppCore.js:411
resolveUaeWarehouseFromCourier_: ShipmentsCore.js:788, ShipmentsCore.js:1259
runCoreTests: AppCore.js:1305
safeAlert_: AppCore.js:634
safeConfirmYesNo_: AppCore.js:674
safeConfirm_: AppCore.js:647
safePromptText_: AppCore.js:657
salesEgBackfillFromCatalog: Sales.js:639
salesEgBackfillFromCatalog_: Sales.js:482
salesEgOnEdit_: Sales.js:411
seedShipmentsUaeEgFromInventoryUae: ShipmentsCore.js:1018, ShipmentsCore.js:1489
setIfBetter_: Orders.js:164, Orders.js:347
setupCatalogEgLayout: CatalogCore.js:68
setupInventoryCoreLayout: InventoryCore.js:75
setupInventoryLedger_: InventoryCore.js:93
setupInventorySnapshotEG_: InventoryCore.js:246
setupInventorySnapshotUAE_: InventoryCore.js:223
setupLogisticsLayout: Logistics.js:118
setupOrdersLayout: Orders.js:56
setupOrdersLayoutHardReset: Orders.js:86
setupPurchasesLayout: Purchases.js:245
setupPurchasesLayoutHardReset: Purchases.js:351
setupQC_UAE_: Logistics.js:134
setupQcLayout: Logistics.js:97
setupSalesLayout: Sales.js:49
setupShipmentsCnUaeStatusValidation_: ShipmentsCore.js:299
setupShipmentsCnUae_: Logistics.js:153
setupShipmentsLayouts: Logistics.js:107
setupShipmentsUaeEg_: Logistics.js:179
shipmentsCnUaeOnEdit_: ShipmentsCore.js:249
sku_backfillPurchasesSku: SkuUtils.js:345
sku_backfillPurchasesSku_: SkuUtils.js:261
sku_cleanText_: SkuUtils.js:46
sku_clearCatalogIndexCache_: SkuUtils.js:133
sku_fingerprint_: SkuUtils.js:126
sku_generateFromText_: SkuUtils.js:89
sku_getCatalogIndex_: SkuUtils.js:145
sku_headers_: SkuUtils.js:11
sku_lookupFromCatalog_: SkuUtils.js:194
sku_normalizeDigits_: SkuUtils.js:30
sku_normalizeSku_: SkuUtils.js:61
sku_rebuildCatalogIndex_: SkuUtils.js:469
sku_registerDraftToCatalog_: SkuUtils.js:206
sku_tokenize_: SkuUtils.js:70
syncPurchasesToShipmentsCnUae: ShipmentsCore.js:526
syncSalesFromOrdersSheet: Sales.js:105
testCatalogLookup: Sales.js:628
testOrdersModule_: Orders.js:649
testPurchasesModule_: Purchases.js:817
testSkuUtils_: SkuUtils.js:458
test_LogisticsSetup: Logistics.js:265
test_manualInventoryTxn: InventoryCore.js:1199
touchesCol_: Sales.js:433
updateAllShipmentsStatusAndTotals: ShipmentsCore.js:493
updateShipmentsCnUaeStatusAndTotals: ShipmentsCore.js:54
updateShipmentsUaeEgStatusAndTotals: ShipmentsCore.js:364
withLock_: AppCore.js:553
write_: InventoryCore.js:474
```

### UsedButNotDefined (occur → file:line)
```text
fn: AppCore.js:560, AppCore.js:571
getOrCreateSheet_: CatalogCore.js:72, InventoryCore.js:97, InventoryCore.js:227, InventoryCore.js:250, Logistics.js:136, Logistics.js:155, Logistics.js:181, Sales.js:53
isDelivered_: Sales.js:183 (locally defined inline; not a global issue)
qcOnEdit_: AppCore.js:1121
qc_generateFromPurchases_: AppCore.js:1731, AppCore.js:1749
shipmentsUaeEgOnEdit_: AppCore.js:1141
syncQCtoInventory_UAE: AppCore.js:893, AppCore.js:1773
syncShipmentsUaeEgToInventory: AppCore.js:894, AppCore.js:1788
```

### DefinedButNeverUsed (non-blocking)
```text
L: Purchases.js:655
R: Purchases.js:656
_canon_: AppCore.js:456
_installPurchaseFormulasCore: Purchases.js:602
backfillShipmentsUaeEgFromInventory: ShipmentsCore.js:987, ShipmentsCore.js:1458
catalog_syncFromInventoryEg: CatalogCore.js:147
coco_debugOrdersSyncStatus: AppCore.js:1864
coco_installTriggers: AppCore.js:1257
coco_onEditInstallable: AppCore.js:1233
coco_preflightAndRepair: AppCore.js:906
coco_processSyncQueueNow: AppCore.js:1850
coco_uninstallTriggers: AppCore.js:1286
colLetter_: AppCore.js:400
debug_listTriggers: AppCore.js:1899
inv_backfillMissingTxnIds: InventoryCore.js:1226
inv_fullRebuildFromLogistics: AppCore.js:882
inv_repairInventoryTransactionsHeaders: InventoryCore.js:136
inv_repairQcUaeLedgerWarehousesFromQcSheet: InventoryCore.js:1301
onEdit: AppCore.js:1215
onOpen: AppCore.js:1004
purchasesOnEdit_: Purchases.js:317
purchases_repairAutofill: Purchases.js:588
resolveCol_: AppCore.js:411
runCoreTests: AppCore.js:1305
safeConfirmYesNo_: AppCore.js:674
safePromptText_: AppCore.js:657
salesEgBackfillFromCatalog: Sales.js:639
setupOrdersLayoutHardReset: Orders.js:86
setupPurchasesLayoutHardReset: Purchases.js:351
sku_lookupFromCatalog_: SkuUtils.js:194
sku_rebuildCatalogIndex_: SkuUtils.js:469
syncSalesFromOrdersSheet: Sales.js:105
testCatalogLookup: Sales.js:628
testOrdersModule_: Orders.js:649
testPurchasesModule_: Purchases.js:817
testSkuUtils_: SkuUtils.js:458
test_LogisticsSetup: Logistics.js:265
test_manualInventoryTxn: InventoryCore.js:1199
updateAllShipmentsStatusAndTotals: ShipmentsCore.js:493
```

## Findings

### BLOCKERS
1) QC generation missing
- Symptom: `coco_processSyncQueue` leaves QC generation flags set (`QC_GEN_ALL/QUEUE`) and records “qc_generateFromPurchases_ is not defined”, so the queue never drains; onOpen menu “Generate QC from Purchases…” throws on click.
- Root cause: `qc_generateFromPurchases_` and its prompt handler are absent from the working tree. Archived implementations exist in `CocoERP_CostPolicy_Files.zip/ShipmentsCore.js` (`qc_generateFromPurchasesPrompt` at ~1369, `qc_generateFromPurchases_` at ~1406).
- References: AppCore.js:1033, 1727-1756; ShipmentsCore.js (no definition).
- Minimal fix proposal (restore archived code):
  ```diff
  --- a/ShipmentsCore.js
  +++ b/ShipmentsCore.js
  +// Restored from CocoERP_CostPolicy_Files.zip/ShipmentsCore.js
  +function qc_generateFromPurchasesPrompt() { /* … archived body … */ }
  +
  +function qc_generateFromPurchases_(optOrderIdOrOrderIds) { /* … archived body … */ }
  ```
- Verification: run `npm run check:syntax`; trigger the onOpen menu item or call `qc_generateFromPurchases_()` in Apps Script and confirm QC_UAE rows are created plus QC_GEN flags clear.

2) QC recalculation + onEdit disabled
- Symptom: QC_UAE edits no longer recompute Qty OK/Missing or QC Result; menu “Recalc QC Quantities & Result” is dead.
- Root cause: `qc_recalcQuantitiesAndResult`, `qc_recalcRows_`, and `qcOnEdit_` were removed. Archived versions exist at ~1785, ~1845, and ~1812 in `CocoERP_CostPolicy_Files.zip/ShipmentsCore.js`.
- References: AppCore.js:1121 (dispatcher), 1034 (menu); ShipmentsCore.js (no definitions).
- Minimal fix proposal:
  ```diff
  --- a/ShipmentsCore.js
  +++ b/ShipmentsCore.js
  +function qc_recalcQuantitiesAndResult(opts) { /* … archived body … */ }
  +function qcOnEdit_(e) { /* … archived body … */ }
  +function qc_recalcRows_(qcSh, qcMap, rowStart, numRows, opts) { /* … archived body … */ }
  ```
- Verification: edit a QC_UAE row touching Qty fields and confirm Qty OK/Missing/QC Result update; menu call should alert updated counts.

3) QC → Inventory sync missing
- Symptom: When QC_INV_SYNC_FLAG is set, `coco_processSyncQueue` re-queues with “syncQCtoInventory_UAE is not defined” and never posts ledger entries; onOpen menu “Sync QC → Inventory (UAE)” fails.
- Root cause: `syncQCtoInventory_UAE` removed; archived implementation at ~1951 in `CocoERP_CostPolicy_Files.zip/ShipmentsCore.js`.
- References: AppCore.js:889, 1042, 1772; ShipmentsCore.js lacks function.
- Minimal fix proposal:
  ```diff
  --- a/ShipmentsCore.js
  +++ b/ShipmentsCore.js
  +function syncQCtoInventory_UAE(opts) { /* … archived body … */ }
  ```
- Verification: run `coco_processSyncQueue` with QC_INV_SYNC_FLAG=1 and ensure ledger rows are written; menu action should complete without “not defined” errors.

4) Shipments_UAE_EG onEdit handler missing
- Symptom: Shipments_UAE_EG edits never recompute totals/status or box costs; relies on stale values while still setting inventory-sync flags.
- Root cause: `shipmentsUaeEgOnEdit_` absent; archived version at ~1292 in `CocoERP_CostPolicy_Files.zip/ShipmentsCore.js`.
- References: AppCore.js:1141-1150.
- Minimal fix proposal:
  ```diff
  --- a/ShipmentsCore.js
  +++ b/ShipmentsCore.js
  +function shipmentsUaeEgOnEdit_(e) { /* … archived body … */ }
  ```
- Verification: edit cost/qty/status cells in Shipments_UAE_EG and confirm totals/status update and sync flags remain consistent.

5) Shipments_UAE_EG → Inventory sync missing
- Symptom: `coco_processSyncQueue` hits SHIP_UAE_EG_INV stage and re-queues with “syncShipmentsUaeEgToInventory is not defined”; onOpen menu “Sync Shipments UAE→EG” fails.
- Root cause: `syncShipmentsUaeEgToInventory` removed; archived implementation at ~2198 in `CocoERP_CostPolicy_Files.zip/ShipmentsCore.js`.
- References: AppCore.js:894, 1043, 1787-1795; ShipmentsCore.js lacks function.
- Minimal fix proposal:
  ```diff
  --- a/ShipmentsCore.js
  +++ b/ShipmentsCore.js
  +function syncShipmentsUaeEgToInventory() { /* … archived body … */ }
  ```
- Verification: set SHIP_UAE_EG_INV_SYNC_FLAG=1 then run `coco_processSyncQueue`; ledger rows should be created and flag cleared.

6) QC onOpen menu actions broken (prompt + recalc)
- Symptom: Selecting “Generate QC from Purchases…” or “Recalc QC Quantities & Result” throws immediately.
- Root cause: menu targets are undefined (see blockers 1 & 2); AppCore still builds menu.
- References: AppCore.js:1033-1035.
- Minimal fix proposal: restore functions per blockers 1 & 2; if deferred, temporarily hide menu items to avoid runtime errors.
- Verification: open spreadsheet menu and ensure items execute without exceptions.

7) Queue resilience: requeue loop risk
- Symptom: Missing QC/Shipments sync handlers cause `coco_processSyncQueue` to re-set flags every run, effectively pinning the queue “busy” forever.
- Root cause: absent handlers (blockers 1,3,5) without guard to clear or backoff.
- References: AppCore.js:1727-1756, 1770-1796.
- Minimal fix proposal: after restoring handlers, keep existing flags; if restoration is delayed, add defensive guard to log once and drop flags to unblock other stages.
- Verification: clear flags, rerun queue, confirm it exits when no pending work.

8) QC helper chain missing (qc_recalcRows_) blocks downstream sync quality
- Symptom: Even if qc_generateFromPurchases_ is restored, recalculation helper is still missing, so QC rows would remain partially computed.
- Root cause: helper removed entirely (see blocker 2); required by qcOnEdit_ and manual recalc.
- References: ShipmentsCore.js (no qc_recalcRows_); archive ~1845.
- Minimal fix proposal: restore helper with exact archived logic to keep QC invariants.
- Verification: run `qc_recalcQuantitiesAndResult({ updateResult: true })` and confirm Qty OK/Missing/Result recomputed.

9) getOrCreateSheet_ undefined across modules
- Symptom: Logistics/Catalog setup falls back to `_fallbackGetOrCreateSheet_`; if fallback diverges from intended behavior, layout repairs may skip sheet creation.
- Root cause: no global `getOrCreateSheet_` implementation; callers already guard with `typeof`.
- References: CatalogCore.js:72; InventoryCore.js:97/227/250; Logistics.js:136/155/181; Sales.js:53.
- Minimal fix proposal: reintroduce the helper (create-if-missing + return sheet) consistent with callers’ expectations or remove guards if fallback is canonical.
- Verification: run `setupShipmentsLayouts` and `setupCatalogEgLayout`, confirm missing sheets are created without manual intervention.

10) Duplicate helper definitions in ShipmentsCore
- Symptom: `_fillShipmentUaeEgFromInventory_` and `_getInventoryUaeInfoForSku_` declared twice (ShipmentsCore.js:766 & 1203; 694 & 1131). Later definition overrides earlier, risking divergent fixes.
- Root cause: merged file contains duplicated blocks.
- References: ShipmentsCore.js:694-1018, 1120-1315.
- Minimal fix proposal: deduplicate to a single authoritative definition (keep latest, remove stale copy) to prevent silent override.
- Verification: rerun `npm run check:syntax` and exercise Shipments_UAE_EG fill/backfill flows to ensure only one implementation is executed.

### HIGH-RISK (non-blocking runtime)
- QC/Shipments menu items currently point to missing handlers (see blockers 1,2,5); until restored, user-facing errors likely.
- getOrCreateSheet_ absence may mask sheet-creation drift; rely on fallback for now.
- Shipments_UAE_EG totals/cost policy likely outdated because cost-policy patch isn’t applied; ensure restored handlers use per-unit extras policy from archive when reintroducing.

### NON-BLOCKING WARNINGS
- Numerous unused globals reported by ESLint (see DefinedButNeverUsed list) — cleanup optional.
- Duplicate helper definitions in ShipmentsCore (see blocker 10) — refactor when restoring missing functions.

## Entry Point/Contract Check
- AppCore dispatcher `_dispatchOnEdit_`: calls `purchasesOnEditDefaults_`, `purchasesOnEditSku_`, `shipmentsCnUaeOnEdit_`, `shipmentsUaeEgOnEdit_`, `salesEgOnEdit_`, `qcOnEdit_`. Missing handlers: `shipmentsUaeEgOnEdit_`, `qcOnEdit_`.
- Sync pipeline: `coco_processSyncQueue` stages — `syncPurchasesToShipmentsCnUae` ✅ defined; `qc_generateFromPurchases_` ❌ missing; `syncQCtoInventory_UAE` ❌ missing; `syncShipmentsUaeEgToInventory` ❌ missing; Orders stage uses `rebuildOrdersSummary`/`orders_syncFromPurchasesByOrderIds_` ✅ defined.
- Listed entry points: `coco_processSyncQueue` (AppCore.js:1661), `coco_processSyncQueueNow` (AppCore.js:1850), `qc_generateFromPurchases_` (missing), `syncPurchasesToShipmentsCnUae` (ShipmentsCore.js:526), `syncQCtoInventory_UAE` (missing), `syncShipmentsUaeEgToInventory` (missing), `rebuildOrdersSummary` (Orders.js:97), `orders_syncFromPurchasesByOrderIds_` (Orders.js:294), `syncSalesFromOrdersSheet` (Sales.js:105).
