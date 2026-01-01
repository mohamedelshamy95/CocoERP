# CocoERP – AI Coding Agent Instructions

This repo is a Google Apps Script–based ERP around Google Sheets. The kernel lives in AppCore.js and defines canonical sheet names, headers, UI helpers, locks, triggers, and the global onEdit/onOpen. Modules implement features for Purchases, Orders, Logistics (QC + Shipments), Inventory, Catalog, Sales, and SKU utilities.

## Big Picture
- Core kernel: AppCore.js exposes `APP` constants (SHEETS, COLS, SETTINGS), header aliasing, `withLock_`, `ensureSheetSchema_`, safe UI helpers, error logging, and the only global `onOpen(e)` and `onEdit(e)`.
- Data flows:
  - Purchases → Orders summary ([Orders.js])
  - Purchases → QC_UAE rows → Inventory ledger IN → Inventory_UAE snapshot ([Logistics.js], [InventoryCore.js])
  - Purchases → Shipments_CN_UAE ([ShipmentsCore.js])
  - Shipments_UAE_EG → Inventory ledger OUT (UAE) and IN (EG) → Inventory snapshots ([ShipmentsCore.js], [InventoryCore.js])
  - Sales_EG → Inventory ledger OUT → Inventory_EG snapshot ([Sales.js], [InventoryCore.js])
- Orchestration: `coco_processSyncQueue()` in AppCore.js processes debounced queues for Orders, QC generation, CN→UAE shipments sync, and inventory sync flags for QC and UAE→EG shipments.

## Workflows
- Lint + syntax checks (local): `npm run lint`, `npm run lint:fix`, `npm run check:syntax` (uses Node’s `--check` to validate file syntax only).
- clasp deployment (daily workflow):
  - Pull latest script state: `clasp pull`
  - Push changes to Apps Script: `clasp push`
- In the spreadsheet:
  - Menu entries (from `onOpen`) drive setup and sync: Preflight/Repair, install/uninstall triggers, Purchases/Orders setup, Logistics/QC setup, Inventory snapshot rebuild, and sync actions.
  - Triggers: `coco_installTriggers()` adds installable `onEdit` handler and a time-based trigger for `coco_processSyncQueue` (typically every 1 minute).

## Conventions & Patterns
- Headers: Always reference via `APP.COLS.*`. Before reads/writes, call `normalizeHeaders_(sheet, 1)` and prefer `ensureSheetSchema_(name, headers, { addMissing: true })` to create missing columns without clearing data.
- Locks: Any write path that appends/updates rows or properties should be wrapped in `withLock_(name, fn)` to avoid concurrent writers and respect re-entrancy in the same execution.
- UI safety: Use `safeAlert_`, `safeConfirm_`, `safePromptText_`. Only show dialogs in explicit UI actions; `onEdit`/time triggers must not call `SpreadsheetApp.getUi()`.
- Idempotency:
  - Inventory ledger writes use deterministic Txn IDs via `_inv_makeTxnId_` and batch append with de-dupe.
  - Snapshot rebuilds clear data rows (keep header) and rewrite from ledger aggregations.
  - Sync queues set DocumentProperties flags/arrays and requeue on failure.
- Header repair: Known drifts are self-healed (e.g., duplicate `Warehouse` headers in ledger) via `inv_repairInventoryTransactionsHeaders_()`.
- Warehouse normalization: `normalizeWarehouseCode_()` maps legacy aliases to canonical codes (UAE: `KOR`, `ATTIA`; EG: `TAN-GH`). Prefer `APP.WAREHOUSE_GROUPS` for regional checks.

## Module Highlights
- Purchases ([Purchases.js]):
  - Layout via `setupPurchasesLayout()` (non-destructive) and `setupPurchasesLayoutHardReset()` (destructive).
  - On edit: `purchasesOnEditDefaults_` fills Currency/FX/Ship/Customs for edited rows; SKU handling comes from `SkuUtils.js` (`purchasesOnEditSku_`).
  - Formulas: `_installPurchaseFormulasCore()` installs ARRAYFORMULAs for per-order aggregation and derived columns; call `purchases_repairAutofill()` for deterministic repair.
- Orders ([Orders.js]): `rebuildOrdersSummary()` aggregates Purchases to 1 row per Order ID. `orders_syncFromPurchasesByOrderIds_()` supports incremental sync (used by the queue).
- Logistics ([Logistics.js]): layout installers for `QC_UAE`, `Shipments_CN_UAE`, `Shipments_UAE_EG`; number/date formats and validations.
- Shipments ([ShipmentsCore.js]):
  - Status/totals updaters for CN→UAE and UAE→EG (`updateShipmentsCnUaeStatusAndTotals`, `updateShipmentsUaeEgStatusAndTotals`).
  - Sync Purchases → `Shipments_CN_UAE` using line-level keys, stable per-order `Shipment ID` sequences (`CN-000001`, ...).
  - UAE→EG integration: auto-fill from `Inventory_UAE` on SKU edits and seed planning rows.
- Inventory ([InventoryCore.js]):
  - Ledger batch writer `logInventoryTxnBatch_()` with de-dupe by Txn ID.
  - Snapshots rebuild (`rebuildInventoryUAEFromLedger`, `rebuildInventoryEGFromLedger`), clamp invariants (e.g., zero qty ⇒ zero value), and overship handling.
  - Full rebuild from logistics: `inv_fullRebuildFromLogistics()`.
- SKU utilities ([SkuUtils.js]): Tokenize name/variant, generate normalized SKU, catalog fingerprint index (`sku_getCatalogIndex_`), Purchases SKU backfill.
- Sales ([Sales.js]): `setupSalesLayout`, `salesEgOnEdit_` auto-fill and delivered-date handling, and `syncSalesFromOrdersSheet()` posting SALE_EG ledger OUT with cost from Inventory_EG or Catalog.

## Integration Points
- Settings: `Settings` sheet seeded/ensured by `ensureSettingsSheet_()`; use `getSetting_()` helpers for defaults: FX AED→EGP, ship per order, customs %.
- Cache: DocumentCache for short-lived indices (e.g., catalog SKU index).
- Properties: DocumentProperties hold sync queues and flags; see `APP.INTERNAL.*` keys.
- Advanced Services: Drive v3 enabled in `appsscript.json` (for invoice image previews).

## Practical Examples
- Adding a new computed Purchases field: extend `PURCHASE_HEADERS` and add an ARRAYFORMULA in `_installPurchaseFormulasCore()` using `colLetter_()` helpers; ensure `ensureSheetSchema_()` runs before writing the formula.
- Writing ledger transactions: prefer `logInventoryTxnBatch_([payloads])` with `type: 'IN'|'OUT'`, canonical headers mapped via `APP.COLS.INV_TXNS`, and wrap call in `withLock_('INV_LEDGER_WRITE', ...)` if writing directly.
- Extending onEdit: never add global `onEdit`; instead, add a module-level handler and wire it in AppCore’s dispatcher by sheet name.

---
Questions or unclear areas? If any sections don’t match your sheet schemas or current flows, say what’s off and we’ll refine this guide.
