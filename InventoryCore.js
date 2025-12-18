/** =============================================================
 * InventoryCore.gs – Inventory Ledger + Snapshots + Sync
 * CocoERP v2.1
 * ============================================================= */

/** === Headers ================================================= */

// Ledger: كل حركة مخزون (شراء، شحن، بيع، تسوية…)
const INV_TXN_HEADERS = [
  APP.COLS.INV_TXNS.TXN_ID,          // 'Txn ID'
  APP.COLS.INV_TXNS.TXN_DATE,        // 'Txn Date'
  APP.COLS.INV_TXNS.SOURCE_TYPE,     // 'Source Type'
  APP.COLS.INV_TXNS.SOURCE_ID,       // 'Source ID'
  APP.COLS.INV_TXNS.BATCH_CODE,      // 'Batch Code'
  APP.COLS.INV_TXNS.SKU,             // 'SKU'
  APP.COLS.INV_TXNS.PRODUCT_NAME,    // 'Product Name'
  APP.COLS.INV_TXNS.VARIANT,         // 'Variant / Color'
  APP.COLS.INV_TXNS.WAREHOUSE,       // 'Warehouse'
  APP.COLS.INV_TXNS.QTY_IN,          // 'Qty In'
  APP.COLS.INV_TXNS.QTY_OUT,         // 'Qty Out'
  APP.COLS.INV_TXNS.UNIT_COST,       // 'Unit Cost (EGP)'
  APP.COLS.INV_TXNS.TOTAL_COST,      // 'Total Cost (EGP)'
  'Currency',
  'Unit Price (Orig)',
  APP.COLS.INV_TXNS.NOTES            // 'Notes'
];

// Snapshot behavior:
// false = default: show only SKUs with non-zero stock/value
// true  = keep rows even if On Hand Qty = 0 and Total Cost = 0
const INV_SNAPSHOT_KEEP_ZERO = false;

// Snapshot: مخزون الإمارات
const INV_UAE_HEADERS = [
  'SKU',
  'Product Name',
  'Variant / Color',
  'Warehouse (UAE)',
  'On Hand Qty',
  'Allocated Qty',
  'Available Qty',
  'Avg Cost (EGP)',
  'Total Cost (EGP)',
  'Last Txn Date',
  'Last Source Type',
  'Last Source ID'
];

// Snapshot: مخزون مصر
const INV_EG_HEADERS = [
  'SKU',
  'Product Name',
  'Variant / Color',
  'Warehouse (EG)',
  'On Hand Qty',
  'Allocated Qty',
  'Available Qty',
  'Avg Cost (EGP)',
  'Total Cost (EGP)',
  'Last Txn Date',
  'Last Source Type',
  'Last Source ID'
];

/** =============================================================
 *  Main Setup – entry point from menu
 * ============================================================= */

/**
 * جهّز شيت Inventory_Transactions + Inventory_UAE + Inventory_EG
 */
function setupInventoryCoreLayout() {
  try {
    setupInventoryLedger_();
    setupInventorySnapshotUAE_();
    setupInventorySnapshotEG_();

    SpreadsheetApp.getUi().alert(
      'تم تجهيز Inventory_Transactions و Inventory_UAE و Inventory_EG ✔️'
    );
  } catch (e) {
    logError_('setupInventoryCoreLayout', e);
    throw e;
  }
}

/** Setup Inventory_Transactions sheet (Ledger) */
function setupInventoryLedger_() {
  const ledgerSh = (typeof ensureSheet_ === 'function')
    ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS)
    : ((typeof getOrCreateSheet_ === 'function')
        ? getOrCreateSheet_(APP.SHEETS.INVENTORY_TXNS)
        : getSheet_(APP.SHEETS.INVENTORY_TXNS));

  _setupSheetWithHeaders_(ledgerSh, INV_TXN_HEADERS);

  // Self-heal common drift (e.g., legacy "Warehouse (EG)" + extra "Warehouse")
  try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) {}

  const map = getHeaderMap_(ledgerSh);
  _applyDateFormat_(ledgerSh, map[APP.COLS.INV_TXNS.TXN_DATE]);
  _applyIntFormat_(ledgerSh, [
    map[APP.COLS.INV_TXNS.QTY_IN],
    map[APP.COLS.INV_TXNS.QTY_OUT]
  ]);
  _applyDecimalFormat_(ledgerSh, [
    map[APP.COLS.INV_TXNS.UNIT_COST],
    map[APP.COLS.INV_TXNS.TOTAL_COST]
  ]);
}


/** =============================================================
 *  Ledger header repair (one-time / self-heal)
 * ============================================================= */

/**
 * Fixes common schema drift in Inventory_Transactions where legacy headers cause duplicates.
 *
 * Typical drift observed:
 * - Legacy header: "Warehouse (EG)" exists in col I (where the system writes).
 * - Preflight adds a missing canonical "Warehouse" header at the far right (empty).
 * - Snapshot rebuild reads the empty "Warehouse" col and outputs nothing.
 *
 * This repair:
 * - Renames the active warehouse column in the first INV_TXN_HEADERS columns to "Warehouse"
 * - Renames any extra "Warehouse" columns to "Warehouse (Extra)" (non-destructive)
 *
 * Safe to run multiple times (idempotent).
 */
function inv_repairInventoryTransactionsHeaders() {
  try {
    const ledgerSh = (typeof ensureSheet_ === 'function')
      ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS)
      : getSheet_(APP.SHEETS.INVENTORY_TXNS);

    inv_repairInventoryTransactionsHeaders_(ledgerSh);

    SpreadsheetApp.getUi().alert('Inventory_Transactions headers repaired ✔️');
  } catch (e) {
    logError_('inv_repairInventoryTransactionsHeaders', e);
    throw e;
  }
}

/**
 * Internal repair implementation (no UI).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ledgerSh
 */
function inv_repairInventoryTransactionsHeaders_(ledgerSh) {
  const sh = ledgerSh;
  if (!sh) return;

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  const headerRow = 1;
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0].map(function (v) {
    return String(v || '').trim();
  });

  const N = INV_TXN_HEADERS.length;

  // Find candidate indices
  const idxWarehouseCanonical = [];
  const idxWarehouseLegacy = []; // Warehouse (EG) / Warehouse (UAE) / any "Warehouse (..)"
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (!h) continue;
    if (h === APP.COLS.INV_TXNS.WAREHOUSE) idxWarehouseCanonical.push(i + 1);
    else if (/^Warehouse\s*\(/i.test(h)) idxWarehouseLegacy.push(i + 1);
  }

  // Decide which column is "active": prefer within the first N columns (where writes happen)
  const activeCandidate = (function () {
    // If a legacy warehouse header exists within first N, it's very likely the active one
    const legacyInFirst = idxWarehouseLegacy.find(function (c) { return c <= N; });
    if (legacyInFirst) return legacyInFirst;

    // Else if canonical exists within first N, use it
    const canonInFirst = idxWarehouseCanonical.find(function (c) { return c <= N; });
    if (canonInFirst) return canonInFirst;

    // Else fallback to first canonical anywhere
    if (idxWarehouseCanonical.length) return idxWarehouseCanonical[0];

    // Else fallback to first legacy anywhere
    if (idxWarehouseLegacy.length) return idxWarehouseLegacy[0];

    return null;
  })();

  if (!activeCandidate) return;

  const targetHeader = APP.COLS.INV_TXNS.WAREHOUSE; // "Warehouse"

  // Rename active candidate header to canonical if needed
  if (headers[activeCandidate - 1] !== targetHeader) {
    sh.getRange(headerRow, activeCandidate).setValue(targetHeader);
    headers[activeCandidate - 1] = targetHeader;
  }

  // Rename any other "Warehouse" header (duplicates) to avoid getHeaderMap_ picking the wrong one.
  for (let i = 0; i < headers.length; i++) {
    const col = i + 1;
    if (col === activeCandidate) continue;

    const h = headers[i];
    if (h === targetHeader) {
      // If it's an empty column, just rename it. If it has values, still rename (non-destructive).
      sh.getRange(headerRow, col).setValue('Warehouse (Extra)');
    }
  }
}


/** Setup Inventory_UAE snapshot */
function setupInventorySnapshotUAE_() {
  const uaeInvSh = (typeof ensureSheet_ === 'function')
    ? ensureSheet_(APP.SHEETS.INVENTORY_UAE)
    : ((typeof getOrCreateSheet_ === 'function')
        ? getOrCreateSheet_(APP.SHEETS.INVENTORY_UAE)
        : getSheet_(APP.SHEETS.INVENTORY_UAE));

  _setupSheetWithHeaders_(uaeInvSh, INV_UAE_HEADERS);

  const map = getHeaderMap_(uaeInvSh);
  _applyIntFormat_(uaeInvSh, [
    map['On Hand Qty'],
    map['Allocated Qty'],
    map['Available Qty']
  ]);
  _applyDecimalFormat_(uaeInvSh, [
    map['Avg Cost (EGP)'],
    map['Total Cost (EGP)']
  ]);
  _applyDateFormat_(uaeInvSh, map['Last Txn Date']);
}

/** Setup Inventory_EG snapshot */
function setupInventorySnapshotEG_() {
  const egInvSh = (typeof ensureSheet_ === 'function')
    ? ensureSheet_(APP.SHEETS.INVENTORY_EG)
    : ((typeof getOrCreateSheet_ === 'function')
        ? getOrCreateSheet_(APP.SHEETS.INVENTORY_EG)
        : getSheet_(APP.SHEETS.INVENTORY_EG));

  _setupSheetWithHeaders_(egInvSh, INV_EG_HEADERS);

  const map = getHeaderMap_(egInvSh);
  _applyIntFormat_(egInvSh, [
    map['On Hand Qty'],
    map['Allocated Qty'],
    map['Available Qty']
  ]);
  _applyDecimalFormat_(egInvSh, [
    map['Avg Cost (EGP)'],
    map['Total Cost (EGP)']
  ]);
  _applyDateFormat_(egInvSh, map['Last Txn Date']);
}

/** =============================================================
 *  Local formatting helpers (Inventory)
 * ============================================================= */

/**
 * Apply integer number format (0) for given column indexes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number[]} indexes
 */
function _applyIntFormat_(sh, indexes) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) return;
  const numRows = maxRows - 1;

  indexes
    .filter(function (col) { return col && col > 0; })
    .forEach(function (col) {
      sh.getRange(2, col, numRows, 1).setNumberFormat('0');
    });
}

/**
 * Apply decimal number format (0.00) for given column indexes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number[]} indexes
 */
function _applyDecimalFormat_(sh, indexes) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) return;
  const numRows = maxRows - 1;

  indexes
    .filter(function (col) { return col && col > 0; })
    .forEach(function (col) {
      sh.getRange(2, col, numRows, 1).setNumberFormat('0.00');
    });
}

/**
 * Apply date format (yyyy-mm-dd) for given column index or list of indexes.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number|number[]} indexes
 */
function _applyDateFormat_(sh, indexes) {
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) return;
  const numRows = maxRows - 1;

  const cols = Array.isArray(indexes) ? indexes : [indexes];
  cols
    .filter(function (col) { return col && col > 0; })
    .forEach(function (col) {
      sh.getRange(2, col, numRows, 1).setNumberFormat('yyyy-mm-dd');
    });
}

/* ============================================================
 * InventoryCore – Inventory Ledger + Snapshots + Sync
 * ============================================================ */

/**
 * Helper: log single inventory transaction (IN / OUT) into ledger.
 *
 * opts:
 *  type         : 'IN' أو 'OUT'
 *  sourceType   : مثال: 'Purchase', 'Manual', 'Sale', 'QC_UAE', 'SHIP_UAE_EG'
 *  sourceId     : مثال: Order ID أو QC ID أو Shipment ID
 *  batchCode    : Batch Code لو موجود
 *  sku          : كود الـ SKU
 *  productName  : اسم المنتج (اختياري)
 *  variant      : اللون / الفاريانت (اختياري)
 *  warehouse    : المخزن (مثلاً 'UAE-DXB' أو 'EG-CAI')
 *  qty          : الكمية المحركة (موجبة دايمًا)
 *  unitCostEgp  : تكلفة الوحدة بالجنيه (لو متوفر)
 *  currency     : عملة السعر الأصلي (اختياري)
 *  unitPriceOrig: سعر الوحدة بالعملة الأصلية (اختياري)
 *  notes        : ملاحظات (اختياري)
 *  txnDate      : تاريخ الحركة (اختياري – لو مش متحدد بيستخدم تاريخ النهاردة)
 */
/**
 * Log a single inventory transaction into Inventory_Transactions ledger.
 * - Keeps backward compatibility with existing callers.
 * - Internally routes to the batch writer for performance + stability.
 *
 * @param {Object} payload
 */
function logInventoryTxn_(payload) {
  try {
    if (!payload) throw new Error('logInventoryTxn_: missing payload');
    logInventoryTxnBatch_([payload]);
  } catch (e) {
    logError_('logInventoryTxn_', e, { payload: payload });
    throw e;
  }
}

/**
 * Batch writer for inventory transactions.
 * - Uses Range.setValues (chunked) instead of appendRow loops.
 * - Safe to call from other modules (Sales, Shipments, QC).
 *
 * @param {Object[]} payloads
 * @param {Object=} opts Optional: { ledgerSheet, headers }
 * @return {number} appended count
 */
function logInventoryTxnBatch_(payloads, opts) {
  const list = (payloads || []).filter(function (p) { return !!p; });
  if (!list.length) return 0;

  const options = opts || {};
  const ledgerSh = options.ledgerSheet || ((typeof ensureSheet_ === 'function') ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS) : getSheet_(APP.SHEETS.INVENTORY_TXNS));

  // Ensure ledger schema exists (non-destructive)
  try { normalizeHeaders_(ledgerSh, 1); } catch (e) {}
  try {
    if (typeof ensureSheetSchema_ === 'function') {
      ensureSheetSchema_(APP.SHEETS.INVENTORY_TXNS, INV_TXN_HEADERS, { addMissing: true, headerRow: 1 });
    }
  } catch (e) {}

  const headers = options.headers || INV_TXN_HEADERS;

  // Map payload keys by header label (canonical)
  const keyByHeader = {};
  keyByHeader[APP.COLS.INV_TXNS.TXN_DATE]    = 'txnDate';
  keyByHeader[APP.COLS.INV_TXNS.TYPE]        = 'type';
  keyByHeader[APP.COLS.INV_TXNS.SOURCE_TYPE] = 'sourceType';
  keyByHeader[APP.COLS.INV_TXNS.SOURCE_ID]   = 'sourceId';
  keyByHeader[APP.COLS.INV_TXNS.BATCH_CODE]  = 'batchCode';
  keyByHeader[APP.COLS.INV_TXNS.SKU]         = 'sku';
  keyByHeader[APP.COLS.INV_TXNS.PRODUCT_NAME]= 'productName';
  keyByHeader[APP.COLS.INV_TXNS.VARIANT]     = 'variant';
  keyByHeader[APP.COLS.INV_TXNS.WAREHOUSE]   = 'warehouse';
  keyByHeader[APP.COLS.INV_TXNS.QTY_IN]      = 'qtyIn';
  keyByHeader[APP.COLS.INV_TXNS.QTY_OUT]     = 'qtyOut';
  keyByHeader[APP.COLS.INV_TXNS.UNIT_COST]      = 'unitCostEgp';
  keyByHeader[APP.COLS.INV_TXNS.TOTAL_COST]     = 'totalCostEgp';
  keyByHeader[APP.COLS.INV_TXNS.CURRENCY]      = 'currency';
  keyByHeader[APP.COLS.INV_TXNS.UNIT_PRICE_ORIG]= 'unitPriceOrig';
  keyByHeader[APP.COLS.INV_TXNS.NOTES]         = 'notes';

  // Normalize payload -> ledger row
  function buildRow_(p) {
    const type = String(p.type || '').toUpperCase().trim();
    const qty  = Number(p.qty || 0);

    const txnDate = p.txnDate ? new Date(p.txnDate) : new Date();
    const wh = (typeof normalizeWarehouseCode_ === 'function')
      ? normalizeWarehouseCode_(p.warehouse || '')
      : String(p.warehouse || '').trim();

    const out = {
      txnDate: txnDate,
      type: type,
      sourceType: String(p.sourceType || '').trim(),
      sourceId: String(p.sourceId || '').trim(),
      batchCode: String(p.batchCode || '').trim(),
      sku: String(p.sku || '').trim(),
      productName: p.productName || '',
      variant: p.variant || '',
      warehouse: wh,
      qtyIn: 0,
      qtyOut: 0,
      unitCostEgp: Number(p.unitCostEgp || 0),
      totalCostEgp: Number(p.totalCostEgp || 0),
      currency: String(p.currency || '').trim(),
      unitPriceOrig: p.unitPriceOrig,
      notes: p.notes || ''
    };

    if (!out.sku) throw new Error('logInventoryTxnBatch_: missing SKU');
    if (!out.warehouse) throw new Error('logInventoryTxnBatch_: missing warehouse');
    if (!qty || qty <= 0) throw new Error('logInventoryTxnBatch_: invalid qty for SKU ' + out.sku);

    if (type === 'IN') out.qtyIn = qty;
    else if (type === 'OUT') out.qtyOut = qty;
    else throw new Error('logInventoryTxnBatch_: invalid type ' + type);

    // Deterministic Total Cost
    if (!out.totalCostEgp) {
      const basisQty = (type === 'IN') ? out.qtyIn : out.qtyOut;
      out.totalCostEgp = out.unitCostEgp ? (out.unitCostEgp * basisQty) : 0;
    }

    return headers.map(function (h) {
      const key = keyByHeader[h];
      return (key ? out[key] : '');
    });
  }

  const rows = [];
  for (let i = 0; i < list.length; i++) {
    rows.push(buildRow_(list[i]));
  }

  // Write in chunks under a lock (to avoid concurrent writers racing on lastRow)
  const CHUNK = 500;

  const write_ = function () {
    let appended = 0;
    for (let i = 0; i < rows.length; i += CHUNK) {
      const chunk = rows.slice(i, i + CHUNK);
      const start = ledgerSh.getLastRow() + 1;
      ledgerSh.getRange(start, 1, chunk.length, headers.length).setValues(chunk);
      appended += chunk.length;
    }
    return appended;
  };

  if (typeof withLock_ === 'function') {
    return withLock_('INV_LEDGER_WRITE', write_);
  }

  // Fallback lock
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return write_();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Create a deterministic Txn ID based on the transaction payload.
 * This keeps IDs stable across rebuilds (as long as the payload is the same).
 */
function _inv_makeTxnId_(o) {
  const parts = [
    o.type || '',
    o.sourceType || '',
    o.sourceId || '',
    o.batchCode || '',
    o.sku || '',
    o.warehouse || '',
    String(o.qtyIn || 0),
    String(o.qtyOut || 0),
    String(o.unitCostEgp || 0),
    o.currency || '',
    String(o.unitPriceOrig || 0),
    (o.txnDate instanceof Date) ? Utilities.formatDate(o.txnDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : ''
  ].join('|');

  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, parts, Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    const b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return 'TXN-' + hex.slice(0, 12).toUpperCase();
}



/**
 * Rebuild Inventory_UAE snapshot sheet from Inventory_Transactions ledger.
 * UAE warehouses only (names: "UAE" أو اللي بتبدأ بـ "UAE-").
 * كل مخزن (UAE-ATTIA / UAE-KOR / UAE-DXB) بيطلع في صف مستقل.
 */
function _inv_normWh_(wh) {
  return String(wh || '').trim().toUpperCase();
}

function _inv_isWhInList_(wh, list) {
  const w = _inv_normWh_(wh);
  if (!w) return false;
  for (let i = 0; i < list.length; i++) {
    if (_inv_normWh_(list[i]) === w) return true;
  }
  return false;
}

function _inv_isUaeWarehouse_(wh) {
  const list = (APP.WAREHOUSE_GROUPS && APP.WAREHOUSE_GROUPS.UAE) ? APP.WAREHOUSE_GROUPS.UAE : ['UAE', 'UAE-DXB', 'UAE-KOR', 'UAE-ATTIA', 'KOR', 'ATTIA'];
  const w = _inv_normWh_(wh);
  return _inv_isWhInList_(w, list) || w.indexOf('UAE-') === 0;
}

function _inv_isEgWarehouse_(wh) {
  const list = (APP.WAREHOUSE_GROUPS && APP.WAREHOUSE_GROUPS.EG) ? APP.WAREHOUSE_GROUPS.EG : ['EG-CAI', 'EG-TANTA', 'TAN-GH'];
  const w = _inv_normWh_(wh);
  return _inv_isWhInList_(w, list) || w.indexOf('EG') === 0;
}

function rebuildInventoryUAEFromLedger() {
  try {
    const ledgerSh = (typeof ensureSheet_ === 'function') ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS) : getSheet_(APP.SHEETS.INVENTORY_TXNS);
    const invSh    = (typeof ensureSheet_ === 'function') ? ensureSheet_(APP.SHEETS.INVENTORY_UAE) : getSheet_(APP.SHEETS.INVENTORY_UAE);

    // Self-heal common ledger header drift (Warehouse duplicates)
    try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) {}

    const ledgerMap = getHeaderMap_(ledgerSh);
    const invMap    = getHeaderMap_(invSh);

    const lastRow = ledgerSh.getLastRow();

    // لو مفيش حركات → امسح Snapshot وخلاص
    if (lastRow < 2) {
      if (invSh.getLastRow() > 1) {
        invSh
          .getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn())
          .clearContent();
      }
      return;
    }

    const data = ledgerSh
      .getRange(2, 1, lastRow - 1, ledgerSh.getLastColumn())
      .getValues();

    const idxSku       = ledgerMap[APP.COLS.INV_TXNS.SKU]          - 1;
    const idxWh        = ledgerMap[APP.COLS.INV_TXNS.WAREHOUSE]    - 1;
    const idxProduct   = ledgerMap[APP.COLS.INV_TXNS.PRODUCT_NAME] - 1;
    const idxVariant   = ledgerMap[APP.COLS.INV_TXNS.VARIANT]      - 1;
    const idxQtyIn     = ledgerMap[APP.COLS.INV_TXNS.QTY_IN]       - 1;
    const idxQtyOut    = ledgerMap[APP.COLS.INV_TXNS.QTY_OUT]      - 1;
    const idxUnitCost  = ledgerMap[APP.COLS.INV_TXNS.UNIT_COST]    - 1;
    const idxTotalCost = ledgerMap[APP.COLS.INV_TXNS.TOTAL_COST]   - 1;
    const idxTxnDate   = ledgerMap[APP.COLS.INV_TXNS.TXN_DATE]     - 1;
    const idxSrcType   = ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE]  - 1;
    const idxSrcId     = ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID]    - 1;

    // تجميع حسب (SKU + Warehouse + Variant) لمخازن الإمارات فقط
    const keyMap = {};
    data.forEach(function (row) {
      const sku = row[idxSku];
      if (!sku) return;

      const whRaw = row[idxWh];
      const wh    = (whRaw || '').toString().trim();
      if (!wh) return;

      const whUpper = wh.toUpperCase();
      if (!_inv_isUaeWarehouse_(wh)) { return; }

      const product   = row[idxProduct];
      const variant   = row[idxVariant];
      const qtyIn     = Number(row[idxQtyIn]  || 0);
      const qtyOut    = Number(row[idxQtyOut] || 0);
      const unitCost  = Number(row[idxUnitCost] || 0);
      const totalCost = Number(row[idxTotalCost] || 0) || unitCost * qtyIn;
      const txnDate   = row[idxTxnDate];
      const srcType   = row[idxSrcType];
      const srcId     = row[idxSrcId];

      const key = sku + '||' + wh + '||' + (variant || '');

      if (!keyMap[key]) {
        keyMap[key] = {
          sku: sku,
          product: product,
          variant: variant,
          warehouse: wh,
          onHand: 0,
          totalCost: 0,
          lastDate: txnDate || null,
          lastSourceType: srcType || '',
          lastSourceId: srcId || ''
        };
      }

      const rec = keyMap[key];
      rec.onHand    += qtyIn - qtyOut;
      rec.totalCost += totalCost;

      if (txnDate && (!rec.lastDate || txnDate > rec.lastDate)) {
        rec.lastDate       = txnDate;
        rec.lastSourceType = srcType || '';
        rec.lastSourceId   = srcId || '';
      }
    });

    // امسح Snapshot القديم
    if (invSh.getLastRow() > 1) {
      invSh
        .getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn())
        .clearContent();
    }

    const invHeaders = Object.keys(invMap).sort(function (a, b) {
      return invMap[a] - invMap[b];
    });
    const rows = [];

    Object.keys(keyMap).forEach(function (key) {
      const r = keyMap[key];

      if (!INV_SNAPSHOT_KEEP_ZERO && !r.onHand && !r.totalCost) return;

      const avgCost = r.onHand ? r.totalCost / r.onHand : 0;

      const rowObj = {};
      rowObj['SKU']              = r.sku;
      rowObj['Product Name']     = r.product;
      rowObj['Variant / Color']  = r.variant;
      rowObj['Warehouse (UAE)']  = r.warehouse;
      rowObj['On Hand Qty']      = r.onHand;
      rowObj['Allocated Qty']    = 0;
      rowObj['Available Qty']    = r.onHand;
      rowObj['Avg Cost (EGP)']   = avgCost;
      rowObj['Total Cost (EGP)'] = r.totalCost;
      rowObj['Last Txn Date']    = r.lastDate;
      rowObj['Last Source Type'] = r.lastSourceType;
      rowObj['Last Source ID']   = r.lastSourceId;

      const rowArr = invHeaders.map(function (h) {
        return rowObj[h] !== undefined ? rowObj[h] : '';
      });
      rows.push(rowArr);
    });

    if (rows.length) {
      invSh.getRange(2, 1, rows.length, invSh.getLastColumn()).setValues(rows);
    }

  } catch (e) {
    logError_('rebuildInventoryUAEFromLedger', e);
    throw e;
  }
}

/**
 * Rebuild Inventory_EG snapshot sheet from Inventory_Transactions ledger.
 * نفس فكرة الإمارات لكن بيفلتر على المخازن اللي بتبدأ بـ "EG".
 */
function rebuildInventoryEGFromLedger() {
  try {
    const ledgerSh = (typeof ensureSheet_ === 'function') ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS) : getSheet_(APP.SHEETS.INVENTORY_TXNS);
    const invSh    = (typeof ensureSheet_ === 'function') ? ensureSheet_(APP.SHEETS.INVENTORY_EG) : getSheet_(APP.SHEETS.INVENTORY_EG);

    // Self-heal common ledger header drift (Warehouse duplicates)
    try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) {}
    const ledgerMap = getHeaderMap_(ledgerSh);
    const invMap    = getHeaderMap_(invSh);

    const lastRow = ledgerSh.getLastRow();
    if (lastRow < 2) {
      if (invSh.getLastRow() > 1) {
        invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
      }
      return;
    }

    const data = ledgerSh.getRange(2, 1, lastRow - 1, ledgerSh.getLastColumn()).getValues();

    const idxSku       = ledgerMap[APP.COLS.INV_TXNS.SKU] - 1;
    const idxWh        = ledgerMap[APP.COLS.INV_TXNS.WAREHOUSE] - 1;
    const idxProd      = ledgerMap[APP.COLS.INV_TXNS.PRODUCT_NAME] - 1;
    const idxVar       = ledgerMap[APP.COLS.INV_TXNS.VARIANT] - 1;
    const idxQtyIn     = ledgerMap[APP.COLS.INV_TXNS.QTY_IN] - 1;
    const idxQtyOut    = ledgerMap[APP.COLS.INV_TXNS.QTY_OUT] - 1;
    const idxUnitCost  = ledgerMap[APP.COLS.INV_TXNS.UNIT_COST] - 1;
    const idxTotalCost = ledgerMap[APP.COLS.INV_TXNS.TOTAL_COST] - 1;
    const idxTxnDate   = ledgerMap[APP.COLS.INV_TXNS.TXN_DATE] - 1;
    const idxSrcType   = ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] - 1;
    const idxSrcId     = ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] - 1;

    const keyMap = {};

    data.forEach(row => {
      const wh = (row[idxWh] || '').toString().trim().toUpperCase();
      if (!_inv_isEgWarehouse_(wh)) return;

      const qtyIn  = Number(row[idxQtyIn] || 0);
      const qtyOut = Number(row[idxQtyOut] || 0);
      if (qtyIn === 0 && qtyOut === 0) return;

      const sku = row[idxSku];
      const key = sku + '||' + wh + '||' + (row[idxVar] || '');

      if (!keyMap[key]) {
        keyMap[key] = {
          sku: sku,
          product: row[idxProd],
          variant: row[idxVar],
          warehouse: wh,
          qtyIn: 0,
          qtyOut: 0,
          totalCostIn: 0,
          lastUnitCost: 0,
          lastDate: null,
          lastSrcType: '',
          lastSrcId: ''
        };
      }

      const rec = keyMap[key];

      // نعتبر فقط الحركات IN للتكلفة
      if (qtyIn > 0) {
        const unitCost = Number(row[idxUnitCost] || 0);
        const totalCost = Number(row[idxTotalCost] || unitCost * qtyIn);
        rec.qtyIn += qtyIn;
        rec.totalCostIn += totalCost;
        rec.lastUnitCost = unitCost;
      }

      rec.qtyOut += qtyOut;

      const txnDate = row[idxTxnDate];
      if (txnDate && (!rec.lastDate || txnDate > rec.lastDate)) {
        rec.lastDate = txnDate;
        rec.lastSrcType = row[idxSrcType];
        rec.lastSrcId = row[idxSrcId];
      }
    });

    // امسح القديم
    if (invSh.getLastRow() > 1) {
      invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
    }

    const invHeaders = Object.keys(invMap).sort((a, b) => invMap[a] - invMap[b]);
    const rows = [];

    Object.values(keyMap).forEach(r => {
      const onHand = r.qtyIn - r.qtyOut;
      if (!INV_SNAPSHOT_KEEP_ZERO && onHand <= 0) return;

      const avgCost = r.qtyIn ? r.totalCostIn / r.qtyIn : r.lastUnitCost;
      const totalCost = avgCost * onHand;

      const rowObj = {
        'SKU': r.sku,
        'Product Name': r.product,
        'Variant / Color': r.variant,
        'Warehouse (EG)': r.warehouse,
        'On Hand Qty': onHand,
        'Allocated Qty': 0,
        'Available Qty': onHand,
        'Avg Cost (EGP)': avgCost,
        'Total Cost (EGP)': totalCost,
        'Last Txn Date': r.lastDate,
        'Last Source Type': r.lastSrcType,
        'Last Source ID': r.lastSrcId
      };

      const rowArr = invHeaders.map(h => rowObj[h] !== undefined ? rowObj[h] : '');
      rows.push(rowArr);
    });

    if (rows.length) {
      invSh.getRange(2, 1, rows.length, invSh.getLastColumn()).setValues(rows);
    }
  } catch (e) {
    logError_('rebuildInventoryEGFromLedger', e);
    throw e;
  }
}

/**
 * Rebuild UAE + EG snapshots مع بعض.
 * المنيو في AppCore بينادي الفنكشن دي.
 */
function inv_rebuildAllSnapshots() {
  try {
    rebuildInventoryUAEFromLedger();
    rebuildInventoryEGFromLedger();
    if (typeof safeAlert_ === 'function') safeAlert_('Inventory snapshots rebuilt (UAE + EG).');
    else Logger.log('Inventory snapshots rebuilt (UAE + EG).');}
}

/**
 * Sync QC_UAE → Inventory_Transactions + Inventory_UAE
 * - لكل صف QC فيه SKU وكميات:
 *    - Qty In = Qty OK (الجديدة) أو (Qty Received - Qty Defective) لو الهيدر القديم لسه موجود.
 * - Unit Cost (EGP) بياخده من Purchases:
 *    - أولاً عن طريق Batch Code (لو موجود)
 *    - أو fallback على (Order ID + SKU)
 * - dedupe عن طريق قراءة الـ ledger (SourceType = QC_UAE + row رقم كـ note)
 */
/** syncQCtoInventory_UAE moved to ShipmentsCore.gs */


/**
 * UI entry point:
 * Generate QC_UAE rows from Purchases.
 * - يطلب منك Order ID (اختياري).
 * - لو سيبته فاضي → يولّد لكل الأوردرات اللي لسه ملهاش QC.
 */
/** qc_generateFromPurchasesPrompt moved to ShipmentsCore.gs */


/**
 * Generate QC_UAE rows from Purchases (one row per SKU)
 * - يدعم فلترة حسب Order ID (اختياري)
 * - يملأ: QC ID, Order ID, Shipment CN→UAE ID, SKU, Batch Code,
 *         Product Name, Variant / Color, Qty Ordered
 */
/** qc_generateFromPurchases_ moved to ShipmentsCore.gs */



/**
 * Recalculate QC_UAE quantities & result:
 * - لو Qty Missing و Qty Defective فاضيين ⇒ نحسب Missing = Qty Ordered - Qty OK
 * - نحدد QC Result = PASS / PARTIAL / FAIL لو فاضية
 * - نملأ QC Date بتاريخ اليوم لو فاضية
 */
/** qc_recalcQuantitiesAndResult moved to ShipmentsCore.gs */


/**
 * Test helper – اختياري
 * جرّب تشغّل الدالة دي من Run وتشوف حركة واحدة تتسجل في
 * Inventory_Transactions، وبعدها شغّل inv_rebuildAllSnapshots.
 */
function test_manualInventoryTxn() {
  logInventoryTxn_({
    type: 'IN',
    sourceType: 'ManualTest',
    sourceId: 'TEST-001',
    batchCode: 'TEST-BATCH-001',
    sku: 'TEST-SKU-001',
    productName: 'Test Product',
    variant: 'Beige',
    warehouse: 'UAE-DXB',
    qty: 10,
    unitCostEgp: 250,
    currency: 'AED',
    unitPriceOrig: 25,
    notes: 'Manual test txn'
  });
  inv_rebuildAllSnapshots();
}
