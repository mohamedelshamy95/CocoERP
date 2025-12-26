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

// Money tolerance used for snapshot invariants (EGP). Helps clamp float noise.
const INV_VALUE_TOL_EGP = 0.05;

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

    if (typeof safeAlert_ === 'function') safeAlert_(
      'تم تجهيز Inventory_Transactions و Inventory_UAE و Inventory_EG ✔️'
    ); else Logger.log(
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
  try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) { }

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

    if (typeof safeAlert_ === 'function') safeAlert_('Inventory_Transactions headers repaired ✔️'); else Logger.log('Inventory_Transactions headers repaired ✔️');
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
  try { normalizeHeaders_(ledgerSh, 1); } catch (e) { }
  try {
    if (typeof ensureSheetSchema_ === 'function') {
      ensureSheetSchema_(APP.SHEETS.INVENTORY_TXNS, INV_TXN_HEADERS, { addMissing: true, headerRow: 1 });
    }
  } catch (e) { }

  const headers = options.headers || INV_TXN_HEADERS;

  // Map payload keys by header label (canonical)
  const keyByHeader = {};
  keyByHeader[APP.COLS.INV_TXNS.TXN_ID] = 'txnId';
  keyByHeader[APP.COLS.INV_TXNS.TXN_DATE] = 'txnDate';
  keyByHeader[APP.COLS.INV_TXNS.SOURCE_TYPE] = 'sourceType';
  keyByHeader[APP.COLS.INV_TXNS.SOURCE_ID] = 'sourceId';
  keyByHeader[APP.COLS.INV_TXNS.BATCH_CODE] = 'batchCode';
  keyByHeader[APP.COLS.INV_TXNS.SKU] = 'sku';
  keyByHeader[APP.COLS.INV_TXNS.PRODUCT_NAME] = 'productName';
  keyByHeader[APP.COLS.INV_TXNS.VARIANT] = 'variant';
  keyByHeader[APP.COLS.INV_TXNS.WAREHOUSE] = 'warehouse';
  keyByHeader[APP.COLS.INV_TXNS.QTY_IN] = 'qtyIn';
  keyByHeader[APP.COLS.INV_TXNS.QTY_OUT] = 'qtyOut';
  keyByHeader[APP.COLS.INV_TXNS.UNIT_COST] = 'unitCostEgp';
  keyByHeader[APP.COLS.INV_TXNS.TOTAL_COST] = 'totalCostEgp';
  keyByHeader[APP.COLS.INV_TXNS.CURRENCY] = 'currency';
  keyByHeader[APP.COLS.INV_TXNS.UNIT_PRICE_ORIG] = 'unitPriceOrig';
  keyByHeader[APP.COLS.INV_TXNS.NOTES] = 'notes';

  // Normalize payload -> ledger row
  function buildRow_(p) {
    const type = String(p.type || '').toUpperCase().trim();
    const qty = Number(p.qty || 0);

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

    // Ensure Txn ID is always populated (unless explicitly provided)
    out.txnId = p.txnId || _inv_makeTxnId_(out);

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

    // Deterministic Txn ID (used for debugging + optional dedupe)
    if (!out.txnId) {
      out.txnId = out.txnId || (typeof _inv_makeTxnId_ === 'function' ? _inv_makeTxnId_(out) : '');
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
    if (!rows.length) return 0;

    // Idempotency: dedupe by stable Txn ID
    const map = getHeaderMap_(ledgerSh);
    const txnIdCol = map[APP.COLS.INV_TXNS.TXN_ID] || map['Txn ID'];
    const existing = new Set();

    if (txnIdCol) {
      const lr = ledgerSh.getLastRow();
      if (lr >= 2) {
        const vals = ledgerSh.getRange(2, txnIdCol, lr - 1, 1).getValues();
        for (let i = 0; i < vals.length; i++) {
          const v = String(vals[i][0] || '').trim();
          if (v) existing.add(v);
        }
      }
    }

    const seen = new Set();
    const toWrite = [];

    // Txn ID is the first column in INV_TXN_HEADERS
    for (const r of rows) {
      const id = String(r[0] || '').trim();
      if (!id) continue;
      if (existing.has(id) || seen.has(id)) continue;
      seen.add(id);
      toWrite.push(r);
    }

    let appended = 0;
    for (let i = 0; i < toWrite.length; i += CHUNK) {
      const chunk = toWrite.slice(i, i + CHUNK);
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
    const tol = (typeof INV_VALUE_TOL_EGP === 'number') ? INV_VALUE_TOL_EGP : 0.05;

    const ledgerSh = (typeof ensureSheet_ === 'function')
      ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS)
      : getSheet_(APP.SHEETS.INVENTORY_TXNS);

    const invSh = (typeof ensureSheet_ === 'function')
      ? ensureSheet_(APP.SHEETS.INVENTORY_UAE)
      : getSheet_(APP.SHEETS.INVENTORY_UAE);

    // Self-heal common ledger header drift (Warehouse duplicates)
    try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) { }

    const ledgerMap = getHeaderMap_(ledgerSh);
    const invMap = getHeaderMap_(invSh);

    const lastRow = ledgerSh.getLastRow();

    // No ledger rows → clear snapshot
    if (lastRow < 2) {
      if (invSh.getLastRow() > 1) {
        invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
      }
      return;
    }

    const data = ledgerSh.getRange(2, 1, lastRow - 1, ledgerSh.getLastColumn()).getValues();

    const idxSku = (ledgerMap[APP.COLS.INV_TXNS.SKU] || ledgerMap['SKU']) - 1;
    const idxWh = (ledgerMap[APP.COLS.INV_TXNS.WAREHOUSE] || ledgerMap['Warehouse']) - 1;
    const idxProduct = (ledgerMap[APP.COLS.INV_TXNS.PRODUCT_NAME] || ledgerMap['Product Name']) - 1;
    const idxVariant = (ledgerMap[APP.COLS.INV_TXNS.VARIANT] || ledgerMap['Variant / Color']) - 1;
    const idxQtyIn = (ledgerMap[APP.COLS.INV_TXNS.QTY_IN] || ledgerMap['Qty In']) - 1;
    const idxQtyOut = (ledgerMap[APP.COLS.INV_TXNS.QTY_OUT] || ledgerMap['Qty Out']) - 1;
    const idxUnitCost = (ledgerMap[APP.COLS.INV_TXNS.UNIT_COST] || ledgerMap['Unit Cost (EGP)']) - 1;
    const idxTotalCost = (ledgerMap[APP.COLS.INV_TXNS.TOTAL_COST] || ledgerMap['Total Cost (EGP)']) - 1;
    const idxTxnDate = (ledgerMap[APP.COLS.INV_TXNS.TXN_DATE] || ledgerMap['Txn Date']) - 1;
    const idxSrcType = (ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] || ledgerMap['Source Type']) - 1;
    const idxSrcId = (ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] || ledgerMap['Source ID']) - 1;

    // Aggregate by (SKU + Warehouse + Variant) for UAE warehouses only
    const keyMap = {};

    data.forEach(function (row) {
      const sku = row[idxSku];
      if (!sku) return;

      let wh = (row[idxWh] || '').toString().trim();
      if (!wh) return;

      if (typeof normalizeWarehouseCode_ === 'function') {
        wh = normalizeWarehouseCode_(wh);
      }
      wh = String(wh || '').trim();

      if (!_inv_isUaeWarehouse_(wh)) return;

      const qtyIn = Number(row[idxQtyIn] || 0);
      const qtyOut = Number(row[idxQtyOut] || 0);
      if (qtyIn === 0 && qtyOut === 0) return;

      const product = row[idxProduct];
      const variant = row[idxVariant];

      const unitCost = Number(row[idxUnitCost] || 0);
      const totalCostCell = Number(row[idxTotalCost] || 0);

      // Signed valuation model:
      // - IN value adds
      // - OUT value subtracts
      const valIn = (qtyIn > 0) ? (totalCostCell || (unitCost * qtyIn)) : 0;
      const valOut = (qtyOut > 0) ? (totalCostCell || (unitCost * qtyOut)) : 0;

      const txnDate = row[idxTxnDate];
      const srcType = row[idxSrcType];
      const srcId = row[idxSrcId];

      const key = String(sku) + '||' + String(wh) + '||' + String(variant || '');

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
      rec.onHand += (qtyIn - qtyOut);
      rec.totalCost += (valIn - valOut);

      if (txnDate && (!rec.lastDate || txnDate > rec.lastDate)) {
        rec.lastDate = txnDate;
        rec.lastSourceType = srcType || '';
        rec.lastSourceId = srcId || '';
      }
    });

    // Clear snapshot
    if (invSh.getLastRow() > 1) {
      invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
    }

    const invHeaders = Object.keys(invMap).sort(function (a, b) {
      return invMap[a] - invMap[b];
    });

    const rows = [];

    Object.keys(keyMap).forEach(function (key) {
      const r = keyMap[key];

      // Clamp float noise
      if (Math.abs(r.onHand) < 1e-9) r.onHand = 0;
      if (Math.abs(r.totalCost) < tol) r.totalCost = 0;

      // Overship / negative on-hand: log + clamp to safe zero to avoid poisoning avg cost.
      if (r.onHand < 0) {
        try {
          logError_('rebuildInventoryUAEFromLedger', new Error('Negative On Hand Qty detected (overship). Snapshot clamped to 0/0.'), {
            sku: r.sku,
            variant: r.variant,
            warehouse: r.warehouse,
            onHand: r.onHand,
            totalCost: r.totalCost
          });
        } catch (e) { }
        r.onHand = 0;
        r.totalCost = 0;
      }

      // Critical invariant: Qty==0 ⇒ Value==0 (prevents “qty clamps but value doesn’t”).
      if (r.onHand === 0) {
        if (Math.abs(r.totalCost) > tol) {
          try {
            logError_('rebuildInventoryUAEFromLedger', new Error('Invariant violation: On Hand Qty is 0 but Total Cost is non-zero. Clamped to 0.'), {
              sku: r.sku,
              variant: r.variant,
              warehouse: r.warehouse,
              totalCost: r.totalCost
            });
          } catch (e) { }
        }
        r.totalCost = 0;
      }

      // Negative valuation (beyond tolerance) is invalid → log + clamp.
      if (r.totalCost < -tol) {
        try {
          logError_('rebuildInventoryUAEFromLedger', new Error('Negative Total Cost detected. Snapshot Total Cost clamped to 0.'), {
            sku: r.sku,
            variant: r.variant,
            warehouse: r.warehouse,
            onHand: r.onHand,
            totalCost: r.totalCost
          });
        } catch (e) { }
        r.totalCost = 0;
      } else if (r.totalCost < 0) {
        r.totalCost = 0;
      }

      if (!INV_SNAPSHOT_KEEP_ZERO && r.onHand === 0) return;

      const avgCost = (r.onHand > 0) ? (r.totalCost / r.onHand) : 0;

      const rowObj = {
        'SKU': r.sku,
        'Product Name': r.product,
        'Variant / Color': r.variant,
        'Warehouse (UAE)': r.warehouse,
        'On Hand Qty': r.onHand,
        'Allocated Qty': 0,
        'Available Qty': r.onHand,
        'Avg Cost (EGP)': avgCost,
        'Total Cost (EGP)': r.totalCost,
        'Last Txn Date': r.lastDate,
        'Last Source Type': r.lastSourceType,
        'Last Source ID': r.lastSourceId
      };

      rows.push(invHeaders.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; }));
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
 * Notes:
 * - Uses the SAME signed valuation model as UAE ONLY if EG OUT rows carry cost.
 * - If any EG OUT postings have 0 Unit Cost and 0 Total Cost, we fall back to IN-only valuation
 *   (avg cost from IN, then Total Cost = Avg * OnHand) and log a TODO.
 */
function rebuildInventoryEGFromLedger() {
  try {
    const tol = (typeof INV_VALUE_TOL_EGP === 'number') ? INV_VALUE_TOL_EGP : 0.05;

    const ledgerSh = (typeof ensureSheet_ === 'function')
      ? ensureSheet_(APP.SHEETS.INVENTORY_TXNS)
      : getSheet_(APP.SHEETS.INVENTORY_TXNS);

    const invSh = (typeof ensureSheet_ === 'function')
      ? ensureSheet_(APP.SHEETS.INVENTORY_EG)
      : getSheet_(APP.SHEETS.INVENTORY_EG);

    // Self-heal common ledger header drift (Warehouse duplicates)
    try { inv_repairInventoryTransactionsHeaders_(ledgerSh); } catch (e) { }

    const ledgerMap = getHeaderMap_(ledgerSh);
    const invMap = getHeaderMap_(invSh);

    const lastRow = ledgerSh.getLastRow();
    if (lastRow < 2) {
      if (invSh.getLastRow() > 1) {
        invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
      }
      return;
    }

    const data = ledgerSh.getRange(2, 1, lastRow - 1, ledgerSh.getLastColumn()).getValues();

    const idxSku = (ledgerMap[APP.COLS.INV_TXNS.SKU] || ledgerMap['SKU']) - 1;
    const idxWh = (ledgerMap[APP.COLS.INV_TXNS.WAREHOUSE] || ledgerMap['Warehouse']) - 1;
    const idxProd = (ledgerMap[APP.COLS.INV_TXNS.PRODUCT_NAME] || ledgerMap['Product Name']) - 1;
    const idxVar = (ledgerMap[APP.COLS.INV_TXNS.VARIANT] || ledgerMap['Variant / Color']) - 1;
    const idxQtyIn = (ledgerMap[APP.COLS.INV_TXNS.QTY_IN] || ledgerMap['Qty In']) - 1;
    const idxQtyOut = (ledgerMap[APP.COLS.INV_TXNS.QTY_OUT] || ledgerMap['Qty Out']) - 1;
    const idxUnitCost = (ledgerMap[APP.COLS.INV_TXNS.UNIT_COST] || ledgerMap['Unit Cost (EGP)']) - 1;
    const idxTotalCost = (ledgerMap[APP.COLS.INV_TXNS.TOTAL_COST] || ledgerMap['Total Cost (EGP)']) - 1;
    const idxTxnDate = (ledgerMap[APP.COLS.INV_TXNS.TXN_DATE] || ledgerMap['Txn Date']) - 1;
    const idxSrcType = (ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] || ledgerMap['Source Type']) - 1;
    const idxSrcId = (ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] || ledgerMap['Source ID']) - 1;

    // We compute both models, then choose:
    // - If ANY EG OUT rows are missing cost → fall back to IN-only valuation.
    let hasCostlessEgOut = false;
    let costlessEgOutLogged = 0;

    // Net-value map
    const keyMapNet = {};
    // IN-only map (safe fallback if OUT cost is missing)
    const keyMapInOnly = {};

    data.forEach(function (row, idx) {
      const sku = row[idxSku];
      if (!sku) return;

      let wh = (row[idxWh] || '').toString().trim();
      if (!wh) return;

      if (typeof normalizeWarehouseCode_ === 'function') {
        wh = normalizeWarehouseCode_(wh);
      }
      wh = String(wh || '').trim();

      const whUpper = wh.toUpperCase();
      if (!_inv_isEgWarehouse_(whUpper)) return;

      const qtyIn = Number(row[idxQtyIn] || 0);
      const qtyOut = Number(row[idxQtyOut] || 0);
      if (qtyIn === 0 && qtyOut === 0) return;

      const unitCost = Number(row[idxUnitCost] || 0);
      const totalCostCell = Number(row[idxTotalCost] || 0);
      const variant = row[idxVar] || '';

      const key = String(sku) + '||' + String(whUpper) + '||' + String(variant);

      // Detect costless OUT rows (historically common when OUT is written without avg-cost).
      if (qtyOut > 0 && unitCost === 0 && totalCostCell === 0) {
        hasCostlessEgOut = true;
        if (costlessEgOutLogged < 5) {
          costlessEgOutLogged++;
          try {
            logError_('rebuildInventoryEGFromLedger', new Error('EG OUT row has 0 Unit Cost and 0 Total Cost (valuation unsafe for net model). Using IN-only model.'), {
              sku: sku,
              variant: variant,
              warehouse: whUpper,
              qtyOut: qtyOut,
              row: idx + 2,
              sourceType: row[idxSrcType],
              sourceId: row[idxSrcId]
            });
          } catch (e) { }
        }
      }

      // ----- Net model (only safe if OUT has cost) -----
      if (!keyMapNet[key]) {
        keyMapNet[key] = {
          sku: sku,
          product: row[idxProd],
          variant: variant,
          warehouse: whUpper,
          onHand: 0,
          totalCost: 0,
          lastDate: null,
          lastSrcType: '',
          lastSrcId: ''
        };
      }
      const recN = keyMapNet[key];

      const valIn = (qtyIn > 0) ? (totalCostCell || (unitCost * qtyIn)) : 0;
      const valOut = (qtyOut > 0) ? (totalCostCell || (unitCost * qtyOut)) : 0;

      recN.onHand += (qtyIn - qtyOut);
      recN.totalCost += (valIn - valOut);

      const txnDate = row[idxTxnDate];
      if (txnDate && (!recN.lastDate || txnDate > recN.lastDate)) {
        recN.lastDate = txnDate;
        recN.lastSrcType = row[idxSrcType] || '';
        recN.lastSrcId = row[idxSrcId] || '';
      }

      // ----- IN-only model (fallback) -----
      if (!keyMapInOnly[key]) {
        keyMapInOnly[key] = {
          sku: sku,
          product: row[idxProd],
          variant: variant,
          warehouse: whUpper,
          qtyIn: 0,
          qtyOut: 0,
          totalCostIn: 0,
          lastUnitCost: 0,
          lastDate: null,
          lastSrcType: '',
          lastSrcId: ''
        };
      }
      const recI = keyMapInOnly[key];

      if (qtyIn > 0) {
        const valInOnly = (totalCostCell || (unitCost * qtyIn));
        recI.qtyIn += qtyIn;
        recI.totalCostIn += valInOnly;
        recI.lastUnitCost = unitCost;
      }
      recI.qtyOut += qtyOut;

      if (txnDate && (!recI.lastDate || txnDate > recI.lastDate)) {
        recI.lastDate = txnDate;
        recI.lastSrcType = row[idxSrcType] || '';
        recI.lastSrcId = row[idxSrcId] || '';
      }
    });

    // Clear old snapshot
    if (invSh.getLastRow() > 1) {
      invSh.getRange(2, 1, invSh.getLastRow() - 1, invSh.getLastColumn()).clearContent();
    }

    const invHeaders = Object.keys(invMap).sort(function (a, b) { return invMap[a] - invMap[b]; });
    const rows = [];

    if (hasCostlessEgOut) {
      // Fallback: IN-only valuation. Total Cost derived from avg IN cost and remaining qty.
      // TODO: Ensure EG OUT postings carry Unit Cost / Total Cost so we can switch to signed net model.
      Object.keys(keyMapInOnly).forEach(function (key) {
        const r = keyMapInOnly[key];
        let onHand = Number(r.qtyIn || 0) - Number(r.qtyOut || 0);

        // Clamp float noise
        if (Math.abs(onHand) < 1e-9) onHand = 0;

        if (onHand < 0) {
          try {
            logError_('rebuildInventoryEGFromLedger', new Error('Negative On Hand Qty detected (overship). Snapshot clamped to 0/0.'), {
              sku: r.sku,
              variant: r.variant,
              warehouse: r.warehouse,
              onHand: onHand,
              qtyIn: r.qtyIn,
              qtyOut: r.qtyOut
            });
          } catch (e) { }
          onHand = 0;
        }

        if (!INV_SNAPSHOT_KEEP_ZERO && onHand === 0) return;

        const avgCost = r.qtyIn ? (r.totalCostIn / r.qtyIn) : Number(r.lastUnitCost || 0);
        let totalCost = avgCost * onHand;

        if (Math.abs(totalCost) < tol) totalCost = 0;

        if (onHand === 0) {
          if (Math.abs(totalCost) > tol) {
            try {
              logError_('rebuildInventoryEGFromLedger', new Error('Invariant violation: On Hand Qty is 0 but Total Cost is non-zero. Clamped to 0.'), {
                sku: r.sku,
                variant: r.variant,
                warehouse: r.warehouse,
                totalCost: totalCost
              });
            } catch (e) { }
          }
          totalCost = 0;
        }

        if (totalCost < -tol) {
          try {
            logError_('rebuildInventoryEGFromLedger', new Error('Negative Total Cost detected. Snapshot Total Cost clamped to 0.'), {
              sku: r.sku,
              variant: r.variant,
              warehouse: r.warehouse,
              onHand: onHand,
              totalCost: totalCost
            });
          } catch (e) { }
          totalCost = 0;
        } else if (totalCost < 0) {
          totalCost = 0;
        }

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

        rows.push(invHeaders.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; }));
      });
    } else {
      // Signed net-value model (safe when OUT carries cost)
      Object.keys(keyMapNet).forEach(function (key) {
        const r = keyMapNet[key];

        if (Math.abs(r.onHand) < 1e-9) r.onHand = 0;
        if (Math.abs(r.totalCost) < tol) r.totalCost = 0;

        if (r.onHand < 0) {
          try {
            logError_('rebuildInventoryEGFromLedger', new Error('Negative On Hand Qty detected (overship). Snapshot clamped to 0/0.'), {
              sku: r.sku,
              variant: r.variant,
              warehouse: r.warehouse,
              onHand: r.onHand,
              totalCost: r.totalCost
            });
          } catch (e) { }
          r.onHand = 0;
          r.totalCost = 0;
        }

        if (r.onHand === 0) {
          if (Math.abs(r.totalCost) > tol) {
            try {
              logError_('rebuildInventoryEGFromLedger', new Error('Invariant violation: On Hand Qty is 0 but Total Cost is non-zero. Clamped to 0.'), {
                sku: r.sku,
                variant: r.variant,
                warehouse: r.warehouse,
                totalCost: r.totalCost
              });
            } catch (e) { }
          }
          r.totalCost = 0;
        }

        if (r.totalCost < -tol) {
          try {
            logError_('rebuildInventoryEGFromLedger', new Error('Negative Total Cost detected. Snapshot Total Cost clamped to 0.'), {
              sku: r.sku,
              variant: r.variant,
              warehouse: r.warehouse,
              onHand: r.onHand,
              totalCost: r.totalCost
            });
          } catch (e) { }
          r.totalCost = 0;
        } else if (r.totalCost < 0) {
          r.totalCost = 0;
        }

        if (!INV_SNAPSHOT_KEEP_ZERO && r.onHand === 0) return;

        const avgCost = (r.onHand > 0) ? (r.totalCost / r.onHand) : 0;

        const rowObj = {
          'SKU': r.sku,
          'Product Name': r.product,
          'Variant / Color': r.variant,
          'Warehouse (EG)': r.warehouse,
          'On Hand Qty': r.onHand,
          'Allocated Qty': 0,
          'Available Qty': r.onHand,
          'Avg Cost (EGP)': avgCost,
          'Total Cost (EGP)': r.totalCost,
          'Last Txn Date': r.lastDate,
          'Last Source Type': r.lastSrcType,
          'Last Source ID': r.lastSrcId
        };

        rows.push(invHeaders.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; }));
      });
    }

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

    // Optional: seed UAE→EG planning rows after snapshots rebuild
    if (typeof seedShipmentsUaeEgFromInventoryUae === 'function') {
      try {
        seedShipmentsUaeEgFromInventoryUae();
      } catch (e) {
        logError_('inv_rebuildAllSnapshots.seedShipmentsUaeEgFromInventoryUae', e);
        // Non-fatal: inventory snapshots should still be considered rebuilt.
      }
    }

    if (typeof safeAlert_ === 'function') {
      safeAlert_('Inventory snapshots rebuilt (UAE + EG).');
    } else {
      Logger.log('Inventory snapshots rebuilt (UAE + EG).');
    }
  } catch (err) {
    try { logError_('inv_rebuildAllSnapshots', err); } catch (e) { }
    throw err;
  }
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



/** ===================== ONE-TIME REPAIR HELPERS (MANUAL RUN) ===================== */

/**
 * Backfill missing Txn IDs in Inventory_Transactions.
 * Safe to re-run; only fills blank Txn ID cells.
 */
function inv_backfillMissingTxnIds() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = getSheet_(APP.SHEETS.INVENTORY_TXNS);
    const map = getHeaderMap_(sh);

    const colTxnId = map[APP.COLS.INV_TXNS.TXN_ID];
    const colTxnDate = map[APP.COLS.INV_TXNS.TXN_DATE];
    const colSrcType = map[APP.COLS.INV_TXNS.SOURCE_TYPE];
    const colSrcId = map[APP.COLS.INV_TXNS.SOURCE_ID];
    const colBatch = map[APP.COLS.INV_TXNS.BATCH_CODE];
    const colSku = map[APP.COLS.INV_TXNS.SKU];
    const colWh = map[APP.COLS.INV_TXNS.WAREHOUSE];
    const colQtyIn = map[APP.COLS.INV_TXNS.QTY_IN];
    const colQtyOut = map[APP.COLS.INV_TXNS.QTY_OUT];
    const colUnit = map[APP.COLS.INV_TXNS.UNIT_COST];
    const colCur = map[APP.COLS.INV_TXNS.CURRENCY];
    const colUPO = map[APP.COLS.INV_TXNS.UNIT_PRICE_ORIG];

    if (!colTxnId || !colTxnDate || !colSku || !colWh || !colQtyIn || !colQtyOut) {
      throw new Error('inv_backfillMissingTxnIds: missing required ledger columns');
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return 0;

    const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
    const outIds = [];
    let touched = 0;

    values.forEach(function (r) {
      const existing = r[colTxnId - 1];
      if (existing && String(existing).trim()) {
        outIds.push([existing]);
        return;
      }

      const qtyIn = Number(r[colQtyIn - 1] || 0);
      const qtyOut = Number(r[colQtyOut - 1] || 0);
      const type = qtyIn ? 'IN' : (qtyOut ? 'OUT' : 'IN');

      const o = {
        type: type,
        sourceType: colSrcType ? r[colSrcType - 1] : '',
        sourceId: colSrcId ? r[colSrcId - 1] : '',
        batchCode: colBatch ? r[colBatch - 1] : '',
        sku: colSku ? r[colSku - 1] : '',
        warehouse: colWh ? r[colWh - 1] : '',
        qtyIn: qtyIn,
        qtyOut: qtyOut,
        unitCostEgp: colUnit ? Number(r[colUnit - 1] || 0) : 0,
        currency: colCur ? r[colCur - 1] : '',
        unitPriceOrig: colUPO ? Number(r[colUPO - 1] || 0) : 0,
        txnDate: r[colTxnDate - 1] ? new Date(r[colTxnDate - 1]) : new Date()
      };

      const id = (typeof _inv_makeTxnId_ === 'function') ? _inv_makeTxnId_(o) : '';
      outIds.push([id]);
      if (id) touched++;
    });

    if (touched) {
      sh.getRange(2, colTxnId, outIds.length, 1).setValues(outIds);
    }
    return touched;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Repair Warehouse column for QC_UAE ledger rows based on QC_UAE sheet (by QC ID).
 * This fixes legacy/default 'UAE-DXB' values and aligns ledger with QC.
 */
function inv_repairQcUaeLedgerWarehousesFromQcSheet() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const qcSh = getSheet_(APP.SHEETS.QC_UAE);
    const lgSh = getSheet_(APP.SHEETS.INVENTORY_TXNS);

    const qcMap = getHeaderMap_(qcSh);
    const lgMap = getHeaderMap_(lgSh);

    const qcColId = qcMap[APP.COLS.QC_UAE.QC_ID];
    const qcColWh = qcMap[APP.COLS.QC_UAE.WAREHOUSE];
    if (!qcColId || !qcColWh) throw new Error('inv_repairQcUaeLedgerWarehousesFromQcSheet: missing QC headers');
    const lgColSrcType = lgMap[APP.COLS.INV_TXNS.SOURCE_TYPE];
    const lgColSrcId = lgMap[APP.COLS.INV_TXNS.SOURCE_ID];
    const lgColWh = lgMap[APP.COLS.INV_TXNS.WAREHOUSE];
    if (!lgColSrcType || !lgColSrcId || !lgColWh) throw new Error('inv_repairQcUaeLedgerWarehousesFromQcSheet: missing ledger headers');

    const qcLast = qcSh.getLastRow();
    const lgLast = lgSh.getLastRow();
    if (qcLast < 2 || lgLast < 2) return 0;

    // Build QC_ID -> canonical warehouse
    const qcData = qcSh.getRange(2, 1, qcLast - 1, qcSh.getLastColumn()).getValues();
    const qcWhById = {};
    qcData.forEach(function (r) {
      const id = (r[qcColId - 1] || '').toString().trim();
      if (!id) return;
      const whRaw = (r[qcColWh - 1] || '').toString().trim();
      const wh = (typeof normalizeWarehouseCode_ === 'function') ? normalizeWarehouseCode_(whRaw) : whRaw;
      if (wh) qcWhById[id] = wh;
    });

    const lgData = lgSh.getRange(2, 1, lgLast - 1, lgSh.getLastColumn()).getValues();
    const outWh = [];
    let touched = 0;

    lgData.forEach(function (r) {
      const srcType = (r[lgColSrcType - 1] || '').toString().trim();
      if (srcType !== 'QC_UAE') {
        outWh.push([r[lgColWh - 1]]);
        return;
      }

      const qcId = (r[lgColSrcId - 1] || '').toString().trim();
      const desired = qcWhById[qcId] || '';
      const currentRaw = (r[lgColWh - 1] || '').toString().trim();
      const current = (typeof normalizeWarehouseCode_ === 'function') ? normalizeWarehouseCode_(currentRaw) : currentRaw;

      if (desired && desired !== current) {
        outWh.push([desired]);
        touched++;
      } else {
        outWh.push([r[lgColWh - 1]]);
      }
    });

    if (touched) {
      lgSh.getRange(2, lgColWh, outWh.length, 1).setValues(outWh);
    }
    return touched;
  } finally {
    lock.releaseLock();
  }
}