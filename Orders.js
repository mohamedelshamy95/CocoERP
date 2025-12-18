/** ============================================================
 * Orders.gs – Orders Summary for CocoERP v2.2+
 * ------------------------------------------------------------
 * - Source: Purchases sheet (line level, 1 row per SKU)
 * - Target: Orders sheet (1 row per Order ID)
 *
 * Depends on AppCore (kernel):
 *  - APP, getSheet_, ensureSheet_, ensureSheetSchema_, normalizeHeaders_
 *  - getHeaderMap_, logError_, ensureErrorLog_
 * ============================================================ */

/** Orders sheet headers (aligned with APP.COLS.ORDERS where possible) */
const ORDER_HEADERS = [
  APP.COLS.ORDERS.ORDER_ID,      // 'Order ID'
  'Order Date',
  'Platform',
  'Seller Name',
  'Currency',
  'Buyer Name',
  APP.COLS.ORDERS.TOTAL_LINES,   // 'Total Lines'
  APP.COLS.ORDERS.TOTAL_QTY,     // 'Total Qty'
  APP.COLS.ORDERS.TOTAL_ORIG,    // 'Total Order (Orig)'
  APP.COLS.ORDERS.TOTAL_EGP,     // 'Order Total (EGP)'
  APP.COLS.ORDERS.SHIP_EG,       // 'Ship UAE→EG (EGP)'
  APP.COLS.ORDERS.CUSTOMS,       // 'Customs/Fees (EGP)'
  APP.COLS.ORDERS.LANDED_COST,   // 'Landed Cost (EGP)'
  APP.COLS.ORDERS.UNIT_LANDED,   // 'Unit Landed Cost (EGP)'
  'Notes'
];

function orders_tryGetUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null; // Trigger / time-driven context
  }
}

function orders_alert_(msg) {
  var text = String(msg);
  var ui = orders_tryGetUi_();
  if (ui) {
    ui.alert(text);
    return;
  }
  if (typeof safeAlert_ === 'function') {
    safeAlert_(text);
    return;
  }
  Logger.log(text);
}

/** ============================================================
 * Layout (SAFE) – ensure headers/schema without wiping data
 * ============================================================ */
function setupOrdersLayout() {
  try {
    ensureErrorLog_();

    const shO = orders_ensureSheet_(APP.SHEETS.ORDERS);

    // Normalize headers if any legacy
    try { normalizeHeaders_(shO, 1); } catch (e) {}

    // Ensure schema (non-destructive)
    if (typeof ensureSheetSchema_ === 'function') {
      ensureSheetSchema_(APP.SHEETS.ORDERS, ORDER_HEADERS, { addMissing: true, headerRow: 1 });
    } else {
      // fallback: set header if empty
      if (shO.getLastRow() === 0) shO.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);
    }

    orders_applyHeaderStyle_(shO, ORDER_HEADERS);
    ensureErrorLog_();

    orders_alert_('✅ Orders layout ensured (بدون مسح بيانات).');
  } catch (e) {
    logError_('setupOrdersLayout', e);
    throw e;
  }
}

/** ============================================================
 * Layout (HARD RESET) – destructive
 * ============================================================ */
function setupOrdersLayoutHardReset() {
  try {
    ensureErrorLog_();

    const shO = orders_ensureSheet_(APP.SHEETS.ORDERS);
    const ui = orders_tryGetUi_();
    if (!ui) throw new Error('This action requires UI (run from spreadsheet menu).');
    const res = ui.alert('تحذير', 'ده هيمسح Orders بالكامل. متأكد؟', ui.ButtonSet.YES_NO);
    if (res !== ui.Button.YES) return;

    orders_removeFilterIfAny_(shO);
    shO.clear();
    shO.setFrozenRows(1);
    shO.setFrozenColumns(3);

    shO.getRange(1, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);

    orders_applyHeaderStyle_(shO, ORDER_HEADERS);
    orders_alert_('✅ Orders HARD RESET done.');
  } catch (e) {
    logError_('setupOrdersLayoutHardReset', e);
    throw e;
  }
}

/** ============================================================
 * Rebuild Orders summary from Purchases rows
 * ============================================================ */
function rebuildOrdersSummary() {
  try {
    ensureErrorLog_();

    const shP = getSheet_(APP.SHEETS.PURCHASES);
    const shO = orders_ensureSheet_(APP.SHEETS.ORDERS);

    // Normalize Purchases headers (just in case)
    try { normalizeHeaders_(shP, 1); } catch (e) {}

    const mapP = getHeaderMap_(shP, 1);
    orders_assertPurchasesHeadersForOrders_(mapP);

    const lastRow = shP.getLastRow();
    const lastCol = shP.getLastColumn();

    // Ensure Orders schema + header
    if (typeof ensureSheetSchema_ === 'function') {
      ensureSheetSchema_(APP.SHEETS.ORDERS, ORDER_HEADERS, { addMissing: true, headerRow: 1 });
    }
    orders_applyHeaderStyle_(shO, ORDER_HEADERS);

    // No data
    if (lastRow < 2) {
      orders_clearOrdersData_(shO);
      return;
    }

    const data = shP.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const idx = (h) => (mapP[h] ? (mapP[h] - 1) : -1);

    const iOrderId        = idx(APP.COLS.PURCHASES.ORDER_ID);
    const iOrderDate      = idx(APP.COLS.PURCHASES.ORDER_DATE);
    const iPlatform       = idx(APP.COLS.PURCHASES.PLATFORM);
    const iSellerName     = idx(APP.COLS.PURCHASES.SELLER);
    const iSku            = idx(APP.COLS.PURCHASES.SKU);
    const iCurrency       = idx(APP.COLS.PURCHASES.CURRENCY);
    const iBuyerName      = idx(APP.COLS.PURCHASES.BUYER_NAME);
    const iQty            = idx(APP.COLS.PURCHASES.QTY);
    const iNotes          = idx(APP.COLS.PURCHASES.NOTES);

    const iTotalOrig      = idx(APP.COLS.PURCHASES.TOTAL_ORIG);
    const iTotalEgp       = idx(APP.COLS.PURCHASES.TOTAL_EGP);
    const iShipEg         = idx(APP.COLS.PURCHASES.SHIP_EG);
    const iCustomsEgp     = idx(APP.COLS.PURCHASES.CUSTOMS_EGP);
    const iLandedEgp      = idx(APP.COLS.PURCHASES.LANDED_COST);
    const iUnitLandedEgp  = idx(APP.COLS.PURCHASES.UNIT_LANDED);

    /** @type {Object<string, any>} */
    const orders = {};

    const setIfBetter_ = (obj, key, val, mode) => {
      // mode: 'text' | 'date' | 'number'
      if (mode === 'text') {
        const v = String(val || '').trim();
        if (!v) return;
        if (!obj[key]) obj[key] = v;
        return;
      }
      if (mode === 'date') {
        const d = orders_parseDate_(val);
        if (!(d instanceof Date)) return;
        if (!(obj[key] instanceof Date)) { obj[key] = d; return; }
        // keep earliest
        if (d.getTime() < obj[key].getTime()) obj[key] = d;
        return;
      }
      // number
      const n = Number(val);
      if (!isFinite(n) || n === 0) return;
      if (!isFinite(Number(obj[key])) || Number(obj[key]) === 0) obj[key] = n;
    };

    const addNoteUnique_ = (obj, note) => {
      const n = String(note || '').trim();
      if (!n) return;
      const existing = String(obj.notes || '');
      if (!existing) { obj.notes = n; return; }
      if (existing.indexOf(n) === -1) obj.notes = existing + ' | ' + n;
    };

    data.forEach(function (row) {
      const orderId = (iOrderId >= 0) ? String(row[iOrderId] || '').trim() : '';
      if (!orderId) return;

      const sku = (iSku >= 0) ? String(row[iSku] || '').trim() : '';
      // SKU may be blank while importing; do not block Orders summary
      // if (!sku) return;

      const qty = (iQty >= 0) ? (Number(row[iQty]) || 0) : 0;

      let o = orders[orderId];
      if (!o) {
        o = orders[orderId] = {
          orderId: orderId,
          orderDate: '',
          platform: '',
          sellerName: '',
          currency: '',
          buyerName: '',
          totalLines: 0,
          totalQty: 0,
          totalOrderOrig: 0,
          orderTotalEGP: 0,
          shipUaeEg: 0,
          customsEGP: 0,
          landedCostEGP: 0,
          unitLandedEGP: 0,
          notes: ''
        };
      }

      o.totalLines += 1;
      o.totalQty   += qty;

      setIfBetter_(o, 'orderDate',      (iOrderDate >= 0 ? row[iOrderDate] : ''), 'date');
      setIfBetter_(o, 'platform',       (iPlatform >= 0 ? row[iPlatform] : ''), 'text');
      setIfBetter_(o, 'sellerName',     (iSellerName >= 0 ? row[iSellerName] : ''), 'text');
      setIfBetter_(o, 'currency',       (iCurrency >= 0 ? row[iCurrency] : ''), 'text');
      setIfBetter_(o, 'buyerName',      (iBuyerName >= 0 ? row[iBuyerName] : ''), 'text');

      setIfBetter_(o, 'totalOrderOrig', (iTotalOrig >= 0 ? row[iTotalOrig] : 0), 'number');
      setIfBetter_(o, 'orderTotalEGP',  (iTotalEgp >= 0 ? row[iTotalEgp] : 0), 'number');
      setIfBetter_(o, 'shipUaeEg',      (iShipEg >= 0 ? row[iShipEg] : 0), 'number');
      setIfBetter_(o, 'customsEGP',     (iCustomsEgp >= 0 ? row[iCustomsEgp] : 0), 'number');
      setIfBetter_(o, 'landedCostEGP',  (iLandedEgp >= 0 ? row[iLandedEgp] : 0), 'number');
      setIfBetter_(o, 'unitLandedEGP',  (iUnitLandedEgp >= 0 ? row[iUnitLandedEgp] : 0), 'number');

      addNoteUnique_(o, (iNotes >= 0 ? row[iNotes] : ''));
    });

    const rows = Object.values(orders)
      .sort(function (a, b) {
        const da = (a.orderDate instanceof Date) ? a.orderDate.getTime() : Number.POSITIVE_INFINITY;
        const db = (b.orderDate instanceof Date) ? b.orderDate.getTime() : Number.POSITIVE_INFINITY;
        if (da !== db) return da - db;
        return String(a.orderId).localeCompare(String(b.orderId));
      })
      .map(function (o) {
        return [
          o.orderId,
          (o.orderDate instanceof Date) ? o.orderDate : '',
          o.platform,
          o.sellerName,
          o.currency,
          o.buyerName,
          o.totalLines,
          o.totalQty,
          o.totalOrderOrig,
          o.orderTotalEGP,
          o.shipUaeEg,
          o.customsEGP,
          o.landedCostEGP,
          o.unitLandedEGP,
          o.notes
        ];
      });

    // Write output
    orders_clearOrdersData_(shO);

    if (rows.length) {
      shO.getRange(2, 1, rows.length, ORDER_HEADERS.length).setValues(rows);
    }

    // Formats
    const mapO = getHeaderMap_(shO, 1);
    orders_applyOrdersFormats_(shO, mapO, rows.length);

  } catch (e) {
    logError_('rebuildOrdersSummary', e);
    throw e;
  }
}

/**
 * Incremental Orders sync for specific Order IDs.
 * Used by AppCore sync queue processor.
 *
 * @param {string[]} orderIds
 */
function orders_syncFromPurchasesByOrderIds_(orderIds) {
  try {
    ensureErrorLog_();

    const ids = (orderIds || [])
      .map(function (x) { return String(x || '').trim(); })
      .filter(Boolean);

    // If empty, fall back to full rebuild when available.
    if (!ids.length) {
      if (typeof rebuildOrdersSummary === 'function') rebuildOrdersSummary();
      return;
    }

    const idSet = {};
    ids.forEach(function (x) { idSet[x] = true; });

    const shP = getSheet_(APP.SHEETS.PURCHASES);
    const shO = orders_ensureSheet_(APP.SHEETS.ORDERS);

    // Ensure Orders schema (non-destructive)
    try { setupOrdersLayout(); } catch (e) {}

    // Normalize Purchases headers (just in case)
    try { normalizeHeaders_(shP, 1); } catch (e) {}

    const mapP = getHeaderMap_(shP, 1);
    orders_assertPurchasesHeadersForOrders_(mapP);

    const lastRow = shP.getLastRow();
    const lastCol = shP.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return;

    const data = shP.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const idx = (h) => (mapP[h] ? (mapP[h] - 1) : -1);

    const iOrderId        = idx(APP.COLS.PURCHASES.ORDER_ID);
    const iOrderDate      = idx(APP.COLS.PURCHASES.ORDER_DATE);
    const iPlatform       = idx(APP.COLS.PURCHASES.PLATFORM);
    const iSellerName     = idx(APP.COLS.PURCHASES.SELLER);
    const iCurrency       = idx(APP.COLS.PURCHASES.CURRENCY);
    const iBuyerName      = idx(APP.COLS.PURCHASES.BUYER_NAME);
    const iQty            = idx(APP.COLS.PURCHASES.QTY);
    const iNotes          = idx(APP.COLS.PURCHASES.NOTES);

    const iTotalOrig      = idx(APP.COLS.PURCHASES.TOTAL_ORIG);
    const iTotalEgp       = idx(APP.COLS.PURCHASES.TOTAL_EGP);
    const iShipEg         = idx(APP.COLS.PURCHASES.SHIP_EG);
    const iCustomsEgp     = idx(APP.COLS.PURCHASES.CUSTOMS_EGP);
    const iLandedEgp      = idx(APP.COLS.PURCHASES.LANDED_COST);
    const iUnitLandedEgp  = idx(APP.COLS.PURCHASES.UNIT_LANDED);

    const setIfBetter_ = (obj, key, val, mode) => {
      if (val == null) return;

      if (mode === 'text') {
        const s = String(val || '').trim();
        if (!s) return;
        if (!obj[key]) obj[key] = s;
        return;
      }

      if (mode === 'date') {
        if (val instanceof Date) {
          if (!obj[key]) obj[key] = val;
        }
        return;
      }

      // number: take max to avoid double counting repeated per-line totals
      const n = Number(val);
      if (!isFinite(n)) return;
      if (obj[key] == null || Number(obj[key]) < n) obj[key] = n;
    };

    const addNoteUnique_ = (obj, note) => {
      const s = String(note || '').trim();
      if (!s) return;
      if (!obj._noteSet) obj._noteSet = {};
      if (obj._noteSet[s]) return;
      obj._noteSet[s] = true;
      obj.notes = obj.notes ? (obj.notes + ' | ' + s) : s;
    };

    const ordersMap = {};

    data.forEach(function (row) {
      const orderId = (iOrderId >= 0) ? String(row[iOrderId] || '').trim() : '';
      if (!orderId) return;
      if (!idSet[orderId]) return;

      const qty = (iQty >= 0) ? (Number(row[iQty]) || 0) : 0;

      let o = ordersMap[orderId];
      if (!o) {
        o = ordersMap[orderId] = {
          orderId: orderId,
          orderDate: '',
          platform: '',
          sellerName: '',
          currency: '',
          buyerName: '',
          totalLines: 0,
          totalQty: 0,
          totalOrderOrig: 0,
          orderTotalEGP: 0,
          shipUaeEg: 0,
          customsEGP: 0,
          landedCostEGP: 0,
          unitLandedEGP: 0,
          notes: ''
        };
      }

      o.totalLines += 1;
      o.totalQty   += qty;

      setIfBetter_(o, 'orderDate',      (iOrderDate >= 0 ? row[iOrderDate] : ''), 'date');
      setIfBetter_(o, 'platform',       (iPlatform >= 0 ? row[iPlatform] : ''), 'text');
      setIfBetter_(o, 'sellerName',     (iSellerName >= 0 ? row[iSellerName] : ''), 'text');
      setIfBetter_(o, 'currency',       (iCurrency >= 0 ? row[iCurrency] : ''), 'text');
      setIfBetter_(o, 'buyerName',      (iBuyerName >= 0 ? row[iBuyerName] : ''), 'text');

      setIfBetter_(o, 'totalOrderOrig', (iTotalOrig >= 0 ? row[iTotalOrig] : 0), 'number');
      setIfBetter_(o, 'orderTotalEGP',  (iTotalEgp >= 0 ? row[iTotalEgp] : 0), 'number');
      setIfBetter_(o, 'shipUaeEg',      (iShipEg >= 0 ? row[iShipEg] : 0), 'number');
      setIfBetter_(o, 'customsEGP',     (iCustomsEgp >= 0 ? row[iCustomsEgp] : 0), 'number');
      setIfBetter_(o, 'landedCostEGP',  (iLandedEgp >= 0 ? row[iLandedEgp] : 0), 'number');
      setIfBetter_(o, 'unitLandedEGP',  (iUnitLandedEgp >= 0 ? row[iUnitLandedEgp] : 0), 'number');

      addNoteUnique_(o, (iNotes >= 0 ? row[iNotes] : ''));
    });

    const outOrders = Object.keys(ordersMap).map(function (k) { return ordersMap[k]; });
    if (!outOrders.length) return;

    // Map existing Orders rows by Order ID
    const mapO = getHeaderMap_(shO, 1);
    const cOrder = mapO[APP.COLS.ORDERS.ORDER_ID] || 1;
    const cNotes = mapO['Notes'];

    const lastO = shO.getLastRow();
    const existingRowById = {};
    const existingNotesById = {};

    if (lastO >= 2) {
      const oidVals = shO.getRange(2, cOrder, lastO - 1, 1).getValues();
      const noteVals = cNotes ? shO.getRange(2, cNotes, lastO - 1, 1).getValues() : null;

      for (let i = 0; i < oidVals.length; i++) {
        const oid = String(oidVals[i][0] || '').trim();
        if (!oid) continue;
        existingRowById[oid] = i + 2;
        if (noteVals) existingNotesById[oid] = String(noteVals[i][0] || '').trim();
      }
    }

    const toUpdate = [];
    const toAppend = [];

    outOrders.forEach(function (o) {
      // Preserve existing notes if new notes are blank
      const existingNote = existingNotesById[o.orderId] || '';
      const finalNote = (String(o.notes || '').trim() || existingNote);

      const row = [
        o.orderId,
        o.orderDate,
        o.platform,
        o.sellerName,
        o.currency,
        o.buyerName,
        o.totalLines,
        o.totalQty,
        o.totalOrderOrig,
        o.orderTotalEGP,
        o.shipUaeEg,
        o.customsEGP,
        o.landedCostEGP,
        o.unitLandedEGP,
        finalNote
      ];

      const r = existingRowById[o.orderId];
      if (r) toUpdate.push({ row: r, values: row });
      else toAppend.push(row);
    });

    // Batched writes
    toUpdate.sort(function (a, b) { return a.row - b.row; });

    let i = 0;
    while (i < toUpdate.length) {
      const startRow = toUpdate[i].row;
      const batch = [toUpdate[i].values];
      let j = i + 1;
      let expected = startRow + 1;

      while (j < toUpdate.length && toUpdate[j].row === expected) {
        batch.push(toUpdate[j].values);
        expected++;
        j++;
      }

      shO.getRange(startRow, 1, batch.length, ORDER_HEADERS.length).setValues(batch);
      i = j;
    }

    if (toAppend.length) {
      const start = shO.getLastRow() + 1;
      shO.getRange(start, 1, toAppend.length, ORDER_HEADERS.length).setValues(toAppend);
    }

    // Formats (cheap; Orders sheet is small)
    try {
      const rowsCount = Math.max(0, shO.getLastRow() - 1);
      orders_applyOrdersFormats_(shO, getHeaderMap_(shO, 1), rowsCount);
    } catch (e) {}

  } catch (e) {
    logError_('orders_syncFromPurchasesByOrderIds_', e);
    throw e;
  }
}


/** ============================================================
 * Helpers
 * ============================================================ */

function orders_ensureSheet_(name) {
  if (typeof ensureSheet_ === 'function') return ensureSheet_(name);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function orders_removeFilterIfAny_(sh) {
  try {
    const f = sh.getFilter();
    if (f) f.remove();
  } catch (e) {}
}

function orders_applyHeaderStyle_(sh, headers) {
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);

  // Ensure header row values for the canonical range
  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setWrap(true)
    .setBackground('#e8f0fe');

  // Filter (safe)
  try {
    const r = sh.getRange(1, 1, 1, headers.length);
    if (!sh.getFilter()) r.createFilter();
  } catch (e) {
    // If an old filter exists with different range, rebuild it safely
    try {
      orders_removeFilterIfAny_(sh);
      sh.getRange(1, 1, 1, headers.length).createFilter();
    } catch (e2) {}
  }
}

function orders_clearOrdersData_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }
}

function orders_applyOrdersFormats_(sh, mapO, numRows) {
  const r = Math.max(0, numRows);
  if (r === 0) return;

  if (mapO['Order Date']) {
    sh.getRange(2, mapO['Order Date'], r, 1).setNumberFormat('yyyy-mm-dd');
  }

  const qtyCols = [
    mapO[APP.COLS.ORDERS.TOTAL_QTY],
    mapO[APP.COLS.ORDERS.TOTAL_LINES]
  ].filter(Boolean);

  qtyCols.forEach(function (c) {
    sh.getRange(2, c, r, 1).setNumberFormat('0');
  });

  const moneyHeaders = [
    APP.COLS.ORDERS.TOTAL_ORIG,
    APP.COLS.ORDERS.TOTAL_EGP,
    APP.COLS.ORDERS.SHIP_EG,
    APP.COLS.ORDERS.CUSTOMS,
    APP.COLS.ORDERS.LANDED_COST,
    APP.COLS.ORDERS.UNIT_LANDED
  ];

  moneyHeaders.forEach(function (h) {
    const col = mapO[h];
    if (col) sh.getRange(2, col, r, 1).setNumberFormat('0.00');
  });
}

function orders_assertPurchasesHeadersForOrders_(mapP) {
  const required = [
    APP.COLS.PURCHASES.ORDER_ID,
    APP.COLS.PURCHASES.ORDER_DATE,
    APP.COLS.PURCHASES.PLATFORM,
    APP.COLS.PURCHASES.SELLER,
    APP.COLS.PURCHASES.SKU,
    APP.COLS.PURCHASES.CURRENCY,
    APP.COLS.PURCHASES.BUYER_NAME,
    APP.COLS.PURCHASES.QTY,
    APP.COLS.PURCHASES.NOTES,
    APP.COLS.PURCHASES.TOTAL_ORIG,
    APP.COLS.PURCHASES.TOTAL_EGP,
    APP.COLS.PURCHASES.SHIP_EG,
    APP.COLS.PURCHASES.CUSTOMS_EGP,
    APP.COLS.PURCHASES.LANDED_COST,
    APP.COLS.PURCHASES.UNIT_LANDED
  ];

  const missing = required.filter(function (h) { return !mapP[h]; });
  if (missing.length) {
    throw new Error(
      'rebuildOrdersSummary: Missing columns in Purchases sheet: ' +
      missing.join(', ') +
      '. Run setupPurchasesLayout / installPurchasesFormulas.'
    );
  }
}

function orders_parseDate_(val) {
  if (val instanceof Date) return val;
  if (val === null || val === '') return '';
  // Try parse string
  if (typeof val === 'string') {
    const s = val.trim();
    if (!s) return '';
    const d = new Date(s);
    return isNaN(d.getTime()) ? '' : d;
  }
  // Numbers rarely appear here; keep safe
  const d2 = new Date(val);
  return isNaN(d2.getTime()) ? '' : d2;
}

/** Quick sanity test */
function testOrdersModule_() {
  try {
    const shP = getSheet_(APP.SHEETS.PURCHASES);
    const mapP = getHeaderMap_(shP, 1);
    orders_assertPurchasesHeadersForOrders_(mapP);

    setupOrdersLayout();
    rebuildOrdersSummary();

    orders_alert_('✅ Orders module basic test passed.');
  } catch (e) {
    logError_('testOrdersModule_', e);
    orders_alert_('❌ Orders module test failed: ' + e.message);
  }
}
