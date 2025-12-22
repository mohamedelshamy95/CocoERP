/** =============================================================
 * Logistics.gs – QC + Shipments (CN→UAE, UAE→EG) for CocoERP
 * Compatible with AppCore3
 * ============================================================= */

const QC_UAE_HEADERS = [
  'QC ID',

  APP.COLS.QC_UAE.ORDER_ID,
  APP.COLS.QC_UAE.SHIPMENT_ID,
  APP.COLS.QC_UAE.SKU,
  APP.COLS.QC_UAE.BATCH_CODE,

  (APP.COLS.PURCHASES.PRODUCT_NAME || APP.COLS.PURCHASES.PRODUCT || 'Product Name'),
  (APP.COLS.PURCHASES.VARIANT || 'Variant / Color'),

  'Qty Ordered',                       // optional business column (not in APP.COLS.QC_UAE)
  APP.COLS.QC_UAE.QTY_RECEIVED,        // canonical column (important)

  APP.COLS.QC_UAE.QTY_OK,
  APP.COLS.QC_UAE.QTY_MISSING,
  APP.COLS.QC_UAE.QTY_DEFECT,

  'QC Result',
  'QC Date',

  APP.COLS.QC_UAE.WAREHOUSE,
  APP.COLS.QC_UAE.PURCHASE_LINE_ID,
  APP.COLS.QC_UAE.NOTES
];

const SHIP_CN_UAE_HEADERS = [
  APP.COLS.SHIP_CN_UAE.SHIPMENT_ID,
  APP.COLS.SHIP_CN_UAE.SUPPLIER || 'Supplier / Factory',
  APP.COLS.SHIP_CN_UAE.FORWARDER || 'Forwarder',
  APP.COLS.SHIP_CN_UAE.TRACKING || 'Tracking / Container',
  'Purchases Line ID',

  'Order ID (Batch)',

  APP.COLS.SHIP_CN_UAE.SHIP_DATE,
  APP.COLS.SHIP_CN_UAE.ETA,
  APP.COLS.SHIP_CN_UAE.ARRIVAL,
  APP.COLS.SHIP_CN_UAE.STATUS,

  APP.COLS.PURCHASES.SKU,
  APP.COLS.PURCHASES.PRODUCT,
  APP.COLS.PURCHASES.VARIANT,
  APP.COLS.PURCHASES.QTY,

  'Gross Weight (kg)',
  'Volume (CBM)',

  APP.COLS.SHIP_CN_UAE.FREIGHT,
  APP.COLS.SHIP_CN_UAE.OTHER,
  APP.COLS.SHIP_CN_UAE.TOTAL_COST,

  APP.COLS.PURCHASES.NOTES
];

const SHIP_UAE_EG_HEADERS = [
  APP.COLS.SHIP_UAE_EG.SHIPMENT_ID,

  APP.COLS.SHIP_UAE_EG.FORWARDER,      // was 'Forwarder'
  APP.COLS.SHIP_UAE_EG.COURIER,        // was 'Courier'
  APP.COLS.SHIP_UAE_EG.AWB,            // was 'AWB / Tracking'

  APP.COLS.SHIP_UAE_EG.BOX_ID,
  APP.COLS.SHIP_UAE_EG.SHIP_DATE,
  APP.COLS.SHIP_UAE_EG.ETA,
  APP.COLS.SHIP_UAE_EG.ARRIVAL,
  APP.COLS.SHIP_UAE_EG.STATUS,

  APP.COLS.SHIP_UAE_EG.SKU,
  APP.COLS.PURCHASES.PRODUCT,
  APP.COLS.PURCHASES.VARIANT,

  APP.COLS.SHIP_UAE_EG.QTY,
  APP.COLS.SHIP_UAE_EG.QTY_SYNCED,

  APP.COLS.SHIP_UAE_EG.SHIP_COST,
  APP.COLS.SHIP_UAE_EG.CUSTOMS,
  APP.COLS.SHIP_UAE_EG.OTHER,
  APP.COLS.SHIP_UAE_EG.TOTAL_COST,


  APP.COLS.PURCHASES.NOTES
];

/** =======================
 * Public layout entry points
 * ======================= */

function setupQcLayout() {
  try {
    setupQC_UAE_();
    safeAlert_('تم تجهيز شيت QC_UAE ✔️');
  } catch (e) {
    logError_('setupQcLayout', e);
    throw e;
  }
}

function setupShipmentsLayouts() {
  try {
    setupShipmentsCnUae_();
    setupShipmentsUaeEg_();
    safeAlert_('تم تجهيز شيتات الشحن (CN→UAE + UAE→EG) ✔️');
  } catch (e) {
    logError_('setupShipmentsLayouts', e);
    throw e;
  }
}

function setupLogisticsLayout() {
  try {
    setupQC_UAE_();
    setupShipmentsCnUae_();
    setupShipmentsUaeEg_();
    safeAlert_('تم تجهيز شيتات اللوجستيك (QC + Shipments) ✔️');
  } catch (e) {
    logError_('setupLogisticsLayout', e);
    throw e;
  }
}

/** =======================
 * Layout implementations
 * ======================= */

function setupQC_UAE_() {
  const qcSh = (typeof getOrCreateSheet_ === 'function')
    ? getOrCreateSheet_(APP.SHEETS.QC_UAE)
    : _fallbackGetOrCreateSheet_(APP.SHEETS.QC_UAE);

  _setupSheetWithHeaders_(qcSh, QC_UAE_HEADERS);

  const qcMap = getHeaderMap_(qcSh);

  _applyDateFormatByHeaders_(qcSh, qcMap, ['QC Date']);
  _applyIntFormatByHeaders_(qcSh, qcMap, [
    'Qty Ordered',
    APP.COLS.QC_UAE.QTY_RECEIVED,
    APP.COLS.QC_UAE.QTY_OK,
    APP.COLS.QC_UAE.QTY_MISSING,
    APP.COLS.QC_UAE.QTY_DEFECT
  ]);
}

function setupShipmentsCnUae_() {
  const sh = (typeof getOrCreateSheet_ === 'function')
    ? getOrCreateSheet_(APP.SHEETS.SHIP_CN_UAE)
    : _fallbackGetOrCreateSheet_(APP.SHEETS.SHIP_CN_UAE);

  _setupSheetWithHeaders_(sh, SHIP_CN_UAE_HEADERS);

  const map = getHeaderMap_(sh);

  _applyDateFormatByHeaders_(sh, map, [
    APP.COLS.SHIP_CN_UAE.SHIP_DATE,
    APP.COLS.SHIP_CN_UAE.ETA,
    APP.COLS.SHIP_CN_UAE.ARRIVAL
  ]);

  _applyIntFormatByHeaders_(sh, map, [APP.COLS.PURCHASES.QTY]);

  _applyDecimalFormatByHeaders_(sh, map, [
    APP.COLS.SHIP_CN_UAE.FREIGHT,
    APP.COLS.SHIP_CN_UAE.OTHER,
    APP.COLS.SHIP_CN_UAE.TOTAL_COST,
    'Gross Weight (kg)',
    'Volume (CBM)'
  ]);
}

function setupShipmentsUaeEg_() {
  const sh = (typeof getOrCreateSheet_ === 'function')
    ? getOrCreateSheet_(APP.SHEETS.SHIP_UAE_EG)
    : _fallbackGetOrCreateSheet_(APP.SHEETS.SHIP_UAE_EG);

  _setupSheetWithHeaders_(sh, SHIP_UAE_EG_HEADERS);

  const map = getHeaderMap_(sh);

  _applyDateFormatByHeaders_(sh, map, [
    APP.COLS.SHIP_UAE_EG.SHIP_DATE,
    APP.COLS.SHIP_UAE_EG.ETA,
    APP.COLS.SHIP_UAE_EG.ARRIVAL
  ]);

  _applyIntFormatByHeaders_(sh, map, [
    APP.COLS.SHIP_UAE_EG.QTY,
    APP.COLS.SHIP_UAE_EG.QTY_SYNCED
  ]);

  _applyDecimalFormatByHeaders_(sh, map, [
    APP.COLS.SHIP_UAE_EG.SHIP_COST,
    APP.COLS.SHIP_UAE_EG.CUSTOMS,
    APP.COLS.SHIP_UAE_EG.OTHER,
    APP.COLS.SHIP_UAE_EG.TOTAL_COST
  ]);
}

/** =======================
 * Local helper (fallback only)
 * ======================= */

function _fallbackGetOrCreateSheet_(name) {
  const ss = getSpreadsheet_(); // from AppCore3
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Safe sheet setup:
 * - DOES NOT clear existing data by default (prevents accidental loss + reduces hang)
 * - Ensures headers exist (adds missing headers to the right)
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} headers
 * @param {Object=} opts { headerRow, freezeRows, freezeCols, clearData }
 */
function _setupSheetWithHeaders_(sheet, headers, opts) {
  const options = opts || {};
  const headerRow = options.headerRow || 1;
  const freezeRows = (options.freezeRows == null) ? 1 : Number(options.freezeRows);
  const freezeCols = Number(options.freezeCols || 0);
  const clearData = !!options.clearData;

  if (!sheet) throw new Error('_setupSheetWithHeaders_: missing sheet');
  if (!headers || !headers.length) throw new Error('_setupSheetWithHeaders_: missing headers');

  // Explicit destructive behavior only when requested
  if (clearData) {
    sheet.clear();
  }

  // Prefer central schema enforcement
  if (typeof ensureSheetSchema_ === 'function') {
    ensureSheetSchema_(sheet.getName(), headers, { addMissing: true, headerRow: headerRow });
  } else {
    // Fallback: write header row if empty
    const lastCol = Math.max(sheet.getLastColumn(), headers.length);
    sheet.getRange(headerRow, 1, 1, lastCol).clearContent();
    sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
  }

  // Header styling (lightweight)
  const hdr = sheet.getRange(headerRow, 1, 1, headers.length);
  hdr.setFontWeight('bold');

  // Freeze header row for usability
  try { sheet.setFrozenRows(Math.max(1, freezeRows)); } catch (e) { }
  if (freezeCols) {
    try { sheet.setFrozenColumns(Math.max(1, freezeCols)); } catch (e) { }
  }
}

/* format helpers unchanged (your versions are fine) */

function test_LogisticsSetup() {
  setupLogisticsLayout();
}
