/** =============================================================
 * Purchases.gs – Purchases Sheet + Formulas Setup for CocoERP v2.2+
 * Depends on AppCore (kernel): APP, getSheet_, ensureSheet_,
 * ensureSheetSchema_, normalizeHeaders_, getHeaderMap_,
 * assertRequiredColumns_, colLetter_, logError_, ensureErrorLog_,
 * ensureSettingsSheet_, getDefaultFxAedEgp_, getDefaultShipUaeEgPerOrder_,
 * getDefaultCustomsPct_, getSettingsListByHeader_, (optional) clearSettingsCache_
 * ============================================================= */

// Purchases sheet header (canonical)
const PURCHASE_HEADERS = [
  'Order ID', 'Order Date', 'Platform', 'Seller Name', 'SKU', 'Batch Code',
  'Product Name', 'Variant / Color', 'Qty',

  'Unit Price (Orig)', 'Currency',
  'Subtotal (Orig)', 'Discount (Order)', 'Shipping Fee (Order)', 'Total Order (Orig)',
  'Final Unit Price',

  'Buyer Name', 'Buyer Phone', 'Buyer Address',

  'Payment Method', 'Payment Card Last4',
  'Invoice File ID', 'Invoice Link', 'Invoice Preview',

  'FX Rate → EGP', 'Order Total (EGP)', 'Ship UAE→EG (EGP)',
  'Customs/Fees %', 'Customs/Fees (EGP)',
  'Landed Cost (EGP)', 'Unit Landed Cost (EGP)', 'Notes',

  'Line Gross (Orig)', 'Discount Alloc (Orig)', 'Shipping Alloc (Orig)',
  'Line Net (Orig)', 'Net Unit Price (Orig)', 'Net Unit Price (EGP)',
  'Line ID'
];


/** =============================================================
 * onEdit (FAST) – Fill order-level defaults for edited rows only
 * - Does NOT override user-entered values
 * - Supports small multi-row pastes
 * ============================================================= */
function purchases_maybeAutoRepairFormulas_(sh, map) {
  try {
    if (!sh || sh.getLastRow() < 2) return;

    // Check a few ARRAYFORMULA anchors that must exist in row 2
    const requiredHeaders = [
      'Order Total (EGP)',
      'Landed Cost (EGP)',
      'Unit Landed Cost (EGP)',
      'Net Unit Price (EGP)',
      'Batch Code'
    ];

    let missing = false;
    for (let i = 0; i < requiredHeaders.length; i++) {
      const h = requiredHeaders[i];
      const c = map[h];
      if (!c) continue; // schema issue handled elsewhere
      const f = sh.getRange(2, c).getFormula();
      if (!f || String(f).trim() === '') {
        missing = true;
        break;
      }
    }
    if (!missing) return;

    // Throttle: do not run more than once per 60s
    const props = PropertiesService.getDocumentProperties();
    const key = 'CocoERP_PurchasesFormulasAutoRepair_LastRun';
    const last = Number(props.getProperty(key) || 0);
    const now = Date.now();
    if (now - last < 60 * 1000) return;

    withLock_('PurchasesAutoRepairFormulas', function () {
      const props2 = PropertiesService.getDocumentProperties();
      const last2 = Number(props2.getProperty(key) || 0);
      const now2 = Date.now();
      if (now2 - last2 < 60 * 1000) return;

      props2.setProperty(key, String(now2));

      if (typeof purchases_installFormulasCore_ === 'function') {
        purchases_installFormulasCore_();
      } else if (typeof installPurchasesFormulas === 'function') {
        installPurchasesFormulas();
      }
    });

  } catch (err) {
    try { logError_('purchases_maybeAutoRepairFormulas_', err); } catch (e) { }
  }
}

function purchases_ensureLineIds_(sh, map, startRow, n) {
  try {
    if (!sh || !map || !startRow || !n) return 0;

    const H = (APP && APP.COLS && APP.COLS.PURCHASES) ? APP.COLS.PURCHASES : {};
    const cLineId = map[H.LINE_ID] || map['Line ID'];
    const cOrder = map[H.ORDER_ID] || map['Order ID'];
    const cSku = map[H.SKU] || map['SKU'];

    if (!cLineId || !cOrder || !cSku) return 0;
    if (n <= 0) return 0;

    const orderIds = sh.getRange(startRow, cOrder, n, 1).getValues();
    const skus = sh.getRange(startRow, cSku, n, 1).getValues();
    const lineIds = sh.getRange(startRow, cLineId, n, 1).getValues();

    let changed = false;
    let count = 0;

    for (let i = 0; i < n; i++) {
      const cur = String(lineIds[i][0] || '').trim();
      if (cur) continue;

      const oid = String(orderIds[i][0] || '').trim();
      const sku = String(skus[i][0] || '').trim();
      if (!oid || !sku) continue;

      // Stable unique ID (won't change if rows move/sort)
      lineIds[i][0] = 'PL-' + Utilities.getUuid().slice(0, 8);
      changed = true;
      count++;
    }

    if (changed) {
      sh.getRange(startRow, cLineId, n, 1).setValues(lineIds);
    }
    return count;
  } catch (err) {
    try { logError_('purchases_ensureLineIds_', err); } catch (e) { }
    return 0;
  }
}

function purchasesOnEditDefaults_(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== APP.SHEETS.PURCHASES) return;

    const row0 = e.range.getRow();
    if (row0 < 2) return;

    const nr = e.range.getNumRows();
    const nc = e.range.getNumColumns();

    // Skip very large pastes (use menu: Repair Purchases Autofill)
    if (nr > 300 || nc > 25) return;

    const map = getHeaderMap_(sh, 1);
    try { purchases_maybeAutoRepairFormulas_(sh, map); } catch (e) { }

    const H = (APP && APP.COLS && APP.COLS.PURCHASES) ? APP.COLS.PURCHASES : {};
    const cOrder = map[H.ORDER_ID] || map['Order ID'];
    const cCurr = map[H.CURRENCY] || map['Currency'];
    const cFx = map[H.FX_RATE] || map['FX Rate → EGP'];
    const cShipEg = map[H.SHIP_EG] || map['Ship UAE→EG (EGP)'];
    const cCustoms = map[H.CUSTOMS_PCT] || map['Customs/Fees %'];

    if (!cOrder) return;

    const startRow = Math.max(2, row0);
    const endRow = Math.min(sh.getLastRow(), row0 + nr - 1);
    const n = endRow - startRow + 1;
    if (n <= 0) return;
    try { purchases_ensureLineIds_(sh, map, startRow, n); } catch (e) { }

    const orderIds = sh.getRange(startRow, cOrder, n, 1).getValues();

    const currVals = cCurr ? sh.getRange(startRow, cCurr, n, 1).getValues() : null;
    const fxVals = cFx ? sh.getRange(startRow, cFx, n, 1).getValues() : null;
    const shipVals = cShipEg ? sh.getRange(startRow, cShipEg, n, 1).getValues() : null;
    const customsVals = cCustoms ? sh.getRange(startRow, cCustoms, n, 1).getValues() : null;

    const defCurrency = (typeof getDefaultCurrency_ === 'function') ? String(getDefaultCurrency_() || '').trim().toUpperCase() : '';
    const defFxAed = (typeof getDefaultFxAedEgp_ === 'function') ? Number(getDefaultFxAedEgp_()) : 0;
    const defFxCny = (typeof getDefaultFxRate_ === 'function') ? Number(getDefaultFxRate_()) : 0; // CNY→EGP
    const defShip = (typeof getDefaultShipUaeEgPerOrder_ === 'function') ? Number(getDefaultShipUaeEgPerOrder_()) : 0;
    const defCustoms = (typeof getDefaultCustomsPct_ === 'function') ? Number(getDefaultCustomsPct_()) : 0;

    const isBlank = (v) => v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

    let currChanged = false, fxChanged = false, shipChanged = false, customsChanged = false;

    for (let i = 0; i < n; i++) {
      const oid = String(orderIds[i][0] || '').trim();
      if (!oid) continue;

      // Currency
      let cur = '';
      if (currVals) {
        cur = String(currVals[i][0] || '').trim().toUpperCase();
        if (!cur && defCurrency) {
          currVals[i][0] = defCurrency;
          cur = defCurrency;
          currChanged = true;
        }
      }

      // FX (currency-aware)
      if (fxVals && isBlank(fxVals[i][0])) {
        const c = cur || defCurrency;
        let fx = 0;
        if (c === 'EGP') fx = 1;
        else if (c === 'AED') fx = defFxAed;
        else if (c === 'CNY') fx = defFxCny;
        else fx = defFxAed || defFxCny || 0;

        if (isFinite(fx) && fx > 0) {
          fxVals[i][0] = fx;
          fxChanged = true;
        }
      }

      // Ship (per order)
      if (shipVals && isBlank(shipVals[i][0]) && isFinite(defShip) && defShip !== 0) {
        shipVals[i][0] = defShip;
        shipChanged = true;
      }

      // Customs %
      if (customsVals && isBlank(customsVals[i][0]) && isFinite(defCustoms) && defCustoms !== 0) {
        customsVals[i][0] = defCustoms;
        customsChanged = true;
      }
    }

    if (currChanged && cCurr) sh.getRange(startRow, cCurr, n, 1).setValues(currVals);
    if (fxChanged && cFx) sh.getRange(startRow, cFx, n, 1).setValues(fxVals);
    if (shipChanged && cShipEg) sh.getRange(startRow, cShipEg, n, 1).setValues(shipVals);
    if (customsChanged && cCustoms) sh.getRange(startRow, cCustoms, n, 1).setValues(customsVals);

  } catch (err) {
    logError_('purchasesOnEditDefaults_', err, {
      sheet: e && e.range && e.range.getSheet && e.range.getSheet().getName(),
      a1: e && e.range && e.range.getA1Notation && e.range.getA1Notation()
    });
  }
}

/** =============================================================
 * SAFE Setup (does NOT clear data)
 * ============================================================= */
function setupPurchasesLayout() {
  try {
    ensureErrorLog_();

    const sh = ensureSheet_(APP.SHEETS.PURCHASES);

    // Normalize headers (aliases like "Product" -> "Product Name")
    try { normalizeHeaders_(sh, 1); } catch (e) { }

    // Ensure Settings structure (non-destructive)
    if (typeof ensureSettingsSheet_ === 'function') {
      ensureSettingsSheet_();
    } else {
      const setSh = ensureSheet_(APP.SHEETS.SETTINGS);
      if (setSh.getLastRow() === 0) {
        setSh.getRange('A1:B1').setValues([['Setting', 'Value']]);
        setSh.getRange('D1:G1').setValues([['Platforms', 'Payment Methods', 'Currencies', 'Stores (optional)']]);
      }
    }

    // Ensure Purchases schema (add missing columns, keep existing data)
    ensureSheetSchema_(APP.SHEETS.PURCHASES, PURCHASE_HEADERS, { addMissing: true, headerRow: 1 });

    // ✅ SKU backfill right after schema (prevents Orders empty runs)
    purchases_backfillSkuSafe_();
    // Basic view config
    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);

    // Header styling (only row 1)
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setFontWeight('bold')
      .setWrap(true)
      .setBackground('#f1f3f4');

    // Filter (safe)
    try {
      const lastCol = sh.getLastColumn();
      const r = sh.getRange(1, 1, 1, lastCol);
      if (!sh.getFilter()) r.createFilter();
    } catch (e) { }

    const map = getHeaderMap_(sh, 1);
    try {
      const n = Math.max(0, sh.getLastRow() - 1);
      if (n > 0) purchases_ensureLineIds_(sh, map, 2, n);
    } catch (e) { }

    applyPurchasesFormats_(sh, map);
    applyPurchasesValidations_(sh, map);

    // Backfill defaults WITHOUT overriding manual edits
    purchases_backfillDefaults_();

    // Install formulas
    installPurchasesFormulas();

    // Enqueue Orders sync (if queue is enabled)
    try { if (typeof coco_enqueueOrdersSync_ === 'function') coco_enqueueOrdersSync_(null, { forceAll: true }); } catch (e) { }

    safeAlert_('✅ Purchases layout ensured (بدون مسح بيانات).');
  } catch (e) {
    logError_('setupPurchasesLayout', e);
    throw e;
  }
}

/**
 * Purchases onEdit hook (row-level, fast):
 * - Applies order-level defaults ONLY for the edited row(s)
 * - Avoids full-sheet backfills on every edit (prevents hang)
 */
function purchasesOnEdit_(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== APP.SHEETS.PURCHASES) return;

    const nr = e.range.getNumRows();
    const nc = e.range.getNumColumns();

    // Skip huge pastes (menu repair/backfill is safer)
    if (nr > 300 || nc > 25) return;

    // Fast, batched defaults + SKU
    if (typeof purchasesOnEditDefaults_ === 'function') purchasesOnEditDefaults_(e);
    if (typeof purchasesOnEditSku_ === 'function') purchasesOnEditSku_(e);

  } catch (err) {
    logError_('purchasesOnEdit_', err, {
      sheet: e && e.range && e.range.getSheet().getName(),
      a1: e && e.range && e.range.getA1Notation()
    });
  }
}

/** =============================================================
 * Auto-refresh defaults + SKU when editing Purchases
 * ============================================================= */
/** purchasesOnEditSku_ moved to SkuUtils.gs (row-level, faster) */


/** =============================================================
 * HARD RESET (explicit destructive action)
 * ============================================================= */
function setupPurchasesLayoutHardReset() {
  try {
    ensureErrorLog_();

    const sh = ensureSheet_(APP.SHEETS.PURCHASES);
    const setSh = ensureSheet_(APP.SHEETS.SETTINGS);

    // Confirm (manual UI only). In triggers/automation this returns false safely.
    if (!safeConfirm_('تحذير', 'ده هيمسح Purchases و Settings بالكامل. متأكد؟')) return;

    // Reset Settings
    setSh.clear();
    setSh.getRange('A1:B1').setValues([['Setting', 'Value']]);
    setSh.getRange('A2:B6').setValues([
      ['Default FX AED→EGP', 0],
      ['Default Ship UAE→EG (EGP) / order', 0],
      ['Default Customs % (e.g. 0.20 = 20%)', 0],
      ['—', '—'],
      ['ملاحظات', 'يمكنك تعديل القوائم يمين الصفحة']
    ]);

    // IMPORTANT: Match kernel naming (and we support legacy too)
    setSh.getRange('D1:G1').setValues([['Platforms', 'Payment Methods', 'Currencies', 'Stores (optional)']]);
    setSh.getRange('D2:D6').setValues([['AliExpress'], ['Alibaba'], ['Shein'], ['Amazon'], ['Noon']]);
    setSh.getRange('E2:E7').setValues([['VISA'], ['MASTERCARD'], ['AMEX'], ['Apple Pay'], ['Google Pay'], ['Cash on Delivery']]);
    setSh.getRange('F2:F6').setValues([['AED'], ['USD'], ['EGP'], ['SAR'], ['EUR']]);

    // Clear settings cache if AppCore implements it
    try {
      if (typeof clearSettingsCache_ === 'function') clearSettingsCache_();
    } catch (e) { }

    // Reset Purchases (remove filter safely BEFORE/AFTER clear)
    purchases_removeFilterIfAny_(sh);
    sh.clear();
    purchases_removeFilterIfAny_(sh);

    sh.setFrozenRows(1);
    sh.setFrozenColumns(3);

    sh.getRange(1, 1, 1, PURCHASE_HEADERS.length)
      .setValues([PURCHASE_HEADERS])
      .setFontWeight('bold')
      .setWrap(true)
      .setBackground('#f1f3f4');

    // Create filter safely
    try { sh.getRange(1, 1, 1, PURCHASE_HEADERS.length).createFilter(); } catch (e) { }

    const map = getHeaderMap_(sh, 1);
    applyPurchasesFormats_(sh, map);
    applyPurchasesValidations_(sh, map);

    // (في الهارد ريست غالبًا مفيش داتا، بس بنحطها كـ safety)
    purchases_backfillSkuSafe_();
    purchases_backfillDefaults_();
    installPurchasesFormulas();

    safeAlert_('✅ HARD RESET done.');
  } catch (e) {
    logError_('setupPurchasesLayoutHardReset', e);
    throw e;
  }
}

function purchases_removeFilterIfAny_(sh) {
  try {
    const f = sh.getFilter();
    if (f) f.remove();
  } catch (e) { }
}

/** =============================================================
 * SKU backfill (safe): calls SkuUtils module if present
 * ============================================================= */
function purchases_backfillSkuSafe_() {
  try {
    if (typeof sku_backfillPurchasesSku === 'function') {
      sku_backfillPurchasesSku();
      return;
    }
    if (typeof backfillPurchasesSku === 'function') { // legacy name
      backfillPurchasesSku();
      return;
    }
  } catch (e) {
    logError_('purchases_backfillSkuSafe_', e);
  }
}


/** =============================================================
 * Defaults backfill (NO override)
 *  - Fill FX / Ship / Customs% only when Order ID exists and cell is blank
 *  - ONLY first line per Order ID
 * ============================================================= */
function purchases_backfillOrderDefaults_() {
  const sh = getSheet_(APP.SHEETS.PURCHASES);
  const map = getHeaderMap_(sh, 1);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const cOrder = map['Order ID'];
  if (!cOrder) return;

  const cCurr = map['Currency'];
  const cFx = map['FX Rate → EGP'];
  const cShipEg = map['Ship UAE→EG (EGP)'];
  const cCustoms = map['Customs/Fees %'];

  const orderIds = sh.getRange(2, cOrder, n, 1).getValues();
  const curr = cCurr ? sh.getRange(2, cCurr, n, 1).getValues() : Array.from({ length: n }, () => ['']);

  const fxVals = cFx ? sh.getRange(2, cFx, n, 1).getValues() : null;
  const shipVals = cShipEg ? sh.getRange(2, cShipEg, n, 1).getValues() : null;
  const customsVals = cCustoms ? sh.getRange(2, cCustoms, n, 1).getValues() : null;

  const defFxAed = (typeof getDefaultFxAedEgp_ === 'function') ? Number(getDefaultFxAedEgp_()) : 0;
  const defFxCny = (typeof getDefaultFxRate_ === 'function') ? Number(getDefaultFxRate_()) : 0; // CNY→EGP
  const defShip = (typeof getDefaultShipUaeEgPerOrder_ === 'function') ? Number(getDefaultShipUaeEgPerOrder_()) : 0;
  const defCustoms = (typeof getDefaultCustomsPct_ === 'function') ? Number(getDefaultCustomsPct_()) : 0;

  const isBlank = (v) => v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '');

  const seen = new Set();
  let fxChanged = false, shipChanged = false, customsChanged = false;

  for (let i = 0; i < n; i++) {
    const oid = String(orderIds[i][0] || '').trim();
    if (!oid) continue;
    if (seen.has(oid)) continue; // only first line per order
    seen.add(oid);

    const cur = String(curr[i][0] || '').trim().toUpperCase();

    if (fxVals && isBlank(fxVals[i][0])) {
      const c = cur || ((typeof getDefaultCurrency_ === 'function') ? String(getDefaultCurrency_() || '').trim().toUpperCase() : '');
      let fx = 0;
      if (c === 'EGP') fx = 1;
      else if (c === 'AED') fx = defFxAed;
      else if (c === 'CNY') fx = defFxCny;
      else fx = defFxAed || defFxCny || 0;
      fxVals[i][0] = fx;
      fxChanged = true;
    }
    if (shipVals && isBlank(shipVals[i][0])) {
      shipVals[i][0] = defShip;
      shipChanged = true;
    }
    if (customsVals && isBlank(customsVals[i][0])) {
      customsVals[i][0] = defCustoms; // ex: 0.20
      customsChanged = true;
    }
  }

  if (fxChanged && cFx) sh.getRange(2, cFx, n, 1).setValues(fxVals);
  if (shipChanged && cShipEg) sh.getRange(2, cShipEg, n, 1).setValues(shipVals);
  if (customsChanged && cCustoms) sh.getRange(2, cCustoms, n, 1).setValues(customsVals);
}

// Backward-compatible alias
function purchases_backfillDefaults_() {
  return purchases_backfillOrderDefaults_();
}

/** =============================================================
 * Formats
 * ============================================================= */
function applyPurchasesFormats_(sh, map) {
  const maxRows = sh.getMaxRows();
  if (maxRows <= 1) return;

  if (map['Order Date']) sh.getRange(2, map['Order Date'], maxRows - 1, 1).setNumberFormat('yyyy-mm-dd');
  if (map['Qty']) sh.getRange(2, map['Qty'], maxRows - 1, 1).setNumberFormat('0');

  const moneyFields = [
    'Unit Price (Orig)', 'Subtotal (Orig)', 'Discount (Order)', 'Shipping Fee (Order)',
    'Total Order (Orig)', 'Final Unit Price',
    'Order Total (EGP)', 'Customs/Fees (EGP)',
    'Landed Cost (EGP)', 'Unit Landed Cost (EGP)',
    'Line Gross (Orig)', 'Discount Alloc (Orig)', 'Shipping Alloc (Orig)',
    'Line Net (Orig)', 'Net Unit Price (Orig)', 'Net Unit Price (EGP)',
    'Ship UAE→EG (EGP)'
  ];

  moneyFields.forEach(function (h) {
    if (map[h]) sh.getRange(2, map[h], maxRows - 1, 1).setNumberFormat('0.00');
  });

  if (map['FX Rate → EGP']) sh.getRange(2, map['FX Rate → EGP'], maxRows - 1, 1).setNumberFormat('0.0000');
  if (map['Customs/Fees %']) sh.getRange(2, map['Customs/Fees %'], maxRows - 1, 1).setNumberFormat('0.00%');
}

/** =============================================================
 * Validations (robust against Settings header naming differences)
 * ============================================================= */
function applyPurchasesValidations_(sh, map) {
  const maxRows = sh.getMaxRows();
  if (maxRows <= 1) return;

  const setSh = ensureSheet_(APP.SHEETS.SETTINGS);

  let platforms = [], payMethods = [], currencies = [];
  try {
    platforms = getSettingsListByHeader_('Platforms') || [];
    payMethods = getSettingsListByHeader_('Payment Methods') || [];
    if (!payMethods.length) payMethods = getSettingsListByHeader_('Payment Method') || [];
    currencies = getSettingsListByHeader_('Currencies') || [];
  } catch (e) { }

  const dvList_ = function (arr, rangeA1) {
    if (arr && arr.length) {
      return SpreadsheetApp.newDataValidation().requireValueInList(arr, true).build();
    }
    return SpreadsheetApp.newDataValidation().requireValueInRange(setSh.getRange(rangeA1), true).build();
  };

  const dvPlatform = dvList_(platforms, 'D2:D');
  const dvPay = dvList_(payMethods, 'E2:E');
  const dvCurr = dvList_(currencies, 'F2:F');

  if (map['Platform']) sh.getRange(2, map['Platform'], maxRows - 1, 1).setDataValidation(dvPlatform);
  if (map['Payment Method']) sh.getRange(2, map['Payment Method'], maxRows - 1, 1).setDataValidation(dvPay);
  if (map['Currency']) sh.getRange(2, map['Currency'], maxRows - 1, 1).setDataValidation(dvCurr);
}

/** =============================================================
 * Install formulas (public)
 * ============================================================= */

/**
 * One-click deterministic repair:
 * - Ensures Purchases schema + ARRAYFORMULAs
 * - Backfills SKU + order defaults (FX/ship/customs)
 * - Safe to run repeatedly (idempotent)
 */
function purchases_repairAutofill() {
  return withLock_('purchases_repairAutofill', function () {
    ensureErrorLog_();
    const res = purchases_installFormulasCore_();
    safeAlert_('✅ Purchases autofill repaired.');
    return res;
  });
}

function installPurchasesFormulas() {
  return purchases_installFormulasCore_();
}

// Backward-compatible alias
function _installPurchaseFormulasCore() {
  return purchases_installFormulasCore_();
}

/** =============================================================
 * Core formulas
 * ============================================================= */

/** =============================================================
 * Clear computed columns so ARRAYFORMULA can expand (safe)
 * ============================================================= */
function purchases_clearComputedColumns_(sh, map) {
  const maxRows = sh.getMaxRows();
  if (maxRows <= 1) return;

  // Columns that MUST be formula-driven (safe to clear)
  const computed = [
    'Invoice Preview',
    'Subtotal (Orig)',
    'Total Order (Orig)',
    'Final Unit Price',
    'Order Total (EGP)',
    'Customs/Fees (EGP)',
    'Landed Cost (EGP)',
    'Unit Landed Cost (EGP)',
    'Line Gross (Orig)',
    'Discount Alloc (Orig)',
    'Shipping Alloc (Orig)',
    'Line Net (Orig)',
    'Net Unit Price (Orig)',
    'Net Unit Price (EGP)',
    'Batch Code'
  ];

  computed.forEach(function (h) {
    const c = map[h];
    if (c) sh.getRange(2, c, maxRows - 1, 1).clearContent();
  });
}

function purchases_installFormulasCore_() {
  try {
    const sh = getSheet_(APP.SHEETS.PURCHASES);

    try { normalizeHeaders_(sh, 1); } catch (e) { }
    ensureSheetSchema_(APP.SHEETS.PURCHASES, PURCHASE_HEADERS, { addMissing: true, headerRow: 1 });

    // ✅ SKU backfill قبل ما نعمل Batch Code (سلسلة النظام بتتبني عليه)
    purchases_backfillSkuSafe_();
    const map = assertRequiredColumns_(sh, PURCHASE_HEADERS);

    // Ensure ARRAYFORMULA can expand (clear computed columns only)
    purchases_clearComputedColumns_(sh, map);
    const L = (name) => colLetter_(map[name]);
    const R = (name) => `${L(name)}2:${L(name)}`;

    // Backfill defaults (first line per order) - does not override manual edits
    try { purchases_backfillOrderDefaults_(); } catch (e) { }

    // Helper: order-level scalar per row (first nonblank value in that column for the order)
    const ORDER_SCALAR = (valueHeader) =>
      `IFERROR(VLOOKUP(${R('Order ID')},
        FILTER({${R('Order ID')},${R(valueHeader)}}, ${R('Order ID')}<>"", ${R(valueHeader)}<>""),
        2, FALSE
      ), 0)`;

    const ORDER_DISCOUNT = ORDER_SCALAR('Discount (Order)');
    const ORDER_SHIPFEE = ORDER_SCALAR('Shipping Fee (Order)');
    const ORDER_FX = ORDER_SCALAR('FX Rate → EGP');
    const ORDER_SHIP_EG = ORDER_SCALAR('Ship UAE→EG (EGP)');
    const ORDER_CUSTOMS = ORDER_SCALAR('Customs/Fees %');

    const ORDER_GROSSSUM = `SUMIF(${R('Order ID')},${R('Order ID')},${R('Line Gross (Orig)')})`;
    const ORDER_QTYSUM = `SUMIF(${R('Order ID')},${R('Order ID')},${R('Qty')})`;

    // 1) Invoice Preview
    sh.getRange(2, map['Invoice Preview']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        IF(LEN(${R('Invoice File ID')})>0,
          IFERROR(IMAGE("https://drive.google.com/uc?export=view&id="&${R('Invoice File ID')}),""),
          IF(LEN(${R('Invoice Link')})>0,
            IFERROR(IMAGE(${R('Invoice Link')}),""),""
          )
        )
      ))`
    );

    // 2) Subtotal (Orig) = Unit Price * Qty
    sh.getRange(2, map['Subtotal (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Unit Price (Orig)')}) * N(${R('Qty')})
      ))`
    );

    // 3) Line Gross (Orig)
    sh.getRange(2, map['Line Gross (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Unit Price (Orig)')}) * N(${R('Qty')})
      ))`
    );

    // 4) Total Order (Orig)
    sh.getRange(2, map['Total Order (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        SUMIF(${R('Order ID')},${R('Order ID')},${R('Subtotal (Orig)')})
        - N(${ORDER_DISCOUNT})
        + N(${ORDER_SHIPFEE})
      ))`
    );

    // 5) Discount Alloc (Orig)
    sh.getRange(2, map['Discount Alloc (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        IF(N(${ORDER_DISCOUNT})=0,0,
          IFERROR(N(${R('Line Gross (Orig)')})/N(${ORDER_GROSSSUM})*N(${ORDER_DISCOUNT}),0)
        )
      ))`
    );

    // 6) Shipping Alloc (Orig)
    sh.getRange(2, map['Shipping Alloc (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        IF(N(${ORDER_SHIPFEE})=0,0,
          IFERROR(N(${R('Line Gross (Orig)')})/N(${ORDER_GROSSSUM})*N(${ORDER_SHIPFEE}),0)
        )
      ))`
    );

    // 7) Line Net (Orig) = Gross - Discount + Shipping
    sh.getRange(2, map['Line Net (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Line Gross (Orig)')})
        - N(${R('Discount Alloc (Orig)')})
        + N(${R('Shipping Alloc (Orig)')})
      ))`
    );

    // 8) Net Unit Price (Orig)
    sh.getRange(2, map['Net Unit Price (Orig)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        IF(N(${R('Qty')})=0,"",
          N(${R('Line Net (Orig)')})/N(${R('Qty')})
        )
      ))`
    );

    // 9) Final Unit Price
    sh.getRange(2, map['Final Unit Price']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","", ${R('Net Unit Price (Orig)')} ))`
    );

    // 10) Order Total (EGP)
    sh.getRange(2, map['Order Total (EGP)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Total Order (Orig)')}) * N(${ORDER_FX})
      ))`
    );

    // 11) Customs/Fees (EGP)
    sh.getRange(2, map['Customs/Fees (EGP)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Order Total (EGP)')}) * N(${ORDER_CUSTOMS})
      ))`
    );

    // 12) Landed Cost (EGP)
    sh.getRange(2, map['Landed Cost (EGP)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Order Total (EGP)')})
        + N(${R('Customs/Fees (EGP)')})
        + N(${ORDER_SHIP_EG})
      ))`
    );

    // 13) Unit Landed Cost (EGP)
    sh.getRange(2, map['Unit Landed Cost (EGP)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        IF(N(${ORDER_QTYSUM})=0,"",
          N(${R('Landed Cost (EGP)')}) / N(${ORDER_QTYSUM})
        )
      ))`
    );

    // 14) Net Unit Price (EGP)
    sh.getRange(2, map['Net Unit Price (EGP)']).setFormula(
      `=ARRAYFORMULA(IF(${R('Order ID')}="","",
        N(${R('Net Unit Price (Orig)')}) * N(${ORDER_FX})
      ))`
    );

    // 15) Batch Code = OrderID || SKU (✅ منع missing data)
    sh.getRange(2, map['Batch Code']).setFormula(
      `=ARRAYFORMULA(IF((${R('Order ID')}="")+(${R('SKU')}=""),"",
        ${R('Order ID')} & "||" & ${R('SKU')}
      ))`
    );

    // Backfill order-level defaults AFTER formulas (non-destructive, fills blanks only)
    try { purchases_backfillDefaults_(); } catch (e) { }

    // Optional: enqueue Orders rebuild (if queue system exists)
    try { if (typeof coco_enqueueOrdersSync_ === 'function') coco_enqueueOrdersSync_(null, { forceAll: true }); } catch (e) { }

    SpreadsheetApp.flush();


  } catch (e) {
    logError_('purchases_installFormulasCore_', e);
    throw e;
  }
}

/** =============================================================
 * Quick sanity test
 * ============================================================= */
function testPurchasesModule_() {
  try {
    const sh = getSheet_(APP.SHEETS.PURCHASES);
    assertRequiredColumns_(sh, PURCHASE_HEADERS);
    purchases_backfillSkuSafe_();
    purchases_backfillDefaults_();
    installPurchasesFormulas();

    safeAlert_('✅ Purchases module basic test passed.');
  } catch (e) {
    logError_('testPurchasesModule_', e);
    safeAlert_('❌ Purchases module test failed: ' + e.message);
  }
}
