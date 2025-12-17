/** =========================================================
 * AppCore.gs – Core Config & Helpers for CocoERP v2.2+
 * "System Kernel" version (stable + extensible)
 * =========================================================
 * Goals:
 *  - Single source of truth for Sheet names + Column headers
 *  - Schema validation + optional repair
 *  - Settings abstraction (FX/Shipping/Customs/etc.)
 *  - Safe triggers + locking + centralized error logging
 *
 * Notes:
 *  - This file owns the ONLY global onOpen(e) and onEdit(e).
 *  - Other modules must NOT define onOpen/onEdit.
 * ========================================================= */

/** ===================== GLOBAL CONFIG ===================== */
const APP = {
  VERSION: '2.2.1',

  BASE: {
    SPREADSHEET_ID: '1CVRbQbQrERmKSFg3jUi6xqLwJjZ4XsjUbhuMDIahCFk',
    TIMEZONE: Session.getScriptTimeZone() || 'Africa/Cairo'
  },


  // Warehouses (canonical codes + legacy aliases)
  // Canonical now used in Sheets:
  //  - UAE: KOR, ATTIA
  //  - EG : TAN-GH
  // Legacy aliases are kept for backward compatibility (older data / older modules).
  WAREHOUSES: {
    // Canonical
    KOR:     'KOR',
    ATTIA:   'ATTIA',
    TAN_GH:  'TAN-GH',

    // Legacy (kept)
    UAE_DXB:   'UAE-DXB',
    UAE_ATTIA: 'UAE-ATTIA',
    UAE_KOR:   'UAE-KOR',
    EG_CAI:    'EG-CAI',
    EG_TANTA:  'EG-TANTA'
  },

  WAREHOUSE_GROUPS: {
    UAE: ['KOR', 'ATTIA', 'UAE-DXB', 'UAE-ATTIA', 'UAE-KOR'],
    EG:  ['TAN-GH', 'EG-CAI', 'EG-TANTA']
  },

  SHEETS: {
    // Core
    PURCHASES: 'Purchases',
    SETTINGS:  'Settings',
    ORDERS:    'Orders',

    // Logistics
    QC_UAE:      'QC_UAE',
    SHIP_CN_UAE: 'Shipments_CN_UAE',
    SHIP_UAE_EG: 'Shipments_UAE_EG',

    // Inventory
    INVENTORY_TXNS: 'Inventory_Transactions',
    INVENTORY_UAE:  'Inventory_UAE',
    INVENTORY_EG:   'Inventory_EG',
    CATALOG_EG:     'Catalog_EG',
    SKU_RATES:      'SKU_Rates',

    // Sales
    SALES_EG: 'Sales_EG',

    // System
    ERROR_LOG: 'ErrorLog',
    DASHBOARD: 'Dashboard'
  },

  /**
   * Canonical column headers (single source of truth).
   * IMPORTANT: Headers are the actual sheet header strings.
   */
  COLS: {
    PURCHASES: {
      ORDER_ID:    'Order ID',
      ORDER_DATE:  'Order Date',
      PLATFORM:    'Platform',
      SELLER:      'Seller Name',
      SKU:         'SKU',

      // Canonical:
      PRODUCT_NAME:'Product Name',
      VARIANT:     'Variant / Color',

      // Backward-compat aliases (do NOT use in new code):
      PRODUCT:     'Product Name',       // legacy alias
      VARIANT_CLR: 'Variant / Color',    // legacy alias

      BATCH_CODE:  'Batch Code',
      QTY:         'Qty',
      TOTAL_ORIG:  'Total Order (Orig)',
      TOTAL_EGP:   'Order Total (EGP)',
      LANDED_COST: 'Landed Cost (EGP)',
      UNIT_LANDED: 'Unit Landed Cost (EGP)',
      CURRENCY:    'Currency',
      BUYER_NAME:  'Buyer Name',
      SHIP_EG:     'Ship UAE→EG (EGP)',
      CUSTOMS_EGP: 'Customs/Fees (EGP)',
      NOTES:       'Notes'
    },

    ORDERS: {
      ORDER_ID:    'Order ID',
      TOTAL_LINES: 'Total Lines',
      TOTAL_QTY:   'Total Qty',
      TOTAL_ORIG:  'Total Order (Orig)',
      TOTAL_EGP:   'Order Total (EGP)',
      SHIP_EG:     'Ship UAE→EG (EGP)',
      CUSTOMS:     'Customs/Fees (EGP)',
      LANDED_COST: 'Landed Cost (EGP)',
      PROFIT_EGP:  'Profit (EGP)',
      UNIT_LANDED: 'Unit Landed Cost (EGP)'
    },

    QC_UAE: {
      QC_ID:        'QC ID',
      ORDER_ID:     'Order ID',
      SHIPMENT_ID:  'Shipment CN→UAE ID',
      SKU:          'SKU',
      BATCH_CODE:   'Batch Code',
      PRODUCT_NAME: 'Product Name',
      VARIANT:      'Variant / Color',
      QTY_ORDERED:  'Qty Ordered',
      QTY_RECEIVED:'Qty Received',
      QTY_OK:       'Qty OK',
      QTY_MISSING:  'Qty Missing',
      QTY_DEFECT:   'Qty Defective',
      QC_RESULT:    'QC Result',
      QC_DATE:      'QC Date',
      WAREHOUSE:    'Warehouse (UAE)',
      NOTES:        'Notes'
    },

    SHIP_CN_UAE: {
      SHIPMENT_ID: 'Shipment ID',
      SUPPLIER:    'Supplier / Factory',
      FORWARDER:   'Forwarder',
      TRACKING:    'Tracking / Container',
      ORDER_BATCH: 'Order ID (Batch)',
      SHIP_DATE:   'Ship Date',
      ETA:         'ETA',
      ARRIVAL:     'Actual Arrival',
      STATUS:      'Status',
      SKU:         'SKU',
      PRODUCT_NAME:'Product Name',
      VARIANT:     'Variant / Color',
      QTY:         'Qty',
      WEIGHT:      'Gross Weight (kg)',
      VOLUME:      'Volume (CBM)',
      FREIGHT_AED: 'Freight (AED)',
      OTHER_AED:   'Other Fees (AED)',
      TOTAL_AED:   'Total Cost (AED)',
      NOTES:       'Notes'
    },

    SHIP_UAE_EG: {
      SHIPMENT_ID: 'Shipment ID',
      FORWARDER:   'Forwarder',
      COURIER:     'Courier',
      AWB:         'AWB / Tracking',
      BOX_ID:      'Box ID',
      SHIP_DATE:   'Ship Date',
      ETA:         'ETA',
      ARRIVAL:     'Actual Arrival',
      STATUS:      'Status',
      SKU:         'SKU',
      PRODUCT_NAME:'Product Name',
      VARIANT:     'Variant / Color',
      QTY:         'Qty',
      QTY_SYNCED:  'Qty Synced',
      SHIP_COST:   'Ship Cost (EGP) – per unit or box',
      CUSTOMS:     'Customs (EGP)',
      OTHER:       'Other (EGP)',
      TOTAL_COST:  'Total Cost (EGP)',
      NOTES:       'Notes'
    },

    INV_TXNS: {
      TXN_ID:       'Txn ID',
      TXN_DATE:     'Txn Date',
      SOURCE_TYPE:  'Source Type',
      SOURCE_ID:    'Source ID',
      BATCH_CODE:   'Batch Code',
      SKU:          'SKU',
      PRODUCT_NAME: 'Product Name',
      VARIANT:      'Variant / Color',
      WAREHOUSE:    'Warehouse',
      QTY_IN:       'Qty In',
      QTY_OUT:      'Qty Out',
      UNIT_COST:    'Unit Cost (EGP)',
      TOTAL_COST:   'Total Cost (EGP)',
      CURRENCY:     'Currency',
      UNIT_PRICE_ORIG:'Unit Price (Orig)',
      NOTES:        'Notes'
    },

    INV_UAE: {
      SKU:         'SKU',
      PRODUCT_NAME:'Product Name',
      VARIANT:     'Variant / Color',
      WAREHOUSE:   'Warehouse (UAE)',
      ON_HAND:     'On Hand Qty',
      ALLOCATED:   'Allocated Qty',
      AVAILABLE:   'Available Qty',
      AVG_COST:    'Avg Cost (EGP)',
      TOTAL_COST:  'Total Cost (EGP)',
      LAST_TXN:    'Last Txn Date',
      LAST_SRC_TYPE:'Last Source Type',
      LAST_SRC_ID: 'Last Source ID'
    },

    INV_EG: {
      SKU:         'SKU',
      PRODUCT_NAME:'Product Name',
      VARIANT:     'Variant / Color',
      WAREHOUSE:   'Warehouse (EG)',
      ON_HAND:     'On Hand Qty',
      ALLOCATED:   'Allocated Qty',
      AVAILABLE:   'Available Qty',
      AVG_COST:    'Avg Cost (EGP)',
      TOTAL_COST:  'Total Cost (EGP)',
      LAST_TXN:    'Last Txn Date',
      LAST_SRC_TYPE:'Last Source Type',
      LAST_SRC_ID: 'Last Source ID'
    },

    CATALOG_EG: {
      SKU:          'SKU',
      PRODUCT_NAME: 'Product Name',
      VARIANT:      'Variant / Color',
      COLOR_GROUP:  'Color Group',
      BRAND:        'Brand',
      CATEGORY:     'Category',
      SUBCATEGORY:  'Subcategory',
      STATUS:       'Status',
      DEFAULT_COST: 'Default Cost (EGP)',
      DEFAULT_PRICE:'Default Price (EGP)',
      BARCODE:      'Barcode',
      NOTES:        'Notes'
    },

    SALES_EG: {
      ORDER_ID:       'Order ID',
      ORDER_DATE:     'Order Date',
      PLATFORM:       'Platform',
      CUSTOMER_NAME:  'Customer Name',
      PHONE:          'Phone',
      CITY:           'City',
      ADDRESS:        'Address',
      SKU:            'SKU',
      PRODUCT_NAME:   'Product Name',
      VARIANT:        'Variant / Color',
      WAREHOUSE:      'Warehouse (EG)',
      QTY:            'Qty',
      UNIT_PRICE:     'Unit Price (EGP)',
      TOTAL_PRICE:    'Total Price (EGP)',
      DISCOUNT:       'Discount (EGP)',
      NET_REVENUE:    'Net Revenue (EGP)',
      SHIPPING_FEE:   'Shipping Fee (EGP)',
      PAYMENT_METHOD: 'Payment Method',
      ORDER_STATUS:   'Order Status',
      DELIVERED_DATE: 'Delivered Date',
      NOTES:          'Notes',
      SOURCE:         'Source',
      COURIER:        'Courier',
      AWB:            'AWB'
    }
  },

  /**
   * Header aliasing for migration (optional repair).
   * IMPORTANT: Keep aliases SAFE across sheets (do not add generic names like "Payment Method" here).
   */
  HEADER_ALIASES: {
    'Product': 'Product Name',
    'ProductName': 'Product Name',
    'Variant/Color': 'Variant / Color',
    'Variant': 'Variant / Color',

    // Dash variants / legacy headers
    'Ship Cost (EGP) - per unit or box': 'Ship Cost (EGP) – per unit or box',
    'Ship Cost (EGP) — per unit or box': 'Ship Cost (EGP) – per unit or box',
    'Ship Cost (EGP) — per unit': 'Ship Cost (EGP) – per unit',
    'Qty (pcs)': 'Qty',
// Arabic common variants (status)
    'تم التسليم': 'Delivered',
    'تم التوصيل': 'Delivered',
    'تم التسليم للعميل': 'Delivered'
  },

  SETTINGS_KEYS: {
    DEFAULT_CURRENCY: 'Default Currency',
    FX_CNY_EGP: 'FX CNY→EGP',
    FX_AED_EGP:        'Default FX AED→EGP',
    SHIP_UAE_EG_ORDER: 'Default Ship UAE→EG (EGP) / order',
    CUSTOMS_PCT:       'Default Customs % (e.g. 0.20 = 20%)'
  },

  SETTINGS_LIST_HEADERS: {
    PLATFORMS:        'Platforms',
    PAYMENT_METHODS:  'Payment Methods',
    CURRENCIES:       'Currencies',
    STORES:           'Stores (optional)',
    WAREHOUSES:       'Warehouses'
  },

  INTERNAL: {
    SETTINGS_CACHE_KEY: 'CocoERP_SettingsMap_v1',
    USE_INSTALLABLE_ONEDIT_PROP: 'CocoERP_UseInstallableOnEdit'
    ,
    // Orders sync queue (Purchases → Orders)
    ORDERS_SYNC_QUEUE_KEY: 'CocoERP_OrdersSyncQueue_v1',
    ORDERS_SYNC_ALL_FLAG:  'CocoERP_OrdersSyncAllFlag_v1',
    ORDERS_SYNC_LAST_RUN:  'CocoERP_OrdersSyncLastRun_v1',
    ORDERS_SYNC_LAST_ERROR:'CocoERP_OrdersSyncLastError_v1'
}
};

/** ===================== CORE HELPERS ===================== */

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  const id = (APP.BASE && APP.BASE.SPREADSHEET_ID) ? String(APP.BASE.SPREADSHEET_ID).trim() : '';
  if (id) return SpreadsheetApp.openById(id);

  throw new Error('No active spreadsheet and APP.BASE.SPREADSHEET_ID is empty/invalid.');
}

function getSheet_(name) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" not found.');
  return sh;
}

function ensureSheet_(name) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function getHeaderMap_(sheet, headerRow) {
  const row = headerRow || 1;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};
  const headers = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach(function (h, i) {
    const key = (h || '').toString().trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function colLetter_(colIndex) {
  let index = colIndex;
  let s = '';
  while (index > 0) {
    const mod = (index - 1) % 26;
    s = String.fromCharCode(65 + mod) + s;
    index = Math.floor((index - mod) / 26);
  }
  return s;
}

function resolveCol_(headerMap, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const k = candidates[i];
    if (headerMap[k]) return headerMap[k];
  }
  return 0;
}

function assertRequiredColumns_(sheet, requiredCols) {
  const map = getHeaderMap_(sheet);
  const missing = requiredCols.filter(function (c) { return !map[c]; });
  if (missing.length) {
    throw new Error('Missing columns in "' + sheet.getName() + '": ' + missing.join(', '));
  }
  return map;
}

function ensureSheetSchema_(sheetName, headers, opts) {
  const options = opts || {};
  const headerRow = options.headerRow || 1;
  const addMissing = options.addMissing !== false;

  const sh = ensureSheet_(sheetName);
  const lastRow = sh.getLastRow();

  if (lastRow === 0) {
    sh.getRange(headerRow, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  const map = getHeaderMap_(sh, headerRow);
  if (!addMissing) return sh;

  const missing = headers.filter(function (h) { return !map[h]; });
  if (missing.length) {
    const startCol = sh.getLastColumn() + 1;
    sh.getRange(headerRow, startCol, 1, missing.length).setValues([missing]);
  }

  return sh;
}

function normalizeHeaders_(sh, headerRow) {
  
  // Normalize common punctuation differences (en/em dash -> hyphen)
  const _canon_ = (v) => String(v || '').replace(/[\u2013\u2014]/g, '-').trim();
const row = headerRow || 1;
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return;

  const range = sh.getRange(row, 1, 1, lastCol);
  const headers = range.getValues()[0];

  let changed = false;
  for (let i = 0; i < headers.length; i++) {
    const raw = (headers[i] || '').toString().trim();
    const alias = APP.HEADER_ALIASES[raw];
    if (alias && alias !== raw) {
      headers[i] = alias;
      changed = true;
    }
  }

  if (changed) range.setValues([headers]);
}

/** ===================== ERROR LOGGING ===================== */

function ensureErrorLog_() {
  const sh = ensureSheet_(APP.SHEETS.ERROR_LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Function', 'Message', 'Stack', 'Context']);
  }
  return sh;
}

function logError_(fnName, error, context) {
  try {
    const sh = ensureErrorLog_();
    const timestamp = Utilities.formatDate(new Date(), APP.BASE.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    let ctxString = '';
    if (context) {
      try { ctxString = JSON.stringify(context); }
      catch (jsonErr) { ctxString = '<<context stringify failed: ' + (jsonErr && jsonErr.message) + '>>'; }
    }

    sh.appendRow([
      timestamp,
      fnName || '',
      (error && error.message) || String(error),
      (error && error.stack) || '',
      ctxString
    ]);
  } catch (e) {
    console.error('Failed to log error for ' + (fnName || '') + ': ' + (e && e.message ? e.message : e));
  }
}

/** ===================== LOCKING / SAFETY ===================== */

function withLock_(lockName, fn) {
  const lock = LockService.getDocumentLock();
  const ok = lock.tryLock(25 * 1000);
  if (!ok) throw new Error('Could not acquire lock: ' + (lockName || 'DocumentLock'));
  try { return fn(); }
  finally { lock.releaseLock(); }
}

/**
 * Normalize warehouse code:
 * - Trim + uppercase
 * - Convert whitespace/underscores to '-'
 * - Collapse multiple '-'
 */
function normalizeWarehouseCode_(wh) {
  if (!wh) return '';
  let s = String(wh).trim().toUpperCase();
  s = s.replace(/[_\s]+/g, '-');
  s = s.replace(/-+/g, '-').replace(/^-+|-+$/g, '');
  return s;
}

/**
 * Delivered status detector (Arabic + English).
 * Used by Sales sync + other flows.
 */
function isDeliveredStatus_(status) {
  if (!status) return false;
  const s = String(status).trim().toLowerCase();
  if (!s) return false;

  // Exact matches
  if (s === 'delivered') return true;
  if (s === 'تم التسليم') return true;
  if (s === 'تم التسليم للعميل') return true;
  if (s === 'تم التوصيل') return true;

  // Loose contains
  if (s.includes('deliv')) return true;
  if (s.includes('تسليم')) return true;
  if (s.includes('توصيل')) return true;

  return false;
}

function safeAlert_(msg) {
  try {
    SpreadsheetApp.getUi().alert(String(msg));
  } catch (e) {
    Logger.log(String(msg));
  }
}

/** ===================== SETTINGS LAYER ===================== */

function ensureSettingsSheet_() {
  const sh = ensureSheet_(APP.SHEETS.SETTINGS);

  // Always ensure header rows exist/correct (safe, idempotent).
  sh.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]);

  // Lists headers row (D1:...)
  sh.getRange(1, 4, 1, 5).setValues([[
    APP.SETTINGS_LIST_HEADERS.PLATFORMS,
    APP.SETTINGS_LIST_HEADERS.PAYMENT_METHODS,
    APP.SETTINGS_LIST_HEADERS.CURRENCIES,
    APP.SETTINGS_LIST_HEADERS.STORES,
    APP.SETTINGS_LIST_HEADERS.WAREHOUSES
  ]]);

  // Seed required Settings keys (non-destructive: fill only if missing)
  try {
    const map = getHeaderMap_(sh, 1);
    const keyCol = map['Setting'] || 1;
    const valCol = map['Value'] || 2;

    const lastRow = sh.getLastRow();
    const kv = (lastRow >= 2)
      ? sh.getRange(2, keyCol, lastRow - 1, 2).getValues()
      : [];

    const existing = {};
    kv.forEach(function (r) {
      const k = (r[0] || '').toString().trim();
      if (k) existing[k] = true;
    });

    const toAppend = [];
    function addIfMissing_(k, v) {
      if (!k) return;
      if (!existing[k]) {
        toAppend.push([k, v]);
        existing[k] = true;
      }
    }

    addIfMissing_(APP.SETTINGS_KEYS.DEFAULT_CURRENCY, 'AED');
    addIfMissing_(APP.SETTINGS_KEYS.FX_AED_EGP, 0);
    // Optional: keep for future if you ever capture USD rows; safe even if unused.
    if (APP.SETTINGS_KEYS.FX_USD_EGP) addIfMissing_(APP.SETTINGS_KEYS.FX_USD_EGP, 0);
    addIfMissing_(APP.SETTINGS_KEYS.DEFAULT_SHIP_UAE_EG, 0);
    addIfMissing_(APP.SETTINGS_KEYS.DEFAULT_CUSTOMS_PCT, 0.20);

    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, keyCol, toAppend.length, 2).setValues(toAppend);
    }
  } catch (e) {
    // ignore - settings sheet is still usable
  }

  // Seed Warehouses list (D lists area) – non-destructive, allow future additions.
  try {
    const listMap = getHeaderMap_(sh, 1);
    const whCol = listMap[APP.SETTINGS_LIST_HEADERS.WAREHOUSES];
    if (whCol) {
      const desired = ['KOR', 'ATTIA', 'TAN-GH', '']; // last blank row as "future slot"
      const last = sh.getLastRow();
      const existingVals = (last >= 2)
        ? sh.getRange(2, whCol, last - 1, 1).getValues().map(function (r) { return (r[0] || '').toString().trim(); })
        : [];
      const hasAny = existingVals.some(function (v) { return v; });
      if (!hasAny) {
        sh.getRange(2, whCol, desired.length, 1).setValues(desired.map(function (v) { return [v]; }));
      }
    }
  } catch (e) {}

  return sh;
}

function clearSettingsCache_() {
  try {
    CacheService.getDocumentCache().remove('CocoERP_SettingsMap_v1');
  } catch (e) {}
}

function getSettingsMap_() {
  const cache = CacheService.getDocumentCache();
  const cacheKey = APP.INTERNAL.SETTINGS_CACHE_KEY;

  const cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch (e) {}
  }

  const sh = ensureSettingsSheet_();
  const lastRow = sh.getLastRow();
  const map = {};

  if (lastRow >= 2) {
    const values = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    values.forEach(function (r) {
      const k = (r[0] || '').toString().trim();
      if (!k) return;
      map[k] = r[1];
    });
  }

  cache.put(cacheKey, JSON.stringify(map), 60); // 60s cache
  return map;
}

function getSetting_(key, defaultValue) {
  const map = getSettingsMap_();
  const v = map[key];
  return (v === undefined || v === null || v === '') ? defaultValue : v;
}

function getDefaultFxAedEgp_() {
  const v = getSetting_(APP.SETTINGS_KEYS.FX_AED_EGP, 0);
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function getDefaultCurrency_() {
  // Settings sheet first; fallback to CNY
  const v = getSetting_(APP.SETTINGS_KEYS.DEFAULT_CURRENCY, 'CNY');
  return String(v || 'CNY').trim().toUpperCase();
}

function getDefaultFxRate_() {
  // Settings sheet first; return null if not configured
  const v = getSetting_(APP.SETTINGS_KEYS.FX_CNY_EGP, null);
  const n = Number(v);
  if (!isFinite(n) || n <= 0) return null;
  return n;
}

function getDefaultShipUaeEgPerOrder_() {
  const v = getSetting_(APP.SETTINGS_KEYS.SHIP_UAE_EG_ORDER, 0);
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function getDefaultCustomsPct_() {
  const v = getSetting_(APP.SETTINGS_KEYS.CUSTOMS_PCT, 0);
  const n = Number(v);
  return isFinite(n) ? n : 0;
}

function getSettingsListByHeader_(headerName) {
  const sh = ensureSettingsSheet_();
  const map = getHeaderMap_(sh, 1);
  const col = map[headerName];
  if (!col) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const raw = sh.getRange(2, col, lastRow - 1, 1).getValues();
  const vals = raw.map(function (r) { return (r[0] || '').toString().trim(); })
    .filter(function (v) { return v; });

  // unique
  const seen = {};
  const out = [];
  vals.forEach(function (v) {
    if (!seen[v]) { seen[v] = true; out.push(v); }
  });
  return out;
}

/** ===================== INVENTORY TXN HEADER ===================== */

function ensureInventoryTxnHeader_() {
  const sh = ensureSheet_(APP.SHEETS.INVENTORY_TXNS);
  const lastRow = sh.getLastRow();

  if (lastRow === 0) {
    let headers;
    if (typeof INV_TXN_HEADERS !== 'undefined' && INV_TXN_HEADERS && INV_TXN_HEADERS.length) {
      headers = INV_TXN_HEADERS;
    } else {
      headers = Object.keys(APP.COLS.INV_TXNS).map(function (k) { return APP.COLS.INV_TXNS[k]; });
    }
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function clearInventoryTransactions_() {
  const sh = ensureSheet_(APP.SHEETS.INVENTORY_TXNS);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }
}

/** ===================== SYSTEM ORCHESTRATION ===================== */

function inv_fullRebuildFromLogistics() {
  return withLock_('inv_fullRebuildFromLogistics', function () {
    try {
      clearInventoryTransactions_();
      ensureInventoryTxnHeader_();
      Logger.log('Inventory Transactions cleared and header ensured.');

      if (typeof syncQCtoInventory_UAE !== 'function') throw new Error('syncQCtoInventory_UAE function not found.');
      if (typeof syncShipmentsUaeEgToInventory !== 'function') throw new Error('syncShipmentsUaeEgToInventory function not found.');
      if (typeof inv_rebuildAllSnapshots !== 'function') throw new Error('inv_rebuildAllSnapshots function not found.');

      syncQCtoInventory_UAE();
      syncShipmentsUaeEgToInventory();
      inv_rebuildAllSnapshots();

      safeAlert_('✅ Full Rebuild Done Successfully!');
    } catch (err) {
      logError_('inv_fullRebuildFromLogistics', err);
      safeAlert_('❌ Error during rebuild: ' + err.message);
      throw err;
    }
  });
}

function coco_preflightAndRepair() {
  return withLock_('coco_preflightAndRepair', function () {
    try {
      ensureErrorLog_();
      ensureSettingsSheet_();

      // Ensure sheets exist (create if missing)
      Object.keys(APP.SHEETS).forEach(function (k) {
        ensureSheet_(APP.SHEETS[k]);
      });

      // If sheets were DELETED, bootstrap layouts to restore headers.
      _coco_bootstrapLayoutsIfMissingHeaders_();

      // Normalize headers (safe) — exclude Settings (handled separately).
      Object.keys(APP.SHEETS).forEach(function (k) {
        const shName = APP.SHEETS[k];
        const sh = ensureSheet_(shName);
        if (shName !== APP.SHEETS.SETTINGS) normalizeHeaders_(sh, 1);
      });

      // Ensure schemas we strongly expect (post-bootstrap)
      ensureSheetSchema_(APP.SHEETS.INVENTORY_TXNS, Object.keys(APP.COLS.INV_TXNS).map(function (k) { return APP.COLS.INV_TXNS[k]; }), { addMissing: true });
      ensureInventoryTxnHeader_();

      ensureSheetSchema_(APP.SHEETS.INVENTORY_UAE, Object.keys(APP.COLS.INV_UAE).map(function (k) { return APP.COLS.INV_UAE[k]; }), { addMissing: true });
      ensureSheetSchema_(APP.SHEETS.INVENTORY_EG,  Object.keys(APP.COLS.INV_EG).map(function (k) { return APP.COLS.INV_EG[k]; }),  { addMissing: true });

      ensureSheetSchema_(APP.SHEETS.CATALOG_EG, Object.keys(APP.COLS.CATALOG_EG).map(function (k) { return APP.COLS.CATALOG_EG[k]; }), { addMissing: true });
      ensureSheetSchema_(APP.SHEETS.SALES_EG,   Object.keys(APP.COLS.SALES_EG).map(function (k) { return APP.COLS.SALES_EG[k]; }),   { addMissing: true });

      safeAlert_('✅ Preflight + Repair completed.');
    } catch (err) {
      logError_('coco_preflightAndRepair', err);
      safeAlert_('❌ Preflight failed: ' + err.message);
      throw err;
    }
  });
}

/**
 * Bootstrap layouts after the user DELETES sheets.
 * This restores headers + formats by calling module setup functions when available.
 * Idempotent: only runs a setup if the target sheet has no header row content.
 */
function _coco_bootstrapLayoutsIfMissingHeaders_() {
  // Purchases & Orders
  if (_sheetHeaderEmpty_(APP.SHEETS.PURCHASES) && typeof setupPurchasesLayout === 'function') setupPurchasesLayout();
  if (_sheetHeaderEmpty_(APP.SHEETS.PURCHASES) && typeof installPurchasesFormulas === 'function') installPurchasesFormulas();
  if (_sheetHeaderEmpty_(APP.SHEETS.ORDERS)    && typeof setupOrdersLayout === 'function')    setupOrdersLayout();

  // Logistics / QC / Shipments
  // Prefer the unified layout installer if available.
  if ((_sheetHeaderEmpty_(APP.SHEETS.QC_UAE) || _sheetHeaderEmpty_(APP.SHEETS.SHIP_CN_UAE) || _sheetHeaderEmpty_(APP.SHEETS.SHIP_UAE_EG)) &&
      typeof setupLogisticsLayout === 'function') {
    setupLogisticsLayout();
  } else {
    if (_sheetHeaderEmpty_(APP.SHEETS.QC_UAE) && typeof setupQcLayout === 'function') setupQcLayout();
    if ((_sheetHeaderEmpty_(APP.SHEETS.SHIP_CN_UAE) || _sheetHeaderEmpty_(APP.SHEETS.SHIP_UAE_EG)) &&
        typeof setupShipmentsLayouts === 'function') setupShipmentsLayouts();
  }

  // Inventory core
  if ((_sheetHeaderEmpty_(APP.SHEETS.INVENTORY_TXNS) || _sheetHeaderEmpty_(APP.SHEETS.INVENTORY_UAE) || _sheetHeaderEmpty_(APP.SHEETS.INVENTORY_EG)) &&
      typeof setupInventoryCoreLayout === 'function') setupInventoryCoreLayout();

  // Catalog & Sales
  if (_sheetHeaderEmpty_(APP.SHEETS.CATALOG_EG) && typeof setupCatalogEgLayout === 'function') setupCatalogEgLayout();
  if (_sheetHeaderEmpty_(APP.SHEETS.SALES_EG)   && typeof setupSalesLayout === 'function')     setupSalesLayout();
}

function _sheetHeaderEmpty_(sheetName, headerRow) {
  const sh = ensureSheet_(sheetName);
  const row = headerRow || 1;
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return true;
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i] || '').trim()) return false;
  }
  return true;
}

/** ===================== UI MENU ===================== */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  const mainMenu = ui.createMenu('CocoERP')
    .addSubMenu(
      ui.createMenu('System')
        .addItem('Preflight + Repair (Schemas/Headers)', 'coco_preflightAndRepair')
        .addSeparator()
        .addItem('Install Triggers (Recommended)', 'coco_installTriggers')
        .addItem('Uninstall Triggers', 'coco_uninstallTriggers')
        .addSeparator()
        .addItem('Run Core Tests', 'runCoreTests')
    )
    .addSubMenu(
      ui.createMenu('Purchases & Orders')
        .addItem('Setup Purchases Layout',     'setupPurchasesLayout')
        .addItem('Install Purchases Formulas', 'installPurchasesFormulas')
        .addSeparator()
        .addItem('Setup Orders Layout',        'setupOrdersLayout')
        .addItem('Rebuild Orders Summary',     'rebuildOrdersSummary')
        .addItem('Process Sync Queue Now',   'coco_processSyncQueueNow')
        .addItem('Debug Orders Sync Status', 'coco_debugOrdersSyncStatus')
    )
    .addSubMenu(
      ui.createMenu('Logistics & QC')
        .addItem('Setup QC Layout',               'setupQcLayout')
        .addItem('Setup Shipments Layouts',       'setupShipmentsLayouts')
        .addSeparator()
        .addItem('Generate QC from Purchases…',   'qc_generateFromPurchasesPrompt')
        .addItem('Recalc QC Quantities & Result', 'qc_recalcQuantitiesAndResult')
        .addItem('Backfill Purchases SKU',        'sku_backfillPurchasesSku')
    )
    .addSubMenu(
      ui.createMenu('Inventory')
        .addItem('Setup Inventory Core Layout', 'setupInventoryCoreLayout')
        .addItem('Rebuild Inventory Snapshots', 'inv_rebuildAllSnapshots')
        .addSeparator()
        .addItem('Sync QC → Inventory (UAE)',   'syncQCtoInventory_UAE')
        .addItem('Sync Shipments UAE→EG',       'syncShipmentsUaeEgToInventory')
        .addSeparator()
        .addItem('Full Rebuild from Logistics', 'inv_fullRebuildFromLogistics')
    );

  if (typeof setupSalesLayout === 'function' ||
      typeof syncSalesFromOrdersSheet === 'function' ||
      typeof syncSalesEgToInventory === 'function') {

    const salesMenu = ui.createMenu('Sales & Revenue');

    if (typeof setupSalesLayout === 'function') {
      salesMenu.addItem('Setup Sales Layout', 'setupSalesLayout');
    }
    if (typeof syncSalesFromOrdersSheet === 'function') {
      salesMenu.addItem('Sync from Orders sheet', 'syncSalesFromOrdersSheet');
    }
    if (typeof syncSalesEgToInventory === 'function') {
      salesMenu.addItem('Sync Sales → Inventory (EG)', 'syncSalesEgToInventory');
    }

    mainMenu.addSubMenu(salesMenu);
  }

  mainMenu.addToUi();
}

/** ===================== onEdit DISPATCHER ===================== */

function _useInstallableOnEditFlag_() {
  try {
    const p = PropertiesService.getDocumentProperties().getProperty(APP.INTERNAL.USE_INSTALLABLE_ONEDIT_PROP);
    return String(p || '') === '1';
  } catch (e) {
    return false;
  }
}

function _setUseInstallableOnEditFlag_(enabled) {
  const props = PropertiesService.getDocumentProperties();
  if (enabled) props.setProperty(APP.INTERNAL.USE_INSTALLABLE_ONEDIT_PROP, '1');
  else props.deleteProperty(APP.INTERNAL.USE_INSTALLABLE_ONEDIT_PROP);
}

function _dispatchOnEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const name  = sheet.getName();

  // Settings edits -> invalidate cache only (A/B area).
  // We keep this light to avoid expensive recalcs on every Settings tweak.
  if (name === APP.SHEETS.SETTINGS) {
    const col = e.range.getColumn();
    if (col === 1 || col === 2) clearSettingsCache_();
    return;
  }

    // Purchases: fast defaults + SKU + enqueue Orders sync (debounced).
  // IMPORTANT: Do not early-return after SKU only, otherwise FX/Ship/Customs & Orders sync will stop.
  if (name === APP.SHEETS.PURCHASES) {
    try { if (typeof purchasesOnEditDefaults_ === 'function') purchasesOnEditDefaults_(e); } catch (err) {
      logError_('purchasesOnEditDefaults_', err, { sheet: name, a1: e && e.range && e.range.getA1Notation() });
    }
    try { if (typeof purchasesOnEditSku_ === 'function') purchasesOnEditSku_(e); } catch (err) {
      logError_('purchasesOnEditSku_', err, { sheet: name, a1: e && e.range && e.range.getA1Notation() });
    }

    // Enqueue Orders sync (time-trigger processes within ~1 minute)
    try { coco_enqueueOrdersSyncFromPurchasesEdit_(e); } catch (err) {
      logError_('coco_enqueueOrdersSyncFromPurchasesEdit_', err, { sheet: name, a1: e && e.range && e.range.getA1Notation() });
    }
    return;
  }




  // QC_UAE: auto-calc Qty Missing / Qty OK + QC Result on edit (script-based; no ARRAYFORMULA).
  if (name === (APP.SHEETS.QC_UAE || 'QC_UAE')) {
    try { if (typeof qcOnEdit_ === 'function') qcOnEdit_(e); } catch (err) {
      logError_('qcOnEdit_', err, { sheet: name, a1: e && e.range && e.range.getA1Notation() });
    }
    return;
  }

  if (name === APP.SHEETS.SHIP_CN_UAE && typeof shipmentsCnUaeOnEdit_ === 'function') {
    shipmentsCnUaeOnEdit_(e);
    return;
  }

  if (name === APP.SHEETS.SHIP_UAE_EG && typeof shipmentsUaeEgOnEdit_ === 'function') {
    shipmentsUaeEgOnEdit_(e);
    return;
  }

  if (name === APP.SHEETS.SALES_EG && typeof salesEgOnEdit_ === 'function') {
    salesEgOnEdit_(e);
    return;
  }
}

function onEdit(e) {
  try {
    // If installable trigger is enabled, avoid double-run:
    // - Simple trigger runs with LIMITED authMode.
    // - Installable runs with FULL.
    if (e && e.authMode === ScriptApp.AuthMode.LIMITED && _useInstallableOnEditFlag_()) return;

    _dispatchOnEdit_(e);
  } catch (err) {
    logError_('onEdit', err, {
      sheet: e && e.range && e.range.getSheet().getName(),
      a1:    e && e.range && e.range.getA1Notation()
    });
  }
}


/** Installable onEdit handler (FULL auth) */
function coco_onEditInstallable(e) {
  try {
    _dispatchOnEdit_(e);
  } catch (err) {
    logError_('coco_onEditInstallable', err, {
      sheet: e && e.range && e.range.getSheet().getName(),
      a1:    e && e.range && e.range.getA1Notation()
    });
  }
}

/** ===================== TRIGGERS (INSTALLABLE) ===================== */

function _coco_deleteTriggersNoLock_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    const fn = t.getHandlerFunction();
    // Delete legacy + current handlers safely
    if (fn === 'onEdit' || fn === 'onOpen' || fn === 'coco_onEditInstallable' || fn === 'coco_processSyncQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
}

function coco_installTriggers() {
  return withLock_('coco_installTriggers', function () {
    try {
      _coco_deleteTriggersNoLock_();

      // Cleanup legacy/broken queue keys (older builds stored under key 'undefined')
      try {
        const dp = PropertiesService.getDocumentProperties();
        dp.deleteProperty('undefined');
      } catch (e) {}

      const ss = getSpreadsheet_();

      // Installable onEdit (FULL auth)
      ScriptApp.newTrigger('coco_onEditInstallable').forSpreadsheet(ss).onEdit().create();
      _setUseInstallableOnEditFlag_(true);

      // Time-based queue processor (debounced Orders sync)
      ScriptApp.newTrigger('coco_processSyncQueue').timeBased().everyMinutes(1).create();

      safeAlert_('✅ Triggers installed:\n- coco_onEditInstallable (onEdit)\n- coco_processSyncQueue (every 1 min)');
    } catch (err) {
      logError_('coco_installTriggers', err);
      safeAlert_('❌ Trigger install failed: ' + err.message);
      throw err;
    }
  });
}

function coco_uninstallTriggers() {
  return withLock_('coco_uninstallTriggers', function () {
    try {
      _coco_deleteTriggersNoLock_();
      try {
        const dp = PropertiesService.getDocumentProperties();
        dp.deleteProperty('undefined');
      } catch (e) {}
      _setUseInstallableOnEditFlag_(false);
      safeAlert_('✅ Installable triggers removed.');
    } catch (err) {
      logError_('coco_uninstallTriggers', err);
      throw err;
    }
  });
}

/** ===================== TESTS ===================== */

function runCoreTests() {
  try {
    const purchases = ensureSheet_(APP.SHEETS.PURCHASES);
    Logger.log('Found Purchases sheet: ' + purchases.getName());

    const errSh = ensureErrorLog_();
    Logger.log('ErrorLog sheet: ' + errSh.getName());

    const fx = getDefaultFxAedEgp_();
    Logger.log('Default FX AED→EGP: ' + fx);

    safeAlert_('✅ Core tests passed.');
  } catch (err) {
    logError_('runCoreTests', err);
    safeAlert_('❌ Core tests failed: ' + err.message);
    throw err;
  }
}

/** ===================== DEEP FREEZE (CONFIG SAFETY) ===================== */

function deepFreeze_(obj) {
  if (!obj || typeof obj !== 'object' || Object.isFrozen(obj)) return obj;
  Object.getOwnPropertyNames(obj).forEach(function (p) {
    try { deepFreeze_(obj[p]); } catch (e) {}
  });
  return Object.freeze(obj);
}

deepFreeze_(APP);


/** ============================================================
 * Formatting Helpers (used by Logistics.gs layouts)
 * ============================================================ */
function _applyNumberFormatByHeaders_(sh, headerMap, headerNames, numberFormat) {
  if (!sh || !headerMap || !headerNames || !headerNames.length) return;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const ranges = [];
  for (const h of headerNames) {
    const col1 = headerMap[h];
    if (!col1) continue;
    const col = Number(col1);
    if (!col || col < 1) continue;
    ranges.push(sh.getRange(2, col, lastRow - 1, 1).getA1Notation());
  }
  if (!ranges.length) return;
  sh.getRangeList(ranges).setNumberFormat(numberFormat);
}

function _applyDateFormatByHeaders_(sh, headerMap, headerNames) {
  _applyNumberFormatByHeaders_(sh, headerMap, headerNames, 'yyyy-mm-dd');
}
function _applyIntFormatByHeaders_(sh, headerMap, headerNames) {
  _applyNumberFormatByHeaders_(sh, headerMap, headerNames, '0');
}
function _applyDecimalFormatByHeaders_(sh, headerMap, headerNames) {
  _applyNumberFormatByHeaders_(sh, headerMap, headerNames, '0.00');
}


function coco_enqueueOrdersSync_(orderIds, opts) {
  opts = opts || {};
  return withLock_('coco_enqueueOrdersSync_', function () {
    const dp = PropertiesService.getDocumentProperties();

    if (opts.forceAll) {
      dp.setProperty(APP.INTERNAL.ORDERS_SYNC_ALL_FLAG, '1');
      dp.deleteProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY);
      return;
    }

    const ids = (orderIds || [])
      .map(function (x) { return String(x || '').trim(); })
      .filter(function (x) { return !!x; });

    if (!ids.length) return;

    let existing = [];
    try {
      const raw = dp.getProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY);
      if (raw) existing = JSON.parse(raw) || [];
    } catch (e) {
      existing = [];
    }

    const set = {};
    existing.forEach(function (x) { set[String(x || '').trim()] = true; });
    ids.forEach(function (x) { set[x] = true; });

    const merged = Object.keys(set).filter(function (x) { return !!x; });

    // Safety cap (avoid oversized properties)
    const capped = (merged.length > 5000) ? merged.slice(0, 5000) : merged;

    dp.setProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY, JSON.stringify(capped));
  });
}


function coco_enqueueOrdersSyncFromPurchasesEdit_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== APP.SHEETS.PURCHASES) return;

  const nr = e.range.getNumRows();
  const nc = e.range.getNumColumns();

  // Huge paste -> full rebuild (safer than trying to deduce IDs)
  if (nr > 300 || nc > 25) {
    coco_enqueueOrdersSync_(null, { forceAll: true });
    return;
  }

  const map = getHeaderMap_(sh, 1);
  const cOrder = map[APP.COLS.PURCHASES.ORDER_ID] || map['Order ID'];
  if (!cOrder) return;

  const startRow = Math.max(2, e.range.getRow());
  const endRow = Math.min(sh.getLastRow(), e.range.getRow() + nr - 1);
  const n = Math.max(0, endRow - startRow + 1);
  if (n <= 0) return;

  const vals = sh.getRange(startRow, cOrder, n, 1).getValues();
  const set = {};
  vals.forEach(function (r) {
    const oid = String(r[0] || '').trim();
    if (oid) set[oid] = true;
  });

  const ids = Object.keys(set);
  if (ids.length) coco_enqueueOrdersSync_(ids);
}


function coco_processSyncQueue() {
  return withLock_('coco_processSyncQueue', function () {
    const dp = PropertiesService.getDocumentProperties();

    const forceAll = String(dp.getProperty(APP.INTERNAL.ORDERS_SYNC_ALL_FLAG) || '') === '1';
    const raw = dp.getProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY);

    if (!forceAll && !raw) return;

    // Clear flags early; if we fail we'll restore.
    dp.deleteProperty(APP.INTERNAL.ORDERS_SYNC_ALL_FLAG);
    dp.deleteProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY);
    // Clear last error before run
    try { dp.deleteProperty(APP.INTERNAL.ORDERS_SYNC_LAST_ERROR); } catch (e) {}

    try {
      const markSuccess_ = () => {
        try { dp.setProperty(APP.INTERNAL.ORDERS_SYNC_LAST_RUN, new Date().toISOString()); } catch (e) {}
      };

      if (forceAll) {
        if (typeof rebuildOrdersSummary === 'function') {
          rebuildOrdersSummary();
        } else if (typeof orders_syncFromPurchasesByOrderIds_ === 'function') {
          // As a fallback, treat as full sync (caller will likely rebuild anyway)
          orders_syncFromPurchasesByOrderIds_([]);
        }
        markSuccess_();
        return;
      }

      let ids = [];
      try { ids = JSON.parse(raw) || []; } catch (e) { ids = []; }
      ids = ids.map(function (x) { return String(x || '').trim(); }).filter(Boolean);

      if (!ids.length) return;

      // If too many IDs, full rebuild is usually cheaper & safer.
      if (ids.length > 200 && typeof rebuildOrdersSummary === 'function') {
        rebuildOrdersSummary();
        markSuccess_();
        return;
      }

      if (typeof orders_syncFromPurchasesByOrderIds_ === 'function') {
        orders_syncFromPurchasesByOrderIds_(ids);
      } else if (typeof rebuildOrdersSummary === 'function') {
        rebuildOrdersSummary();
      }

      // If we got here without throwing, mark run success
      markSuccess_();
    } catch (err) {
      // Restore queue for retry
      if (forceAll) dp.setProperty(APP.INTERNAL.ORDERS_SYNC_ALL_FLAG, '1');
      if (raw) dp.setProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY, raw);
      try { dp.setProperty(APP.INTERNAL.ORDERS_SYNC_LAST_ERROR, String((err && err.stack) || err)); } catch (e) {}
      logError_('coco_processSyncQueue', err);
      throw err;
    }
  });
}


function coco_processSyncQueueNow() {
  try {
    coco_processSyncQueue();
    safeAlert_('✅ Sync queue processed.');
  } catch (e) {
    logError_('coco_processSyncQueueNow', e);
    safeAlert_('❌ Sync queue failed: ' + e.message);
    throw e;
  }
}

/**
 * Debug helper: shows Orders sync queue status + availability of sync functions.
 */
function coco_debugOrdersSyncStatus() {
  try {
    const dp = PropertiesService.getDocumentProperties();
    const rawQueue = dp.getProperty(APP.INTERNAL.ORDERS_SYNC_QUEUE_KEY) || '[]';
    const forceAll = String(dp.getProperty(APP.INTERNAL.ORDERS_SYNC_ALL_FLAG) || '') === '1';
    const lastRun  = dp.getProperty(APP.INTERNAL.ORDERS_SYNC_LAST_RUN) || '';
    const lastErr  = dp.getProperty(APP.INTERNAL.ORDERS_SYNC_LAST_ERROR) || '';

    let queueLen = -1;
    try { queueLen = (JSON.parse(rawQueue) || []).length; } catch (e) { queueLen = -1; }

    const hasIncremental = (typeof orders_syncFromPurchasesByOrderIds_ === 'function');
    const hasRebuild     = (typeof rebuildOrdersSummary === 'function');

    const msg = [
      'Orders Sync Status',
      '------------------',
      'forceAll: ' + forceAll,
      'queueLen: ' + queueLen,
      'queueRaw: ' + (rawQueue ? rawQueue.slice(0, 200) : '(empty)') + (rawQueue && rawQueue.length > 200 ? '...' : ''),
      '',
      'hasIncremental: ' + hasIncremental,
      'hasRebuild: ' + hasRebuild,
      '',
      'lastRun: ' + (lastRun || '(none)'),
      'lastError: ' + (lastErr ? lastErr.slice(0, 400) : '(none)')
    ].join('\n');

    safeAlert_(msg);
  } catch (e) {
    logError_('coco_debugOrdersSyncStatus', e);
    safeAlert_('Debug failed: ' + (e && e.message ? e.message : e));
  }
}

