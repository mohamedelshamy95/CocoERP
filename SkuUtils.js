/** =========================================================
 * SkuUtils.gs – SKU policy + generation + catalog lookup + Purchases hooks
 * CocoERP v2.2+
 *
 * Depends on AppCore (kernel):
 *  - APP, getSheet_, ensureSheet_, ensureSheetSchema_, normalizeHeaders_,
 *    getHeaderMap_, logError_, ensureErrorLog_, withLock_ (optional)
 * ========================================================= */

/** ---------- Headers (canonical from AppCore, with safe fallbacks) ---------- */
function sku_headers_() {
  const P = (APP && APP.COLS && APP.COLS.PURCHASES) ? APP.COLS.PURCHASES : {};
  const C = (APP && APP.COLS && APP.COLS.CATALOG_EG) ? APP.COLS.CATALOG_EG : {};

  return {
    // Purchases
    P_ORDER: P.ORDER_ID     || 'Order ID',
    P_SKU:   P.SKU          || 'SKU',
    P_PROD:  P.PRODUCT_NAME || 'Product Name',
    P_VAR:   P.VARIANT      || 'Variant / Color',

    // Catalog
    C_SKU:   C.SKU          || 'SKU',
    C_PROD:  C.PRODUCT_NAME || 'Product Name',
    C_VAR:   C.VARIANT      || 'Variant / Color'
  };
}

/** ---------- Text normalization helpers ---------- */
function sku_normalizeDigits_(s) {
  const str = String(s || '');
  const map = {
    '٠':'0','١':'1','٢':'2','٣':'3','٤':'4','٥':'5','٦':'6','٧':'7','٨':'8','٩':'9',
    '۰':'0','۱':'1','۲':'2','۳':'3','۴':'4','۵':'5','۶':'6','۷':'7','۸':'8','۹':'9'
  };
  return str.replace(/[٠-٩۰-۹]/g, function (d) { return map[d] || d; });
}

/**
 * Clean text for SKU tokenization:
 * - Uppercase
 * - Convert Arabic/Persian digits
 * - Keep A-Z 0-9
 * - Normalize separators/spaces
 */
function sku_cleanText_(s) {
  return sku_normalizeDigits_(s)
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Normalize SKU string itself (if user typed it):
 * - Uppercase
 * - Keep A-Z 0-9 and '-'
 * - Convert spaces/underscores to '-'
 * - Collapse multiple '-'
 */
function sku_normalizeSku_(sku) {
  let s = sku_normalizeDigits_(sku).toUpperCase();
  s = s.replace(/[_\s]+/g, '-');
  s = s.replace(/[^A-Z0-9-]+/g, '');
  s = s.replace(/-+/g, '-');
  s = s.replace(/^-+|-+$/g, '');
  return s;
}

function sku_tokenize_(s) {
  const t = sku_cleanText_(s);
  if (!t) return [];
  const parts = t.split(' ').filter(Boolean);

  // de-dup tokens while preserving order
  const seen = {};
  const out = [];
  for (let i = 0; i < parts.length; i++) {
    const p = parts[i];
    if (!seen[p]) {
      seen[p] = true;
      out.push(p);
    }
  }
  return out;
}

/** ---------- SKU generation (stable heuristic) ---------- */
function sku_generateFromText_(productName, variant) {
  const nameTokens = sku_tokenize_(productName);
  const varTokens  = sku_tokenize_(variant);

  if (!nameTokens.length && !varTokens.length) return '';

  // brand: first token from productName, else from variant
  const brand = (nameTokens[0] || varTokens[0] || '').trim();

  // model: first token containing digit from (nameTokens then varTokens)
  let model = '';
  const all = nameTokens.concat(varTokens);
  for (let i = 0; i < all.length; i++) {
    if (/\d/.test(all[i])) { model = all[i]; break; }
  }

  // color: prefer last token from variant that is not brand/model
  let color = '';
  for (let i = varTokens.length - 1; i >= 0; i--) {
    const tok = varTokens[i];
    if (!tok) continue;
    if (tok === brand) continue;
    if (model && tok === model) continue;
    color = tok;
    break;
  }

  const parts = [];
  if (brand) parts.push(brand);
  if (model && model !== brand) parts.push(model);
  if (color && color !== brand && color !== model) parts.push(color);

  const rawSku = parts.join('-');
  return sku_normalizeSku_(rawSku);
}

/** ---------- Fingerprint for Catalog lookup ---------- */
function sku_fingerprint_(productName, variant) {
  const p = sku_cleanText_(productName);
  const v = sku_cleanText_(variant);
  return (p + '|' + v).trim();
}

/** ---------- Catalog index (fingerprint -> SKU) with cache ---------- */
function sku_clearCatalogIndexCache_() {
  try {
    const cache = CacheService.getDocumentCache();
    cache.remove('CocoERP_CatalogSkuIndex_v1');
  } catch (e) {}
}

/**
 * Build or load Catalog index:
 * - Key: fingerprint(product, variant)
 * - Value: normalized SKU
 */
function sku_getCatalogIndex_(forceRebuild) {
  const cacheKey = 'CocoERP_CatalogSkuIndex_v1';
  const cache = CacheService.getDocumentCache();

  if (!forceRebuild) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (e) {}
    }
  }

  const H = sku_headers_();

  // Ensure Catalog exists + schema (non-destructive)
  const sh = ensureSheet_(APP.SHEETS.CATALOG_EG);
  try { normalizeHeaders_(sh, 1); } catch (e) {}
  try { ensureSheetSchema_(APP.SHEETS.CATALOG_EG, Object.values(APP.COLS.CATALOG_EG), { addMissing: true, headerRow: 1 }); } catch (e) {}

  const map = getHeaderMap_(sh, 1);
  const cSku  = map[H.C_SKU];
  const cProd = map[H.C_PROD];
  const cVar  = map[H.C_VAR];

  const idx = {};
  const lastRow = sh.getLastRow();
  if (lastRow >= 2 && cSku && cProd && cVar) {
    const lastCol = sh.getLastColumn();
    const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const skuRaw = row[cSku - 1];
      if (!skuRaw) continue;

      const sku = sku_normalizeSku_(skuRaw);
      if (!sku) continue;

      const fp = sku_fingerprint_(row[cProd - 1], row[cVar - 1]);
      if (!fp) continue;

      // do not overwrite existing mapping (first wins)
      if (!idx[fp]) idx[fp] = sku;
    }
  }

  cache.put(cacheKey, JSON.stringify(idx), 300); // 5 minutes
  return idx;
}

function sku_lookupFromCatalog_(productName, variant) {
  try {
    const fp = sku_fingerprint_(productName, variant);
    if (!fp) return '';
    const idx = sku_getCatalogIndex_(false);
    return idx[fp] || '';
  } catch (e) {
    return '';
  }
}

/** ---------- Optional: Register generated SKU as Draft in Catalog ---------- */
function sku_registerDraftToCatalog_(sku, productName, variant, source) {
  const normalized = sku_normalizeSku_(sku);
  if (!normalized) return;

  const H = sku_headers_();
  const sh = ensureSheet_(APP.SHEETS.CATALOG_EG);

  try { normalizeHeaders_(sh, 1); } catch (e) {}
  try { ensureSheetSchema_(APP.SHEETS.CATALOG_EG, Object.values(APP.COLS.CATALOG_EG), { addMissing: true, headerRow: 1 }); } catch (e) {}

  const map = getHeaderMap_(sh, 1);
  const cSku  = map[H.C_SKU];
  const cProd = map[H.C_PROD];
  const cVar  = map[H.C_VAR];

  if (!cSku || !cProd || !cVar) return;

  // Check existence by SKU (simple scan on column; OK for small/medium catalogs)
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    const skus = sh.getRange(2, cSku, lastRow - 1, 1).getValues();
    for (let i = 0; i < skus.length; i++) {
      if (sku_normalizeSku_(skus[i][0]) === normalized) return; // already exists
    }
  }

  // Append draft row (keep other columns blank)
  const row = [];
  const lastCol = sh.getLastColumn();
  for (let i = 0; i < lastCol; i++) row.push('');

  row[cSku - 1]  = normalized;
  row[cProd - 1] = productName || '';
  row[cVar - 1]  = variant || '';

  // Optional: if columns exist in your Catalog schema, fill them when present
  if (map['Status']) row[map['Status'] - 1] = 'Draft';
  if (map['Notes'])  row[map['Notes'] - 1]  = (source ? ('Auto-registered from ' + source) : 'Auto-registered');

  sh.appendRow(row);

  // Refresh cache
  sku_clearCatalogIndexCache_();
}

/** ---------- Purchases: backfill SKU (safe, non-destructive) ---------- */
/**
 * Backfill SKU in Purchases:
 * - Does NOT overwrite existing SKU
 * - Default: only fill rows where Order ID exists
 *
 * opts:
 *  - onlyIfOrderId: boolean (default true)
 *  - registerDraftToCatalog: boolean (default false)
 */
function sku_backfillPurchasesSku_(opts) {
  const options = opts || {};
  const onlyIfOrderId = (options.onlyIfOrderId !== false);
  const registerDraft = !!options.registerDraftToCatalog;

  const H = sku_headers_();
  const sh = getSheet_(APP.SHEETS.PURCHASES);
  const map = getHeaderMap_(sh, 1);

  const cOrder = map[H.P_ORDER];
  const cSku   = map[H.P_SKU]   || map['SKU'];
  const cProd  = map[H.P_PROD];
  const cVar   = map[H.P_VAR];

  if (!cSku || !cProd || !cVar) {
    throw new Error('Purchases missing required columns for SKU backfill (SKU / Product Name / Variant / Color).');
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { changed: 0, scanned: 0 };

  const lastCol = sh.getLastColumn();
  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // build catalog index once
  const catalogIdx = sku_getCatalogIndex_(false);

  let changed = 0;
  const outSku = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    const orderId = cOrder ? String(row[cOrder - 1] || '').trim() : '';
    const prod    = row[cProd - 1];
    const vari    = row[cVar  - 1];

    let sku = row[cSku - 1];

    // Keep existing SKU but normalize it
    if (sku && String(sku).trim() !== '') {
      const normalized = sku_normalizeSku_(sku);
      outSku.push([normalized]);
      if (normalized !== String(sku)) changed++;
      continue;
    }

    // If restriction enabled: require Order ID to fill
    if (onlyIfOrderId && !orderId) {
      outSku.push(['']);
      continue;
    }

    // Need at least some product context
    const fp = sku_fingerprint_(prod, vari);
    if (!fp) {
      outSku.push(['']);
      continue;
    }

    // 1) Lookup from catalog
    let generated = catalogIdx[fp] || '';

    // 2) Fallback generate
    if (!generated) generated = sku_generateFromText_(prod, vari);

    outSku.push([generated]);

    if (generated) {
      changed++;
      if (registerDraft) {
        try { sku_registerDraftToCatalog_(generated, prod, vari, 'Purchases'); } catch (e) {}
      }
    }
  }

  if (changed > 0) {
    sh.getRange(2, cSku, outSku.length, 1).setValues(outSku);
  }

  return { changed: changed, scanned: outSku.length };
}

/** Public menu-friendly wrapper (kept stable name) */
function sku_backfillPurchasesSku() {
  try {
    ensureErrorLog_();
    const res = sku_backfillPurchasesSku_({ onlyIfOrderId: true, registerDraftToCatalog: false });
    SpreadsheetApp.getUi().alert('✅ SKU Backfill done. Changed: ' + res.changed + ' / Scanned: ' + res.scanned);
  } catch (e) {
    logError_('sku_backfillPurchasesSku', e);
    SpreadsheetApp.getUi().alert('❌ SKU Backfill failed: ' + e.message);
    throw e;
  }
}

/** Backward-compatible alias (your old function name) */
function backfillPurchasesSku() {
  return sku_backfillPurchasesSku();
}

/** ---------- Purchases onEdit hook (auto fill/normalize per-row) ---------- */
function purchasesOnEditSku_(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    if (sh.getName() !== APP.SHEETS.PURCHASES) return;

    const headerRow = 1;
    const map = getHeaderMap_(sh, headerRow);

    const H = sku_headers_();
    const cOrder = map[H.P_ORDER];
    const cSku   = map[H.P_SKU]   || map['SKU'];
    const cProd  = map[H.P_PROD];
    const cVar   = map[H.P_VAR];

    if (!cSku || !cProd || !cVar) return;

    const row0 = e.range.getRow();
    if (row0 < 2) return;

    const nr = e.range.getNumRows();
    const nc = e.range.getNumColumns();

    // Hard guard against very large pastes (use menu backfill instead)
    if (nr > 120 || nc > 25) return;

    // If user typed SKU in a single cell -> normalize only
    if (nr === 1 && nc === 1 && e.range.getColumn() === cSku) {
      const raw = e.range.getValue();
      if (raw && String(raw).trim() !== '') {
        const normalized = sku_normalizeSku_(raw);
        if (normalized !== String(raw)) e.range.setValue(normalized);
      }
      return;
    }

    const startRow = row0;
    const endRow   = Math.min(sh.getLastRow(), row0 + nr - 1);
    const n = endRow - startRow + 1;
    if (n <= 0) return;

    // Read minimal row values in batch
    const skuVals  = sh.getRange(startRow, cSku,  n, 1).getValues();
    const prodVals = sh.getRange(startRow, cProd, n, 1).getValues();
    const varVals  = sh.getRange(startRow, cVar,  n, 1).getValues();
    const ordVals  = cOrder ? sh.getRange(startRow, cOrder, n, 1).getValues() : null;

    const catalogIdx = sku_getCatalogIndex_(false) || {};

    let changed = 0;

    for (let i = 0; i < n; i++) {
      const orderId = ordVals ? String(ordVals[i][0] || '').trim() : '';
      const prod    = prodVals[i][0];
      const vari    = varVals[i][0];

      const skuRaw = skuVals[i][0];

      // Keep existing SKU but normalize it
      if (skuRaw && String(skuRaw).trim() !== '') {
        const normalized = sku_normalizeSku_(skuRaw);
        if (normalized !== String(skuRaw)) {
          skuVals[i][0] = normalized;
          changed++;
        }
        continue;
      }

      // Need product context (and optionally Order ID)
      const fp = sku_fingerprint_(prod, vari);
      if (!fp) continue;

      // catalog first (fingerprint match), then generate
      let skuOut = catalogIdx[fp] || '';
      if (!skuOut) skuOut = sku_generateFromText_(prod, vari);

      if (skuOut) {
        skuVals[i][0] = skuOut;
        changed++;
      }
    }

    if (changed > 0) {
      sh.getRange(startRow, cSku, n, 1).setValues(skuVals);
    }
  } catch (err) {
    logError_('purchasesOnEditSku_', err, {
      sheet: e && e.range && e.range.getSheet && e.range.getSheet().getName(),
      a1: e && e.range && e.range.getA1Notation && e.range.getA1Notation()
    });
  }
}

/** ---------- Quick tests ---------- */
function testSkuUtils_() {
  const a = sku_generateFromText_('Monster MQT65 Monster MQT65', 'MQT65 Beige');
  Logger.log(a); // MONSTER-MQT65-BEIGE

  const b = sku_generateFromText_('', 'MQT65 White');
  Logger.log(b); // MQT65-WHITE

  const c = sku_normalizeSku_(' monster  mqt65  beige ');
  Logger.log(c); // MONSTER-MQT65-BEIGE
}

function sku_rebuildCatalogIndex_() {
  sku_clearCatalogIndexCache_();
  sku_getCatalogIndex_(true);
  SpreadsheetApp.getUi().alert('✅ Catalog SKU index rebuilt.');
}
