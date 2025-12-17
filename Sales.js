/** ============================================================
 * Sales.gs – Sales_EG sheet + sync to Inventory
 * CocoERP v2.1.2
 *
 * Depends on:
 *  - APP (AppCore3.gs)
 *  - getSheet_, getHeaderMap_, logError_
 *  - logInventoryTxn_, rebuildInventoryEGFromLedger (InventoryCore3)
 * ============================================================ */

/**
 * Sales_EG headers – 1 row per order line.
 * لازم تكون متوافقة مع APP.COLS.SALES_EG
 */
const SALES_EG_HEADERS = [
  APP.COLS.SALES_EG.ORDER_ID,         // 'Order ID'
  APP.COLS.SALES_EG.ORDER_DATE,       // 'Order Date'
  APP.COLS.SALES_EG.PLATFORM,         // 'Platform'
  APP.COLS.SALES_EG.CUSTOMER_NAME,    // 'Customer Name'
  APP.COLS.SALES_EG.PHONE,            // 'Phone'
  APP.COLS.SALES_EG.CITY,             // 'City'
  APP.COLS.SALES_EG.ADDRESS,          // 'Address'
  APP.COLS.SALES_EG.SKU,              // 'SKU'
  APP.COLS.SALES_EG.PRODUCT_NAME,     // 'Product Name'
  APP.COLS.SALES_EG.VARIANT,          // 'Variant / Color'
  APP.COLS.SALES_EG.WAREHOUSE,        // 'Warehouse (EG)'
  APP.COLS.SALES_EG.QTY,              // 'Qty'
  APP.COLS.SALES_EG.UNIT_PRICE,       // 'Unit Price (EGP)'
  APP.COLS.SALES_EG.TOTAL_PRICE,      // 'Total Price (EGP)'
  APP.COLS.SALES_EG.DISCOUNT,         // 'Discount (EGP)'
  APP.COLS.SALES_EG.NET_REVENUE,      // 'Net Revenue (EGP)'
  APP.COLS.SALES_EG.SHIPPING_FEE,     // 'Shipping Fee (EGP)'
  APP.COLS.SALES_EG.PAYMENT_METHOD,   // 'Payment Method'
  APP.COLS.SALES_EG.ORDER_STATUS,     // 'Order Status'
  APP.COLS.SALES_EG.DELIVERED_DATE,   // 'Delivered Date'
  APP.COLS.SALES_EG.SOURCE,           // 'Source'
  APP.COLS.SALES_EG.COURIER,          // 'Courier'
  APP.COLS.SALES_EG.AWB,              // 'AWB'
  APP.COLS.SALES_EG.NOTES             // 'Notes'
];

/** =============================================================
 * Layout – تجهيز شيت Sales_EG (إنشاء + الهيدرز + فورمات)
 * ============================================================= */

/**
 * ينشئ / يجهز شيت Sales_EG بالهيدرز + فورمات بسيطة للهيدر.
 */
function setupSalesLayout() {
  try {
    const sh =
      (typeof getOrCreateSheet_ === 'function')
        ? getOrCreateSheet_(APP.SHEETS.SALES_EG)
        : getSheet_(APP.SHEETS.SALES_EG);

    _setupSheetWithHeaders_(sh, SALES_EG_HEADERS);

    const map = getHeaderMap_(sh);

    // أرقام: Qty
    _applyIntFormat_(sh, [map[APP.COLS.SALES_EG.QTY]]);

    // أرقام عشرية: الأسعار / الإيراد / الشحن / الخصم
    _applyDecimalFormat_(sh, [
      map[APP.COLS.SALES_EG.UNIT_PRICE],
      map[APP.COLS.SALES_EG.TOTAL_PRICE],
      map[APP.COLS.SALES_EG.DISCOUNT],
      map[APP.COLS.SALES_EG.NET_REVENUE],
      map[APP.COLS.SALES_EG.SHIPPING_FEE]
    ]);

    // تواريخ: Order Date + Delivered Date
    _applyDateFormat_(sh, [
      map[APP.COLS.SALES_EG.ORDER_DATE],
      map[APP.COLS.SALES_EG.DELIVERED_DATE]
    ]);

    // تنسيق للهيدر
    const headerRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#e2f3ff');
    headerRange.setHorizontalAlignment('center');

    sh.setFrozenRows(1);

  } catch (e) {
    logError_('setupSalesLayout', e);
    throw e;
  }
}

/** =============================================================
 * Sync Sales_EG → Inventory_Transactions + Inventory_EG
 * ============================================================= */

/**
 * المنطق:
 * 1) نقرأ صفوف Sales_EG الـ Delivered فقط (ومعاها Delivered Date).
 * 2) نجمع الكمية لكل مفتاح: (OrderID + SKU + Warehouse).
 * 3) نقرأ من Inventory_Transactions كل الحركات اللي SourceType = 'SALE_EG'
 *    ونجمع الكمية OUT لنفس المفتاح.
 * 4) لو SalesQty > LedgerQty → نعمل حركة جديدة بالـ delta بس.
 * 5) بعد ما نخلص → نعيد بناء Inventory_EG من الـ Ledger.
 */
function syncSalesFromOrdersSheet() {
  try {
    const salesSh   = getSheet_(APP.SHEETS.SALES_EG);
    const ledgerSh  = getSheet_(APP.SHEETS.INVENTORY_TXNS);
    const invEgSh   = getSheet_(APP.SHEETS.INVENTORY_EG);
    const catalogSh = getSheet_(APP.SHEETS.CATALOG_EG);

    const salesMap  = getHeaderMap_(salesSh);
    const lastRow   = salesSh.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('Sales_EG: لا توجد صفوف بيانات لمزامنتها.');
      return;
    }

    const data = salesSh
      .getRange(2, 1, lastRow - 1, salesSh.getLastColumn())
      .getValues();

    const idxOrderId   = salesMap[APP.COLS.SALES_EG.ORDER_ID]       - 1;
    const idxSku       = salesMap[APP.COLS.SALES_EG.SKU]            - 1;
    const idxWh        = salesMap[APP.COLS.SALES_EG.WAREHOUSE]      - 1;
    const idxQty       = salesMap[APP.COLS.SALES_EG.QTY]            - 1;
    const idxStatus    = salesMap[APP.COLS.SALES_EG.ORDER_STATUS]   - 1;
    const idxDelivDate = salesMap[APP.COLS.SALES_EG.DELIVERED_DATE] - 1;
    const idxUnitPrice = salesMap[APP.COLS.SALES_EG.UNIT_PRICE]     - 1;
    const idxProdName  = salesMap[APP.COLS.SALES_EG.PRODUCT_NAME]   - 1;
    const idxVariant   = salesMap[APP.COLS.SALES_EG.VARIANT]        - 1;
    // Delivered detector: prefer AppCore helper when available
    const isDelivered_ = (typeof isDeliveredStatus_ === 'function')
      ? isDeliveredStatus_
      : function (st) {
          if (!st) return false;
          const s = String(st).trim().toLowerCase();
          return (s === 'delivered' || s.includes('deliv') || s.includes('تسليم') || s.includes('توصيل'));
        };
   // Strong guards: required columns (fail fast with clear error)
     assertRequiredColumns_(salesSh, [
       APP.COLS.SALES_EG.ORDER_ID,
       APP.COLS.SALES_EG.SKU,
       APP.COLS.SALES_EG.WAREHOUSE,
       APP.COLS.SALES_EG.QTY,
       APP.COLS.SALES_EG.ORDER_STATUS,
       APP.COLS.SALES_EG.DELIVERED_DATE,
       APP.COLS.SALES_EG.UNIT_PRICE,
       APP.COLS.SALES_EG.PRODUCT_NAME,
       APP.COLS.SALES_EG.VARIANT
     ]); 

     assertRequiredColumns_(ledgerSh, [
       APP.COLS.INV_TXNS.SOURCE_TYPE,
       APP.COLS.INV_TXNS.SOURCE_ID,
       APP.COLS.INV_TXNS.SKU,
       APP.COLS.INV_TXNS.WAREHOUSE,
       APP.COLS.INV_TXNS.QTY_OUT
     ]);

    // --------- 1) نجمع مبيعات Sales_EG ----------
    /**
     * key: OrderID||SKU||Warehouse
     * value: { orderId, sku, warehouse, qty, unitPrice, productName, variant, lastDeliveredDate }
     */
    const salesAgg = {};

    data.forEach(function (row) {
      const orderId      = row[idxOrderId];
      const skuRaw       = row[idxSku];
      const warehouseRaw = row[idxWh];
      const qty          = Number(row[idxQty] || 0);
      const status       = row[idxStatus];
      const delDate      = row[idxDelivDate];

      const sku = (skuRaw || '').toString().trim();
      let wh    = (warehouseRaw || '').toString().trim();

      if (!orderId || !sku || !qty) return;
      if (!wh) wh = (APP && APP.WAREHOUSES && APP.WAREHOUSES.EG_CAI) ? APP.WAREHOUSES.EG_CAI : 'EG-CAI';
      if (typeof normalizeWarehouseCode_ === 'function') wh = normalizeWarehouseCode_(wh);
      if (!String(wh).toUpperCase().startsWith('EG-')) return;
      if (!isDelivered_(status)) return;
      if (!delDate) return; // لازم يبقى في Delivered Date

      const key = orderId + '||' + sku + '||' + wh;

      if (!salesAgg[key]) {
        salesAgg[key] = {
          orderId: orderId,
          sku: sku,
          warehouse: wh,
          qty: 0,
          unitPrice: Number(row[idxUnitPrice] || 0),
          productName: row[idxProdName] || '',
          variant: row[idxVariant] || '',
          lastDeliveredDate: delDate
        };
      }

      const rec = salesAgg[key];
      rec.qty += qty;

      // ناخد آخر تاريخ تسليم
      if (delDate && (!rec.lastDeliveredDate || delDate > rec.lastDeliveredDate)) {
        rec.lastDeliveredDate = delDate;
      }
    });

    // لو مفيش ولا صف Delivered
    if (!Object.keys(salesAgg).length) {
      SpreadsheetApp.getUi().alert('Sales_EG: لا توجد صفوف Delivered لمزامنتها.');
      return;
    }

    // --------- 2) نقرأ Ledger الحالى لحركات SALE_EG ----------
    const ledgerMap = getHeaderMap_(ledgerSh);
    const lastLedgerRow = ledgerSh.getLastRow();

    /**
     * ledgerAgg: key → alreadyOutQty
     */
    const ledgerAgg = {};

    if (lastLedgerRow >= 2) {
      const ledData = ledgerSh
        .getRange(2, 1, lastLedgerRow - 1, ledgerSh.getLastColumn())
        .getValues();

      const idxSrcType = ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] - 1;
      const idxSrcId   = ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID]   - 1;
      const idxSkuL    = ledgerMap[APP.COLS.INV_TXNS.SKU]         - 1;
      const idxWhL     = ledgerMap[APP.COLS.INV_TXNS.WAREHOUSE]   - 1;
      const idxQtyOut  = ledgerMap[APP.COLS.INV_TXNS.QTY_OUT]     - 1;

      ledData.forEach(function (row) {
        const srcType = (row[idxSrcType] || '').toString().trim();
        if (srcType !== 'SALE_EG') return;

        const orderId = row[idxSrcId];
        const sku     = (row[idxSkuL] || '').toString().trim();
        const wh      = (row[idxWhL]  || '').toString().trim();
        const qtyOut  = Number(row[idxQtyOut] || 0);

        if (!orderId || !sku || !wh || !qtyOut) return;

        const key = orderId + '||' + sku + '||' + wh;
        ledgerAgg[key] = (ledgerAgg[key] || 0) + qtyOut;
      });
    }

    // --------- 3) Cost map من Inventory_EG + fallback Catalog ----------
    const costBySkuWh = {};
    const invEgMap    = getHeaderMap_(invEgSh);
    const invEgLast   = invEgSh.getLastRow();

    if (invEgLast >= 2) {
      const invData = invEgSh
        .getRange(2, 1, invEgLast - 1, invEgSh.getLastColumn())
        .getValues();

      const idxSkuInv  = invEgMap['SKU']             - 1;
      const idxWhInv   = invEgMap['Warehouse (EG)']  - 1;
      const idxAvgCost = invEgMap['Avg Cost (EGP)']  - 1;

      invData.forEach(function (row) {
        const sku = (row[idxSkuInv] || '').toString().trim();
        const wh  = (row[idxWhInv]  || '').toString().trim();
        if (!sku || !wh) return;
        const key  = sku + '||' + wh;
        const cost = Number(row[idxAvgCost] || 0);
        if (cost) costBySkuWh[key] = cost;
      });
    }

    // كتالوج: فى حالة عدم توفر Cost من Inventory_EG
    const catalogMap = getHeaderMap_(catalogSh);
    const catLast    = catalogSh.getLastRow();
    let catData      = [];
    let idxSkuCat, idxProdCat, idxVarCat, idxCostCat;

    if (catLast >= 2) {
      catData   = catalogSh.getRange(2, 1, catLast - 1, catalogSh.getLastColumn()).getValues();
      idxSkuCat = catalogMap['SKU']                - 1;
      idxProdCat= catalogMap['Product Name']       - 1;
      idxVarCat = catalogMap['Variant / Color']    - 1;
      idxCostCat= catalogMap['Default Cost (EGP)'] - 1;
    }

    function normSku_(s) {
  return String(s || '').trim().toUpperCase().replace(/\s+/g, '');
}

const catBySku = {};
if (catData.length && idxSkuCat >= 0) {
  for (var i = 0; i < catData.length; i++) {
    const k = normSku_(catData[i][idxSkuCat]);
    if (!k) continue;
    if (!catBySku[k]) {
      catBySku[k] = {
        product: catData[i][idxProdCat] || '',
        variant: catData[i][idxVarCat]  || '',
        cost: Number(catData[i][idxCostCat] || 0)
      };
    }
  }
}

function lookupCatalog_(sku) {
  const k = normSku_(sku);
  return k ? (catBySku[k] || null) : null;
}

    // --------- 4) حساب الـ delta وتسجيل الحركات ----------
    let newTxns = 0;
    let skipped = 0;
    const txns = [];

    Object.keys(salesAgg).forEach(function (key) {
      const rec       = salesAgg[key];
      const already   = ledgerAgg[key] || 0;
      const deltaQty  = rec.qty - already;

      if (deltaQty <= 0) {
        skipped++;
        return;
      }

      const costKey = rec.sku + '||' + rec.warehouse;
      let unitCost  = costBySkuWh[costKey] || 0;
      const catInfo = lookupCatalog_(rec.sku);

      if (!unitCost && catInfo && catInfo.cost) {
        unitCost = catInfo.cost;
      }

      const prodName = rec.productName || (catInfo && catInfo.product) || rec.sku;
      const variant  = rec.variant     || (catInfo && catInfo.variant) || '';

      txns.push({
        type: 'OUT',
        sourceType: 'SALE_EG',
        sourceId: String(rec.orderId),
        batchCode: '',
        sku: rec.sku,
        productName: prodName,
        variant: variant,
        warehouse: rec.warehouse,
        qty: deltaQty,
        unitCostEgp: unitCost,
        currency: 'EGP',
        unitPriceOrig: rec.unitPrice || '',
        txnDate: rec.lastDeliveredDate || new Date(),
        notes: 'SALE_EG (delta=' + deltaQty + ')'
      });

      newTxns++;
    });

    // --------- 5) Write new txns + rebuild Snapshot مصر ----------
    if (txns.length) {
      if (typeof logInventoryTxnBatch_ === 'function') {
        logInventoryTxnBatch_(txns);
      } else {
        // Fallback (slower)
        txns.forEach(function (t) { logInventoryTxn_(t); });
      }
    }

    if (txns.length && typeof rebuildInventoryEGFromLedger === 'function') {
      rebuildInventoryEGFromLedger();
    }
SpreadsheetApp.getUi().alert(
      'Sales_EG sync done.\n\n' +
      'New txns: ' + newTxns + '\n' +
      'Skipped (not delivered / already synced / no delta / invalid): ' + skipped
    );

  } catch (e) {
    logError_('syncSalesFromOrdersSheet', e);
    throw e;
  }
}

/** =============================================================
 * Sales_EG – onEdit helpers (auto-fill + recalcs)
 * ============================================================= */

/**
 * Trigger داخلي يستدعيه onEdit العام في AppCore3:
 *  - يكمل بيانات الصف لما SKU يتكتب (Product / Variant / Default Price).
 *  - يعيد حساب الأسعار + Net Revenue.
 *  - يحدد Delivered Date تلقائياً عند تغيير الحالة إلى Delivered.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function salesEgOnEdit_(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== APP.SHEETS.SALES_EG) return;

    const map = getHeaderMap_(sheet);

    const rowStart = e.range.getRow();
    const rowEnd   = rowStart + e.range.getNumRows() - 1;
    const colStart = e.range.getColumn();
    const colEnd   = colStart + e.range.getNumColumns() - 1;

    const skuCol       = map[APP.COLS.SALES_EG.SKU];
    const qtyCol       = map[APP.COLS.SALES_EG.QTY];
    const unitPriceCol = map[APP.COLS.SALES_EG.UNIT_PRICE];
    const discountCol  = map[APP.COLS.SALES_EG.DISCOUNT];
    const shipFeeCol   = map[APP.COLS.SALES_EG.SHIPPING_FEE];
    const statusCol    = map[APP.COLS.SALES_EG.ORDER_STATUS];
    const deliveredCol = map[APP.COLS.SALES_EG.DELIVERED_DATE];

    const touchesCol_ = function (c) {
      return c && c >= colStart && c <= colEnd;
    };

    for (let row = rowStart; row <= rowEnd; row++) {
      if (row === 1) continue; // header

      // 1) SKU backfill
      if (touchesCol_(skuCol)) {
        salesEgBackfillFromCatalog_(sheet, map, row);
        recalcSalesRowAmounts_(sheet, map, row);
        continue;
      }

      // 2) Recalc when numeric fields change
      if (
        touchesCol_(qtyCol) ||
        touchesCol_(unitPriceCol) ||
        touchesCol_(discountCol) ||
        touchesCol_(shipFeeCol)
      ) {
        recalcSalesRowAmounts_(sheet, map, row);
      }

      // 3) Delivered date
      if (touchesCol_(statusCol) && statusCol && deliveredCol) {
        const statusVal = String(sheet.getRange(row, statusCol).getValue() || '')
          .trim()
          .toLowerCase();

        const deliveredCell = sheet.getRange(row, deliveredCol);

        const deliveredNow = (typeof isDeliveredStatus_ === 'function') ? isDeliveredStatus_(statusVal) : (statusVal === 'delivered' || statusVal === 'تم التسليم' || statusVal === 'تم التوصيل' || statusVal === 'تم التسليم للعميل');
        if (deliveredNow) {
          if (!deliveredCell.getValue()) deliveredCell.setValue(new Date());
        }
      }
    }
  } catch (err) {
    logError_('salesEgOnEdit_', err, {
      sheet: e && e.range && e.range.getSheet().getName(),
      a1: e && e.range && e.range.getA1Notation()
    });
  }
}

/**
 * يملأ بيانات صف واحد في Sales_EG من الكتالوج بناءً على الـ SKU.
 */
function salesEgBackfillFromCatalog_(sheet, map, row) {
  const skuCol       = map[APP.COLS.SALES_EG.SKU];
  const prodCol      = map[APP.COLS.SALES_EG.PRODUCT_NAME];
  const variantCol   = map[APP.COLS.SALES_EG.VARIANT];
  const unitPriceCol = map[APP.COLS.SALES_EG.UNIT_PRICE];
  const whCol        = map[APP.COLS.SALES_EG.WAREHOUSE];

  if (!skuCol) return;

  const sku = String(sheet.getRange(row, skuCol).getValue() || '').trim();
  if (!sku) return;

  const cat = catalogLookupBySku_(sku);
  if (cat) {
    if (prodCol)    sheet.getRange(row, prodCol).setValue(cat.productName);
    if (variantCol) sheet.getRange(row, variantCol).setValue(cat.variant);
    if (unitPriceCol && cat.defaultPriceEgp != null) {
      const cell = sheet.getRange(row, unitPriceCol);
      if (!cell.getValue()) cell.setValue(cat.defaultPriceEgp);
    }
  }

  // Default warehouse لو فاضي
  if (whCol) {
    const whVal = String(sheet.getRange(row, whCol).getValue() || '').trim();
    if (!whVal) {
      sheet.getRange(row, whCol).setValue('EG-CAI');
    }
  }
}

/**
 * قراءة صف الكتالوج حسب الـ SKU (بشكل مرن شوية).
 * - بيقارن الـ SKU بعد ما يشيل المسافات ويحوله UPPERCASE
 *
 * @param {string} sku
 * @return {{productName:string, variant:string, defaultPriceEgp:number, brand:string}|null}
 */
/**
 * Cached catalog lookup by SKU (fast).
 * - Builds a normalized SKU index in DocumentCache for 5 minutes.
 *
 * @param {string} sku
 * @return {{productName:string, variant:string, defaultPriceEgp:number, brand:string, defaultCostEgp:number}|null}
 */
function catalogLookupBySku_(sku) {
  if (!sku) return null;

  const normalizeSku = function (s) {
    return String(s || '')
      .trim()
      .toUpperCase()
      .replace(/\s+/g, '')
      .replace(/[_]+/g, '');
  };

  const target = normalizeSku(sku);
  if (!target) return null;

  const cacheKey = 'CocoERP_CatalogSkuIndexBySku_v1';
  const cache = CacheService.getDocumentCache();

  let idx = null;
  const cached = cache.get(cacheKey);
  if (cached) {
    try { idx = JSON.parse(cached); } catch (e) {}
  }

  if (!idx) {
    const catSh  = getSheet_(APP.SHEETS.CATALOG_EG);
    try { normalizeHeaders_(catSh, 1); } catch (e) {}
    const catMap = getHeaderMap_(catSh);

    const lastRow = catSh.getLastRow();
    idx = {};

    if (lastRow >= 2) {
      const skuCol      = catMap['SKU'];
      const prodCol     = catMap['Product Name'];
      const variantCol  = catMap['Variant / Color'];
      const brandCol    = catMap['Brand'];
      const defPriceCol = catMap['Default Price (EGP)'];
      const defCostCol  = catMap['Default Cost (EGP)'];

      if (skuCol) {
        const data = catSh
          .getRange(2, 1, lastRow - 1, catSh.getLastColumn())
          .getValues();

        for (let i = 0; i < data.length; i++) {
          const rowSku = normalizeSku(data[i][skuCol - 1]);
          if (!rowSku) continue;
          if (!idx[rowSku]) {
            idx[rowSku] = {
              productName:     prodCol     ? (data[i][prodCol - 1] || '') : '',
              variant:         variantCol  ? (data[i][variantCol - 1] || '') : '',
              defaultPriceEgp: defPriceCol ? Number(data[i][defPriceCol - 1] || 0) : null,
              defaultCostEgp:  defCostCol  ? Number(data[i][defCostCol - 1] || 0) : 0,
              brand:           brandCol    ? (data[i][brandCol - 1] || '') : ''
            };
          }
        }
      }
    }

    cache.put(cacheKey, JSON.stringify(idx), 300);
  }

  return idx[target] || null;
}

/**
 * يعيد حساب:
 *  - Total Price (EGP) = Qty * Unit Price
 *  - Net Revenue (EGP) = Total - Discount - Shipping Fee
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} map header map بتاع الشيت
 * @param {number} row رقم الصف المراد حسابه
 */
function recalcSalesRowAmounts_(sheet, map, row) {
  const qtyCol       = map[APP.COLS.SALES_EG.QTY];
  const unitPriceCol = map[APP.COLS.SALES_EG.UNIT_PRICE];
  const totalCol     = map[APP.COLS.SALES_EG.TOTAL_PRICE];
  const discountCol  = map[APP.COLS.SALES_EG.DISCOUNT];
  const shipFeeCol   = map[APP.COLS.SALES_EG.SHIPPING_FEE];
  const netRevCol    = map[APP.COLS.SALES_EG.NET_REVENUE];

  if (!qtyCol || !unitPriceCol) return;

  const qty       = Number(sheet.getRange(row, qtyCol).getValue()       || 0);
  const unitPrice = Number(sheet.getRange(row, unitPriceCol).getValue() || 0);
  const discount  = discountCol ? Number(sheet.getRange(row, discountCol).getValue() || 0) : 0;
  const shipFee   = shipFeeCol  ? Number(sheet.getRange(row, shipFeeCol).getValue()  || 0) : 0;

  const total = qty * unitPrice;
  if (totalCol) {
    sheet.getRange(row, totalCol).setValue(total || '');
  }

  if (netRevCol) {
    const net = total - discount - shipFee;
    sheet.getRange(row, netRevCol).setValue(net || '');
  }
}

function testCatalogLookup() {
  const sku = 'MONSTER-MQT65-BEIGE'; // أو أي SKU من اللي في الكتالوج
  const info = catalogLookupBySku_(sku);
  Logger.log(JSON.stringify(info));
}

/**
 * يملي Product Name / Variant / Unit Price
 * لكل صف في Sales_EG عنده SKU ومفيهوش البيانات دي.
 * تقدر تشغلها من Run أو نضيف لها Menu بعدين.
 */
function salesEgBackfillFromCatalog() {
  try {
    const sh  = getSheet_(APP.SHEETS.SALES_EG);
    const map = getHeaderMap_(sh);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('Sales_EG: لا يوجد بيانات.');
      return;
    }

    const skuCol       = map[APP.COLS.SALES_EG.SKU];
    const prodCol      = map[APP.COLS.SALES_EG.PRODUCT_NAME];
    const variantCol   = map[APP.COLS.SALES_EG.VARIANT];
    const unitPriceCol = map[APP.COLS.SALES_EG.UNIT_PRICE];

    const rng = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn());
    const values = rng.getValues();

    for (let r = 0; r < values.length; r++) {
      const row = values[r];
      const sku = String(row[skuCol - 1] || '').trim();
      if (!sku) continue;

      const cat = catalogLookupBySku_(sku);
      if (!cat) continue;

      if (prodCol && !row[prodCol - 1]) {
        row[prodCol - 1] = cat.productName;
      }
      if (variantCol && !row[variantCol - 1]) {
        row[variantCol - 1] = cat.variant;
      }
      if (unitPriceCol && !row[unitPriceCol - 1] && cat.defaultPriceEgp != null) {
        row[unitPriceCol - 1] = cat.defaultPriceEgp;
      }
    }

    rng.setValues(values);

    SpreadsheetApp.getUi().alert('تم ملء بيانات المنتج من الكتالوج بنجاح ✅');

  } catch (err) {
    logError_('salesEgBackfillFromCatalog', err);
    throw err;
  }
}
