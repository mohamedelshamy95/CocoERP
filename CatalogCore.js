/** =============================================================
 * CatalogEg.gs – Product Catalog (EG)
 * CocoERP v2.1
 * -------------------------------------------------------------
 * - Sheet: Catalog_EG
 * - Role : Master data للمنتجات في السوق المصري
 * - يعتمد على:
 *    - APP.SHEETS.CATALOG_EG
 *    - APP.SHEETS.INVENTORY_EG
 *    - APP.SHEETS.PURCHASES
 *    - الهيلبرز: getSheet_, getOrCreateSheet_, getHeaderMap_,
 *                assertRequiredColumns_, logError_, _setupSheetWithHeaders_
 * ============================================================= */

/** Catalog headers (ثابتة) */
const CATALOG_HEADERS = [
  'SKU',
  'Product Name',
  'Variant / Color',
  'Color Group',
  'Brand',
  'Category',
  'Subcategory',
  'Status',
  'Default Cost (EGP)',
  'Default Price (EGP)',
  'Barcode',
  'Notes'
];

/** اختيارات الـ Drop-down */
const CATALOG_CATEGORY_OPTIONS = [
  'Earphones',
  'Earphone Accessories',
  'Speakers',
  'Kidswear',
  'Other'
];

const CATALOG_SUBCATEGORY_OPTIONS = [
  'Ear-cuff Bundle',        // سماعة + إكسسوار (المنتج الأساسي)
  'Ear-cuff Only',          // الإكسسوار لوحده
  'Charm / Crystal Add-on', // قطع كريستال إضافية
  'Case',
  'Cleaning Kit',
  'Cable',
  'Gift Box',
  'Bundle Pack (Multi Items)',
  'Other'
];

const CATALOG_STATUS_OPTIONS = [
  'Active',
  'Inactive',
  'Test'
];

/** =============================================================
 * Setup – إنشاء شيت الكتالوج و الهيدر + Data Validation
 * ============================================================= */

/**
 * تجهيز شيت Catalog_EG:
 *  - يتأكد إن الشيت موجود
 *  - يكتب الهيدر الموحد
 *  - يضيف الـ Drop-down للـ Category / Subcategory / Status
 */
function setupCatalogEgLayout() {
  try {
    const sh =
      (typeof getOrCreateSheet_ === 'function')
        ? getOrCreateSheet_(APP.SHEETS.CATALOG_EG)
        : getSheet_(APP.SHEETS.CATALOG_EG);

    // هيدر موحّد
    if (typeof _setupSheetWithHeaders_ === 'function') {
      _setupSheetWithHeaders_(sh, CATALOG_HEADERS);
    } else {
      const lastRow = sh.getLastRow();
      if (lastRow === 0) {
        sh.getRange(1, 1, 1, CATALOG_HEADERS.length).setValues([CATALOG_HEADERS]);
      }
    }

    // إضافة الـ Data Validation للأعمدة المطلوبة
    catalog_applyDataValidation_(sh);

    safeAlert_('Catalog_EG sheet is ready ✔️');

  } catch (e) {
    logError_('setupCatalogEgLayout', e);
    throw e;
  }
}

/**
 * إضافة / تحديث الـ Drop-down للـ Category / Subcategory / Status
 * ممكن نستدعيها بعد كل Sync للتأكد إن أي صفوف جديدة معاها نفس الفاليديشن.
 */
function catalog_applyDataValidation_(sh) {
  const map = getHeaderMap_(sh);
  const maxRows = sh.getMaxRows();
  if (maxRows < 2) return;
  const numRows = maxRows - 1;

  // Helper داخلي
  function applyListValidation(colIndex, options) {
    if (!colIndex || colIndex <= 0) return;

    const optionsWithBlank = [''].concat(options); // يسمح بترك الخانة فاضية

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(optionsWithBlank, true)
      .setAllowInvalid(false)
      .build();

    sh.getRange(2, colIndex, numRows, 1).setDataValidation(rule);
  }

  // Category
  applyListValidation(map['Category'], CATALOG_CATEGORY_OPTIONS);

  // Subcategory
  applyListValidation(map['Subcategory'], CATALOG_SUBCATEGORY_OPTIONS);

  // Status
  applyListValidation(map['Status'], CATALOG_STATUS_OPTIONS);
}

/** =============================================================
 * Sync – سحب بيانات من Inventory_EG + Purchases للكتالوج
 * ============================================================= */

/**
 * Sync Catalog_EG from:
 *  - Inventory_EG: Product Name, Variant / Color, Avg Cost
 *  - Purchases: Brand = Seller Name
 *
 * قواعد التحديث:
 *  - لو SKU مش موجود في الكتالوج → نضيف صف جديد.
 *  - لو SKU موجود:
 *      * Product Name / Variant / Brand → نملأها فقط لو فاضية.
 *      * Default Cost (EGP) → دايمًا يتحدث من Avg Cost (EGP) في Inventory_EG.
 *  - باقي الأعمدة (Color Group, Category, Subcategory,
 *    Default Price, Barcode, Notes) ما بنلمسهاش.
 */
function catalog_syncFromInventoryEg() {
  try {
    const catalogSh = getSheet_(APP.SHEETS.CATALOG_EG);
    const invSh = getSheet_(APP.SHEETS.INVENTORY_EG);
    const purchSh = getSheet_(APP.SHEETS.PURCHASES);

    // ===== Validate & maps =====
    const catMap = assertRequiredColumns_(catalogSh, CATALOG_HEADERS);

    const maxHeaderCol = Math.max.apply(null, CATALOG_HEADERS.map(h => catMap[h] || 0));

    const invMap = assertRequiredColumns_(invSh, [
      'SKU',
      'Product Name',
      'Variant / Color',
      'Avg Cost (EGP)'
    ]);

    const purchMap = getHeaderMap_(purchSh);

    // ===== Build Brand map من Purchases (Brand = Seller Name) =====
    const brandBySku = {};
    const lastPurRow = purchSh.getLastRow();

    if (lastPurRow >= 2 &&
      purchMap[APP.COLS.PURCHASES.SKU] &&
      purchMap[APP.COLS.PURCHASES.SELLER]) {

      const purData = purchSh
        .getRange(2, 1, lastPurRow - 1, purchSh.getLastColumn())
        .getValues();

      const idxSku = purchMap[APP.COLS.PURCHASES.SKU] - 1;
      const idxSeller = purchMap[APP.COLS.PURCHASES.SELLER] - 1;

      purData.forEach(function (row) {
        const sku = (row[idxSku] || '').toString().trim();
        if (!sku) return;

        const seller = (row[idxSeller] || '').toString().trim();
        if (!seller) return;

        if (!brandBySku[sku]) {
          brandBySku[sku] = seller;
        }
      });
    }

    // ===== Build Inventory snapshot map by SKU =====
    const lastInvRow = invSh.getLastRow();
    if (lastInvRow < 2) {
      safeAlert_('No data in Inventory_EG to sync from.');
      return;
    }

    const invData = invSh
      .getRange(2, 1, lastInvRow - 1, invSh.getLastColumn())
      .getValues();

    const idxInvSku = invMap['SKU'] - 1;
    const idxInvProd = invMap['Product Name'] - 1;
    const idxInvVar = invMap['Variant / Color'] - 1;
    const idxInvAvg = invMap['Avg Cost (EGP)'] - 1;

    const invBySku = {};
    invData.forEach(function (row) {
      const sku = (row[idxInvSku] || '').toString().trim();
      if (!sku) return;

      if (!invBySku[sku]) {
        invBySku[sku] = {
          productName: row[idxInvProd] || '',
          variant: row[idxInvVar] || '',
          avgCost: Number(row[idxInvAvg] || 0)
        };
      }
    });

    if (Object.keys(invBySku).length === 0) {
      safeAlert_('No SKUs found in Inventory_EG to sync.');
      return;
    }

    // ===== Read current catalog data =====
    const lastCatRow = catalogSh.getLastRow();
    const totalCols = maxHeaderCol;

    let catalogData = [];
    if (lastCatRow >= 2) {
      catalogData = catalogSh
        .getRange(2, 1, lastCatRow - 1, totalCols)
        .getValues();
    }

    const idxCatSku = catMap['SKU'] - 1;
    const idxCatProd = catMap['Product Name'] - 1;
    const idxCatVar = catMap['Variant / Color'] - 1;
    const idxCatBrand = catMap['Brand'] - 1;
    const idxCatDefCost = catMap['Default Cost (EGP)'] - 1;

    const catRowBySku = {};
    catalogData.forEach(function (row, i) {
      const sku = (row[idxCatSku] || '').toString().trim();
      if (sku && !catRowBySku[sku]) {
        catRowBySku[sku] = i;
      }
    });

    // لا نحتاج headersOrdered طالما سنبني صفًا بطول totalCols ونضع القيم حسب مواقع الأعمدة الفعلية.

    let added = 0;
    let updated = 0;

    // ===== Sync loop =====
    Object.keys(invBySku).forEach(function (sku) {
      const invInfo = invBySku[sku];
      const brand = brandBySku[sku] || '';
      const avgCost = invInfo.avgCost || 0;

      if (catRowBySku.hasOwnProperty(sku)) {
        // تحديث صف موجود
        const rowIdx = catRowBySku[sku];
        const row = catalogData[rowIdx];

        if (!row[idxCatProd] && invInfo.productName) {
          row[idxCatProd] = invInfo.productName;
        }
        if (!row[idxCatVar] && invInfo.variant) {
          row[idxCatVar] = invInfo.variant;
        }
        if (!row[idxCatBrand] && brand) {
          row[idxCatBrand] = brand;
        }

        if (avgCost) {
          row[idxCatDefCost] = avgCost;
        }

        updated++;

      } else {
        // إضافة SKU جديد
        const newRow = new Array(totalCols).fill('');
        newRow[idxCatSku] = sku;
        if (idxCatProd >= 0) newRow[idxCatProd] = invInfo.productName || '';
        if (idxCatVar >= 0) newRow[idxCatVar] = invInfo.variant || '';
        if (idxCatBrand >= 0) newRow[idxCatBrand] = brand || '';
        if (idxCatDefCost >= 0) newRow[idxCatDefCost] = avgCost || '';

        catalogData.push(newRow);
        added++;
      }
    });

    // ===== Write back =====
    if (lastCatRow > 1) {
      catalogSh
        .getRange(2, 1, lastCatRow - 1, totalCols)
        .clearContent();
    }

    if (catalogData.length) {
      catalogSh
        .getRange(2, 1, catalogData.length, totalCols)
        .setValues(catalogData);
    }

    // نضمن إن الـ Drop-down موجودة لأي صفوف جديدة
    catalog_applyDataValidation_(catalogSh);

    safeAlert_(
      'Catalog sync done.\n\n' +
      'New SKUs added: ' + added + '\n' +
      'Existing SKUs updated: ' + updated
    );

  } catch (e) {
    logError_('catalog_syncFromInventoryEg', e);
    throw e;
  }
}
