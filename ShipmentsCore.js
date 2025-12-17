/** =============================================================
 * ShipmentsCore.gs – Shipments + QC + Inventory Integration
 * CocoERP v2.1
 *
 * Depends on:
 *  - APP.SHEETS.SHIP_CN_UAE   → 'Shipments_CN_UAE'
 *  - APP.SHEETS.SHIP_UAE_EG   → 'Shipments_UAE_EG'
 *  - APP.SHEETS.QC_UAE        → 'QC_UAE'
 *  - APP.SHEETS.PURCHASES     → 'Purchases'
 *  - APP.SHEETS.INVENTORY_TXNS → 'Inventory_Transactions'
 *  - Helpers: getSheet_, getHeaderMap_, logError_, logInventoryTxn_,
 *             inv_rebuildAllSnapshots
 *
 * NOTE:
 *  - This file هو مركز الشحن بالكامل:
 *      • CN→UAE Shipments (status + totals + sync from Purchases)
 *      • UAE→EG Shipments (status + totals + UI integration)
 *      • QC_UAE generation + recalc
 *      • QC_UAE → Inventory Ledger (IN to UAE warehouses)
 *      • Shipments_UAE_EG → Inventory Ledger (OUT from UAE warehouses, IN to Egypt)
 *
 *  - InventoryCore3.gs مسئول عن:
 *      • Inventory_Transactions ledger
 *      • Inventory_UAE / Inventory_EG snapshots
 *      • Catalog + basic helpers
 * ============================================================= */

/** Unified shipment status constants (used in both CN→UAE & UAE→EG) */
const SHIPMENT_STATUS = {
  PLANNED:     'Planned',
  IN_TRANSIT:  'In Transit',
  DELAYED:     'Delayed',
  ARRIVED_UAE: 'Arrived UAE',
  ARRIVED_EG:  'Arrived EG'
};

/* ===================================================================
 * CN → UAE – Status + Totals
 * =================================================================== */

/**
 * Update Shipments_CN_UAE:
 *  - Total Cost (AED) = Freight (AED) + Other Fees (AED) if any of them is set.
 *  - Status:
 *      • If Actual Arrival exists → Arrived UAE
 *      • Else if Ship Date + ETA:
 *            - ETA < today       → Delayed
 *            - otherwise         → In Transit
 *      • Else if Ship Date only  → In Transit
 *      • Else                    → Planned
 *
 *  - Ignores empty rows (no Shipment ID) and clears Status + Total Cost.
 */
function updateShipmentsCnUaeStatusAndTotals() {
  try {
    const sh = getSheet_(APP.SHEETS.SHIP_CN_UAE);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No data found in Shipments_CN_UAE.');
      return;
    }

    const map = getHeaderMap_(sh);

    const colShipmentId = map[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || map['Shipment ID'];
    const colShipDate   = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE]   || map['Ship Date'];
    const colEta        = map[APP.COLS.SHIP_CN_UAE.ETA]         || map['ETA'];
    const colArrival    = map[APP.COLS.SHIP_CN_UAE.ARRIVAL]     || map['Actual Arrival'];
    const colStatus     = map[APP.COLS.SHIP_CN_UAE.STATUS]      || map['Status'];
    const colFreight    = map[APP.COLS.SHIP_CN_UAE.FREIGHT]     || map['Freight (AED)'];
    const colOther      = map[APP.COLS.SHIP_CN_UAE.OTHER]       || map['Other Fees (AED)'];
    const colTotal      = map[APP.COLS.SHIP_CN_UAE.TOTAL_COST]  || map['Total Cost (AED)'];

    if (!colShipmentId || !colStatus) {
      SpreadsheetApp.getUi().alert('Missing required headers in Shipments_CN_UAE (Shipment ID / Status).');
      return;
    }

    const numRows = lastRow - 1;
    const lastCol = sh.getLastColumn();
    const data = sh.getRange(2, 1, numRows, lastCol).getValues();

    const today = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    const statusOut = [];
    const totalOut = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const shipmentId = row[colShipmentId - 1];

      // Empty row → clear status + total
      if (!shipmentId) {
        statusOut.push(['']);
        totalOut.push(['']);
        continue;
      }

      // ----- Total Cost (AED) -----
      let totalVal = '';
      if (colTotal && colFreight) {
        const rawFreight = row[colFreight - 1];
        const rawOther   = colOther ? row[colOther - 1] : '';

        if (rawFreight === '' && (rawOther === '' || rawOther === undefined)) {
          totalVal = '';
        } else {
          const freight = Number(rawFreight || 0);
          const other   = colOther ? Number(rawOther || 0) : 0;
          totalVal = freight + other;
        }
      }

      // ----- Status -----
      const shipDate = colShipDate ? row[colShipDate - 1] : null;
      const eta      = colEta      ? row[colEta - 1]      : null;
      const arr      = colArrival  ? row[colArrival - 1]  : null;

      let status;
      if (arr) {
        status = SHIPMENT_STATUS.ARRIVED_UAE;
      } else if (shipDate && eta) {
        const etaMid = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());
        status = (etaMid < todayMid) ? SHIPMENT_STATUS.DELAYED : SHIPMENT_STATUS.IN_TRANSIT;
      } else if (shipDate) {
        status = SHIPMENT_STATUS.IN_TRANSIT;
      } else {
        status = SHIPMENT_STATUS.PLANNED;
      }

      statusOut.push([status]);
      totalOut.push([totalVal]);
    }

    // Write back only the computed columns (faster + safer)
    sh.getRange(2, colStatus, numRows, 1).setValues(statusOut);

    if (colTotal) {
      sh.getRange(2, colTotal, numRows, 1).setValues(totalOut);
      sh.getRange(2, colTotal, numRows, 1).setNumberFormat('0.00');
    }

    SpreadsheetApp.getUi().alert('Shipments_CN_UAE updated (status + totals) ✔️');
  } catch (e) {
    logError_('updateShipmentsCnUaeStatusAndTotals', e);
    throw e;
  }
}

/**
 * Update Status + Total Cost for a single row in Shipments_CN_UAE.
 *
 * Rules:
 *  - If no Shipment ID → clear Status and Total Cost.
 *  - Total Cost (AED) = Freight (AED) + Other Fees (AED) if any is set.
 *  - Status:
 *      • If Actual Arrival → Arrived UAE
 *      • If Ship Date + ETA → Delayed / In Transit
 *      • If Ship Date only → In Transit
 *      • Otherwise → Planned
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowIndex 1-based row index
 * @param {Object<string, number>=} headerMap optional header map to reuse
 */
function _updateShipmentCnUaeStatusForRow_(sh, rowIndex, headerMap) {
  const map = headerMap || getHeaderMap_(sh);

  const colShipmentId = map[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || map['Shipment ID'];
  const colShipDate   = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE]   || map['Ship Date'];
  const colEta        = map[APP.COLS.SHIP_CN_UAE.ETA]         || map['ETA'];
  const colArrival    = map[APP.COLS.SHIP_CN_UAE.ARRIVAL]     || map['Actual Arrival'];
  const colStatus     = map[APP.COLS.SHIP_CN_UAE.STATUS]      || map['Status'];
  const colFreight    = map[APP.COLS.SHIP_CN_UAE.FREIGHT]     || map['Freight (AED)'];
  const colOther      = map[APP.COLS.SHIP_CN_UAE.OTHER]       || map['Other Fees (AED)'];
  const colTotal      = map[APP.COLS.SHIP_CN_UAE.TOTAL_COST]  || map['Total Cost (AED)'];

  if (!colShipmentId || !colStatus) {
    // Required columns missing (header renamed or incomplete layout)
    return;
  }

  const lastCol = sh.getLastColumn();
  const row = sh.getRange(rowIndex, 1, 1, lastCol).getValues()[0];

  const shipmentId = row[colShipmentId - 1];

  // Empty row → clear Status + Total Cost (if present) and exit
  if (!shipmentId) {
    if (colStatus) sh.getRange(rowIndex, colStatus).clearContent();
    if (colTotal)  sh.getRange(rowIndex, colTotal).clearContent();
    return;
  }

  // ----- Total Cost (AED) -----
  if (colTotal && colFreight) {
    const rawFreight = row[colFreight - 1];
    const rawOther   = colOther ? row[colOther - 1] : '';

    // If no numbers in Freight / Other → leave Total Cost empty
    if (rawFreight === '' && (rawOther === '' || rawOther === undefined)) {
      row[colTotal - 1] = '';
    } else {
      const freight = Number(rawFreight || 0);
      const other   = colOther ? Number(rawOther || 0) : 0;
      row[colTotal - 1] = freight + other;
    }
  }

  // ----- Status -----
  const ship = colShipDate ? row[colShipDate - 1] : null;
  const eta  = colEta      ? row[colEta - 1]      : null;
  const arr  = colArrival  ? row[colArrival - 1]  : null;

  let status;

  if (arr) {
    status = SHIPMENT_STATUS.ARRIVED_UAE;
  } else if (ship && eta) {
    const today    = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const etaMid   = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());

    status = etaMid < todayMid
      ? SHIPMENT_STATUS.DELAYED
      : SHIPMENT_STATUS.IN_TRANSIT;
  } else if (ship) {
    status = SHIPMENT_STATUS.IN_TRANSIT;
  } else {
    status = SHIPMENT_STATUS.PLANNED;
  }

  row[colStatus - 1] = status;

  sh.getRange(rowIndex, 1, 1, lastCol).setValues([row]);
}

/**
 * Handle edits in Shipments_CN_UAE:
 * - If one of the key columns changes → recalculate Status + Total for that row:
 *    Shipment ID / Ship Date / ETA / Actual Arrival / Freight / Other Fees
 *
 * Intended to be called from the global onEdit(e) dispatcher in AppCore3.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function shipmentsCnUaeOnEdit_(e) {
  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    if (sheetName !== APP.SHEETS.SHIP_CN_UAE && sheetName !== 'Shipments_CN_UAE') {
      return;
    }

    const editedCol = e.range.getColumn();
    const editedRow = e.range.getRow();
    if (editedRow === 1) return; // header row

    const map = getHeaderMap_(sh);

    const colShipmentId = map[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || map['Shipment ID'];
    const colShipDate   = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE]   || map['Ship Date'];
    const colEta        = map[APP.COLS.SHIP_CN_UAE.ETA]         || map['ETA'];
    const colArrival    = map[APP.COLS.SHIP_CN_UAE.ARRIVAL]     || map['Actual Arrival'];
    const colFreight    = map[APP.COLS.SHIP_CN_UAE.FREIGHT]     || map['Freight (AED)'];
    const colOther      = map[APP.COLS.SHIP_CN_UAE.OTHER]       || map['Other Fees (AED)'];

    if (!colShipmentId) return;

    const interestingCols = [
      colShipmentId,
      colShipDate,
      colEta,
      colArrival,
      colFreight,
      colOther
    ].filter(function (c) { return !!c; });

    if (interestingCols.indexOf(editedCol) === -1) {
      return;
    }

    _updateShipmentCnUaeStatusForRow_(sh, editedRow, map);
  } catch (err) {
    logError_('shipmentsCnUaeOnEdit_', err, {
      a1: e && e.range ? e.range.getA1Notation() : ''
    });
  }
}

/**
 * Apply Data Validation on Status column in Shipments_CN_UAE
 * using SHIPMENT_STATUS constants. Run once after layout is ready.
 */
function setupShipmentsCnUaeStatusValidation_() {
  try {
    const sh = getSheet_(APP.SHEETS.SHIP_CN_UAE);
    const map = getHeaderMap_(sh);

    const colStatus = map[APP.COLS.SHIP_CN_UAE.STATUS] || map['Status'];
    if (!colStatus) return;

    const allowedStatuses = [
      SHIPMENT_STATUS.PLANNED,
      SHIPMENT_STATUS.IN_TRANSIT,
      SHIPMENT_STATUS.DELAYED,
      SHIPMENT_STATUS.ARRIVED_UAE
    ];

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(allowedStatuses, true) // dropdown + free typing from same list
      .setAllowInvalid(false)
      .build();

    const maxRows = sh.getMaxRows();
    if (maxRows <= 1) return;

    sh.getRange(2, colStatus, maxRows - 1, 1).setDataValidation(rule);
  } catch (e) {
    logError_('setupShipmentsCnUaeStatusValidation_', e);
    throw e;
  }
}

/**
 * Rebuild Status + Totals for all rows in Shipments_CN_UAE.
 * Uses the same helper as the onEdit handler.
 */
function rebuildShipmentsCnUaeStatus_() {
  try {
    const sh = getSheet_(APP.SHEETS.SHIP_CN_UAE);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const map = getHeaderMap_(sh);
    for (let r = 2; r <= lastRow; r++) {
      _updateShipmentCnUaeStatusForRow_(sh, r, map);
    }
  } catch (e) {
    logError_('rebuildShipmentsCnUaeStatus_', e);
    throw e;
  }
}

/* ===================================================================
 * UAE → EG – Status + Totals (sheet-level)
 * =================================================================== */

/**
 * Update Shipments_UAE_EG:
 * - Total Cost (EGP) = Ship Cost (EGP) * Qty + Customs (EGP) + Other (EGP)
 *   (Ship Cost is treated as cost per unit or per box, depending on
 *    how the user enters it. In all cases, formula is cost * Qty + customs + other.)
 * - Status:
 *    * If Actual Arrival => Arrived EG
 *    * If Ship Date + ETA => Delayed / In Transit
 *    * If Ship Date only => In Transit
 *    * Otherwise => Planned
 */
function updateShipmentsUaeEgStatusAndTotals() {
  try {
    const sh  = getSheet_(APP.SHEETS.SHIP_UAE_EG);
    const map = getHeaderMap_(sh);

    const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
                        map['Ship Date'] ||
                        map['Ship Date (UAE)'];

    const colEta      = map[APP.COLS.SHIP_UAE_EG.ETA] ||
                        map['ETA'];

    const colArr      = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
                        map['Actual Arrival'] ||
                        map['Actual Arrival (EG)'];

    const colStatus   = map[APP.COLS.SHIP_UAE_EG.STATUS] ||
                        map['Status'];

    const colQty      = map[APP.COLS.SHIP_UAE_EG.QTY] ||
                        map['Qty'];

    const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
                        map['Ship Cost (EGP) - per unit or box'] ||
                        map['Ship Cost (EGP)'];

    const colCustoms  = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
                        map['Customs (EGP)'] ||
                        map['Customs / Clearance (EGP)'];

    const colOther    = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
                        map['Other (EGP)'] ||
                        map['Other Fees (EGP)'];

    const colTotal    = map[APP.COLS.SHIP_UAE_EG.TOTAL_COST] ||
                        map['Total Cost (EGP)'];

    const requiredCols = [
      colShipDate, colEta, colArr, colStatus,
      colQty, colShipCost, colCustoms, colOther, colTotal
    ];

    if (requiredCols.some(function (c) { return !c; })) {
      throw new Error('Missing one or more required columns in Shipments_UAE_EG.');
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      SpreadsheetApp.getUi().alert('No data found in Shipments_UAE_EG.');
      return;
    }

    const numRows = lastRow - 1;
    const range   = sh.getRange(2, 1, numRows, sh.getLastColumn());
    const values  = range.getValues();

    const idx = {
      shipDate: colShipDate - 1,
      eta:      colEta      - 1,
      arr:      colArr      - 1,
      status:   colStatus   - 1,
      qty:      colQty      - 1,
      shipCost: colShipCost - 1,
      customs:  colCustoms  - 1,
      other:    colOther    - 1,
      total:    colTotal    - 1
    };

    const today    = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    values.forEach(function (row) {
      // ----- Total Cost -----
      const qty             = Number(row[idx.qty]      || 0);
      const shipCostPerUnit = Number(row[idx.shipCost] || 0);
      const customs         = Number(row[idx.customs]  || 0);
      const other           = Number(row[idx.other]    || 0);

      const totalForShipment = (shipCostPerUnit * qty) + customs + other;
      row[idx.total] = totalForShipment;

      // ----- Status -----
      const ship = row[idx.shipDate];
      const eta  = row[idx.eta];
      const arr  = row[idx.arr];

      let status;
      if (arr) {
        status = SHIPMENT_STATUS.ARRIVED_EG;
      } else if (ship && eta) {
        const etaMid = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());
        status = etaMid < todayMid
          ? SHIPMENT_STATUS.DELAYED
          : SHIPMENT_STATUS.IN_TRANSIT;
      } else if (ship) {
        status = SHIPMENT_STATUS.IN_TRANSIT;
      } else {
        status = SHIPMENT_STATUS.PLANNED;
      }

      row[idx.status] = status;
    });

    range.setValues(values);

    // Ensure number format for Total Cost (EGP)
    sh.getRange(2, colTotal, numRows, 1).setNumberFormat('0.00');

    SpreadsheetApp.getUi().alert('Shipments_UAE_EG updated (status + totals) ✔️');
  } catch (e) {
    logError_('updateShipmentsUaeEgStatusAndTotals', e);
    throw e;
  }
}

/**
 * Convenience: update all shipments at once (CN→UAE + UAE→EG).
 * Used from the menu: Logistics & Inventory → Update Shipments Status & Totals
 */
function updateAllShipmentsStatusAndTotals() {
  updateShipmentsCnUaeStatusAndTotals();
  updateShipmentsUaeEgStatusAndTotals();
}

/* ===================================================================
 * Sync Purchases → Shipments_CN_UAE
 * =================================================================== */

/**
 * Sync Purchases -> Shipments_CN_UAE
 *
 * - For each row in Purchases that has invoice signals
 *    (Invoice Link / Invoice Preview / Order Total EGP > 0):
 *    create a row in Shipments_CN_UAE if not already present.
 * - De-dup key: {Order ID} + {SKU}
 * - Each Order ID gets one Shipment ID, reused for all SKUs.
 * - Shipment ID format: CN-000001, CN-000002, ...
 */

/**
 * Sync Purchases -> Shipments_CN_UAE (UPSERT + AGGREGATION)
 *
 * Why this version:
 * - Some orders contain multiple Purchases rows with the same Order ID + SKU (e.g. split lines).
 * - Old logic used a de-dup key and SKIPPED duplicates, so Shipments/QC qty became wrong.
 *
 * Behavior:
 * - Aggregates Purchases Qty by key: Order ID + SKU + Variant (fallback to empty Variant)
 * - If row exists in Shipments_CN_UAE => updates Qty (and optionally fills blanks)
 * - If not exists => inserts a new line row
 * - Shipment ID: one per Order ID (CN-000001 sequence), reused across that order.
 */
function syncPurchasesToShipmentsCnUae() {
  try {
    const purchSh = getSheet_(APP.SHEETS.PURCHASES);
    const shipSh  = getSheet_(APP.SHEETS.SHIP_CN_UAE);

    const pMap = getHeaderMap_(purchSh);
    const sMap = getHeaderMap_(shipSh);

    const lastPurRow = purchSh.getLastRow();
    if (lastPurRow < 2) {
      safeAlert_('No data found in Purchases.');
      return;
    }

    // Purchases columns (robust)
    const colOrderId   = pMap[APP.COLS.PURCHASES.ORDER_ID]   || pMap['Order ID'];
    const colOrderDate = pMap[APP.COLS.PURCHASES.ORDER_DATE] || pMap['Order Date'];
    const colPlatform  = pMap[APP.COLS.PURCHASES.PLATFORM]   || pMap['Platform'];
    const colSeller    = pMap[APP.COLS.PURCHASES.SELLER]     || pMap['Seller Name'];
    const colSku       = pMap[APP.COLS.PURCHASES.SKU]        || pMap['SKU'];
    const colProduct   = pMap[APP.COLS.PURCHASES.PRODUCT_NAME] || pMap[APP.COLS.PURCHASES.PRODUCT] || pMap['Product Name'];
    const colVariant   = pMap[APP.COLS.PURCHASES.VARIANT]      || pMap['Variant / Color'];
    const colQty       = pMap[APP.COLS.PURCHASES.QTY]        || pMap['Qty'];
    const colLineId    = pMap[APP.COLS.PURCHASES.LINE_ID]    || pMap['Line ID'];

    // Invoice indicators (to decide "ready to ship")
    const colInvoiceLink   = pMap['Invoice Link'];
    const colInvoicePrev   = pMap['Invoice Preview'];
    const colOrderTotalEgp = pMap[APP.COLS.PURCHASES.TOTAL_EGP] || pMap['Order Total (EGP)'];

    if (!colOrderId || !colSku || !colQty) {
      throw new Error('Missing required Purchases columns (Order ID / SKU / Qty).');
    }
    if (!colLineId) {
      throw new Error('Missing Purchases column "Line ID". Run Purchases layout/repair first.');
    }

    // Ensure missing Line IDs are generated (idempotent)
    try {
      if (typeof purchases_ensureLineIds_ === 'function') {
        purchases_ensureLineIds_(purchSh, pMap, 2, lastPurRow - 1);
      }
    } catch (e) {}

    const purchData = purchSh
      .getRange(2, 1, lastPurRow - 1, purchSh.getLastColumn())
      .getValues();

    // Shipments columns
    const shipColId        = sMap[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID]       || sMap['Shipment ID'];
    const shipColOrderId   = sMap[APP.COLS.SHIP_CN_UAE.ORDER_BATCH]       || sMap['Order ID (Batch)'] || sMap['Order ID'];
    const shipColLineId    = sMap[APP.COLS.SHIP_CN_UAE.PURCHASE_LINE_ID]  || sMap['Purchases Line ID'];
    const shipColSku       = sMap[APP.COLS.SHIP_CN_UAE.SKU]               || sMap['SKU'];
    const shipColVariant   = sMap[APP.COLS.SHIP_CN_UAE.VARIANT]           || sMap['Variant / Color'] || sMap[APP.COLS.PURCHASES.VARIANT];
    const shipColQty       = sMap[APP.COLS.SHIP_CN_UAE.QTY]               || sMap[APP.COLS.PURCHASES.QTY] || sMap['Qty'];
    const shipColProd      = sMap[APP.COLS.SHIP_CN_UAE.PRODUCT_NAME]      || sMap['Product Name'];

    if (!shipColOrderId || !shipColSku || !shipColQty || !shipColLineId) {
      throw new Error('Missing required Shipments_CN_UAE columns (Order ID (Batch) / Purchases Line ID / SKU / Qty). Run Logistics → Setup Shipments Layouts.');
    }

    // Existing Shipments map + detect Shipment ID max sequence
    const lastShipRow = shipSh.getLastRow();
    const existingRows = (lastShipRow >= 2)
      ? shipSh.getRange(2, 1, lastShipRow - 1, shipSh.getLastColumn()).getValues()
      : [];

    /** @type {Object<string, {row:number, shipId:string}>} */
    const existingByLineId = {}; // key: Purchases Line ID -> { row, shipId }
    /** @type {Object<string, string>} */
    const orderToShipmentId = {}; // OrderID -> ShipmentID
    let maxSeq = 0;

    if (existingRows.length) {
      existingRows.forEach(function (row, i) {
        const sheetRow = i + 2;
        const orderId = row[shipColOrderId - 1];
        const lineId  = String(row[shipColLineId - 1] || '').trim();
        const shipId  = shipColId ? String(row[shipColId - 1] || '').trim() : '';

        if (orderId && shipId) orderToShipmentId[String(orderId)] = shipId;

        if (shipId) {
          const m = shipId.match(/(\d+)$/);
          if (m) {
            const n = parseInt(m[1], 10);
            if (n > maxSeq) maxSeq = n;
          }
        }

        if (lineId && !existingByLineId[lineId]) {
          existingByLineId[lineId] = { row: sheetRow, shipId: shipId };
        }
      });
    }

    const shipLastCol = shipSh.getLastColumn();
    const shipHeaders = shipSh.getRange(1, 1, 1, shipLastCol).getValues()[0];

    const newRows = [];
    const qtyUpdates = [];
    const fillUpdates = [];

    // Line-level sync (NO aggregation; no de-dup by OrderID+SKU)
    purchData.forEach(function (r) {
      const orderIdRaw = colOrderId ? r[colOrderId - 1] : '';
      const orderId = String(orderIdRaw || '').trim();
      const sku     = colSku ? String(r[colSku - 1] || '').trim() : '';
      const qty     = colQty ? Number(r[colQty - 1] || 0) : 0;
      const lineId  = colLineId ? String(r[colLineId - 1] || '').trim() : '';

      if (!orderId || !sku || !qty) return;
      if (!lineId) return;

      // Only sync if invoice exists (any of these signals)
      const hasInvoice =
        (colInvoiceLink   && r[colInvoiceLink - 1]) ||
        (colInvoicePrev   && r[colInvoicePrev - 1]) ||
        (colOrderTotalEgp && Number(r[colOrderTotalEgp - 1] || 0) > 0);

      if (!hasInvoice) return;

      const variant = colVariant ? String(r[colVariant - 1] || '').trim() : '';

      // Shipment ID per Order ID
      let shipmentId = orderToShipmentId[orderId];
      if (!shipmentId) {
        maxSeq++;
        shipmentId = 'CN-' + Utilities.formatString('%06d', maxSeq);
        orderToShipmentId[orderId] = shipmentId;
      }

      const existing = existingByLineId[lineId];
      if (existing && existing.row) {
        qtyUpdates.push({ row: existing.row, qty: qty });

        if (shipColProd || shipColVariant) {
          const prod = colProduct ? String(r[colProduct - 1] || '').trim() : '';
          fillUpdates.push({ row: existing.row, product: prod, variant: String(variant || '') });
        }
        return;
      }

      /** @type {Object<string, any>} */
      const rowObj = {};

      rowObj['Shipment ID']          = shipmentId;
      rowObj['Supplier / Factory']   = colSeller ? (r[colSeller - 1] || '') : '';
      rowObj['Forwarder']            = colPlatform ? (r[colPlatform - 1] || '') : '';
      rowObj['Tracking / Container'] = '';
      rowObj['Purchases Line ID']    = lineId;

      rowObj['Order ID (Batch)']     = orderId;

      const orderDate = colOrderDate ? r[colOrderDate - 1] : null;
      rowObj['Ship Date']            = (orderDate instanceof Date) ? orderDate : new Date();
      rowObj['ETA']                  = '';
      rowObj['Actual Arrival']       = '';

      rowObj['SKU']             = sku;
      rowObj['Product Name']    = colProduct ? (r[colProduct - 1] || '') : '';
      rowObj['Variant / Color'] = variant || '';
      rowObj['Qty']             = qty;

      rowObj['Gross Weight (kg)'] = '';
      rowObj['Volume (CBM)']      = '';
      rowObj['Freight (AED)']     = '';
      rowObj['Other Fees (AED)']  = '';
      rowObj['Total Cost (AED)']  = '';

      rowObj['Notes'] = 'Auto (line-level) from Purchases';

      const outRow = shipHeaders.map(function (h) {
        return (rowObj[h] !== undefined) ? rowObj[h] : '';
      });
      newRows.push(outRow);
    });

    if (!newRows.length && !qtyUpdates.length) {
      safeAlert_('No Shipments_CN_UAE changes detected.');
      return;
    }

    // Apply qty updates in batches (contiguous runs)
    let updatedQtyCount = 0;
    if (qtyUpdates.length) {
      qtyUpdates.sort(function (a, b) { return a.row - b.row; });

      const rowToQty = {};
      qtyUpdates.forEach(function (u) { rowToQty[u.row] = u.qty; });

      let i = 0;
      while (i < qtyUpdates.length) {
        const startRow = qtyUpdates[i].row;
        let endRow = startRow;
        while (i + 1 < qtyUpdates.length && qtyUpdates[i + 1].row === endRow + 1) {
          i++;
          endRow = qtyUpdates[i].row;
        }
        const n = endRow - startRow + 1;
        const vals = [];
        for (let r = startRow; r <= endRow; r++) {
          vals.push([rowToQty[r]]);
        }
        shipSh.getRange(startRow, shipColQty, n, 1).setValues(vals);
        updatedQtyCount += n;
        i++;
      }

      // Optional: fill Product/Variant if blank (best-effort)
      try {
        if (fillUpdates.length) {
          fillUpdates.forEach(function (u) {
            const r = u.row;
            if (shipColProd) {
              const cur = shipSh.getRange(r, shipColProd).getValue();
              if (!cur && u.product) shipSh.getRange(r, shipColProd).setValue(u.product);
            }
            if (shipColVariant) {
              const curV = shipSh.getRange(r, shipColVariant).getValue();
              if (!curV && u.variant) shipSh.getRange(r, shipColVariant).setValue(u.variant);
            }
          });
        }
      } catch (e) {}
    }

    // Append new rows
    if (newRows.length) {
      const startRow = shipSh.getLastRow() + 1;
      shipSh.getRange(startRow, 1, newRows.length, shipLastCol).setValues(newRows);
    }

    if (newRows.length || updatedQtyCount) {
      try { rebuildShipmentsCnUaeStatus_(); } catch (e) {}
      try { setupShipmentsCnUaeStatusValidation_(); } catch (e) {}
    }

    safeAlert_(
      'Purchases → Shipments_CN_UAE sync done.\n' +
      'Inserted rows: ' + newRows.length + '\n' +
      'Updated qty rows: ' + updatedQtyCount
    );

  } catch (e) {
    logError_('syncPurchasesToShipmentsCnUae', e);
    throw e;
  }
}

/* ===================================================================
 * Inventory integration (Inventory_UAE ↔ Shipments_UAE_EG)
 * =================================================================== */

/**
 * Helper: حوّل اسم الـ Courier إلى كود مخزن الإمارات.
 * - Attia / عطية → UAE-ATTIA
 * - Kor / الكور → UAE-KOR
 * - لو المستخدم كتب UAE-ATTIA / UAE-KOR مباشرة → نرجعهم كما هم.
 * - غير ذلك → UAE-DXB.
 */
function resolveUaeWarehouseFromCourier_(courierRaw) {
  const s = (courierRaw || '').toString().toLowerCase().trim();
  if (!s) return 'UAE-DXB';

  // لو المستخدم كتب الكود مباشرة
  if (s.indexOf('uae-attia') !== -1) return 'UAE-ATTIA';
  if (s.indexOf('uae-kor')   !== -1) return 'UAE-KOR';

  // أسماء عربية/إنجليزية
  if (s.indexOf('attia') !== -1 || s.indexOf('عطية') !== -1 || s.indexOf('عطيه') !== -1) {
    return 'UAE-ATTIA';
  }
  if (s.indexOf('kor') !== -1 || s.indexOf('الكور') !== -1) {
    return 'UAE-KOR';
  }

  return 'UAE-DXB';
}

/**
 * Read Inventory_UAE info for a given SKU (and optional warehouse).
 * Returns an object or null if not found.
 *
 * @param {string} sku
 * @param {string=} optWarehouse
 * @returns {{productName: string, variant: string, warehouse: string, onHand: number, available: number, avgCost: number} | null}
 */
function _getInventoryUaeInfoForSku_(sku, optWarehouse) {
  try {
    const normalizedSku = (sku || '').toString().trim();
    if (!normalizedSku) return null;

    const invSh = getSheet_(APP.SHEETS.INVENTORY_UAE);
    const map   = getHeaderMap_(invSh);

    const colSku     = map['SKU'];
    const colWh      = map['Warehouse (UAE)'];
    const colProduct = map['Product Name'];
    const colVar     = map['Variant / Color'];
    const colOnHand  = map['On Hand Qty'];
    const colAvail   = map['Available Qty'];
    const colAvgCost = map['Avg Cost (EGP)'];

    if (!colSku || !colWh) return null;

    const lastRow = invSh.getLastRow();
    if (lastRow < 2) return null;

    const data = invSh
      .getRange(2, 1, lastRow - 1, invSh.getLastColumn())
      .getValues();

    const targetWhRaw = (optWarehouse || '').toString().trim();
    const targetWhUpper = targetWhRaw ? targetWhRaw.toUpperCase() : '';

    /** @type {{productName: string, variant: string, warehouse: string, onHand: number, available: number, avgCost: number} | null} */
    let fallbackMatch = null;

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const rowSku = (row[colSku - 1] || '').toString().trim();
      if (!rowSku || rowSku !== normalizedSku) continue;

      const rowWhRaw = (row[colWh - 1] || '').toString().trim();
      const rowWhUpper = rowWhRaw.toUpperCase();

      const info = {
        productName: colProduct ? (row[colProduct - 1] || '') : '',
        variant:     colVar     ? (row[colVar     - 1] || '') : '',
        warehouse:   rowWhRaw,
        onHand:      colOnHand ? Number(row[colOnHand - 1] || 0) : 0,
        available:   colAvail  ? Number(row[colAvail  - 1] || 0) : 0,
        avgCost:     colAvgCost ? Number(row[colAvgCost - 1] || 0) : 0
      };

      // لو محدد Warehouse معين
      if (targetWhUpper) {
        if (rowWhUpper === targetWhUpper) {
          return info; // match perfect
        }
        // غير مطابق → نخليه fallback لو مفيش غيره
        if (!fallbackMatch) {
          fallbackMatch = info;
        }
      } else {
        // مفيش Warehouse محدد → أول صف مطابق للـ SKU يعتبر fallback
        fallbackMatch = info;
        break;
      }
    }

    return fallbackMatch;
  } catch (e) {
    logError_('_getInventoryUaeInfoForSku_', e, { sku: sku, wh: optWarehouse });
    return null;
  }
}

/**
 * Auto-fill one row in Shipments_UAE_EG from Inventory_UAE when SKU is edited.
 *
 * - يحاول يحدد Warehouse (UAE) من نفس الصف الأول.
 * - لو مش موجود، يرجع لـ Courier → Warehouse code.
 * - يملا Product / Variant / Notes + ممكن يملا Warehouse (UAE) و Courier لو فاضيين.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowIndex
 * @param {Object<string, number>=} headerMapOpt
 */
function _fillShipmentUaeEgFromInventory_(sh, rowIndex, headerMapOpt) {
  const map = headerMapOpt || getHeaderMap_(sh);

  const colSku     = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
  const colProd    = map['Product Name'];
  const colVar     = map['Variant / Color'];
  const colNotes   = map['Notes'];
  const colCourier = map['Courier'] || map['Courier Name'];
  const colWhUae   = map['Warehouse (UAE)']; // اختياري في Shipments_UAE_EG

  if (!colSku) return;

  const sku = (sh.getRange(rowIndex, colSku).getValue() || '').toString().trim();

  // If SKU cleared → clear related cells
  if (!sku) {
    if (colProd)  sh.getRange(rowIndex, colProd).clearContent();
    if (colVar)   sh.getRange(rowIndex, colVar).clearContent();
    if (colNotes) sh.getRange(rowIndex, colNotes).setNote('');
    // مش هنلعب في Warehouse (UAE) / Courier هنا
    return;
  }

  // 1) Warehouse hint من نفس الصف لو موجود
  let whHint = '';
  if (colWhUae) {
    const whCell = sh.getRange(rowIndex, colWhUae).getValue();
    whHint = (whCell || '').toString().trim();
  }

  // 2) لو مفيش Warehouse، استخدم Courier كـ hint
  if (!whHint && colCourier) {
    const courierVal = sh.getRange(rowIndex, colCourier).getValue();
    whHint = resolveUaeWarehouseFromCourier_(courierVal);
  }

  // 3) Inventory lookup
  const info = _getInventoryUaeInfoForSku_(sku, whHint);

  if (!info) {
    if (colProd)  sh.getRange(rowIndex, colProd).setValue('');
    if (colVar)   sh.getRange(rowIndex, colVar).setValue('');
    if (colNotes) {
      sh.getRange(rowIndex, colNotes)
        .setNote('SKU not found in Inventory_UAE.');
    }
    return;
  }

  if (colProd) sh.getRange(rowIndex, colProd).setValue(info.productName);
  if (colVar)  sh.getRange(rowIndex, colVar).setValue(info.variant);

  // لو Warehouse (UAE) فاضي في Shipments_UAE_EG → املاه باللي جاي من Inventory_UAE
  if (colWhUae) {
    const currentWh = (sh.getRange(rowIndex, colWhUae).getValue() || '').toString().trim();
    if (!currentWh && (info.warehouse || whHint)) {
      sh.getRange(rowIndex, colWhUae).setValue(info.warehouse || whHint);
    }
  }

  // لو Courier فاضي ومعانا Warehouse واضح → املاه علشان UI يبقى واضح
  if (colCourier) {
    const curCourier = (sh.getRange(rowIndex, colCourier).getValue() || '').toString().trim();
    if (!curCourier) {
      const resolvedWh = (info.warehouse || whHint || '').toUpperCase();
      if (resolvedWh === 'UAE-ATTIA') {
        sh.getRange(rowIndex, colCourier).setValue('Attia');
      } else if (resolvedWh === 'UAE-KOR') {
        sh.getRange(rowIndex, colCourier).setValue('Kor');
      }
    }
  }

  if (colNotes) {
    const labelWh = info.warehouse || whHint || '';
    const note =
      'From Inventory_UAE' +
      (labelWh ? ' [' + labelWh + ']' : '') +
      '\nOn Hand: ' + info.onHand +
      ', Available: ' + info.available +
      ', Avg Cost: ' + info.avgCost + ' EGP';
    sh.getRange(rowIndex, colNotes).setNote(note);
  }
}

/**
 * One-time helper:
 * Fill Product Name / Variant / Notes in Shipments_UAE_EG
 * for all rows that already have a SKU.
 */
function backfillShipmentsUaeEgFromInventory() {
  try {
    const sh  = getSheet_(APP.SHEETS.SHIP_UAE_EG);
    const map = getHeaderMap_(sh);

    const colSku = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
    if (!colSku) {
      throw new Error('SKU column not found in Shipments_UAE_EG.');
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    for (let r = 2; r <= lastRow; r++) {
      const sku = (sh.getRange(r, colSku).getValue() || '').toString().trim();
      if (!sku) continue;
      _fillShipmentUaeEgFromInventory_(sh, r, map);
    }

  } catch (e) {
    logError_('backfillShipmentsUaeEgFromInventory', e);
    throw e;
  }
}

/**
 * Recalculate Total Cost + Status + stock warning for a single row in Shipments_UAE_EG.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowIndex
 * @param {Object<string, number>=} headerMapOpt
 */
function _updateShipmentUaeEgRowTotalsAndStatus_(sh, rowIndex, headerMapOpt) {
  const map = headerMapOpt || getHeaderMap_(sh);

  const colQty      = map[APP.COLS.SHIP_UAE_EG.QTY] ||
                      map['Qty'];
  const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
                      map['Ship Cost (EGP) - per unit or box'] ||
                      map['Ship Cost (EGP)'];
  const colCustoms  = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
                      map['Customs (EGP)'] ||
                      map['Customs / Clearance (EGP)'];
  const colOther    = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
                      map['Other (EGP)'] ||
                      map['Other Fees (EGP)'];
  const colTotal    = map[APP.COLS.SHIP_UAE_EG.TOTAL_COST] ||
                      map['Total Cost (EGP)'];

  const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
                      map['Ship Date'] ||
                      map['Ship Date (UAE)'];
  const colEta      = map[APP.COLS.SHIP_UAE_EG.ETA] ||
                      map['ETA'];
  const colArr      = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
                      map['Actual Arrival'] ||
                      map['Actual Arrival (EG)'];
  const colStatus   = map[APP.COLS.SHIP_UAE_EG.STATUS] ||
                      map['Status'];

  const colSku      = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
  const colNotes    = map['Notes'];
  const colCourier  = map['Courier'] || map['Courier Name'];

  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

  // ----- Total Cost (EGP) -----
  const qty      = colQty      ? Number(row[colQty - 1]      || 0) : 0;
  const shipCost = colShipCost ? Number(row[colShipCost - 1] || 0) : 0;
  const customs  = colCustoms  ? Number(row[colCustoms - 1]  || 0) : 0;
  const other    = colOther    ? Number(row[colOther - 1]    || 0) : 0;

  if (colTotal) {
    const totalShipment = (shipCost * qty) + customs + other;
    sh.getRange(rowIndex, colTotal).setValue(totalShipment);
    sh.getRange(rowIndex, colTotal).setNumberFormat('0.00');
  }

  // ----- Status -----
  const ship = colShipDate ? row[colShipDate - 1] : '';
  const eta  = colEta      ? row[colEta - 1]      : '';
  const arr  = colArr      ? row[colArr - 1]      : '';

  let status = '';

  if (arr) {
    status = SHIPMENT_STATUS.ARRIVED_EG;
  } else if (ship && eta) {
    const today    = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const etaMid   = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());
    status = (etaMid < todayMid)
      ? SHIPMENT_STATUS.DELAYED
      : SHIPMENT_STATUS.IN_TRANSIT;
  } else if (ship) {
    status = SHIPMENT_STATUS.IN_TRANSIT;
  } else {
    status = SHIPMENT_STATUS.PLANNED;
  }

  if (colStatus) {
    sh.getRange(rowIndex, colStatus).setValue(status);
  }

  // ----- Stock check vs correct UAE warehouse (warning only) -----
  if (colSku && colNotes) {
    const sku = (row[colSku - 1] || '').toString().trim();
    const rangeNotes = sh.getRange(rowIndex, colNotes);

    if (sku && qty > 0) {
      let whHint = '';
      if (colCourier) {
        const courierVal = row[colCourier - 1];
        whHint = resolveUaeWarehouseFromCourier_(courierVal);
      }
      const info = _getInventoryUaeInfoForSku_(sku, whHint);
      if (info && info.onHand && qty > info.onHand) {
        const whLabel = info.warehouse || whHint || 'UAE';
        const warn =
          '⚠ Qty ' + qty + ' > on-hand ' + info.onHand +
          ' in ' + whLabel;
        rangeNotes.setNote(warn);
      } else {
        // Clear previous warning note if any
        const note = rangeNotes.getNote();
        if (note) {
          rangeNotes.setNote('');
        }
      }
    }
  }
}

/**
 * Handle edits in Shipments_UAE_EG:
 * - If SKU changes → pull Product / Variant / note from Inventory_UAE.
 * - If Qty or Ship Cost or Customs or Other or Ship/ETA/Arrival changes
 *     → recalc Total + Status + stock warning.
 *
 * Intended to be called from the global onEdit(e) dispatcher in AppCore3.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function shipmentsUaeEgOnEdit_(e) {
  try {
    if (!e || !e.range) return;

    const sh        = e.range.getSheet();
    const sheetName = sh.getName();

    if (sheetName !== APP.SHEETS.SHIP_UAE_EG && sheetName !== 'Shipments_UAE_EG') {
      return;
    }

    const rowIndex = e.range.getRow();
    if (rowIndex === 1) return; // header

    const map = getHeaderMap_(sh);

    const colSku      = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
    const colQty      = map[APP.COLS.SHIP_UAE_EG.QTY] ||
                        map['Qty'];
    const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
                        map['Ship Cost (EGP) - per unit or box'] ||
                        map['Ship Cost (EGP)'];
    const colCustoms  = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
                        map['Customs (EGP)'] ||
                        map['Customs / Clearance (EGP)'];
    const colOther    = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
                        map['Other (EGP)'] ||
                        map['Other Fees (EGP)'];

    const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
                        map['Ship Date'] ||
                        map['Ship Date (UAE)'];
    const colEta      = map[APP.COLS.SHIP_UAE_EG.ETA] ||
                        map['ETA'];
    const colArr      = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
                        map['Actual Arrival'] ||
                        map['Actual Arrival (EG)'];

    const editedCol = e.range.getColumn();

    // 1) If SKU changed → fill from Inventory_UAE
    if (colSku && editedCol === colSku) {
      _fillShipmentUaeEgFromInventory_(sh, rowIndex, map);
    }

    // 2) If any column that affects costs or status changes → recalc the row
    const recalcCols = [
      colQty,
      colShipCost,
      colCustoms,
      colOther,
      colShipDate,
      colEta,
      colArr
    ].filter(function (c) { return !!c; });

    if (recalcCols.indexOf(editedCol) !== -1) {
      _updateShipmentUaeEgRowTotalsAndStatus_(sh, rowIndex, map);
    }

  } catch (e2) {
    logError_('shipmentsUaeEgOnEdit_', e2, {
      a1: e && e.range ? e.range.getA1Notation() : ''
    });
  }
}

/* ===================================================================
 * QC_UAE – Generate + Recalc + Sync → Inventory
 * =================================================================== */

/**
 * UI entry point:
 * Generate QC_UAE rows from Purchases.
 * - يطلب منك Order ID (اختياري).
 * - لو سيبته فاضي → يولّد لكل الأوردرات اللي لسه ملهاش QC.
 */
function qc_generateFromPurchasesPrompt() {
  try {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.prompt(
      'Generate QC_UAE rows from Purchases',
      'Enter a single Order ID to generate QC rows for.\n' +
      'Leave empty to generate for ALL orders that do not yet have QC rows.',
      ui.ButtonSet.OK_CANCEL
    );

    if (resp.getSelectedButton() !== ui.Button.OK) {
      return; // user cancelled
    }

    const orderIdText  = (resp.getResponseText() || '').trim();
    const orderIdFilter = orderIdText || null;

    const result = qc_generateFromPurchases_(orderIdFilter) || { added: 0, skipped: 0 };

    ui.alert(
      'QC_UAE generation done.\n\n' +
      'New QC rows added: ' + result.added + '\n' +
      'Existing rows skipped: ' + result.skipped
    );

  } catch (e) {
    logError_('qc_generateFromPurchasesPrompt', e);
    throw e;
  }
}

/**
 * Generate QC_UAE rows from Purchases (one row per SKU)
 * - يدعم فلترة حسب Order ID (اختياري)
 * - يملأ: QC ID, Order ID, Shipment CN→UAE ID, SKU, Batch Code,
 *         Product Name, Variant / Color, Qty Ordered
 */

/**
 * Generate / Upsert QC_UAE rows from Purchases (AGGREGATION SAFE)
 *
 * - Aggregates Qty Ordered for duplicate Purchases lines that share the same Batch Code
 *   (or fallback key if Batch Code is blank).
 * - If QC row already exists for that Batch Code => updates Qty Ordered (and fills blanks)
 * - If not exists => inserts a new QC row (at the TOP under headers)
 *
 * NOTE:
 * - Script edits do NOT trigger onEdit, so we call qc_recalcRows_ after write to fill:
 *   Qty Missing = Qty Ordered - Qty Received
 *   Qty OK      = Qty Received - Qty Defective
 */
function qc_generateFromPurchases_(orderIdFilter) {
  const purchasesSh = getSheet_(APP.SHEETS.PURCHASES);
  const qcSh        = getSheet_(APP.SHEETS.QC_UAE);
  const shipCnUaeSh = getSheet_(APP.SHEETS.SHIP_CN_UAE);

  const pMap        = getHeaderMap_(purchasesSh);
  const qcMap       = getHeaderMap_(qcSh);
  const shipHdrMap  = getHeaderMap_(shipCnUaeSh);

  const lastPurRow = purchasesSh.getLastRow();
  if (lastPurRow < 2) return { added: 0, updated: 0, skipped: 0 };

  let data = purchasesSh
    .getRange(2, 1, lastPurRow - 1, purchasesSh.getLastColumn())
    .getValues();

  // Optional filter by Order ID
  if (orderIdFilter) {
    const idxOrderPurch = (pMap[APP.COLS.PURCHASES.ORDER_ID] || pMap['Order ID'] || 0) - 1;
    if (idxOrderPurch >= 0) {
      data = data.filter(r => String(r[idxOrderPurch]) === String(orderIdFilter));
    }
  }

  if (!data.length) return { added: 0, updated: 0, skipped: 0 };

  // QC headers (canonical labels)
  const QC_ID_H        = 'QC ID';
  const ORDER_H        = 'Order ID';
  const SHIP_CN_UAE_H  = 'Shipment CN→UAE ID';
  const SKU_H          = 'SKU';
  const BATCH_H        = 'Batch Code';
  const PROD_H         = 'Product Name';
  const VAR_H          = 'Variant / Color';
  const QTY_ORD_H      = 'Qty Ordered';

  // Purchases column indexes
  const idxOrder   = (pMap[APP.COLS.PURCHASES.ORDER_ID] || pMap['Order ID'] || 0) - 1;
  const idxSku     = (pMap[APP.COLS.PURCHASES.SKU]      || pMap['SKU']      || 0) - 1;
  const idxVariant = (pMap[APP.COLS.PURCHASES.VARIANT]  || pMap['Variant / Color'] || 0) - 1;
  const idxQty     = (pMap[APP.COLS.PURCHASES.QTY]      || pMap['Qty']      || 0) - 1;
  const idxBatch   = (pMap[APP.COLS.PURCHASES.BATCH_CODE] || pMap['Batch Code'] || 0) - 1;

  const idxProd = ((pMap[APP.COLS.PURCHASES.PRODUCT_NAME] || pMap[APP.COLS.PURCHASES.PRODUCT] || pMap['Product Name'] || 0) - 1);

  if (idxOrder < 0 || idxSku < 0 || idxQty < 0) {
    throw new Error('qc_generateFromPurchases_: Missing Purchases columns (Order ID / SKU / Qty).');
  }

  // Build Shipments map: OrderID||SKU||Variant -> Shipment ID (+ Product fallback)
  const shipIdCol   = shipHdrMap['Shipment ID'] || shipHdrMap[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID];
  const shipOrdCol  = shipHdrMap['Order ID (Batch)'] || shipHdrMap['Order ID'];
  const shipSkuCol  = shipHdrMap['SKU'];
  const shipVarCol  = shipHdrMap['Variant / Color'] || shipHdrMap[APP.COLS.PURCHASES.VARIANT];
  const shipProdCol = shipHdrMap['Product Name'];

  /** @type {Record<string, {shipId:string, product:string}>} */
  const shipByKey = {};

  if (shipIdCol && shipOrdCol && shipSkuCol && shipCnUaeSh.getLastRow() >= 2) {
    const shipData = shipCnUaeSh
      .getRange(2, 1, shipCnUaeSh.getLastRow() - 1, shipCnUaeSh.getLastColumn())
      .getValues();

    shipData.forEach(r => {
      const shipId = r[shipIdCol - 1];
      const ordId  = r[shipOrdCol - 1];
      const skuS   = (r[shipSkuCol - 1] || '').toString().trim();
      const varS   = shipVarCol ? String(r[shipVarCol - 1] || '').trim() : '';
      if (!shipId || !ordId || !skuS) return;

      const key = String(ordId) + '||' + skuS + '||' + varS;
      if (!shipByKey[key]) {
        shipByKey[key] = {
          shipId : String(shipId),
          product: shipProdCol ? (r[shipProdCol - 1] || '') : ''
        };
      }
    });
  }

  // Existing QC rows by Batch Code
  const lastQcRow = qcSh.getLastRow();
  const qcBatchColIndex = qcMap[BATCH_H] || qcMap[APP.COLS.QC_UAE.BATCH_CODE];
  const qcQtyOrdCol     = qcMap[QTY_ORD_H] || qcMap['Qty Ordered'];
  const qcIdColIndex    = qcMap[QC_ID_H] || qcMap[APP.COLS.QC_UAE.QC_ID];

  if (!qcBatchColIndex || !qcQtyOrdCol) {
    throw new Error('qc_generateFromPurchases_: Missing QC_UAE columns (Batch Code / Qty Ordered).');
  }

  /** @type {Record<string, number>} */
  const batchToRow = {}; // BatchCode -> sheet row number

  if (lastQcRow >= 2) {
    const vals = qcSh.getRange(2, qcBatchColIndex, lastQcRow - 1, 1).getValues();
    vals.forEach((r, i) => {
      const bc = String(r[0] || '').trim();
      if (bc && !batchToRow[bc]) batchToRow[bc] = i + 2;
    });
  }

  // Next QC sequence
  let nextQcSeq = 1;
  if (qcIdColIndex && lastQcRow >= 2) {
    const existingIds = qcSh
      .getRange(2, qcIdColIndex, lastQcRow - 1, 1)
      .getValues()
      .map(r => r[0])
      .filter(v => v);

    existingIds.forEach(id => {
      const m = String(id).match(/(\d+)$/);
      if (m) {
        const num = Number(m[1]);
        if (num >= nextQcSeq) nextQcSeq = num + 1;
      }
    });
  }

  // Aggregate purchases by BatchCode (or fallback key)
  /** @type {Record<string, any>} */
  const agg = {};
  let skipped = 0;

  data.forEach(row => {
    const orderId = row[idxOrder];
    const sku     = (row[idxSku] || '').toString().trim();
    const qty     = Number(row[idxQty] || 0);
    if (!orderId || !sku || !qty) { skipped++; return; }

    const variant = (idxVariant >= 0) ? String(row[idxVariant] || '').trim() : '';
    let batchCode = (idxBatch >= 0) ? String(row[idxBatch] || '').trim() : '';
    if (!batchCode) batchCode = variant ? (String(orderId) + '||' + sku + '||' + variant) : (String(orderId) + '||' + sku);

    if (!agg[batchCode]) {
      agg[batchCode] = {
        batchCode: batchCode,
        orderId: String(orderId),
        sku: sku,
        variant: variant,
        qty: 0,
        product: (idxProd >= 0) ? (row[idxProd] || '') : ''
      };
    }
    agg[batchCode].qty += qty;
    if (!agg[batchCode].product && idxProd >= 0) agg[batchCode].product = row[idxProd] || '';
  });

  const batchCodes = Object.keys(agg);
  if (!batchCodes.length) return { added: 0, updated: 0, skipped: skipped };

  const headers = qcSh.getRange(1, 1, 1, qcSh.getLastColumn()).getValues()[0];

  // Build new rows (to insert at top) + updates for existing rows
  const newRows = [];
  const updateRows = []; // {row:number, qty:number, shipId:string, product:string, variant:string, orderId:string, sku:string, batchCode:string}
  let added = 0;
  let updated = 0;

  batchCodes.forEach(bc => {
    const it = agg[bc];
    const key = it.orderId + '||' + it.sku + '||' + (it.variant || '');
    const shipInfo = shipByKey[key] || {};

    const productFinal = it.product || shipInfo.product || '';
    const shipIdFinal  = shipInfo.shipId || '';

    const existingRow = batchToRow[bc];
    if (existingRow) {
      updateRows.push({
        row: existingRow,
        qty: Number(it.qty || 0),
        shipId: shipIdFinal,
        product: productFinal,
        variant: it.variant || '',
        orderId: it.orderId,
        sku: it.sku,
        batchCode: bc
      });
      return;
    }

    /** @type {Record<string, any>} */
    const obj = {};
    if (qcIdColIndex) obj[QC_ID_H] = 'QC-' + Utilities.formatString('%06d', nextQcSeq++);

    obj[ORDER_H]       = it.orderId;
    obj[SHIP_CN_UAE_H] = shipIdFinal;
    obj[SKU_H]         = it.sku;
    obj[BATCH_H]       = bc;
    obj[PROD_H]        = productFinal;
    obj[VAR_H]         = it.variant || '';
    obj[QTY_ORD_H]     = Number(it.qty || 0);

    const outRow = headers.map(h => (obj[h] !== undefined ? obj[h] : ''));
    newRows.push(outRow);
    added++;
  });

  // Apply updates (qty + fill blanks best-effort)
  if (updateRows.length) {
    updateRows.forEach(u => {
      if (qcQtyOrdCol) qcSh.getRange(u.row, qcQtyOrdCol).setValue(u.qty);

      // Fill blanks (do not overwrite user's edits)
      try {
        if (qcMap[SHIP_CN_UAE_H]) {
          const c = qcMap[SHIP_CN_UAE_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.shipId) qcSh.getRange(u.row, c).setValue(u.shipId);
        }
        if (qcMap[PROD_H]) {
          const c = qcMap[PROD_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.product) qcSh.getRange(u.row, c).setValue(u.product);
        }
        if (qcMap[VAR_H]) {
          const c = qcMap[VAR_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.variant) qcSh.getRange(u.row, c).setValue(u.variant);
        }
        if (qcMap[ORDER_H]) {
          const c = qcMap[ORDER_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.orderId) qcSh.getRange(u.row, c).setValue(u.orderId);
        }
        if (qcMap[SKU_H]) {
          const c = qcMap[SKU_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.sku) qcSh.getRange(u.row, c).setValue(u.sku);
        }
        if (qcMap[BATCH_H]) {
          const c = qcMap[BATCH_H];
          const cur = qcSh.getRange(u.row, c).getValue();
          if (!cur && u.batchCode) qcSh.getRange(u.row, c).setValue(u.batchCode);
        }
      } catch (e) {}
    });
    updated = updateRows.length;
  }

  // Insert new rows at TOP (row 2)
  if (newRows.length) {
    qcSh.insertRowsBefore(2, newRows.length);
    qcSh.getRange(2, 1, newRows.length, headers.length).setValues(newRows);
  }

  // Recalc computed QC columns for affected rows (since script edits don't fire onEdit)
  try {
    const startRow = 2;
    const nRows = Math.max(0, (newRows.length + updateRows.length));
    if (nRows) qc_recalcRows_(qcSh, qcMap, startRow, Math.min(qcSh.getLastRow() - 1, nRows), { silent: true, setDate: false });
  } catch (e) {
    logError_('qc_generateFromPurchases_recalc', e);
  }

  return { added: added, updated: updated, skipped: skipped };
}




/**
 * Recalculate QC_UAE quantities & result:
 * - لو Qty Missing و Qty Defective فاضيين ⇒ نحسب Missing = Qty Ordered - Qty OK
 * - نحدد QC Result = PASS / PARTIAL / FAIL لو فاضية
 * - نملأ QC Date بتاريخ اليوم لو فاضية
 */

/**
 * QC recalculation engine (row-based; safe for triggers)
 *
 * Business rules:
 * - Qty Missing   = Qty Ordered - Qty Received (>=0)
 * - Qty OK        = Qty Received - Qty Defective (>=0)
 * - QC Result (optional auto): PASS / PARTIAL / FAIL (only if blank)
 *
 * NOTE:
 * - This function is safe to call from installable triggers (no UI).
 * - For manual run from menu, call qc_recalcQuantitiesAndResult({silent:false}).
 */
function qc_recalcQuantitiesAndResult(opts) {
  const options = opts || {};
  const qcSh  = getSheet_(APP.SHEETS.QC_UAE);
  const qcMap = getHeaderMap_(qcSh);

  const lastRow = qcSh.getLastRow();
  if (lastRow < 2) return { updated: 0 };

  const rowStart = Math.max(2, Number(options.rowStart || 2));
  const numRows = Number(options.numRows || (lastRow - rowStart + 1));

  const res = qc_recalcRows_(qcSh, qcMap, rowStart, numRows, {
    silent: true,
    setDate: !!options.setDate,
    updateResult: (options.updateResult == null) ? true : !!options.updateResult
  });

  if (!options.silent) {
    safeAlert_('QC_UAE recalculation done.\nUpdated rows: ' + res.updated);
  }
  return res;
}

/**
 * Installable onEdit handler for QC_UAE (called from AppCore._dispatchOnEdit_)
 * - Recalculates only edited rows, only when edit touches Qty Ordered/Received/Defective.
 */
function qcOnEdit_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== (APP.SHEETS.QC_UAE || 'QC_UAE')) return;

  const qcMap = getHeaderMap_(sh);

  const colOrdered = qcMap['Qty Ordered'];
  const colRecv    = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
  const colDef     = qcMap[APP.COLS.QC_UAE.QTY_DEFECT]   || qcMap['Qty Defective'];

  if (!colOrdered || !colRecv || !colDef) return;

  // Only react when edit intersects the driving columns
  const colStart = e.range.getColumn();
  const colEnd   = colStart + e.range.getNumColumns() - 1;

  const drives = [colOrdered, colRecv, colDef];
  const intersects = drives.some(c => c >= colStart && c <= colEnd);
  if (!intersects) return;

  const rowStart = Math.max(2, e.range.getRow());
  const rowEnd   = Math.max(rowStart, e.range.getRow() + e.range.getNumRows() - 1);
  const numRows  = rowEnd - rowStart + 1;

  qc_recalcRows_(sh, qcMap, rowStart, numRows, { silent: true, setDate: false, updateResult: true });
}

/**
 * Internal worker: recalculates computed QC columns for a row window.
 * Uses batched reads/writes.
 */
function qc_recalcRows_(qcSh, qcMap, rowStart, numRows, opts) {
  const options = opts || {};
  const lastRow = qcSh.getLastRow();
  if (lastRow < 2) return { updated: 0 };

  const start = Math.max(2, rowStart || 2);
  const end = Math.min(lastRow, start + Math.max(0, numRows || (lastRow - start + 1)) - 1);
  if (end < start) return { updated: 0 };

  const lastCol = qcSh.getLastColumn();
  const data = qcSh.getRange(start, 1, end - start + 1, lastCol).getValues();

  const idxAnchor = qcMap['QC ID'] || qcMap[APP.COLS.QC_UAE.QC_ID] || qcMap[APP.COLS.QC_UAE.ORDER_ID] || qcMap['Order ID'];
  const idxOrd    = qcMap['Qty Ordered'];
  const idxRecv   = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
  const idxDef    = qcMap[APP.COLS.QC_UAE.QTY_DEFECT]   || qcMap['Qty Defective'];
  const idxMiss   = qcMap[APP.COLS.QC_UAE.QTY_MISSING]  || qcMap['Qty Missing'];
  const idxOk     = qcMap[APP.COLS.QC_UAE.QTY_OK]       || qcMap['Qty OK'];
  const idxRes    = qcMap['QC Result'];
  const idxDate   = qcMap['QC Date'];

  if (!idxOrd || !idxRecv || !idxDef || !idxMiss || !idxOk) {
    throw new Error('qc_recalcRows_: Missing QC_UAE columns (Qty Ordered/Received/Defective/OK/Missing).');
  }

  const today = new Date();
  const missOut = [];
  const okOut = [];
  const resOut = [];
  const dateOut = [];

  let updated = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const anchor = idxAnchor ? row[idxAnchor - 1] : row[idxOrd - 1];
    if (!anchor) {
      missOut.push([row[idxMiss - 1]]);
      okOut.push([row[idxOk - 1]]);
      if (idxRes) resOut.push([row[idxRes - 1]]);
      if (idxDate) dateOut.push([row[idxDate - 1]]);
      continue;
    }

    const qtyOrd  = Number(row[idxOrd - 1] || 0);
    const qtyRecv = Number(row[idxRecv - 1] || 0);
    const qtyDef  = Number(row[idxDef - 1] || 0);

    const missing = Math.max(qtyOrd - qtyRecv, 0);
    const ok      = Math.max(qtyRecv - qtyDef, 0);

    missOut.push([missing]);
    okOut.push([ok]);

    // QC Result: fill only if blank (do not overwrite manual)
    if (idxRes) {
      let cur = row[idxRes - 1];
      if (!cur && (options.updateResult !== false)) {
        if (qtyOrd <= 0 && qtyRecv <= 0 && qtyDef <= 0) {
          cur = '';
        } else if (missing === 0 && qtyDef === 0 && qtyRecv >= qtyOrd) {
          cur = 'PASS';
        } else if (ok > 0) {
          cur = 'PARTIAL';
        } else {
          cur = 'FAIL';
        }
      }
      resOut.push([cur || '']);
    }

    // QC Date: optional (default false here; set via manual action if needed)
    if (idxDate) {
      let d = row[idxDate - 1];
      if (!d && options.setDate) d = today;
      dateOut.push([d || '']);
    }

    updated++;
  }

  // Batched writes (1 col per call)
  qcSh.getRange(start, idxMiss, missOut.length, 1).setValues(missOut);
  qcSh.getRange(start, idxOk, okOut.length, 1).setValues(okOut);
  if (idxRes) qcSh.getRange(start, idxRes, resOut.length, 1).setValues(resOut);
  if (idxDate) qcSh.getRange(start, idxDate, dateOut.length, 1).setValues(dateOut);

  return { updated: updated };
}



/**
 * Sync QC_UAE → Inventory_Transactions + Inventory_UAE
 * - لكل صف QC فيه SKU وكميات:
 *    - Qty In = Qty OK (الجديدة) أو (Qty Received - Qty Defective) لو الهيدر القديم لسه موجود.
 * - Unit Cost (EGP) بياخده من Purchases:
 *    - أولاً عن طريق Batch Code (لو موجود)
 *    - أو fallback على (Order ID + SKU)
 * - dedupe عن طريق قراءة الـ ledger (SourceType = QC_UAE + row رقم كـ note)
 *
 * - Warehouse (UAE):
 *    - يفضل تملأه في QC_UAE بـ:
 *        UAE-ATTIA أو UAE-KOR أو UAE-DXB
 *    - لو فاضي → default = 'UAE-DXB'
 */
function syncQCtoInventory_UAE() {
  try {
    const qcSh     = getSheet_(APP.SHEETS.QC_UAE);
    const purchSh  = getSheet_(APP.SHEETS.PURCHASES);
    const ledgerSh = getSheet_(APP.SHEETS.INVENTORY_TXNS);

    const qcMap     = getHeaderMap_(qcSh);
    const purchMap  = getHeaderMap_(purchSh);
    const ledgerMap = getHeaderMap_(ledgerSh);

    const lastQcRow = qcSh.getLastRow();
    if (lastQcRow < 2) return; // مفيش بيانات

    const qcData = qcSh.getRange(2, 1, lastQcRow - 1, qcSh.getLastColumn()).getValues();

    const qcIdCol    = qcMap[APP.COLS.QC_UAE.QC_ID] || qcMap['QC ID'];
    const qcDateCol  = qcMap[APP.COLS.QC_UAE.QC_DATE] || qcMap['QC Date'];

    // ===== Build cost map from Purchases =====
    const purchLast = purchSh.getLastRow();
    const costMap = {};

    if (purchLast >= 2) {
      const purchData = purchSh.getRange(2, 1, purchLast - 1, purchSh.getLastColumn()).getValues();

      const idxOrder = purchMap[APP.COLS.PURCHASES.ORDER_ID] ? purchMap[APP.COLS.PURCHASES.ORDER_ID] - 1 : null;
      const idxSku   = purchMap[APP.COLS.PURCHASES.SKU]      ? purchMap[APP.COLS.PURCHASES.SKU]      - 1 : null;
      const idxBatch = purchMap[APP.COLS.PURCHASES.BATCH_CODE] ? purchMap[APP.COLS.PURCHASES.BATCH_CODE] - 1 : null;

      // Prefer Unit Landed Cost, fallback to Unit Price (EGP)
      const idxUnitLanded = purchMap[APP.COLS.PURCHASES.UNIT_LANDED] ? purchMap[APP.COLS.PURCHASES.UNIT_LANDED] - 1 : null;
      const idxUnitNet    = purchMap[APP.COLS.PURCHASES.NET_UNIT_PRICE] ? purchMap[APP.COLS.PURCHASES.NET_UNIT_PRICE] - 1 : null;

      purchData.forEach(function (r) {
        if (idxOrder == null || idxSku == null) return;
        const orderId = r[idxOrder];
        const sku     = (r[idxSku] || '').toString().trim();
        if (!orderId || !sku) return;

        const baseKey = String(orderId) + '||' + String(sku);
        const unitCost = idxUnitLanded != null
          ? Number(r[idxUnitLanded] || 0)
          : (idxUnitNet != null ? Number(r[idxUnitNet] || 0) : 0);

        if (unitCost) {
          costMap['ORDSKU||' + baseKey] = unitCost;
        }

        if (idxBatch != null) {
          const batch = (r[idxBatch] || '').toString().trim();
          if (batch && unitCost) {
            costMap['BATCH||' + batch] = unitCost;
          }
        }
      });
    }

    // ===== Determine already-synced QC rows (new: by QC ID, legacy: by notes row number) =====
    const syncedQcIds = new Set();
    const syncedLegacyRows = new Set();

    const ledgerLast = ledgerSh.getLastRow();
    if (ledgerLast >= 2) {
      const ledgerData = ledgerSh.getRange(2, 1, ledgerLast - 1, ledgerSh.getLastColumn()).getValues();

      const idxSrcType = ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] ? ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] - 1 : null;
      const idxSrcId   = ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID]   ? ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID]   - 1 : null;
      const idxNotes   = ledgerMap[APP.COLS.INV_TXNS.NOTES]       ? ledgerMap[APP.COLS.INV_TXNS.NOTES]       - 1 : null;

      ledgerData.forEach(function (row) {
        const srcType = idxSrcType != null ? row[idxSrcType] : '';
        if (String(srcType) !== 'QC_UAE') return;

        const srcId = idxSrcId != null ? row[idxSrcId] : '';
        if (srcId) {
          syncedQcIds.add(String(srcId).trim());
          return;
        }

        // Legacy fallback: parse "row 12" from notes
        const notes = idxNotes != null ? (row[idxNotes] || '') : '';
        const m = String(notes).match(/row\s+(\d+)/i);
        if (m) {
          const rNum = Number(m[1]);
          if (rNum) syncedLegacyRows.add(rNum);
        }
      });
    }

    const hasBatchQC = !!qcMap[APP.COLS.QC_UAE.BATCH_CODE];

    let newTxns = 0;
    let skipped = 0;
    const txns = [];

    // ----- Loop QC rows and add IN movements (UAE) -----
    qcData.forEach(function (row, idx) {
      const sheetRow = idx + 2;

      const orderId = row[qcMap[APP.COLS.QC_UAE.ORDER_ID] - 1];
      const sku     = (row[qcMap[APP.COLS.QC_UAE.SKU] - 1] || '').toString().trim();
      if (!orderId || !sku) return;

      let qcId = qcIdCol ? String(row[qcIdCol - 1] || '').trim() : '';
      if (!qcId) qcId = 'QC_ROW_' + sheetRow;

      if (syncedQcIds.has(qcId) || syncedLegacyRows.has(sheetRow)) {
        skipped++;
        return;
      }

      const batchCodeRaw = hasBatchQC ? row[qcMap[APP.COLS.QC_UAE.BATCH_CODE] - 1] : '';
      let batchCode = (batchCodeRaw || '').toString().trim();
      if (!batchCode) batchCode = String(orderId) + '||' + String(sku);

      const product   = row[qcMap['Product Name'] - 1] || '';
      const variant   = row[qcMap['Variant / Color'] - 1] || '';
      const warehouse = (row[qcMap['Warehouse (UAE)'] - 1] || 'UAE-DXB').toString().trim();

      // Qty In logic:
      //  - If Qty OK exists → use it
      //  - Else fallback Qty Received - Qty Defective
      let qtyIn = 0;
      const qtyOkCol  = qcMap['Qty OK'];
      const qtyRecCol = qcMap['Qty Received'];
      const qtyDefCol = qcMap['Qty Defective'];

      if (qtyOkCol) {
        qtyIn = Number(row[qtyOkCol - 1] || 0);
      } else if (qtyRecCol) {
        const rec = Number(row[qtyRecCol - 1] || 0);
        const def = qtyDefCol ? Number(row[qtyDefCol - 1] || 0) : 0;
        qtyIn = Math.max(0, rec - def);
      }

      if (!qtyIn) return;

      const qcDate = qcDateCol ? row[qcDateCol - 1] : null;

      const baseKey = String(orderId) + '||' + String(sku);
      let unitCostEgp = 0;

      if (batchCode) {
        unitCostEgp = costMap['BATCH||' + batchCode] || 0;
      }
      if (!unitCostEgp) {
        unitCostEgp = costMap['ORDSKU||' + baseKey] || 0;
      }

      txns.push({
        type: 'IN',
        sourceType: 'QC_UAE',
        sourceId: qcId,
        batchCode: batchCode,
        sku: sku,
        productName: product,
        variant: variant,
        warehouse: warehouse,
        qty: qtyIn,
        unitCostEgp: unitCostEgp,
        currency: 'EGP',
        txnDate: qcDate || new Date(),
        notes: 'Imported from QC_UAE (QC ID: ' + qcId + ', row ' + sheetRow + ', Order ' + orderId + ')'
      });

      newTxns++;
    });

    if (txns.length) {
      if (typeof logInventoryTxnBatch_ === 'function') {
        logInventoryTxnBatch_(txns);
      } else {
        txns.forEach(function (t) { logInventoryTxn_(t); });
      }
    }

    inv_rebuildAllSnapshots();

    SpreadsheetApp.getUi().alert(
      'QC_UAE Sync Done.\n\nNew txns: ' + newTxns +
      '\nSkipped (already synced): ' + skipped
    );

  } catch (e) {
    logError_('syncQCtoInventory_UAE', e);
    throw e;
  }
}

/**
 * Sync Shipments_UAE_EG → Inventory_Transactions + snapshots
 *
 * المنطق:
 *  - Qty = إجمالي اللي شحنته لمصر لحد دلوقتي.
 *  - Qty Synced = الكمية اللي اتحولت بالفعل لحركات OUT/IN في الـ ledger.
 *  - delta = Qty - Qty Synced.
 *
 *  - Warehouse (UAE) يحدد المخزن اللي هنخصم منه:
 *      * لو موجود في نفس صف الشحنة → هو المصدر الأساسي.
 *      * لو فاضي → نحاول نستنتجه من Courier.
 *      * لو لسه مش واضح → نرجع لأول Warehouse للـ SKU من Inventory_UAE.
 *
 *  - النتيجة: OUT من Warehouse الحقيقي (UAE-ATTIA / UAE-KOR / ...),
 *             IN في EG-CAI بالـ landed cost.
 */
function syncShipmentsUaeEgToInventory() {
  try {
    const runner_ = function () {
      const shipSh   = getSheet_(APP.SHEETS.SHIP_UAE_EG);
      const invUaeSh = getSheet_(APP.SHEETS.INVENTORY_UAE);
      const ledgerSh = getSheet_(APP.SHEETS.INVENTORY_TXNS);

      const shipMap   = getHeaderMap_(shipSh);
      const invUaeMap = getHeaderMap_(invUaeSh);

      const lastShipRow = shipSh.getLastRow();
      if (lastShipRow < 2) {
        SpreadsheetApp.getUi().alert('No rows in Shipments_UAE_EG to sync.');
        return;
      }

      const shipData = shipSh.getRange(2, 1, lastShipRow - 1, shipSh.getLastColumn()).getValues();

      // Required columns
      const colShipmentId = shipMap[APP.COLS.SHIP_UAE_EG.SHIPMENT_ID] || shipMap['Shipment ID'];
      const colSku        = shipMap[APP.COLS.SHIP_UAE_EG.SKU]         || shipMap['SKU'];
      const colQty        = shipMap[APP.COLS.SHIP_UAE_EG.QTY]         || shipMap['Qty'];
      const colShipDate   = shipMap[APP.COLS.SHIP_UAE_EG.SHIP_DATE]   || shipMap['Ship Date'] || shipMap['Ship Date (UAE)'];
      const colArrival    = shipMap[APP.COLS.SHIP_UAE_EG.ARRIVAL]     || shipMap['Actual Arrival'] || shipMap['Actual Arrival (EG)'];
      const colQtySynced  = shipMap[APP.COLS.SHIP_UAE_EG.QTY_SYNCED]  || shipMap['Qty Synced'];

      if (!colShipmentId || !colSku || !colQty || !colShipDate || !colQtySynced) {
        SpreadsheetApp.getUi().alert('Missing required headers in Shipments_UAE_EG (Shipment ID / SKU / Qty / Ship Date / Qty Synced).');
        return;
      }

      const idxShipShipmentId = colShipmentId - 1;
      const idxShipSku        = colSku        - 1;
      const idxShipQty        = colQty        - 1;
      const idxShipShipDate   = colShipDate   - 1;
      const idxShipArrival    = colArrival    ? colArrival    - 1 : null;
      const idxShipQtySynced  = colQtySynced  - 1;

      // Cost columns
      const colShipCost =
        shipMap[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
        shipMap['Ship Cost (EGP) – per unit or box'] ||
        shipMap['Ship Cost (EGP) - per unit or box'] ||
        shipMap['Ship Cost (EGP)'];
      const colCustoms =
        shipMap[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
        shipMap['Customs (EGP)'];
      const colOther =
        shipMap[APP.COLS.SHIP_UAE_EG.OTHER] ||
        shipMap['Other (EGP)'];
      const colTotal =
        shipMap[APP.COLS.SHIP_UAE_EG.TOTAL_COST] ||
        shipMap['Total Cost (EGP)'];

      const idxShipShipCost = colShipCost ? colShipCost - 1 : null;
      const idxShipCustoms  = colCustoms  ? colCustoms  - 1 : null;
      const idxShipOther    = colOther    ? colOther    - 1 : null;
      const idxShipTotal    = colTotal    ? colTotal    - 1 : null;

      // Optional columns: Product / Variant / Courier / Warehouse (UAE)
      const idxShipProdName = shipMap['Product Name']    ? shipMap['Product Name']    - 1 : null;
      const idxShipVariant  = shipMap['Variant / Color'] ? shipMap['Variant / Color'] - 1 : null;
      const colCourier      = shipMap['Courier'] || shipMap['Courier Name'];
      const idxShipCourier  = colCourier ? colCourier - 1 : null;
      const colShipWhUae    = shipMap['Warehouse (UAE)'];
      const idxShipWhUae    = colShipWhUae ? colShipWhUae - 1 : null;

      // ===== Inventory_UAE map: SKU+Warehouse → cost info =====
      const uaeLastRow = invUaeSh.getLastRow();
      const uaeData = (uaeLastRow >= 2)
        ? invUaeSh.getRange(2, 1, uaeLastRow - 1, invUaeSh.getLastColumn()).getValues()
        : [];

      const idxUaeSku       = invUaeMap['SKU']             ? invUaeMap['SKU']             - 1 : null;
      const idxUaeProd      = invUaeMap['Product Name']    ? invUaeMap['Product Name']    - 1 : null;
      const idxUaeVariant   = invUaeMap['Variant / Color'] ? invUaeMap['Variant / Color'] - 1 : null;
      const idxUaeAvgCost   = invUaeMap['Avg Cost (EGP)']  ? invUaeMap['Avg Cost (EGP)']  - 1 : null;
      const idxUaeLastSrcId = invUaeMap['Last Source ID']  ? invUaeMap['Last Source ID']  - 1 : null;
      const idxUaeWh        = invUaeMap['Warehouse (UAE)'] ? invUaeMap['Warehouse (UAE)'] - 1 : null;

      /** SKU+Warehouse → info */
      const uaeByWh = {};
      /** SKU → default info (first warehouse found for SKU) */
      const uaeBySku = {};

      uaeData.forEach(function (r) {
        if (idxUaeSku == null) return;
        const skuVal = r[idxUaeSku];
        if (!skuVal) return;

        const wh  = idxUaeWh != null ? (r[idxUaeWh] || '').toString().trim() : '';
        const key = skuVal + '||' + wh;

        const info = {
          product: idxUaeProd != null ? (r[idxUaeProd] || '') : '',
          variant: idxUaeVariant != null ? (r[idxUaeVariant] || '') : '',
          avgCost: idxUaeAvgCost != null ? Number(r[idxUaeAvgCost] || 0) : 0,
          lastSourceId: idxUaeLastSrcId != null ? (r[idxUaeLastSrcId] || '') : '',
          warehouse: wh
        };

        if (wh) uaeByWh[key] = info;
        if (!uaeBySku[skuVal]) uaeBySku[skuVal] = info;
      });

      let outCount = 0;
      let inCount  = 0;

      const txns = [];

      // Process shipments rows
      shipData.forEach(function (row, idx) {
        const shipmentId = row[idxShipShipmentId];
        const sku        = row[idxShipSku];
        const qtyCurrent = Number(row[idxShipQty] || 0);

        if (!shipmentId || !sku || !qtyCurrent) return;

        const qtySynced = Number(row[idxShipQtySynced] || 0);
        const delta     = qtyCurrent - qtySynced;

        if (delta === 0) return;
        if (delta < 0) {
          logError_(
            'syncShipmentsUaeEgToInventory',
            new Error('Negative delta for shipment row (Qty < Qty Synced).'),
            { shipmentId: shipmentId, sku: sku, qtyCurrent: qtyCurrent, qtySynced: qtySynced }
          );
          return;
        }

        const shipDate  = row[idxShipShipDate] || new Date();
        const arrDate   = (idxShipArrival != null) ? (row[idxShipArrival] || null) : null;
        const inTxnDate = arrDate || shipDate;

        // Total extras (ship/customs/other) per unit (for landed cost)
        const shipCost = idxShipShipCost != null ? Number(row[idxShipShipCost] || 0) : 0;
        const customs  = idxShipCustoms  != null ? Number(row[idxShipCustoms]  || 0) : 0;
        const other    = idxShipOther    != null ? Number(row[idxShipOther]    || 0) : 0;
        const totalExtras = (idxShipTotal != null)
          ? Number(row[idxShipTotal] || (shipCost + customs + other))
          : (shipCost + customs + other);

        // === Determine UAE warehouse for OUT ===
        let fromWarehouse = '';
        const whCell = (idxShipWhUae != null) ? (row[idxShipWhUae] || '').toString().trim() : '';

        // 1) If sheet has Warehouse (UAE), use it
        if (whCell) fromWarehouse = whCell;

        // 2) Else infer from courier label (Kor / Attia) if present
        let courierLabel = (idxShipCourier != null) ? (row[idxShipCourier] || '').toString().trim() : '';
        if (!fromWarehouse && courierLabel) {
          const c = courierLabel.toUpperCase();
          if (c.indexOf('KOR') >= 0) fromWarehouse = 'UAE-KOR';
          if (c.indexOf('ATTIA') >= 0) fromWarehouse = 'UAE-ATTIA';
        }

        // 3) Else infer from inventory (first warehouse for SKU)
        if (!fromWarehouse) {
          const invGuess = uaeBySku[sku] || {};
          if (invGuess.warehouse) fromWarehouse = invGuess.warehouse;
        }

        if (!fromWarehouse) fromWarehouse = 'UAE-DXB';

        // ===== Inventory info for this SKU+Warehouse =====
        const keyWh  = sku + '||' + fromWarehouse;
        let info     = uaeByWh[keyWh] || uaeBySku[sku] || {};

        // If Warehouse (UAE) is empty in sheet and inventory has a warehouse → fill it
        if (idxShipWhUae != null && !whCell && info.warehouse) {
          row[idxShipWhUae] = info.warehouse;
          fromWarehouse = info.warehouse;
        }

        const baseCost      = info.avgCost || 0; // UAE avg cost
        const extrasPerUnit = qtyCurrent > 0 ? totalExtras / qtyCurrent : 0;
        const landedCost    = baseCost + extrasPerUnit;

        // Fill Product / Variant if missing
        if (idxShipProdName != null && !row[idxShipProdName] && info.product) {
          row[idxShipProdName] = info.product;
        }
        if (idxShipVariant != null && !row[idxShipVariant] && info.variant) {
          row[idxShipVariant] = info.variant;
        }

        // Auto-fill Courier label if missing and warehouse indicates an office
        if (idxShipCourier != null && !courierLabel && fromWarehouse) {
          const whUpper = fromWarehouse.toUpperCase();
          if (whUpper === 'UAE-ATTIA') {
            row[idxShipCourier] = 'Attia';
          } else if (whUpper === 'UAE-KOR') {
            row[idxShipCourier] = 'Kor';
          }
          courierLabel = row[idxShipCourier];
        }

        // Batch Code: prefer inventory last source id
        const batchCode = info.lastSourceId
          ? String(info.lastSourceId) + '||' + String(sku)
          : String(shipmentId) + '||' + String(sku);

        // ===== OUT from UAE warehouse =====
        txns.push({
          type        : 'OUT',
          sourceType  : 'SHIP_UAE_EG',
          sourceId    : String(shipmentId),
          batchCode   : batchCode,
          sku         : String(sku),
          productName : info.product || (idxShipProdName != null ? row[idxShipProdName] : '') || '',
          variant     : info.variant || (idxShipVariant != null ? row[idxShipVariant] : '') || '',
          warehouse   : fromWarehouse,
          qty         : delta,
          unitCostEgp : baseCost,
          currency    : 'EGP',
          txnDate     : shipDate,
          notes       : 'UAE→EG OUT (' + fromWarehouse + '), delta=' + delta
        });
        outCount++;

        // ===== IN to EG-CAI at landed cost =====
        txns.push({
          type        : 'IN',
          sourceType  : 'SHIP_UAE_EG',
          sourceId    : String(shipmentId),
          batchCode   : batchCode,
          sku         : String(sku),
          productName : info.product || (idxShipProdName != null ? row[idxShipProdName] : '') || '',
          variant     : info.variant || (idxShipVariant != null ? row[idxShipVariant] : '') || '',
          warehouse   : 'EG-CAI',
          qty         : delta,
          unitCostEgp : landedCost,
          currency    : 'EGP',
          txnDate     : inTxnDate,
          notes       : 'UAE→EG IN (EG-CAI), delta=' + delta + ', landedCost=' + landedCost.toFixed(2)
        });
        inCount++;

        // Update Qty Synced
        row[idxShipQtySynced] = qtyCurrent;
      });

      // Write txns in one batch (faster + one lock)
      if (txns.length) {
        if (typeof logInventoryTxnBatch_ === 'function') {
          logInventoryTxnBatch_(txns);
        } else {
          txns.forEach(function (t) { logInventoryTxn_(t); });
        }
      }

      // Write sheet updates (Qty Synced + optional autofills)
      shipSh.getRange(2, 1, shipData.length, shipSh.getLastColumn()).setValues(shipData);

      // Rebuild snapshots
      inv_rebuildAllSnapshots();

      SpreadsheetApp.getUi().alert(
        'Sync Shipments_UAE_EG → Inventory done.\n' +
        'New OUT txns: ' + outCount + '\n' +
        'New IN txns: ' + inCount
      );
    };

    if (typeof withLock_ === 'function') {
      return withLock_('SYNC_SHIP_UAE_EG_TO_INV', runner_);
    }

    const lock = LockService.getDocumentLock();
    lock.waitLock(30000);
    try {
      return runner_();
    } finally {
      lock.releaseLock();
    }

  } catch (e) {
    logError_('syncShipmentsUaeEgToInventory', e);
    throw e;
  }
}

/**
 * Debug helper for Inventory lookup.
 */
function debugTestInventoryLookup_() {
  var info = _getInventoryUaeInfoForSku_('MONSTER-MQT65-PURPLE', 'UAE-DXB');
  Logger.log(JSON.stringify(info));
}
