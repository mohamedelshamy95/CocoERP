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
  PLANNED: 'Planned',
  IN_TRANSIT: 'In Transit',
  DELAYED: 'Delayed',
  ARRIVED_UAE: 'Arrived UAE',
  ARRIVED_EG: 'Arrived EG'
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
function updateShipmentsCnUaeStatusAndTotals(opts) {
  const interactive = !!(opts && opts.interactive);
  try {
    const sh = getSheet_(APP.SHEETS.SHIP_CN_UAE);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) {
      safeAlert_('No data found in Shipments_CN_UAE.');
      return;
    }

    const map = getHeaderMap_(sh);

    const colShipmentId = map[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || map['Shipment ID'];
    const colShipDate = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE] || map['Ship Date'];
    const colEta = map[APP.COLS.SHIP_CN_UAE.ETA] || map['ETA'];
    const colArrival = map[APP.COLS.SHIP_CN_UAE.ARRIVAL] || map['Actual Arrival'];
    const colStatus = map[APP.COLS.SHIP_CN_UAE.STATUS] || map['Status'];
    const colFreight = map[APP.COLS.SHIP_CN_UAE.FREIGHT_AED] || map['Freight (AED)'];
    const colOther = map[APP.COLS.SHIP_CN_UAE.OTHER_AED] || map['Other Fees (AED)'];
    const colTotal = map[APP.COLS.SHIP_CN_UAE.TOTAL_AED] || map['Total Cost (AED)'];

    if (!colShipmentId || !colStatus) {
      safeAlert_('Missing required headers in Shipments_CN_UAE (Shipment ID / Status).');
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
        const rawOther = colOther ? row[colOther - 1] : '';

        if (rawFreight === '' && (rawOther === '' || rawOther === undefined)) {
          totalVal = '';
        } else {
          const freight = Number(rawFreight || 0);
          const other = colOther ? Number(rawOther || 0) : 0;
          totalVal = freight + other;
        }
      }

      // ----- Status -----
      const shipDate = colShipDate ? row[colShipDate - 1] : null;
      const eta = colEta ? row[colEta - 1] : null;
      const arr = colArrival ? row[colArrival - 1] : null;

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

    if (interactive && typeof safeAlert_ === 'function') safeAlert_('Shipments_CN_UAE updated (status + totals)');
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
  const colShipDate = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE] || map['Ship Date'];
  const colEta = map[APP.COLS.SHIP_CN_UAE.ETA] || map['ETA'];
  const colArrival = map[APP.COLS.SHIP_CN_UAE.ARRIVAL] || map['Actual Arrival'];
  const colStatus = map[APP.COLS.SHIP_CN_UAE.STATUS] || map['Status'];
  const colFreight = map[APP.COLS.SHIP_CN_UAE.FREIGHT_AED] || map['Freight (AED)'];
  const colOther = map[APP.COLS.SHIP_CN_UAE.OTHER_AED] || map['Other Fees (AED)'];
  const colTotal = map[APP.COLS.SHIP_CN_UAE.TOTAL_AED] || map['Total Cost (AED)'];

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
    if (colTotal) sh.getRange(rowIndex, colTotal).clearContent();
    return;
  }

  // ----- Total Cost (AED) -----
  if (colTotal && colFreight) {
    const rawFreight = row[colFreight - 1];
    const rawOther = colOther ? row[colOther - 1] : '';

    // If no numbers in Freight / Other → leave Total Cost empty
    if (rawFreight === '' && (rawOther === '' || rawOther === undefined)) {
      row[colTotal - 1] = '';
    } else {
      const freight = Number(rawFreight || 0);
      const other = colOther ? Number(rawOther || 0) : 0;
      row[colTotal - 1] = freight + other;
    }
  }

  // ----- Status -----
  const ship = colShipDate ? row[colShipDate - 1] : null;
  const eta = colEta ? row[colEta - 1] : null;
  const arr = colArrival ? row[colArrival - 1] : null;

  let status;

  if (arr) {
    status = SHIPMENT_STATUS.ARRIVED_UAE;
  } else if (ship && eta) {
    const today = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const etaMid = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());

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
    const colShipDate = map[APP.COLS.SHIP_CN_UAE.SHIP_DATE] || map['Ship Date'];
    const colEta = map[APP.COLS.SHIP_CN_UAE.ETA] || map['ETA'];
    const colArrival = map[APP.COLS.SHIP_CN_UAE.ARRIVAL] || map['Actual Arrival'];
    const colFreight = map[APP.COLS.SHIP_CN_UAE.FREIGHT_AED] || map['Freight (AED)'];
    const colOther = map[APP.COLS.SHIP_CN_UAE.OTHER_AED] || map['Other Fees (AED)'];

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
function updateShipmentsUaeEgStatusAndTotals(opts) {
  const interactive = !!(opts && opts.interactive);
  try {
    const sh = getSheet_(APP.SHEETS.SHIP_UAE_EG);
    const map = getHeaderMap_(sh);

    const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
      map['Ship Date'] ||
      map['Ship Date (UAE)'];

    const colEta = map[APP.COLS.SHIP_UAE_EG.ETA] ||
      map['ETA'];

    const colArr = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
      map['Actual Arrival'] ||
      map['Actual Arrival (EG)'];

    const colStatus = map[APP.COLS.SHIP_UAE_EG.STATUS] ||
      map['Status'];

    const colQty = map[APP.COLS.SHIP_UAE_EG.QTY] ||
      map['Qty'];

    const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
      map['Ship Cost (EGP) - per unit or box'] ||
      map['Ship Cost (EGP)'];

    const colCustoms = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
      map['Customs (EGP)'] ||
      map['Customs / Clearance (EGP)'];

    const colOther = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
      map['Other (EGP)'] ||
      map['Other Fees (EGP)'];

    const colTotal = map[APP.COLS.SHIP_UAE_EG.TOTAL_COST] ||
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
      if (interactive && typeof safeAlert_ === 'function') safeAlert_('No data found in Shipments_UAE_EG.');
      else Logger.log('No data found in Shipments_UAE_EG.');
      return;
    }

    const numRows = lastRow - 1;
    const range = sh.getRange(2, 1, numRows, sh.getLastColumn());
    const values = range.getValues();

    const idx = {
      shipDate: colShipDate - 1,
      eta: colEta - 1,
      arr: colArr - 1,
      status: colStatus - 1,
      qty: colQty - 1,
      shipCost: colShipCost - 1,
      customs: colCustoms - 1,
      other: colOther - 1,
      total: colTotal - 1
    };

    const today = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());

    values.forEach(function (row) {
      // ----- Total Cost -----
      const qty = Number(row[idx.qty] || 0);
      const shipCostPerUnit = Number(row[idx.shipCost] || 0);
      const customs = Number(row[idx.customs] || 0);
      const other = Number(row[idx.other] || 0);

      const totalForShipment = (shipCostPerUnit * qty) + customs + other;
      row[idx.total] = totalForShipment;

      // ----- Status -----
      const ship = row[idx.shipDate];
      const eta = row[idx.eta];
      const arr = row[idx.arr];

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

    if (interactive && typeof safeAlert_ === 'function') safeAlert_('Shipments_UAE_EG updated (status + totals)');
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
  updateShipmentsCnUaeStatusAndTotals({ interactive: true });
  updateShipmentsUaeEgStatusAndTotals({ interactive: true });
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
    const shipSh = getSheet_(APP.SHEETS.SHIP_CN_UAE);

    const pMap = getHeaderMap_(purchSh);
    const sMap = getHeaderMap_(shipSh);

    const lastPurRow = purchSh.getLastRow();
    if (lastPurRow < 2) {
      safeAlert_('No data found in Purchases.');
      return;
    }

    // Purchases columns (robust)
    const colOrderId = pMap[APP.COLS.PURCHASES.ORDER_ID] || pMap['Order ID'];
    const colOrderDate = pMap[APP.COLS.PURCHASES.ORDER_DATE] || pMap['Order Date'];
    const colPlatform = pMap[APP.COLS.PURCHASES.PLATFORM] || pMap['Platform'];
    const colSeller = pMap[APP.COLS.PURCHASES.SELLER] || pMap['Seller Name'];
    const colSku = pMap[APP.COLS.PURCHASES.SKU] || pMap['SKU'];
    const colProduct = pMap[APP.COLS.PURCHASES.PRODUCT_NAME] || pMap[APP.COLS.PURCHASES.PRODUCT] || pMap['Product Name'];
    const colVariant = pMap[APP.COLS.PURCHASES.VARIANT] || pMap['Variant / Color'];
    const colQty = pMap[APP.COLS.PURCHASES.QTY] || pMap['Qty'];
    const colLineId = pMap[APP.COLS.PURCHASES.LINE_ID] || pMap['Line ID'];

    // Invoice indicators (to decide "ready to ship")
    const colInvoiceLink = pMap['Invoice Link'];
    const colInvoicePrev = pMap['Invoice Preview'];
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
    } catch (e) { }

    const purchData = purchSh
      .getRange(2, 1, lastPurRow - 1, purchSh.getLastColumn())
      .getValues();

    // Shipments columns
    const shipColId = sMap[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || sMap['Shipment ID'];
    const shipColOrderId = sMap[APP.COLS.SHIP_CN_UAE.ORDER_BATCH] || sMap['Order ID (Batch)'] || sMap['Order ID'];
    const shipColLineId = sMap[APP.COLS.SHIP_CN_UAE.PURCHASE_LINE_ID] || sMap['Purchases Line ID'];
    const shipColSku = sMap[APP.COLS.SHIP_CN_UAE.SKU] || sMap['SKU'];
    const shipColVariant = sMap[APP.COLS.SHIP_CN_UAE.VARIANT] || sMap['Variant / Color'] || sMap[APP.COLS.PURCHASES.VARIANT];
    const shipColQty = sMap[APP.COLS.SHIP_CN_UAE.QTY] || sMap[APP.COLS.PURCHASES.QTY] || sMap['Qty'];
    const shipColProd = sMap[APP.COLS.SHIP_CN_UAE.PRODUCT_NAME] || sMap['Product Name'];

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
        const lineId = String(row[shipColLineId - 1] || '').trim();
        const shipId = shipColId ? String(row[shipColId - 1] || '').trim() : '';

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
      const sku = colSku ? String(r[colSku - 1] || '').trim() : '';
      const qty = colQty ? Number(r[colQty - 1] || 0) : 0;
      const lineId = colLineId ? String(r[colLineId - 1] || '').trim() : '';

      if (!orderId || !sku || !qty) return;
      if (!lineId) return;

      // Only sync if invoice exists (any of these signals)
      const hasInvoice =
        (colInvoiceLink && r[colInvoiceLink - 1]) ||
        (colInvoicePrev && r[colInvoicePrev - 1]) ||
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

      rowObj['Shipment ID'] = shipmentId;
      rowObj['Supplier / Factory'] = colSeller ? (r[colSeller - 1] || '') : '';
      rowObj['Forwarder'] = colPlatform ? (r[colPlatform - 1] || '') : '';
      rowObj['Tracking / Container'] = '';
      rowObj['Purchases Line ID'] = lineId;

      rowObj['Order ID (Batch)'] = orderId;

      const orderDate = colOrderDate ? r[colOrderDate - 1] : null;
      rowObj['Ship Date'] = (orderDate instanceof Date) ? orderDate : new Date();
      rowObj['ETA'] = '';
      rowObj['Actual Arrival'] = '';

      rowObj['SKU'] = sku;
      rowObj['Product Name'] = colProduct ? (r[colProduct - 1] || '') : '';
      rowObj['Variant / Color'] = variant || '';
      rowObj['Qty'] = qty;

      rowObj['Gross Weight (kg)'] = '';
      rowObj['Volume (CBM)'] = '';
      rowObj['Freight (AED)'] = '';
      rowObj['Other Fees (AED)'] = '';
      rowObj['Total Cost (AED)'] = '';

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
      } catch (e) { }
    }

    // Append new rows
    if (newRows.length) {
      const startRow = shipSh.getLastRow() + 1;
      shipSh.getRange(startRow, 1, newRows.length, shipLastCol).setValues(newRows);
    }

    if (newRows.length || updatedQtyCount) {
      try { rebuildShipmentsCnUaeStatus_(); } catch (e) { }
      try { setupShipmentsCnUaeStatusValidation_(); } catch (e) { }
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
  if (s.indexOf('uae-kor') !== -1) return 'UAE-KOR';

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
    const map = getHeaderMap_(invSh);

    const colSku = map['SKU'];
    const colWh = map['Warehouse (UAE)'];
    const colProduct = map['Product Name'];
    const colVar = map['Variant / Color'];
    const colOnHand = map['On Hand Qty'];
    const colAvail = map['Available Qty'];
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
        variant: colVar ? (row[colVar - 1] || '') : '',
        warehouse: rowWhRaw,
        onHand: colOnHand ? Number(row[colOnHand - 1] || 0) : 0,
        available: colAvail ? Number(row[colAvail - 1] || 0) : 0,
        avgCost: colAvgCost ? Number(row[colAvgCost - 1] || 0) : 0
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

  const colSku = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
  const colProd = map['Product Name'];
  const colVar = map['Variant / Color'];
  const colNotes = map['Notes'];
  const colCourier = map['Courier'] || map['Courier Name'];
  const colWhUae = map['Warehouse (UAE)']; // اختياري في Shipments_UAE_EG

  if (!colSku) return;

  const sku = (sh.getRange(rowIndex, colSku).getValue() || '').toString().trim();

  // If SKU cleared → clear related cells
  if (!sku) {
    if (colProd) sh.getRange(rowIndex, colProd).clearContent();
    if (colVar) sh.getRange(rowIndex, colVar).clearContent();
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
    if (colProd) sh.getRange(rowIndex, colProd).setValue('');
    if (colVar) sh.getRange(rowIndex, colVar).setValue('');
    if (colNotes) {
      sh.getRange(rowIndex, colNotes)
        .setNote('SKU not found in Inventory_UAE.');
    }
    return;
  }

  if (colProd) sh.getRange(rowIndex, colProd).setValue(info.productName);
  if (colVar) sh.getRange(rowIndex, colVar).setValue(info.variant);

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
    const sh = getSheet_(APP.SHEETS.SHIP_UAE_EG);
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
 * Seed Shipments_UAE_EG planning rows from Inventory_UAE.
 * - Creates/updates rows where Shipment ID is blank (planning rows)
 * - Key = SKU + Warehouse (UAE)
 * - Qty remains 0 (you fill shipped qty when you actually ship)
 */
function seedShipmentsUaeEgFromInventoryUae() {
  try {
    ensureErrorLog_();

    const invSh = getSheet_(APP.SHEETS.INVENTORY_UAE);
    const shipSh = getSheet_(APP.SHEETS.SHIP_UAE_EG);

    // Repair known blank header (Status → Warehouse (UAE)) and ensure required schema additions (e.g., Line ID).
    try {
      if (typeof SHIP_UAE_EG_HEADERS !== 'undefined') {
        repairBlankHeadersByPosition_(APP.SHEETS.SHIP_UAE_EG, SHIP_UAE_EG_HEADERS, 1);
        ensureSheetSchema_(APP.SHEETS.SHIP_UAE_EG, SHIP_UAE_EG_HEADERS, { addMissing: true });
      }
    } catch (e) {
      logError_('seedShipmentsUaeEgFromInventoryUae.preflight', e);
    }

    const invMap = getHeaderMap_(invSh);
    const shipMap = getHeaderMap_(shipSh);

    const invSkuCol = invMap[APP.COLS.INV_UAE.SKU];
    const invWhCol = invMap[APP.COLS.INV_UAE.WAREHOUSE];
    const invPnCol = invMap[APP.COLS.INV_UAE.PRODUCT_NAME];
    const invVarCol = invMap[APP.COLS.INV_UAE.VARIANT];
    const invQtyCol = invMap[APP.COLS.INV_UAE.ON_HAND];

    const shipIdCol = shipMap[APP.COLS.SHIP_UAE_EG.SHIPMENT_ID];
    const shipWhCol = shipMap[APP.COLS.SHIP_UAE_EG.WAREHOUSE_UAE];
    const shipSkuCol = shipMap[APP.COLS.SHIP_UAE_EG.SKU];
    const shipPnCol = shipMap[APP.COLS.SHIP_UAE_EG.PRODUCT_NAME];
    const shipVarCol = shipMap[APP.COLS.SHIP_UAE_EG.VARIANT];
    const shipQtyCol = shipMap[APP.COLS.SHIP_UAE_EG.QTY];
    const shipQtySyncedCol = shipMap[APP.COLS.SHIP_UAE_EG.QTY_SYNCED];
    const shipStatusCol = shipMap[APP.COLS.SHIP_UAE_EG.STATUS];
    const shipNotesCol = shipMap['Notes'];

    if (!invSkuCol || !invWhCol) throw new Error('Inventory_UAE missing required headers (SKU/Warehouse).');
    if (!shipIdCol || !shipWhCol || !shipSkuCol) {
      const err = new Error('Shipments_UAE_EG missing required headers (Shipment ID / Warehouse (UAE) / SKU). Run Setup Shipments Layouts.');
      logError_('seedShipmentsUaeEgFromInventoryUae', err, {
        hasShipmentId: !!shipIdCol,
        hasWarehouseUAE: !!shipWhCol,
        hasSku: !!shipSkuCol
      });
      return; // non-fatal: allow snapshot rebuild to continue
    }

    const invLast = invSh.getLastRow();
    if (invLast < 2) return;

    const invData = invSh.getRange(2, 1, invLast - 1, invSh.getLastColumn()).getValues();

    // Build inventory map: key = SKU||WH
    const invByKey = new Map();
    for (const r of invData) {
      const sku = String(r[invSkuCol - 1] || '').trim();
      const wh = normalizeWarehouseCode_(r[invWhCol - 1]);
      if (!sku || !wh) continue;

      const onHand = invQtyCol ? Number(r[invQtyCol - 1] || 0) : 0;
      if (onHand <= 0) continue; // seed only what exists

      invByKey.set(`${sku}||${wh}`, {
        sku,
        wh,
        productName: invPnCol ? String(r[invPnCol - 1] || '').trim() : '',
        variant: invVarCol ? String(r[invVarCol - 1] || '').trim() : '',
        onHand
      });
    }

    if (invByKey.size === 0) return;

    // Read shipments existing rows
    const shipLast = shipSh.getLastRow();
    const shipCols = shipSh.getLastColumn();
    const shipData = shipLast >= 2 ? shipSh.getRange(2, 1, shipLast - 1, shipCols).getValues() : [];

    // Index existing planning rows (Shipment ID blank): key = SKU||WH => rowIndex
    const existingPlanRowByKey = new Map();
    for (let i = 0; i < shipData.length; i++) {
      const row = shipData[i];
      const shipId = String(row[shipIdCol - 1] || '').trim();
      if (shipId) continue; // only planning rows

      const sku = String(row[shipSkuCol - 1] || '').trim();
      const wh = normalizeWarehouseCode_(row[shipWhCol - 1]);
      if (!sku || !wh) continue;

      existingPlanRowByKey.set(`${sku}||${wh}`, i);
    }

    // Update existing planning rows + append new ones
    const rowsToWrite = [];
    const rowNumbersToWrite = []; // 2-based sheet row number

    // Update existing
    for (const [key, info] of invByKey.entries()) {
      if (!existingPlanRowByKey.has(key)) continue;

      const i = existingPlanRowByKey.get(key);
      const row = shipData[i];

      // Fill standard fields but don't touch user-entered qty/shipping fields
      row[shipWhCol - 1] = info.wh;
      row[shipSkuCol - 1] = info.sku;

      if (shipPnCol && !String(row[shipPnCol - 1] || '').trim()) row[shipPnCol - 1] = info.productName;
      if (shipVarCol && !String(row[shipVarCol - 1] || '').trim()) row[shipVarCol - 1] = info.variant;

      if (shipQtyCol && !Number(row[shipQtyCol - 1] || 0)) row[shipQtyCol - 1] = 0;
      if (shipQtySyncedCol && !Number(row[shipQtySyncedCol - 1] || 0)) row[shipQtySyncedCol - 1] = 0;

      if (shipStatusCol && !String(row[shipStatusCol - 1] || '').trim()) row[shipStatusCol - 1] = 'Planned';

      if (shipNotesCol) {
        const note = `Seeded from Inventory_UAE (OnHand: ${info.onHand})`;
        if (!String(row[shipNotesCol - 1] || '').includes('Seeded from Inventory_UAE')) {
          row[shipNotesCol - 1] = note;
        }
      }

      rowsToWrite.push(row);
      rowNumbersToWrite.push(i + 2); // +2 because data starts at row 2
    }

    // Append missing planning rows
    const headerRow = shipSh.getRange(1, 1, 1, shipCols).getValues()[0];
    const headerIndex = {};
    headerRow.forEach((h, idx) => headerIndex[String(h || '').trim()] = idx);

    const newRows = [];
    for (const [key, info] of invByKey.entries()) {
      if (existingPlanRowByKey.has(key)) continue;

      const newRow = new Array(shipCols).fill('');
      newRow[shipIdCol - 1] = '';                // planning row
      newRow[shipWhCol - 1] = info.wh;
      newRow[shipSkuCol - 1] = info.sku;
      if (shipPnCol) newRow[shipPnCol - 1] = info.productName;
      if (shipVarCol) newRow[shipVarCol - 1] = info.variant;
      if (shipQtyCol) newRow[shipQtyCol - 1] = 0;
      if (shipQtySyncedCol) newRow[shipQtySyncedCol - 1] = 0;
      if (shipStatusCol) newRow[shipStatusCol - 1] = 'Planned';
      if (shipNotesCol) newRow[shipNotesCol - 1] = `Seeded from Inventory_UAE (OnHand: ${info.onHand})`;

      newRows.push(newRow);
    }

    // Write updates (row-by-row, but only the changed rows)
    for (let k = 0; k < rowsToWrite.length; k++) {
      shipSh.getRange(rowNumbersToWrite[k], 1, 1, shipCols).setValues([rowsToWrite[k]]);
    }

    // Append new ones in a single batch
    if (newRows.length) {
      shipSh.getRange(shipSh.getLastRow() + 1, 1, newRows.length, shipCols).setValues(newRows);
    }

  } catch (e) {
    logError_('seedShipmentsUaeEgFromInventoryUae', e);
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

  const colQty = map[APP.COLS.SHIP_UAE_EG.QTY] ||
    map['Qty'];
  const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
    map['Ship Cost (EGP) - per unit or box'] ||
    map['Ship Cost (EGP)'];
  const colCustoms = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
    map['Customs (EGP)'] ||
    map['Customs / Clearance (EGP)'];
  const colOther = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
    map['Other (EGP)'] ||
    map['Other Fees (EGP)'];
  const colTotal = map[APP.COLS.SHIP_UAE_EG.TOTAL_COST] ||
    map['Total Cost (EGP)'];

  const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
    map['Ship Date'] ||
    map['Ship Date (UAE)'];
  const colEta = map[APP.COLS.SHIP_UAE_EG.ETA] ||
    map['ETA'];
  const colArr = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
    map['Actual Arrival'] ||
    map['Actual Arrival (EG)'];
  const colStatus = map[APP.COLS.SHIP_UAE_EG.STATUS] ||
    map['Status'];

  const colSku = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
  const colNotes = map['Notes'];
  const colCourier = map['Courier'] || map['Courier Name'];

  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

  // ----- Total Cost (EGP) -----
  const qty = colQty ? Number(row[colQty - 1] || 0) : 0;
  const shipCost = colShipCost ? Number(row[colShipCost - 1] || 0) : 0;
  const customs = colCustoms ? Number(row[colCustoms - 1] || 0) : 0;
  const other = colOther ? Number(row[colOther - 1] || 0) : 0;

  if (colTotal) {
    const totalShipment = (shipCost * qty) + customs + other;
    sh.getRange(rowIndex, colTotal).setValue(totalShipment);
    sh.getRange(rowIndex, colTotal).setNumberFormat('0.00');
  }

  // ----- Status -----
  const ship = colShipDate ? row[colShipDate - 1] : '';
  const eta = colEta ? row[colEta - 1] : '';
  const arr = colArr ? row[colArr - 1] : '';

  let status = '';

  if (arr) {
    status = SHIPMENT_STATUS.ARRIVED_EG;
  } else if (ship && eta) {
    const today = new Date();
    const todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const etaMid = new Date(eta.getFullYear(), eta.getMonth(), eta.getDate());
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

    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    if (sheetName !== APP.SHEETS.SHIP_UAE_EG && sheetName !== 'Shipments_UAE_EG') {
      return;
    }

    const rowIndex = e.range.getRow();
    if (rowIndex === 1) return; // header

    const map = getHeaderMap_(sh);

    const colSku = map[APP.COLS.SHIP_UAE_EG.SKU] || map['SKU'];
    const colQty = map[APP.COLS.SHIP_UAE_EG.QTY] ||
      map['Qty'];
    const colShipCost = map[APP.COLS.SHIP_UAE_EG.SHIP_COST] ||
      map['Ship Cost (EGP) - per unit or box'] ||
      map['Ship Cost (EGP)'];
    const colCustoms = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
      map['Customs (EGP)'] ||
      map['Customs / Clearance (EGP)'];
    const colOther = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
      map['Other (EGP)'] ||
      map['Other Fees (EGP)'];

    const colShipDate = map[APP.COLS.SHIP_UAE_EG.SHIP_DATE] ||
      map['Ship Date'] ||
      map['Ship Date (UAE)'];
    const colEta = map[APP.COLS.SHIP_UAE_EG.ETA] ||
      map['ETA'];
    const colArr = map[APP.COLS.SHIP_UAE_EG.ARRIVAL] ||
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
    const orderIdFilter = (typeof safePromptText_ === 'function')
      ? safePromptText_(
        'Generate QC_UAE',
        'Optional: Enter a single Order ID to generate QC only for that order.\n\nLeave empty to generate QC for all eligible orders.'
        , null, { ui: true })
      : null;
    if (orderIdFilter === null) return;
    qc_generateFromPurchases_(orderIdFilter || null);
    if (typeof safeAlert_ === 'function') safeAlert_('QC_UAE generation triggered.', null, { ui: true });
  } catch (err) {
    logError_('qc_generateFromPurchasesPrompt', err);
    throw err;
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
function qc_generateFromPurchases_(optOrderIdOrOrderIds) {
  try {
    const purchSh = getSheet_(APP.SHEETS.PURCHASES);
    const qcSh = getSheet_(APP.SHEETS.QC_UAE);
    const shipSh = getSheet_(APP.SHEETS.SHIP_CN_UAE);

    const pMap = getHeaderMap_(purchSh);
    const qcMap = getHeaderMap_(qcSh);
    const shipMap = getHeaderMap_(shipSh);

    const lastPurRow = purchSh.getLastRow();
    if (lastPurRow < 2) {
      safeAlert_('No data found in Purchases.');
      return;
    }

    // Ensure Purchases Line IDs exist (idempotent)
    try {
      if (typeof purchases_ensureLineIds_ === 'function') {
        purchases_ensureLineIds_(purchSh, pMap, 2, lastPurRow - 1);
      }
    } catch (e) { }

    // Purchases columns
    const colOrderId = pMap[APP.COLS.PURCHASES.ORDER_ID] || pMap['Order ID'];
    const colSku = pMap[APP.COLS.PURCHASES.SKU] || pMap['SKU'];
    const colQty = pMap[APP.COLS.PURCHASES.QTY] || pMap['Qty'];
    const colLineId = pMap[APP.COLS.PURCHASES.LINE_ID] || pMap['Line ID'];

    const colProduct = pMap[APP.COLS.PURCHASES.PRODUCT_NAME] || pMap[APP.COLS.PURCHASES.PRODUCT] || pMap['Product Name'];
    const colVariant = pMap[APP.COLS.PURCHASES.VARIANT] || pMap['Variant / Color'];
    const colBatch = pMap[APP.COLS.PURCHASES.BATCH_CODE] || pMap['Batch Code'];

    if (!colOrderId || !colSku || !colQty) {
      throw new Error('Missing required Purchases columns (Order ID / SKU / Qty).');
    }
    if (!colLineId) {
      throw new Error('Missing Purchases column "Line ID". Run Purchases layout/repair first.');
    }

    const purchData = purchSh
      .getRange(2, 1, lastPurRow - 1, purchSh.getLastColumn())
      .getValues();

    // QC columns
    const qcColQcId = qcMap[APP.COLS.QC_UAE.QC_ID] || qcMap['QC ID'];
    const qcColOrderId = qcMap[APP.COLS.QC_UAE.ORDER_ID] || qcMap['Order ID'];
    const qcColShipId = qcMap[APP.COLS.QC_UAE.SHIPMENT_ID] || qcMap['Shipment CN→UAE ID'] || qcMap['Shipment ID'];
    const qcColSku = qcMap[APP.COLS.QC_UAE.SKU] || qcMap['SKU'];
    const qcColBatch = qcMap[APP.COLS.QC_UAE.BATCH_CODE] || qcMap['Batch Code'];
    const qcColProduct = qcMap['Product Name'] || qcMap[APP.COLS.QC_UAE.PRODUCT_NAME] || qcMap[APP.COLS.PURCHASES.PRODUCT_NAME];
    const qcColVariant = qcMap['Variant / Color'] || qcMap[APP.COLS.QC_UAE.VARIANT] || qcMap[APP.COLS.PURCHASES.VARIANT];
    const qcColQtyOrdered = qcMap['Qty Ordered'] || qcMap[APP.COLS.QC_UAE.QTY_ORDERED];
    const qcColPurchLineId = qcMap[APP.COLS.QC_UAE.PURCHASE_LINE_ID] || qcMap['Purchases Line ID'];

    const qcColQtyReceived = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
    const qcColQtyDef = qcMap[APP.COLS.QC_UAE.QTY_DEFECT] || qcMap['Qty Defective'];
    const qcColQtyMissing = qcMap[APP.COLS.QC_UAE.QTY_MISSING] || qcMap['Qty Missing'];
    const qcColQtyOk = qcMap[APP.COLS.QC_UAE.QTY_OK] || qcMap['Qty OK'];

    if (!qcColPurchLineId) {
      throw new Error('Missing QC_UAE column "Purchases Line ID". Run Logistics → Setup QC Layouts.');
    }
    if (!qcColOrderId || !qcColSku) {
      throw new Error('Missing required QC_UAE columns (Order ID / SKU).');
    }

    // Shipments map (prefer line-level) + ARRIVED gate
    const shipColShipId = shipMap[APP.COLS.SHIP_CN_UAE.SHIPMENT_ID] || shipMap['Shipment ID'];
    const shipColOrderId = shipMap[APP.COLS.SHIP_CN_UAE.ORDER_BATCH] || shipMap['Order ID (Batch)'] || shipMap['Order ID'];
    const shipColSku = shipMap[APP.COLS.SHIP_CN_UAE.SKU] || shipMap['SKU'];
    const shipColVariant = shipMap[APP.COLS.SHIP_CN_UAE.VARIANT] || shipMap['Variant / Color'];
    const shipColLineId = shipMap[APP.COLS.SHIP_CN_UAE.PURCHASE_LINE_ID] || shipMap['Purchases Line ID'];
    const shipColStatus = shipMap[APP.COLS.SHIP_CN_UAE.STATUS] || shipMap['Status'];
    const shipColArrival = shipMap[APP.COLS.SHIP_CN_UAE.ARRIVAL] || shipMap['Actual Arrival'];

    const shipIdByLineId = {};
    const shipIdByKey = {};

    // Eligibility (only Arrived UAE lines / keys)
    const arrivedShipIdByLineId = {};
    const arrivedShipIdByKey = {};

    const shipLastRow = shipSh.getLastRow();
    if (shipLastRow >= 2 && shipColShipId) {
      const shipData = shipSh.getRange(2, 1, shipLastRow - 1, shipSh.getLastColumn()).getValues();

      shipData.forEach(function (r) {
        const sid = String(r[shipColShipId - 1] || '').trim();
        if (!sid) return;

        const oid = shipColOrderId ? String(r[shipColOrderId - 1] || '').trim() : '';
        const sku = shipColSku ? String(r[shipColSku - 1] || '').trim() : '';
        const v = shipColVariant ? String(r[shipColVariant - 1] || '').trim() : '';

        const status = shipColStatus ? String(r[shipColStatus - 1] || '').trim() : '';
        const arrival = shipColArrival ? r[shipColArrival - 1] : null;

        const isArrivedUAE = !!arrival || status === 'Arrived UAE';

        // Line-level key
        if (shipColLineId) {
          const lid = String(r[shipColLineId - 1] || '').trim();
          if (lid) {
            if (!shipIdByLineId[lid]) shipIdByLineId[lid] = sid;
            if (isArrivedUAE && !arrivedShipIdByLineId[lid]) arrivedShipIdByLineId[lid] = sid;
          }
        }

        // Fallback key (order+sku+variant)
        if (oid && sku) {
          const key = oid + '||' + sku + '||' + v;
          if (!shipIdByKey[key]) shipIdByKey[key] = sid;
          if (isArrivedUAE && !arrivedShipIdByKey[key]) arrivedShipIdByKey[key] = sid;
        }
      });
    }


    // QC sheet metadata
    const qcLastCol = qcSh.getLastColumn();
    const qcHeaders = qcSh.getRange(1, 1, 1, qcLastCol).getValues()[0];
    const formulasRow2 = qcSh.getRange(2, 1, 1, qcLastCol).getFormulas()[0];

    function _pickAnchorCol_() {
      const candidates = [qcColPurchLineId, qcColQcId, qcColOrderId, qcColSku].filter(function (c) { return !!c; });
      for (let i = 0; i < candidates.length; i++) {
        const c = candidates[i];
        if (!formulasRow2[c - 1]) return c;
      }
      return candidates.length ? candidates[0] : 1;
    }

    const anchorCol = _pickAnchorCol_();
    let dataLastRow = 1;
    try {
      dataLastRow = qcSh.getRange(qcSh.getMaxRows(), anchorCol)
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      if (dataLastRow < 2) dataLastRow = 1;
    } catch (e) {
      dataLastRow = qcSh.getLastRow();
      if (dataLastRow < 2) dataLastRow = 1;
    }

    const existingN = (dataLastRow >= 2) ? (dataLastRow - 1) : 0;
    const existingData = existingN
      ? qcSh.getRange(2, 1, existingN, qcLastCol).getValues()
      : [];

    const existingByLineId = {};
    let maxSeq = 0;

    existingData.forEach(function (row, i) {
      const sheetRow = i + 2;
      const lid = String(row[qcColPurchLineId - 1] || '').trim();
      if (lid && !existingByLineId[lid]) existingByLineId[lid] = sheetRow;

      if (qcColQcId) {
        const qcId = String(row[qcColQcId - 1] || '').trim();
        const m = qcId.match(/(\d+)$/);
        if (m) {
          const n = parseInt(m[1], 10);
          if (n > maxSeq) maxSeq = n;
        }
      }
    });

    const updates = [];
    const newRowsFull = [];
    const filterSet = (function () {
      if (Array.isArray(optOrderIdOrOrderIds)) {
        const set = {};
        (optOrderIdOrOrderIds || []).forEach(function (x) {
          const s = String(x || '').trim();
          if (s) set[s] = true;
        });
        return Object.keys(set).length ? set : null;
      }
      const s = String(optOrderIdOrOrderIds || '').trim();
      if (!s) return null;
      const set = {};
      set[s] = true;
      return set;
    })();

    const seenInRun = {};

    purchData.forEach(function (r) {
      const orderId = String(r[colOrderId - 1] || '').trim();
      if (!orderId) return;
      if (filterSet && !filterSet[orderId]) return;

      const sku = String(r[colSku - 1] || '').trim();
      const qty = Number(r[colQty - 1] || 0);
      const lineId = String(r[colLineId - 1] || '').trim();

      if (!sku || !qty || !lineId) return;

      if (seenInRun[lineId]) return;
      seenInRun[lineId] = true;

      const variant = colVariant ? String(r[colVariant - 1] || '').trim() : '';
      const product = colProduct ? String(r[colProduct - 1] || '').trim() : '';
      const batch = colBatch ? String(r[colBatch - 1] || '').trim() : '';

      const shipKey = orderId + '||' + sku + '||' + variant;

      // Only allow QC when Arrived UAE
      const arrivedShipId = (arrivedShipIdByLineId[lineId] || arrivedShipIdByKey[shipKey] || '');
      if (!arrivedShipId) return;

      const shipId = arrivedShipId;

      const existingRow = existingByLineId[lineId];
      if (existingRow) {
        updates.push({
          row: existingRow,
          qtyOrdered: qty,
          shipId: shipId,
          product: product,
          variant: variant,
          batch: batch
        });
        return;
      }

      maxSeq++;
      const qcId = 'QC-' + Utilities.formatString('%06d', maxSeq);

      const rowObj = {};
      rowObj['QC ID'] = qcId;
      rowObj['Order ID'] = orderId;
      rowObj['Shipment CN→UAE ID'] = shipId;
      rowObj['SKU'] = sku;
      rowObj['Batch Code'] = batch || '';
      rowObj['Product Name'] = product || '';
      rowObj['Variant / Color'] = variant || '';
      rowObj['Qty Ordered'] = qty;
      rowObj['Qty Received'] = '';
      rowObj['Qty Defective'] = '';
      rowObj['Qty Missing'] = '';
      rowObj['Qty OK'] = '';
      rowObj['Notes'] = 'Auto (line-level) from Purchases';
      rowObj['Purchases Line ID'] = lineId;

      const outRow = qcHeaders.map(function (h) {
        return (rowObj[h] !== undefined) ? rowObj[h] : '';
      });

      newRowsFull.push(outRow);
    });

    if (!newRowsFull.length && !updates.length) {
      safeAlert_('No new QC rows to generate.');
      return;
    }

    // 1) Updates
    if (updates.length) {
      updates.forEach(function (u) {
        const r = u.row;

        if (qcColQtyOrdered) qcSh.getRange(r, qcColQtyOrdered).setValue(u.qtyOrdered);

        if (qcColShipId && u.shipId) {
          const cur = qcSh.getRange(r, qcColShipId).getValue();
          if (!cur) qcSh.getRange(r, qcColShipId).setValue(u.shipId);
        }
        if (qcColBatch && u.batch) {
          const cur = qcSh.getRange(r, qcColBatch).getValue();
          if (!cur) qcSh.getRange(r, qcColBatch).setValue(u.batch);
        }
        if (qcColProduct && u.product) {
          const cur = qcSh.getRange(r, qcColProduct).getValue();
          if (!cur) qcSh.getRange(r, qcColProduct).setValue(u.product);
        }
        if (qcColVariant && u.variant) {
          const cur = qcSh.getRange(r, qcColVariant).getValue();
          if (!cur) qcSh.getRange(r, qcColVariant).setValue(u.variant);
        }
      });
    }

    // 2) Insert at TOP without breaking ARRAYFORMULA anchors (shift non-formula columns only)
    if (newRowsFull.length) {
      const insertCount = newRowsFull.length;

      const shiftCols = [];
      for (let c = 1; c <= qcLastCol; c++) {
        if (!formulasRow2[c - 1]) shiftCols.push(c);
      }

      const segments = [];
      if (shiftCols.length) {
        let start = shiftCols[0];
        let prev = start;
        for (let i = 1; i < shiftCols.length; i++) {
          const cur = shiftCols[i];
          if (cur === prev + 1) {
            prev = cur;
            continue;
          }
          segments.push({ startCol: start, numCols: (prev - start + 1) });
          start = cur;
          prev = cur;
        }
        segments.push({ startCol: start, numCols: (prev - start + 1) });
      }

      const targetLastRow = 1 + existingN + insertCount;
      if (qcSh.getMaxRows() < targetLastRow) {
        qcSh.insertRowsAfter(qcSh.getMaxRows(), targetLastRow - qcSh.getMaxRows());
      }

      segments.forEach(function (seg) {
        const startCol = seg.startCol;
        const numCols = seg.numCols;

        const existingSeg = existingN
          ? qcSh.getRange(2, startCol, existingN, numCols).getValues()
          : [];

        const out = [];
        for (let i = 0; i < insertCount; i++) {
          out.push(newRowsFull[i].slice(startCol - 1, startCol - 1 + numCols));
        }
        for (let i = 0; i < existingSeg.length; i++) out.push(existingSeg[i]);

        qcSh.getRange(2, startCol, out.length, numCols).setValues(out);
      });
    }

    // Optional: recalc derived columns only if NOT formula-driven at row2
    try {
      const canRecalcMissingOk =
        qcColQtyMissing && qcColQtyOk &&
        !formulasRow2[qcColQtyMissing - 1] &&
        !formulasRow2[qcColQtyOk - 1];

      if (typeof qc_recalcRows_ === 'function' && canRecalcMissingOk) {
        const rowsToRecalc = Math.min(1000, (existingN + newRowsFull.length + 5));
        qc_recalcRows_(qcSh, 2, rowsToRecalc);
      }
    } catch (e) { }

    try { setupQcWarehouseValidation_(); } catch (e) { }

    safeAlert_(
      'QC_UAE generation done.\n' +
      'Inserted rows: ' + newRowsFull.length + '\n' +
      'Updated rows: ' + updates.length
    );

  } catch (e) {
    logError_('qc_generateFromPurchases_', e, { optOrderIdOrOrderIds: optOrderIdOrOrderIds });
    throw e;
  }
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
  const qcSh = getSheet_(APP.SHEETS.QC_UAE);
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
  const colRecv = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
  const colDef = qcMap[APP.COLS.QC_UAE.QTY_DEFECT] || qcMap['Qty Defective'];

  if (!colOrdered || !colRecv || !colDef) return;

  // Only react when edit intersects the driving columns
  const colStart = e.range.getColumn();
  const colEnd = colStart + e.range.getNumColumns() - 1;

  const drives = [colOrdered, colRecv, colDef];
  const intersects = drives.some(c => c >= colStart && c <= colEnd);
  if (!intersects) return;

  const rowStart = Math.max(2, e.range.getRow());
  const rowEnd = Math.max(rowStart, e.range.getRow() + e.range.getNumRows() - 1);
  const numRows = rowEnd - rowStart + 1;

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
  const idxOrd = qcMap['Qty Ordered'];
  const idxRecv = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
  const idxDef = qcMap[APP.COLS.QC_UAE.QTY_DEFECT] || qcMap['Qty Defective'];
  const idxMiss = qcMap[APP.COLS.QC_UAE.QTY_MISSING] || qcMap['Qty Missing'];
  const idxOk = qcMap[APP.COLS.QC_UAE.QTY_OK] || qcMap['Qty OK'];
  const idxRes = qcMap['QC Result'];
  const idxDate = qcMap['QC Date'];

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

    const qtyOrd = Number(row[idxOrd - 1] || 0);
    const qtyRecv = Number(row[idxRecv - 1] || 0);
    const qtyDef = Number(row[idxDef - 1] || 0);

    const missing = Math.max(qtyOrd - qtyRecv, 0);
    const ok = Math.max(qtyRecv - qtyDef, 0);

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
function syncQCtoInventory_UAE(opts) {
  const interactive = !!(opts && opts.interactive);
  try {
    const qcSh = getSheet_(APP.SHEETS.QC_UAE);
    const purchSh = getSheet_(APP.SHEETS.PURCHASES);
    const ledgerSh = getSheet_(APP.SHEETS.INVENTORY_TXNS);

    const qcMap = getHeaderMap_(qcSh);
    const purchMap = getHeaderMap_(purchSh);
    const ledgerMap = getHeaderMap_(ledgerSh);

    const lastQcRow = qcSh.getLastRow();
    if (lastQcRow < 2) return; // مفيش بيانات

    const qcData = qcSh.getRange(2, 1, lastQcRow - 1, qcSh.getLastColumn()).getValues();

    const qcIdCol = qcMap[APP.COLS.QC_UAE.QC_ID] || qcMap['QC ID'];
    const qcDateCol = qcMap[APP.COLS.QC_UAE.QC_DATE] || qcMap['QC Date'];

    // ===== Build cost map from Purchases =====
    const purchLast = purchSh.getLastRow();
    const costMap = {};
    const priceMap = {};
    const currencyMap = {};

    if (purchLast >= 2) {
      const purchData = purchSh.getRange(2, 1, purchLast - 1, purchSh.getLastColumn()).getValues();

      const idxOrder = purchMap[APP.COLS.PURCHASES.ORDER_ID] ? purchMap[APP.COLS.PURCHASES.ORDER_ID] - 1 : null;
      const idxSku = purchMap[APP.COLS.PURCHASES.SKU] ? purchMap[APP.COLS.PURCHASES.SKU] - 1 : null;
      const idxBatch = purchMap[APP.COLS.PURCHASES.BATCH_CODE] ? purchMap[APP.COLS.PURCHASES.BATCH_CODE] - 1 : null;

      // Prefer Unit Landed Cost, fallback to Unit Price (EGP)
      const idxUnitLanded = purchMap[APP.COLS.PURCHASES.UNIT_LANDED] ? purchMap[APP.COLS.PURCHASES.UNIT_LANDED] - 1 : null;
      const idxUnitNet = purchMap[APP.COLS.PURCHASES.NET_UNIT_PRICE] ? purchMap[APP.COLS.PURCHASES.NET_UNIT_PRICE] - 1 : null;
      const idxTotalOrig = purchMap[APP.COLS.PURCHASES.TOTAL_ORIG] ? purchMap[APP.COLS.PURCHASES.TOTAL_ORIG] - 1 : null;
      const idxQty = purchMap[APP.COLS.PURCHASES.QTY] ? purchMap[APP.COLS.PURCHASES.QTY] - 1 : null;
      const idxCurrency = purchMap[APP.COLS.PURCHASES.CURRENCY] ? purchMap[APP.COLS.PURCHASES.CURRENCY] - 1 : null;

      const parseNum_ = function (v) {
        if (typeof v === 'number') return v;
        const s = String(v || '').replace(/[^0-9.\-]/g, '');
        return s ? Number(s) : 0;
      };

      purchData.forEach(function (r) {
        if (idxOrder == null || idxSku == null) return;

        const orderId = r[idxOrder];
        const sku = (r[idxSku] || '').toString().trim();
        if (!orderId || !sku) return;

        const baseKey = String(orderId) + '||' + String(sku);
        const keyOrdSku = 'ORDSKU||' + baseKey;

        const cur = (idxCurrency != null) ? (r[idxCurrency] || '').toString().trim() : '';
        if (cur) currencyMap[keyOrdSku] = cur;

        let unitPriceOrig = 0;
        if (idxTotalOrig != null && idxQty != null) {
          const tot = parseNum_(r[idxTotalOrig]);
          const q = parseNum_(r[idxQty]);
          if (q > 0) unitPriceOrig = tot / q;
        }
        if (unitPriceOrig > 0) priceMap[keyOrdSku] = unitPriceOrig;

        // Prefer Unit Landed Cost, fallback to Net Unit Price (EGP)
        const unitCost = (idxUnitLanded != null)
          ? Number(r[idxUnitLanded] || 0)
          : (idxUnitNet != null ? Number(r[idxUnitNet] || 0) : 0);

        if (unitCost > 0) costMap[keyOrdSku] = unitCost;

        // Batch mappings (optional but helpful)
        if (idxBatch != null) {
          const batch = (r[idxBatch] || '').toString().trim();
          if (batch) {
            const keyBatch = 'BATCH||' + batch;
            if (unitCost > 0) costMap[keyBatch] = unitCost;
            if (cur) currencyMap[keyBatch] = cur;
            if (unitPriceOrig > 0) priceMap[keyBatch] = unitPriceOrig;
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
      const idxSrcId = ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] ? ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] - 1 : null;
      const idxNotes = ledgerMap[APP.COLS.INV_TXNS.NOTES] ? ledgerMap[APP.COLS.INV_TXNS.NOTES] - 1 : null;

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
      const sku = (row[qcMap[APP.COLS.QC_UAE.SKU] - 1] || '').toString().trim();
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

      const product = row[qcMap['Product Name'] - 1] || '';
      const variant = row[qcMap['Variant / Color'] - 1] || '';
      const warehouseRaw = (row[qcMap[APP.COLS.QC_UAE.WAREHOUSE] - 1] || '').toString().trim();
      const warehouse = (typeof normalizeWarehouseCode_ === 'function') ? normalizeWarehouseCode_(warehouseRaw) : warehouseRaw;
      if (!warehouse) {
        logError_('syncQCtoInventory_UAE', new Error('Missing Warehouse (UAE) in QC_UAE'), { qcId: qcId, qcRow: (2 + idx) });
        return; // skip this QC row
      }

      // Qty In logic:
      //  - If Qty OK exists → use it
      //  - Else fallback Qty Received - Qty Defective
      let qtyIn = 0;

      const qtyOkCol = qcMap[APP.COLS.QC_UAE.QTY_OK] || qcMap['Qty OK'];
      const qtyRecCol = qcMap[APP.COLS.QC_UAE.QTY_RECEIVED] || qcMap['Qty Received'];
      const qtyDefCol = qcMap[APP.COLS.QC_UAE.QTY_DEFECT] || qcMap['Qty Defective'];

      const qtyOkVal = qtyOkCol ? Number(row[qtyOkCol - 1] || 0) : 0;


      if (qtyOkVal > 0) {
        qtyIn = qtyOkVal;
      } else if (qtyRecCol) {
        const rec = Number(row[qtyRecCol - 1] || 0);
        const def = qtyDefCol ? Number(row[qtyDefCol - 1] || 0) : 0;
        qtyIn = Math.max(0, rec - def);
      }

      if (qtyIn <= 0) return; // skip this QC row

      const qcDate = qcDateCol ? row[qcDateCol - 1] : null;

      const baseKey = String(orderId) + '||' + String(sku);
      let unitCostEgp = 0;

      if (batchCode) {
        unitCostEgp = costMap['BATCH||' + batchCode] || 0;
      }
      if (!unitCostEgp) {
        unitCostEgp = costMap['ORDSKU||' + baseKey] || 0;
      }

      const currency = currencyMap['BATCH||' + batchCode] || currencyMap['ORDSKU||' + baseKey] || '';
      const unitPriceOrig = priceMap['BATCH||' + batchCode] || priceMap['ORDSKU||' + baseKey] || 0;

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
        currency: currency,            // <-- FIX
        unitPriceOrig: unitPriceOrig,  // <-- FIX
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

    if (newTxns > 0 && typeof inv_rebuildAllSnapshots === 'function') inv_rebuildAllSnapshots();
    if (interactive && typeof safeAlert_ === 'function') {
      safeAlert_('QC_UAE Sync Done.\n\nNew txns: ' + newTxns + '\nSkipped (already synced): ' + skipped);
    } else {
      Logger.log('QC_UAE Sync Done. New txns=' + newTxns + ', skipped=' + skipped);
    }
  } catch (err) {
    logError_('syncQCtoInventory_UAE', err);
    throw err;
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
 *             IN في TAN-GH بالـ landed cost.
 */
function syncShipmentsUaeEgToInventory() {
  try {
    const runner_ = function () {
      const shipSh = getSheet_(APP.SHEETS.SHIP_UAE_EG);
      const invUaeSh = getSheet_(APP.SHEETS.INVENTORY_UAE);
      const ledgerSh = getSheet_(APP.SHEETS.INVENTORY_TXNS);

      // Repair known blank header (Status → Warehouse (UAE)) in-place and ensure persistent Line ID column exists.
      try {
        if (typeof SHIP_UAE_EG_HEADERS !== 'undefined') {
          repairBlankHeadersByPosition_(APP.SHEETS.SHIP_UAE_EG, SHIP_UAE_EG_HEADERS, 1);
          ensureSheetSchema_(APP.SHEETS.SHIP_UAE_EG, SHIP_UAE_EG_HEADERS, { addMissing: true });
        }
      } catch (e) {
        logError_('syncShipmentsUaeEgToInventory.preflight', e);
      }

      const shipMap = getHeaderMap_(shipSh);
      const invUaeMap = getHeaderMap_(invUaeSh);

      const lastShipRow = shipSh.getLastRow();
      if (lastShipRow < 2) {
        Logger.log('[SHIP_UAE_EG] No rows to sync.');
        return;
      }

      const shipData = shipSh.getRange(2, 1, lastShipRow - 1, shipSh.getLastColumn()).getValues();

      // Required columns
      const colShipmentId = shipMap[APP.COLS.SHIP_UAE_EG.SHIPMENT_ID] || shipMap['Shipment ID'];
      const colSku = shipMap[APP.COLS.SHIP_UAE_EG.SKU] || shipMap['SKU'];
      const colQty = shipMap[APP.COLS.SHIP_UAE_EG.QTY] || shipMap['Qty'];
      const colShipDate = shipMap[APP.COLS.SHIP_UAE_EG.SHIP_DATE] || shipMap['Ship Date'] || shipMap['Ship Date (UAE)'];
      const colArrival = shipMap[APP.COLS.SHIP_UAE_EG.ARRIVAL] || shipMap['Actual Arrival'] || shipMap['Actual Arrival (EG)'];
      const colQtySynced = shipMap[APP.COLS.SHIP_UAE_EG.QTY_SYNCED] || shipMap['Qty Synced'];

      if (!colShipmentId || !colSku || !colQty || !colShipDate || !colQtySynced) {
        logError_('syncShipmentsUaeEgToInventory', new Error('Missing required headers in Shipments_UAE_EG (Shipment ID / SKU / Qty / Ship Date / Qty Synced).'));
        return;
      }

      const idxShipShipmentId = colShipmentId - 1;
      const idxShipSku = colSku - 1;
      const idxShipQty = colQty - 1;
      const idxShipShipDate = colShipDate - 1;
      const idxShipArrival = colArrival ? colArrival - 1 : null;
      const idxShipQtySynced = colQtySynced - 1;

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
      const idxShipCustoms = colCustoms ? colCustoms - 1 : null;
      const idxShipOther = colOther ? colOther - 1 : null;
      const idxShipTotal = colTotal ? colTotal - 1 : null;

      // Optional columns: Product / Variant / Courier / Warehouse (UAE)
      const idxShipProdName = shipMap['Product Name'] ? shipMap['Product Name'] - 1 : null;
      const idxShipVariant = shipMap['Variant / Color'] ? shipMap['Variant / Color'] - 1 : null;
      const colBoxId = shipMap[APP.COLS.SHIP_UAE_EG.BOX_ID] || shipMap['Box ID'];
      const idxShipBoxId = colBoxId ? colBoxId - 1 : null;
      const colLineId = shipMap[APP.COLS.SHIP_UAE_EG.LINE_ID] || shipMap['Line ID'];
      const idxShipLineId = colLineId ? colLineId - 1 : null;
      const colCourier = shipMap['Courier'] || shipMap['Courier Name'];
      const idxShipCourier = colCourier ? colCourier - 1 : null;
      const colShipWhUae = shipMap[APP.COLS.SHIP_UAE_EG.WAREHOUSE_UAE] || shipMap['Warehouse (UAE)'];
      const idxShipWhUae = colShipWhUae ? colShipWhUae - 1 : null;

      // ===== Inventory_UAE map: SKU+Warehouse → cost info =====
      const uaeLastRow = invUaeSh.getLastRow();
      const uaeData = (uaeLastRow >= 2)
        ? invUaeSh.getRange(2, 1, uaeLastRow - 1, invUaeSh.getLastColumn()).getValues()
        : [];

      const idxUaeSku = invUaeMap['SKU'] ? invUaeMap['SKU'] - 1 : null;
      const idxUaeProd = invUaeMap['Product Name'] ? invUaeMap['Product Name'] - 1 : null;
      const idxUaeVariant = invUaeMap['Variant / Color'] ? invUaeMap['Variant / Color'] - 1 : null;
      const idxUaeAvgCost = invUaeMap['Avg Cost (EGP)'] ? invUaeMap['Avg Cost (EGP)'] - 1 : null;
      const idxUaeLastSrcId = invUaeMap['Last Source ID'] ? invUaeMap['Last Source ID'] - 1 : null;
      const idxUaeWh = invUaeMap['Warehouse (UAE)'] ? invUaeMap['Warehouse (UAE)'] - 1 : null;

      /** SKU+Warehouse → info */
      const uaeByWh = {};
      /** SKU → default info (first warehouse found for SKU) */
      const uaeBySku = {};

      uaeData.forEach(function (r) {
        if (idxUaeSku == null) return;
        const skuVal = r[idxUaeSku];
        if (!skuVal) return;

        const wh = idxUaeWh != null ? (r[idxUaeWh] || '').toString().trim() : '';
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
      let inCount = 0;
      let lineIdsChanged = false;

      const txns = [];

      // Build dedupe sets for SHIP_UAE_EG to prevent duplicates on retries/crashes and across code revisions
      const ledgerMap = getHeaderMap_(ledgerSh);
      const idxLTxnId =
        (ledgerMap[APP.COLS.INV_TXNS.TXN_ID] || ledgerMap['Txn ID'] || 0) - 1;
      const idxLSourceType =
        (ledgerMap[APP.COLS.INV_TXNS.SOURCE_TYPE] || ledgerMap['Source Type'] || 0) - 1;
      const idxLSourceId =
        (ledgerMap[APP.COLS.INV_TXNS.SOURCE_ID] || ledgerMap['Source ID'] || 0) - 1;

      const existingSourceIds = new Set();
      const existingTxnIds = new Set();
      const lr = ledgerSh.getLastRow();

      if (lr >= 2 && idxLSourceType >= 0) {
        const lvals = ledgerSh.getRange(2, 1, lr - 1, ledgerSh.getLastColumn()).getValues();
        for (const r of lvals) {
          const st = String(r[idxLSourceType] || '').trim();
          if (st !== 'SHIP_UAE_EG') continue;

          if (idxLSourceId >= 0) {
            const sid = String(r[idxLSourceId] || '').trim();
            if (sid) existingSourceIds.add(sid);
          }
          if (idxLTxnId >= 0) {
            const tid = String(r[idxLTxnId] || '').trim();
            if (tid) existingTxnIds.add(tid);
          }
        }
      }

      // Process shipments rows
      shipData.forEach(function (row, idx) {
        const shipmentId = row[idxShipShipmentId];
        const sku = row[idxShipSku];
        const qtyCurrent = Number(row[idxShipQty] || 0);

        if (!shipmentId || !sku || !qtyCurrent) return;

        const qtySynced = Number(row[idxShipQtySynced] || 0);
        const delta = qtyCurrent - qtySynced;

        if (delta === 0) return;
        if (delta < 0) {
          logError_(
            'syncShipmentsUaeEgToInventory',
            new Error('Negative delta for shipment row (Qty < Qty Synced).'),
            { shipmentId: shipmentId, sku: sku, qtyCurrent: qtyCurrent, qtySynced: qtySynced }
          );
          return;
        }

        const sheetRow = idx + 2; // actual row number in sheet

        // Line discriminator: SKU can repeat within the same Shipment ID (across boxes/rows)
        const boxId = (idxShipBoxId != null) ? String(row[idxShipBoxId] || '').trim() : '';
        const vKey = (idxShipVariant != null) ? String(row[idxShipVariant] || '').trim() : '';
        // Policy: Box ID first; if Box ID is empty use a persistent per-row Line ID (never sheet row number, since sorting/filtering occurs).
        let lineId = '';
        if (!boxId) {
          if (idxShipLineId == null) {
            logError_('syncShipmentsUaeEgToInventory', new Error('Missing Line ID column in Shipments_UAE_EG (required when Box ID is empty).'), { shipmentId: shipmentId, sku: sku, row: sheetRow });
            return;
          }
          lineId = String(row[idxShipLineId] || '').trim();
          if (!lineId) {
            lineId = 'L-' + Utilities.getUuid();
            row[idxShipLineId] = lineId;
            lineIdsChanged = true;
          }
        }
        const lineDiscriminator = boxId ? ('B' + boxId) : ('L' + lineId);

        // Operation key encodes the pre-sync Qty Synced and the delta, so retries don't duplicate
        const baseKey = `SUEG|${shipmentId}|${lineDiscriminator}|${sku}|${vKey}|S${qtySynced}|D${delta}`;
        const sourceIdOut = `${baseKey}|OUT`;
        const sourceIdIn = `${baseKey}|IN`;

        // Dedupe: prefer Txn ID check (covers legacy sourceId=Shipment ID), fallback to Source ID check
        // Fast dedupe (crash-retry) by Source ID; full Txn-ID dedupe is done after we compute cost/warehouse.
        let outExists = existingSourceIds.has(sourceIdOut);
        let inExists = existingSourceIds.has(sourceIdIn);


        const shipDate = row[idxShipShipDate] || new Date();
        const arrDate = (idxShipArrival != null) ? (row[idxShipArrival] || null) : null;
        const inTxnDate = arrDate || shipDate;

        // Total extras (ship/customs/other) per unit (for landed cost)
        const shipCost = idxShipShipCost != null ? Number(row[idxShipShipCost] || 0) : 0;
        const customs = idxShipCustoms != null ? Number(row[idxShipCustoms] || 0) : 0;
        const other = idxShipOther != null ? Number(row[idxShipOther] || 0) : 0;
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

        if (typeof normalizeWarehouseCode_ === 'function') {
          fromWarehouse = normalizeWarehouseCode_(fromWarehouse);
        }

        if (!fromWarehouse) {
          logError_('syncShipmentsUaeEgToInventory', new Error('Cannot determine UAE warehouse for OUT'), { shipmentId: shipmentId, sku: sku, row: (2 + idx) });
          return; // skip row
        }

        // ===== Inventory info for this SKU+Warehouse =====
        const keyWh = sku + '||' + fromWarehouse;
        const info = uaeByWh[keyWh] || uaeBySku[sku] || {};

        // If Warehouse (UAE) is empty in sheet and inventory has a warehouse → fill it
        if (idxShipWhUae != null && !whCell && info.warehouse) {
          row[idxShipWhUae] = info.warehouse;
          fromWarehouse = info.warehouse;
        }

        const baseCost = info.avgCost || 0; // UAE avg cost
        const extrasPerUnit = qtyCurrent > 0 ? totalExtras / qtyCurrent : 0;
        const landedCost = baseCost + extrasPerUnit;

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

        const productName =
          (info && (info.productName || info.product)) ||
          (idxShipProdName != null ? row[idxShipProdName] : '') ||
          '';

        const variant =
          (info && info.variant) ||
          (idxShipVariant != null ? row[idxShipVariant] : '') ||
          '';


        const toWarehouse = (APP.WAREHOUSES && APP.WAREHOUSES.TAN_GH) ? APP.WAREHOUSES.TAN_GH : 'TAN-GH';

        // Txn-ID based dedupe (covers legacy Source ID scheme and prevents duplicates across code revisions)
        const outTxnId = (typeof _inv_makeTxnId_ === 'function') ? _inv_makeTxnId_({
          type: 'OUT',
          sourceType: 'SHIP_UAE_EG',
          sourceId: sourceIdOut,
          batchCode: batchCode,
          sku: String(sku),
          warehouse: fromWarehouse,
          qtyIn: 0,
          qtyOut: delta,
          unitCostEgp: baseCost,
          currency: 'EGP',
          unitPriceOrig: 0,
          txnDate: shipDate
        }) : '';

        const inTxnId = (typeof _inv_makeTxnId_ === 'function') ? _inv_makeTxnId_({
          type: 'IN',
          sourceType: 'SHIP_UAE_EG',
          sourceId: sourceIdIn,
          batchCode: batchCode,
          sku: String(sku),
          warehouse: toWarehouse,
          qtyIn: delta,
          qtyOut: 0,
          unitCostEgp: landedCost,
          currency: 'EGP',
          unitPriceOrig: 0,
          txnDate: inTxnDate
        }) : '';

        const outLegacyTxnId = (typeof _inv_makeTxnId_ === 'function') ? _inv_makeTxnId_({
          type: 'OUT',
          sourceType: 'SHIP_UAE_EG',
          sourceId: String(shipmentId),
          batchCode: batchCode,
          sku: String(sku),
          warehouse: fromWarehouse,
          qtyIn: 0,
          qtyOut: delta,
          unitCostEgp: baseCost,
          currency: 'EGP',
          unitPriceOrig: 0,
          txnDate: shipDate
        }) : '';

        const inLegacyTxnId = (typeof _inv_makeTxnId_ === 'function') ? _inv_makeTxnId_({
          type: 'IN',
          sourceType: 'SHIP_UAE_EG',
          sourceId: String(shipmentId),
          batchCode: batchCode,
          sku: String(sku),
          warehouse: toWarehouse,
          qtyIn: delta,
          qtyOut: 0,
          unitCostEgp: landedCost,
          currency: 'EGP',
          unitPriceOrig: 0,
          txnDate: inTxnDate
        }) : '';

        if (!outExists && ((outTxnId && existingTxnIds.has(outTxnId)) || (outLegacyTxnId && existingTxnIds.has(outLegacyTxnId)))) {
          outExists = true;
        }
        if (!inExists && ((inTxnId && existingTxnIds.has(inTxnId)) || (inLegacyTxnId && existingTxnIds.has(inLegacyTxnId)))) {
          inExists = true;
        }

        // ===== OUT from UAE warehouse =====
        if (!outExists) {
          txns.push({
            txnId: outTxnId || undefined,
            type: 'OUT',
            sourceType: 'SHIP_UAE_EG',
            sourceId: sourceIdOut,
            batchCode: batchCode,
            sku: String(sku),
            productName: String(productName || ''),
            variant: String(variant || ''),
            warehouse: fromWarehouse,
            qty: delta,
            unitCostEgp: baseCost,
            currency: 'EGP',
            txnDate: shipDate,
            notes: 'UAE→EG OUT (' + fromWarehouse + '), delta=' + delta
          });

          outCount++;
          existingSourceIds.add(sourceIdOut);
          if (outTxnId) existingTxnIds.add(outTxnId);
        }

        // ===== IN to TAN-GH at landed cost =====
        if (!inExists) {
          txns.push({
            txnId: inTxnId || undefined,
            type: 'IN',
            sourceType: 'SHIP_UAE_EG',
            sourceId: sourceIdIn,
            batchCode: batchCode,
            sku: String(sku),
            productName: String(productName || ''),
            variant: String(variant || ''),
            warehouse: toWarehouse,
            qty: delta,
            unitCostEgp: landedCost,
            currency: 'EGP',
            txnDate: inTxnDate,
            notes: 'UAE→EG IN (' + toWarehouse + '), delta=' + delta + ', landedCost=' + landedCost.toFixed(2)
          });

          inCount++;
          existingSourceIds.add(sourceIdIn);
          if (inTxnId) existingTxnIds.add(inTxnId);
        }


        // Update Qty Synced (safe حتى لو الاتنين كانوا موجودين)
        row[idxShipQtySynced] = qtyCurrent;
      });

      // Persist any newly generated Line IDs BEFORE writing ledger rows (crash-safety for idempotency).
      if (lineIdsChanged && colLineId && idxShipLineId != null) {
        shipSh.getRange(2, colLineId, shipData.length, 1).setValues(shipData.map(function (r) {
          return [r[idxShipLineId] || ''];
        }));
      }

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
      rebuildInventoryUAEFromLedger();
      rebuildInventoryEGFromLedger();


      if (typeof safeAlert_ === 'function') {
        safeAlert_(
          'Sync Shipments_UAE_EG → Inventory done.\n' +
          'New OUT txns: ' + outCount + '\n' +
          'New IN txns: ' + inCount
        );
      } else {
        Logger.log('Sync Shipments_UAE_EG → Inventory done. OUT=' + outCount + ', IN=' + inCount);
      }
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
  const info = _getInventoryUaeInfoForSku_('MONSTER-MQT65-PURPLE', 'UAE-DXB');
  Logger.log(JSON.stringify(info));
}
