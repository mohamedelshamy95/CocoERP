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
      map['Ship Cost (EGP) – per unit'] ||
      map['Ship Cost (EGP) - per unit'] ||
      map['Ship Cost (EGP) – per unit or box'] ||
      map['Ship Cost (EGP) - per unit or box'] ||
      map['Ship Cost (EGP)'];

    const colCustoms = map[APP.COLS.SHIP_UAE_EG.CUSTOMS] ||
      map['Customs (EGP) – per unit'] ||
      map['Customs (EGP) - per unit'] ||
      map['Customs (EGP)'] ||
      map['Customs / Clearance (EGP)'];

    const colOther = map[APP.COLS.SHIP_UAE_EG.OTHER] ||
      map['Other (EGP) – per unit'] ||
      map['Other (EGP) - per unit'] ||
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

      const extrasPerUnit = shipCostPerUnit + customs + other;
      const totalForShipment = qty ? (qty * extrasPerUnit) : 0;
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
      rowObj['Purchases Line ID'] = lineId;

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

    if (!shipIdCol || !shipWhCol || !shipSkuCol) {
      throw new Error('Missing required Shipments_UAE_EG columns (Shipment ID / Warehouse (UAE) / SKU). Run Logistics → Setup Shipments Layouts.');
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
      rowObj['Purchases Line ID'] = lineId;

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

    if (!shipIdCol || !shipWhCol || !shipSkuCol) {
      throw new Error('Missing required Shipments_UAE_EG columns (Shipment ID / Warehouse (UAE) / SKU). Run Logistics → Setup Shipments Layouts.');
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
      rowObj['Purchases Line ID'] = lineId;

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