# Patchset-00: Shipments_UAE_EG Customs/Other Per-Unit Migration

**What changed**
- Shipments_UAE_EG cost semantics are per-unit: Ship Cost + Customs + Other are per-unit inputs; Total Cost = Qty * (ship + customs + other).

**One-time migration for legacy rows that stored totals**
1) Filter Shipments_UAE_EG to rows where Qty > 0.
2) For each row that previously held totals:
   - Customs per unit = (legacy Customs total) / Qty
   - Other per unit   = (legacy Other total) / Qty
   - If Qty = 0, skip the row and update Qty first; do not divide by zero.
3) Recompute Total Cost = Qty * (Ship Cost per unit + Customs per unit + Other per unit).
4) Rerun `updateShipmentsUaeEgStatusAndTotals()` to normalize totals/formats, then rerun `syncShipmentsUaeEgToInventory()` to align ledger costs.

**Verification**
- Pick a sample row with Qty > 1 and confirm Total Cost equals Qty * (ship + customs + other).
- After inventory sync, check Inventory_Transactions entries for that shipment: Unit Cost = base cost + extrasPerUnit (ship + customs + other), no divide-by-qty adjustments.
