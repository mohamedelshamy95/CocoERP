# Coco ERP System — Roadmap (The Coco Club)

**Document:** Roadmap.md  
**Owner:** The Coco Club (CocoERP)  
**Primary Stack:** Google Sheets + Google Apps Script (clasp) + Node/Eslint (local)  
**Operating Model:** Strict sequential pipeline + queue runner (time trigger) + minimal onEdit

---

## 1) Vision & North Star

Build a lightweight, reliable ERP that manages **end-to-end commerce operations** for The Coco Club:

- Procurement (China) → inbound shipping to UAE → QC in UAE → UAE inventory
- Outbound shipping UAE → Egypt → Egypt inventory
- Catalog + pricing + unit economics
- Sales + delivery tracking + returns
- Advertising/campaign tracking & profitability attribution (later)
- Future: Shopify API integration for automated sales + inventory sync

**North Star KPIs**
- Accurate inventory (UAE & EG) with auditable ledger
- True landed cost & unit economics per SKU/batch
- Fast, non-blocking automations (no sheet freezing)
- High trust: deterministic, idempotent sync with strong error observability

---

## 2) Core Operating Principle

### 2.1 Sequential Flow (Must be preserved)

The system must execute in a **strict, sequential pipeline** (not parallel), to avoid sheet hanging and to keep ordering correct:

1. **Settings**
2. **Purchases**
3. **Orders**
4. **Shipments_CN_UAE**
5. **QC_UAE**
6. **Inventory_Transactions** (UAE posting from QC)
7. **Inventory_UAE** (snapshot/view rebuild)
8. **Shipments_UAE_EG**
9. **Inventory_Transactions** (EG posting from UAE→EG transfer)
10. **Inventory_EG** (snapshot/view rebuild)
11. **Catalog_EG**
12. **Sales_Inbox_EasyOrder** (raw intake / staging)
13. **Sales_EG** (operational sales + fulfillment + stock-out)
14. (Next) **Returns_EG**
15. (Next) **Ads / Campaigns**

**Orchestration rule:** a single “queue runner” should call modules in this order with locks + time slicing + delta processing.

---

## 3) Warehouses, Couriers, and Logistics Model

### 3.1 UAE Warehouses (and also UAE→EG Couriers)

You have two operational entities in UAE that function as:
- **Warehouses inside UAE**
- **Couriers/forwarders from UAE-DXB to EG-TAN**

**Canonical UAE Warehouses / Couriers**
- `KOR`
- `ATTIA`

> Normalization must map any aliases (e.g., UAE-KOR, Kor, Kor - Attia text, etc.) → canonical `KOR` / `ATTIA`.

### 3.2 Egypt Warehouse (Primary)
- `EG-TAN` (Tanta main office warehouse)

---

## 4) Sales Channels & Order Intake

### 4.1 Platforms (Where demand originates)
- `Facebook`
- `Instagram`
- `TikTok`
- `WhatsApp`

### 4.2 Source (How the order enters Sheets)
- `EasyOrder` (auto from landing page to sheet)
- `Manual-WhatsApp` (confirmed manually)
- `Manual-Messenger` (confirmed manually)

**Rule:** Keep `Platform` separate from `Source` for clean reporting.

---

## 5) Data Model: Sheets as Modules (System of Record)

Below is the authoritative module list and intended responsibility for each sheet.

### 5.1 Settings
**Purpose:** global configuration, enums, defaults, and policies.  
**Headers:**
- Setting | Value | Platforms | Payment Methods | Currencies | Stores (optional) | Warehouses

**Canonical master lists**
- Platforms: Facebook / Instagram / TikTok / WhatsApp (+ aliases: FB, IG, WA)
- Payment Methods: Cash / Transfer / Card (Card uses Last4 when available)
- Currencies: AED / EGP / USD
- Warehouses: UAE: KOR, ATTIA / EG: EG-TAN

---

### 5.2 Purchases (line-level truth)
**Purpose:** procurement lines, costs, FX, landed-cost math, and stable line identity.  
**Headers:**
- Order ID | Order Date | Platform | Seller Name | SKU | Batch Code | Product Name | Variant / Color | Qty | Unit Price (Orig) | Currency | Subtotal (Orig) | Discount (Order) | Shipping Fee (Order) | Total Order (Orig) | Final Unit Price | Buyer Name | Buyer Phone | Buyer Address | Payment Method | Payment Card Last4 | Invoice File ID | Invoice Link | Invoice Preview | FX Rate → EGP | Order Total (EGP) | Ship UAE→EG (EGP) | Customs/Fees % | Customs/Fees (EGP) | Landed Cost (EGP) | Unit Landed Cost (EGP) | Notes | Line Gross (Orig) | Discount Alloc (Orig) | Shipping Alloc (Orig) | Line Net (Orig) | Net Unit Price (Orig) | Net Unit Price (EGP) | Line ID

**Key rule:** `Line ID` must be stable + unique per purchase line (downstream identity anchor).

---

### 5.3 Orders (Aggregated purchase orders)
**Purpose:** order-level rollup derived from Purchases.  
**Headers:**
- Order ID | Order Date | Platform | Seller Name | Currency | Buyer Name | Total Lines | Total Qty | Total Order (Orig) | Order Total (EGP) | Ship UAE→EG (EGP) | Customs/Fees (EGP) | Landed Cost (EGP) | Unit Landed Cost (EGP) | Notes

**Rule:** duplicates → choose canonical row and log duplicates once (audit).

---

### 5.4 Shipments_CN_UAE (Inbound logistics)
**Purpose:** inbound shipments from supplier/factory to UAE; line mapping via Purchases Line ID.  
**Headers:**
- Shipment ID | Supplier / Factory | Forwarder | Tracking / Container | Purchases Line ID | Order ID (Batch) | Ship Date | ETA | Actual Arrival | Status | SKU | Product Name | Variant / Color | Qty | Gross Weight (kg) | Volume (CBM) | Freight (AED) | Other Fees (AED) | Total Cost (AED) | Notes

**Status dictionary (canonical)**
- Planned / In Transit / Delayed / Arrived UAE

**Output:** QC_UAE becomes eligible when status becomes `Arrived UAE` (or arrival date is present).

---

### 5.5 QC_UAE (Quality control in UAE)
**Purpose:** QC records per purchase line once shipment arrives UAE.  
**Headers:**
- QC ID | Order ID | Shipment CN→UAE ID | SKU | Batch Code | Product Name | Variant / Color | Qty Ordered | Qty Received | Qty OK | Qty Missing | Qty Defective | QC Result | QC Date | Warehouse (UAE) | Purchases Line ID | Notes

**Rules**
- `Warehouse (UAE)` is **manual** (must be one of `KOR` / `ATTIA`).
- QC Result is derived:
  - Blank until meaningful input exists
  - PASS when OK == ordered and missing/defect == 0
  - PARTIAL when OK > 0 and (missing or defect) > 0
  - FAIL when OK == 0 and (missing or defect) > 0 (or received==0 with ordered>0)

**Output:** Drives inventory transactions into UAE after QC completion.  
**Hard guard:** If Warehouse is missing → skip posting + rate-limit log (no spam).

---

### 5.6 Inventory_Transactions (Ledger)
**Purpose:** append-only inventory movements.  
**Headers:**
- Txn ID | Txn Date | Source Type | Source ID | Batch Code | SKU | Product Name | Variant / Color | Warehouse | Qty In | Qty Out | Unit Cost (EGP) | Total Cost (EGP) | Currency | Unit Price (Orig) | Notes

**Rule:** This is the single source of truth for inventory valuation.  
Inventory_UAE and Inventory_EG are derived snapshots.

---

### 5.7 Inventory_UAE (Snapshot)
**Purpose:** computed UAE stock state and valuation (derived from ledger).  
**Headers:**
- SKU | Product Name | Variant / Color | Warehouse (UAE) | On Hand Qty | Allocated Qty | Available Qty | Avg Cost (EGP) | Total Cost (EGP) | Last Txn Date | Last Source Type | Last Source ID

Warehouses must be canonical: `KOR`, `ATTIA`.

---

### 5.8 Shipments_UAE_EG (Outbound logistics)
**Purpose:** shipments from UAE to Egypt; also cost allocation driver.  
**Headers:**
- Shipment ID | Forwarder | Courier | AWB / Tracking | Box ID | Ship Date | ETA | Actual Arrival | Status | Warehouse (UAE) | SKU | Product Name | Variant / Color | Qty | Qty Synced | Ship Cost (EGP) – per unit | Customs (EGP) – per unit | Other (EGP) – per unit | Total Cost (EGP) | Line ID | Notes

**Status dictionary (canonical)**
- Planned / In Transit / Delayed / Arrived EG

**Rules**
- `Warehouse (UAE)` must be `KOR` or `ATTIA` (also doubles as courier identity when needed).
- Ledger posting is idempotent by `Line ID` + `Qty Synced` delta.

---

### 5.9 Inventory_EG (Snapshot)
**Purpose:** computed Egypt stock snapshot (derived from ledger).  
**Headers:**
- SKU | Product Name | Variant / Color | Warehouse (EG) | On Hand Qty | Allocated Qty | Available Qty | Avg Cost (EGP) | Total Cost (EGP) | Last Txn Date | Last Source Type | Last Source ID

Primary warehouse: `EG-TAN`.

---

### 5.10 Catalog_EG (Master data & pricing)
**Purpose:** SKU master, classification, default cost/price, status.  
**Headers:**
- SKU | Product Name | Variant / Color | Color Group | Brand | Category | Subcategory | Status | Default Cost (EGP) | Default Price (EGP) | Barcode | Notes

---

### 5.11 Sales_Inbox_EasyOrder (Raw intake / staging)
**Purpose:** raw EasyOrder feed + manual enrichment, kept separate from operational Sales_EG.  
**Headers (as provided):**
- Name | Phone | City | Address | Product Quantity | Product Total Cost | Order Total Cost | Shipping Cost | Product Name | Product Variant | Order ID | Taager ID | Date | SKU | Extra Data | Extra Data 2 | UTM Source | UTM Campaign | Payment Method | Coupon | Discount | Alt Phone | Customer note | Referral Code | Taager Order ID | Confirm Status

**Rules**
- This sheet is **not** the inventory stock-out source directly.
- It is a staging inbox to prevent collisions while orders are being confirmed/edited.
- We promote rows from Sales_Inbox_EasyOrder → Sales_EG only when Confirm Status reaches shipping stage (see 5.12).

---

### 5.12 Sales_EG (Sales operations)
**Purpose:** operational sales orders + fulfillment status + inventory stock-out when shipped.  
**Headers:**
- Order ID | Order Date | Platform | Customer Name | Phone | City | Address | SKU | Product Name | Variant / Color | Warehouse (EG) | Qty | Unit Price (EGP) | Total Price (EGP) | Discount (EGP) | Net Revenue (EGP) | Shipping Fee (EGP) | Payment Method | Order Status | Delivered Date | Source | Courier | AWB | Notes

**Order Status dictionary (canonical Arabic flow)**
- تم تأكيد الأوردر
- تم تجهيز الأوردر
- تم تسليم الاوردر لشركة الشحن
- قيد الشحن
- تم التسليم للعميل
- قيد الاسترجاع
- تم الاسترجاع
- مفقود
- قيد التعويض
- تم التعويض

**Promotion + Stock-out rule (critical)**
- Orders are promoted into Sales_EG (or updated there) when:
  - Confirm Status (in Sales_Inbox_EasyOrder) becomes **"تم تسليم الاوردر لشركة الشحن"** or **"قيد الشحن"** (or later states).
- Inventory stock-out from EG happens when Sales_EG status enters shipped states (same two above).
- Returns are handled later via Returns_EG (planned) or extended logic.

---

### 5.13 ErrorLog (Observability)
**Purpose:** structured logging for failures and important anomalies.  
**Headers:**
- Timestamp | Function | Message | Stack | Context

**Rule:** rate-limit noisy permanent-data issues (e.g., missing warehouse) to prevent spam.

---

## 6) Automation & Orchestration Roadmap

### 6.1 Queue Runner (Core)
- Single orchestrator executes modules sequentially.
- Uses `LockService` for concurrency safety.
- Uses delta processing (new/changed lines only).
- Cursor/state stored in DocumentProperties.

### 6.2 Triggers (Current)
- `coco_processSyncQueue` runs every **1 minute**
- `coco_onEditInstallable` is installed for minimal safe enqueueing only

### 6.3 Performance Guardrails
- Batch reads/writes only (avoid per-cell operations)
- Avoid whole-sheet scans where possible
- Stable keys:
  - Purchases: `Line ID`
  - Shipments_UAE_EG: `Line ID`
  - Ledger: `Txn ID`

---

## 7) Integrations Roadmap

### Phase A — Current
- EasyOrder → Sales_Inbox_EasyOrder (auto)
- Manual confirmation workflows supported

### Phase B — Shopify (Planned)
- Pull orders, customers, products, inventory adjustments via API
- Map Shopify variants to `SKU + Variant/Color`
- Two-way sync:
  - Inventory EG → Shopify
  - Shopify orders/fulfillments → Sales_EG

### Phase C — Couriers (Optional)
- AWB tracking auto updates in Sales_EG
- Delivery status can drive returns automation

---

## 8) Analytics & Reporting Roadmap
- Landed cost by SKU/batch (CN→UAE + UAE→EG + customs/fees)
- Profit per order / per SKU / per channel
- Stock aging + dead stock flags
- Ads spend + ROAS (manual first, then integrations)

---

## 9) Milestones & Phases

### Phase 0 — Baseline Stabilization (Done / In Progress)
- Core modules exist with headers aligned
- QC_UAE: manual warehouse, QC Date from arrival, correct QC Result
- Error log hygiene improvements
- Ensure queue runner strictly respects sequential order end-to-end

### Phase 1 — Operational MVP (Next)
- Enforce canonical warehouses: `KOR`, `ATTIA`, `EG-TAN`
- Add/verify Sales_Inbox_EasyOrder → Sales_EG promotion logic
- Expand guardrails: rate-limit noisy logs + summarize anomalies
- Preflight checker (missing headers/sheets + actionable repair)

### Phase 2 — Sales & Fulfillment Hardening
- Delivery workflow standardization + courier/AWB normalization
- Returns module (Returns_EG) + ledger integration

### Phase 3 — Catalog & Pricing Governance
- Catalog validation (SKU uniqueness, status)
- Default pricing policy & barcode support

### Phase 4 — Ads & Attribution
- Campaigns + spend + ROAS dashboards

### Phase 5 — Shopify Integration
- Shopify orders ingestion + EG inventory sync back

---

## 10) Glossary (Canonical Terms)
- SKU: unique sellable item identifier
- Variant/Color: matching dimension used for matching
- Batch Code: procurement batch identifier
- Line ID: stable line-level identifier propagated downstream
- Warehouse (UAE): `KOR` or `ATTIA`
- Warehouse (EG): `EG-TAN`
- Platform: Facebook/Instagram/TikTok/WhatsApp
- Source: EasyOrder / Manual-WhatsApp / Manual-Messenger
- Ledger: Inventory_Transactions (source of truth)

---

**End of Roadmap.md**
