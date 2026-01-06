# CocoERP — System Contract Pack (v1.0)

This document is the binding contract for workbook schema, keys, statuses, triggers, and sync rules.

---

## A) Workbook & Apps Script

### A.1 Spreadsheet
- Name: CocoERP v2
- Timezone/Locale: Africa/Cairo

### A.2 Tabs (System of Record)
1) Settings  
2) Purchases  
3) Orders  
4) Shipments_CN_UAE  
5) QC_UAE  
6) Inventory_Transactions  
7) Inventory_UAE  
8) Shipments_UAE_EG  
9) Inventory_EG  
10) Catalog_EG  
11) Sales_Inbox_EasyOrder  
12) Sales_EG  
13) ErrorLog  

### A.3 Triggers
- Installable onEdit: `coco_onEditInstallable`
- Time-driven: `coco_processSyncQueue` every 1 minute

---

## B) Master Data (Enums)

### B.1 Platforms (canonical + accepted aliases)
- Facebook (FB)
- Instagram (IG)
- TikTok
- WhatsApp (WA)

### B.2 Payment Methods
- Cash
- Transfer
- Card (supports Last4 when available)

### B.3 Currencies
- AED
- EGP
- USD

### B.4 Warehouses (canonical)
- UAE: KOR, ATTIA
- EG: EG-TAN

**Important semantic rule:**  
KOR and ATTIA are both:
- Warehouses in UAE
- Courier/Forwarder identity for UAE→EG (UAE-DXB → EG-TAN)

---

## C) Status Dictionaries

### C.1 Shipments_CN_UAE.Status (canonical)
- Planned
- In Transit
- Delayed
- Arrived UAE

### C.2 Shipments_UAE_EG.Status (canonical)
- Planned
- In Transit
- Delayed
- Arrived EG

### C.3 Sales_EG.Order Status (canonical operational Arabic)
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

**Synonyms handling:** allowed, but must be normalized to canonical values for automation rules.

---

## D) Sheet Contracts (Headers)

### D.1 Settings
- Setting | Value | Platforms | Payment Methods | Currencies | Stores (optional) | Warehouses

### D.2 Purchases
(As provided; Line ID is mandatory and stable)

### D.3 Orders
(As provided)

### D.4 Shipments_CN_UAE
(As provided; Purchases Line ID required)

### D.5 QC_UAE
(As provided; Warehouse (UAE) manual; Purchases Line ID required)

### D.6 Inventory_Transactions
(As provided; Txn ID deterministic; append-only behavior preferred)

### D.7 Inventory_UAE / Inventory_EG
(As provided; derived snapshots)

### D.8 Shipments_UAE_EG
(As provided; Line ID required; Qty Synced supports delta posting)
- Cost semantics (canonical): `Ship Cost (EGP) – per unit`, `Customs (EGP) – per unit`, and `Other (EGP) – per unit` are all **per-unit** inputs.
- Total Cost formula: `Total Cost (EGP) = Qty * (Ship Cost per unit + Customs per unit + Other per unit)`.

### D.9 Catalog_EG
(As provided)

### D.10 Sales_Inbox_EasyOrder (staging)
- Name | Phone | City | Address | Product Quantity | Product Total Cost | Order Total Cost | Shipping Cost | Product Name | Product Variant | Order ID | Taager ID | Date | SKU | Extra Data | Extra Data 2 | UTM Source | UTM Campaign | Payment Method | Coupon | Discount | Alt Phone | Customer note | Referral Code | Taager Order ID | Confirm Status

### D.11 Sales_EG (operational)
(As provided)

### D.12 ErrorLog
- Timestamp | Function | Message | Stack | Context

---

## E) Keys & Idempotency Rules (Critical)

### E.1 Purchases
- Primary identity: `Line ID` (stable unique)

### E.2 Shipments_CN_UAE
- Primary identity: Purchases Line ID

### E.3 QC_UAE
- Primary identity: Purchases Line ID

### E.4 Shipments_UAE_EG
- Primary identity: `Line ID`
- Delta control: `Qty Synced`

### E.5 Inventory_Transactions
- Primary identity: `Txn ID` deterministic
- Must not duplicate Txn IDs on reruns

---

## F) Promotion and Posting Rules

### F.1 EasyOrder → Sales_Inbox_EasyOrder
- Data enters automatically (current) OR via API later.
- Sheet can be enriched manually (SKU mapping, notes, etc.).

### F.2 Sales_Inbox_EasyOrder → Sales_EG (Promotion rule)
Promote/update Sales_EG line(s) when Confirm Status becomes:
- "تم تسليم الاوردر لشركة الشحن"
- "قيد الشحن"
(or later statuses)

### F.3 Sales_EG → Inventory posting (Stock-out)
Stock-out from EG warehouse happens when Sales_EG enters shipped states:
- "تم تسليم الاوردر لشركة الشحن"
- "قيد الشحن"

Returns are handled later by Returns_EG (planned) or equivalent extension.

---

## G) Webhooks vs Pull Jobs (Conceptual)

### G.1 Webhooks
A webhook means the source system pushes updates to us immediately (event-driven).
- Pros: near real-time, less polling.
- Cons: requires a public endpoint; Apps Script web app/security must be handled carefully.

### G.2 Pull Jobs
A pull job means we periodically fetch new/changed records from the source (time-driven).
- Pros: simpler ops; matches our existing 1-minute queue runner model.
- Cons: not real-time; must handle paging/delta markers.

**Practical guidance for CocoERP**
- Start with Pull Jobs (fits queue runner + simplicity).
- Add Webhooks later if EasyOrder/Shopify supports it cleanly and we need real-time.

---

## H) Operational Constraints (Never violate)
- No UI calls inside triggers (no SpreadsheetApp.getUi()).
- Batch reads/writes only.
- One stage per run (or bounded batch) to avoid timeouts.
- Rate-limit ErrorLog noise.
- Maintain strict stage ordering.

---

**End of SYSTEM_CONTRACT_PACK.md v1.0**
