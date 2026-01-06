# CocoERP — RUNBOOK.md (v1.1)

**Scope:** How we work, ship, debug, and collaborate (VS Code + Git + clasp + Copilot + Codex).  
**Primary rule:** Protect the sequential pipeline and keep changes small, testable, and reversible.

---

## 1) Tooling Baseline

### 1.1 Editor / IDE Choice
- We use **VS Code** as the primary workspace (fast, extensible, strong Git + JS tooling).
- A full IDE (e.g., WebStorm) typically adds deeper refactor/navigation features out-of-the-box.
- For this project, **VS Code is sufficient and recommended** because:
  - repo is Apps Script + Node lint tooling
  - workflows depend more on Git discipline and deployment (clasp) than IDE complexity

### 1.2 Required Local Tools
- Node.js (LTS)
- npm
- Git
- clasp (Google Apps Script CLI)
- ESLint configured in repo

### 1.3 Standard Commands
- Syntax check:
  - `npm run check:syntax`
- Lint:
  - `npm run lint`
- Fix lint (when safe):
  - `npm run lint:fix`
- Apps Script deployment:
  - `clasp pull`
  - `clasp push`

---

## 2) Operational Architecture (Non-Negotiables)

### 2.1 Strict Sequential Pipeline
Stages run in a strict order (one stage per minute run or a small batch):

1) Settings  
2) Purchases  
3) Orders  
4) Shipments_CN_UAE  
5) QC_UAE  
6) Inventory_Transactions (UAE)  
7) Inventory_UAE  
8) Shipments_UAE_EG  
9) Inventory_Transactions (EG)  
10) Inventory_EG  
11) Catalog_EG  
12) Sales_Inbox_EasyOrder  
13) Sales_EG  

**Cost entry rule (Stage 8):** Shipments_UAE_EG uses per-unit inputs for Ship Cost, Customs, and Other; Total Cost must equal `Qty * (ship + customs + other)`.

### 2.2 Trigger Model
- `coco_processSyncQueue` runs **every 1 minute**
- `coco_onEditInstallable` sets flags only (must remain minimal)

### 2.3 Ledger-First Accounting
- `Inventory_Transactions` is the source of truth.
- Inventory snapshots are derived views.

---

## 3) Collaboration Model: Copilot vs Codex (Two Complementary Roles)

### 3.1 Why split roles?
If Copilot and Codex both edit the same files concurrently, we get:
- conflicting diffs
- duplicated logic
- unstable behavior in triggers

We avoid this by assigning **non-overlapping ownership per task**.

---

## 4) Role Definitions (Copilot vs Codex)

### 4.1 Copilot Role (Local, surgical edits)
**Best at**
- Small local refactors inside the file you are editing
- Quick iteration on a single function
- Formatting, guards, small bug fixes
- Assisting with VS Code inline suggestions + short diffs

**Copilot constraints**
- Avoid large multi-file changes
- Avoid sweeping edits to core constants unless explicitly planned

**Copilot “ownership” examples**
- Editing a single module file for a targeted bug fix:
  - e.g., one function in ShipmentsCore.js
- Adding guard clauses, rate-limit wrappers, small helpers

---

### 4.2 Codex Role (Repo-wide, multi-file changes + patch artifacts)
**Best at**
- Reading the repo holistically
- Multi-file edits that preserve contracts
- Generating patchsets + audit summaries
- Ensuring referenced functions exist + restoring missing handlers

**Codex constraints**
- Must follow the contract pack and pipeline invariants
- Should produce changes as a coherent patchset (reversible)

**Codex “ownership” examples**
- Implementing a new pipeline stage end-to-end:
  - Sales_Inbox_EasyOrder → Sales_EG promotion
- Restoring missing functions referenced by AppCore
- Large refactors that require coordinated edits across modules

---

## 5) “No Collision” File Ownership Rule (Per Task)

For any task, declare file ownership explicitly:

- **Codex-owned files** (multi-file/system changes):
  - AppCore.js (orchestrator, triggers, constants)
  - any new shared helpers / contract enforcement
- **Copilot-owned files** (localized changes):
  - one module file at a time (e.g., Sales.js OR ShipmentsCore.js)

**Hard rule:** Only one agent edits AppCore.js in a given patchset.

---

## 6) Standard Workflow per Change (Patchset Discipline)

### 6.1 Branching
1) `git checkout -b patchset-XX-short-title`
2) Keep scope small and reversible.

### 6.2 Define Acceptance Criteria (before coding)
Every change must have:
- a one-paragraph goal statement
- a checklist of acceptance criteria
- explicit success signals in Sheets (what should appear / stop appearing)

### 6.3 Implement
- If multi-file: Codex leads; Copilot supports local edits only.
- If single-file: Copilot leads; Codex reviews and checks invariants.

### 6.4 Local Verification (mandatory)
Run:
- `npm run check:syntax`
- `npm run lint`

Lint warnings are acceptable; lint errors are not.

### 6.5 Deploy to Apps Script
- `clasp push`
- Run a manual safe entrypoint (non-UI from trigger context), e.g.:
  - `coco_processSyncQueueNow()` (preferred if present)
  - or run `coco_processSyncQueue()` manually once for validation

### 6.6 Commit
- Include patchset label and intent:
  - `git commit -m "Patchset-XX: <intent>"`
- Push branch:
  - `git push -u origin patchset-XX-short-title`

---

## 7) Sales Intake Rule (Staging Sheet)

### 7.1 Why Sales_Inbox_EasyOrder exists
EasyOrder rows are edited while orders are being confirmed. If we write directly into Sales_EG too early, we risk:
- duplicates
- mismatched status
- inventory posting too early

So we use:
- `Sales_Inbox_EasyOrder` = raw intake + manual enrichment
- `Sales_EG` = operational truth

### 7.2 Promotion Rule (Critical)
Promote from `Sales_Inbox_EasyOrder` to `Sales_EG` when Confirm Status becomes:
- "تم تسليم الاوردر لشركة الشحن"
- "قيد الشحن"
(or later states)

### 7.3 Inventory Posting Rule (Critical)
Inventory stock-out from `EG-TAN` happens when Sales_EG enters shipped states:
- "تم تسليم الاوردر لشركة الشحن"
- "قيد الشحن"

Returns are handled later via Returns_EG.

---

## 8) Debugging & Operations

### 8.1 First-line checks
- Verify triggers exist and are enabled:
  - installable onEdit: `coco_onEditInstallable`
  - time trigger: `coco_processSyncQueue` every 1 minute
- Check ErrorLog:
  - look for repeating issues (should be rate-limited)
- Confirm the pipeline is not stuck on a stage due to missing required fields:
  - e.g., missing Warehouse (UAE) in QC_UAE blocks posting

### 8.2 Common failure modes
- Missing sheet/headers:
  - must be surfaced via preflight and rate-limited
- Spam logging:
  - rate-limit by error signature + time window
- Duplicate writes:
  - enforce idempotency keys:
    - Purchases: Line ID
    - Shipments_UAE_EG: Line ID + Qty Synced delta
    - Ledger: Txn ID deterministic

---

## 9) Copilot Chat / Codex Prompting Standards (Inside VS Code)

### 9.1 Use repo context intentionally
When using Copilot Chat, prefer:
- `@workspace` for repo-wide questions
- `@file` when you want edits to a specific file

### 9.2 Model selection guidance (Copilot)
- For coding multi-file patches: choose a Codex-optimized model (e.g., GPT-5.1-Codex / Codex-Max).
- For architecture review and contracts: choose a reasoning-optimized model (e.g., GPT-5.2).

### 9.3 Prompt template (patch request)
Include:
- Goal
- Constraints (sequential pipeline, idempotency, no UI calls in triggers)
- Files allowed to edit (ownership rule)
- Acceptance criteria checklist
- Test steps (npm run check:syntax, npm run lint, clasp push, manual run)

---

## 10) Change Safety Checklists

### 10.1 Pre-change checklist
- Pull latest: `git pull`
- Ensure clean working tree: `git status`
- Confirm triggers in spreadsheet
- Confirm required sheets exist
- Export/record DocumentProperties keys if you will touch queues

### 10.2 Post-change checklist
- `npm run check:syntax`
- `npm run lint`
- `clasp push`
- Manual run queue once
- Confirm:
  - no new spam in ErrorLog
  - ledger idempotency preserved
  - snapshots rebuild correctly

---

**End of RUNBOOK.md v1.1**
