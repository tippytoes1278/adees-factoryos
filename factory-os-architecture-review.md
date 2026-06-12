# Factory OS — Full Architecture & Code Review

**Date:** 12 June 2026
**Scope:** Read-only audit of `/Users/Ayush/adees-factoryos/` — `Server.js` (2,087 lines), `Index.html` (2,318 lines), `appsscript.json`. No code was modified.

---

## 1. Executive Summary

Factory OS is a Google Apps Script web app over a Google Sheets database that digitises piece-rate contractor management for a footwear export factory: order intake (TS → BOM → ART sheets), per-department activity setup with approval gates, weekly/period quantity entry, payment submission and approval, and WIP reconciliation. The approval-request pattern (everything flows through the `REQUESTS` sheet and lands on Ayush's Approve screen) is the system's strongest idea and it is applied consistently — orders, activity setups, rate changes, entry edits, payments, and (recently) setup edits all go through it.

**What's genuinely good:**
- The domain model matches the factory. Department sequence (Cutting → Preparation → Fitter → Lasting → Finishing → Dispatch), per-article activity setup, lot-size caps, skip semantics, and sequence-based edit locking are real business rules encoded correctly, not generic ERP abstractions.
- The append-only event direction is improving: the new `ACTIVITY_SETUP_RESET` mechanism reverts state by *appending* a row rather than mutating history — exactly the right pattern.
- The UI is honest mobile-first vanilla JS with no framework debt, and role-based navigation keeps each user's screen simple.

**What's holding it back:**
1. **The two-layer event-log design exists on paper but not fully in code.** `AUDIT_LOG` and `PERIOD_ENTRIES` — described as part of the schema — are never referenced anywhere in `Server.js`. The ART sheet grids are mutable working state, and `clearWeekForNext` destructively wipes them with no event trail.
2. **No concurrency control at all.** There is not a single `LockService` call in the codebase. Request IDs, BOM numbers, and ART sheet numbers are all derived from row counts or sheet-name maxima — two simultaneous users can collide.
3. **Server functions are unguarded.** `google.script.run` exposes *every* global function to *any* logged-in Google user (web app access is `ANYONE`). `WIPE_AND_RESET` — which deletes every ART sheet — has no role check. Neither do `saveWIP`, `clearWeekForNext`, or `submitRequest`.
4. **Client-supplied financial data is trusted.** `addContinuationRow` accepts rate and commission from the browser and writes them to the sheet. Rates should only ever come from the server-side approved setup.
5. **Two parallel order-creation code paths** (legacy pipe-delimited `createNewArtSheet` at ~line 1158, JSON `createOrder` at ~line 1991) and **two divergent `PAYMENT_HISTORY` header schemas** mean the same data means different things depending on which code path wrote it.

**Bottom line:** the product design is ahead of the engineering. The system can run the factory today because one trusted operator (Ayush) approves everything, but it cannot yet survive the milestone bar — *one complete payment cycle where nobody asks Ayush what to do* — because the data-integrity and authorisation gaps currently rely on Ayush being the integrity layer. Fixing §4's top five items is cheaper than any feature and is a precondition for the Labour Rate Card work.

---

## 2. Architecture & Data Model Review

### 2.1 Overall shape

```
Browser (Index.html, single file, vanilla JS)
   │  google.script.run RPC
   ▼
Server.js (Apps Script, V8)
   │  SpreadsheetApp
   ▼
Google Sheet (DEV or LIVE, selected by CONFIG.ENV)
   ├─ ART-XXX sheets        ← per-order working grids (rows 5–49)
   ├─ REQUESTS              ← approval event log (the real source of truth for state changes)
   ├─ MASTER_ACTIVITIES, MASTER_CONTRACTORS, MASTER_RATES, TS_MASTER
   ├─ ORDER_INDEX, ORDER_TRACKER, WIP_RECONCILIATION, DEPT_STATUS
   └─ PAYMENT_PERIODS, PAYMENT_HISTORY, PAYMENT_LOG, WEEKLY PAYMENT MASTER, CONFIG
```

The client boots via a single `getAllData`-style bundle, caches per-tab in the global `D` object, and renders everything imperatively with the `el()` helper. The server reads whole sheets with `getDataRange().getValues()` and filters in memory.

### 2.2 Flow consistency — order → activity setup → entry → payment

The happy path is coherent:

1. **Order**: `NEW_ORDER` request (JSON payload from `renderNO_Step2`) → approval → `createOrder` builds the ART sheet, registers in `ORDER_INDEX`/`ORDER_TRACKER`.
2. **Setup**: `ACTIVITY_SETUP` request per dept → approval → activities land in the ART grid with rates; `getDeptStatus` derives PENDING/APPROVED/SKIPPED/NOT_SET by replaying `REQUESTS` rows in order.
3. **Entry**: accounts saves contractor/qty rows (`saveEntry`, `addContinuationRow`) into the ART grid (col L: DRAFT → SUBMITTED), gated by dept status and lot caps.
4. **Payment**: `PAYMENT_SUBMISSION` request → approval → rows marked APPROVED, written to `PAYMENT_HISTORY`/`PAYMENT_LOG`.

**Inconsistencies found:**

- **Dual order paths.** Legacy `NEW ORDER` (pipe-delimited, parsed by regex in `createNewArtSheet` ~line 1158) coexists with `NEW_ORDER` (JSON, `createOrder` ~line 1991). Both are still reachable from `processRequest`. The legacy path should be deleted once no legacy rows remain pending.
- **PAYMENT_HISTORY schema fork.** `approvePaymentSubmission` creates the sheet with a `WeekEnding`-style header set; `approvePeriodPayment` creates it with `PeriodID`-style headers. Whichever runs first on a fresh sheet wins, and the other then writes misaligned columns. This is a live data-corruption risk, not just untidiness.
- **Skip is forever, edit is not.** The skip modal (Index.html:1847) warns depts are "permanently skipped … cannot be filled in any future week," and `getDeptStatus` honours that (a SKIPPED dept ignores `ACTIVITY_SETUP_RESET`). But the new setup-edit flow can reset an APPROVED dept to NOT_SET. The asymmetry is defensible but undocumented — and there is no admin path to un-skip a dept skipped by mistake.
- **`markDeptSkipped` failure is swallowed** (Index.html:1885): `withFailureHandler(next)` advances the loop on error, so a dept can fail to be marked skipped while the client proceeds as if it were.
- **Multi-dept setup submission is non-atomic** (Index.html:2145): `renderArticleLevelForm` fires N parallel `requestActivitySetup` calls; partial failure leaves some depts submitted and others not, with only the first error surfaced.

### 2.3 Two-layer event-log design — adherence

The stated design is: **Layer 1** = append-only request/event log (`REQUESTS`, plus `AUDIT_LOG`); **Layer 2** = derived working state (ART grids, `DEPT_STATUS`, trackers). Audit findings:

| Design element | Status in code |
|---|---|
| `REQUESTS` as append-only approval log | ✅ Followed well. State transitions (incl. new `ACTIVITY_SETUP_RESET`) are appended, originals untouched. |
| `AUDIT_LOG` sheet | ❌ **Never referenced anywhere in Server.js.** Grep confirms zero occurrences. |
| `PERIOD_ENTRIES` sheet | ❌ **Never referenced anywhere in Server.js.** Period entry data lives only in mutable ART grid cells (col K = PeriodId). |
| ART grids as derived/rebuildable state | ❌ ART grids are the *primary* store of entry data. They cannot be rebuilt from any log. |
| Destructive ops leave an event trail | ❌ `clearWeekForNext` (~line 1295) wipes D/H/J across rows 5–49 of every ART sheet with no log entry, no snapshot, no role check. `WIPE_AND_RESET` (~line 1610) likewise. |
| Payment trail | ⚠️ Partial. `PAYMENT_HISTORY` + `PAYMENT_LOG` capture approved payments, but `PAYMENT_LOG` coverage is not uniform across both approval paths. |

**Verdict:** the architecture document describes the right system; the code implements roughly half of it. The most consequential gap: **entry quantities have no immutable record.** Once `clearWeekForNext` runs, the only surviving trace of a week's work is whatever made it into `PAYMENT_HISTORY`. A dispute ("I cut 300 pairs, you paid 250") cannot be adjudicated from the system. This is the single biggest design debt and it directly blocks the milestone bar.

### 2.4 Schema coherence

- **Dept vocabulary is duplicated at least six times.** `DEPT_SEQ`/`DEPT_KEY` (and the `prep`/`finish` short-key mapping) appear in Server.js and in Index.html at lines ~1368, ~1773, ~1863, ~2017, plus the contractors/activities tabs' `DEPT_ORDER`. The `'preparation'→'prep'`, `'finishing'→'finish'` normalisation is re-implemented inline in *two* copy-from-article handlers (Index.html:1943, 2081). One drifted copy = silent mis-bucketing of activities.
- **Cross-sheet formulas are load-bearing.** `ORDER_TRACKER` and `WIP_RECONCILIATION` rows hold formulas referencing `'ART-XXX'!Q2`. Rename or delete an ART sheet and the trackers break with `#REF!`. `fixFormulas` exists as a manual repair tool — a symptom, not a fix.
- **The ART template hardcodes activity names.** The M4 summary formula references "Upper Making", "Lasting", "Finish" literally. Any article whose activity names differ gets a wrong/blank summary cell.
- **Rate truth is split three ways**: `MASTER_ACTIVITIES` (std rate), `MASTER_RATES` (edited via stored `rowIndex` — fragile against row insertion/deletion), and the per-article ART grid col E (the rate actually paid). There is no single answer to "what is the rate for activity X on article Y today?" — this is precisely the gap the Labour Rate Card must close.
- **Hardcoded business data in the UI**: buyer list (`Snitch`, `Tata/Westside`, `Van Heusen`, `Mothercare`, `Other` — Index.html:1429), seasons (`SS26`–`AW27`, line 1366), size run fixed at UK 6–11 (lines 1438–1443; kids/Mothercare orders will not fit). All three belong in master sheets.

### 2.5 Supabase/Postgres migration readiness

Honest answer: **the data model is closer to migration-ready than the code is, and neither is close.**

What maps cleanly:
- `REQUESTS` → an `events`/`requests` table (append-only, JSON payload column → `jsonb`). The newer JSON-payload request types port trivially; the legacy pipe-delimited rows need a one-time parse-and-backfill.
- Master sheets → straightforward lookup tables (`contractors`, `activities`, `rates`, `technical_specs`, `orders`).
- The dept-status state machine (`getDeptStatus`) → either a replayed view or a materialised `dept_status` table updated by trigger.

What does not map:
- **ART grids are position-semantic.** Meaning is encoded in *where* a cell is (rows 5–49, col L = status, col K = period) and in formulas (col G, col I compute totals). Postgres needs this normalised into an `entries` table: `(order_id, activity_id, contractor_id, period_id, qty, rate, comm, conveyance, status, …)` — which is exactly the `PERIOD_ENTRIES` design that was specified and never built. **Building `PERIOD_ENTRIES` now, in Sheets, is the cheapest possible migration prep**: it forces the normalisation while the system is small.
- Spreadsheet formulas (commission totals, tracker rollups) become computed columns or queries — fine, but every formula must first be inventoried as business logic.
- `Session.getActiveUser().getEmail()` + hardcoded `ROLES` map → Supabase Auth + a `users`/`roles` table. The role model is small enough that this is easy, but every server function needs an explicit auth check *anyway* (see §4), so do that work now in Apps Script — it transfers directly to RLS policies later.
- Sequential IDs from row counts → Postgres sequences/identity columns solve the race conditions for free. Until then, `LockService` is the stopgap.

**Recommended migration posture:** don't migrate yet. Instead, make the Sheets system *log-shaped* (build `PERIOD_ENTRIES` + an `AUDIT_LOG`, stop destructive wipes). Once every fact is a row in an append-only sheet, migration becomes an export, not a rewrite.

---

## 3. Code Quality & Technical Debt (prioritised)

### P0 — will cause data loss or corruption

1. **No `LockService` anywhere.** Concurrent submits can produce duplicate `REQ-xxx` IDs (`'REQ-' + pad(lastRow - 3)`), duplicate BOM numbers (row-count derived), duplicate ART numbers (max-of-sheet-names derived), and interleaved writes to the same ART rows. Two users (Arvind + Ayush) is already enough to trigger this. Every read-modify-write server function needs `LockService.getScriptLock()`.
2. **`PAYMENT_HISTORY` header fork** (`approvePaymentSubmission` vs `approvePeriodPayment` — see §2.2). Unify the schema and add a header-validation guard before append.
3. **`clearWeekForNext` is an un-logged destructive wipe** of D/H/J across all ART sheets. Minimum fix: snapshot the cleared values to an archive sheet first; better fix: stop clearing and key everything on PeriodId (col K already exists for this).

### P1 — wrong results under normal use

4. **Client-supplied rates in `addContinuationRow`** (Index.html:1180, 1258 pass `dr`/`dc` from the DOM). Server must look up the approved rate from the ART grid/setup and ignore the client value.
5. **`MASTER_RATES` edits by stored `rowIndex`** (Server.js ~855, 1837–1838). Any row insertion between request and approval writes the new rate onto the wrong activity. Match by activity name (or a stable ID) at approval time.
6. **Pervasive empty `catch(e){}`** throughout Server.js. Failures vanish silently — the worst possible behaviour in a financial system. Every catch should at minimum log to a sheet/Stackdriver and return `{success:false, error}`.
7. **Legacy pipe-parsing order path** (`createNewArtSheet`) still live alongside `createOrder`. Delete after confirming no pending legacy rows.

### P2 — maintainability debt that slows every future feature

8. **`getEntryData` scans every ART sheet on every load** (~lines 343–590). O(orders × 45 rows) full-range reads; with 50+ orders this will crawl into Apps Script's 6-minute ceiling and burn read quota. The eventual fix is the `PERIOD_ENTRIES` normalisation; an interim fix is caching with `CacheService`.
9. **Dept vocabulary duplication** (≥6 copies — §2.4). One injected constant from server to client (`getAllData` already exists as the vehicle) eliminates the class of bug.
10. **`renderEntry` is a ~600-line function** with 4-level IIFE closure pyramids (e.g., Index.html:1173, 1249, 1291 are three near-identical inline save/edit handlers). The trio `saveRow` / `addContinuationRow` success-handler / `buildContRow` success-handler is copy-paste with drift already visible. Extract one `makeSavedRowState()` helper.
11. **Hardcoded buyers/seasons/sizes in the UI** (§2.4) — every new buyer is a code deploy.
12. **No tests, no CI, manual `clasp push`/`deploy`, and ENV switching by editing a source constant.** `CONFIG.ENV` is currently `'DEV'` in committed code — deploying this commit to the LIVE deployment ID would silently point production users at the dev sheet. At minimum, derive ENV from `ScriptApp` deployment ID or a Script Property, not source code.
13. **`Index.html` is a single 2,318-line file** mixing CSS, markup, and JS. Apps Script supports multiple HTML files with `include()` — splitting by tab would help, though this is the *least* urgent item on this list.

### What's good (keep doing this)

- `processRequest`'s new double-processing guard (`row[5]` APPROVED/REJECTED → reject) is the right idempotency primitive.
- The setup-edit flow (sequence-based payment lock, `deptHasPayment`, append-style `ACTIVITY_SETUP_RESET`) is the most architecturally sound feature in the codebase.
- `getDeptStatus` as a replay-the-log derivation is exactly the pattern everything else should follow.
- Consistent `{success, error}` RPC envelope and consistent optimistic-UI save patterns on the client.

---

## 4. Security & Data Integrity Findings

Threat model context: web app is `ANYONE` access, `USER_DEPLOYING` execution (appsscript.json). "ANYONE" means **any Google account on the internet** that obtains the URL can load the app and — critically — call **any global server function** via `google.script.run`, regardless of what the UI shows them. The hardcoded `ROLES` map only controls *navigation*, not *capability*, unless each server function checks the caller.

**Findings, in severity order:**

1. **`WIPE_AND_RESET` (Server.js ~1610) has no role check.** Any authenticated Google user who finds the URL can delete every ART sheet and clear the trackers. This is a one-line-of-code-from-disaster situation. Add an admin-email check *and* consider requiring a confirmation token argument.
2. **`clearWeekForNext` (~1295), `saveWIP` (~824), `archiveWeek`, `fixFormulas` — no role checks.** Same exposure class: destructive or state-mutating, callable by anyone.
3. **`submitRequest` (~989) has no role restriction** — anyone can flood `REQUESTS`. Lower severity (requests still need approval) but enables spam/confusion, and combined with finding 4 it matters more.
4. **Client-trusted financial inputs:** rates/comm via `addContinuationRow` (see P1-4), and `periodId` is taken from the client (`ENTRY_PERIOD`) — a stale or manipulated client can write entries into a closed period. Server should validate the period is open.
5. **Approval authority is not server-verified per-request-type.** `processRequest` should verify the caller is the admin for payment/order approvals; currently the protection is primarily that only admin *sees* the approve tab.
6. **No audit trail for who did what** beyond approver columns on some rows. The unbuilt `AUDIT_LOG` is the fix: log `(timestamp, email, function, args-summary, result)` on every mutating call. Cheap to add as a wrapper.
7. **Sheet IDs in committed source** (`CONFIG.LIVE_SHEET_ID`/`DEV_SHEET_ID`). Low severity (the sheet is the real ACL), but Script Properties are the right home.
8. **Integrity positive:** the new idempotency guard in `processRequest` closes the double-approval/double-payment hole, and lot-cap checks (`lotStatus` OVER warnings) exist on entry. But note the lot cap *warns* and saves anyway (`toast('Over lot cap!')` after a successful save, Index.html:1176) — over-cap entries are recorded, relying on a human to notice the toast.

**Single most important fix in this whole document:** a 10-line `requireRole(['admin'])` / `requireRole(['admin','accounts'])` helper applied to every mutating server function. It costs an hour and removes the catastrophic tail risk.

---

## 5. Benchmark: Factory OS vs Off-the-Shelf ERP

Reasoning from general knowledge of these products (no live evaluation performed):

| Capability | Odoo (Community/Ent.) | Zoho One / Books | Tally Prime | SAP Business One | Factory OS today |
|---|---|---|---|---|---|
| Double-entry accounting | ✅ Full GL | ✅ Full GL | ✅ Its core strength | ✅ Full GL | ❌ None — payment records only |
| GST/statutory (India) | ⚠️ Via localisation | ✅ Strong | ✅ The Indian default | ⚠️ Partner-dependent | ❌ None |
| Audit trail | ✅ | ✅ | ✅ (Edit Log) | ✅ | ⚠️ REQUESTS only; no AUDIT_LOG |
| Multi-level BOM / MRP | ✅ Strong | ⚠️ Basic (Inventory) | ⚠️ Basic | ✅ Strong | ❌ Single-level TS activity list |
| Inventory valuation (FIFO/WAvg) | ✅ | ✅ | ✅ | ✅ | ❌ No material tracking at all |
| **Piece-rate contractor payments** | ❌ Not native — heavy customisation | ❌ Not native | ❌ Payroll ≠ piece-rate | ❌ Custom dev | ✅ **The core competency** |
| Dept-sequence gating w/ approval | ❌ Generic workflows, poor fit | ❌ | ❌ | ⚠️ Custom | ✅ Native, matches the factory exactly |
| Per-article rate cards w/ approval | ⚠️ Costing exists, approval flow custom | ⚠️ | ❌ | ⚠️ | 🔶 Designed (Labour Rate Card), not built |
| Mobile-first for low-tech users | ⚠️ Generic mobile UI | ⚠️ | ❌ Desktop | ❌ | ✅ Built for Arvind's phone |
| Cost to adapt to *this* factory | High (dev + consultant) | Medium-high | High (wrong shape) | Very high | Already adapted |

### What makes ERP-grade software hard (and where Factory OS stands)

- **Double-entry correctness.** Real ERPs guarantee every rupee has a debit and a credit; balances are provable. Factory OS records payments as standalone rows — there is no ledger, so "how much do we owe contractor X right now?" is computed by re-scanning history, and advances/deductions (a planned feature) have nowhere principled to live. *This is the conceptual gap the contractor-ledger work must respect: model advances and payments as signed ledger entries against a contractor account, even in Sheets.*
- **Statutory compliance.** GST invoicing, TDS on contractor payments (potentially relevant for piece-rate contractors above thresholds), e-way bills. Factory OS does none of this — which is fine *as long as Tally (or the accountant) remains the books of record*. The realistic architecture is Factory OS as the operational layer feeding summaries into Tally, not replacing it.
- **Audit trails.** ERPs make every change attributable and irreversible-by-default. Factory OS is halfway there (REQUESTS) with the other half missing (AUDIT_LOG, immutable entries).
- **Inventory valuation & multi-level BOM.** Genuinely hard (cost layers, variance, scrap). Factory OS hasn't started; the planned material-consumption feature should begin with simple quantity tracking, not valuation.

### Where custom beats off-the-shelf — the honest case for this codebase

Adees' actual operating reality — piece-rate contractors per department, per-article negotiated rates, weekly cash/UPI payment cycles, lot-size caps, dept-sequence dependencies, one approver, users who need a three-tap phone UI — is the *worst-case fit* for every product in the table above. Odoo could be bent to it for ₹15–30L of implementation effort and would still feel wrong to Arvind. Tally is the right *books of record* and the wrong *shop floor tool*. The strategic position writes itself: **Factory OS owns the shop floor and the contractor relationship; Tally/CA owns statutory books; the integration is a periodic export.** The risk to manage is not "should we have built custom" — it's that custom software without ERP-grade integrity discipline (locks, auth, ledgers, audit) eventually produces a payment dispute it cannot resolve. §4 and §6 Phase 0 are that discipline.

---

## 6. Recommended Phase Plan for Pending Features

Dependency-ordered. The bar for "done" throughout: **one complete end-to-end payment cycle (order → setup → entry → submission → approval → contractor paid → records reconciled) where nobody has to ask Ayush what to do.**

### Phase 0 — Integrity foundation (prerequisite for everything; ~the cheapest phase)
1. `requireRole()` guard on every mutating server function (§4.1–3).
2. `LockService` on every read-modify-write (§3 P0-1).
3. Unify `PAYMENT_HISTORY` headers (§3 P0-2).
4. `AUDIT_LOG` sheet + logging wrapper on mutating calls.
5. Snapshot-before-wipe (or retire `clearWeekForNext` in favour of PeriodId filtering).
6. Move ENV to Script Properties.

*Dependencies: none. Blocks: everything below, because every later feature writes money-adjacent data.*

### Phase 1 — Labour Rate Card (the stated biggest blocker)
Single authoritative `RATE_CARD` sheet: `(article, activity, contractor-class?, rate, comm, effective_from, approved_by, status)`. All entry/setup flows read rates from it server-side (kills the client-supplied-rate hole, P1-4, as a side effect). Rate changes go through the existing RATE_EDIT request pattern but match by ID, not rowIndex (P1-5).

*Depends on: Phase 0. Blocks: pre-costing, contractor dashboard accuracy, salary automation, AI cost analysis — essentially everything financial.*

### Phase 2 — Contractor ledger: payments dashboard + advance tracking
Model as a signed ledger (`CONTRACTOR_LEDGER`: date, contractor, type [EARNED/ADVANCE/PAYMENT/DEDUCTION], amount, ref). Earned entries derive from approved payment submissions; advances are a new request type; the dashboard is a per-contractor running balance. This is the double-entry-lite discipline from §5 and the foundation for salary automation later.

*Depends on: Phase 1 (correct rates → correct earned amounts) and Phase 0.3 (one PAYMENT_HISTORY schema).*

### Phase 3 — Prakash WIP role + job cards / stage caps
- Add the `prakash` role (one line in ROLES once `requireRole` exists) with a WIP-entry-only nav; harden `saveWIP` (currently unguarded, P0-class).
- Job cards / stage quantity caps: enforce that a dept cannot record more output than the prior dept's approved output (the data already exists in `approvedByAct`). Turn the current over-cap *warning* into a server-side *block* with an admin override request type.

*Depends on: Phase 0. Independent of Phases 1–2 — can run in parallel.*

### Phase 4 — Rejection/rework + dispatch
- Rejection: a `REJECTED_QTY` concept on entries (extra column / linked row) flowing into WIP reconciliation — rejected pairs must subtract from the qty eligible for downstream depts and payment.
- Dispatch: the sixth dept is currently just another activity row; give it order-completion semantics (dispatched qty vs lot size → order CLOSED state in `ORDER_TRACKER`).

*Depends on: Phase 3 (stage caps define the qty pipeline rejection subtracts from).*

### Phase 5 — Pre-costing + material consumption
- Pre-costing: at order creation, compute projected labour cost from the Rate Card (Phase 1) per pair × lot size; store on `ORDER_INDEX`; later compare actual vs projected.
- Material consumption: start with quantity-only issue/return per order (no valuation — see §5). This is the first feature that genuinely strains Sheets; consider it the Supabase trigger point.

*Depends on: Phase 1 (pre-costing), Phase 4 (actuals need rejection-adjusted quantities).*

### Phase 6 — Salary automation, buyer reports, AI cost analysis
- Salary automation = Phase 2 ledger + period close → generated payment instructions per contractor (UPI/cash/bank from `MASTER_CONTRACTORS`).
- Buyer reports: order status/dispatch summaries from `ORDER_TRACKER` — cheap once Phase 4 lands.
- AI cost analysis: only meaningful once Phases 1–5 produce clean, normalised data. Doing it earlier analyses noise.

*Depends on: everything above. This phase is the reward, not the path.*

**Sequencing note:** Phases 0→1→2 are strictly serial. Phase 3 can run parallel to 1–2. The milestone bar ("nobody asks Ayush") is realistically achieved at the end of **Phase 2** — at that point entry, rates, and contractor balances are all self-serve and self-consistent, and Ayush's role reduces to tapping Approve.

---

## 7. Open Questions

1. **Where do the statutory books live?** Is Tally (or the CA) the system of record for GST/TDS, with Factory OS feeding it? This determines whether Phase 2's ledger needs export formatting and whether TDS-on-contractor-payments ever becomes a Factory OS concern.
2. **Is `clearWeekForNext` still in active use**, or has the PeriodId mechanism (col K) fully replaced weekly wipes? If replaced, it should be deleted; if in use, Phase 0.5 is urgent.
3. **Are there pending legacy `NEW ORDER` (pipe-delimited) rows in LIVE's REQUESTS sheet?** If not, `createNewArtSheet` can be deleted now.
4. **What is the real concurrency level?** Today: Ayush + Arvind + store. Adding Prakash makes four. The LockService work assumes ≤10 concurrent users — fine for Sheets, but worth confirming no plan for contractor self-service logins (which would force the Supabase move earlier).
5. **Un-skip policy:** should a SKIPPED dept be recoverable via an admin request type (mirroring the new SETUP_EDIT_REQUEST), or is permanence intentional?
6. **Rate card granularity:** are rates per (article × activity), or do they also vary by contractor (senior vs junior karigars)? This decides the Phase 1 schema and is much cheaper to decide before building.
7. **Lot-cap policy:** should over-cap entries be blocked server-side (recommended in Phase 3) or remain warn-and-record? If contractors are sometimes legitimately paid for over-lot work (samples, replacements), a cap-override request type is needed.
8. **Size-run flexibility:** UK 6–11 is hardcoded; do Mothercare (kids) orders need different size sets soon? Determines whether the size-run becomes data-driven in Phase 4's dispatch work or earlier.
9. **DEV/LIVE data drift:** is the DEV sheet a stale copy or refreshed from LIVE? Testing payment logic against unrepresentative data is how the PAYMENT_HISTORY header fork survived this long.
10. **Migration trigger:** what observable event forces the Supabase move — order count (sheet read latency), contractor logins, or material valuation? Naming the trigger now prevents both premature migration and a panic migration.

---

*End of review. No code files were modified during this audit.*
