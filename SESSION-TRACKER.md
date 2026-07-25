# Session Tracker

## 22 Jul 2026 — Incident triage (Claude Code, blocked)

- **Issue 1 — Master activity submissions vanish silently.** Arvind submitted Punch +
  Hand Lasting ~2 days prior; no row reached REQUESTS, no visible error. Resubmit threw
  "Activities already approved for cutting department on this order". Suspected misroute
  through the order-activity-setup handler. Audit prompt (Q1–Q5) prepared but never run.
- **Issue 2 — ART collision.** WO-2026-046 registered ART-024; WO-2026-047's approval
  minutes later tried ART-024 again, failed, stranded a "Copy of ART-TEMPLATE", succeeded
  on retry as ART-025. ORDER_INDEX healthy for both. Suspected stale cached read
  (Speed-3, 5-min TTL). 1.13 failure handling worked as designed.
- **Issue 3 — Home showed "1 pending request", Requests tab showed 0.** Suspected
  type-filter mismatch.
- **Tooling blockers:** prompt pasted into bash; 401 fixed via /login; Advisor Tool 400
  (Sonnet 4.6 advisor × Opus 4.7 main) — fix: /advisor off.
- **Open items carried:** delete stranded template copy; LIVE REQUESTS row 61 (1.15)
  resubmission; 2.13 Upper Skiving flow for Arvind.

## 23 Jul 2026 — Investigation + fixes + LIVE deploy (Cowork)

- **Read-only investigation completed** (code audit + LIVE sheet + REQUESTS data;
  execution-log page cut short by a Chrome drop, approval timestamps substituted).
  Full detail in FINDINGS-2026-07-23.md. Verdicts:
  - Issue 1: confirmed misroute, but the real damage is the guard matching
    `'' === ''` on the missing sheet field — every dept was permanently locked
    (all six have approved master rows). Batch rejected before any write; error
    was only a transient toast.
  - Issue 2: cache theory killed. `createOrder` had no lock; REQUESTS rows
    131/132/133 all approved 22-Jul 12:26 (same minute) → overlapping executions
    both read max=ART-023; `setName` threw; copy stranded with no cleanup.
  - Issue 3: not a filter — `submitRequest`/`processRequest` never invalidated the
    5-min `dashboardData_` cache (all other mutation flows do).
  - Cross-cutting stale-cache theory: holds for Issue 3 only.
- **Fixes shipped** (DEV trial passed → 4-step LIVE protocol, LIVE @349, DEV @350,
  HEAD back at ENV='DEV'; consoles clean on both):
  - Issue 1: dedicated MASTER_ACTIVITY request type — `requestMasterActivity()`
    (per-item dup checks, partial-batch submit with skipped reporting),
    `processRequest` approval branch, client rerouted, admin card rendering,
    legacy guard hardened to never match empty sheet.
  - Issue 2: script lock in `createOrder`, single `nextArtName_()` source of truth,
    rename retry with recompute, stranded-copy cleanup on failure, plus
    `auditArtSheets()` diagnostic.
  - Issue 3: `clearDashCache_()` called from all 14 REQUESTS-mutating paths.
- **Manual cleanups done:** "Copy of ART-TEMPLATE" deleted from LIVE (verified
  template-only first).
- **Closed:** 1.15 — no PENDING/REVISION rows remain in LIVE REQUESTS.
- **Committed:** 59e26e0 (9 files, +349/−20) — rollback point per protocol.
- **Session closed 23-Jul-2026.** Carried forward to next session:
  - Arvind resubmits Punch + Hand Lasting on LIVE (real-world Issue-1 confirmation),
    then 2.13 Setup Edit Request flow. Note: master list has "Sneakers Upper Skiving
    (Men/Ladies)" under prep — no standalone "Upper Skiving".
  - Optional: run `auditArtSheets()` on LIVE for the old D4 orphan question
    (known gaps: ART-004/005/008/011).

## 23 Jul 2026 (later) — 8.B1 + 8.P2 fixes (Cowork)

Environment note: no shell/clasp available this session — code changes only, in
DEV local files. Deploy + live trial are user-driven (4-step protocol). ENV
confirmed `DEV` in config.js before starting.

- **8.B1 — Stage validation returned 0 (BLOCKING).** Root cause was NOT dept
  mismatch or a column read: predecessor "received" only counted job cards with
  status `COMPLETE`/`PAYMENT_PENDING`/`PAID`, excluding `PARTIAL`. A partially-
  received Cutting card (106 received, still PARTIAL) contributed 0 → Preparation
  showed 0 available. **Discrepancy vs the continuation brief:** the brief assumed
  this fix was already on DEV @352; the local code did NOT contain it (verified —
  the PARTIAL clause was absent in both call sites). So the bug reproduces on DEV
  @352 too; promoting @352 as-is would not fix LIVE.
  - Fix: added `PARTIAL` to the predecessor-received status set in both
    `getMaxIssuableForStage()` (jobcards.js ~784) and the `issueJobCard()`
    predecessor lock (jobcards.js ~134). DEV needs a redeploy to carry it.
- **8.P2 — Copy/Import activities (REGRESSION).** Rebuilt to match pre-refactor
  behavior confirmed by Ayush: import from a **Work Order** or **another Article**,
  pre-fill the editable setup form, edit if needed, then submit through the normal
  approval flow (no direct-apply).
  - Server (activities.js): `getActivitiesForCopy(source, mode)` — read-only;
    `mode:'article'` reads one ART sheet's approved activities, `mode:'workorder'`
    resolves the BOM via ORDER_INDEX to every ART sheet under it and merges
    (dedupe by dept+name). Returns `{success, byDept, count}`.
  - Client (js_arvind.html renderActivitySetup): "Import activities from…" panel
    with a Work Order picker and an Article picker; a shared `addRow()` builder now
    backs both the manual "+ Add Activity" button and import pre-fill (with
    per-dept duplicate guard). Submission still goes through `requestActivitySetup`.
- **8.I1** — MASTER_ACTIVITY friendly renderer confirmed present (js_admin.html
  renderRequests ~660–665). No code change; reload to confirm on running DEV.
- **WhatsApp** — added `waDiagnostic()` (requests.js, admin-only): reports masked
  Twilio Script Property presence and does a live test send returning HTTP status,
  to tell "creds missing" apart from the sandbox 24h-session expiry.
- **Note:** the brief referenced `factory-os-change-tracker.md`, which does not
  exist; logged here in SESSION-TRACKER.md instead.
- **Deployed DEV @353** (17 files) after a `clasp login` reauth (`invalid_rapt`).
  8.B1 validated on DEV: a PARTIAL Cutting card's received pairs now count toward
  the next stage's cap (Fitter showed 20 = Cutting received 50 − Fitter issued 30).

### 23 Jul 2026 (later still) — follow-ups from DEV trial

- **UX:** Import "Work Order" dropdown now labelled `WO · article — customer`
  (was bare BOM); approved departments on the Accounts/Admin setup screen now list
  their approved activities + rates (was only a green chip; Store already saw them).
- **`finish` vs `finishing` vocabulary bug (partial fix).** Canonical stored dept
  key is `finish` (activities.js/orders.js); `jobcards.js` and `payments.js` used
  `finishing`. This broke Packing job-card issuance (dept check never matched) and
  Packing-as-predecessor detection at Dispatch. Fixed in `jobcards.js` (4 maps) and
  `js_store.html` `DEPT_KEY_MAP`. **Still TODO:** `payments.js` has `finishing` in
  ~6 places — some are broken dept-matches (e.g. ~L859 rate filter), others are
  benign internal pipeline buckets (~L260/262 with a `finishing` bucket key). Needs
  a per-occurrence pass + a Finishing/Packing payment test before touching (critical
  file). `js_store.html`/`payments.js` `_MS`/`_plCounts` buckets left as-is.
- **Skipped-stage message** on Store issue screen reworded to cover skipped vs
  not-set-up ("this stage may be skipped, or not set up yet"). Hiding/flagging
  skipped movements outright needs dept-status plumbed to the Store screen — deferred.
- **Size-grid block (diagnosed, not a code bug in current logic).** On the DEV
  trial, issuing a partial Cutting qty to Fitter was blocked with "Only 0 of UK 9
  left" although the stage cap correctly allowed 20. Root cause: `getOrderSizeBalance`
  (orders.js) caps each stage per-size against the ORDER size run minus that stage's
  already-issued `SIZE_BREAKDOWN`. Prior trial cards JC-013/JC-014 have a stored
  breakdown summing to 60 (10×6) while only 30 pairs were issued — inconsistent data
  from an earlier trial, which consumed the full per-size run. **Latent integrity
  gap:** `issueDepartmentJobCard`/`issueJobCard` do not enforce
  `sum(sizeBreakdown) === pairsIssued`, so a mismatched breakdown can be stored.
  Recommended: validate with a FRESH order on @353; optionally add the sum guard.
- **Carried forward:** payments.js `finishing` pass (+ payment test); optional
  sizeBreakdown==pairsIssued guard; optional Store skipped-stage plumbing; then the
  9-step trial and 4-step LIVE promotion (promotes everything since @346).

### 23 Jul 2026 (later still²) — dept-vocabulary unification + size guard

- **Unified the department vocabulary to ONE scheme everywhere:**
  `cutting, prep, fitter, lasting, finish, dispatch` (display labels unchanged,
  e.g. Fitter still shows "Upper Making", finish shows "Finishing & Packing").
  Two vocabularies had coexisted: the stage/activity logic keys (…/finish) and the
  pipeline-widget buckets (…/upper/finishing), which is what let the finish/finishing
  bug hide. Changes:
  - **Logic (correctness):** payments.js — 5 dept-match maps `finishing`→`finish`
    (movement→deptKey used to resolve Finishing/Packing activity rates + the
    PAYMENT_HISTORY Department value). jobcards.js + js_store `DEPT_KEY_MAP` already
    done in the prior entry.
  - **Pipeline widget (display):** aligned `js_core` FOS_PIPELINE_STAGES keys
    (`upper`→`fitter`, `finishing`→`finish`), and every caller — `js_admin` (tracker
    + on-floor total), `js_store` (`_plCounts`/`_MS`), `js_arvind` (demo), and the
    server `pipeline` bucket in payments.js. Labels kept.
  - Removed the temporary `finishing` alias in `js_arvind` SHORT_TO_DEPT.
  - Full-repo sweep: zero `finishing` / `'upper'` bucket keys remain; legit
    `Upper IN` / `Upper Store` / `Upper Making` labels untouched.
- **Size-integrity guard (both decisions cleared):** `issueJobCard` and
  `issueDepartmentJobCard` now reject an issue where a provided `SIZE_BREAKDOWN`
  doesn't total the pairs issued (empty breakdown still allowed). Matching
  client-side check added to the Store issue screen (pairs and the size grid are
  independent inputs there — that's how the 60-vs-30 cards were created).
- **Needs testing after deploy:** a Finishing/Packing job card + its payment
  (proves the payments finish fix), and the CEO/Home pipeline tracker numbers
  (proves the widget re-key didn't zero a stage). Then fresh-order 9-step trial.
- **Repair util:** added `repairMismatchedSizeBreakdowns()` (jobcards.js) —
  rescales any card whose SIZE_BREAKDOWN ≠ PAIRS_ISSUED down to match (run from
  the Apps Script editor). Ran on DEV: fixed JC-2026-013 & 014 (60→30, 5/size).
### 25 Jul 2026 — per-size receive tracking (new)

Root cause the trial surfaced: `receiveJobCard` stored only an aggregate
`PAIRS_RECEIVED`, never which sizes came back — so the next stage couldn't know
per-size availability and fell back to the order run, disagreeing with the
aggregate stage cap on partials. Built the fix (user chose "before LIVE"):

- **Schema:** new JOB_CARDS column R `RECEIVED_BREAKDOWN` (ensureJobCardsSheet
  header + migration guard; getJobCards reads col 18 → `receivedBreakdown`).
- **receiveJobCard:** accepts `receivedSizeBreakdown`, validates it totals the
  pairs credited, checks cumulative per-size ≤ issued, and accumulates it into
  col 18 across partial receives.
- **Receive UI (js_store.html):** a FULL receive asks for no sizes (auto-derives
  them as the remaining un-received sizes); dropping the qty below the balance
  reveals a per-size grid (bounded per size, must total the qty) for partials.
- **getOrderSizeBalance (orders.js):** an intermediate stage's per-size cap now
  comes from the predecessor's received sizes; **safe fallback** — if the
  predecessor was received before per-size tracking (aggregate only), or it's the
  first stage, it falls back to the order run so in-flight/legacy orders (e.g. the
  current 203090) are never blocked. Fresh orders get exact per-size that agrees
  with the aggregate stage cap.

- **Store issue-screen size UX** (js_store.html): the grid used to pre-fill each
  box with its size-run remaining, so PAIRS defaulted to the size-run total (30)
  even when the stage cap was lower (20) — confusing. Now: an "Order size run"
  summary line above the grid; boxes default to 0; and a live "Selected: N — this
  stage can take X" total that turns red and blocks issue when N exceeds the stage
  cap. Per-size "X left" hints and the admin size-run override are unchanged.
