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
