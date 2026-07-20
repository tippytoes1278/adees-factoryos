# Period Model — Audit Findings (read-only, 2026-07-20)

## How PERIOD_ID gets created

`ensureCurrentPeriod()` (payments.js:316) is the only place a period row is born.
It buckets by calendar week, Saturday–Friday: `periodId = 'W-' + yyyyMMdd` of that
week's Saturday. On each call it checks `PAYMENT_PERIODS` for that ID; if missing,
it appends a new row with `Status = OPEN`. It never closes a period — nothing in
the codebase flips a row from OPEN to CLOSED (no writer sets Status elsewhere).

`ensureCurrentPeriod()` has exactly one caller: `getEntryData()` (activities.js:9).
It runs on effectively every screen load that touches entry/dashboard data (home,
approvals, payment, requests tabs all call `getEntryData` per js_core.html).

## How PERIOD_ID gets stamped on Job Card flow

`issueJobCard`, `issueDepartmentJobCard`, and `receiveJobCard` (jobcards.js:221,
356, 529 — three separate copies of the same 12-line block) resolve the period
by reading `PAYMENT_PERIODS` directly, filtering `Status === 'OPEN'`, sorting
IDs ascending, and taking the first. This periodId is passed into `saveWipEntry`.
**None of these three functions call `ensureCurrentPeriod()` themselves** — they
only consume whatever row already exists. If no OPEN row exists (e.g. sheet
missing, or the week just rolled over and nothing has called `getEntryData` yet),
they fall back to a synthetic `'JC-' + today's date` string that will never match
a real weekly period and won't be picked up by anything that groups on `W-*` IDs.

**Practical risk:** correctness of Job Card period-stamping is soft-coupled to
`getEntryData` having run at least once since the last Saturday rollover, in the
same execution context. In practice this is very likely (home tab alone triggers
it), but it is not guaranteed by the job-card code path itself — it's an implicit
dependency, not an explicit one.

## Is period assignment fully automatic? Yes for entry/WIP — no manual selector exists

- `ENTRY_PERIOD` is a client global (`js_core.html:5`), declared `null`, reset to
  `null` on every `refreshTab()` (js_core.html:453), and **never assigned a real
  value anywhere in the codebase.** It is a leftover from the old Entry-tab period
  selector. Every read of it (`js_core.html:211,289,316,415`) evaluates to the
  `null` branch, so `getEntryData(null)` is always what actually executes — which
  is fine, because `getEntryData` already auto-resolves to the first OPEN period
  itself (activities.js:29: `effectivePeriodId = periodId || firstOpenId`).
- No period `<select>`/dropdown exists anywhere for *choosing* a period to enter
  data against. Confirmed via grep across js_admin.html, js_arvind.html,
  js_store.html — zero hits for an entry-side period selector.
- **Conclusion:** the old manual period selector is gone and its removal did not
  break anything, because the fallback path it left behind (`ENTRY_PERIOD ?
  ENTRY_PERIOD.periodId : null`) always takes the `null`/auto branch anyway.
  `ENTRY_PERIOD` itself is dead code (flagged for Section F).

## Gap found: Payment/Advance submission does NOT use the auto-resolved period

`submitCardPayment` and `submitCardAdvance` (payments.js:983, 1398) take
`data.periodId` **verbatim from the client** and stamp it straight into
`PAYMENT_HISTORY` — there is no server-side re-resolution against
`PAYMENT_PERIODS` here (unlike the Job Card WIP path).

The client call sites (js_arvind.html:609, 682) build that payload as:
```
periodId: ENTRY_PERIOD ? ENTRY_PERIOD.periodId : 'PAY-' + new Date().toISOString().slice(0,10)
```
Since `ENTRY_PERIOD` is always `null` (see above), **every payment and advance is
always stamped with the synthetic fallback `'PAY-yyyy-mm-dd'` (today's date),
never with the real weekly `W-yyyyMMdd` period ID** that Job Card WIP entries use.

Consequences:
- Payments made on different days within the same weekly period get different
  `periodId` values (`PAY-2026-07-20` vs `PAY-2026-07-21`), so they will never
  group together — whereas the WIP/Job-Card side of the same work correctly
  groups under one `W-*` ID for the whole week.
- Payment History's "Payouts by Period" and the Pay History period filter
  (js_arvind.html:726–773) are grouping/filtering on this daily synthetic ID, not
  the real period. The UI still works (it groups on whatever string is present),
  but it is effectively grouping by *day paid*, not by *pay period* — likely not
  the intended semantics for "Payment History by period."

This looks like a real inconsistency, not intentional design — recommend
fixing in a follow-up: either resolve `periodId` server-side in
`submitCardPayment`/`submitCardAdvance` the same way `jobcards.js` does, or have
the client fetch the current open period instead of relying on the dead
`ENTRY_PERIOD` variable.

## Pay History grouping

`renderPayHistory()` (js_arvind.html:707) pulls `D.payBatches` (from
`getPaymentBatches`, payments.js:1109), groups client-side by `b.periodId`
descending (`periodOrder.sort` newest-first), with dropdown filters for period/
article/contractor. This is a pure display/filter layer — it does not assign or
resolve periods, it just groups whatever `periodId` string was stamped at
submission time (see gap above).

## Summary for Section F candidates

- `ENTRY_PERIOD` (js_core.html) — dead variable, always null, safe to remove
  along with its `? .periodId : null` guards, once the payment-period gap above
  is fixed or explicitly deferred (removing it without fixing the payment path
  would just hardcode the `null` branch, no behavior change either way).
