# Payment `periodId` gap — assessment (no code changes)

Priority item flagged in the Step-1 period audit. This is analysis only — nothing
here is applied.

## Recap of the defect

`submitCardPayment` (payments.js:983) and `submitCardAdvance` (payments.js:1398)
both take `data.periodId` **verbatim from the client** and stamp it into
`PAYMENT_HISTORY` column A. The client builds that value as
`ENTRY_PERIOD ? ENTRY_PERIOD.periodId : 'PAY-' + today` (js_arvind.html:609, 682),
and `ENTRY_PERIOD` is dead — always `null`. So every payment/advance is stamped
`PAY-yyyy-mm-dd` (the day it was paid), never the real weekly `W-yyyyMMdd` period
that Job Card WIP entries use. Result: Pay History "groups by period" but is
actually grouping by day-paid.

## (a) Where to resolve `periodId` server-side

Do it inside the two server functions, and stop trusting the client value.

The correct source of truth already exists and is used by the Job Card path
(jobcards.js:221–232, duplicated at 356–368 and 529–541): read `PAYMENT_PERIODS`,
filter `Status === 'OPEN'`, sort IDs, take the first. Recommended shape:

1. Add one shared helper in payments.js, next to `ensureCurrentPeriod()`:
   ```js
   function resolveOpenPeriodId() {
     ensureCurrentPeriod();                    // guarantee this week's row exists
     var ss = SpreadsheetApp.openById(SHEET_ID);
     var pp = ss.getSheetByName('PAYMENT_PERIODS');
     if (pp && pp.getLastRow() > 1) {
       var v = pp.getRange(2,1,pp.getLastRow()-1,7).getValues(), open = [];
       v.forEach(function(r){ if(safeStr(r[6]).trim().toUpperCase()==='OPEN') open.push(safeStr(r[0])); });
       open.sort();
       if (open.length) return open[0];
     }
     return 'W-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
   }
   ```
   Note this calls `ensureCurrentPeriod()` first — closing the secondary gap the
   audit raised, where the Job Card path only *reads* an OPEN row and silently
   falls back to a synthetic `JC-<date>` if none exists yet.

2. In `submitCardPayment` and `submitCardAdvance`, replace the current
   `var periodId = safeStr(data.periodId || '').trim(); if (!periodId) return err;`
   with `var periodId = resolveOpenPeriodId();` and drop the "periodId is required"
   guard (it's now server-derived, never client-required).

3. Optional cleanup once the above lands: delete the dead `ENTRY_PERIOD` global and
   its `? .periodId : 'PAY-'+…` guards in js_arvind.html (609, 682) and js_core.html
   (5, 211, 289, 316, 415, 453). Behaviour won't change (server ignores the client
   value), but it removes the misleading dead variable. Fold this into the same
   commit or a trailing cleanup commit — not into Section F.

Also worth folding in: the three duplicated OPEN-period blocks in jobcards.js should
call `resolveOpenPeriodId()` too, so entry/WIP and payment share one definition of
"current period." Not required for the fix, but it's the reason the two sides
drifted in the first place.

## (b) Existing LIVE rows: leave as mixed history, or backfill?

**Recommendation: leave them; do not backfill automatically.** Reasoning:

- The fix is forward-only by nature — it changes what future rows get stamped.
  Past `PAY-yyyy-mm-dd` rows remain valid records of real payments; only their
  grouping key differs.
- A backfill means rewriting column A on historical `PAYMENT_HISTORY` rows in the
  **LIVE** sheet — money-adjacent, irreversible, and exactly the kind of thing the
  4-step LIVE protocol exists to gate. Not worth the risk for a reporting-only key.
- Mapping a `PAY-2026-07-14` row back to its `W-*` week is deterministic (the pay
  date falls in exactly one Sat–Fri window), so IF a clean historical view is ever
  wanted, it can be done as a **read-time** derivation in the Pay History grouping
  code (map any `PAY-<date>` to the enclosing `W-<Saturday>` when grouping) — no
  data mutation. That's the safe way to get unified history without touching LIVE
  rows.
- Consequence of leaving as-is: Pay History will show a boundary — older entries
  grouped by day (`PAY-*`), newer ones grouped by week (`W-*`). If that visible
  seam is undesirable, prefer the read-time mapping above over a data backfill.

Net: fix forward, optionally add a read-time `PAY-*`→`W-*` normaliser in the Pay
History view, and do **not** run a write-backfill against LIVE.

## (c) Interaction with the Section A advance ledger

This is the reassuring part. **The advance ledger's money math does not depend on
`periodId` at all.**

- Advances and final payments are netted by `_paidPairsMap()` (payments.js:1225),
  which keys strictly on `jobCardId || contractorId` and sums paid pairs. It never
  reads column A (period). So an advance stamped `PAY-2026-07-14` still correctly
  offsets a final payment stamped `W-20260711` for the same card+contractor — the
  "pay only pairs not already paid" logic (submitCardPayment ~1078–1092) stays
  correct regardless of period strings.
- Therefore fixing `periodId` **cannot change any payable amount, advance offset, or
  card balance.** It is safe to land before the Section A ledger trial — it won't
  perturb the numbers the trial validates.
- The only place the two interact is *reporting*: an advance and its final payment
  currently can land in different period buckets (different `PAY-*` days). After the
  fix, both resolve to the same OPEN `W-*` period if paid in the same week, so a
  period-grouped ledger view becomes internally consistent. That's a reporting
  improvement, not a math change.

### Sequencing recommendation
Because the fix is math-neutral for the ledger, either order works, but cleanest is:
land the `periodId` server-resolution fix **first** (small, forward-only, no LIVE
data change), then run the Section A advance-ledger trial on top of a consistent
period model — so the trial exercises the corrected grouping rather than the
day-stamped legacy behaviour. Confirm before doing either, since both touch money
code and Section A is still on hold per your instruction.
