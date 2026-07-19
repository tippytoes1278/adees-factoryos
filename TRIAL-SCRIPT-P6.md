# FactoryOS — Phase 6 Trial Script (DEV)

Run this on the **DEV** app before any LIVE push. It exercises the full job-card
lifecycle plus the four locking behaviours that changed this cycle, and the new
**one-card-per-department** model (multi-activity, per-contractor payment).

## Before you start

- **DEV only.** Confirm the top-left badge shows **DEV**.
- You need to switch between three Google logins (each is a role):
  - **Arvind** → `accounts@adeesexports.com` (New Order, Activities, Payment submit)
  - **Ayush** → `ayush@adeesexports.com` (Approvals, Payment approval)
  - **Prakash** → `hr@adeesexports.com` (Issue, Job Cards, Receive)
- Use the **/dev** link (Deploy → Manage deployments → its `.../dev` URL) so you see the latest pushed code. Hard-refresh after each login switch.
- Open DevTools console (⌥⌘J) for each role and watch for **red errors** throughout — zero red is the baseline pass condition at every step.

Pick a fresh article for the run (call it **TEST**), with a **lot size of 100** and set up activities in **at least Cutting, Fitter, and Dispatch** — deliberately leave **Preparation with no approved activities** so step 6 can test the skip-stage lookback.

---

## Step 1 — Create the order (Arvind)
New Order → create TEST with lot size **100** and a size run (e.g. UK6–UK11). Submit.
- **Pass:** order appears; a NEW_ORDER request goes to Ayush.

## Step 2 — Approve the order (Ayush)
Approvals/Requests → approve the new order.
- **Pass:** order becomes active; it now appears in Prakash's Issue dropdown.

## Step 3 — Activity setup, multiple departments (Arvind)
Activities → for TEST, set up activities in **Cutting** (2 activities to test multi-activity, e.g. "Synthetic Cutting" + "Lining Cutting"), **Fitter** (1+), **Dispatch** (1+). **Do NOT set up Preparation.** Submit each for approval.

## Step 4 — Approve activity setup (Ayush)
Approvals → approve each department's activity setup for TEST.
- **Pass:** approved activities now show on the Issue form for those departments.

## Step 5 — Issue, with the locking checks (Prakash)

### 5a — Multi-activity issue creates ONE card
Issue → Order **TEST** → Store **Upper Store** → Movement **Cutting IN**.
- The two Cutting activities appear as rows, each with its rate and a contractor dropdown; **Pairs to Issue** is one field (defaults to the stage max = **100**, the lot size, since Cutting has no predecessor).
- Assign a **different contractor** to each activity. Leave Pairs at **50**. Press **Issue All**.
- **Pass:** toast "Issued JC-… — 2 activities". In **Job Cards**, there is **ONE** new card (not two) showing both `activity — contractor` lines. In **Grid**, Cutting IN shows **50** (not 100).

### 5b — Lot-size cap blocks over-issue
Issue again → TEST → Upper Store → **Cutting IN**. Set Pairs to **60** (50 already issued, lot 100 → only 50 left). Assign a contractor. **Issue All**.
- **Pass (block):** fails with "…Order lot size is 100, and 50 already issued… Maximum available: 50 pairs."
- Redo with Pairs **50** → **Pass:** succeeds (Cutting now 100/100 issued).

### 5c — Next stage blocked before predecessor received
Receive **nothing** yet. Issue → TEST → **Fitter IN** (Fitter's nearest active predecessor is Cutting — Preparation is skipped). Any pairs. **Issue All**.
- **Pass (block):** "Maximum available for this stage: 0 pairs" — because 0 Cutting pairs have been *received* back yet.

### 5d — Next stage unlocks once predecessor has some received
Go to **Job Cards** → on one of the Cutting cards, **Receive 40** pairs (leave the other Cutting card open).
Now Issue → TEST → **Fitter IN** → Pairs **40** (max should read 40). Assign contractor. **Issue All**.
- **Pass:** succeeds even though a Cutting card is still open — the **received** 40 is what unlocks Fitter, not full completion.
- Try Pairs **50** → **Pass (block):** blocked at "Maximum available: 40 pairs".

## Step 6 — Skip-stage lookback (Prakash)
Still on TEST: confirm that issuing **Fitter IN** keyed off **Cutting** (received), **not** Preparation. Since Preparation has no approved activities, it must be skipped in the predecessor walk.
- **Pass:** step 5d working *is* this check — Fitter's max tracked Cutting-received, proving Preparation was skipped. (If Preparation had wrongly been treated as the predecessor, Fitter's max would have been 0 with a "Preparation completed 0" message.)

## Step 7 — Receive job cards (Prakash)
Job Cards → receive the remaining open cards to COMPLETE (receive full outstanding on each). Confirm each card flips **ISSUED → PARTIAL → COMPLETE** as expected.
- **Pass:** completed cards show the green COMPLETE state; balances reach 0.

## Step 8 — Submit payment (Arvind)
Payment → each **completed department card** shows as one card with a **per-contractor breakdown** (each contractor, their activities, `rate × pairs`, amount) and a card **total**. Press **Pay Card** on the multi-activity Cutting card.
- **Pass:** toast "Paid JC-… — ₹… · PAY-…". The card leaves the pending list. Two contractors were on that card → each is paid **only their** activity's rate (not the department sum).

## Step 9 — Approve payment (Ayush) → Pay History
Payment (admin) → approve the PAY-… batch. Then open **Pay History**.
- **Pass:**
  - The batch total **₹** equals the sum of the card's contractors' amounts.
  - **Pairs shows the physical count (e.g. 50), NOT 100** — even though two contractors were paid for those same 50 pairs.
  - The card's status is now **PAID**.
  - Each contractor appears in the batch line detail.

---

## Results checklist (tick each)

- [ ] Zero red console errors across all roles/steps
- [ ] 5a: multi-activity issue creates ONE card; Grid shows physical pairs once
- [ ] 5b: over-lot issue blocked with the lot-size message
- [ ] 5c: next stage blocked at 0 before any predecessor pairs received
- [ ] 5d: next stage unlocks on predecessor **received** (partial ok), and caps at received qty
- [ ] 6: Fitter keyed off Cutting, skipping empty Preparation
- [ ] 7: receive flow ISSUED→PARTIAL→COMPLETE works
- [ ] 8: payment shows per-contractor breakdown; each paid only their activity rate
- [ ] 9: approval → PAID; Pay History total correct; **pairs not double-counted**

Any red error or a ✗ on 5b/5c/5d/8/9 → stop and report which step + the exact message before considering LIVE.
