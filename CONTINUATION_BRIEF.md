# FACTORY OS — CONTINUATION BRIEF (paste into new Cowork window)

## Read first
- Project folder (ONLY correct copy): /Users/Ayush/adees-factoryos
- Read /Users/Ayush/adees-factoryos/CLAUDE.md (deploy protocol, file map, client/server rules)
- Read /Users/Ayush/adees-factoryos/PHASE8_OPEN_ISSUES.md
- Google Apps Script + Google Sheets project, deployed via clasp.
- RULES: all changes go to DEV first. NEVER run `clasp push`/`deploy` without an explicit
  separate user instruction. Confirm CONFIG.ENV='DEV' in config.js before starting.
  LIVE only via the 4-step protocol in CLAUDE.md.

## Current deploy state
- DEV  = @352  (deployment AKfycbzqXhBq6rVaCZ0fcAA2lzQiBYVGKaV3qiIfQ6iZRxnIBpd2pL2nWzSGJQtMHBWfnHusXQ, sheet 1eHnrG7IWn5PhreW1ywkdhgpzjOzYs6Y53vC4EIxwTvg)
- LIVE = @346  (deployment AKfycbwnXfDSJ9AwkOrTGkb5h88QHTyL2ZUSPKxZ1_RousLQgkc5x9e0B5n7slrCj3lXnNLlxw, sheet 1FLPeuQFPx0nQXRy-16P2-1-e5SjDu7nLE-1ycNZ-IH0)
- DEV @352 already contains: Phase 8 (8.1 contractor add/delete, 8.2 profile+Drive docs [new
  drive.file scope -> one-time re-auth], 8.3 balance-to-pay-per-activity, 8.P1 numeric+label size
  grid), the Notifications-tab rewrite, and the Add-Contractor profile fields.
- NOTE 8.P1: New Order now saves sizeBreakdown keyed by NUMERIC size; new orders show numeric
  sizes on floor/issue screens (existing orders unchanged).

## WORK QUEUE (in order)

### 1) 8.B1 — Stage validation returns 0 (BLOCKING) [do first]
Symptom: getMaxIssuableForStage() shows 0 pairs for Preparation though 106 received in Cutting.
Reported on the LIVE @346 trial.
CRITICAL: The DEV code already has a documented 8.B1 fix. In jobcards.js
`getMaxIssuableForStage()` (~line 720-819), predecessor "received" now counts PARTIAL cards too
(status in PARTIAL/COMPLETE/PAYMENT_PENDING/PAID, ~line 784). LIVE @346 predates this fix.
So:
  a) FIRST reproduce on DEV @352 with a real order.
     - If NOT reproducible on DEV -> no code change needed; the resolution is promoting DEV->LIVE
       (4-step), because the fix simply isn't on LIVE yet.
     - If STILL reproducible on DEV -> debug deeper:
        * Predecessor detection: `orderActiveDepts` is built from getApprovedActivitiesForArticle()
          dept strings matched against STAGE_DEPT_KEY ('cutting','prep','fitter','lasting',
          'finishing','dispatch'). A dept-string mismatch makes predecessorStage=null and it falls
          back to the lot-size branch. Verify the order's approved activities carry the expected dept.
        * Confirm the Cutting JOB_CARDS row: STATUS is PARTIAL/COMPLETE/etc AND pairsReceived (col H)
          truly = 106.
        * Movement strings: predecessor of Preparation = Cutting; predMovements = ['Cutting IN'].
          Confirm receiveJobCard() writes received pairs where this read expects them.
     - Add Logger.log for predecessorStage / predReceived / alreadyIssued and inspect execution logs.
Key functions: getMaxIssuableForStage (jobcards.js ~720), getApprovedActivitiesForArticle
(activities.js ~359), receiveJobCard (jobcards.js ~481), getJobCards (jobcards.js ~595).
JOB_CARDS cols: A JOB_CARD_ID ... F CONTRACTOR_ID, G PAIRS_ISSUED, H PAIRS_RECEIVED ... N STATUS ... Q ASSIGNMENTS.

### 2) 8.P2 — Copy Activity feature (REGRESSION) [do second]
Arvind lost the ability to copy activities from one article/order to another (lost in the Job Card
refactor). Grep confirms NO remaining copy code in activities.js or js_arvind.html -> rebuild.
Plan:
  - Check git history first: `git -C /Users/Ayush/adees-factoryos log --oneline` and search "copy".
  - Backend: copyActivities(fromOrderRef, toOrderRef) reading approved activities from source and
    applying to target via the normal activity-setup path. Understand how ACTIVITY_SETUP requests +
    approval currently persist approved activities (requests.js submitRequest/processRequest;
    activities.js activity storage + getApprovedActivitiesForArticle).
  - Frontend: add a "Copy activities from another order" control on the Activities screen
    (js_arvind.html renderActivities, ~line 1200+). Source-order dropdown from D.entry.articles.
  - Decide with user: should copy go through Ayush approval or apply directly (admin)? Match pre-refactor behavior.

### 3) Final 9-step trial on DEV
Fresh order -> activity setup -> COPY activities -> issue job card -> receive -> progress each stage
(verify getMaxIssuableForStage unlocks Preparation/Fitter/etc.) -> payment. All green before LIVE.

### 4) LIVE promotion (only after trial passes AND explicit user go-ahead)
Use the 4-step protocol in CLAUDE.md. This promotes EVERYTHING since @346: all Phase 8, notifications,
add-contractor, plus 8.B1/8.P2. LIVE will prompt a one-time Drive re-auth (drive.file scope) on first
document action. Confirm CONFIG.ENV value + resolved SHEET_ID in the report-back, not just version #.

## Deploy commands
DEV:  cd /Users/Ayush/adees-factoryos && npx @google/clasp push && npx @google/clasp deploy --deploymentId AKfycbzqXhBq6rVaCZ0fcAA2lzQiBYVGKaV3qiIfQ6iZRxnIBpd2pL2nWzSGJQtMHBWfnHusXQ --description "dev"
LIVE (4-step only): ...--deploymentId AKfycbwnXfDSJ9AwkOrTGkb5h88QHTyL2ZUSPKxZ1_RousLQgkc5x9e0B5n7slrCj3lXnNLlxw --description "live"

## Also open (lower priority, from prior session)
- WhatsApp notifications not arriving: likely Twilio SANDBOX 24-hour session expiry. Code path is
  fine (requests.js notifyNewRequest_ -> _sendWhatsApp_ ~125) but errors are muted
  (muteHttpExceptions + no status check) and it no-ops if any Script Property is missing
  (TWILIO_ACCOUNT_SID/TWILIO_AUTH_TOKEN/TWILIO_WHATSAPP_FROM/NOTIFY_WHATSAPP_TO). Email path
  (MailApp) is separate. Optional: add waDiagnostic() that reports masked cred presence + live
  Twilio HTTP status.
- 8.I1: MASTER_ACTIVITY cards showed raw JSON in admin Pending Requests. Friendly renderer already
  exists in js_admin.html renderRequests (~line 660); verify on @352 with a hard reload.
- factory-os-change-tracker.md not yet updated for Phase 8 + these fixes.
