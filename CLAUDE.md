## LIVE DEPLOY PROTOCOL — must follow exactly, every time

Every LIVE deploy is a 4-step sequence. Never split across separate prompts.

Step 1: Change CONFIG.ENV to 'LIVE' in config.js
Step 2: Push and deploy to LIVE deployment ID
Step 3: Immediately change CONFIG.ENV back to 'DEV' in config.js  
Step 4: Push and deploy to DEV deployment ID

Never deploy to LIVE without completing all 4 steps in the same prompt.
Never leave HEAD at CONFIG.ENV = 'LIVE'.

CONFIG.ENV now lives in config.js (not Server.js — that file no longer exists).
Every deploy session must confirm ENV='DEV' in config.js before starting.

## DEPLOYMENT COMMANDS (permanent reference)

DEV deploy (all changes go here first):
cd /Users/Ayush/adees-factoryos && npx @google/clasp push && npx @google/clasp deploy --deploymentId AKfycbzqXhBq6rVaCZ0fcAA2lzQiBYVGKaV3qiIfQ6iZRxnIBpd2pL2nWzSGJQtMHBWfnHusXQ --description "dev"

LIVE deploy (4-step protocol only, never standalone):
cd /Users/Ayush/adees-factoryos && npx @google/clasp push && npx @google/clasp deploy --deploymentId AKfycbwnXfDSJ9AwkOrTGkb5h88QHTyL2ZUSPKxZ1_RousLQgkc5x9e0B5n7slrCj3lXnNLlxw --description "live"

DEV Sheet ID:  1eHnrG7IWn5PhreW1ywkdhgpzjOzYs6Y53vC4EIxwTvg
LIVE Sheet ID: 1FLPeuQFPx0nQXRy-16P2-1-e5SjDu7nLE-1ycNZ-IH0

LIVE is the factory's running system. DEV changes only until
trial script passes. Never touch both in the same session.

## Deployment and run rules

The `/run` skill may only be invoked when the user types `/run` with a
leading slash. The bare word "run" should never trigger it.

`clasp push` and `clasp deploy` must never be run without an explicit
separate user instruction to deploy. "Run the app" or "see how it looks"
is NOT such an instruction. If showing the live app requires a push,
stop and ask first.

## Working Directory — CRITICAL
The only correct project folder is `/Users/Ayush/adees-factoryos`.
There is NO other copy of this project anywhere on this machine.

Before making ANY changes, always confirm:
1. You are working in `/Users/Ayush/adees-factoryos`

FILE MAP (verify sizes before any edit session):
- config.js        ~2 KB  — CONFIG, SHEET_ID, ROLES, doGet, helpers
- orders.js       ~21 KB  — createOrder, deleteOrder, BOM/TS functions
- activities.js   ~32 KB  — getEntryData, activity setup/approval flows
- requests.js     ~17 KB  — submitRequest, processRequest, notifications
- contractors.js   ~9 KB  — contractor CRUD, enrollments
- wip.js          ~22 KB  — WIP entries, job card daily reports, grid
- jobcards.js     ~12 KB  — issueJobCard, receiveJobCard, getJobCards
- payments.js     ~41 KB  — getDashboardData, payment batch lifecycle
- legacy.js        ~4 KB  — WIPE_AND_RESET etc, do not edit
- Index.html       ~3 KB  — shell only: head, static markup, 6 includes
- css.html        ~21 KB  — all CSS styles
- js_core.html    ~27 KB  — boot, nav, helpers, common functions
- js_admin.html   ~29 KB  — Ayush's screens (home, orders, approvals, requests)
- js_arvind.html  ~79 KB  — Arvind's screens (new order, payment, contractors, activities)
- js_store.html   ~44 KB  — Prakash's screens (WIP, job cards, floor, grid)
- js_tutorial.html ~7 KB  — tutorial, help, what's new overlays

SIZE GUARD: No file should shrink more than 20% in a single edit
without explicit approval. If Index.html drops below 2KB or any
component file drops below 50% of listed size — stop and verify before
proceeding.

Never run `clasp pull` without explicit instruction from the user.
`clasp pull` overwrites local files with server state and can destroy local work.
Git history is the recovery source for all rollbacks.

## CLIENT-SIDE vs SERVER-SIDE RULES

Server utility functions (safeNum, safeStr etc.) are defined in
config.js and available to all server .js files automatically.
They are NOT available in the browser.

RULE 1: Never call server utility functions in js_*.html files.
Frontend files must define their own versions in js_core.html.

Current client-side utilities (defined in js_core.html):
- safeNum(v) — parseFloat with NaN fallback to 0
- fmt(n) — formats number as ₹ with Indian locale
- fmtPd(ds,yr) — formats period date string

If a fix in js_*.html needs safeNum or safeStr,
add the client-side version to js_core.html first.

RULE 2: Always open browser DevTools console after every deploy.
Zero red errors = pass. Any red error = something is broken
even if the page looks okay. Fix before calling deploy done.

RULE 3: When adding a new server utility to config.js that might
be used in frontend code, add a matching client-side version
to js_core.html in the same commit.
