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
2. Run `wc -c Index.html` — it should be approximately 180KB
3. If Index.html is less than 150KB — STOP immediately and report before proceeding

Never run `clasp pull` without explicit instruction from the user.
`clasp pull` overwrites local files with server state and can destroy local work.
Git history is the recovery source for all rollbacks.
