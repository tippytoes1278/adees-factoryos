## Deployment and run rules

The `/run` skill may only be invoked when the user types `/run` with a
leading slash. The bare word "run" should never trigger it.

`clasp push` and `clasp deploy` must never be run without an explicit
separate user instruction to deploy. "Run the app" or "see how it looks"
is NOT such an instruction. If showing the live app requires a push,
stop and ask first.
