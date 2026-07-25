# Phase 8 — Open Issues / To-Fix List

_Working list for issues found during Phase 8. Fold into factory-os-change-tracker.md when convenient._

| # | Area | Issue | Status | Notes |
|---|------|-------|--------|-------|
| 8.I1 | Admin → Pending Requests | MASTER_ACTIVITY request cards showed the raw JSON payload (e.g. `{"dept":"prep","activityName":"PUNCH","rate":12,"comm":0}`) instead of a readable summary. Seen on DEV @347. | 🟡 Verify | Root cause: card fell through to the generic `else` in `renderRequests` (js_admin.html) that prints `r.details` verbatim. Local code already has a friendly MASTER_ACTIVITY branch (js_admin.html ~660–665) rendering "New Master Activity / Activity / Dept · Rate". Shipped in DEV @351 — **reload and confirm**. If still raw, verify `r.type` equals `'MASTER_ACTIVITY'` exactly and the branch is reached. |

## Phase 8 build — shipped to DEV @351

- 8.1 Contractor Add/Delete (soft deactivate + reactivate, enrollment cleanup)
- 8.2 Contractor Profile & Documents (CONTRACTOR_PROFILE + CONTRACTOR_DOCS sheets; per-contractor Google Drive folder; `drive.file` scope added — one-time re-auth on next app open)
- 8.3 Balance to Pay per Activity (`getContractorAccount`, netted against advances/pending/paid)
- 8.P1 Size-grid Option A (numeric size + grade label side by side; also fixes label-keyed sizeBreakdown so ORDER_INDEX numeric size columns populate)

### Watch during trial
- New OAuth scope `drive.file` → re-authorization prompt on first document action.
- 8.P1: new orders store size run keyed by numeric size, so downstream floor/issue screens show numeric sizes (e.g. 39) for **new** orders; existing orders unaffected.
