# WIP_RECONCILIATION teardown — PREPARED DIFF (do not apply until reviewed)

Status: **not applied.** This is the plan for a separate reviewed commit, kept out
of the Section F orphan-removal commit per instruction.

## Why this is a coordinated change, not an orphan removal

WIP_RECONCILIATION is the old "Arvind-entered qty vs Prakash-produced qty" cross-
check. Its data-input column (D, "produced") was only ever written by `saveWIP()`,
which is itself dead (no caller). With no writer, column D stays blank, the status
formula in column G always resolves to `"AWAITING"`, and the dashboard mismatch
count is therefore **permanently 0**. The Job Card flow superseded this cross-check.

But unlike the four clean orphans, the sheet is still threaded through live code in
four files, including a visible dashboard tile. So it must be removed as one
reviewed unit.

## Exact edits (5 sites across 4 files)

### 1. wip.js — delete `saveWIP()` (currently lines 10–20)
```js
function saveWIP(rowNum, produced) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var wr = ss.getSheetByName('WIP_RECONCILIATION');
    wr.getRange('D'+rowNum).setValue(produced);
    SpreadsheetApp.flush();
    return { success:true,
      status: safeStr(wr.getRange('G'+rowNum).getValue()),
      diff:   safeStr(wr.getRange('F'+rowNum).getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}
```
Zero callers (grep confirms definition-only). Safe to delete outright.

### 2. orders.js — delete WIP_RECON writer in the order-creation helper (currently ~128–138)
The block that seeds a reconciliation row + formulas on new-article creation:
```js
try {
  var wr = ss.getSheetByName('WIP_RECONCILIATION');
  if (wr) {
    var wrRow = Math.max(wr.getLastRow(), 4) + 1;
    wr.getRange(wrRow, 1, 1, 3).setValues([[newName, displayName, 'Upper Making']]);
    wr.getRange(wrRow, 5).setFormula("=IFERROR('"+newName+"'!Q2,0)");
    wr.getRange(wrRow, 6).setFormula('=IF(D'+wrRow+'="","--",D'+wrRow+'-E'+wrRow+')');
    wr.getRange(wrRow, 7).setFormula(
      '=IF(D'+wrRow+'="","AWAITING",IF(D'+wrRow+'=E'+wrRow+',"MATCH",IF(E'+wrRow+'>D'+wrRow+',"PAID > MADE","UNDER-PAID")))');
  }
} catch(e) { Logger.log('WR error: ' + e.message); }
```

### 3. orders.js — delete the second WIP_RECON writer in `createOrder` (currently ~313–320)
Same pattern, keyed on `artSheet`/`article`. Remove the whole `var wr = ...` block
through its closing brace. Leave the surrounding `ORDER_TRACKER` writes and the
`SpreadsheetApp.flush(); return {success:true,...}` intact.

### 4. orders.js — delete the WIP_RECON cleanup in `deleteOrder` (currently ~27–33)
```js
try {
  var wr = ss.getSheetByName('WIP_RECONCILIATION');
  if (wr && wr.getLastRow()>4) {
    var wrD=wr.getRange(5,1,wr.getLastRow()-4,1).getValues();
    for(var k=0;k<wrD.length;k++){if(safeStr(wrD[k][0])===sheetName){wr.deleteRow(k+5);break;}}
  }
} catch(e3) { failures.push('WIP_RECONCILIATION: '+e3.message); }
```
Removing this also removes one `failures.push` path — the `failures` array and its
warning logic stay valid (ORDER_TRACKER cleanup still uses them).

### 5. payments.js — delete the `mismatches` computation in the dashboard builder (currently ~134–143)
```js
var mismatches = 0;
try {
  var wr = ss.getSheetByName('WIP_RECONCILIATION');
  if (wr && wr.getLastRow() > 4) {
    var wrData = wr.getRange(5, 7, wr.getLastRow()-4, 1).getValues();
    wrData.forEach(function(r){
      if (safeStr(r[0]).indexOf('PAID > MADE') > -1) mismatches++;
    });
  }
} catch(e) { Logger.log('WR error: ' + e.message); }
```
Then remove `mismatches:mismatches` from the dashboard return object (currently ~304).
**Caution:** the return object is consumed client-side as `d.mismatches`; do site 6
in the same commit or the admin dashboard JS will read `undefined`.

### 6. js_admin.html — remove the tile + two alerts that read `d.mismatches`
- Line ~82 — the metric tile:
  ```js
  mrow.appendChild(fosMetricEl('WIP',String(d.mismatches),d.mismatches>0?'mismatch':'all matched',d.mismatches>0?'red':'blue'));
  ```
- Line ~107 — the yellow alert:
  ```js
  if(d.mismatches>0)c.appendChild(alert_('yellow',d.mismatches+' WIP mismatch','Prakash qty does not match Arvind entry'));
  ```
- Line ~109 — the "all clear" condition also references `d.mismatches===0`; simplify to
  `if(approved&&d.redCount===0)` once mismatches is gone.

## What the admin dashboard looks like afterward

The metric row is a **fixed 4-column CSS grid** (`css.html:54`,
`grid-template-columns:repeat(4,1fr)`). The row currently holds: **Payout · Orders ·
WIP · (Requests|Pending)**. If we simply delete the WIP tile, the grid keeps four
columns and the three survivors sit in columns 1–3 with an **empty hole in column 4**
— visually broken. So "does the tile disappear?" — not cleanly, not without one more
decision. Two clean options:

- **Option A — repurpose the tile (recommended).** Keep a tile labelled "WIP" but feed
  it real in-progress volume instead of the dead mismatch count, e.g. total pairs
  currently on the floor = sum of the `d.pipeline` stage values already computed for
  the pipeline tracker (`js_admin.html:74`). One-line value change, grid stays 4-wide,
  and the metric finally reflects the Job Card reality. The "all matched"/"mismatch"
  sub-label would change to something like "on floor".
- **Option B — drop to a 3-metric row.** Delete the tile and change `css.html:54` to
  `repeat(3,1fr)` (or make the grid `auto-fit`). Fewer moving parts, but loses a slot
  you may want later.

Alerts at 107/109 are pure removals either way (no layout hole — they're conditional
list items).

## Sheet itself

Leave the physical `WIP_RECONCILIATION` tab in both DEV and LIVE spreadsheets in
place for now — it's harmless once no code reads/writes it, and deleting tabs is a
separate, irreversible data decision. Flag for a later cleanup pass if desired.

## Suggested commit boundary

All six sites in ONE commit (they're interdependent: site 5's removal of
`d.mismatches` will break site 6's reads if split). Post-deploy verification:
open the admin dashboard, confirm the metric row renders with no empty 4th cell,
and check DevTools console is clean (no `undefined` from a stray `d.mismatches`).
