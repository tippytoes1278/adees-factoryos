// LEGACY — scheduled for removal.

function WIPE_AND_RESET() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().filter(isArtSheet).forEach(function(s){ ss.deleteSheet(s); });
  ['ORDER_TRACKER','ORDER_INDEX','PAYMENT_HISTORY','REQUESTS'].forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 3)
      sh.getRange(4, 1, sh.getLastRow()-3, 20).clearContent();
  });
  var wr = ss.getSheetByName('WIP_RECONCILIATION');
  if (wr && wr.getLastRow() > 4)
    wr.getRange(5, 1, wr.getLastRow()-4, 10).clearContent();
  var wm = ss.getSheetByName('WEEKLY PAYMENT MASTER');
  if (wm) wm.getRange('B57').clearContent();
  return { success:true };
}

function deleteOrphanedArt005() {
  var ss = SpreadsheetApp.openById(CONFIG.LIVE_SHEET_ID);
  var ws = ss.getSheetByName('ART-005');
  if (!ws) return 'ABORT: ART-005 sheet not found in LIVE';

  var article = String(ws.getRange('B2').getValue());
  var log = ['ART-005 B2 (article) = ' + article];

  // Safety: confirm no registry entries exist before deleting
  var registryRefs = [];
  ['ORDER_INDEX', 'ORDER_TRACKER', 'WIP_RECONCILIATION'].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (!s || s.getLastRow() < 2) return;
    var vals = s.getRange(1, 1, s.getLastRow(), s.getLastColumn()).getValues();
    vals.forEach(function(row, ri) {
      if (row.some(function(cell) { return String(cell).trim() === 'ART-005'; }))
        registryRefs.push(name + ' row ' + (ri + 1));
    });
  });

  if (registryRefs.length > 0) {
    return 'ABORT: ART-005 found in registry — ' + registryRefs.join(', ');
  }
  log.push('Registry check: clean (no ORDER_INDEX / ORDER_TRACKER / WIP_RECONCILIATION rows)');

  ss.deleteSheet(ws);
  SpreadsheetApp.flush();
  log.push('ART-005 sheet deleted from LIVE');

  var result = log.join('\n');
  Logger.log(result);
  return result;
}

function fixFormulas() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().filter(function(s){return s.getName().indexOf('ART-')===0;}).forEach(function(ws){
    for (var r=5; r<=49; r++) {
      ws.getRange('G'+r).setFormula('=IF(D'+r+'="",0,D'+r+'*F'+r+')');
      ws.getRange('I'+r).setFormula('=IF(D'+r+'="",0,(D'+r+'*E'+r+')+G'+r+'+IF(H'+r+'="",0,H'+r+'))');
    }
    ws.getRange('I50').setFormula('=SUM(I5:I49)');
  });
  return 'Formulas fixed';
}
