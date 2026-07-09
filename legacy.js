// LEGACY — scheduled for removal.

function WIPE_AND_RESET() {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Not authorised' };
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

function fixFormulas() {
  var user = getUserInfo();
  if (user.role !== 'admin') return 'Not authorised';
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
