function getPendingRequests() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rq = ss.getSheetByName('REQUESTS');
  var requests = [];
  try {
    if (rq && rq.getLastRow() > 3) {
      rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues().forEach(function(r,i){
        if (!r[0]) return;
        requests.push({ rowNum:i+4, reqId:safeStr(r[0]),
          date:safeStr(r[1]), submittedBy:safeStr(r[2]),
          type:safeStr(r[3]), details:safeStr(r[4]),
          status:safeStr(r[5]), ayushNotes:safeStr(r[6]),
          approvedOn:safeStr(r[7]), executed:safeStr(r[8]),
          sheetCreated:safeStr(r[9]) });
      });
    }
  } catch(e) {}
  var mrMap = {};
  try {
    var maS = ss.getSheetByName('MASTER_ACTIVITIES');
    if (maS && maS.getLastRow() > 1)
      maS.getRange(2, 1, maS.getLastRow()-1, 2).getValues().forEach(function(r,i){ mrMap[2+i] = safeStr(r[1]); });
  } catch(e) {}
  var oiMap = {};
  try {
    var oiS = ss.getSheetByName('ORDER_INDEX');
    if (oiS && oiS.getLastRow() > 3)
      oiS.getRange(4, 1, oiS.getLastRow()-3, 5).getValues().forEach(function(r) {
        var sn = safeStr(r[1]);
        if (sn) oiMap[sn] = { bom:safeStr(r[0]), article:safeStr(r[2]), color:safeStr(r[3]), customer:safeStr(r[4]) };
      });
  } catch(e) {}
  requests.forEach(function(req) {
    if (req.type === 'RATE_EDIT') {
      try { var pl = JSON.parse(req.details); if (pl && pl.rowIndex) req.activityName = mrMap[pl.rowIndex] || ''; } catch(e) {}
    }
    if (req.type === 'ACTIVITY_SETUP' || req.type === 'SETUP_EDIT_REQUEST' || req.type === 'EDIT_REQUEST') {
      try {
        var plOi = JSON.parse(req.details);
        var snOi = safeStr(plOi.sheet || '');
        if (snOi && oiMap[snOi]) req.orderInfo = oiMap[snOi];
      } catch(e) {}
    }
    if (req.type === 'PAYMENT_SUBMISSION') {
      try {
        var pl2 = JSON.parse(req.details);
        if (pl2 && pl2.sheet) {
          var ws2 = ss.getSheetByName(pl2.sheet);
          if (ws2) {
            var psActs = [], psTotal = 0;
            ws2.getRange(5, 1, 45, 12).getValues().forEach(function(row) {
              if (!safeStr(row[1]).trim() || safeNum(row[0]) <= 0) return;
              var st2 = safeStr(row[11]).toUpperCase();
              if (st2 !== 'SUBMITTED' && st2 !== 'APPROVED') return;
              if (pl2.periodId && safeStr(row[10]) !== pl2.periodId) return;
              var qty2 = safeNum(row[3]);
              if (!qty2) return;
              var tot2 = safeNum(row[8]);
              psTotal += tot2;
              psActs.push({ activity:safeStr(row[1]).trim(), contractor:safeStr(row[2]),
                qty:qty2, rate:safeNum(row[4]), total:tot2 });
            });
            req.psActivities = psActs;
            req.psGrandTotal = psTotal;
            req.psArticle = pl2.article || '';
            req.psPeriodId = pl2.periodId || '';
          }
        }
      } catch(pe2) {}
    }
  });
  return { requests:requests };
}

function getMyRequests() {
  var user = getUserInfo();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rq = ss.getSheetByName('REQUESTS');
  var requests = [];
  try {
    if (rq && rq.getLastRow() > 3) {
      rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues().forEach(function(r,i){
        if (!r[0] || safeStr(r[2]) !== user.name) return;
        requests.push({
          rowNum: i+4, reqId: safeStr(r[0]), date: safeStr(r[1]),
          submittedBy: safeStr(r[2]), type: safeStr(r[3]), details: safeStr(r[4]),
          status: safeStr(r[5]), ayushNotes: safeStr(r[6]),
          approvedOn: safeStr(r[7]), processed: safeStr(r[8]), revisionHistory: safeStr(r[9])
        });
      });
    }
  } catch(e) {}
  var mrMap = {};
  try {
    var maS = ss.getSheetByName('MASTER_ACTIVITIES');
    if (maS && maS.getLastRow() > 1)
      maS.getRange(2, 1, maS.getLastRow()-1, 2).getValues().forEach(function(r,i){ mrMap[2+i] = safeStr(r[1]); });
  } catch(e) {}
  var oiMap = {}, oiBomMR = {};
  try {
    var oiS2 = ss.getSheetByName('ORDER_INDEX');
    if (oiS2 && oiS2.getLastRow() > 3)
      oiS2.getRange(4, 1, oiS2.getLastRow()-3, 4).getValues().forEach(function(r) {
        var sn = safeStr(r[1]);
        if (sn) { oiMap[sn] = safeStr(r[2]) + (r[3] ? ' - ' + safeStr(r[3]) : ''); oiBomMR[sn] = safeStr(r[0]); }
      });
  } catch(e) {}
  requests.forEach(function(req) {
    if (req.type === 'RATE_EDIT') {
      try { var pl = JSON.parse(req.details); if (pl && pl.rowIndex) req.activityName = mrMap[pl.rowIndex] || ''; } catch(e) {}
    }
    if (req.type === 'SETUP_EDIT_REQUEST' || req.type === 'ACTIVITY_SETUP') {
      try { var pl = JSON.parse(req.details); if (pl && pl.sheet) { req.articleLabel = oiMap[pl.sheet] || pl.sheet; req.bom = oiBomMR[pl.sheet] || ''; } } catch(e) {}
    }
  });
  requests.reverse();
  return { success:true, requests:requests };
}

function notifyNewRequest_(reqId, type, rawDetails, submittedBy, now) {
  try {
    var pl = {};
    try { pl = JSON.parse(rawDetails); } catch(e) {}
    var subject = 'Factory OS — New ' + type + ' request';
    var lines = [];
    switch (type) {
      case 'NEW_ORDER':
        lines.push((pl.styleName||'?') + (pl.color ? ' — ' + pl.color : '') + ' for ' + (pl.buyer||'?'));
        lines.push('BOM: ' + (pl.tsNumber||'?') + ' | Lot: ' + (pl.lotSize||0) + ' pairs | Del: ' + (pl.deliveryDate||'?'));
        break;
      case 'PAYMENT':
        lines.push('Week entries submitted for approval: ' + (pl.weekLabel||'?'));
        break;
      case 'CUSTOM_PERIOD_REQUEST':
        lines.push('Custom Period: ' + (pl.fromDate||'?') + ' → ' + (pl.toDate||'?'));
        if (pl.reason) lines.push('Reason: ' + pl.reason);
        break;
      case 'EDIT_REQUEST':
        lines.push('Edit Request: ' + (pl.activityName||'?') + ' on ' + (pl.sheet||'?'));
        if (pl.rowNum) lines.push('Row: ' + pl.rowNum);
        break;
      case 'SETUP_EDIT_REQUEST':
        lines.push('Setup Edit: ' + (pl.dept||'?') + ' dept on ' + (pl.sheet||'?'));
        break;
      case 'ACTIVITY_SETUP':
        if (pl.activityName) {
          lines.push('New Activity: ' + pl.activityName);
          lines.push('Article: ' + (pl.sheet||'?') + ' | Dept: ' + (pl.dept||'?') + ' | Rate: ₹' + (pl.rate||0) + '/pr' + (pl.comm ? ' | Comm: ₹' + pl.comm + '/pr' : ''));
        } else {
          lines.push('Activity Setup: ' + (pl.sheet||'?') + (pl.dept ? ' — ' + pl.dept + ' dept' : ''));
          var acts = pl.activities || [];
          acts.slice(0, 5).forEach(function(a) { lines.push('  • ' + (a.activityName||'?') + ' | ₹' + (a.rate||0) + '/pr'); });
          if (acts.length > 5) lines.push('  … and ' + (acts.length - 5) + ' more');
        }
        break;
      case 'RATE_EDIT':
        lines.push('Rate Edit: Row ' + (pl.rowIndex||'?') + ' → ₹' + (pl.newRate||0) + ' rate, ₹' + (pl.newComm||0) + ' comm');
        break;
      case 'PAYMENT_SUBMISSION':
        lines.push('Article: ' + (pl.article||pl.sheet||'?'));
        lines.push('Period: ' + (pl.periodId||'?') + ' | Activities: ' + (pl.count||'?'));
        break;
      default:
        lines.push(rawDetails);
    }
    lines.push('');
    lines.push('Submitted by: ' + submittedBy + ' on ' + now);
    lines.push('Request ID: ' + reqId);
    lines.push('');
    lines.push('Open Factory OS to review.');
    MailApp.sendEmail('ayush@adeesexports.com', subject, lines.join('\n'));
  } catch(mailErr) {
    Logger.log('notifyNewRequest_ mail error: ' + mailErr.message);
  }
}

function submitRequest(type, details) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var user = getUserInfo();
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var rq = ss.getSheetByName('REQUESTS');
      var lastRow = Math.max(rq.getLastRow(), 3) + 1;
      var reqId = 'REQ-' + String(lastRow-3).padStart ? String(lastRow-3).padStart(3,'0') : ('00'+(lastRow-3)).slice(-3);
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
      var revHistory = '';
      try { var dp_ = JSON.parse(details); if (dp_ && dp_.revisionRemark) revHistory = 'REVISION_HISTORY: ' + dp_.revisionRemark; } catch(e_) {}
      rq.getRange(lastRow, 1, 1, 10).setValues([[
        reqId, now, user.name, type, details, 'PENDING', '', '', 'No', revHistory
      ]]);
      SpreadsheetApp.flush();
      notifyNewRequest_(reqId, type, details, user.name, now);
      return { success:true, reqId:reqId };
    } catch(e) { return { success:false, error:e.message }; }
  } finally {
    lock.releaseLock();
  }
}

function processRequest(rowNum, action, notes) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var rq  = ss.getSheetByName('REQUESTS');
    var row = rq.getRange(rowNum, 1, 1, 10).getValues()[0];
    var rowStatus = safeStr(row[5]);
    if (rowStatus === 'APPROVED' || rowStatus === 'REJECTED') return { success: false, error: 'Already processed' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(rowNum, 6).setValue(action === 'REJECT' ? 'REJECTED' : 'APPROVED');
    rq.getRange(rowNum, 7).setValue(notes || (action==='APPROVE'?'Approved':'Rejected'));
    rq.getRange(rowNum, 8).setValue(now);
    rq.getRange(rowNum, 9).setValue('Yes');
    var sheetCreated = '';
    if (action === 'APPROVE' && safeStr(row[3]) === 'ACTIVITY_SETUP') {
      try {
        var setupPayload = JSON.parse(safeStr(row[4]));
        if (setupPayload && setupPayload.sheet && setupPayload.activities) {
          saveActivitySetup(setupPayload.sheet, setupPayload.activities, setupPayload.dept);
        } else if (setupPayload && setupPayload.activityName) {
          var ma2 = ss.getSheetByName('MASTER_ACTIVITIES');
          if (!ma2) {
            ma2 = ss.insertSheet('MASTER_ACTIVITIES');
            ma2.getRange(1, 1, 1, 7).setValues([['Dept','ActivityName','Rate','Comm','Status','RequestedBy','RequestedDate']]);
          }
          ma2.getRange(ma2.getLastRow() + 1, 1, 1, 7).setValues([[
            safeStr(setupPayload.dept), safeStr(setupPayload.activityName),
            safeNum(setupPayload.rate), safeNum(setupPayload.comm),
            'APPROVED', safeStr(row[2]), safeStr(row[1])
          ]]);
        }
      } catch(pe) { Logger.log('ACTIVITY_SETUP error: ' + pe.message); }
    }
    if (action === 'APPROVE' && safeStr(row[3]) === 'CUSTOM_PERIOD_REQUEST') {
      try {
        var cpPayload = JSON.parse(safeStr(row[4]));
        if (cpPayload && cpPayload.fromDate && cpPayload.toDate) {
          var pp = ss.getSheetByName('PAYMENT_PERIODS');
          if (pp) {
            if (pp.getLastRow() > 1) {
              var ppVals = pp.getRange(2, 1, pp.getLastRow()-1, 7).getValues();
              for (var pi = 0; pi < ppVals.length; pi++) {
                if (safeStr(ppVals[pi][6]).trim().toUpperCase() === 'OPEN')
                  pp.getRange(pi + 2, 7).setValue('CLOSED');
              }
            }
            var cpDays = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
            var cpMonths = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
            var fmtCp = function(ds) { var p=ds.split('-'); var d=new Date(+p[0],+p[1]-1,+p[2]); return cpDays[d.getDay()]+' '+d.getDate()+' '+cpMonths[d.getMonth()]+' '+d.getFullYear(); };
            var cpLabel = fmtCp(cpPayload.fromDate) + ' – ' + fmtCp(cpPayload.toDate);
            var cpId = 'W-' + cpPayload.fromDate.replace(/-/g, '');
            pp.getRange(pp.getLastRow() + 1, 1, 1, 7).setValues([[
              cpId, 'Custom', cpLabel, cpPayload.fromDate, cpPayload.toDate,
              cpPayload.reason || '', 'OPEN'
            ]]);
          }
        }
      } catch(pe) { Logger.log('CUSTOM_PERIOD_REQUEST error: ' + pe.message); }
    }
    if (action === 'APPROVE' && safeStr(row[3]) === 'NEW_ORDER') {
      try {
        var noPayload = JSON.parse(safeStr(row[4]));
        var noResult = createOrder(noPayload);
        if (noResult.success) {
          sheetCreated = noResult.bomNumber + ' → ' + noResult.artSheet;
        } else {
          rq.getRange(rowNum, 6).setValue('PENDING');
          rq.getRange(rowNum, 7).setValue('⚠ Order failed: ' + noResult.error);
          rq.getRange(rowNum, 8).setValue('');
          rq.getRange(rowNum, 9).setValue('No');
          SpreadsheetApp.flush();
          return { success:false, error: noResult.error };
        }
      } catch(pe) {
        rq.getRange(rowNum, 6).setValue('PENDING');
        rq.getRange(rowNum, 7).setValue('⚠ Order failed: ' + pe.message);
        rq.getRange(rowNum, 8).setValue('');
        rq.getRange(rowNum, 9).setValue('No');
        SpreadsheetApp.flush();
        return { success:false, error: pe.message };
      }
    }
    SpreadsheetApp.flush();
    return { success:true, action:action, sheetCreated:sheetCreated };
  } catch(e) { return { success:false, error:e.message }; }
}

function rejectRequest(reqId, remark) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can reject' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:false, error:'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === reqId) {
        var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
        rq.getRange(i + 4, 6).setValue('REJECTED');
        rq.getRange(i + 4, 7).setValue(remark || 'Rejected');
        rq.getRange(i + 4, 8).setValue(now);
        rq.getRange(i + 4, 9).setValue('Yes');
        SpreadsheetApp.flush();
        return { success:true };
      }
    }
    return { success:false, error:'Request not found: ' + reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveEditRequest(reqId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:false, error:'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues();
    var targetRow = -1, payload = null;
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === reqId && safeStr(data[i][3]) === 'EDIT_REQUEST') {
        targetRow = i + 4;
        try { payload = JSON.parse(safeStr(data[i][4])); } catch(pe) {}
        break;
      }
    }
    if (targetRow === -1) return { success:false, error:'Request not found: '+reqId };
    if (!payload) return { success:false, error:'Invalid payload' };
    var ws = ss.getSheetByName(payload.sheet);
    if (!ws) return { success:false, error:'Sheet not found: '+payload.sheet };
    ws.getRange('L'+payload.rowNum).setValue('DRAFT');
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(targetRow, 6).setValue('APPROVED');
    rq.getRange(targetRow, 7).setValue('Edit approved');
    rq.getRange(targetRow, 8).setValue(now);
    rq.getRange(targetRow, 9).setValue('Yes');
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveSetupEditRequest(reqId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success: false, error: 'Only Ayush can approve' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success: false, error: 'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 6).getValues();
    var targetRow = -1, payload = null;
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === reqId && safeStr(data[i][3]) === 'SETUP_EDIT_REQUEST') {
        targetRow = i + 4; try { payload = JSON.parse(safeStr(data[i][4])); } catch(pe) {} break;
      }
    }
    if (targetRow === -1) return { success: false, error: 'Request not found: ' + reqId };
    if (!payload || !payload.sheet || !payload.dept) return { success: false, error: 'Invalid payload' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(targetRow, 6).setValue('APPROVED');
    rq.getRange(targetRow, 7).setValue('Setup edit approved');
    rq.getRange(targetRow, 8).setValue(now);
    rq.getRange(targetRow, 9).setValue('Yes');
    var resetRow = rq.getLastRow() + 1;
    var resetSeq = String(resetRow - 3); while (resetSeq.length < 3) resetSeq = '0' + resetSeq;
    rq.getRange(resetRow, 1, 1, 10).setValues([[
      'REQ-' + resetSeq, now, user.name, 'ACTIVITY_SETUP_UNLOCKED',
      JSON.stringify({ sheet: payload.sheet, dept: payload.dept }),
      'APPROVED', 'Setup edit unlocked', now, 'Yes', ''
    ]]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
