function getCurrentWeek() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = Session.getScriptTimeZone();

  // Try CONFIG sheet for a manual override first
  try {
    var cfg = ss.getSheetByName('CONFIG');
    if (cfg && cfg.getLastRow() > 0) {
      var cfgRows = cfg.getDataRange().getValues();
      var ws = '', we = '';
      cfgRows.forEach(function(r) {
        var k = safeStr(r[0]).toUpperCase();
        if (k === 'CURRENT_WEEK_START') ws = safeStr(r[1]);
        if (k === 'CURRENT_WEEK_END')   we = safeStr(r[1]);
      });
      if (ws && we) return { weekStart:ws, weekEnd:we, weekLabel:'Week ending '+we };
    }
  } catch(e) { Logger.log('CW config: ' + e.message); }

  // Auto-calculate the current Sat–Fri window
  var today = new Date();
  var dow = today.getDay(); // 0=Sun … 6=Sat
  var daysToSat = (dow === 6) ? 0 : (dow + 1); // steps back to reach Saturday
  var sat = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysToSat);
  var fri = new Date(sat.getFullYear(), sat.getMonth(), sat.getDate() + 6);
  var weekStart = Utilities.formatDate(sat, tz, 'dd-MMM-yyyy');
  var weekEnd   = Utilities.formatDate(fri, tz, 'dd-MMM-yyyy');
  return { weekStart:weekStart, weekEnd:weekEnd, weekLabel:'Week ending '+weekEnd };
}

function setCustomWeek(startDate, endDate) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can set the week' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var tz = Session.getScriptTimeZone();
  var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
  try {
    var cfg = ss.getSheetByName('CONFIG');
    if (!cfg) {
      cfg = ss.insertSheet('CONFIG');
      cfg.getRange(1, 1, 1, 3).setValues([['KEY', 'VALUE', 'LAST_UPDATED']]);
    }
    var cfgData = cfg.getDataRange().getValues();
    var startRow = -1, endRow = -1;
    for (var i = 0; i < cfgData.length; i++) {
      var k = safeStr(cfgData[i][0]).toUpperCase();
      if (k === 'CURRENT_WEEK_START') startRow = i + 1;
      if (k === 'CURRENT_WEEK_END')   endRow   = i + 1;
    }
    var next = Math.max(cfg.getLastRow(), 1) + 1;
    if (startRow === -1) { startRow = next++; }
    if (endRow   === -1) { endRow   = next;   }
    cfg.getRange(startRow, 1, 1, 3).setValues([['CURRENT_WEEK_START', startDate, now]]);
    cfg.getRange(endRow,   1, 1, 3).setValues([['CURRENT_WEEK_END',   endDate,   now]]);
    SpreadsheetApp.flush();
    return { success:true, weekStart:startDate, weekEnd:endDate };
  } catch(e) { return { success:false, error:e.message }; }
}

function getDashboardData(ss) {
  if (!ss) {
    try {
      var _cache = CacheService.getScriptCache();
      var _cached = _cache.get('dashboardData_' + CONFIG.ENV);
      if (_cached) return JSON.parse(_cached);
    } catch(ce) {}
  }
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var weeklyPayout = 0, approvalStatus = '', weekEnding = '';
  try {
    var wk = getCurrentWeek();
    weekEnding = wk ? (wk.weekLabel || '') : '';
  } catch(we) {}

  // Compute weeklyPayout from PAYMENT_HISTORY (new job card flow)
  try {
    var ph2 = ss.getSheetByName('PAYMENT_HISTORY');
    if (ph2 && ph2.getLastRow() > 1) {
      var phRows = ph2.getRange(2, 1, ph2.getLastRow()-1, 12).getValues();
      var pendingAmt = 0, approvedAmt = 0;
      phRows.forEach(function(r) {
        var approvedBy = safeStr(r[6]).trim();
        var amt = safeNum(r[5]);
        if (!approvedBy) {
          pendingAmt += amt;
        } else if (approvedBy.indexOf('REJECTED:') !== 0) {
          approvedAmt += amt;
        }
      });
      weeklyPayout = approvedAmt;
      approvalStatus = pendingAmt > 0 ?
        'pending:' + pendingAmt.toFixed(2) : 'all_approved';
    }
  } catch(ph2Err) { Logger.log('payout err: ' + ph2Err.message); }

  var orders = [];
  var redCount = 0, completeCount = 0;
  var oiBomMap = {};
  var oiArticleMap = {};
  try {
    var oiSd = ss.getSheetByName('ORDER_INDEX');
    if (oiSd && oiSd.getLastRow() > 3)
      oiSd.getRange(4, 1, oiSd.getLastRow()-3, 3).getValues().forEach(function(r) {
        var sn = safeStr(r[1]);
        if (sn) {
          oiBomMap[sn] = safeStr(r[0]);
          oiArticleMap[sn] = safeStr(r[2]);
        }
      });
  } catch(e) {}
  try {
    var ot = ss.getSheetByName('ORDER_TRACKER');
    if (ot && ot.getLastRow() > 3) {
      var otData = ot.getRange(4, 1, ot.getLastRow()-3, 9).getValues();
      otData.forEach(function(r) {
        if (!r[0]) return;
        var oQty   = safeNum(r[3]);
        var cumul  = safeNum(r[6]);
        var status = safeStr(r[8]);
        orders.push({
          sheet:r[0],
          article: safeStr(r[1]) || oiArticleMap[safeStr(r[0])] || '',
          customer:safeStr(r[2]),
          orderQty:oQty, prior:safeNum(r[4]),
          thisWeek:safeNum(r[5]), cumul:cumul,
          remaining:safeNum(r[7]), status:status, bom:oiBomMap[safeStr(r[0])]||''
        });
        if (oQty > 0 && cumul > oQty)   redCount++;
        if (oQty > 0 && cumul === oQty) completeCount++;
      });
    }
  } catch(e) { Logger.log('OT error: ' + e.message); }

  var pendingCount = 0;
  try {
    var rq = ss.getSheetByName('REQUESTS');
    if (rq && rq.getLastRow() > 3) {
      var rqData = rq.getRange(4, 6, rq.getLastRow()-3, 1).getValues();
      rqData.forEach(function(r){
        if (safeStr(r[0]).toUpperCase() === 'PENDING') pendingCount++;
      });
    }
  } catch(e) { Logger.log('RQ error: ' + e.message); }

  var contractorSummary = [];
  try {
    var pmMapD = {};
    try {
      var mcD = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mcD && mcD.getLastRow() > 3)
        mcD.getRange(4, 2, mcD.getLastRow()-3, 2).getValues().forEach(function(r){
          if (r[0]) pmMapD[safeStr(r[0])] = safeStr(r[1]) || 'Cash';
        });
    } catch(e) {}
    var curPeriodId = '';
    try {
      var ppD = ss.getSheetByName('PAYMENT_PERIODS');
      if (ppD && ppD.getLastRow() > 1) {
        var ppDV = ppD.getRange(2, 1, ppD.getLastRow()-1, 7).getValues();
        var openIds = [];
        ppDV.forEach(function(r){ if(safeStr(r[6]).trim().toUpperCase()==='OPEN') openIds.push(safeStr(r[0])); });
        openIds.sort();
        if (openIds.length) curPeriodId = openIds[0];
      }
    } catch(e) {}
    var csMap = {};
    ss.getSheets().filter(isArtSheet).forEach(function(ws) {
      try {
        ws.getRange(5, 1, 45, 12).getValues().forEach(function(r) {
          var ctr = safeStr(r[2]);
          var qty = safeNum(r[3]);
          if (!ctr || !qty) return;
          var st = safeStr(r[11]).toUpperCase();
          if (st !== 'SUBMITTED' && st !== 'APPROVED') return;
          if (curPeriodId && safeStr(r[10]) !== curPeriodId) return;
          var total = safeNum(r[8]);
          if (!csMap[ctr]) csMap[ctr] = {name:ctr, qty:0, amount:0, method:pmMapD[ctr]||'Cash'};
          csMap[ctr].qty += qty;
          csMap[ctr].amount += total;
        });
      } catch(e) {}
    });
    contractorSummary = Object.keys(csMap).map(function(k){return csMap[k];}).sort(function(a,b){return b.amount-a.amount;});
  } catch(e) {}

  try {
    var maActDeptMap = {};
    try {
      var maS2 = ss.getSheetByName('MASTER_ACTIVITIES');
      if (maS2 && maS2.getLastRow() > 1)
        maS2.getRange(2, 1, maS2.getLastRow()-1, 2).getValues().forEach(function(r){
          if (r[1]) maActDeptMap[safeStr(r[1])] = safeStr(r[0]);
        });
    } catch(e) {}
    var curPidOrders = '';
    try {
      var ppO = ss.getSheetByName('PAYMENT_PERIODS');
      if (ppO && ppO.getLastRow() > 1) {
        var openIdsO = [];
        ppO.getRange(2, 1, ppO.getLastRow()-1, 7).getValues().forEach(function(r){
          if (safeStr(r[6]).trim().toUpperCase() === 'OPEN') openIdsO.push(safeStr(r[0]));
        });
        openIdsO.sort();
        if (openIdsO.length) curPidOrders = openIdsO[0];
      }
    } catch(e) {}
    var sheetIdx = {};
    orders.forEach(function(o, i){ sheetIdx[o.sheet] = i; });
    ss.getSheets().filter(isArtSheet).forEach(function(ws) {
      var sn = ws.getName();
      if (!(sn in sheetIdx)) return;
      var idx = sheetIdx[sn];
      var totalPaid = 0, thisWeekQty = 0, deptBkMap = {};
      try {
        ws.getRange(5, 1, 45, 12).getValues().forEach(function(r) {
          if (!safeStr(r[1]).trim() || safeNum(r[0]) <= 0) return;
          var st = safeStr(r[11]).toUpperCase();
          var qty = safeNum(r[3]);
          if (!qty) return;
          var dept = maActDeptMap[safeStr(r[1]).trim()] || 'other';
          if (!deptBkMap[dept]) deptBkMap[dept] = {dept:dept, thisWeek:0, paid:0};
          if (st === 'APPROVED') { totalPaid += qty; deptBkMap[dept].paid += qty; }
          if (curPidOrders && safeStr(r[10]) === curPidOrders && (st === 'SUBMITTED' || st === 'DRAFT')) {
            thisWeekQty += qty; deptBkMap[dept].thisWeek += qty;
          }
        });
      } catch(e) {}
      orders[idx].totalPaid = totalPaid;
      orders[idx].thisWeekQty = thisWeekQty;
      orders[idx].deptBreakdown = Object.keys(deptBkMap).map(function(k){ return deptBkMap[k]; });
    });
  } catch(e) {}

  var periodList = [];
  try {
    var pidMap = {};
    ss.getSheets().filter(isArtSheet).forEach(function(ws) {
      try {
        ws.getRange(5, 1, 45, 12).getValues().forEach(function(r) {
          var st = safeStr(r[11]).toUpperCase();
          if (st !== 'SUBMITTED' && st !== 'APPROVED') return;
          var pid = safeStr(r[10]);
          if (!pid) return;
          var total = safeNum(r[8]);
          if (!total) return;
          if (!pidMap[pid]) pidMap[pid] = { periodId:pid, total:0, submitted:0, approved:0 };
          pidMap[pid].total += total;
          if (st === 'SUBMITTED') pidMap[pid].submitted += total;
          else pidMap[pid].approved += total;
        });
      } catch(e) {}
    });
    periodList = Object.keys(pidMap).sort().map(function(k){ return pidMap[k]; });
  } catch(e) {}

  // CEO dashboard: real production pipeline (outstanding pairs per stage) and
  // delivery-risk (overdue open job cards), both from the job-card model.
  var MOVEMENT_STAGE_D = {
    'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter',
    'Upper IN':'lasting','Lasting IN':'lasting','Packing IN':'finishing','Dispatch IN':'dispatch'
  };
  var pipeline = { cutting:0, prep:0, fitter:0, lasting:0, finishing:0, dispatch:0 };
  var deliveryRisk = [];
  try {
    var _jcsD = getJobCards({});
    if (Array.isArray(_jcsD)) {
      var _todayD = new Date(); _todayD.setHours(0,0,0,0);
      _jcsD.forEach(function(jc) {
        var st = safeStr(jc.status).toUpperCase();
        if (st !== 'ISSUED' && st !== 'PARTIAL') return;
        var bal = safeNum(jc.pairsIssued) - safeNum(jc.pairsReceived);
        if (bal <= 0) return;
        var stage = MOVEMENT_STAGE_D[safeStr(jc.movement).trim()];
        if (stage) pipeline[stage] += bal;
        if (jc.expectedReturn) {
          var er = new Date(jc.expectedReturn);
          if (!isNaN(er.getTime())) {
            er.setHours(0,0,0,0);
            if (er < _todayD) deliveryRisk.push({
              orderRef: safeStr(jc.orderRef), movement: safeStr(jc.movement),
              balance: bal, daysOverdue: Math.floor((_todayD - er) / 86400000)
            });
          }
        }
      });
    }
  } catch(e) {}
  deliveryRisk.sort(function(a,b){ return b.daysOverdue - a.daysOverdue; });

  var _dashResult = {
    weeklyPayout:weeklyPayout, approvalStatus:approvalStatus,
    weekEnding:weekEnding, orders:orders, redCount:redCount,
    completeCount:completeCount,
    pendingCount:pendingCount, totalOrders:orders.length,
    contractorSummary:contractorSummary, periodList:periodList,
    pipeline:pipeline, deliveryRisk:deliveryRisk
  };
  try {
    CacheService.getScriptCache()
      .put('dashboardData_' + CONFIG.ENV, JSON.stringify(_dashResult), 300);
  } catch(ce) {}
  return _dashResult;
}

function ensureCurrentPeriod() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = Session.getScriptTimeZone();
    var today = new Date();
    var dow = today.getDay();
    var daysToSat = (dow === 6) ? 0 : (dow + 1);
    var sat = new Date(today.getFullYear(), today.getMonth(), today.getDate() - daysToSat);
    var fri = new Date(sat.getFullYear(), sat.getMonth(), sat.getDate() + 6);
    var periodId = 'W-' + Utilities.formatDate(sat, tz, 'yyyyMMdd');
    var weekStart = Utilities.formatDate(sat, tz, 'dd-MMM-yyyy');
    var weekEnd   = Utilities.formatDate(fri, tz, 'dd-MMM-yyyy');
    var weekLabel = 'Week ending ' + weekEnd;
    var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
    var pp = ss.getSheetByName('PAYMENT_PERIODS');
    if (!pp) {
      pp = ss.insertSheet('PAYMENT_PERIODS');
      pp.getRange(1, 1, 1, 10).setValues([['PeriodID','Type','Label','StartDate','EndDate','Reason','Status','SubmissionRow','ApprovedBy','CreatedAt']]);
    }
    if (pp.getLastRow() > 1) {
      var existing = pp.getRange(2, 1, pp.getLastRow()-1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (safeStr(existing[i][0]) === periodId) return;
      }
    }
    pp.getRange(pp.getLastRow() + 1, 1, 1, 10).setValues([[
      periodId, 'Auto', weekLabel, weekStart, weekEnd, '', 'OPEN', '', '', now
    ]]);
    SpreadsheetApp.flush();
  } catch(e) { Logger.log('ensureCurrentPeriod error: ' + e.message); }
}

// Single source of truth for "the current pay period". Guarantees this week's
// PAYMENT_PERIODS row exists, then returns the first OPEN period ID. Mirrors the
// resolution the Job Card / WIP path uses so entry and payment can never drift.
function resolveOpenPeriodId() {
  ensureCurrentPeriod();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var pp = ss.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var v = pp.getRange(2, 1, pp.getLastRow()-1, 7).getValues();
      var open = [];
      v.forEach(function(r){ if (safeStr(r[6]).trim().toUpperCase() === 'OPEN') open.push(safeStr(r[0])); });
      open.sort();
      if (open.length) return open[0];
    }
  } catch(e) { Logger.log('resolveOpenPeriodId: ' + e.message); }
  return 'W-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

function saveEntry(sheetName, row, contractor, qty, conveyance, remarks, rate, comm, periodId) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found' };
    var prevQty = safeNum(ws.getRange('D'+row).getValue());
    if (prevQty > 0 && (qty||0) !== prevQty) {
      try {
        var plog = ss.getSheetByName('PAYMENT_LOG');
        if (!plog) { plog = ss.insertSheet('PAYMENT_LOG'); plog.getRange(1,1,1,7).setValues([['Timestamp','User','Sheet','Row','Activity','Qty','Status']]); }
        plog.appendRow([Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm'), user.name, sheetName, row, safeStr(ws.getRange('B'+row).getValue()), prevQty, 'VOIDED']);
      } catch(le) {}
    }
    if (periodId) {
      try {
        var vActName = safeStr(ws.getRange('B'+row).getValue());
        if (vActName) {
          ws.getRange(5, 1, 45, 12).getValues().forEach(function(r, vi) {
            var vrn = vi + 5;
            if (vrn === row) return;
            var vst = safeStr(r[11]).toUpperCase();
            if (safeStr(r[1]) === vActName && safeStr(r[10]) === periodId && vst !== 'VOIDED')
              ws.getRange('L'+vrn).setValue('VOIDED');
          });
        }
      } catch(ve) {}
    }
    if (contractor) ws.getRange('C'+row).setValue(contractor);
    ws.getRange('D'+row).setValue(qty||0);
    if (rate)  ws.getRange('E'+row).setValue(rate);
    if (comm !== null && comm !== undefined) ws.getRange('F'+row).setValue(comm);
    if (conveyance) ws.getRange('H'+row).setValue(conveyance);
    if (remarks)    ws.getRange('J'+row).setValue(remarks);
    if (periodId)   ws.getRange('K'+row).setValue(periodId);
    ws.getRange('L'+row).setValue('DRAFT');
    SpreadsheetApp.flush();
    return { success:true,
      lotStatus: safeStr(ws.getRange('M7').getValue()),
      thisWeek:  safeNum(ws.getRange('M4').getValue()),
      remaining: safeNum(ws.getRange('M6').getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}

function clearEntry(sheetName, rowNum, activityName) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found' };
    var prevQty = safeNum(ws.getRange('D'+rowNum).getValue());
    if (prevQty > 0) {
      try {
        var plog = ss.getSheetByName('PAYMENT_LOG');
        if (!plog) { plog = ss.insertSheet('PAYMENT_LOG'); plog.getRange(1,1,1,7).setValues([['Timestamp','User','Sheet','Row','Activity','Qty','Status']]); }
        plog.appendRow([Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy HH:mm'), user.name, sheetName, rowNum, activityName, prevQty, 'VOIDED']);
      } catch(le) {}
    }
    ws.getRange('C'+rowNum).clearContent();
    ws.getRange('D'+rowNum).clearContent();
    ws.getRange('H'+rowNum).clearContent();
    ws.getRange('J'+rowNum).clearContent();
    ws.getRange('K'+rowNum).clearContent();
    ws.getRange('L'+rowNum).clearContent();
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function submitArticleEntries(sheetName, periodId) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found' };
    var data = ws.getRange(5, 1, 45, 12).getValues();
    var submittedActs = [];
    data.forEach(function(r, i) {
      if (!safeStr(r[1]).trim() || safeNum(r[0]) <= 0) return;
      var st = safeStr(r[11]).toUpperCase();
      if (st === 'DRAFT' || (safeNum(r[3]) > 0 && st !== 'SUBMITTED')) {
        ws.getRange('L'+(i+5)).setValue('SUBMITTED');
        submittedActs.push(safeStr(r[1]));
      }
    });
    if (!submittedActs.length) return { success:false, error:'No saved entries to submit' };
    SpreadsheetApp.flush();
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq) return { success:true, count:submittedActs.length };
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var article = safeStr(ws.getRange('B2').getValue());
    var psDetails = JSON.stringify({sheet:sheetName, article:article, periodId:periodId||'', count:submittedActs.length});
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, 'PAYMENT_SUBMISSION', psDetails, 'PENDING', '', '', 'No', ''
    ]]);
    SpreadsheetApp.flush();
    notifyNewRequest_(reqId, 'PAYMENT_SUBMISSION', psDetails, user.name, now);
    return { success:true, count:submittedActs.length, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function addContinuationRow(sheetName, activityName, rate, comm, contractor, qty, conveyance, remarks, periodId) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found' };
    var data = ws.getRange(5, 1, 45, 2).getValues();
    var emptyRow = -1, maxSno = 0;
    for (var i = 0; i < data.length; i++) {
      var sno = safeNum(data[i][0]);
      if (sno > maxSno) maxSno = sno;
      if (emptyRow === -1 && !safeStr(data[i][1]).trim() && !sno) emptyRow = i + 5;
    }
    if (emptyRow === -1) return { success:false, error:'No empty rows in sheet' };
    ws.getRange(emptyRow, 1).setValue(maxSno + 1);
    ws.getRange(emptyRow, 2).setValue(activityName);
    if (rate)  ws.getRange(emptyRow, 5).setValue(rate);
    if (comm !== null && comm !== undefined) ws.getRange(emptyRow, 6).setValue(comm);
    ws.getRange('G'+emptyRow).setFormula('=IF(D'+emptyRow+'="",0,D'+emptyRow+'*F'+emptyRow+')');
    ws.getRange('I'+emptyRow).setFormula('=IF(D'+emptyRow+'="",0,(D'+emptyRow+'*E'+emptyRow+')+G'+emptyRow+'+IF(H'+emptyRow+'="",0,H'+emptyRow+'))');
    if (contractor) ws.getRange(emptyRow, 3).setValue(contractor);
    ws.getRange(emptyRow, 4).setValue(qty||0);
    if (conveyance) ws.getRange(emptyRow, 8).setValue(conveyance);
    if (remarks)    ws.getRange(emptyRow, 10).setValue(remarks);
    if (periodId)   ws.getRange('K'+emptyRow).setValue(periodId);
    ws.getRange('L'+emptyRow).setValue('DRAFT');
    SpreadsheetApp.flush();
    return { success:true, row:emptyRow,
      lotStatus: safeStr(ws.getRange('M7').getValue()),
      thisWeek:  safeNum(ws.getRange('M4').getValue()),
      remaining: safeNum(ws.getRange('M6').getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveWeek(initials) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var wm  = ss.getSheetByName('WEEKLY PAYMENT MASTER');
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    wm.getRange('B57').setValue(initials + ' — ' + now);
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPaymentSubmissions() {
  var user = getUserInfo();
  if (user.role !== 'admin') return { submissions:[], pmMap:{} };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var oiBomPS = {};
    try {
      var oiPS = ss.getSheetByName('ORDER_INDEX');
      if (oiPS && oiPS.getLastRow() > 3)
        oiPS.getRange(4, 1, oiPS.getLastRow()-3, 2).getValues().forEach(function(r) {
          var sn = safeStr(r[1]); if (sn) oiBomPS[sn] = safeStr(r[0]);
        });
    } catch(e) {}
    var rq = ss.getSheetByName('REQUESTS');
    var pmMap = {};
    try {
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() > 3)
        mc.getRange(4, 2, mc.getLastRow()-3, 2).getValues().forEach(function(r){
          if (r[0]) pmMap[safeStr(r[0])] = safeStr(r[1]) || 'Cash';
        });
    } catch(e) {}
    var submissions = [];
    if (!rq || rq.getLastRow() < 4) return { submissions:submissions, pmMap:pmMap };
    var reqData = rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues();
    reqData.forEach(function(r, i) {
      if (safeStr(r[3]) !== 'PAYMENT_SUBMISSION' || safeStr(r[5]).toUpperCase() !== 'PENDING') return;
      try {
        var pl = JSON.parse(safeStr(r[4]));
        if (!pl || !pl.sheet) return;
        var ws = ss.getSheetByName(pl.sheet);
        if (!ws) return;
        var customer = safeStr(ws.getRange('E2').getValue());
        var actData = ws.getRange(5, 1, 45, 12).getValues();
        var activities = [], grandTotal = 0;
        actData.forEach(function(row) {
          var act = safeStr(row[1]);
          if (!act.trim() || safeNum(row[0]) <= 0) return;
          var st = safeStr(row[11]).toUpperCase();
          if (st !== 'SUBMITTED' && st !== 'APPROVED') return;
          if (pl.periodId && safeStr(row[10]) !== pl.periodId) return;
          var qty = safeNum(row[3]);
          if (!qty) return;
          var total = safeNum(row[8]);
          grandTotal += total;
          activities.push({ activity:act.trim(), contractor:safeStr(row[2]),
            qty:qty, rate:safeNum(row[4]), comm:safeNum(row[5]),
            conv:safeNum(row[7]), total:total });
        });
        submissions.push({ reqId:safeStr(r[0]), rowNum:i+4, date:safeStr(r[1]),
          submittedBy:safeStr(r[2]), sheet:pl.sheet, article:pl.article||pl.sheet,
          bom:oiBomPS[pl.sheet]||'',
          customer:customer, periodId:pl.periodId||'', activities:activities, grandTotal:grandTotal });
      } catch(e2) { Logger.log('getPaymentSubmissions: '+e2.message); }
    });
    return { submissions:submissions, pmMap:pmMap };
  } catch(e) { return { submissions:[], pmMap:{} }; }
}

function approvePaymentSubmission(reqId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:false, error:'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues();
    var targetRow = -1, payload = null;
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === reqId && safeStr(data[i][3]) === 'PAYMENT_SUBMISSION') {
        targetRow = i + 4; try { payload = JSON.parse(safeStr(data[i][4])); } catch(pe) {} break;
      }
    }
    if (targetRow === -1) return { success:false, error:'Request not found' };
    if (!payload || !payload.sheet) return { success:false, error:'Invalid payload' };
    var ws = ss.getSheetByName(payload.sheet);
    if (!ws) return { success:false, error:'Sheet not found: '+payload.sheet };
    var article = safeStr(ws.getRange('B2').getValue());
    var customer = safeStr(ws.getRange('E2').getValue());
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph) { ph = ss.insertSheet('PAYMENT_HISTORY'); ph.getRange(1,1,1,8).setValues([['WeekEnding','Article','Customer','Contractor','Qty','Amount','ApprovedBy','Date']]); }
    var actData = ws.getRange(5, 1, 45, 12).getValues();
    var approvedCount = 0;
    actData.forEach(function(r, i) {
      if (!safeStr(r[1]).trim() || safeNum(r[0]) <= 0) return;
      if (payload.periodId && safeStr(r[10]) !== payload.periodId) return;
      if (safeStr(r[11]).toUpperCase() !== 'SUBMITTED') return;
      ws.getRange('L'+(i+5)).setValue('APPROVED');
      var qty = safeNum(r[3]), total = safeNum(r[8]);
      if (qty > 0 && total > 0)
        ph.appendRow([payload.periodId||now, article, customer, safeStr(r[2]), qty, total, user.name+' — '+now, new Date()]);
      approvedCount++;
    });
    rq.getRange(targetRow, 6).setValue('APPROVED');
    rq.getRange(targetRow, 7).setValue('Payment approved');
    rq.getRange(targetRow, 8).setValue(now);
    rq.getRange(targetRow, 9).setValue('Yes');
    SpreadsheetApp.flush();
    return { success:true, count:approvedCount };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPaymentPeriods() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 2) return { success:true, records:[] };
    var tz = Session.getScriptTimeZone();
    var data = ph.getRange(2, 1, ph.getLastRow()-1, 8).getValues();
    var records = [];
    data.forEach(function(r) {
      var periodId = safeStr(r[0]);
      var amount = safeNum(r[5]);
      if (!periodId || !amount) return;
      var dv = r[7];
      records.push({ periodId:periodId, article:safeStr(r[1]), customer:safeStr(r[2]),
        contractor:safeStr(r[3]), qty:safeNum(r[4]), amount:amount,
        approvedBy:safeStr(r[6]),
        date:dv instanceof Date ? Utilities.formatDate(dv,tz,'dd-MMM-yyyy') : safeStr(dv) });
    });
    return { success:true, records:records };
  } catch(e) { return { success:false, error:e.message, records:[] }; }
}

function getPaymentHistory() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 2) return { success:true, records:[] };
    var oiBomPH = {};
    try {
      var oiPH = ss.getSheetByName('ORDER_INDEX');
      if (oiPH && oiPH.getLastRow() > 3)
        oiPH.getRange(4, 1, oiPH.getLastRow()-3, 2).getValues().forEach(function(r) {
          var sn = safeStr(r[1]); if (sn) oiBomPH[sn] = safeStr(r[0]);
        });
    } catch(e) {}
    var tz = Session.getScriptTimeZone();
    var data = ph.getRange(2, 1, ph.getLastRow()-1, 8).getValues();
    var records = [];
    data.forEach(function(r) {
      var periodId = safeStr(r[0]);
      if (!periodId) return;
      var dv = r[7];
      var sht = safeStr(r[1]);
      records.push({
        periodId: periodId,
        sheet: sht,
        bom: oiBomPH[sht] || '',
        qty: safeNum(r[4]),
        amount: safeNum(r[5]),
        approvedBy: safeStr(r[6]),
        date: dv instanceof Date ? Utilities.formatDate(dv, tz, 'dd-MMM-yyyy') : safeStr(dv)
      });
    });
    return { success:true, records:records };
  } catch(e) { return { success:false, error:e.message, records:[] }; }
}

// ── PAYMENT — Phase 5.6a ─────────────────────────────────────────────────────

function getCompletedUnpaidJobCards() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // contractor name lookup
    var ctrNameById = {};
    try {
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() >= 4) {
        mc.getRange(4, 1, mc.getLastRow()-3, 2).getValues().forEach(function(r) {
          var id = safeStr(r[0]).trim(); if (id) ctrNameById[id] = safeStr(r[1]).trim();
        });
      }
    } catch(e) {}

    // ORDER_INDEX lookup: artSheet → {article, color, customer}
    var orderInfo = {};
    try {
      var oi = ss.getSheetByName('ORDER_INDEX');
      if (oi && oi.getLastRow() >= 4) {
        oi.getRange(4, 1, oi.getLastRow()-3, 5).getValues().forEach(function(r) {
          var sh = safeStr(r[1]).trim();
          if (sh) orderInfo[sh] = { article: safeStr(r[2]).trim(), color: safeStr(r[3]).trim(), customer: safeStr(r[4]).trim() };
        });
      }
    } catch(e) {}

    var MOVEMENT_DEPT_KEY = {
      'Cutting IN':     'cutting',
      'Preparation IN': 'prep',
      'Fitter IN':      'fitter',
      'Upper IN':       'lasting',
      'Lasting IN':     'lasting',
      'Packing IN':     'finishing',
      'Dispatch IN':    'dispatch'
    };

    var paidMapC = _paidPairsMap(ss);
    var ws = ensureJobCardsSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow()-1, 17).getValues();
    var result = [];

    rows.forEach(function(r) {
      if (safeStr(r[0]).trim() === '') return;
      if (safeStr(r[13]).trim() !== 'COMPLETE') return;

      var jobCardId     = safeStr(r[0]).trim();
      var orderRef      = safeStr(r[1]).trim();
      var workOrder     = safeStr(r[2]).trim();
      var store         = safeStr(r[3]).trim();
      var movement      = safeStr(r[4]).trim();
      var contractorId  = safeStr(r[5]).trim();
      var pairsIssued   = safeNum(r[6]);
      var pairsReceived = safeNum(r[7]);
      var issuedAt      = safeStr(r[10]).trim();
      var expectedReturn= safeStr(r[11]).trim();
      var deptKey       = MOVEMENT_DEPT_KEY[movement] || '';
      var oiEntry       = orderInfo[orderRef] || {};

      var activities = [], ratePerPair = 0;
      try {
        var actRes = getApprovedActivitiesForArticle(orderRef);
        if (actRes && actRes.success && actRes.activities) {
          activities = deptKey
            ? actRes.activities.filter(function(a) { return safeStr(a.dept).toLowerCase().trim().indexOf(deptKey) >= 0; })
            : actRes.activities;
          activities.forEach(function(a) { ratePerPair += safeNum(a.rate) + safeNum(a.comm); });
        }
      } catch(ae) {}

      // Per-contractor payment lines. New department cards carry ASSIGNMENTS
      // (one line per contractor, paid only their activities' rate); legacy
      // single-contractor cards fall back to the whole department rate.
      var assignments = [];
      try { assignments = JSON.parse(safeStr(r[16])) || []; } catch(e) {}
      // Per-contractor payable lines, net of any advances already paid on this card.
      var lines = _cardContractorLines(r, ctrNameById, paidMapC, ss).filter(function(l){ return l.payablePairs > 0; });
      if (!lines.length) return;   // complete card already fully paid via advances
      var cardTotal = lines.reduce(function(s, l){ return s + l.amount; }, 0);

      result.push({
        jobCardId:      jobCardId,
        orderRef:       orderRef,
        workOrder:      workOrder,
        store:          store,
        movement:       movement,
        contractorId:   contractorId,
        contractorName: ctrNameById[contractorId] || contractorId,
        pairsIssued:    pairsIssued,
        pairsReceived:  pairsReceived,
        department:     deptKey,
        activities:     activities,
        assignments:    Array.isArray(assignments) ? assignments : [],
        lines:          lines,
        ratePerPair:    ratePerPair,
        totalAmount:    cardTotal,
        article:        oiEntry.article  || '',
        color:          oiEntry.color    || '',
        customer:       oiEntry.customer || '',
        issuedAt:       issuedAt,
        expectedReturn: expectedReturn,
        status:         'COMPLETE'
      });
    });

    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

function submitJobCardPayment(data) {
  var contractorId = safeStr(data.contractorId || '').trim();
  var jobCardIds   = Array.isArray(data.jobCardIds) ? data.jobCardIds : [];
  var periodId     = safeStr(data.periodId     || '').trim();
  var notes        = safeStr(data.notes        || '').trim();

  if (!contractorId)      return { success: false, error: 'contractorId is required' };
  if (!jobCardIds.length) return { success: false, error: 'jobCardIds must not be empty' };
  if (!periodId)          return { success: false, error: 'periodId is required' };

  var MOVEMENT_DEPT_KEY = {
    'Cutting IN':     'cutting',
    'Preparation IN': 'prep',
    'Fitter IN':      'fitter',
    'Upper IN':       'lasting',
    'Lasting IN':     'lasting',
    'Packing IN':     'finishing',
    'Dispatch IN':    'dispatch'
  };

  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // contractor name
    var contractorName = contractorId;
    try {
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() >= 4) {
        var mcRows = mc.getRange(4, 1, mc.getLastRow()-3, 2).getValues();
        for (var mi = 0; mi < mcRows.length; mi++) {
          if (safeStr(mcRows[mi][0]).trim() === contractorId) {
            contractorName = safeStr(mcRows[mi][1]).trim() || contractorId; break;
          }
        }
      }
    } catch(e) {}

    // ORDER_INDEX: artSheet → {customer}
    var orderInfo = {};
    try {
      var oi2 = ss.getSheetByName('ORDER_INDEX');
      if (oi2 && oi2.getLastRow() >= 4) {
        oi2.getRange(4, 1, oi2.getLastRow()-3, 5).getValues().forEach(function(r) {
          var sh = safeStr(r[1]).trim();
          if (sh) orderInfo[sh] = { customer: safeStr(r[4]).trim() };
        });
      }
    } catch(e) {}

    // index all JOB_CARDS by ID
    var jcWs = ensureJobCardsSheet();
    var jcAllRows = jcWs.getLastRow() > 1 ? jcWs.getRange(2, 1, jcWs.getLastRow()-1, 15).getValues() : [];
    var jcIndexById = {};
    jcAllRows.forEach(function(r, i) {
      var id = safeStr(r[0]).trim(); if (id) jcIndexById[id] = { sheetRow: i + 2, data: r };
    });

    // activities cache
    var actCache = {};
    function deptActs(orderRef, deptKey) {
      if (!actCache[orderRef]) {
        try {
          var res = getApprovedActivitiesForArticle(orderRef);
          actCache[orderRef] = (res && res.success && res.activities) ? res.activities : [];
        } catch(e) { actCache[orderRef] = []; }
      }
      if (!deptKey) return actCache[orderRef];
      return actCache[orderRef].filter(function(a) {
        return safeStr(a.dept).toLowerCase().trim().indexOf(deptKey) >= 0;
      });
    }

    // PAYMENT_HISTORY — get or create sheet, generate unique PAYMENT_ID
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph) {
      ph = ss.insertSheet('PAYMENT_HISTORY');
      ph.getRange(1, 1, 1, 12).setValues([[
        'PeriodID','Article','Customer','Contractor','Qty','Amount',
        'ApprovedBy','Date','Contractor_ID','Job_Card_Ref','Department','Payment_ID'
      ]]);
    }
    // count distinct existing PAY-IDs for sequence
    var existingPayIds = {};
    if (ph.getLastRow() > 1) {
      ph.getRange(2, 12, ph.getLastRow()-1, 1).getValues().forEach(function(r) {
        var pid = safeStr(r[0]).trim(); if (pid) existingPayIds[pid] = true;
      });
    }
    var payYear  = new Date().getFullYear();
    var paySeq   = Object.keys(existingPayIds).length + 1;
    var paySeqStr = String(paySeq); while (paySeqStr.length < 3) paySeqStr = '0' + paySeqStr;
    var PAYMENT_ID = 'PAY-' + payYear + '-' + paySeqStr;

    // process each job card
    var totalPairs = 0, totalAmount = 0, written = [];

    jobCardIds.forEach(function(jcId) {
      var entry = jcIndexById[safeStr(jcId).trim()];
      if (!entry) return;
      var r = entry.data;
      if (safeStr(r[13]).trim() !== 'COMPLETE')       return;
      if (safeStr(r[5]).trim()  !== contractorId)     return;

      var orderRef = safeStr(r[1]).trim();
      var movement = safeStr(r[4]).trim();
      var pairs    = safeNum(r[6]);
      var deptKey  = MOVEMENT_DEPT_KEY[movement] || '';
      var customer = (orderInfo[orderRef] || {}).customer || '';

      var acts = deptActs(orderRef, deptKey);
      var ratePerPair = 0;
      acts.forEach(function(a) { ratePerPair += safeNum(a.rate) + safeNum(a.comm); });
      var pairsRcvd = safeNum(r[7]);
      var amount = ratePerPair * pairsRcvd;

      ph.appendRow([
        periodId, orderRef, customer, contractorName,
        pairs, amount, '', new Date(),
        contractorId, safeStr(jcId).trim(), deptKey, PAYMENT_ID
      ]);

      totalPairs  += pairs;
      totalAmount += amount;
      written.push(safeStr(jcId).trim());
    });

    // update job card statuses to PAYMENT_PENDING
    written.forEach(function(jcId) {
      var entry = jcIndexById[jcId];
      if (entry) jcWs.getRange(entry.sheetRow, 14).setValue('PAYMENT_PENDING');
    });

    SpreadsheetApp.flush();
    return { success: true, paymentId: PAYMENT_ID, totalPairs: totalPairs, totalAmount: totalAmount, jobCardCount: written.length };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// Pay a whole completed department card in one action. Writes one PAYMENT_HISTORY
// row per contractor on the card (each paid only their assigned activities' rate),
// all sharing one PAYMENT_ID, then flips the card to PAYMENT_PENDING.
// Legacy single-contractor cards (no ASSIGNMENTS) pay the full department rate.
function submitCardPayment(data) {
  var jobCardId = safeStr(data.jobCardId || '').trim();
  var periodId  = resolveOpenPeriodId();
  var notes     = safeStr(data.notes     || '').trim();
  if (!jobCardId) return { success:false, error:'jobCardId is required' };

  var MOVEMENT_DEPT_KEY = {
    'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter',
    'Upper IN':'lasting','Lasting IN':'lasting','Packing IN':'finishing','Dispatch IN':'dispatch'
  };

  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);

    var ctrNameById = {};
    try {
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() >= 4)
        mc.getRange(4, 1, mc.getLastRow()-3, 2).getValues().forEach(function(r) {
          var id = safeStr(r[0]).trim(); if (id) ctrNameById[id] = safeStr(r[1]).trim() || id;
        });
    } catch(e) {}

    var orderInfo = {};
    try {
      var oi = ss.getSheetByName('ORDER_INDEX');
      if (oi && oi.getLastRow() >= 4)
        oi.getRange(4, 1, oi.getLastRow()-3, 5).getValues().forEach(function(r) {
          var sh = safeStr(r[1]).trim(); if (sh) orderInfo[sh] = { customer: safeStr(r[4]).trim() };
        });
    } catch(e) {}

    var jcWs = ensureJobCardsSheet();
    var jcRows = jcWs.getLastRow() > 1 ? jcWs.getRange(2, 1, jcWs.getLastRow()-1, 17).getValues() : [];
    var targetRow = -1, r = null;
    for (var i = 0; i < jcRows.length; i++) {
      if (safeStr(jcRows[i][0]).trim() === jobCardId) { targetRow = i + 2; r = jcRows[i]; break; }
    }
    if (targetRow < 0)                          return { success:false, error:'Job card not found: ' + jobCardId };
    if (safeStr(r[13]).trim() !== 'COMPLETE')   return { success:false, error:'Job card ' + jobCardId + ' is not COMPLETE (status: ' + safeStr(r[13]) + ')' };

    var orderRef      = safeStr(r[1]).trim();
    var pairsCol      = safeNum(r[6]);
    var pairsReceived = safeNum(r[7]);
    var movement      = safeStr(r[4]).trim();
    var deptKey       = MOVEMENT_DEPT_KEY[movement] || '';
    var customer      = (orderInfo[orderRef] || {}).customer || '';

    var assignments = [];
    try { assignments = JSON.parse(safeStr(r[16])) || []; } catch(e) {}

    // Per-contractor amounts
    var lines = [];
    if (Array.isArray(assignments) && assignments.length) {
      var byCtr = {};
      assignments.forEach(function(a) {
        var cid = safeStr(a.contractorId).trim();
        if (!cid) return;
        if (!byCtr[cid]) byCtr[cid] = 0;
        byCtr[cid] += safeNum(a.rate) + safeNum(a.comm);
      });
      Object.keys(byCtr).forEach(function(cid) { lines.push({ contractorId:cid, ratePerPair:byCtr[cid] }); });
    } else {
      var legacyCid = safeStr(r[5]).trim();
      var rate = 0;
      try {
        var ar = getApprovedActivitiesForArticle(orderRef);
        if (ar && ar.success && Array.isArray(ar.activities))
          ar.activities.filter(function(a){ return !deptKey || safeStr(a.dept).toLowerCase().trim().indexOf(deptKey) >= 0; })
                       .forEach(function(a){ rate += safeNum(a.rate) + safeNum(a.comm); });
      } catch(e) {}
      lines.push({ contractorId:legacyCid, ratePerPair:rate });
    }
    if (!lines.length) return { success:false, error:'No contractor assignments on this card' };

    // PAYMENT_HISTORY sheet + unique PAYMENT_ID
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph) {
      ph = ss.insertSheet('PAYMENT_HISTORY');
      ph.getRange(1, 1, 1, 12).setValues([[
        'PeriodID','Article','Customer','Contractor','Qty','Amount',
        'ApprovedBy','Date','Contractor_ID','Job_Card_Ref','Department','Payment_ID'
      ]]);
    }
    var existingPayIds = {};
    if (ph.getLastRow() > 1)
      ph.getRange(2, 12, ph.getLastRow()-1, 1).getValues().forEach(function(x) {
        var pid = safeStr(x[0]).trim(); if (pid) existingPayIds[pid] = true;
      });
    var paySeqStr = String(Object.keys(existingPayIds).length + 1); while (paySeqStr.length < 3) paySeqStr = '0' + paySeqStr;
    var PAYMENT_ID = 'PAY-' + new Date().getFullYear() + '-' + paySeqStr;

    // Net out any prior advances: pay only pairs not already paid per contractor.
    var paidMap = _paidPairsMap(ss);
    var totalAmount = 0;
    lines.forEach(function(l) {
      var alreadyPaid = paidMap[jobCardId + '||' + l.contractorId] || 0;
      var payable = pairsReceived - alreadyPaid;
      if (payable <= 0) return;   // already covered by advances
      var amount = safeNum(l.ratePerPair) * payable;
      ph.appendRow([
        periodId, orderRef, customer, ctrNameById[l.contractorId] || l.contractorId,
        payable, amount, '', new Date(),
        l.contractorId, jobCardId, deptKey, PAYMENT_ID
      ]);
      totalAmount += amount;
    });

    // Flip card to PAYMENT_PENDING (STATUS is column 14)
    jcWs.getRange(targetRow, 14).setValue('PAYMENT_PENDING');
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
    try { CacheService.getScriptCache().remove('storeScreenData_' + CONFIG.ENV); } catch(ce) {}

    return { success:true, paymentId:PAYMENT_ID, jobCardId:jobCardId, contractors:lines.length,
             pairs:pairsReceived, totalAmount:totalAmount };
  } catch(e) {
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
  }
}

function getPaymentBatches(filters) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 2) return [];

    var rows = ph.getRange(2, 1, ph.getLastRow()-1, 12).getValues();
    var batchMap = {}, batchOrder = [];
    var tz = Session.getScriptTimeZone();

    rows.forEach(function(r) {
      var paymentId = safeStr(r[11]).trim();
      if (!paymentId) return;  // skip legacy rows without payment ID

      if (!batchMap[paymentId]) {
        var dv = r[7];
        var dateStr = dv instanceof Date ? Utilities.formatDate(dv, tz, 'dd-MMM-yyyy') : safeStr(dv);
        batchMap[paymentId] = {
          paymentId:      paymentId,
          contractorId:   safeStr(r[8]).trim(),
          contractorName: safeStr(r[3]).trim(),
          periodId:       safeStr(r[0]).trim(),
          approvedBy:     safeStr(r[6]).trim(),
          date:           dateStr,
          _dateMs:        dv instanceof Date ? dv.getTime() : 0,
          status:         safeStr(r[6]).trim() ? 'APPROVED' : 'PENDING',
          lines:          [],
          _seenJC:        {},
          totalPairs:     0,
          totalAmount:    0
        };
        batchOrder.push(paymentId);
      }

      var pairs  = safeNum(r[4]);
      var amount = safeNum(r[5]);
      var lineJc = safeStr(r[9]).trim();
      batchMap[paymentId].lines.push({
        jobCardId:      lineJc,
        orderRef:       safeStr(r[1]).trim(),
        customer:       safeStr(r[2]).trim(),
        contractorName: safeStr(r[3]).trim(),
        department:     safeStr(r[10]).trim(),
        pairs:          pairs,
        amount:         amount
      });
      // Count physical pairs once per job card (a department card can have several
      // contractor rows, all for the same pairs); amount always sums.
      if (!batchMap[paymentId]._seenJC[lineJc]) {
        batchMap[paymentId].totalPairs += pairs;
        batchMap[paymentId]._seenJC[lineJc] = true;
      }
      batchMap[paymentId].totalAmount += amount;
    });

    var result = batchOrder.map(function(pid) { var b = batchMap[pid]; delete b._seenJC; return b; });

    if (filters) {
      if (filters.periodId)     result = result.filter(function(b){ return b.periodId     === safeStr(filters.periodId); });
      if (filters.contractorId) result = result.filter(function(b){ return b.contractorId === safeStr(filters.contractorId); });
      if (filters.orderRef) {
        var filterOr = safeStr(filters.orderRef);
        result = result.filter(function(b){ return b.lines.some(function(l){ return l.orderRef === filterOr; }); });
      }
    }

    result.sort(function(a, b) { return b._dateMs - a._dateMs; });
    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

function approvePaymentBatch(paymentId) {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var _user = getUserInfo();
    if (_user.role !== 'admin') return { success:false, error:'Only Ayush can approve payments' };
    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var user = getUserInfo();
    var tz   = Session.getScriptTimeZone();
    var now  = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
    var approverStr = (user.name || user.email || 'Unknown') + ' — ' + now;

    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 2) return { success: false, error: 'Payment batch not found' };

    var rows = ph.getRange(2, 1, ph.getLastRow()-1, 12).getValues();
    var matchedRows = [], jobCardIds = [];
    rows.forEach(function(r, i) {
      if (safeStr(r[11]).trim() !== safeStr(paymentId).trim()) return;
      matchedRows.push(i + 2);
      var jcId = safeStr(r[9]).trim(); if (jcId) jobCardIds.push(jcId);
    });

    if (!matchedRows.length) return { success: false, error: 'Payment batch not found' };

    matchedRows.forEach(function(sheetRow) {
      ph.getRange(sheetRow, 7).setValue(approverStr);
    });

    var jcWs = ensureJobCardsSheet();
    if (jcWs.getLastRow() > 1) {
      jcWs.getRange(2, 1, jcWs.getLastRow()-1, 14).getValues().forEach(function(r, i) {
        // Only a final payment (card PAYMENT_PENDING) becomes PAID; advance batches
        // leave the card open (still ISSUED/PARTIAL).
        if (jobCardIds.indexOf(safeStr(r[0]).trim()) >= 0 && safeStr(r[13]).toUpperCase() === 'PAYMENT_PENDING')
          jcWs.getRange(i + 2, 14).setValue('PAID');
      });
    }

    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
    return { success: true, paymentId: paymentId, jobCardCount: matchedRows.length };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function rejectPaymentBatch(paymentId, reason) {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var _user = getUserInfo();
    if (_user.role !== 'admin') return { success:false, error:'Only Ayush can reject payments' };
    var ss         = SpreadsheetApp.openById(SHEET_ID);
    var reasonStr  = 'REJECTED: ' + safeStr(reason || '').trim();

    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 2) return { success: false, error: 'Payment batch not found' };

    var rows = ph.getRange(2, 1, ph.getLastRow()-1, 12).getValues();
    var matchedRows = [], jobCardIds = [];
    rows.forEach(function(r, i) {
      if (safeStr(r[11]).trim() !== safeStr(paymentId).trim()) return;
      matchedRows.push(i + 2);
      var jcId = safeStr(r[9]).trim(); if (jcId) jobCardIds.push(jcId);
    });

    if (!matchedRows.length) return { success: false, error: 'Payment batch not found' };

    matchedRows.forEach(function(sheetRow) {
      ph.getRange(sheetRow, 7).setValue(reasonStr);
    });

    var jcWs = ensureJobCardsSheet();
    if (jcWs.getLastRow() > 1) {
      jcWs.getRange(2, 1, jcWs.getLastRow()-1, 14).getValues().forEach(function(r, i) {
        // Only revert a final payment (card PAYMENT_PENDING) back to COMPLETE; a
        // rejected advance leaves the card open and frees its pairs for re-payment.
        if (jobCardIds.indexOf(safeStr(r[0]).trim()) >= 0 && safeStr(r[13]).toUpperCase() === 'PAYMENT_PENDING')
          jcWs.getRange(i + 2, 14).setValue('COMPLETE');
      });
    }

    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── PARTIAL-CARD ADVANCES — Phase 7 ──────────────────────────────────────────
// An advance is a partial payment against a still-open (ISSUED/PARTIAL) card for
// pairs already returned. It reuses PAYMENT_HISTORY + the normal approval flow;
// the final payment pays only the remaining (un-advanced) pairs, so nothing is
// double-paid. Netting is per job card + contractor. No floating balances.

// Pairs already paid (advance or final, non-rejected) per jobCard+contractor.
function _paidPairsMap(ss) {
  var map = {};
  var ph = ss.getSheetByName('PAYMENT_HISTORY');
  if (ph && ph.getLastRow() > 1) {
    ph.getRange(2, 1, ph.getLastRow()-1, 12).getValues().forEach(function(r) {
      if (safeStr(r[6]).trim().indexOf('REJECTED:') === 0) return;
      var jcId = safeStr(r[9]).trim(), cid = safeStr(r[8]).trim();
      if (!jcId || !cid) return;
      map[jcId + '||' + cid] = (map[jcId + '||' + cid] || 0) + safeNum(r[4]);
    });
  }
  return map;
}

// Per-contractor payable lines for a card, netting pairs already paid.
function _cardContractorLines(r, ctrNameById, paidMap, ss) {
  var MDK = {'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter','Upper IN':'lasting','Lasting IN':'lasting','Packing IN':'finishing','Dispatch IN':'dispatch'};
  var jobCardId = safeStr(r[0]).trim();
  var orderRef  = safeStr(r[1]).trim();
  var deptKey   = MDK[safeStr(r[4]).trim()] || '';
  var pairsReceived = safeNum(r[7]);
  var assignments = [];
  try { assignments = JSON.parse(safeStr(r[16])) || []; } catch(e) {}
  var byCtr = {};
  if (Array.isArray(assignments) && assignments.length) {
    assignments.forEach(function(a) {
      var cid = safeStr(a.contractorId).trim(); if (!cid) return;
      if (!byCtr[cid]) byCtr[cid] = { ratePerPair:0, activities:[] };
      byCtr[cid].ratePerPair += safeNum(a.rate) + safeNum(a.comm);
      byCtr[cid].activities.push(safeStr(a.activity));
    });
  } else {
    var cid = safeStr(r[5]).trim();
    var rate = 0;
    try {
      var ar = getApprovedActivitiesForArticle(orderRef, ss);
      if (ar && ar.success && Array.isArray(ar.activities))
        ar.activities.filter(function(a){ return !deptKey || safeStr(a.dept).toLowerCase().trim().indexOf(deptKey) >= 0; })
                     .forEach(function(a){ rate += safeNum(a.rate) + safeNum(a.comm); });
    } catch(e) {}
    byCtr[cid] = { ratePerPair:rate, activities:[] };
  }
  var lines = [];
  Object.keys(byCtr).forEach(function(cid) {
    var alreadyPaid = paidMap[jobCardId + '||' + cid] || 0;
    var payablePairs = pairsReceived - alreadyPaid;
    if (payablePairs < 0) payablePairs = 0;
    lines.push({
      contractorId: cid, contractorName: ctrNameById[cid] || cid,
      ratePerPair: byCtr[cid].ratePerPair, activities: byCtr[cid].activities,
      alreadyPaidPairs: alreadyPaid, payablePairs: payablePairs,
      amount: byCtr[cid].ratePerPair * payablePairs
    });
  });
  return lines;
}

function _ctrNameMap(ss) {
  var m = {};
  try {
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() >= 4)
      mc.getRange(4, 1, mc.getLastRow()-3, 2).getValues().forEach(function(r) {
        var id = safeStr(r[0]).trim(); if (id) m[id] = safeStr(r[1]).trim() || id;
      });
  } catch(e) {}
  return m;
}

function _orderInfoMap(ss) {
  var m = {};
  try {
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() >= 4)
      oi.getRange(4, 1, oi.getLastRow()-3, 5).getValues().forEach(function(r) {
        var sh = safeStr(r[1]).trim();
        if (sh) m[sh] = { article: safeStr(r[2]).trim(), color: safeStr(r[3]).trim(), customer: safeStr(r[4]).trim() };
      });
  } catch(e) {}
  return m;
}

// Open (ISSUED/PARTIAL) cards with received pairs not yet paid — advance candidates.
function getAdvanceableJobCards() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ctrNameById = _ctrNameMap(ss);
    var orderInfo   = _orderInfoMap(ss);
    var paidMap     = _paidPairsMap(ss);
    var ws = ensureJobCardsSheet(ss);
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow()-1, 17).getValues();
    var result = [];
    rows.forEach(function(r) {
      if (safeStr(r[0]).trim() === '') return;
      var status = safeStr(r[13]).toUpperCase();
      if (status !== 'ISSUED' && status !== 'PARTIAL') return;
      if (safeNum(r[7]) <= 0) return;
      var lines = _cardContractorLines(r, ctrNameById, paidMap, ss).filter(function(l){ return l.payablePairs > 0; });
      if (!lines.length) return;
      var oiEntry = orderInfo[safeStr(r[1]).trim()] || {};
      result.push({
        jobCardId: safeStr(r[0]).trim(), orderRef: safeStr(r[1]).trim(),
        store: safeStr(r[3]).trim(), movement: safeStr(r[4]).trim(),
        pairsIssued: safeNum(r[6]), pairsReceived: safeNum(r[7]),
        article: oiEntry.article || '', color: oiEntry.color || '', customer: oiEntry.customer || '',
        status: status, lines: lines, totalAmount: lines.reduce(function(s,l){ return s + l.amount; }, 0)
      });
    });
    return result;
  } catch(e) { return { success:false, error:e.message }; }
}

// Pay an advance against an open card for the pairs received-so-far (net of prior
// payments). Records ADV- rows in PAYMENT_HISTORY (pending approval); does NOT
// change the card status — the card stays open for more work.
function submitCardAdvance(data) {
  var jobCardId = safeStr(data.jobCardId || '').trim();
  var periodId  = resolveOpenPeriodId();
  if (!jobCardId) return { success:false, error:'jobCardId is required' };
  var MDK = {'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter','Upper IN':'lasting','Lasting IN':'lasting','Packing IN':'finishing','Dispatch IN':'dispatch'};
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ctrNameById = _ctrNameMap(ss);
    var orderInfo   = _orderInfoMap(ss);
    var paidMap     = _paidPairsMap(ss);
    var jcWs = ensureJobCardsSheet(ss);
    var rows = jcWs.getLastRow() > 1 ? jcWs.getRange(2, 1, jcWs.getLastRow()-1, 17).getValues() : [];
    var r = null;
    for (var i = 0; i < rows.length; i++) { if (safeStr(rows[i][0]).trim() === jobCardId) { r = rows[i]; break; } }
    if (!r) return { success:false, error:'Job card not found: ' + jobCardId };
    var status = safeStr(r[13]).toUpperCase();
    if (status !== 'ISSUED' && status !== 'PARTIAL')
      return { success:false, error:'Advances are only for open cards. ' + jobCardId + ' is ' + status + ' — use Pay Card instead.' };
    if (safeNum(r[7]) <= 0) return { success:false, error:'No pairs received yet on this card' };
    var lines = _cardContractorLines(r, ctrNameById, paidMap, ss).filter(function(l){ return l.payablePairs > 0 && l.ratePerPair > 0; });
    if (!lines.length) return { success:false, error:'Nothing to advance — received pairs already paid, or rate not set' };

    var orderRef = safeStr(r[1]).trim();
    var deptKey  = MDK[safeStr(r[4]).trim()] || '';
    var customer = (orderInfo[orderRef] || {}).customer || '';

    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph) {
      ph = ss.insertSheet('PAYMENT_HISTORY');
      ph.getRange(1, 1, 1, 12).setValues([[
        'PeriodID','Article','Customer','Contractor','Qty','Amount',
        'ApprovedBy','Date','Contractor_ID','Job_Card_Ref','Department','Payment_ID'
      ]]);
    }
    var advSeq = 0;
    if (ph.getLastRow() > 1)
      ph.getRange(2, 12, ph.getLastRow()-1, 1).getValues().forEach(function(x) {
        var m = safeStr(x[0]).match(/^ADV-\d{4}-(\d+)$/); if (m) { var n = parseInt(m[1],10); if (n > advSeq) advSeq = n; }
      });
    var seqStr = String(advSeq + 1); while (seqStr.length < 3) seqStr = '0' + seqStr;
    var ADVANCE_ID = 'ADV-' + new Date().getFullYear() + '-' + seqStr;

    var totalAmount = 0;
    lines.forEach(function(l) {
      ph.appendRow([
        periodId, orderRef, customer, l.contractorName,
        l.payablePairs, l.amount, '', new Date(),
        l.contractorId, jobCardId, deptKey, ADVANCE_ID
      ]);
      totalAmount += l.amount;
    });
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
    return { success:true, advanceId:ADVANCE_ID, jobCardId:jobCardId, contractors:lines.length, totalAmount:totalAmount };
  } catch(e) {
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
  }
}
