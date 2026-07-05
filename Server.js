/**
 * ADEES EXPORTS — FACTORY OS
 * Final Server.gs — Clean, no DASHBOARD dependency
 */

var CONFIG = {
  LIVE_SHEET_ID: '1FLPeuQFPx0nQXRy-16P2-1-e5SjDu7nLE-1ycNZ-IH0',
  DEV_SHEET_ID: '1eHnrG7IWn5PhreW1ywkdhgpzjOzYs6Y53vC4EIxwTvg',
  ENV: 'DEV'
};
var SHEET_ID = CONFIG.ENV === 'DEV' ? CONFIG.DEV_SHEET_ID : CONFIG.LIVE_SHEET_ID;

const ROLES = {
  "ayush@adeesexports.com":   "admin",
  "accounts@adeesexports.com": "accounts",
  "admin@adeesexports.com":   "store",
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('FactoryOS — Adees Exports' + (CONFIG.ENV === 'DEV' ? ' (DEV)' : ''))
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserInfo() {
  var email = Session.getActiveUser().getEmail();
  var role  = ROLES[email] || 'viewer';
  var first = email.split('@')[0];
  var name  = first.charAt(0).toUpperCase() + first.slice(1);
  return { email:email, role:role, name:name, env:CONFIG.ENV };
}

function safeNum(val) {
  if (val === null || val === undefined) return 0;
  var n = Number(val);
  return isNaN(n) ? 0 : n;
}

function safeStr(val) {
  if (val === null || val === undefined) return '';
  return String(val);
}

function isArtSheet(s) {
  var n = s.getName();
  return n.indexOf('ART-') === 0 && n !== 'ART-TEMPLATE';
}

function getAllData() {
  var result = { ok:true, dash:null, entry:null, wip:null, reqs:null, user:null };
  try {
    var user = getUserInfo();
    result.user = user;
    Logger.log('[getAllData] user=' + user.email + ' role=' + user.role);
    result.dash = getDashboardData();
    if (user.role === 'accounts') {
      result.entry = getEntryData();
      result.reqs = getPendingRequests();
    } else if (user.role === 'admin') {
      result.reqs = getPendingRequests();
    }
  } catch(e) {
    Logger.log('getAllData error: ' + e.message + ' | stack: ' + e.stack);
    result.error = e.message;
  }
  return result;
}

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

function getDashboardData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var weeklyPayout = 0, approvalStatus = '', weekEnding = '';

  var week = getCurrentWeek();
  weekEnding = week.weekLabel;

  var orders = [];
  var redCount = 0, completeCount = 0;
  var oiBomMap = {};
  try {
    var oiSd = ss.getSheetByName('ORDER_INDEX');
    if (oiSd && oiSd.getLastRow() > 3)
      oiSd.getRange(4, 1, oiSd.getLastRow()-3, 2).getValues().forEach(function(r) {
        var sn = safeStr(r[1]); if (sn) oiBomMap[sn] = safeStr(r[0]);
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
          sheet:r[0], article:safeStr(r[1]), customer:safeStr(r[2]),
          orderQty:oQty, prior:safeNum(r[4]),
          thisWeek:safeNum(r[5]), cumul:cumul,
          remaining:safeNum(r[7]), status:status, bom:oiBomMap[safeStr(r[0])]||''
        });
        if (oQty > 0 && cumul > oQty)   redCount++;
        if (oQty > 0 && cumul === oQty) completeCount++;
      });
    }
  } catch(e) { Logger.log('OT error: ' + e.message); }

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

  return {
    weeklyPayout:weeklyPayout, approvalStatus:approvalStatus,
    weekEnding:weekEnding, orders:orders, redCount:redCount,
    completeCount:completeCount, mismatches:mismatches,
    pendingCount:pendingCount, totalOrders:orders.length,
    contractorSummary:contractorSummary, periodList:periodList
  };
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

function getEntryData(periodId) {
  ensureCurrentPeriod();
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var firstOpenId = '';
  try {
    var ppE = ss.getSheetByName('PAYMENT_PERIODS');
    if (ppE && ppE.getLastRow() > 1) {
      var ppEV = ppE.getRange(2, 1, ppE.getLastRow()-1, 7).getValues();
      var openIds = [];
      ppEV.forEach(function(r){ if(safeStr(r[6]).trim().toUpperCase()==='OPEN') openIds.push(safeStr(r[0])); });
      openIds.sort();
      if(openIds.length) firstOpenId = openIds[0];
    }
  } catch(ep) {}
  var effectivePeriodId = periodId || firstOpenId;
  var oiBomEnt = {};
  try {
    var oiEnt = ss.getSheetByName('ORDER_INDEX');
    if (oiEnt && oiEnt.getLastRow() > 3)
      oiEnt.getRange(4, 1, oiEnt.getLastRow()-3, 2).getValues().forEach(function(r) {
        var sn = safeStr(r[1]); if (sn) oiBomEnt[sn] = safeStr(r[0]);
      });
  } catch(e) {}
  var artSheets = ss.getSheets().filter(isArtSheet);
  var articles = [];

  artSheets.forEach(function(ws) {
    try {
      var name     = ws.getName();
      var hdr      = ws.getRange('B2:J2').getValues()[0];
      var article  = safeStr(hdr[0]);
      var customer = safeStr(hdr[3]);
      var orderQty = safeNum(hdr[6]);
      var status   = safeStr(ws.getRange('M7').getValue());
      var thisWeek = safeNum(ws.getRange('M4').getValue());
      var remaining= safeNum(ws.getRange('M6').getValue());

      var actData = ws.getRange(5, 1, 45, 12).getValues();
      var activities = [];
      var approvedByAct = {};
      actData.forEach(function(r, i) {
        var act = safeStr(r[1]);
        if (act && act.trim() && !act.match(/^[-=]/) && safeNum(r[0]) > 0) {
          var _actSt = safeStr(r[11]).toUpperCase();
          if ((_actSt === 'APPROVED' || _actSt === 'SUBMITTED') && safeNum(r[3]) > 0)
            approvedByAct[act.trim()] = (approvedByAct[act.trim()] || 0) + safeNum(r[3]);
          var rowPeriodId = safeStr(r[10]);
          var inPeriod = rowPeriodId === effectivePeriodId;
          activities.push({
            row:i+5, activity:act.trim(),
            contractor: inPeriod ? safeStr(r[2]) : '',
            qty:        inPeriod ? safeNum(r[3]) : 0,
            rate:safeNum(r[4]), comm:safeNum(r[5]),
            total:      inPeriod ? safeNum(r[8]) : 0,
            entryStatus: inPeriod ? safeStr(r[11]) : '',
            conveyance: inPeriod ? safeNum(r[7]) : 0,
            remarks:    inPeriod ? safeStr(r[9]) : ''
          });
        }
      });
      articles.push({ sheet:name, article:article, customer:customer,
        orderQty:orderQty, status:status, thisWeek:thisWeek,
        remaining:remaining, activities:activities, _aqa:approvedByAct, bom:oiBomEnt[name]||'' });
    } catch(e) { Logger.log('ART error ' + ws.getName() + ': ' + e.message); }
  });

  try {
    var rqSetup = ss.getSheetByName('REQUESTS');
    if (rqSetup && rqSetup.getLastRow() > 3) {
      rqSetup.getRange(4, 1, rqSetup.getLastRow()-3, 6).getValues().forEach(function(rr) {
        if (safeStr(rr[3]) !== 'ACTIVITY_SETUP' || safeStr(rr[5]).toUpperCase() !== 'PENDING') return;
        try {
          var setupPl = JSON.parse(safeStr(rr[4]));
          if (setupPl && setupPl.sheet)
            for (var ai = 0; ai < articles.length; ai++)
              if (articles[ai].sheet === setupPl.sheet) articles[ai].hasPendingSetup = true;
        } catch(pe) {}
      });
    }
  } catch(e4) {}

  var contractors = [];
  try {
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() > 3) {
      mc.getRange(4, 2, mc.getLastRow()-3, 3).getValues().forEach(function(r){
        if (r[0] && safeStr(r[2]).toLowerCase() === 'active')
          contractors.push(safeStr(r[0]));
      });
    }
  } catch(e) {}

  var masterActivities = [];
  try {
    var maSheet = ss.getSheetByName('MASTER_ACTIVITIES');
    if (maSheet && maSheet.getLastRow() > 1) {
      maSheet.getRange(2, 1, maSheet.getLastRow()-1, 7).getValues().forEach(function(r, i){
        if (safeStr(r[1]) && safeStr(r[4]).toUpperCase() === 'APPROVED')
          masterActivities.push({ name:safeStr(r[1]), section:safeStr(r[0]),
            stdRate:safeNum(r[2]), comm:safeNum(r[3]),
            rowIndex: 2 + i });
      });
      try {
        var rq2 = ss.getSheetByName('REQUESTS');
        if (rq2 && rq2.getLastRow() > 3) {
          rq2.getRange(4, 1, rq2.getLastRow()-3, 6).getValues().forEach(function(rr) {
            if (safeStr(rr[3]) !== 'RATE_EDIT') return;
            var st = safeStr(rr[5]).toUpperCase();
            if (st !== 'PENDING' && st !== 'REJECTED') return;
            try {
              var pl = JSON.parse(safeStr(rr[4]));
              if (!pl || !pl.rowIndex) return;
              masterActivities.forEach(function(ma) {
                if (ma.rowIndex === pl.rowIndex) {
                  if (st === 'PENDING') { ma.hasPendingRateEdit = true; ma.pendingReqId = safeStr(rr[0]); }
                  else { ma.hasRejectedRateEdit = true; ma.rejectedRate = pl.newRate; ma.rejectedComm = pl.newComm; ma.rejectedReqId = safeStr(rr[0]); }
                }
              });
            } catch(pe) {}
          });
        }
      } catch(e2) {}
    }
  } catch(e) {}

  var oiLotMap = {};
  try {
    var oiS = ss.getSheetByName('ORDER_INDEX');
    if (oiS && oiS.getLastRow() > 3)
      oiS.getRange(4, 1, oiS.getLastRow()-3, 9).getValues().forEach(function(r) {
        var sn = safeStr(r[1]), ls = safeNum(r[8]);  // col I (index 8) = LOT QTY
        if (sn && ls) oiLotMap[sn] = ls;
      });
  } catch(e) {}

  var articleDeptMaps = {};
  try {
    var rqAS = ss.getSheetByName('REQUESTS');
    if (rqAS && rqAS.getLastRow() > 3) {
      rqAS.getRange(4, 1, rqAS.getLastRow()-3, 6).getValues().forEach(function(rr) {
        if (safeStr(rr[3]) !== 'ACTIVITY_SETUP' || safeStr(rr[5]).toUpperCase() !== 'APPROVED') return;
        try {
          var asPl = JSON.parse(safeStr(rr[4]));
          if (!asPl || !asPl.sheet) return;
          if (!articleDeptMaps[asPl.sheet]) articleDeptMaps[asPl.sheet] = {};
          var acts = asPl.activities || (asPl.activityName ? [asPl] : []);
          acts.forEach(function(a) {
            if (a.activityName)
              articleDeptMaps[asPl.sheet][safeStr(a.activityName)] = safeStr(a.dept || '');
          });
        } catch(pe) {}
      });
    }
  } catch(e) {}

  var actDeptMap = {};
  masterActivities.forEach(function(ma) { actDeptMap[ma.name] = ma.section || ''; });
  var _batchReqsData = [];
  var _batchDsData = [];
  try { var _bRq = ss.getSheetByName('REQUESTS'); if (_bRq && _bRq.getLastRow() > 3) _batchReqsData = _bRq.getRange(4, 1, _bRq.getLastRow()-3, 6).getValues(); } catch(e) {}
  try { var _bDs = ss.getSheetByName('DEPT_STATUS'); if (_bDs && _bDs.getLastRow() > 1) _batchDsData = _bDs.getRange(2, 1, _bDs.getLastRow()-1, 3).getValues(); } catch(e) {}
  var _deptStatusBatch = getDeptStatusBatch(articles.map(function(a){return a.sheet;}), _batchReqsData, _batchDsData);
  articles.forEach(function(art) {
    var perArtMap = articleDeptMaps[art.sheet] || {};
    var deptApproved = {};
    Object.keys(art._aqa || {}).forEach(function(an) {
      var d = perArtMap[an] || actDeptMap[an] || 'other';
      deptApproved[d] = (deptApproved[d] || 0) + (art._aqa[an] || 0);
    });
    var lotSize = oiLotMap[art.sheet] || art.orderQty;
    art.activities.forEach(function(ac) {
      ac.section = perArtMap[ac.activity] || actDeptMap[ac.activity] || '';
      if (lotSize > 0) {
        var d = perArtMap[ac.activity] || actDeptMap[ac.activity] || 'other';
        if (deptApproved[d] && deptApproved[d] >= lotSize) ac.deptLocked = true;
      }
    });
    var deptStatus = _deptStatusBatch[art.sheet] || {};
    var DS_KEY = {'Cutting':'cutting','Preparation':'prep','Fitter':'fitter','Lasting':'lasting','Finishing':'finish','Dispatch':'dispatch'};
    if (lotSize > 0) {
      Object.keys(deptStatus).forEach(function(dn) {
        if (deptStatus[dn] === 'APPROVED') {
          var dv = DS_KEY[dn] || dn.toLowerCase();
          if ((deptApproved[dv] || 0) >= lotSize) deptStatus[dn] = 'CAP_REACHED';
        }
      });
    }
    art.deptHasPayment = {};
    Object.keys(deptApproved).forEach(function(dk) {
      if (deptApproved[dk] > 0) art.deptHasPayment[dk] = true;
    });
    art.deptStatus = deptStatus;
    delete art._aqa;
  });

  try {
    var rqSer = ss.getSheetByName('REQUESTS');
    if (rqSer && rqSer.getLastRow() > 3) {
      var pendingEditMap = {};
      rqSer.getRange(4, 1, rqSer.getLastRow()-3, 6).getValues().forEach(function(r) {
        if (safeStr(r[3]) !== 'SETUP_EDIT_REQUEST' || safeStr(r[5]).toUpperCase() !== 'PENDING') return;
        try {
          var pl = JSON.parse(safeStr(r[4]));
          if (!pl || !pl.sheet || !pl.dept) return;
          if (!pendingEditMap[pl.sheet]) pendingEditMap[pl.sheet] = {};
          pendingEditMap[pl.sheet][pl.dept] = true;
        } catch(pe) {}
      });
      articles.forEach(function(art) { art.pendingSetupEdits = pendingEditMap[art.sheet] || {}; });
    }
  } catch(e) {}

  var week = null;
  var periods = [];
  try {
    var pp = ss.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var ppTz = Session.getScriptTimeZone();
    var fmtPpD = function(v) { return v instanceof Date ? Utilities.formatDate(v, ppTz, 'dd-MMM-yyyy') : safeStr(v); };
    var ppData = pp.getRange(2, 1, pp.getLastRow()-1, 7).getValues();
      for (var pi = 0; pi < ppData.length; pi++) {
        var ppSt=safeStr(ppData[pi][6]).trim().toUpperCase();
        if (ppSt==='OPEN'||ppSt==='CLOSED') {
          periods.push({
            periodId:  safeStr(ppData[pi][0]),
            weekLabel: safeStr(ppData[pi][2]),
            weekStart: fmtPpD(ppData[pi][3]),
            weekEnd:   fmtPpD(ppData[pi][4]),
            status:    ppSt
          });
        }
      }
      periods.sort(function(a,b){return a.periodId<b.periodId?-1:a.periodId>b.periodId?1:0;});
      var openOnes=periods.filter(function(p){return p.status==='OPEN';});
      if(openOnes.length) week=openOnes[0]; else if(periods.length) week=periods[0];
    }
  } catch(e3) {}
  var pendingActivities = [];
  var pendingActsCount = 0;
  try {
    var rqp = ss.getSheetByName('REQUESTS');
    if (rqp && rqp.getLastRow() > 3) {
      rqp.getRange(4, 1, rqp.getLastRow()-3, 6).getValues().forEach(function(rr) {
        var type = safeStr(rr[3]);
        var status = safeStr(rr[5]).toUpperCase();
        if (status !== 'PENDING') return;
        if (type === 'ACTIVITY_SETUP' || type === 'RATE_EDIT') pendingActsCount++;
        if (type === 'ACTIVITY_SETUP') {
          try {
            var pl = JSON.parse(safeStr(rr[4]));
            if (pl && pl.activityName) {
              pendingActivities.push({ activityName:safeStr(pl.activityName), dept:safeStr(pl.dept), rate:safeNum(pl.rate), comm:safeNum(pl.comm) });
            }
          } catch(pe) {}
        }
      });
    }
  } catch(e5) {}
  if (!week) week = getCurrentWeek();
  return { articles:articles, contractors:contractors, masterActivities:masterActivities, pendingActivities:pendingActivities, pendingActsCount:pendingActsCount, week:week, periods:periods };
}

function getContractorsData() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    var contractors = [];
    if (mc && mc.getLastRow() > 3) {
      mc.getRange(4, 2, mc.getLastRow()-3, 6).getValues().forEach(function(r) {
        if (!r[0]) return;
        contractors.push({
          name: safeStr(r[0]),
          paymentMethod: safeStr(r[1]) || 'Cash',
          status: safeStr(r[2]),
          dept: safeStr(r[3]),
          phone: safeStr(r[4])
        });
      });
    }
    return { success: true, contractors: contractors };
  } catch(e) { return { success: false, error: e.message, contractors: [] }; }
}

function saveContractor(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc) return { success: false, error: 'MASTER_CONTRACTORS sheet not found' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    mc.getRange(mc.getLastRow() + 1, 2, 1, 6).setValues([[
      safeStr(payload.name), safeStr(payload.paymentMethod) || 'Cash',
      'ACTIVE', safeStr(payload.dept), safeStr(payload.phone), now
    ]]);
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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

function deleteOrder(sheetName) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can delete orders' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found: '+sheetName };
    var hasPaid = false;
    ws.getRange(5, 4, 45, 1).getValues().forEach(function(r){ if(safeNum(r[0])>0) hasPaid=true; });
    if (hasPaid) return { success:false, error:'Cannot delete — order has paid entries' };
    ss.deleteSheet(ws);
    var failures = [];
    try {
      var oi = ss.getSheetByName('ORDER_INDEX');
      if (oi && oi.getLastRow()>3) {
        var oiD=oi.getRange(4,2,oi.getLastRow()-3,1).getValues();
        for(var i=0;i<oiD.length;i++){if(safeStr(oiD[i][0])===sheetName){oi.deleteRow(i+4);break;}}
      }
    } catch(e1) { failures.push('ORDER_INDEX: '+e1.message); }
    try {
      var ot = ss.getSheetByName('ORDER_TRACKER');
      if (ot && ot.getLastRow()>3) {
        var otD=ot.getRange(4,1,ot.getLastRow()-3,1).getValues();
        for(var j=0;j<otD.length;j++){if(safeStr(otD[j][0])===sheetName){ot.deleteRow(j+4);break;}}
      }
    } catch(e2) { failures.push('ORDER_TRACKER: '+e2.message); }
    try {
      var wr = ss.getSheetByName('WIP_RECONCILIATION');
      if (wr && wr.getLastRow()>4) {
        var wrD=wr.getRange(5,1,wr.getLastRow()-4,1).getValues();
        for(var k=0;k<wrD.length;k++){if(safeStr(wrD[k][0])===sheetName){wr.deleteRow(k+5);break;}}
      }
    } catch(e3) { failures.push('WIP_RECONCILIATION: '+e3.message); }
    SpreadsheetApp.flush();
    if (failures.length) {
      Logger.log('deleteOrder('+sheetName+'): sheet deleted but row cleanup failed — '+failures.join('; '));
      return { success:true, warning:'Sheet deleted, but stale rows remain in: '+failures.join('; ') };
    }
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// getWIPData removed — replaced by getWipEntries / getWipGrid in Phase 5

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

// getPaymentList removed — replaced by getPaymentBatches in Phase 5

function createNewArtSheet(detailsStr) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var parse = function(key) {
    var m = detailsStr.match(new RegExp(key + ':\\s*([^|]+)'));
    return m ? m[1].trim() : '';
  };
  var orderId  = parse('OrderID') || ('ADE-' + Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyMM') + '-NEW');
  var article  = parse('Article');
  var color    = parse('Color');
  var customer = parse('Customer');
  var brand    = parse('Brand') || customer;
  var lot      = parseInt(parse('Lot')) || 0;
  var artType  = parse('Type');
  var material = parse('Material');
  var season   = parse('Season');
  var month    = parse('Month');

  var artSheets = ss.getSheets().filter(isArtSheet);
  var nums = artSheets.map(function(s){ return parseInt(s.getName().replace('ART-',''))||0; });
  var nextNum = String(Math.max.apply(null,[0].concat(nums))+1);
  while(nextNum.length < 3) nextNum = '0' + nextNum;
  var newName = 'ART-' + nextNum;

  var template = ss.getSheetByName('ART-TEMPLATE') || ss.getSheetByName('ART-001');
  var newSheet = template.copyTo(ss);
  newSheet.setName(newName);

  newSheet.getRange('C5:D49').clearContent();
  newSheet.getRange('H5:H49').clearContent();
  newSheet.getRange('J5:J49').clearContent();
  newSheet.getRange('B5:B49').clearContent();
  newSheet.getRange('E5:E49').clearContent();
  newSheet.getRange('F5:F49').clearContent();

  newSheet.getRange('B2').setValue(article + (color?' - '+color:''));
  newSheet.getRange('E2').setValue(customer);
  newSheet.getRange('H2').setValue(lot);
  newSheet.getRange('J2').setValue(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd-MMM-yyyy'));
  newSheet.getRange('M14').setValue(orderId);
  newSheet.getRange('M15').setValue(color);
  newSheet.getRange('M16').setValue(season);
  newSheet.getRange('M17').setValue(month);
  newSheet.getRange('M18').setValue(brand);
  if(artType)  newSheet.getRange('M11').setValue(artType);
  if(material) newSheet.getRange('M12').setValue(material);

  // Add formulas for totals
  for (var r = 5; r <= 49; r++) {
    newSheet.getRange('G'+r).setFormula('=IF(D'+r+'="",0,D'+r+'*F'+r+')');
    newSheet.getRange('I'+r).setFormula('=IF(D'+r+'="",0,(D'+r+'*E'+r+')+G'+r+'+IF(H'+r+'="",0,H'+r+'))');
  }

  updateTrackers(ss, newName, orderId, article, color, customer, brand, season, month, lot);
  SpreadsheetApp.flush();
  return newName;
}

function updateTrackers(ss, newName, orderId, article, color, customer, brand, season, month, lot) {
  var displayName = article + (color?' - '+color:'');
  try {
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi) {
      var oiRow = Math.max(oi.getLastRow(), 3) + 1;
      oi.getRange(oiRow, 1, 1, 11).setValues([[
        orderId, newName, article, color, customer, brand, season, month, lot,
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy'), 'Active'
      ]]);
    }
  } catch(e) { Logger.log('OI error: ' + e.message); }

  try {
    var ot = ss.getSheetByName('ORDER_TRACKER');
    if (ot) {
      var otRow = Math.max(ot.getLastRow(), 3) + 1;
      ot.getRange(otRow, 1, 1, 3).setValues([[newName, displayName, customer]]);
      ot.getRange(otRow, 4).setFormula("='"+newName+"'!H2");
      ot.getRange(otRow, 5).setValue(0);
      ot.getRange(otRow, 6).setFormula("=IFERROR('"+newName+"'!Q2,0)");
      ot.getRange(otRow, 7).setFormula("=E"+otRow+"+F"+otRow);
      ot.getRange(otRow, 8).setFormula("=IF(D"+otRow+'="","--",D'+otRow+"-G"+otRow+")");
      ot.getRange(otRow, 9).setFormula(
        '=IF(D'+otRow+'="","NO LOT SET",IF(G'+otRow+'>D'+otRow+',"OVER BY "&(G'+otRow+'-D'+otRow+')&" PAIRS",IF(G'+otRow+'=D'+otRow+',"LOT COMPLETE","OK - "&(D'+otRow+'-G'+otRow+')&" LEFT")))');
    }
  } catch(e) { Logger.log('OT error: ' + e.message); }

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
}

// archiveWeek removed — replaced by job card payment batch system in Phase 5

// clearWeekForNext removed — week lifecycle handled by payment period system in Phase 5

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

function getPrintSummary() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Build ORDER_INDEX lookup: artName → {sheetName, customer}
  var articleToSheet = {};
  try {
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() > 3) {
      oi.getRange(4, 1, oi.getLastRow()-3, 5).getValues().forEach(function(r) {
        var sheetName = safeStr(r[1]);
        var article   = safeStr(r[2]);
        var customer  = safeStr(r[4]);
        if (article) articleToSheet[article] = { sheetName:sheetName, customer:customer };
      });
    }
  } catch(e) { Logger.log('PS OI: ' + e.message); }

  var records = [];
  try {
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (ph && ph.getLastRow() > 3) {
      ph.getRange(4, 1, ph.getLastRow()-3, 8).getValues().forEach(function(r) {
        var weekRaw    = r[0];
        var weekEnding = weekRaw instanceof Date
          ? Utilities.formatDate(weekRaw, Session.getScriptTimeZone(), 'dd-MMM-yyyy')
          : safeStr(weekRaw);
        var artName     = safeStr(r[1]);
        var customer    = safeStr(r[2]);
        var contractor  = safeStr(r[3]);
        var qty         = safeNum(r[4]);
        var amount      = safeNum(r[5]);
        var approval    = safeStr(r[6]);
        var archiveDate = safeStr(r[7]);
        if (!approval || !contractor || !amount) return;
        var oiEntry = articleToSheet[artName] || {};
        records.push({
          periodId:    'P-' + weekEnding.replace(/[^a-zA-Z0-9]/g, ''),
          periodLabel: weekEnding,
          sheetName:   oiEntry.sheetName || '',
          article:     artName,
          customer:    customer || oiEntry.customer || '',
          activityName:'',
          contractor:  contractor,
          qty:         qty,
          rate:        0,
          comm:        0,
          conv:        0,
          amount:      amount,
          approvedDate:archiveDate
        });
      });
    }
  } catch(e) { Logger.log('PS PH: ' + e.message); }

  return records;
}

// approvePeriodPayment removed — replaced by approvePaymentBatch in Phase 5

function getActivitiesFromTS(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tsNumber = '';
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() > 3) {
      var oiData = oi.getRange(4, 1, oi.getLastRow()-3, 12).getValues();
      for (var i = 0; i < oiData.length; i++) {
        if (safeStr(oiData[i][1]) === sheetName) {
          // col L (index 11) = tsNumber for new rows; col F (index 5) fallback for old buggy rows
          tsNumber = safeStr(oiData[i][11]) || safeStr(oiData[i][5]);
          break;
        }
      }
    }
    if (!tsNumber) return { success:false, error:'No BOM linked to this article' };
    var tm = ss.getSheetByName('TS_MASTER');
    if (!tm || tm.getLastRow() < 2) return { success:false, error:'TS_MASTER not found' };
    var tmData = tm.getRange(2, 1, tm.getLastRow()-1, 7).getValues();
    for (var j = 0; j < tmData.length; j++) {
      if (safeStr(tmData[j][0]) === tsNumber) {
        var raw = safeStr(tmData[j][6]);
        if (!raw) return { success:true, activities:[] };
        try { return { success:true, activities:JSON.parse(raw) }; }
        catch(pe) { return { success:false, error:'Invalid activities JSON in BOM' }; }
      }
    }
    return { success:false, error:'BOM not found: ' + tsNumber };
  } catch(e) { return { success:false, error:e.message }; }
}

function getApprovedActivitiesForArticle(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:true, activities:[] };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 6).getValues();
    var activities = [];
    data.forEach(function(r) {
      if (safeStr(r[3]) !== 'ACTIVITY_SETUP' || safeStr(r[5]).toUpperCase() !== 'APPROVED') return;
      try {
        var pl = JSON.parse(safeStr(r[4]));
        if (!pl || safeStr(pl.sheet) !== sheetName) return;
        var dept = safeStr(pl.dept || '');
        var acts = pl.activities || (pl.activityName ? [{activityName:pl.activityName,rate:safeNum(pl.rate),comm:safeNum(pl.comm)}] : []);
        acts.forEach(function(a) {
          activities.push({ activityName:safeStr(a.activityName), dept:dept||safeStr(a.department||''), rate:safeNum(a.rate), comm:safeNum(a.comm) });
        });
      } catch(pe) {}
    });
    return { success:true, activities:activities };
  } catch(e) { return { success:false, error:e.message }; }
}

function getDeptStatus(sheetName) {
  var DEPT_KEYS = ['Cutting','Preparation','Fitter','Lasting','Finishing','Dispatch'];
  var status = {};
  DEPT_KEYS.forEach(function(d) { status[d] = 'NOT_SET'; });
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var DMAP = {'cutting':'Cutting','prep':'Preparation','fitter':'Fitter','lasting':'Lasting','finish':'Finishing','dispatch':'Dispatch'};
  try {
    var ds = ss.getSheetByName('DEPT_STATUS');
    if (ds && ds.getLastRow() > 1) {
      ds.getRange(2, 1, ds.getLastRow()-1, 3).getValues().forEach(function(r) {
        var sn = safeStr(r[0]), dn = safeStr(r[1]), st = safeStr(r[2]).toUpperCase();
        if (sn === sheetName && st === 'SKIPPED' && DEPT_KEYS.indexOf(dn) >= 0) status[dn] = 'SKIPPED';
      });
    }
  } catch(e) {}
  try {
    var rq = ss.getSheetByName('REQUESTS');
    if (rq && rq.getLastRow() > 3) {
      rq.getRange(4, 1, rq.getLastRow()-3, 6).getValues().forEach(function(r) {
        var rType = safeStr(r[3]);
        if (rType !== 'ACTIVITY_SETUP' && rType !== 'ACTIVITY_SETUP_RESET' && rType !== 'ACTIVITY_SETUP_UNLOCKED') return;
        var reqSt = safeStr(r[5]).toUpperCase();
        try {
          var pl = JSON.parse(safeStr(r[4]));
          if (!pl || safeStr(pl.sheet) !== sheetName) return;
          if (rType === 'ACTIVITY_SETUP_RESET') {
            var resetDn = DMAP[safeStr(pl.dept || '').toLowerCase()] || safeStr(pl.dept || '');
            if (DEPT_KEYS.indexOf(resetDn) >= 0 && status[resetDn] !== 'SKIPPED') status[resetDn] = 'NOT_SET';
            return;
          }
          if (rType === 'ACTIVITY_SETUP_UNLOCKED') {
            var unlockDn = DMAP[safeStr(pl.dept || '').toLowerCase()] || safeStr(pl.dept || '');
            if (DEPT_KEYS.indexOf(unlockDn) >= 0 && status[unlockDn] !== 'SKIPPED') status[unlockDn] = 'EDIT_UNLOCKED';
            return;
          }
          if (reqSt !== 'APPROVED' && reqSt !== 'PENDING') return;
          var depts = [];
          if (pl.dept) {
            var dn = DMAP[safeStr(pl.dept).toLowerCase()] || safeStr(pl.dept);
            if (DEPT_KEYS.indexOf(dn) >= 0) depts = [dn];
          }
          if (!depts.length && pl.activities && pl.activities.length) {
            pl.activities.forEach(function(a) {
              var dv = safeStr(a.department || '').toLowerCase();
              var dn2 = DMAP[dv] || '';
              if (dn2 && DEPT_KEYS.indexOf(dn2) >= 0 && depts.indexOf(dn2) < 0) depts.push(dn2);
            });
          }
          depts.forEach(function(dept) {
            if (status[dept] === 'SKIPPED') return;
            if (reqSt === 'APPROVED') status[dept] = 'APPROVED';
            else if (reqSt === 'PENDING') {
              if (status[dept] === 'NOT_SET') status[dept] = 'PENDING';
              else if (status[dept] === 'EDIT_UNLOCKED') status[dept] = 'EDIT_PENDING';
            } else if (reqSt === 'REJECTED' && status[dept] === 'EDIT_PENDING') {
              status[dept] = 'EDIT_UNLOCKED';
            }
          });
        } catch(pe) {}
      });
    }
  } catch(e) {}
  return status;
}

function getDeptStatusBatch(allSheetNames, requestsData, deptStatusData) {
  var DEPT_KEYS = ['Cutting','Preparation','Fitter','Lasting','Finishing','Dispatch'];
  var DMAP = {'cutting':'Cutting','prep':'Preparation','fitter':'Fitter','lasting':'Lasting','finish':'Finishing','dispatch':'Dispatch'};
  var result = {};
  allSheetNames.forEach(function(sn) {
    var status = {};
    DEPT_KEYS.forEach(function(d) { status[d] = 'NOT_SET'; });
    result[sn] = status;
  });
  deptStatusData.forEach(function(r) {
    var sn = safeStr(r[0]), dn = safeStr(r[1]), st = safeStr(r[2]).toUpperCase();
    if (result[sn] && st === 'SKIPPED' && DEPT_KEYS.indexOf(dn) >= 0) result[sn][dn] = 'SKIPPED';
  });
  requestsData.forEach(function(r) {
    var rType = safeStr(r[3]);
    if (rType !== 'ACTIVITY_SETUP' && rType !== 'ACTIVITY_SETUP_RESET' && rType !== 'ACTIVITY_SETUP_UNLOCKED') return;
    var reqSt = safeStr(r[5]).toUpperCase();
    try {
      var pl = JSON.parse(safeStr(r[4]));
      if (!pl) return;
      var sn = safeStr(pl.sheet);
      if (!result[sn]) return;
      var status = result[sn];
      if (rType === 'ACTIVITY_SETUP_RESET') {
        var resetDn = DMAP[safeStr(pl.dept || '').toLowerCase()] || safeStr(pl.dept || '');
        if (DEPT_KEYS.indexOf(resetDn) >= 0 && status[resetDn] !== 'SKIPPED') status[resetDn] = 'NOT_SET';
        return;
      }
      if (rType === 'ACTIVITY_SETUP_UNLOCKED') {
        var unlockDn = DMAP[safeStr(pl.dept || '').toLowerCase()] || safeStr(pl.dept || '');
        if (DEPT_KEYS.indexOf(unlockDn) >= 0 && status[unlockDn] !== 'SKIPPED') status[unlockDn] = 'EDIT_UNLOCKED';
        return;
      }
      if (reqSt !== 'APPROVED' && reqSt !== 'PENDING') return;
      var depts = [];
      if (pl.dept) {
        var dn = DMAP[safeStr(pl.dept).toLowerCase()] || safeStr(pl.dept);
        if (DEPT_KEYS.indexOf(dn) >= 0) depts = [dn];
      }
      if (!depts.length && pl.activities && pl.activities.length) {
        pl.activities.forEach(function(a) {
          var dv = safeStr(a.department || '').toLowerCase();
          var dn2 = DMAP[dv] || '';
          if (dn2 && DEPT_KEYS.indexOf(dn2) >= 0 && depts.indexOf(dn2) < 0) depts.push(dn2);
        });
      }
      depts.forEach(function(dept) {
        if (status[dept] === 'SKIPPED') return;
        if (reqSt === 'APPROVED') status[dept] = 'APPROVED';
        else if (reqSt === 'PENDING') {
          if (status[dept] === 'NOT_SET') status[dept] = 'PENDING';
          else if (status[dept] === 'EDIT_UNLOCKED') status[dept] = 'EDIT_PENDING';
        } else if (reqSt === 'REJECTED' && status[dept] === 'EDIT_PENDING') {
          status[dept] = 'EDIT_UNLOCKED';
        }
      });
    } catch(pe) {}
  });
  return result;
}

function markDeptSkipped(sheetName, dept) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin') return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ds = ss.getSheetByName('DEPT_STATUS');
    if (!ds) {
      ds = ss.insertSheet('DEPT_STATUS');
      ds.getRange(1, 1, 1, 4).setValues([['SheetName','Dept','Status','MarkedBy']]);
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    if (ds.getLastRow() > 1) {
      var existing = ds.getRange(2, 1, ds.getLastRow()-1, 3).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (safeStr(existing[i][0]) === sheetName && safeStr(existing[i][1]) === dept) {
          ds.getRange(i+2, 3).setValue('SKIPPED');
          ds.getRange(i+2, 4).setValue(user.name + ' — ' + now);
          SpreadsheetApp.flush();
          return { success:true };
        }
      }
    }
    ds.appendRow([sheetName, dept, 'SKIPPED', user.name + ' — ' + now]);
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function submitDeptActivities(sheetName, depts) {
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq) return { success:false, error:'REQUESTS sheet not found' };
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var lastReqId = '';
    depts.forEach(function(d) {
      var lastRow = Math.max(rq.getLastRow(), 3) + 1;
      var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
      rq.getRange(lastRow, 1, 1, 10).setValues([[
        reqId, now, user.name, 'ACTIVITY_SETUP',
        JSON.stringify({sheet:sheetName, dept:d.dept, activities:d.activities}),
        'PENDING', '', '', 'No', ''
      ]]);
      lastReqId = reqId;
    });
    SpreadsheetApp.flush();
    return { success:true, count:depts.length, reqId:lastReqId };
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

function saveActivitySetup(sheet, newActivities, dept) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheet);
    if (!ws) return { success:false, error:'Sheet not found: ' + sheet };
    var DEPT_ORDER = ['Cutting','Preparation','Fitter','Lasting','Finishing','Dispatch'];
    var NCOLS = 12;

    // Rebuild full activity list from all approved ACTIVITY_SETUP requests for this sheet
    var deptActs = {};
    var rq = ss.getSheetByName('REQUESTS');
    if (rq && rq.getLastRow() > 3) {
      rq.getRange(4, 1, rq.getLastRow()-3, 6).getValues().forEach(function(r) {
        if (safeStr(r[3]) !== 'ACTIVITY_SETUP' || safeStr(r[5]).toUpperCase() !== 'APPROVED') return;
        try {
          var pl = JSON.parse(safeStr(r[4]));
          if (!pl || safeStr(pl.sheet) !== sheet || !pl.dept || !pl.activities) return;
          if (DEPT_ORDER.indexOf(pl.dept) < 0) return;
          deptActs[pl.dept] = pl.activities.map(function(a) {
            return { activityName:safeStr(a.activityName), rate:safeNum(a.rate), comm:safeNum(a.comm) };
          });
        } catch(pe) {}
      });
    }
    // Override with the dept currently being approved
    deptActs[dept] = newActivities.map(function(a) {
      return { activityName:safeStr(a.activityName), rate:safeNum(a.rate), comm:safeNum(a.comm) };
    });

    // Preserve existing entry data (C,D,H,J,K,L) keyed by activity name
    var entryByAct = {};
    ws.getRange(5, 1, 45, NCOLS).getValues().forEach(function(r) {
      var nm = safeStr(r[1]).trim().toLowerCase();
      if (nm) entryByAct[nm] = { c:r[2], d:r[3], h:r[7], j:r[9], k:r[10], l:r[11] };
    });

    // Build rows in standard dept order
    var allRows = [];
    DEPT_ORDER.forEach(function(d) {
      if (!deptActs[d]) return;
      deptActs[d].forEach(function(a) {
        var nm = safeStr(a.activityName);
        var ent = entryByAct[nm.trim().toLowerCase()] || {};
        var row = new Array(NCOLS).fill('');
        row[0]  = allRows.length + 1;
        row[1]  = nm;
        row[2]  = ent.c || '';  // Contractor
        row[3]  = ent.d || '';  // Qty
        row[4]  = a.rate;       // Rate
        row[5]  = a.comm;       // Comm
        row[7]  = ent.h || '';  // Conveyance
        row[9]  = ent.j || '';  // Remarks
        row[10] = ent.k || '';  // PeriodId
        row[11] = ent.l || '';  // Status
        allRows.push(row);
      });
    });

    ws.getRange(5, 1, 45, NCOLS).clearContent();
    if (allRows.length > 0) ws.getRange(5, 1, allRows.length, NCOLS).setValues(allRows);

    // Restore G (comm total) and I (total) formulas
    var gFormulas = [], iFormulas = [];
    for (var r = 5; r <= 49; r++) {
      gFormulas.push(['=IF(D'+r+'="",0,D'+r+'*F'+r+')']);
      iFormulas.push(['=IF(D'+r+'="",0,(D'+r+'*E'+r+')+G'+r+'+IF(H'+r+'="",0,H'+r+'))']);
    }
    ws.getRange(5, 7, 45, 1).setFormulas(gFormulas);
    ws.getRange(5, 9, 45, 1).setFormulas(iFormulas);
    SpreadsheetApp.flush();
    return { success:true };
    } catch(e) { return { success:false, error:e.message }; }
  } finally {
    lock.releaseLock();
  }
}

function getActivitySetup(sheet) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(sheet);
    if (!ws) return [];
    var data = ws.getRange(5, 1, 45, 4).getValues();
    var result = [];
    data.forEach(function(r, i) {
      var actName = safeStr(r[1]);
      if (!actName || actName.match(/^[-=]/)) return;
      result.push({ rowIndex:i+5, activityName:actName,
        rate:safeNum(r[2]), comm:safeNum(r[3]), department:'' });
    });
    return result;
  } catch(e) { return []; }
}

function checkDuplicateRequest(type, subjectKey) {
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { isDuplicate: false };
    var rows = rq.getRange(4, 1, rq.getLastRow() - 3, 6).getValues();
    var now = new Date();
    var ms48 = 48 * 60 * 60 * 1000;
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var rowName = safeStr(r[2]);
      var rowType = safeStr(r[3]);
      var rowStatus = safeStr(r[5]).toUpperCase();
      if (rowName !== user.name) continue;
      if (rowType !== type) continue;
      if (rowStatus !== 'PENDING' && rowStatus !== 'REJECTED') continue;
      if (rowStatus === 'REJECTED') {
        var dateStr = safeStr(r[1]);
        var d = new Date(dateStr.replace(/(\d+)-(\w+)-(\d+)/, '$2 $1, $3'));
        if (isNaN(d.getTime()) || (now - d) > ms48) continue;
      }
      try {
        var pl = JSON.parse(safeStr(r[4]));
        var match = false;
        if (type === 'RATE_EDIT') {
          match = String(pl.rowIndex) === String(subjectKey);
        } else if (type === 'ACTIVITY_SETUP') {
          match = pl.sheet === subjectKey.sheet && pl.dept === subjectKey.dept;
        }
        if (match) return { isDuplicate: true, status: rowStatus };
      } catch(pe) {}
    }
    return { isDuplicate: false };
  } catch(e) { return { isDuplicate: false }; }
}

function requestActivityRateEdit(rowIndex, newRate, newComm, revisionRemark) {
  var user = getUserInfo();
  try {
    var dup = checkDuplicateRequest('RATE_EDIT', rowIndex);
    if (dup.isDuplicate) {
      if (dup.status === 'PENDING') return { success:false, error:'A request for this activity is already pending approval. Wait for Ayush to approve or reject it first.' };
      if (dup.status === 'REJECTED') return { success:false, error:'A rejected request for this activity is still open. Use Revise & Resubmit from your Alerts tab, or wait 48 hours for it to expire.' };
    }
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var payload = JSON.stringify({ rowIndex:rowIndex, newRate:newRate, newComm:newComm });
    var revHistory = revisionRemark ? 'REVISION_HISTORY: ' + revisionRemark : '';
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, 'RATE_EDIT', payload, 'PENDING', '', '', 'No', revHistory
    ]]);
    SpreadsheetApp.flush();
    notifyNewRequest_(reqId, 'RATE_EDIT', payload, user.name, now);
    return { success:true, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function requestActivitySetup(payload, revisionRemark) {
  var user = getUserInfo();
  try {
    var items = Array.isArray(payload) ? payload : [payload];
    for (var _di = 0; _di < items.length; _di++) {
      var _dup = checkDuplicateRequest('ACTIVITY_SETUP', {sheet: items[_di].sheet, dept: items[_di].dept});
      if (_dup.isDuplicate) {
        if (_dup.status === 'PENDING') return { success:false, error:'A request for this activity is already pending approval. Wait for Ayush to approve or reject it first.' };
        if (_dup.status === 'REJECTED') return { success:false, error:'A rejected request for this activity is still open. Use Revise & Resubmit from your Alerts tab, or wait 48 hours for it to expire.' };
      }
    }
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var lastReqId = '';
    var createdReqs = [];
    items.forEach(function(item) {
      var lastRow = Math.max(rq.getLastRow(), 3) + 1;
      var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
      rq.getRange(lastRow, 1, 1, 10).setValues([[
        reqId, now, user.name, 'ACTIVITY_SETUP', JSON.stringify(item), 'PENDING', '', '', 'No', revisionRemark ? 'REVISION_HISTORY: ' + revisionRemark : ''
      ]]);
      lastReqId = reqId;
      createdReqs.push({reqId:reqId, details:JSON.stringify(item)});
    });
    SpreadsheetApp.flush();
    if (createdReqs.length === 1) {
      notifyNewRequest_(createdReqs[0].reqId, 'ACTIVITY_SETUP', createdReqs[0].details, user.name, now);
    } else if (createdReqs.length > 1) {
      try {
        var batchLines = ['Activity Setup batch: ' + createdReqs.length + ' dept(s)'];
        createdReqs.forEach(function(r) {
          try { var p=JSON.parse(r.details); batchLines.push('  • '+(p.dept||'?')+' on '+(p.sheet||'?')+' ('+(p.activities?p.activities.length:0)+' activities)'); } catch(e) {}
        });
        batchLines.push(''); batchLines.push('Submitted by: ' + user.name + ' on ' + now);
        batchLines.push('Last Request ID: ' + lastReqId);
        batchLines.push(''); batchLines.push('Open Factory OS to review.');
        MailApp.sendEmail('ayush@adeesexports.com', 'Factory OS — New ACTIVITY_SETUP request', batchLines.join('\n'));
      } catch(mailErr) { Logger.log('notifyNewRequest_ batch mail error: ' + mailErr.message); }
    }
    return { success:true, reqId:lastReqId, count:items.length };
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

function approveRateEdit(requestId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:false, error:'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues();
    var targetRow = -1, payload = null;
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === requestId && safeStr(data[i][3]) === 'RATE_EDIT') {
        targetRow = i + 4;
        try { payload = JSON.parse(safeStr(data[i][4])); } catch(pe) {}
        break;
      }
    }
    if (targetRow === -1) return { success:false, error:'Request not found: ' + requestId };
    if (!payload) return { success:false, error:'Invalid payload for: ' + requestId };
    var ma = ss.getSheetByName('MASTER_RATES');
    if (!ma) return { success:false, error:'MASTER_RATES sheet not found' };
    ma.getRange(payload.rowIndex, 3).setValue(payload.newRate);
    ma.getRange(payload.rowIndex, 4).setValue(payload.newComm);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(targetRow, 6).setValue('APPROVED');
    rq.getRange(targetRow, 8).setValue(now);
    rq.getRange(targetRow, 9).setValue('Yes');
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function dismissRateEdit(reqId) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq || rq.getLastRow() < 4) return { success:false, error:'No requests found' };
    var data = rq.getRange(4, 1, rq.getLastRow()-3, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (safeStr(data[i][0]) === reqId) {
        rq.getRange(i + 4, 6).setValue('DISMISSED');
        SpreadsheetApp.flush();
        return { success:true };
      }
    }
    return { success:false, error:'Request not found: ' + reqId };
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

function createArtTemplate() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var existing = ss.getSheetByName('ART-TEMPLATE');
  if (existing) ss.deleteSheet(existing);
  var ws = ss.insertSheet('ART-TEMPLATE');
  ws.getRange('A1').setValue('ADEES EXPORTS — POST PRODUCTION COST SUMMARY');
  ws.getRange('A2').setValue('ARTICLE:');
  ws.getRange('D2').setValue('CUSTOMER:');
  ws.getRange('G2').setValue('Order Qty');
  ws.getRange('I2').setValue('DATE:');
  ws.getRange(3,1,1,10).setValues([['S.No','ACTIVITY','CONTRACTOR','QTY (pairs)','RATE (Rs/pr)','COMM (Rs/pr)','COMM TOTAL','CONVEYANCE','TOTAL (Rs)','REMARKS']]);
  for (var r=5; r<=49; r++) {
    ws.getRange('G'+r).setFormula('=IF(D'+r+'="",0,D'+r+'*F'+r+')');
    ws.getRange('I'+r).setFormula('=IF(D'+r+'="",0,(D'+r+'*E'+r+')+G'+r+'+IF(H'+r+'="",0,H'+r+'))');
  }
  ws.getRange('I50').setFormula('=SUM(I5:I49)');
  ws.getRange('A52').setValue('PAYMENT SUMMARY');
  ws.getRange('B53').setValue('CONTRACTOR'); ws.getRange('I53').setValue('TOTAL PAYABLE');
  for (var i=0; i<30; i++) {
    ws.getRange('I'+(54+i)).setFormula('=IF(B'+(54+i)+'="",0,SUMIF($C$5:$C$49,B'+(54+i)+',$I$5:$I$49))');
  }
  ws.getRange('L1').setValue('LOT CAP MONITOR');
  ws.getRange('L2').setValue('Order Qty'); ws.getRange('M2').setFormula('=H2');
  ws.getRange('L3').setValue('Prior Weeks'); ws.getRange('M3').setValue(0);
  ws.getRange('L4').setValue('This Week');
  ws.getRange('M4').setFormula('=MAX(SUMIF($B$5:$B$49,"Upper Making",$D$5:$D$49),SUMIF($B$5:$B$49,"Lasting",$D$5:$D$49),SUMIF($B$5:$B$49,"Finish",$D$5:$D$49),IFERROR(MAX($D$5:$D$49),0))');
  ws.getRange('L5').setValue('Cumulative'); ws.getRange('M5').setFormula('=M3+M4');
  ws.getRange('L6').setValue('Remaining'); ws.getRange('M6').setFormula('=IF(H2="","--",H2-M5)');
  ws.getRange('L7').setValue('STATUS');
  ws.getRange('M7').setFormula('=IF(H2="","NO LOT SET",IF(M5>H2,"OVER BY "&(M5-H2)&" PAIRS",IF(M5=H2,"LOT COMPLETE","OK - "&(H2-M5)&" LEFT")))');
  ws.getRange('L9').setValue('APPROVED BY');
  ws.getRange('L14').setValue('ORDER ID'); ws.getRange('L15').setValue('COLOR');
  ws.getRange('L16').setValue('SEASON');   ws.getRange('L17').setValue('MONTH');
  ws.getRange('L18').setValue('BRAND');
  ws.getRange('Q2').setFormula('=M4');
  return 'ART-TEMPLATE created';
}

function searchTS(query) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tm = ss.getSheetByName('TS_MASTER');
    if (!tm || tm.getLastRow() < 2) return [];
    var data = tm.getRange(2, 1, tm.getLastRow()-1, 7).getValues();
    var q = safeStr(query).trim().toLowerCase();
    var results = [];
    if (!q) {
      var start = Math.max(0, data.length - 10);
      for (var i = start; i < data.length; i++) {
        if (data[i][0]) results.push({ tsNumber:safeStr(data[i][0]), styleName:safeStr(data[i][1]), category:safeStr(data[i][2]), season:safeStr(data[i][3]), activitiesJSON:safeStr(data[i][6]) });
      }
    } else {
      data.forEach(function(r) {
        if (r[0] && safeStr(r[1]).toLowerCase().indexOf(q) > -1)
          results.push({ tsNumber:safeStr(r[0]), styleName:safeStr(r[1]), category:safeStr(r[2]), season:safeStr(r[3]), activitiesJSON:safeStr(r[6]) });
      });
    }
    return results;
  } catch(e) { Logger.log('searchTS error: ' + e.message); return []; }
}

function createTS(styleName, category, season, activities) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var user = getUserInfo();
    var tm = ss.getSheetByName('TS_MASTER');
    if (!tm) return { success:false, error:'TS_MASTER sheet not found' };
    var lastRow = Math.max(tm.getLastRow(), 1);
    var seq = String(lastRow);
    while (seq.length < 3) seq = '0' + seq;
    var tsNumber = 'BOM-' + (season||'SS26') + '-' + seq;
    tm.getRange(lastRow + 1, 1, 1, 9).setValues([[tsNumber, styleName, category||'', season||'SS26', '', '', JSON.stringify(activities||[]), new Date(), user.name]]);
    SpreadsheetApp.flush();
    return { success:true, tsNumber:tsNumber };
  } catch(e) { return { success:false, error:e.message }; }
}

function createOrder(payload) {
  if (safeNum(payload.lotSize) <= 0) return { success:false, error:'Lot size must be greater than 0' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy');
    var oi = ss.getSheetByName('ORDER_INDEX');
    var oiCount = oi ? Math.max(oi.getLastRow() - 3, 0) : 0;
    var bomSeq = String(oiCount + 1);
    while (bomSeq.length < 3) bomSeq = '0' + bomSeq;
    var bomNumber = 'WO-' + new Date().getFullYear() + '-' + bomSeq;
    var artSheets = ss.getSheets().filter(isArtSheet);
    var nums = artSheets.map(function(s){ return parseInt(s.getName().replace('ART-',''))||0; });
    var nextNum = String(Math.max.apply(null,[0].concat(nums))+1);
    while (nextNum.length < 3) nextNum = '0' + nextNum;
    var artSheet = 'ART-' + nextNum;
    var template = ss.getSheetByName('ART-TEMPLATE') || ss.getSheetByName('ART-001');
    var ws = template.copyTo(ss);
    ws.setName(artSheet);
    ws.getRange('B5:B49').clearContent();
    ws.getRange('C5:F49').clearContent();
    ws.getRange('H5:H49').clearContent();
    ws.getRange('J5:J49').clearContent();
    var lotSize = safeNum(payload.lotSize);
    var article = safeStr(payload.styleName) + (payload.color ? ' - ' + safeStr(payload.color) : '');
    ws.getRange('B2').setValue(article);
    ws.getRange('E2').setValue(safeStr(payload.buyer));
    ws.getRange('H2').setValue(lotSize);
    ws.getRange('J2').setValue(now);
    ws.getRange('M14').setValue(bomNumber);
    ws.getRange('M15').setValue(safeStr(payload.color));
    if (payload.brand) ws.getRange('M18').setValue(safeStr(payload.brand));
    if (payload.poReceiveDate) ws.getRange('M19').setValue(safeStr(payload.poReceiveDate));
    for (var r = 5; r <= 49; r++) {
      ws.getRange('G'+r).setFormula('=IF(D'+r+'="",0,D'+r+'*F'+r+')');
      ws.getRange('I'+r).setFormula('=IF(D'+r+'="",0,(D'+r+'*E'+r+')+G'+r+'+IF(H'+r+'="",0,H'+r+'))');
    }
    if (payload.tsNumber) {
      try {
        var tm = ss.getSheetByName('TS_MASTER');
        if (tm && tm.getLastRow() > 1) {
          var tmData = tm.getRange(2, 1, tm.getLastRow()-1, 7).getValues();
          for (var ti = 0; ti < tmData.length; ti++) {
            if (safeStr(tmData[ti][0]) === safeStr(payload.tsNumber)) {
              var tsActJSON = safeStr(tmData[ti][6]);
              if (tsActJSON) {
                var tsActs = JSON.parse(tsActJSON);
                if (tsActs && tsActs.length) {
                  var tsRows = tsActs.map(function(act, idx) {
                    return [idx+1, safeStr(act.activityName), '', '', safeNum(act.rate), safeNum(act.comm)];
                  });
                  ws.getRange(5, 1, tsRows.length, 6).setValues(tsRows);
                }
              }
              break;
            }
          }
        }
      } catch(tse) { Logger.log('TS activity inheritance: ' + tse.message); }
    }
    if (oi) {
      var oiRow = Math.max(oi.getLastRow(), 3) + 1;
      var _sizeRun = {};
      try { if (payload.sizeBreakdown) _sizeRun = JSON.parse(payload.sizeBreakdown); } catch(se) {}
      var _sizeVals = [23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46].map(function(s){ return parseInt(_sizeRun[String(s)])||0; });
      var _tsParts = safeStr(payload.tsNumber||'').split('-');
      var _season = _tsParts.length >= 2 ? _tsParts[1] : '';
      // Columns A–K match updateTrackers layout; L–Q are new extended fields; R–AO are size values
      oi.getRange(oiRow, 1, 1, 41).setValues([[
        bomNumber, artSheet, safeStr(payload.styleName), safeStr(payload.color),
        safeStr(payload.buyer), safeStr(payload.brand||''), _season, safeStr(payload.deliveryDate||''), lotSize,
        now, 'Active',
        safeStr(payload.tsNumber||''), safeStr(payload.poNumber||''), safeStr(payload.poReceiveDate||''),
        safeStr(payload.grading||''), safeStr(payload.category||''), safeStr(payload.sizeBreakdown||'')
      ].concat(_sizeVals)]);
    }
    var ot = ss.getSheetByName('ORDER_TRACKER');
    if (ot) {
      var otRow = Math.max(ot.getLastRow(), 3) + 1;
      ot.getRange(otRow, 1, 1, 3).setValues([[artSheet, article, safeStr(payload.buyer)]]);
      ot.getRange(otRow, 4).setFormula("='"+artSheet+"'!H2");
      ot.getRange(otRow, 5).setValue(0);
      ot.getRange(otRow, 6).setFormula("=IFERROR('"+artSheet+"'!Q2,0)");
      ot.getRange(otRow, 7).setFormula('=E'+otRow+'+F'+otRow);
      ot.getRange(otRow, 8).setFormula('=IF(D'+otRow+'="","--",D'+otRow+'-G'+otRow+')');
      ot.getRange(otRow, 9).setFormula('=IF(D'+otRow+'="","NO LOT SET",IF(G'+otRow+'>D'+otRow+',"OVER BY "&(G'+otRow+'-D'+otRow+')&" PAIRS",IF(G'+otRow+'=D'+otRow+',"LOT COMPLETE","OK - "&(D'+otRow+'-G'+otRow+')&" LEFT")))');
    }
    var wr = ss.getSheetByName('WIP_RECONCILIATION');
    if (wr) {
      var wrRow = Math.max(wr.getLastRow(), 4) + 1;
      wr.getRange(wrRow, 1, 1, 3).setValues([[artSheet, article, 'Upper Making']]);
      wr.getRange(wrRow, 5).setFormula("=IFERROR('"+artSheet+"'!Q2,0)");
      wr.getRange(wrRow, 6).setFormula('=IF(D'+wrRow+'="","--",D'+wrRow+'-E'+wrRow+')');
      wr.getRange(wrRow, 7).setFormula('=IF(D'+wrRow+'="","AWAITING",IF(D'+wrRow+'=E'+wrRow+',"MATCH",IF(E'+wrRow+'>D'+wrRow+',"PAID > MADE","UNDER-PAID")))');
    }
    SpreadsheetApp.flush();
    return { success:true, bomNumber:bomNumber, artSheet:artSheet };
  } catch(e) { Logger.log('createOrder error: ' + e.message); return { success:false, error:e.message }; }
}

// ONE-TIME CLEANUP — run once from Apps Script editor, then delete this function.
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

function getOrderProgress(artSheet) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var actDeptMap = {};
    var maS = ss.getSheetByName('MASTER_ACTIVITIES');
    if (maS && maS.getLastRow() > 1) {
      maS.getRange(2, 1, maS.getLastRow()-1, 5).getValues().forEach(function(r) {
        var dept = safeStr(r[0]).trim();
        var actName = safeStr(r[1]).trim();
        if (actName && safeStr(r[4]).trim().toUpperCase() === 'APPROVED') {
          actDeptMap[actName.toLowerCase()] = dept;
        }
      });
    }
    var ws = ss.getSheetByName(artSheet);
    if (!ws) return { error: 'Sheet not found: ' + artSheet };
    var lotSize = safeNum(ws.getRange('H2').getValue());
    var lastRow = ws.getLastRow();
    var deptQty = {};
    if (lastRow >= 5) {
      ws.getRange(5, 1, lastRow - 4, 12).getValues().forEach(function(r) {
        var actName = safeStr(r[1]).trim();
        var qty = safeNum(r[3]);
        if (!actName || qty <= 0) return;
        var dept = actDeptMap[actName.toLowerCase()] || 'unknown';
        deptQty[dept] = (deptQty[dept] || 0) + qty;
      });
    }
    var DEPT_KEY_MAP = {
      'cutting':'cutting','preparation':'prep','fitter':'fitter',
      'lasting':'lasting','finishing':'finish','dispatch':'dispatch'
    };
    var sq = {cutting:0,prep:0,fitter:0,lasting:0,finish:0,dispatch:0};
    Object.keys(deptQty).forEach(function(dept) {
      var key = DEPT_KEY_MAP[dept.toLowerCase()];
      if (key) sq[key] += deptQty[dept];
    });
    return {
      lotSize: lotSize,
      stages: [
        {key:'cutting',  label:'Cutting',             paidQty:sq.cutting},
        {key:'prep',     label:'Preparation',         paidQty:sq.prep},
        {key:'fitter',   label:'Upper Making',        paidQty:sq.fitter},
        {key:'lasting',  label:'Lasting & Pasting',   paidQty:sq.lasting},
        {key:'finish',   label:'Finishing & Packing', paidQty:sq.finish},
        {key:'dispatch', label:'Dispatch',             paidQty:sq.dispatch}
      ]
    };
  } catch(e) {
    return { error: e.message };
  }
}

// ── WIP ENTRIES — Phase 5.1 ───────────────────────────────────────────────────

function ensureWipEntriesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('WIP_ENTRIES');
  var NEW_HEADERS = ['WIP_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT','ENTRY_TYPE','PAIRS','SUBMITTED_BY','SUBMITTED_AT','PERIOD_ID','STATUS','NOTES','CONTRACTORS','JOB_CARD_REF'];
  if (!ws) {
    ws = ss.insertSheet('WIP_ENTRIES');
    ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
    ws.setFrozenRows(1);
  } else {
    if (safeStr(ws.getRange(1, 1).getValue()) !== 'WIP_ID') {
      ws.clearContents();
      ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
      ws.setFrozenRows(1);
    }
  }
  return ws;
}

function saveWipEntry(data, status) {
  var STORE_MOVEMENT_MAP = {
    'Upper Store':              ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store':  ['Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':           ['Dispatch IN','Dispatch OUT']
  };
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var orderRef    = safeStr(data.orderRef).trim();
      var workOrder   = safeStr(data.workOrder   || '').trim();
      var store       = safeStr(data.store).trim();
      var movement    = safeStr(data.movement).trim();
      var pairs       = safeNum(data.pairs);
      var periodId    = safeStr(data.periodId).trim();
      var notes       = safeStr(data.notes       || '').trim();
      var contractors = Array.isArray(data.contractors) ? data.contractors : [];
      var jobCardRef  = safeStr(data.jobCardRef  || '').trim();

      if (!orderRef) throw new Error('orderRef is required');
      if (!STORE_MOVEMENT_MAP[store]) throw new Error('Invalid store: ' + store);
      if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0) throw new Error('Invalid movement for store: ' + movement);
      if (!pairs || pairs <= 0 || Math.floor(pairs) !== pairs) throw new Error('pairs must be a positive integer');
      if (!periodId) throw new Error('periodId is required');

      status = status || 'SUBMITTED';
      var entryType = movement.slice(-2) === 'IN' ? 'IN' : 'OUT';

      var ws = ensureWipEntriesSheet();
      var dataRows = Math.max(0, ws.getLastRow() - 1);
      var nextNum  = dataRows + 1;
      var year     = new Date().getFullYear();
      var wipId    = 'WIP-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));

      var user          = getUserInfo();
      var now           = new Date().toISOString();
      var contractorsStr = contractors.join(',');

      ws.appendRow([wipId, orderRef, workOrder, store, movement, entryType, pairs, user.email, now, periodId, status, notes, contractorsStr, jobCardRef]);
      SpreadsheetApp.flush();
      return { success: true, wipId: wipId };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function getWipEntries(filters) {
  try {
    var ws = ensureWipEntriesSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows = ws.getRange(2, 1, lastRow - 1, 14).getValues();
    var entries = rows.map(function(r) {
      return {
        wipId:       safeStr(r[0]),
        orderRef:    safeStr(r[1]),
        workOrder:   safeStr(r[2]),
        store:       safeStr(r[3]),
        movement:    safeStr(r[4]),
        entryType:   safeStr(r[5]),
        pairs:       safeNum(r[6]),
        submittedBy: safeStr(r[7]),
        submittedAt: safeStr(r[8]),
        periodId:    safeStr(r[9]),
        status:      safeStr(r[10]),
        notes:       safeStr(r[11]),
        contractors: safeStr(r[12]),
        jobCardRef:  safeStr(r[13])
      };
    });
    if (filters) {
      if (filters.orderRef)   entries = entries.filter(function(e){ return e.orderRef   === filters.orderRef; });
      if (filters.store)      entries = entries.filter(function(e){ return e.store      === filters.store; });
      if (filters.movement)   entries = entries.filter(function(e){ return e.movement   === filters.movement; });
      if (filters.entryType)  entries = entries.filter(function(e){ return e.entryType  === filters.entryType; });
      if (filters.periodId)   entries = entries.filter(function(e){ return e.periodId   === filters.periodId; });
      if (filters.status)     entries = entries.filter(function(e){ return e.status     === filters.status; });
      if (filters.jobCardRef) entries = entries.filter(function(e){ return e.jobCardRef === filters.jobCardRef; });
    }
    return entries;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function voidWipEntry(wipId) {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ws = ensureWipEntriesSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'WIP entry not found' };
      var ids = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      var targetRow = -1;
      for (var i = 0; i < ids.length; i++) {
        if (safeStr(ids[i][0]) === wipId) { targetRow = i + 2; break; }
      }
      if (targetRow < 0) return { success: false, error: 'WIP entry not found' };
      var status = safeStr(ws.getRange(targetRow, 11).getValue());
      if (status === 'LINKED') return { success: false, error: 'Cannot void a linked entry' };
      ws.getRange(targetRow, 11).setValue('VOID');
      SpreadsheetApp.flush();
      return { success: true };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

// ── CTR-ID MIGRATION & CONTRACTOR LOOKUP — Phase 5.2b-pre ────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Factory OS Admin')
    .addItem('Run CTR-ID Migration', 'assignContractorIds')
    .addItem('Backfill Order Size Columns', 'backfillOrderSizesMenu')
    .addItem('Migrate WIP_ENTRIES Schema', 'migrateWipEntriesMenu')
    .addToUi();
}

function assignContractorIds() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc) return { success: false, error: 'MASTER_CONTRACTORS sheet not found' };

    var lastRow = mc.getLastRow();
    if (lastRow < 4) return { success: true, assigned: 0, duplicate: 'none' };

    // Write CTR-ID header into A3
    mc.getRange(3, 1).setValue('CTR-ID');

    // Read all data rows: cols A-G (1-7), starting row 4
    var numRows = lastRow - 3;
    var data = mc.getRange(4, 1, numRows, 7).getValues();

    // Detect duplicate "Jai Prakash Press" — keep first occurrence (lower row index)
    var jpSeen = -1;
    var duplicateRowNum = -1;
    for (var i = 0; i < data.length; i++) {
      var name = safeStr(data[i][1]).trim();
      if (name.toLowerCase() === 'jai prakash press') {
        if (jpSeen < 0) {
          jpSeen = i;
        } else {
          duplicateRowNum = i + 4;
          mc.getRange(duplicateRowNum, 4).setValue('INACTIVE');
          mc.getRange(duplicateRowNum, 5).setValue('DUPLICATE - REMOVED');
          data[i][3] = 'INACTIVE';
          Logger.log('Marked row ' + duplicateRowNum + ' (Jai Prakash Press) as DUPLICATE - REMOVED');
          break;
        }
      }
    }

    // Find highest existing CTR-nnn to avoid collisions on re-run
    var maxExisting = 0;
    for (var i = 0; i < data.length; i++) {
      var existing = safeStr(data[i][0]).trim();
      if (/^CTR-\d+$/.test(existing)) {
        var n = parseInt(existing.replace('CTR-', ''), 10);
        if (n > maxExisting) maxExisting = n;
      }
    }
    var nextId = maxExisting + 1;

    // Assign CTR-IDs to active, non-blank, ID-less rows
    var counter = 0;
    for (var i = 0; i < data.length; i++) {
      var ctrId  = safeStr(data[i][0]).trim();
      var rName  = safeStr(data[i][1]).trim();
      var status = safeStr(data[i][3]).trim().toUpperCase();
      if (!rName) continue;
      if (status === 'INACTIVE') continue;
      if (ctrId) continue;
      var newId = 'CTR-' + (String(nextId).padStart ? String(nextId).padStart(3,'0') : ('00'+nextId).slice(-3));
      mc.getRange(i + 4, 1).setValue(newId);
      Logger.log('Assigned ' + newId + ' → row ' + (i+4) + ': ' + rName);
      nextId++;
      counter++;
    }

    SpreadsheetApp.flush();
    var dupMsg = duplicateRowNum > 0 ? 'row ' + duplicateRowNum + ' marked INACTIVE' : 'none found';
    Logger.log('assignContractorIds done — assigned: ' + counter + ', duplicate: ' + dupMsg);
    return { success: true, assigned: counter, duplicate: dupMsg };
  } catch(e) {
    Logger.log('assignContractorIds error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getContractors() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc || mc.getLastRow() < 4) return [];
    var rows = mc.getRange(4, 1, mc.getLastRow()-3, 4).getValues();
    var result = [];
    rows.forEach(function(r) {
      var name   = safeStr(r[1]).trim();
      var status = safeStr(r[3]).trim().toUpperCase();
      if (!name) return;
      if (status === 'INACTIVE') return;
      result.push({
        ctrId:         safeStr(r[0]).trim(),
        name:          name,
        paymentMethod: safeStr(r[2]).trim() || 'Cash',
        status:        status || 'ACTIVE'
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function ensureEnrollmentsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('CONTRACTOR_ENROLLMENTS');
  if (!ws) {
    ws = ss.insertSheet('CONTRACTOR_ENROLLMENTS');
    ws.getRange(1, 1, 1, 7).setValues([[
      'ENROLLMENT_ID', 'CONTRACTOR_ID', 'CONTRACTOR_NAME',
      'DEPARTMENT', 'ENROLLED_BY', 'ENROLLED_AT', 'STATUS'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function enrollContractor(data) {
  var VALID_DEPTS = [
    'Cutting', 'Preparation', 'Fitter',
    'Lasting/Pasting', 'Finishing/Packing', 'Dispatch'
  ];
  const lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var contractorId = safeStr(data.contractorId).trim();
      var department   = safeStr(data.department).trim();
      if (!contractorId) throw new Error('contractorId is required');
      if (VALID_DEPTS.indexOf(department) < 0) throw new Error('Invalid department: ' + department);

      var ws = ensureEnrollmentsSheet();
      var lastRow = ws.getLastRow();
      if (lastRow > 1) {
        var existing = ws.getRange(2, 1, lastRow - 1, 7).getValues();
        for (var i = 0; i < existing.length; i++) {
          if (safeStr(existing[i][1]).trim() === contractorId &&
              safeStr(existing[i][3]).trim() === department &&
              safeStr(existing[i][6]).trim().toUpperCase() === 'ACTIVE') {
            return { success: false, error: 'Already enrolled in this department' };
          }
        }
      }

      var contractorName = '';
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() >= 4) {
        var mcRows = mc.getRange(4, 1, mc.getLastRow() - 3, 2).getValues();
        for (var j = 0; j < mcRows.length; j++) {
          if (safeStr(mcRows[j][0]).trim() === contractorId) {
            contractorName = safeStr(mcRows[j][1]).trim();
            break;
          }
        }
      }

      var dataRows = Math.max(0, lastRow - 1);
      var nextNum  = dataRows + 1;
      var year     = new Date().getFullYear();
      var enrollmentId = 'ENR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3, '0') : ('00' + nextNum).slice(-3));
      var user = getUserInfo();
      var now  = new Date().toISOString();
      ws.appendRow([enrollmentId, contractorId, contractorName, department, user.email, now, 'ACTIVE']);
      SpreadsheetApp.flush();
      return { success: true, enrollmentId: enrollmentId };
    } catch(e) { return { success: false, error: e.message }; }
  } finally { lock.releaseLock(); }
}

function unenrollContractor(enrollmentId) {
  const lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ws = ensureEnrollmentsSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'Enrollment not found' };
      var colA = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < colA.length; i++) {
        if (safeStr(colA[i][0]).trim() === safeStr(enrollmentId).trim()) {
          ws.getRange(i + 2, 7).setValue('INACTIVE');
          SpreadsheetApp.flush();
          return { success: true };
        }
      }
      return { success: false, error: 'Enrollment not found' };
    } catch(e) { return { success: false, error: e.message }; }
  } finally { lock.releaseLock(); }
}

function getEnrollments(filters) {
  try {
    var ws = ensureEnrollmentsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows = ws.getRange(2, 1, lastRow - 1, 7).getValues();
    var result = [];
    rows.forEach(function(r) {
      var obj = {
        enrollmentId:   safeStr(r[0]).trim(),
        contractorId:   safeStr(r[1]).trim(),
        contractorName: safeStr(r[2]).trim(),
        department:     safeStr(r[3]).trim(),
        enrolledBy:     safeStr(r[4]).trim(),
        enrolledAt:     safeStr(r[5]).trim(),
        status:         safeStr(r[6]).trim()
      };
      if (!obj.enrollmentId) return;
      if (filters) {
        if (filters.contractorId && obj.contractorId !== filters.contractorId) return;
        if (filters.department   && obj.department   !== filters.department)   return;
        if (filters.status       && obj.status.toUpperCase() !== filters.status.toUpperCase()) return;
      }
      result.push(obj);
    });
    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

// ── WIP GRID — Phase 5.2f ────────────────────────────────────────────────────

function getWipGrid() {
  var ALL_MOVEMENTS = [
    'Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT',
    'Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT',
    'Dispatch IN','Dispatch OUT'
  ];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var oi = ss.getSheetByName('ORDER_INDEX');
    var orders = [];
    if (oi && oi.getLastRow() >= 4) {
      var oiRows = oi.getRange(4, 1, oi.getLastRow() - 3, 9).getValues();
      oiRows.forEach(function(r) {
        var orderId  = safeStr(r[0]).trim();
        var artSheet = safeStr(r[1]).trim();
        var article  = safeStr(r[2]).trim();
        var color    = safeStr(r[3]).trim();
        var customer = safeStr(r[4]).trim();
        var lotSize  = safeNum(r[8]);
        if (!orderId || lotSize <= 0) return;
        orders.push({ orderId: orderId, artSheet: artSheet, article: article, color: color, customer: customer, lotSize: lotSize });
      });
    }
    var ws = ensureWipEntriesSheet();
    var today = new Date();
    var mm = String(today.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(today.getDate());      if (dd.length < 2) dd = '0' + dd;
    var todayStr = today.getFullYear() + '-' + mm + '-' + dd;
    var map = {};
    if (ws.getLastRow() >= 2) {
      var wipRows = ws.getRange(2, 1, ws.getLastRow() - 1, 14).getValues();
      wipRows.forEach(function(r) {
        var orderRef    = safeStr(r[1]).trim();  // WIP_ENTRIES col B = ART sheet name
        var movement    = safeStr(r[4]).trim();
        var pairs       = safeNum(r[6]);
        var submittedAt = safeStr(r[8]).trim();
        var status      = safeStr(r[10]).trim();
        if (!orderRef || !movement || status === 'VOID') return;
        // Match WIP entry to order via artSheet (col B of ORDER_INDEX)
        var matched = null;
        for (var oi2 = 0; oi2 < orders.length; oi2++) {
          if (orders[oi2].artSheet === orderRef) { matched = orders[oi2]; break; }
        }
        if (!matched) return;
        var key = matched.orderId;
        if (!map[key]) map[key] = {};
        if (!map[key][movement]) map[key][movement] = { today: 0, total: 0 };
        map[key][movement].total += pairs;
        if (submittedAt && submittedAt.indexOf(todayStr) === 0) map[key][movement].today += pairs;
      });
    }
    var entries = [];
    Object.keys(map).forEach(function(orderId) {
      ALL_MOVEMENTS.forEach(function(mv) {
        var m = map[orderId][mv] || { today: 0, total: 0 };
        entries.push({ orderRef: orderId, movement: mv, pairsToday: m.today, pairsTotal: m.total });
      });
    });
    return { orders: orders, entries: entries, movements: ALL_MOVEMENTS };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function submitWipGrid(gridEntries) {
  var STORE_MOVEMENT_MAP = {
    'Upper Store':              ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store':  ['Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':           ['Dispatch IN','Dispatch OUT']
  };
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      if (!Array.isArray(gridEntries) || !gridEntries.length) throw new Error('No entries provided');
      var ws = ensureWipEntriesSheet();
      var today = new Date();
      var mm = String(today.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
      var dd = String(today.getDate());      if (dd.length < 2) dd = '0' + dd;
      var defaultPeriodId = 'GRID-' + today.getFullYear() + '-' + mm + '-' + dd;
      var user = getUserInfo();
      var now  = new Date().toISOString();
      var toInsert = [];
      gridEntries.forEach(function(entry) {
        var orderRef  = safeStr(entry.orderRef).trim();
        var workOrder = safeStr(entry.workOrder  || '').trim();
        var store     = safeStr(entry.store).trim();
        var movement  = safeStr(entry.movement).trim();
        var pairs     = safeNum(entry.pairs);
        var periodId  = safeStr(entry.periodId   || defaultPeriodId).trim();
        if (!orderRef) throw new Error('orderRef required');
        if (!STORE_MOVEMENT_MAP[store]) throw new Error('Invalid store: ' + store);
        if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0) throw new Error('Invalid movement for store: ' + movement);
        if (!pairs || pairs <= 0) throw new Error('pairs must be positive for: ' + orderRef);
        var entryType = movement.slice(-2) === 'IN' ? 'IN' : 'OUT';
        toInsert.push([orderRef, workOrder, store, movement, entryType, pairs, periodId]);
      });
      var dataRows = Math.max(0, ws.getLastRow() - 1);
      var saved = 0;
      toInsert.forEach(function(row, i) {
        var nextNum = dataRows + i + 1;
        var seq = String(nextNum); while (seq.length < 3) seq = '0' + seq;
        var wipId = 'WIP-' + today.getFullYear() + '-' + seq;
        ws.appendRow([wipId, row[0], row[1], row[2], row[3], row[4], row[5], user.email, now, row[6], 'PENDING', 'Grid entry', '', '']);
        saved++;
      });
      SpreadsheetApp.flush();
      return { success: true, saved: saved };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

// ── WIP_ENTRIES SCHEMA MIGRATION — Phase 5.3 ─────────────────────────────────

function migrateWipEntries() {
  var STAGE_MAP = {
    'Cutting':             { store: 'Upper Store',             movement: 'Cutting IN' },
    'Preparation':         { store: 'Upper Store',             movement: 'Preparation IN' },
    'Upper Making':        { store: 'Upper Store',             movement: 'Fitter IN' },
    'Lasting & Pasting':   { store: 'Lasting & Packing Store', movement: 'Lasting IN' },
    'Finishing & Packing': { store: 'Lasting & Packing Store', movement: 'Packing IN' },
    'Dispatch':            { store: 'Dispatch Store',          movement: 'Dispatch IN' }
  };
  var NEW_HEADERS = ['WIP_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT','ENTRY_TYPE','PAIRS','SUBMITTED_BY','SUBMITTED_AT','PERIOD_ID','STATUS','NOTES','CONTRACTORS','JOB_CARD_REF'];
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var ws = ss.getSheetByName('WIP_ENTRIES');
      if (!ws) {
        ensureWipEntriesSheet();
        return { success: true, message: 'Created WIP_ENTRIES with new schema. No migration needed.' };
      }
      var header3 = safeStr(ws.getRange(1, 3).getValue());
      if (header3 === 'WORK_ORDER') {
        return { success: true, message: 'Already on new schema. No migration needed.' };
      }
      var header1 = safeStr(ws.getRange(1, 1).getValue());
      if (header1 !== 'WIP_ID') {
        ws.clearContents();
        ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
        ws.setFrozenRows(1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Reinitialized blank/corrupt sheet with new schema.' };
      }
      var lastRow = ws.getLastRow();
      var migratedRows = 0;
      var skippedRows  = 0;
      var newRows      = [];
      if (lastRow >= 2) {
        var oldData = ws.getRange(2, 1, lastRow - 1, 12).getValues();
        oldData.forEach(function(r) {
          var wipId       = safeStr(r[0]).trim();
          var orderRef    = safeStr(r[1]).trim();
          var workOrder   = safeStr(r[2]).trim();
          var stage       = safeStr(r[3]).trim();
          var pairs       = safeNum(r[5]);
          var submittedBy = safeStr(r[6]).trim();
          var submittedAt = safeStr(r[7]).trim();
          var periodId    = safeStr(r[8]).trim();
          var status      = safeStr(r[9]).trim();
          var notes       = safeStr(r[10]).trim();
          var contractors = safeStr(r[11]).trim();
          if (!wipId) { skippedRows++; return; }
          var mapped    = STAGE_MAP[stage] || { store: 'Upper Store', movement: 'Cutting IN' };
          newRows.push([wipId, orderRef, workOrder, mapped.store, mapped.movement, 'IN', pairs, submittedBy, submittedAt, periodId, status, notes, contractors, '']);
          migratedRows++;
        });
      }
      ws.clearContents();
      ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
      if (newRows.length) ws.getRange(2, 1, newRows.length, 14).setValues(newRows);
      ws.setFrozenRows(1);
      SpreadsheetApp.flush();
      return { success: true, migrated: migratedRows, skipped: skippedRows };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function migrateWipEntriesMenu() {
  var result = migrateWipEntries();
  var ui = SpreadsheetApp.getUi();
  var msg = result.success
    ? (result.message || ('Migrated: ' + result.migrated + ' rows. Skipped: ' + result.skipped + ' blank rows.'))
    : ('Error: ' + result.error);
  ui.alert('Migrate WIP_ENTRIES Schema', msg, ui.ButtonSet.OK);
}

// ── ORDER SIZE LOOKUP — Phase pre-5.2g ───────────────────────────────────────

var _SIZES_RANGE = [23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46];

function getOrderSizes(orderRef) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (!oi) return { success: false, error: 'ORDER_INDEX sheet not found' };
    var lastRow = oi.getLastRow();
    if (lastRow < 4) return { success: false, error: 'Order not found' };
    var ncols = Math.max(oi.getLastColumn(), 41);
    var data = oi.getRange(4, 1, lastRow - 3, ncols).getValues();
    var ref = safeStr(orderRef).trim();
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      // col A (index 0) = workOrder/WO-...; col B (index 1) = artSheet/ART-...
      if (safeStr(r[0]).trim() === ref || safeStr(r[1]).trim() === ref) {
        var sizes = {};
        var totalQty = 0;
        _SIZES_RANGE.forEach(function(s, si) {
          var qty = safeNum(r[17 + si]);  // size cols start at col R (index 17)
          if (qty > 0) { sizes[String(s)] = qty; totalQty += qty; }
        });
        return { sizes: sizes, totalQty: totalQty };
      }
    }
    return { success: false, error: 'Order not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function backfillOrderSizes() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (!oi) return { success: false, error: 'ORDER_INDEX sheet not found' };
    var rq = ss.getSheetByName('REQUESTS');
    if (!rq) return { success: false, error: 'REQUESTS sheet not found' };

    // Build lookup: tsNumber → sizeBreakdown JSON string (from NEW_ORDER requests)
    var tsToSize = {};
    if (rq.getLastRow() >= 4) {
      var reqRows = rq.getRange(4, 1, rq.getLastRow() - 3, 6).getValues();
      reqRows.forEach(function(r) {
        if (safeStr(r[3]).trim() !== 'NEW_ORDER') return;
        try {
          var pl = JSON.parse(safeStr(r[4]));
          if (pl && pl.tsNumber && pl.sizeBreakdown) {
            tsToSize[safeStr(pl.tsNumber).trim()] = safeStr(pl.sizeBreakdown);
          }
        } catch(e) {}
      });
    }

    if (oi.getLastRow() < 4) return { success: true, updated: 0, skipped: 0 };
    var oiRows = oi.getRange(4, 1, oi.getLastRow() - 3, 12).getValues();
    var updated = 0, skipped = 0;
    oiRows.forEach(function(r, i) {
      // col L (index 11) = tsNumber for new rows; col F (index 5) fallback for old buggy rows
      var tsNumber = safeStr(r[11]).trim() || safeStr(r[5]).trim();
      if (!tsNumber || !tsToSize[tsNumber]) { skipped++; return; }
      var sizeRun = {};
      try { sizeRun = JSON.parse(tsToSize[tsNumber]); } catch(e) { skipped++; return; }
      var sizeVals = _SIZES_RANGE.map(function(s){ return parseInt(sizeRun[String(s)])||0; });
      oi.getRange(4 + i, 18, 1, 24).setValues([sizeVals]);  // col R (1-based 18) = SIZE_23
      updated++;
    });
    SpreadsheetApp.flush();
    return { success: true, updated: updated, skipped: skipped };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function backfillOrderSizesMenu() {
  var result = backfillOrderSizes();
  var msg = result.success
    ? 'Done. Updated: ' + result.updated + ' rows | Skipped: ' + result.skipped + ' rows'
    : 'Error: ' + result.error;
  SpreadsheetApp.getUi().alert('Backfill Order Sizes', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── DAILY REPORTS — Phase 5.3a ────────────────────────────────────────────────

function ensureDailyReportsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('DAILY_REPORTS');
  if (!ws) {
    ws = ss.insertSheet('DAILY_REPORTS');
    ws.getRange(1, 1, 1, 7).setValues([[
      'REPORT_ID', 'DATE', 'SUBMITTED_BY', 'SUBMITTED_AT',
      'ENTRY_COUNT', 'TOTAL_PAIRS', 'STATUS'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function submitDay() {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ss        = SpreadsheetApp.openById(SHEET_ID);
      var email     = Session.getActiveUser().getEmail();
      var todayDate = new Date().toISOString().slice(0, 10);
      var now       = new Date().toISOString();

      var ws      = ensureWipEntriesSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'No draft entries to submit' };

      var rows        = ws.getRange(2, 1, lastRow - 1, 14).getValues();
      var matchedIdxs = [];
      rows.forEach(function(r, i) {
        var status      = safeStr(r[10]).trim();
        var submittedBy = safeStr(r[7]).trim();
        var submittedAt = safeStr(r[8]).trim();
        if (status === 'DRAFT' && submittedBy === email && submittedAt.slice(0, 10) === todayDate) {
          matchedIdxs.push(i);
        }
      });

      if (!matchedIdxs.length) return { success: false, error: 'No draft entries to submit' };

      matchedIdxs.forEach(function(i) {
        ws.getRange(i + 2, 11).setValue('SUBMITTED');
      });

      var entryCount = matchedIdxs.length;
      var totalPairs = 0;
      matchedIdxs.forEach(function(i) { totalPairs += safeNum(rows[i][6]); });

      var dr       = ensureDailyReportsSheet();
      var drRows   = Math.max(0, dr.getLastRow() - 1);
      var nextNum  = drRows + 1;
      var year     = new Date().getFullYear();
      var reportId = 'DR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));

      dr.appendRow([reportId, todayDate, email, now, entryCount, totalPairs, 'SUBMITTED']);
      SpreadsheetApp.flush();
      return { success: true, reportId: reportId, entryCount: entryCount, totalPairs: totalPairs, date: todayDate };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function generateDailyReport() {
  try {
    var email     = Session.getActiveUser().getEmail();
    var todayDate = new Date().toISOString().slice(0, 10);
    var now       = new Date().toISOString();

    var wipWs   = ensureWipEntriesSheet();
    var lastRow = wipWs.getLastRow();
    var entryCount = 0, totalPairs = 0;
    if (lastRow >= 2) {
      var rows = wipWs.getRange(2, 1, lastRow - 1, 11).getValues();
      rows.forEach(function(r) {
        if (safeStr(r[10]).trim() === 'SUBMITTED' &&
            safeStr(r[7]).trim() === email &&
            safeStr(r[8]).trim().slice(0, 10) === todayDate) {
          entryCount++;
          totalPairs += safeNum(r[6]);
        }
      });
    }
    if (entryCount === 0) return;

    var dr      = ensureDailyReportsSheet();
    var drLast  = dr.getLastRow();
    var foundRow = -1;
    if (drLast >= 2) {
      var drRows = dr.getRange(2, 1, drLast - 1, 4).getValues();
      for (var i = 0; i < drRows.length; i++) {
        if (safeStr(drRows[i][1]).trim() === todayDate &&
            safeStr(drRows[i][2]).trim() === email) {
          foundRow = i + 2; break;
        }
      }
    }
    if (foundRow > 0) {
      dr.getRange(foundRow, 4).setValue(now);
      dr.getRange(foundRow, 5).setValue(entryCount);
      dr.getRange(foundRow, 6).setValue(totalPairs);
    } else {
      var drRowCount = Math.max(0, dr.getLastRow() - 1);
      var nextNum    = drRowCount + 1;
      var year       = new Date().getFullYear();
      var reportId   = 'DR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));
      dr.appendRow([reportId, todayDate, email, now, entryCount, totalPairs, 'AUTO']);
    }
    SpreadsheetApp.flush();
  } catch(e) {}
}

function getDailyReports(filters) {
  try {
    var ws      = ensureDailyReportsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows    = ws.getRange(2, 1, lastRow - 1, 7).getValues();
    var result  = [];
    rows.forEach(function(r) {
      var reportId    = safeStr(r[0]).trim();
      if (!reportId) return;
      var date        = safeStr(r[1]).trim();
      var submittedBy = safeStr(r[2]).trim();
      var submittedAt = safeStr(r[3]).trim();
      var entryCount  = safeNum(r[4]);
      var totalPairs  = safeNum(r[5]);
      var status      = safeStr(r[6]).trim();
      if (filters) {
        if (filters.date        && date        !== filters.date)        return;
        if (filters.submittedBy && submittedBy !== filters.submittedBy) return;
      }
      result.push({ reportId: reportId, date: date, submittedBy: submittedBy, submittedAt: submittedAt, entryCount: entryCount, totalPairs: totalPairs, status: status });
    });
    result.sort(function(a, b) { return b.date < a.date ? -1 : b.date > a.date ? 1 : 0; });
    if (filters && filters.limit) result = result.slice(0, filters.limit);
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getTodaysDrafts() {
  try {
    var ss        = SpreadsheetApp.openById(SHEET_ID);
    var todayDate = new Date().toISOString().slice(0, 10);
    var ws        = ensureWipEntriesSheet();
    var lastRow   = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows      = ws.getRange(2, 1, lastRow - 1, 14).getValues();
    var result    = [];
    rows.forEach(function(r) {
      var status      = safeStr(r[10]).trim();
      var submittedAt = safeStr(r[8]).trim();
      if (status !== 'DRAFT' || submittedAt.slice(0, 10) !== todayDate) return;
      result.push({
        wipId:       safeStr(r[0]),
        orderRef:    safeStr(r[1]),
        workOrder:   safeStr(r[2]),
        store:       safeStr(r[3]),
        movement:    safeStr(r[4]),
        entryType:   safeStr(r[5]),
        pairs:       safeNum(r[6]),
        submittedBy: safeStr(r[7]),
        submittedAt: submittedAt,
        periodId:    safeStr(r[9]),
        status:      status,
        notes:       safeStr(r[11]),
        contractors: safeStr(r[12]),
        jobCardRef:  safeStr(r[13])
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── JOB CARDS ─────────────────────────────────────────────────────────────────

function ensureJobCardsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('JOB_CARDS');
  if (!ws) {
    ws = ss.insertSheet('JOB_CARDS');
    // STATUS values: ISSUED | PARTIAL | COMPLETE | PAYMENT_PENDING | PAID | CANCELLED
    ws.appendRow([
      'JOB_CARD_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT',
      'CONTRACTOR_ID','PAIRS_ISSUED','PAIRS_RECEIVED','SIZE_BREAKDOWN',
      'ISSUED_BY','ISSUED_AT','EXPECTED_RETURN','RECEIVED_AT','STATUS','NOTES'
    ]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function issueJobCard(data) {
  var STORE_MOVEMENT_MAP = {
    'Upper Store':             ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store': ['Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':          ['Dispatch IN','Dispatch OUT']
  };
  var orderRef       = safeStr(data.orderRef       || '').trim();
  var workOrder      = safeStr(data.workOrder      || '').trim();
  var store          = safeStr(data.store          || '').trim();
  var movement       = safeStr(data.movement       || '').trim();
  var contractorId   = safeStr(data.contractorId   || '').trim();
  var pairsIssued    = safeNum(data.pairsIssued);
  var sizeBreakdown  = data.sizeBreakdown || {};
  var expectedReturn = safeStr(data.expectedReturn || '').trim();
  var notes          = safeStr(data.notes          || '').trim();

  if (!orderRef)                                                                    return { success: false, error: 'orderRef is required' };
  if (!STORE_MOVEMENT_MAP[store])                                                   return { success: false, error: 'Invalid store: ' + store };
  if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0)                             return { success: false, error: 'Invalid movement for store: ' + movement };
  if (!contractorId)                                                                return { success: false, error: 'contractorId is required' };
  if (!pairsIssued || pairsIssued <= 0 || Math.floor(pairsIssued) !== pairsIssued) return { success: false, error: 'pairsIssued must be a positive integer' };
  if (!expectedReturn)                                                              return { success: false, error: 'expectedReturn is required' };

  // Check approved activities exist for this order + department
  var deptKey = {
    'Cutting IN':     'cutting',
    'Preparation IN': 'prep',
    'Fitter IN':      'fitter',
    'Upper IN':       'lasting',
    'Lasting IN':     'lasting',
    'Packing IN':     'finishing',
    'Dispatch IN':    'dispatch'
  }[movement] || '';

  if (deptKey) {
    var actResult = getApprovedActivitiesForArticle(orderRef);
    if (actResult && actResult.success && Array.isArray(actResult.activities)) {
      var deptActs = actResult.activities.filter(function(a) {
        return safeStr(a.dept).toLowerCase().indexOf(deptKey) === 0;
      });
      if (deptActs.length === 0) {
        return {
          success: false,
          error: 'No approved activities for this department on order ' + orderRef +
                 '. Ask Arvind to set up and get activities approved first.'
        };
      }
    }
  }

  var jobCardId;
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var ws       = ensureJobCardsSheet();
    var dataRows = Math.max(0, ws.getLastRow() - 1);
    var nextNum  = dataRows + 1;
    var year     = new Date().getFullYear();
    var seq      = String(nextNum); while (seq.length < 3) seq = '0' + seq;
    jobCardId    = 'JC-' + year + '-' + seq;
    var issuedBy = Session.getActiveUser().getEmail();
    var now      = new Date().toISOString();
    ws.appendRow([
      jobCardId, orderRef, workOrder, store, movement, contractorId,
      pairsIssued, 0, JSON.stringify(sizeBreakdown), issuedBy,
      now, expectedReturn, '', 'ISSUED', notes
    ]);
    SpreadsheetApp.flush();
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }

  // Resolve current open periodId; fall back to synthetic JC date if none open
  var periodId = 'JC-' + new Date().toISOString().slice(0, 10);
  try {
    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var pp  = ss2.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var ppV  = pp.getRange(2, 1, pp.getLastRow() - 1, 7).getValues();
      var oids = [];
      ppV.forEach(function(r){ if (safeStr(r[6]).trim().toUpperCase() === 'OPEN') oids.push(safeStr(r[0])); });
      oids.sort();
      if (oids.length) periodId = oids[0];
    }
  } catch(pe) {}

  // Create IN-side WIP entry; saveWipEntry manages its own lock
  var wipWarning;
  try {
    var wipResult = saveWipEntry({
      orderRef:    orderRef,
      workOrder:   workOrder,
      store:       store,
      movement:    movement,
      pairs:       pairsIssued,
      periodId:    periodId,
      notes:       'Job Card ' + jobCardId,
      contractors: [contractorId],
      jobCardRef:  jobCardId
    });
    if (wipResult && wipResult.success === false) wipWarning = wipResult.error;
    else try { generateDailyReport(); } catch(e) {}
  } catch(wipErr) { wipWarning = wipErr.message; }

  var issueResult = { success: true, jobCardId: jobCardId };
  if (wipWarning) issueResult.warning = 'WIP entry not created: ' + wipWarning;
  return issueResult;
}

function receiveJobCard(data) {
  var jobCardId     = safeStr(data.jobCardId     || '').trim();
  var pairsReceived = safeNum(data.pairsReceived);
  var notes         = safeStr(data.notes         || '').trim();

  if (!jobCardId)                                                                           return { success: false, error: 'jobCardId is required' };
  if (!pairsReceived || pairsReceived <= 0 || Math.floor(pairsReceived) !== pairsReceived) return { success: false, error: 'pairsReceived must be a positive integer' };

  var orderRef, workOrder, store, inMovement, contractorId;
  var pairsIssued, effectivePairs, newReceived, newStatus;

  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    var ws      = ensureJobCardsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return { success: false, error: 'Job Card not found' };
    var rows     = ws.getRange(2, 1, lastRow - 1, 15).getValues();
    var rowIndex = -1;
    for (var i = 0; i < rows.length; i++) {
      if (safeStr(rows[i][0]).trim() === jobCardId) { rowIndex = i; break; }
    }
    if (rowIndex < 0) return { success: false, error: 'Job Card not found' };

    var row           = rows[rowIndex];
    var sheetRow      = rowIndex + 2;
    var currentStatus = safeStr(row[13]).trim();
    if (currentStatus === 'COMPLETE')  return { success: false, error: 'Job Card already complete' };
    if (currentStatus === 'CANCELLED') return { success: false, error: 'Job Card is cancelled' };

    pairsIssued      = safeNum(row[6]);
    var currentRecvd = safeNum(row[7]);
    orderRef         = safeStr(row[1]).trim();
    workOrder        = safeStr(row[2]).trim();
    store            = safeStr(row[3]).trim();
    inMovement       = safeStr(row[4]).trim();
    contractorId     = safeStr(row[5]).trim();
    var existingNotes = safeStr(row[14]).trim();

    if (inMovement.slice(-2) !== 'IN') return { success: false, error: 'Job card movement must be an IN movement' };

    effectivePairs = pairsReceived;
    var finalNotes = notes;
    if (currentRecvd + pairsReceived > pairsIssued) {
      var excess   = (currentRecvd + pairsReceived) - pairsIssued;
      effectivePairs = pairsIssued - currentRecvd;
      finalNotes   = (notes ? notes + '; ' : '') + 'Capped: ' + excess + ' excess pairs ignored';
    }
    if (effectivePairs <= 0) return { success: false, error: 'No remaining capacity on this job card' };

    newReceived = currentRecvd + effectivePairs;
    newStatus   = newReceived >= pairsIssued ? 'COMPLETE' : 'PARTIAL';
    var now     = new Date().toISOString();

    ws.getRange(sheetRow, 8).setValue(newReceived);
    if (!safeStr(row[12]).trim()) ws.getRange(sheetRow, 13).setValue(now);
    ws.getRange(sheetRow, 14).setValue(newStatus);
    if (finalNotes) {
      ws.getRange(sheetRow, 15).setValue(existingNotes ? existingNotes + '; ' + finalNotes : finalNotes);
    }
    SpreadsheetApp.flush();
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }

  // Resolve current open periodId; fall back to synthetic JC date if none open
  var periodId = 'JC-' + new Date().toISOString().slice(0, 10);
  try {
    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var pp  = ss2.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var ppV  = pp.getRange(2, 1, pp.getLastRow() - 1, 7).getValues();
      var oids = [];
      ppV.forEach(function(r){ if (safeStr(r[6]).trim().toUpperCase() === 'OPEN') oids.push(safeStr(r[0])); });
      oids.sort();
      if (oids.length) periodId = oids[0];
    }
  } catch(pe) {}

  var rcvWarning;
  // 'Upper IN' is a one-way transfer into Lasting & Packing Store — no OUT counterpart exists.
  // All other IN movements pair with an OUT movement of the same prefix.
  if (inMovement === 'Upper IN') {
    // No OUT-side WIP entry for Upper IN receives; job card is simply marked COMPLETE.
  } else {
    var outMovement = inMovement.slice(0, -2) + 'OUT';
    try {
      var wipResult = saveWipEntry({
        orderRef:    orderRef,
        workOrder:   workOrder,
        store:       store,
        movement:    outMovement,
        pairs:       effectivePairs,
        periodId:    periodId,
        notes:       'Job Card ' + jobCardId + ' receive',
        contractors: [contractorId],
        jobCardRef:  jobCardId
      });
      if (wipResult && wipResult.success === false) rcvWarning = wipResult.error;
      else try { generateDailyReport(); } catch(e) {}
    } catch(wipErr) { rcvWarning = wipErr.message; }
  }

  var rcvResult = { success: true, jobCardId: jobCardId, totalReceived: newReceived, pairsIssued: pairsIssued, status: newStatus };
  if (rcvWarning) rcvResult.warning = 'WIP entry not created: ' + rcvWarning;
  return rcvResult;
}

function getJobCards(filters) {
  try {
    var ws      = ensureJobCardsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows   = ws.getRange(2, 1, lastRow - 1, 15).getValues();
    var result = [];
    rows.forEach(function(r) {
      if (!safeStr(r[0]).trim()) return;
      var sd = {};
      try { sd = JSON.parse(safeStr(r[8])) || {}; } catch(e) {}
      result.push({
        jobCardId:      safeStr(r[0]),
        orderRef:       safeStr(r[1]),
        workOrder:      safeStr(r[2]),
        store:          safeStr(r[3]),
        movement:       safeStr(r[4]),
        contractorId:   safeStr(r[5]),
        pairsIssued:    safeNum(r[6]),
        pairsReceived:  safeNum(r[7]),
        sizeBreakdown:  sd,
        issuedBy:       safeStr(r[9]),
        issuedAt:       safeStr(r[10]),
        expectedReturn: safeStr(r[11]),
        receivedAt:     safeStr(r[12]),
        status:         safeStr(r[13]),
        notes:          safeStr(r[14])
      });
    });
    if (filters) {
      if (filters.orderRef)     result = result.filter(function(c){ return c.orderRef     === filters.orderRef; });
      if (filters.store)        result = result.filter(function(c){ return c.store        === filters.store; });
      if (filters.status)       result = result.filter(function(c){ return c.status       === filters.status; });
      if (filters.contractorId) result = result.filter(function(c){ return c.contractorId === filters.contractorId; });
    }
    result.sort(function(a, b){ return a.issuedAt < b.issuedAt ? 1 : a.issuedAt > b.issuedAt ? -1 : 0; });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getOpenJobCards(store) {
  try {
    var all = store ? getJobCards({ store: store }) : getJobCards({});
    if (!Array.isArray(all)) return all;
    return all.filter(function(c){ return c.status === 'ISSUED' || c.status === 'PARTIAL'; });
  } catch(e) {
    return { success: false, error: e.message };
  }
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

    var ws = ensureJobCardsSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow()-1, 15).getValues();
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
          activities.forEach(function(a) { ratePerPair += safeNum(a.rate); });
        }
      } catch(ae) {}

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
        ratePerPair:    ratePerPair,
        totalAmount:    ratePerPair * pairsIssued,
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
      acts.forEach(function(a) { ratePerPair += safeNum(a.rate); });
      var amount = ratePerPair * pairs;

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
          totalPairs:     0,
          totalAmount:    0
        };
        batchOrder.push(paymentId);
      }

      var pairs  = safeNum(r[4]);
      var amount = safeNum(r[5]);
      batchMap[paymentId].lines.push({
        jobCardId:  safeStr(r[9]).trim(),
        orderRef:   safeStr(r[1]).trim(),
        customer:   safeStr(r[2]).trim(),
        department: safeStr(r[10]).trim(),
        pairs:      pairs,
        amount:     amount
      });
      batchMap[paymentId].totalPairs  += pairs;
      batchMap[paymentId].totalAmount += amount;
    });

    var result = batchOrder.map(function(pid) { return batchMap[pid]; });

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
        if (jobCardIds.indexOf(safeStr(r[0]).trim()) >= 0) jcWs.getRange(i + 2, 14).setValue('PAID');
      });
    }

    SpreadsheetApp.flush();
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
        if (jobCardIds.indexOf(safeStr(r[0]).trim()) >= 0) jcWs.getRange(i + 2, 14).setValue('COMPLETE');
      });
    }

    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
