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
  var result = { ok:true, dash:null, entry:null, wip:null, reqs:null };
  try {
    var user = getUserInfo();
    result.dash = getDashboardData();
    if (user.role === 'accounts') {
      result.entry = getEntryData();
    } else if (user.role === 'store') {
      result.wip = getWIPData();
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

  try {
    var wm = ss.getSheetByName('WEEKLY PAYMENT MASTER');
    if (wm) {
      weeklyPayout   = safeNum(wm.getRange('R55').getValue());
      approvalStatus = safeStr(wm.getRange('B57').getValue());
      weekEnding     = safeStr(wm.getRange('B2').getValue());
    }
  } catch(e) { Logger.log('WM error: ' + e.message); }

  var week = getCurrentWeek();
  if (!weekEnding) weekEnding = week.weekLabel;

  var orders = [];
  var redCount = 0, completeCount = 0;
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
          remaining:safeNum(r[7]), status:status
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
        remaining:remaining, activities:activities, _aqa:approvedByAct });
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
      oiS.getRange(4, 1, oiS.getLastRow()-3, 6).getValues().forEach(function(r) {
        var sn = safeStr(r[1]), ls = safeNum(r[5]);
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
    var deptStatus = getDeptStatus(art.sheet);
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

function getWIPData() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var wr = ss.getSheetByName('WIP_RECONCILIATION');
  var rows = [];
  try {
    if (wr && wr.getLastRow() > 4) {
      wr.getRange(5, 1, wr.getLastRow()-4, 8).getValues().forEach(function(r,i){
        if (!r[0]) return;
        rows.push({ rowNum:i+5, sheet:safeStr(r[0]), article:safeStr(r[1]),
          activity:safeStr(r[2]), produced:safeStr(r[3]),
          paid:safeNum(r[4]), difference:safeStr(r[5]), status:safeStr(r[6]) });
      });
    }
  } catch(e) {}
  return { rows:rows };
}

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
    var mr = ss.getSheetByName('MASTER_RATES');
    if (mr && mr.getLastRow() > 3)
      mr.getRange(4, 2, mr.getLastRow()-3, 1).getValues().forEach(function(r,i){ mrMap[4+i] = safeStr(r[0]); });
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

function getPaymentSubmissions() {
  var user = getUserInfo();
  if (user.role !== 'admin') return { submissions:[], pmMap:{} };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
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
        lines.push('TS: ' + (pl.tsNumber||'?') + ' | Lot: ' + (pl.lotSize||0) + ' pairs | Del: ' + (pl.deliveryDate||'?'));
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
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + String(lastRow-3).padStart ? String(lastRow-3).padStart(3,'0') : ('00'+(lastRow-3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, type, details, 'PENDING', '', '', 'No', ''
    ]]);
    SpreadsheetApp.flush();
    notifyNewRequest_(reqId, type, details, user.name, now);
    return { success:true, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function processRequest(rowNum, action, notes) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var rq  = ss.getSheetByName('REQUESTS');
    var row = rq.getRange(rowNum, 1, 1, 10).getValues()[0];
    if (safeStr(row[5]) === 'APPROVED' || safeStr(row[5]) === 'REJECTED') return { success: false, error: 'Already processed' };
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
    if (action === 'APPROVE' && safeStr(row[3]) === 'PAYMENT') {
      try {
        var wm2 = ss.getSheetByName('WEEKLY PAYMENT MASTER');
        if (wm2) wm2.getRange('B57').setValue(user.name + ' — ' + now);
      } catch(pe) { Logger.log('PAYMENT approval error: ' + pe.message); }
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

function getPaymentList() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var paymentMethods = {};
  try {
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() > 3) {
      mc.getRange(4, 2, mc.getLastRow()-3, 2).getValues().forEach(function(r){
        if (r[0]) paymentMethods[safeStr(r[0])] = safeStr(r[1]) || 'Cash';
      });
    }
  } catch(e) {}

  var contractors = {};
  var totalPayout = 0;
  var weekEnding = '', approvedBy = '';

  try {
    var wm = ss.getSheetByName('WEEKLY PAYMENT MASTER');
    if (wm) {
      weekEnding = safeStr(wm.getRange('B2').getValue());
      approvedBy = safeStr(wm.getRange('B57').getValue());
    }
  } catch(e) {}

  ss.getSheets().filter(isArtSheet).forEach(function(ws) {
    try {
      var article = safeStr(ws.getRange('B2').getValue());
      var actData = ws.getRange(5, 2, 45, 8).getValues();
      actData.forEach(function(r) {
        var activity   = safeStr(r[0]);
        var contractor = safeStr(r[1]);
        var qty        = safeNum(r[2]);
        var rate       = safeNum(r[3]);
        var comm       = safeNum(r[4]);
        var conveyance = safeNum(r[6]);
        var total      = safeNum(r[7]);
        if (!contractor || !qty) return;
        var amount = total || ((qty*rate) + (qty*comm) + conveyance);
        if (!amount) return;
        if (!contractors[contractor]) {
          contractors[contractor] = { name:contractor,
            method:paymentMethods[contractor]||'Cash', total:0, details:[] };
        }
        contractors[contractor].total += amount;
        contractors[contractor].details.push({ article:article,
          activity:activity, qty:qty, rate:rate, amount:amount });
        totalPayout += amount;
      });
    } catch(e) {}
  });

  var list = Object.keys(contractors).map(function(k){ return contractors[k]; })
    .sort(function(a,b){ return b.total - a.total; });

  return { list:list, totalPayout:totalPayout,
    weekEnding:weekEnding, approvedBy:approvedBy,
    approved:approvedBy !== '' };
}

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

function archiveWeek() {
  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var wm  = ss.getSheetByName('WEEKLY PAYMENT MASTER');
  var approval = safeStr(wm.getRange('B57').getValue());
  if (!approval || approval.trim() === '')
    return { success:false, error:'Week not approved yet. Ayush must approve first.' };
  var ph  = ss.getSheetByName('PAYMENT_HISTORY');
  var ot  = ss.getSheetByName('ORDER_TRACKER');
  var artSheets = ss.getSheets().filter(isArtSheet);
  var archiveDate = new Date();
  var weekEnding  = wm.getRange('B2').getValue() || archiveDate;
  var rowsAdded = 0;
  artSheets.forEach(function(ws) {
    try {
      var pairsThisWk = safeNum(ws.getRange('Q2').getValue());
      if (!pairsThisWk) return;
      var artName  = safeStr(ws.getRange('B2').getValue());
      var customer = safeStr(ws.getRange('E2').getValue());
      ws.getRange(54, 2, 30, 8).getValues().forEach(function(r) {
        var contractor = safeStr(r[0]), total = safeNum(r[7]);
        if (contractor && total > 0) {
          ph.getRange(ph.getLastRow()+1, 1, 1, 8).setValues([[
            weekEnding, artName, customer, contractor,
            pairsThisWk, total, approval, archiveDate
          ]]);
          rowsAdded++;
        }
      });
      if (ot && ot.getLastRow() > 3) {
        ot.getRange(4, 1, ot.getLastRow()-3, 5).getValues().forEach(function(r,i){
          if (safeStr(r[0]) === ws.getName())
            ot.getRange(4+i, 5).setValue(safeNum(r[4]) + pairsThisWk);
        });
      }
    } catch(e) { Logger.log('Archive error: ' + e.message); }
  });
  return { success:true, rowsAdded:rowsAdded };
}

function clearWeekForNext() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  ss.getSheets().filter(isArtSheet).forEach(function(ws){
    ws.getRange('D5:D49').clearContent();
    ws.getRange('H5:H49').clearContent();
    ws.getRange('J5:J49').clearContent();
  });
  var wm = ss.getSheetByName('WEEKLY PAYMENT MASTER');
  if (wm) wm.getRange('B57').clearContent();
  return { success:true };
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

function approvePeriodPayment(periodId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only admin can approve' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy HH:mm');
    var rq = ss.getSheetByName('REQUESTS');
    var approvedRows = 0, histRows = 0;
    if (rq && rq.getLastRow() > 3) {
      var reqData = rq.getRange(4, 1, rq.getLastRow()-3, 10).getValues();
      reqData.forEach(function(r, i) {
        if (safeStr(r[3]) !== 'PAYMENT_SUBMISSION' || safeStr(r[5]).toUpperCase() !== 'PENDING') return;
        try {
          var pl = JSON.parse(safeStr(r[4]));
          if (!pl || !pl.sheet || safeStr(pl.periodId) !== safeStr(periodId)) return;
          var ws = ss.getSheetByName(pl.sheet);
          if (!ws) return;
          var article = safeStr(ws.getRange('B2').getValue());
          var customer = safeStr(ws.getRange('E2').getValue());
          var ph = ss.getSheetByName('PAYMENT_HISTORY');
          if (!ph) { ph = ss.insertSheet('PAYMENT_HISTORY'); ph.getRange(1,1,1,8).setValues([['PeriodID','Article','Customer','Contractor','Qty','Amount','ApprovedBy','Date']]); }
          ws.getRange(5, 1, 45, 12).getValues().forEach(function(ar, ai) {
            if (!safeStr(ar[1]).trim() || safeNum(ar[0]) <= 0) return;
            if (safeStr(ar[10]) !== safeStr(periodId)) return;
            if (safeStr(ar[11]).toUpperCase() !== 'SUBMITTED') return;
            ws.getRange('L'+(ai+5)).setValue('APPROVED');
            var qty = safeNum(ar[3]), total = safeNum(ar[8]);
            if (qty > 0 && total > 0) {
              ph.appendRow([periodId, article, customer, safeStr(ar[2]), qty, total, user.name+' — '+now, new Date()]);
              histRows++;
            }
            approvedRows++;
          });
          rq.getRange(i+4, 6).setValue('APPROVED');
          rq.getRange(i+4, 7).setValue('Payment approved');
          rq.getRange(i+4, 8).setValue(now);
          rq.getRange(i+4, 9).setValue('Yes');
        } catch(e2) { Logger.log('approvePeriodPayment: '+e2.message); }
      });
    }
    try {
      var ph2 = ss.getSheetByName('PAYMENT_HISTORY');
      if (ph2 && ph2.getLastRow() > 1) {
        ph2.getRange(2, 1, ph2.getLastRow()-1, 8).getValues().forEach(function(r, i) {
          if (safeStr(r[0]) === safeStr(periodId) && !safeStr(r[6]))
            ph2.getRange(i+2, 7).setValue(user.name+' — '+now);
        });
      }
    } catch(e3) {}
    SpreadsheetApp.flush();
    return { success:true, approvedRows:approvedRows, histRows:histRows };
  } catch(e) { return { success:false, error:e.message }; }
}

function getActivitiesFromTS(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var tsNumber = '';
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() > 3) {
      var oiData = oi.getRange(4, 1, oi.getLastRow()-3, 6).getValues();
      for (var i = 0; i < oiData.length; i++) {
        if (safeStr(oiData[i][1]) === sheetName) { tsNumber = safeStr(oiData[i][5]); break; }
      }
    }
    if (!tsNumber) return { success:false, error:'No TS linked to this article' };
    var tm = ss.getSheetByName('TS_MASTER');
    if (!tm || tm.getLastRow() < 2) return { success:false, error:'TS_MASTER not found' };
    var tmData = tm.getRange(2, 1, tm.getLastRow()-1, 7).getValues();
    for (var j = 0; j < tmData.length; j++) {
      if (safeStr(tmData[j][0]) === tsNumber) {
        var raw = safeStr(tmData[j][6]);
        if (!raw) return { success:true, activities:[] };
        try { return { success:true, activities:JSON.parse(raw) }; }
        catch(pe) { return { success:false, error:'Invalid activities JSON in TS' }; }
      }
    }
    return { success:false, error:'TS not found: ' + tsNumber };
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

function rejectRequest(reqId) {
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
        rq.getRange(i + 4, 7).setValue('Rejected');
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

function requestActivityRateEdit(rowIndex, newRate, newComm) {
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var payload = JSON.stringify({ rowIndex:rowIndex, newRate:newRate, newComm:newComm });
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, 'RATE_EDIT', payload, 'PENDING', '', '', 'No', ''
    ]]);
    SpreadsheetApp.flush();
    notifyNewRequest_(reqId, 'RATE_EDIT', payload, user.name, now);
    return { success:true, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function requestActivitySetup(payload) {
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rq = ss.getSheetByName('REQUESTS');
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var items = Array.isArray(payload) ? payload : [payload];
    var lastReqId = '';
    var createdReqs = [];
    items.forEach(function(item) {
      var lastRow = Math.max(rq.getLastRow(), 3) + 1;
      var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
      rq.getRange(lastRow, 1, 1, 10).setValues([[
        reqId, now, user.name, 'ACTIVITY_SETUP', JSON.stringify(item), 'PENDING', '', '', 'No', ''
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
    var tz = Session.getScriptTimeZone();
    var data = ph.getRange(2, 1, ph.getLastRow()-1, 8).getValues();
    var records = [];
    data.forEach(function(r) {
      var periodId = safeStr(r[0]);
      if (!periodId) return;
      var dv = r[7];
      records.push({
        periodId: periodId,
        sheet: safeStr(r[1]),
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
    var tsNumber = 'TS-' + (season||'SS26') + '-' + seq;
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
    var bomNumber = 'BOM-' + new Date().getFullYear() + '-' + bomSeq;
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
      oi.getRange(oiRow, 1, 1, 12).setValues([[
        bomNumber, artSheet, safeStr(payload.styleName), safeStr(payload.color),
        safeStr(payload.buyer), safeStr(payload.tsNumber), '', '', lotSize,
        now, 'Active', safeStr(payload.poNumber)
      ]]);
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
