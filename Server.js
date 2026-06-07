/**
 * ADEES EXPORTS — FACTORY OS
 * Final Server.gs — Clean, no DASHBOARD dependency
 */

const ROLES = {
  "ayush@adeesexports.com":   "admin",
  "aneesh@adeesexports.com":  "accounts",
  "admin@adeesexports.com":   "store",
};

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('FactoryOS — Adees Exports')
    .addMetaTag('viewport','width=device-width, initial-scale=1')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getUserInfo() {
  var email = Session.getActiveUser().getEmail();
  var role  = ROLES[email] || 'viewer';
  var first = email.split('@')[0];
  var name  = first.charAt(0).toUpperCase() + first.slice(1);
  Logger.log('User: ' + email + ' | Role: ' + role);
  return { email:email, role:role, name:name };
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
        var status = safeStr(r[8]);
        orders.push({
          sheet:r[0], article:safeStr(r[1]), customer:safeStr(r[2]),
          orderQty:safeNum(r[3]), prior:safeNum(r[4]),
          thisWeek:safeNum(r[5]), cumul:safeNum(r[6]),
          remaining:safeNum(r[7]), status:status
        });
        if (status.indexOf('OVER') > -1)     redCount++;
        if (status.indexOf('COMPLETE') > -1) completeCount++;
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

  return {
    weeklyPayout:weeklyPayout, approvalStatus:approvalStatus,
    weekEnding:weekEnding, orders:orders, redCount:redCount,
    completeCount:completeCount, mismatches:mismatches,
    pendingCount:pendingCount, totalOrders:orders.length,
    costPerPair:[]
  };
}

function getEntryData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

      var actData = ws.getRange(5, 1, 45, 10).getValues();
      var activities = [];
      actData.forEach(function(r, i) {
        var act = safeStr(r[1]);
        if (act && act.trim() && !act.match(/^[-=]/)) {
          activities.push({
            row:i+5, activity:act.trim(),
            contractor:safeStr(r[2]), qty:safeNum(r[3]),
            rate:safeNum(r[4]), comm:safeNum(r[5]), total:safeNum(r[8])
          });
        }
      });
      articles.push({ sheet:name, article:article, customer:customer,
        orderQty:orderQty, status:status, thisWeek:thisWeek,
        remaining:remaining, activities:activities });
    } catch(e) { Logger.log('ART error ' + ws.getName() + ': ' + e.message); }
  });

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
    var mr = ss.getSheetByName('MASTER_RATES');
    if (mr && mr.getLastRow() > 3) {
      mr.getRange(4, 2, mr.getLastRow()-3, 5).getValues().forEach(function(r, i){
        if (r[0] && r[4] && safeStr(r[4]).toLowerCase() !== 'tbd')
          masterActivities.push({ name:safeStr(r[0]), section:safeStr(r[1]),
            minRate:safeNum(r[2]), maxRate:safeNum(r[3]), stdRate:safeNum(r[4]),
            rowIndex: 4 + i });
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

  var week = null;
  try {
    var pp = ss.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var ppData = pp.getRange(2, 1, pp.getLastRow()-1, 7).getValues();
      for (var pi = 0; pi < ppData.length; pi++) {
        if (safeStr(ppData[pi][6]).trim().toUpperCase() === 'OPEN') {
          week = { weekLabel: safeStr(ppData[pi][2]), weekStart: safeStr(ppData[pi][3]), weekEnd: safeStr(ppData[pi][4]) };
          break;
        }
      }
    }
  } catch(e3) {}
  if (!week) week = getCurrentWeek();
  return { articles:articles, contractors:contractors, masterActivities:masterActivities, week:week };
}

function saveEntry(sheetName, row, contractor, qty, conveyance, remarks) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success:false, error:'Not authorised' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName(sheetName);
    if (!ws) return { success:false, error:'Sheet not found' };
    if (contractor) ws.getRange('C'+row).setValue(contractor);
    ws.getRange('D'+row).setValue(qty||0);
    if (conveyance) ws.getRange('H'+row).setValue(conveyance);
    if (remarks)    ws.getRange('J'+row).setValue(remarks);
    SpreadsheetApp.flush();
    return { success:true,
      lotStatus: safeStr(ws.getRange('M7').getValue()),
      thisWeek:  safeNum(ws.getRange('M4').getValue()),
      remaining: safeNum(ws.getRange('M6').getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}

function getWIPData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var wr = ss.getSheetByName('WIP_RECONCILIATION');
    wr.getRange('D'+rowNum).setValue(produced);
    SpreadsheetApp.flush();
    return { success:true,
      status: safeStr(wr.getRange('G'+rowNum).getValue()),
      diff:   safeStr(wr.getRange('F'+rowNum).getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPendingRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  return { requests:requests };
}

function submitRequest(type, details) {
  var user = getUserInfo();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rq = ss.getSheetByName('REQUESTS');
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + String(lastRow-3).padStart ? String(lastRow-3).padStart(3,'0') : ('00'+(lastRow-3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, type, details, 'PENDING', '', '', 'No', ''
    ]]);
    SpreadsheetApp.flush();
    return { success:true, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function processRequest(rowNum, action, notes) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var rq  = ss.getSheetByName('REQUESTS');
    var row = rq.getRange(rowNum, 1, 1, 10).getValues()[0];
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(rowNum, 6).setValue(action === 'REJECT' ? 'REJECTED' : 'APPROVED');
    rq.getRange(rowNum, 7).setValue(notes || (action==='APPROVE'?'Approved':'Rejected'));
    rq.getRange(rowNum, 8).setValue(now);
    rq.getRange(rowNum, 9).setValue('Yes');
    var sheetCreated = '';
    if (action === 'APPROVE' && safeStr(row[3]) === 'NEW ORDER') {
      sheetCreated = createNewArtSheet(safeStr(row[4]));
      rq.getRange(rowNum, 10).setValue(sheetCreated);
    }
    if (action === 'APPROVE' && safeStr(row[3]) === 'ACTIVITY_SETUP') {
      try {
        var setupPayload = JSON.parse(safeStr(row[4]));
        Logger.log('ACTIVITY_SETUP: sheet=' + (setupPayload&&setupPayload.sheet) + ' activities=' + (setupPayload&&setupPayload.activities?setupPayload.activities.length:0));
        if (setupPayload && setupPayload.sheet && setupPayload.activities)
          saveActivitySetup(setupPayload.sheet, setupPayload.activities);
      } catch(pe) { Logger.log('ACTIVITY_SETUP error: ' + pe.message); }
    }
    SpreadsheetApp.flush();
    return { success:true, action:action, sheetCreated:sheetCreated };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveWeek(initials) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var wm  = ss.getSheetByName('WEEKLY PAYMENT MASTER');
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    wm.getRange('B57').setValue(initials + ' — ' + now);
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function getPaymentList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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
        var weekEnding  = safeStr(r[0]);
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


function rejectRequest(reqId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can reject' };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  Logger.log('WIPE COMPLETE');
  return { success:true };
}

function saveActivitySetup(sheet, activities) {
  Logger.log('saveActivitySetup: sheet=' + sheet + ' count=' + (activities?activities.length:0));
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ws = ss.getSheetByName(sheet);
    if (!ws) return { success:false, error:'Sheet not found: ' + sheet };
    ws.getRange(5, 1, 45, 6).clearContent();
    var rows = activities.map(function(a, i) {
      return [i + 1, safeStr(a.activityName), '', '', safeNum(a.rate), safeNum(a.comm)];
    });
    if (rows.length > 0) ws.getRange(5, 1, rows.length, 6).setValues(rows);
    SpreadsheetApp.flush();
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function getActivitySetup(sheet) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rq = ss.getSheetByName('REQUESTS');
    var lastRow = Math.max(rq.getLastRow(), 3) + 1;
    var reqId = 'REQ-' + ('00' + (lastRow - 3)).slice(-3);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var payload = JSON.stringify({ rowIndex:rowIndex, newRate:newRate, newComm:newComm });
    rq.getRange(lastRow, 1, 1, 10).setValues([[
      reqId, now, user.name, 'RATE_EDIT', payload, 'PENDING', '', '', 'No', ''
    ]]);
    SpreadsheetApp.flush();
    return { success:true, reqId:reqId };
  } catch(e) { return { success:false, error:e.message }; }
}

function approveRateEdit(requestId) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can approve' };
  Logger.log('approveRateEdit: requestId=' + requestId);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    Logger.log('approveRateEdit: rowIndex=' + payload.rowIndex + ' newRate=' + payload.newRate + ' newComm=' + payload.newComm);
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ph = ss.getSheetByName('PAYMENT_HISTORY');
    if (!ph || ph.getLastRow() < 4) return { success:true, periods:[] };
    var data = ph.getRange(4, 1, ph.getLastRow()-3, 8).getValues();
    var periodMap = {};
    data.forEach(function(r) {
      var weekEnding = safeStr(r[0]);
      if (!weekEnding) return;
      var periodId = 'P-' + weekEnding.replace(/[^a-zA-Z0-9]/g, '');
      if (!periodMap[periodId]) {
        periodMap[periodId] = { periodId:periodId, weekLabel:weekEnding,
          totalAmount:0, status:safeStr(r[6]) ? 'APPROVED' : 'PENDING' };
      }
      periodMap[periodId].totalAmount += safeNum(r[5]);
    });
    var periods = Object.keys(periodMap).map(function(k){ return periodMap[k]; });
    return { success:true, periods:periods };
  } catch(e) { return { success:false, error:e.message, periods:[] }; }
}

function createArtTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
  Logger.log('ART-TEMPLATE created');
  return 'ART-TEMPLATE created';
}