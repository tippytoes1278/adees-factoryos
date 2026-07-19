function getEntryData(periodId, ss) {
  if (!ss && !periodId) {
    try {
      var _ec = CacheService.getScriptCache();
      var _ev = _ec.get('entryData_' + CONFIG.ENV);
      if (_ev) return JSON.parse(_ev);
    } catch(ce) {}
  }
  ensureCurrentPeriod();
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  // Read REQUESTS once and reuse across every pass below (was ~6 separate reads).
  var _REQ = [];
  try {
    var _rqAll = ss.getSheetByName('REQUESTS');
    if (_rqAll && _rqAll.getLastRow() > 3)
      _REQ = _rqAll.getRange(4, 1, _rqAll.getLastRow()-3, 6).getValues();
  } catch(e) {}
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
      var _mCol    = ws.getRange('M4:M7').getValues();  // one read: M4,M5,M6,M7
      var status   = safeStr(_mCol[3][0]);   // M7
      var thisWeek = safeNum(_mCol[0][0]);   // M4
      var remaining= safeNum(_mCol[2][0]);   // M6

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
    _REQ.forEach(function(rr) {
      if (safeStr(rr[3]) !== 'ACTIVITY_SETUP' || safeStr(rr[5]).toUpperCase() !== 'PENDING') return;
      try {
        var setupPl = JSON.parse(safeStr(rr[4]));
        if (setupPl && setupPl.sheet)
          for (var ai = 0; ai < articles.length; ai++)
            if (articles[ai].sheet === setupPl.sheet) articles[ai].hasPendingSetup = true;
      } catch(pe) {}
    });
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
        _REQ.forEach(function(rr) {
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
    {
      _REQ.forEach(function(rr) {
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
  var _batchReqsData = _REQ;
  var _batchDsData = [];
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
    if (_REQ.length) {
      var pendingEditMap = {};
      _REQ.forEach(function(r) {
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
    _REQ.forEach(function(rr) {
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
  } catch(e5) {}
  if (!week) week = getCurrentWeek();
  var _entryResult = { articles:articles, contractors:contractors, masterActivities:masterActivities, pendingActivities:pendingActivities, pendingActsCount:pendingActsCount, week:week, periods:periods };
  if (!periodId) {
    try {
      CacheService.getScriptCache()
        .put('entryData_' + CONFIG.ENV, JSON.stringify(_entryResult), 300);
    } catch(ce) {}
  }
  return _entryResult;
}

// Lightweight article list for the store/floor screens.
// Same authoritative source as getEntryData (each ART sheet's B2:J2 header,
// orderQty = H2), but ONE read per sheet and none of the activities / REQUESTS /
// masterActivities / deptStatus work. Store screens only ever use articles[].
// Returns the same { articles:[...] } shape store render code already expects.
function getArticlesLite(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var bomMap = {};
  try {
    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() > 3)
      oi.getRange(4, 1, oi.getLastRow()-3, 2).getValues().forEach(function(r) {
        var sn = safeStr(r[1]); if (sn) bomMap[sn] = safeStr(r[0]);
      });
  } catch(e) {}
  var articles = [];
  ss.getSheets().filter(isArtSheet).forEach(function(ws) {
    try {
      var name = ws.getName();
      var hdr  = ws.getRange('B2:J2').getValues()[0];  // B2..J2 in one read
      articles.push({
        sheet:    name,
        article:  safeStr(hdr[0]),   // B2
        customer: safeStr(hdr[3]),   // E2
        orderQty: safeNum(hdr[6]),   // H2 — authoritative order/lot qty
        bom:      bomMap[name] || ''
      });
    } catch(e) { Logger.log('getArticlesLite ' + ws.getName() + ': ' + e.message); }
  });
  return { articles: articles };
}

// One round-trip bundle for the store screens (WIP / Job Cards / Floor):
// lite articles + contractors + active enrollments + job cards. Cached 300s under
// storeScreenData_<ENV>; the cache is removed by issueJobCard / issueJobCardBatch /
// receiveJobCard so a write is reflected on the next load.
function getStoreScreenData() {
  try {
    var _cc = CacheService.getScriptCache();
    var _cv = _cc.get('storeScreenData_' + CONFIG.ENV);
    if (_cv) return JSON.parse(_cv);
  } catch(ce) {}
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var result = { articles: [], contractors: [], enrollments: [], jobCards: [] };
  try { var lite = getArticlesLite(ss); result.articles = (lite && lite.articles) || []; } catch(e) {}
  try { result.contractors = getContractors(ss); } catch(e) {}
  try { result.enrollments = getEnrollments({status:'ACTIVE'}, ss); } catch(e) {}
  try { result.jobCards = getJobCards({}, ss); } catch(e) {}
  try {
    CacheService.getScriptCache()
      .put('storeScreenData_' + CONFIG.ENV, JSON.stringify(result), 300);
  } catch(ce) {}
  return result;
}

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

function getApprovedActivitiesForArticle(sheetName, ss) {
  try {
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
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
    // Block if already APPROVED for this order+dept
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var rqChk = ss.getSheetByName('REQUESTS');
    if (rqChk && rqChk.getLastRow() > 3) {
      var rqChkData = rqChk.getRange(4, 1, rqChk.getLastRow()-3, 6).getValues();
      var itemsArr = Array.isArray(payload) ? payload : [payload];
      for (var _ai = 0; _ai < itemsArr.length; _ai++) {
        var _item = itemsArr[_ai];
        for (var _ri = 0; _ri < rqChkData.length; _ri++) {
          var _r = rqChkData[_ri];
          if (safeStr(_r[3]) !== 'ACTIVITY_SETUP') continue;
          if (safeStr(_r[5]).toUpperCase() !== 'APPROVED') continue;
          try {
            var _pl = JSON.parse(safeStr(_r[4]));
            if (_pl && safeStr(_pl.sheet) === safeStr(_item.sheet) &&
                safeStr(_pl.dept) === safeStr(_item.dept)) {
              return {
                success: false,
                error: 'Activities already approved for ' +
                       safeStr(_item.dept) + ' department on this order. ' +
                       'Contact Ayush if changes are needed.'
              };
            }
          } catch(pe) {}
        }
      }
    }
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
    try { CacheService.getScriptCache().remove('entryData_' + CONFIG.ENV); } catch(ce) {}
    return { success:true, reqId:lastReqId, count:items.length };
  } catch(e) { return { success:false, error:e.message }; }
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
    var ma = ss.getSheetByName('MASTER_ACTIVITIES');
    if (!ma) return { success:false, error:'MASTER_ACTIVITIES sheet not found' };
    ma.getRange(payload.rowIndex, 3).setValue(payload.newRate);
    ma.getRange(payload.rowIndex, 4).setValue(payload.newComm);
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    rq.getRange(targetRow, 6).setValue('APPROVED');
    rq.getRange(targetRow, 8).setValue(now);
    rq.getRange(targetRow, 9).setValue('Yes');
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('entryData_' + CONFIG.ENV); } catch(ce) {}
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
