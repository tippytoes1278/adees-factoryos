// S.6: locked — writer
function deleteOrder(sheetName) {
  var user = getUserInfo();
  if (user.role !== 'admin') return { success:false, error:'Only Ayush can delete orders' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
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
    SpreadsheetApp.flush();
    if (failures.length) {
      Logger.log('deleteOrder('+sheetName+'): sheet deleted but row cleanup failed — '+failures.join('; '));
      return { success:true, warning:'Sheet deleted, but stale rows remain in: '+failures.join('; ') };
    }
    return { success:true };
    } catch(e) { return { success:false, error:e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// Next free ART-### computed from live sheet names (monotonic: max+1,
// survives deletes, never reuses a number). Single source of truth for
// ART numbering — do not compute ART numbers anywhere else.
function nextArtName_(ss) {
  var nums = ss.getSheets().filter(isArtSheet)
    .map(function(s){ return parseInt(s.getName().replace('ART-',''))||0; });
  var n = String(Math.max.apply(null,[0].concat(nums)) + 1);
  while (n.length < 3) n = '0' + n;
  return 'ART-' + n;
}

// S.2: the legacy pipe-delimited creation path (createNewArtSheet/updateTrackers)
// was deleted — unlocked, computed ART numbers pre-write, and had no callers.
// createOrder (locked, live-sheet numbering) is the ONLY creation path.

// S.6: locked — rebuild (returns a string; errors propagate, lock still released)
function createArtTemplate() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
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
  } finally {
    lock.releaseLock();
  }
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

// S.6: locked — sequence
function createTS(styleName, category, season, activities) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
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
  } finally {
    lock.releaseLock();
  }
}

function createOrder(payload) {
  if (safeNum(payload.lotSize) <= 0) return { success:false, error:'Lot size must be greater than 0' };
  // Script lock: on 22-Jul two same-minute NEW_ORDER approvals both read
  // max=ART-023 (processRequest has no lock; only submitRequest does) and
  // collided on ART-024. Serialise all order/sheet creation.
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch(lockErr) {
    return { success:false, error:'Another order is being created right now — wait a few seconds and approve again' };
  }
  var ss = null, ws = null;
  try {
    ss = SpreadsheetApp.openById(SHEET_ID);
    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, 'dd-MMM-yyyy');
    var oi = ss.getSheetByName('ORDER_INDEX');
    // WO sequence = highest existing WO number + 1 (monotonic — survives deletes,
    // never reuses a number). Count-based numbering duplicated after a delete.
    var maxWoSeq = 0;
    if (oi && oi.getLastRow() > 3) {
      oi.getRange(4, 1, oi.getLastRow() - 3, 1).getValues().forEach(function(r) {
        var m = safeStr(r[0]).match(/^WO-\d{4}-(\d+)$/);
        if (m) { var n = parseInt(m[1], 10) || 0; if (n > maxWoSeq) maxWoSeq = n; }
      });
    }
    var bomSeq = String(maxWoSeq + 1);
    while (bomSeq.length < 3) bomSeq = '0' + bomSeq;
    var bomNumber = 'WO-' + new Date().getFullYear() + '-' + bomSeq;
    var artSheet = nextArtName_(ss);
    var template = ss.getSheetByName('ART-TEMPLATE') || ss.getSheetByName('ART-001');
    ws = template.copyTo(ss);
    // Collision-proof rename (belt-and-braces on top of the lock): if the
    // computed name is taken or setName throws, recompute from live sheet
    // names and retry. Never leaves HEAD holding a "Copy of ART-TEMPLATE".
    var renamed = false;
    for (var _na = 0; _na < 5 && !renamed; _na++) {
      if (_na > 0) { Utilities.sleep(200); artSheet = nextArtName_(ss); }
      if (ss.getSheetByName(artSheet)) continue;
      try { ws.setName(artSheet); renamed = true; } catch(nameErr) {}
    }
    if (!renamed) throw new Error('Could not allocate a free ART number (last tried ' + artSheet + '). Nothing was created — approve again.');
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
      // Columns A–K match the classic ORDER_INDEX layout; L–Q are new extended fields; R–AO are size values
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
    SpreadsheetApp.flush();
    return { success:true, bomNumber:bomNumber, artSheet:artSheet };
  } catch(e) {
    // If the template copy never got renamed, delete it so no stranded
    // "Copy of ART-TEMPLATE" is left behind (22-Jul incident).
    try {
      if (ss && ws && ws.getName().indexOf('Copy of') === 0) ss.deleteSheet(ws);
    } catch(cleanErr) {}
    Logger.log('createOrder error: ' + e.message);
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
  }
}

// Read-only diagnostic. Run from the Apps Script editor: reconciles ART-*
// sheets against ORDER_INDEX, flags orphans on both sides, stranded
// "Copy of" sheets, and numbering gaps (old D4 orphan question).
function auditArtSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var report = { artSheets:[], orphanSheets:[], orphanIndexRows:[], strandedCopies:[], gaps:[] };
  var names = ss.getSheets().map(function(s){ return s.getName(); });
  names.forEach(function(n) { if (n.indexOf('Copy of') === 0) report.strandedCopies.push(n); });
  var arts = names.filter(function(n){ return n.indexOf('ART-') === 0 && n !== 'ART-TEMPLATE'; });
  report.artSheets = arts;
  var idxSheets = {};
  var oi = ss.getSheetByName('ORDER_INDEX');
  if (oi && oi.getLastRow() > 3) {
    oi.getRange(4, 1, oi.getLastRow()-3, 2).getValues().forEach(function(r) {
      var sn = safeStr(r[1]);
      if (sn) idxSheets[sn] = safeStr(r[0]);
    });
  }
  arts.forEach(function(n){ if (!idxSheets[n]) report.orphanSheets.push(n); });
  Object.keys(idxSheets).forEach(function(n){
    if (arts.indexOf(n) < 0) report.orphanIndexRows.push(n + ' (' + idxSheets[n] + ')');
  });
  var nums = arts.map(function(n){ return parseInt(n.replace('ART-',''))||0; })
    .sort(function(a,b){ return a-b; });
  var maxN = nums.length ? nums[nums.length-1] : 0;
  for (var i = 1; i <= maxN; i++) {
    if (nums.indexOf(i) < 0) report.gaps.push('ART-' + ('00'+i).slice(-3));
  }
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}

function getOrderProgress(artSheet) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName(artSheet);
    if (!ws) return { error: 'Sheet not found: ' + artSheet };
    var lotSize = safeNum(ws.getRange('H2').getValue());

    // Progress per stage = pairs RECEIVED back through that stage's job cards
    // (job-card model). One card per department, so no double counting.
    var MOVEMENT_STAGE = {
      'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter',
      'Upper IN':'lasting','Lasting IN':'lasting','Packing IN':'finish','Dispatch IN':'dispatch'
    };
    var sq = {cutting:0,prep:0,fitter:0,lasting:0,finish:0,dispatch:0};
    var jcs = getJobCards({ orderRef: artSheet });
    if (Array.isArray(jcs)) {
      jcs.forEach(function(jc) {
        var stage = MOVEMENT_STAGE[safeStr(jc.movement).trim()];
        if (!stage) return;
        if (safeStr(jc.status).toUpperCase() === 'CANCELLED') return;
        sq[stage] += safeNum(jc.pairsReceived);
      });
    }
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
      if (safeStr(r[0]).trim() === ref || safeStr(r[1]).trim() === ref) {
        var rawSizeCol = r[16]; // col Q (index 16) = raw sizeBreakdown JSON written by createOrder
        var sizes = {};
        var totalQty = 0;
        try {
          var parsed = JSON.parse(safeStr(rawSizeCol));
          if (parsed && typeof parsed === 'object') {
            Object.keys(parsed).forEach(function(k) {
              var qty = safeNum(parsed[k]);
              if (qty > 0) { sizes[k] = qty; totalQty += qty; }
            });
          }
        } catch(pe) {}
        // Fallback: legacy orders that only have numeric size columns (R–AO)
        if (Object.keys(sizes).length === 0) {
          _SIZES_RANGE.forEach(function(s, si) {
            var qty = safeNum(r[17 + si]);
            if (qty > 0) { sizes[String(s)] = qty; totalQty += qty; }
          });
        }
        return { success: true, sizes: sizes, totalQty: totalQty };
      }
    }
    return { success: false, error: 'Order not found' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Per-size balance for the Issue form: approved size run minus what's already been
// issued at this movement/stage. Returns remaining per size + whether the caller
// may override the cap (admin only).
function getOrderSizeBalance(orderRef, movement) {
  var base = getOrderSizes(orderRef);
  if (!base || !base.success) return base || { success:false, error:'Order not found' };
  var ordered = base.sizes || {};
  var ref = safeStr(orderRef).trim(), mv = safeStr(movement).trim();

  // Unified stage vocabulary: cutting/prep/fitter/lasting/finish/dispatch.
  var STAGE_ORDER = ['Cutting','Preparation','Fitter','Lasting','Packing','Dispatch'];
  var STAGE_DEPT_KEY = {'Cutting':'cutting','Preparation':'prep','Fitter':'fitter','Lasting':'lasting','Packing':'finish','Dispatch':'dispatch'};
  var STAGE_OWN_MOVEMENTS = {'Cutting':['Cutting IN'],'Preparation':['Preparation IN'],'Fitter':['Fitter IN'],'Lasting':['Upper IN','Lasting IN'],'Packing':['Packing IN'],'Dispatch':['Dispatch IN']};
  var MOVEMENT_TO_STAGE = {'Cutting IN':'Cutting','Preparation IN':'Preparation','Fitter IN':'Fitter','Upper IN':'Lasting','Lasting IN':'Lasting','Packing IN':'Packing','Dispatch IN':'Dispatch'};

  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Nearest active predecessor stage for this order (skips skipped departments).
  var predMovements = null;
  var currentStage = MOVEMENT_TO_STAGE[mv] || '';
  if (currentStage) {
    var idx = STAGE_ORDER.indexOf(currentStage);
    var actRes = getApprovedActivitiesForArticle(ref, ss);
    var activeDepts = {};
    if (actRes && actRes.success && Array.isArray(actRes.activities)) {
      actRes.activities.forEach(function(a){
        var dk = deptKeyOf(a.dept); // S.7: canonical short key
        Object.keys(STAGE_DEPT_KEY).forEach(function(s){ if (STAGE_DEPT_KEY[s] === dk) activeDepts[s] = true; });
      });
    }
    for (var si = idx-1; si >= 0; si--) { if (activeDepts[STAGE_ORDER[si]]) { predMovements = STAGE_OWN_MOVEMENTS[STAGE_ORDER[si]]; break; } }
  }

  // This-stage issued per size + predecessor received per size (and totals).
  var issued = {}, predRecv = {}, predRecvSizeSum = 0, predReceivedTotal = 0;
  var jcs = getJobCards({ orderRef: ref }, ss);
  if (Array.isArray(jcs)) {
    jcs.forEach(function(jc) {
      if (safeStr(jc.status).toUpperCase() === 'CANCELLED') return;
      if (mv && jc.movement === mv) {
        var sb = jc.sizeBreakdown || {};
        Object.keys(sb).forEach(function(k){ issued[k] = safeNum(issued[k]) + safeNum(sb[k]); });
      }
      if (predMovements && predMovements.indexOf(jc.movement) >= 0) {
        predReceivedTotal += safeNum(jc.pairsReceived);
        var rb = jc.receivedBreakdown || {};
        Object.keys(rb).forEach(function(k){ var q = safeNum(rb[k]); predRecv[k] = safeNum(predRecv[k]) + q; predRecvSizeSum += q; });
      }
    });
  }

  // Use predecessor-delivered sizes as the per-size cap ONLY when the predecessor's
  // per-size receipts fully account for its received total. Otherwise (first stage,
  // or a predecessor received before per-size tracking existed) fall back to the
  // order size run so in-flight/legacy orders are never blocked.
  var usePred = !!predMovements && predReceivedTotal > 0 && predRecvSizeSum === predReceivedTotal;

  var sizes = {};
  Object.keys(ordered).forEach(function(k) {
    var o = safeNum(ordered[k]), iss = safeNum(issued[k]);
    var cap = usePred ? safeNum(predRecv[k]) : o;
    sizes[k] = {
      ordered:   o,
      issued:    iss,
      delivered: usePred ? safeNum(predRecv[k]) : o,
      remaining: Math.max(0, Math.min(o, cap) - iss)
    };
  });
  var role = '';
  try { role = getUserInfo().role; } catch(e) {}
  return { success:true, sizes:sizes, isAdmin: role === 'admin', boundedByPredecessor: usePred };
}

// ── Admin override gate (Phase 7.1) ──────────────────────────────────────────
// Two-step verification for admin overrides. The password lives in Script
// Properties (key ADMIN_OVERRIDE_PW), never in code. Returns {success} so callers
// can gate any sensitive override (size-run, order edits, etc.).
function verifyAdminOverride(password) {
  var u = getUserInfo();
  if (u.role !== 'admin') return { success:false, error:'Not authorised' };
  var pw = '';
  try { pw = PropertiesService.getScriptProperties().getProperty('ADMIN_OVERRIDE_PW') || ''; } catch(e) {}
  if (!pw) return { success:false, error:'Override password not set. Add ADMIN_OVERRIDE_PW in Script Properties.' };
  if (safeStr(password) !== pw) return { success:false, error:'Incorrect override password' };
  return { success:true };
}

// Admin override (7.1): edit an order's details directly, gated by the override
// password. Updates lot size (ART H2), customer (ART E2 + trackers), and/or the
// size run (ORDER_INDEX col Q JSON). Only the fields provided are changed.
// S.6: locked — writer
function adminEditOrder(data) {
  var ov = verifyAdminOverride(data.overridePassword);
  if (!ov.success) return { success:false, error: ov.error || 'Override not verified' };
  var sheetName = safeStr(data.sheetName).trim();
  if (!sheetName) return { success:false, error:'sheetName is required' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var ss  = SpreadsheetApp.openById(SHEET_ID);
    var art = ss.getSheetByName(sheetName);
    if (!art) return { success:false, error:'Order sheet not found: ' + sheetName };

    var hasLot  = (data.lot !== undefined && data.lot !== null && safeStr(data.lot) !== '');
    var hasCust = (data.customer !== undefined && safeStr(data.customer).trim() !== '');
    var hasSize = (data.sizeRun && typeof data.sizeRun === 'object' && Object.keys(data.sizeRun).length > 0);
    var lotN = hasLot ? safeNum(data.lot) : null;
    var cust = hasCust ? safeStr(data.customer).trim() : null;
    var cleanSize = null;
    if (hasSize) { cleanSize = {}; Object.keys(data.sizeRun).forEach(function(k){ var q = safeNum(data.sizeRun[k]); if (q > 0) cleanSize[k] = q; }); }
    var hasRates = Array.isArray(data.rates) && data.rates.length > 0;
    if (!hasLot && !hasCust && !hasSize && !hasRates) return { success:false, error:'Nothing to change' };

    var changed = [];
    if (hasLot)  { art.getRange('H2').setValue(lotN); changed.push('lot=' + lotN); }
    if (hasCust) { art.getRange('E2').setValue(cust); changed.push('customer'); }

    var oi = ss.getSheetByName('ORDER_INDEX');
    if (oi && oi.getLastRow() >= 4) {
      var oiV = oi.getRange(4, 1, oi.getLastRow()-3, 17).getValues();
      for (var i = 0; i < oiV.length; i++) {
        if (safeStr(oiV[i][1]).trim() === sheetName) {
          var row = 4 + i;
          if (hasLot)  oi.getRange(row, 9).setValue(lotN);      // col I = lot
          if (hasCust) oi.getRange(row, 5).setValue(cust);      // col E = customer
          if (hasSize) { oi.getRange(row, 17).setValue(JSON.stringify(cleanSize)); changed.push('sizeRun'); }
          break;
        }
      }
    }
    if (hasCust) {
      var ot = ss.getSheetByName('ORDER_TRACKER');
      if (ot && ot.getLastRow() > 3) {
        var otV = ot.getRange(4, 1, ot.getLastRow()-3, 1).getValues();
        for (var j = 0; j < otV.length; j++) {
          if (safeStr(otV[j][0]).trim() === sheetName) { ot.getRange(4 + j, 3).setValue(cust); break; }
        }
      }
    }

    // Rates: update matching activities inside this order's APPROVED ACTIVITY_SETUP
    // payloads (the source payment reads). Matched by dept + activity name.
    if (hasRates) {
      var newByKey = {};
      data.rates.forEach(function(x) {
        var k = safeStr(x.dept).toLowerCase().trim() + '|' + safeStr(x.activityName).toLowerCase().trim();
        newByKey[k] = { rate: safeNum(x.rate), comm: safeNum(x.comm) };
      });
      var rq = ss.getSheetByName('REQUESTS');
      if (rq && rq.getLastRow() >= 4) {
        var rqV = rq.getRange(4, 1, rq.getLastRow()-3, 6).getValues();
        for (var k2 = 0; k2 < rqV.length; k2++) {
          if (safeStr(rqV[k2][3]) !== 'ACTIVITY_SETUP' || safeStr(rqV[k2][5]).toUpperCase() !== 'APPROVED') continue;
          var pl; try { pl = JSON.parse(safeStr(rqV[k2][4])); } catch(e) { continue; }
          if (!pl || safeStr(pl.sheet) !== sheetName) continue;
          var deptKey = safeStr(pl.dept).toLowerCase().trim();
          var dirty = false;
          if (Array.isArray(pl.activities)) {
            pl.activities.forEach(function(a) {
              var kk = deptKey + '|' + safeStr(a.activityName).toLowerCase().trim();
              if (newByKey[kk]) { a.rate = newByKey[kk].rate; a.comm = newByKey[kk].comm; dirty = true; }
            });
          } else if (pl.activityName) {
            var kk1 = deptKey + '|' + safeStr(pl.activityName).toLowerCase().trim();
            if (newByKey[kk1]) { pl.rate = newByKey[kk1].rate; pl.comm = newByKey[kk1].comm; dirty = true; }
          }
          if (dirty) rq.getRange(4 + k2, 5).setValue(JSON.stringify(pl));
        }
        changed.push('rates');
      }
    }
    SpreadsheetApp.flush();
    ['dashboardData_', 'storeScreenData_', 'entryData_'].forEach(function(k){ try { CacheService.getScriptCache().remove(k + CONFIG.ENV); } catch(e) {} });
    return { success:true, changed:changed };
    } catch(e) { return { success:false, error:e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// S.6: locked — migration
function backfillOrderSizes(targetEnv) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var _sid = (targetEnv === 'LIVE') ? CONFIG.LIVE_SHEET_ID
             : (targetEnv === 'DEV')  ? CONFIG.DEV_SHEET_ID
             : SHEET_ID;
    var ss = SpreadsheetApp.openById(_sid);
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
    return { success: true, updated: updated, skipped: skipped, env: (targetEnv || CONFIG.ENV) };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function _bfsReport(result, title) {
  var msg = result.success
    ? ('Target: ' + result.env + '\nUpdated: ' + result.updated + ' rows | Skipped: ' + result.skipped + ' rows')
    : ('Error: ' + result.error);
  SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK);
}
function backfillOrderSizesLive() { _bfsReport(backfillOrderSizes('LIVE'), 'Backfill Order Sizes — LIVE'); }
function backfillOrderSizesDev()  { _bfsReport(backfillOrderSizes('DEV'),  'Backfill Order Sizes — DEV'); }
// Kept for backward compatibility (targets current ENV).
function backfillOrderSizesMenu() { _bfsReport(backfillOrderSizes(), 'Backfill Order Sizes'); }
