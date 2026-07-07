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
