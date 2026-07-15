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
      'ISSUED_BY','ISSUED_AT','EXPECTED_RETURN','RECEIVED_AT','STATUS','NOTES',
      'BATCH_ID'
    ]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function issueJobCard(data) {
  var _user = getUserInfo();
  if (_user.role !== 'store' && _user.role !== 'admin') return { success:false, error:'Not authorised' };
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

  // Predecessor stage lock — finds nearest EARLIER active stage
  // for this specific order, skipping stages that don't apply
  var STAGE_ORDER = ['Cutting','Preparation','Fitter','Lasting','Packing','Dispatch'];
  var STAGE_DEPT_KEY = {
    'Cutting':'cutting','Preparation':'prep','Fitter':'fitter',
    'Lasting':'lasting','Packing':'finishing','Dispatch':'dispatch'
  };
  var STAGE_OWN_MOVEMENTS = {
    'Cutting':['Cutting IN'],
    'Preparation':['Preparation IN'],
    'Fitter':['Fitter IN'],
    'Lasting':['Upper IN','Lasting IN'],
    'Packing':['Packing IN'],
    'Dispatch':['Dispatch IN']
  };
  var MOVEMENT_TO_STAGE = {
    'Cutting IN':'Cutting','Preparation IN':'Preparation',
    'Fitter IN':'Fitter','Upper IN':'Lasting','Lasting IN':'Lasting',
    'Packing IN':'Packing','Dispatch IN':'Dispatch'
  };

  var currentStage = MOVEMENT_TO_STAGE[movement] || '';
  if (currentStage) {
    var currentStageIdx = STAGE_ORDER.indexOf(currentStage);

    // Get which stages are active for THIS order (have approved activities)
    var orderActRes = getApprovedActivitiesForArticle(orderRef);
    var orderActiveDepts = {};
    if (orderActRes && orderActRes.success && Array.isArray(orderActRes.activities)) {
      orderActRes.activities.forEach(function(a) {
        var dk = safeStr(a.dept).toLowerCase();
        Object.keys(STAGE_DEPT_KEY).forEach(function(stageName) {
          if (STAGE_DEPT_KEY[stageName] === dk) orderActiveDepts[stageName] = true;
        });
      });
    }

    // Walk backward from current stage to find nearest active predecessor
    var predecessorStage = null;
    for (var si = currentStageIdx - 1; si >= 0; si--) {
      var candidateStage = STAGE_ORDER[si];
      if (orderActiveDepts[candidateStage]) {
        predecessorStage = candidateStage;
        break;
      }
    }

    if (predecessorStage) {
      var predMovements = STAGE_OWN_MOVEMENTS[predecessorStage] || [];
      var allJCsForLock = getJobCards({orderRef: orderRef});
      if (!Array.isArray(allJCsForLock)) allJCsForLock = [];

      var predJCs = allJCsForLock.filter(function(jc) {
        return predMovements.indexOf(jc.movement) >= 0;
      });

      if (predJCs.length === 0) {
        return {
          success: false,
          error: predecessorStage + ' has no job cards yet for this order. ' +
                 'Issue and complete ' + predecessorStage + ' first.'
        };
      }

      var predIncomplete = predJCs.filter(function(jc) {
        var st = safeStr(jc.status).toUpperCase();
        return st === 'ISSUED' || st === 'PARTIAL';
      });
      if (predIncomplete.length > 0) {
        return {
          success: false,
          error: predecessorStage + ' still has ' + predIncomplete.length +
                 ' open job card(s). Complete ' + predecessorStage +
                 ' before issuing ' + currentStage + '.'
        };
      }

      // Pairs cap: cannot issue more than predecessor received
      var predReceived = 0;
      allJCsForLock.forEach(function(jc) {
        if (predMovements.indexOf(jc.movement) >= 0) {
          var st = safeStr(jc.status).toUpperCase();
          if (st === 'COMPLETE' || st === 'PAYMENT_PENDING' || st === 'PAID') {
            predReceived += safeNum(jc.pairsReceived);
          }
        }
      });

      var thisStageMovements = STAGE_OWN_MOVEMENTS[currentStage] || [];
      var thisStageAlreadyIssued = 0;
      allJCsForLock.forEach(function(jc) {
        if (thisStageMovements.indexOf(jc.movement) >= 0) {
          var st = safeStr(jc.status).toUpperCase();
          if (st !== 'CANCELLED') thisStageAlreadyIssued += safeNum(jc.pairsIssued);
        }
      });

      var availableForThisStage = predReceived - thisStageAlreadyIssued;
      if (pairsIssued > availableForThisStage) {
        return {
          success: false,
          error: 'Cannot issue ' + pairsIssued + ' pairs. ' + predecessorStage +
                 ' completed ' + predReceived + ' pairs, and ' +
                 thisStageAlreadyIssued + ' already issued for ' + currentStage +
                 '. Maximum available: ' + availableForThisStage + ' pairs.'
        };
      }
    }
    // If predecessorStage is null, this is the first active stage for
    // this order — no lock, draws against order lot size (existing
    // validation elsewhere already caps against order balance)
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

function issueJobCardBatch(data) {
  var _user = getUserInfo();
  if (_user.role !== 'store' && _user.role !== 'admin')
    return { success:false, error:'Not authorised' };

  var orderRef       = safeStr(data.orderRef       || '').trim();
  var store          = safeStr(data.store          || '').trim();
  var movement       = safeStr(data.movement       || '').trim();
  var items          = Array.isArray(data.items) ? data.items : [];
  var expectedReturn = safeStr(data.expectedReturn || '').trim();
  var notes          = safeStr(data.notes          || '').trim();

  if (!orderRef)      return { success:false, error:'orderRef is required' };
  if (!items.length)  return { success:false, error:'At least one activity-contractor row is required' };
  if (!expectedReturn) return { success:false, error:'expectedReturn is required' };

  var batchId = 'BATCH-' + new Date().getFullYear() + '-' +
                Utilities.getUuid().slice(0,8);

  var results = [];
  var anyFailed = false;

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var singleResult = issueJobCard({
      orderRef:       orderRef,
      workOrder:      data.workOrder,
      store:          store,
      movement:       movement,
      contractorId:   item.contractorId,
      pairsIssued:    item.pairsIssued,
      sizeBreakdown:  data.sizeBreakdown,
      expectedReturn: expectedReturn,
      notes:          notes,
      activityName:   item.activityName
    });
    if (singleResult.success) {
      try {
        var ws = ensureJobCardsSheet();
        var lastRow = ws.getLastRow();
        var idCol = ws.getRange(2, 1, lastRow-1, 1).getValues();
        for (var r = 0; r < idCol.length; r++) {
          if (safeStr(idCol[r][0]).trim() === singleResult.jobCardId) {
            ws.getRange(r+2, 16).setValue(batchId); // column P
            break;
          }
        }
      } catch(tagErr) {}
      results.push({ activityName: item.activityName, jobCardId: singleResult.jobCardId, success: true });
    } else {
      anyFailed = true;
      results.push({ activityName: item.activityName, success: false, error: singleResult.error });
    }
  }

  return {
    success: !anyFailed,
    batchId: batchId,
    results: results,
    partialSuccess: results.some(function(r){ return r.success; }) && anyFailed
  };
}

function receiveJobCard(data) {
  var _user = getUserInfo();
  if (_user.role !== 'store' && _user.role !== 'admin') return { success:false, error:'Not authorised' };
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
