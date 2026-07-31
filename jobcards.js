// ── JOB CARDS ─────────────────────────────────────────────────────────────────

// S.6: locked — bootstrap (double-checked: fast path when sheet + headers exist)
function ensureJobCardsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('JOB_CARDS');
  if (ws &&
      safeStr(ws.getRange(1, 17).getValue()) !== '' &&
      safeStr(ws.getRange(1, 18).getValue()) !== '') return ws;             // fast path — no lock
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('JOB_CARDS');                                    // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('JOB_CARDS');
      // STATUS values: ISSUED | PARTIAL | COMPLETE | PAYMENT_PENDING | PAID | CANCELLED
      ws.appendRow([
        'JOB_CARD_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT',
        'CONTRACTOR_ID','PAIRS_ISSUED','PAIRS_RECEIVED','SIZE_BREAKDOWN',
        'ISSUED_BY','ISSUED_AT','EXPECTED_RETURN','RECEIVED_AT','STATUS','NOTES',
        'BATCH_ID','ASSIGNMENTS','RECEIVED_BREAKDOWN'
      ]);
      ws.setFrozenRows(1);
    } else {
      // Label added columns once for sheets that predate them.
      if (safeStr(ws.getRange(1, 17).getValue()) === '') ws.getRange(1, 17).setValue('ASSIGNMENTS');
      if (safeStr(ws.getRange(1, 18).getValue()) === '') ws.getRange(1, 18).setValue('RECEIVED_BREAKDOWN');
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

function issueJobCard(data) {
  var _user = getUserInfo();
  if (_user.role !== 'store' && _user.role !== 'admin') return { success:false, error:'Not authorised' };
  var STORE_MOVEMENT_MAP = {
    'Upper Store':             ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store': ['Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
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
  // S.9: contractor identity must be a real CTR-ID — a name or typo here is how
  // payments end up unattributable.
  if (!/^CTR-\d+$/.test(contractorId) || !_validContractorIds_()[contractorId])     return { success: false, error: 'Unknown contractor id "' + contractorId + '" — reload the app and pick the contractor again.' };
  if (!pairsIssued || pairsIssued <= 0 || Math.floor(pairsIssued) !== pairsIssued) return { success: false, error: 'pairsIssued must be a positive integer' };
  if (!expectedReturn)                                                              return { success: false, error: 'expectedReturn is required' };

  // Size breakdown, when provided, must total the pairs issued. Prevents the
  // per-size balance corruption seen on trial cards (breakdown 60 vs 30 issued).
  var _sbSum = 0; Object.keys(sizeBreakdown || {}).forEach(function(_k){ _sbSum += safeNum(sizeBreakdown[_k]); });
  if (_sbSum > 0 && _sbSum !== pairsIssued) return { success: false, error: 'Size breakdown totals ' + _sbSum + ' but pairs issued is ' + pairsIssued + ' — they must match.' };

  // S.6: lock acquired BEFORE the cap/validation reads below — with the reads
  // outside the lock, two concurrent issues could both pass the cap.
  var jobCardId;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

  // Check approved activities exist for this order + department
  var deptKey = {
    'Cutting IN':     'cutting',
    'Preparation IN': 'prep',
    'Fitter IN':      'fitter',
    'Lasting IN':     'lasting',
    'Packing IN':     'finish',
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
    'Lasting':'lasting','Packing':'finish','Dispatch':'dispatch'
  };
  var STAGE_OWN_MOVEMENTS = {
    'Cutting':['Cutting IN'],
    'Preparation':['Preparation IN'],
    'Fitter':['Fitter IN'],
    'Lasting':['Lasting IN'],
    'Packing':['Packing IN'],
    'Dispatch':['Dispatch IN']
  };
  var MOVEMENT_TO_STAGE = {
    'Cutting IN':'Cutting','Preparation IN':'Preparation',
    'Fitter IN':'Fitter','Lasting IN':'Lasting',
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
        var dk = deptKeyOf(a.dept); // S.7: canonical short key
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

      // Pairs cap: cannot issue more than predecessor received
      var predReceived = 0;
      allJCsForLock.forEach(function(jc) {
        if (predMovements.indexOf(jc.movement) >= 0) {
          var st = safeStr(jc.status).toUpperCase();
          // Include PARTIAL: partially-received predecessor cards hold real pairs
          // available to this stage (bug 8.B1 — must match getMaxIssuableForStage).
          if (st === 'PARTIAL' || st === 'COMPLETE' || st === 'PAYMENT_PENDING' || st === 'PAID') {
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
    if (!predecessorStage) {
      // First active stage — cap against order lot size
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var oi = ss.getSheetByName('ORDER_INDEX');
      var orderLotSize = 0;
      if (oi && oi.getLastRow() > 3) {
        var oiRows = oi.getRange(4, 1, oi.getLastRow()-3, 9).getValues();
        for (var oiR = 0; oiR < oiRows.length; oiR++) {
          if (safeStr(oiRows[oiR][1]).trim() === orderRef) {
            orderLotSize = safeNum(oiRows[oiR][8]);
            break;
          }
        }
      }
      if (orderLotSize > 0) {
        var firstStageMovements = STAGE_OWN_MOVEMENTS[currentStage] || [];
        var firstStageJCs = getJobCards({orderRef: orderRef});
        if (!Array.isArray(firstStageJCs)) firstStageJCs = [];
        var firstStageAlreadyIssued = 0;
        firstStageJCs.forEach(function(jc) {
          if (firstStageMovements.indexOf(jc.movement) >= 0) {
            var st = safeStr(jc.status).toUpperCase();
            if (st !== 'CANCELLED') firstStageAlreadyIssued += safeNum(jc.pairsIssued);
          }
        });
        var availableFirstStage = orderLotSize - firstStageAlreadyIssued;
        if (pairsIssued > availableFirstStage) {
          return {
            success: false,
            error: 'Cannot issue ' + pairsIssued + ' pairs. Order lot size is ' +
                   orderLotSize + ', and ' + firstStageAlreadyIssued +
                   ' already issued for ' + currentStage +
                   '. Maximum available: ' + availableFirstStage + ' pairs.'
          };
        }
      }
    }
  }

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

  try { CacheService.getScriptCache().remove('storeScreenData_' + CONFIG.ENV); } catch(ce) {}
  try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
  var issueResult = { success: true, jobCardId: jobCardId };
  if (wipWarning) issueResult.warning = 'WIP entry not created: ' + wipWarning;
  return issueResult;
}

// One-card-per-department issue: a single job card carries `pairs` for the whole
// department (one movement, one WIP entry, one stage-cap consumption) plus a list
// of {activity, contractorId} assignments for per-contractor payment.
function issueDepartmentJobCard(data) {
  var _user = getUserInfo();
  if (_user.role !== 'store' && _user.role !== 'admin') return { success:false, error:'Not authorised' };

  var STORE_MOVEMENT_MAP = {
    'Upper Store':             ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store': ['Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':          ['Dispatch IN','Dispatch OUT']
  };
  var DEPT_KEY = {
    'Cutting IN':'cutting','Preparation IN':'prep','Fitter IN':'fitter',
    'Lasting IN':'lasting','Packing IN':'finish','Dispatch IN':'dispatch'
  };

  var orderRef       = safeStr(data.orderRef       || '').trim();
  var workOrder      = safeStr(data.workOrder      || '').trim();
  var store          = safeStr(data.store          || '').trim();
  var movement       = safeStr(data.movement       || '').trim();
  var pairs          = safeNum(data.pairs);
  var assignments    = Array.isArray(data.assignments) ? data.assignments : [];
  var sizeBreakdown  = data.sizeBreakdown || {};
  var expectedReturn = safeStr(data.expectedReturn || '').trim();
  var notes          = safeStr(data.notes          || '').trim();

  if (!orderRef)                                                       return { success:false, error:'orderRef is required' };
  if (!STORE_MOVEMENT_MAP[store])                                      return { success:false, error:'Invalid store: ' + store };
  if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0)                return { success:false, error:'Invalid movement for store: ' + movement };
  if (!assignments.length)                                            return { success:false, error:'At least one activity-contractor assignment is required' };
  if (!pairs || pairs <= 0 || Math.floor(pairs) !== pairs)            return { success:false, error:'pairs must be a positive integer' };
  if (!expectedReturn)                                                return { success:false, error:'expectedReturn is required' };

  // Size breakdown, when provided, must total the pairs issued (data-integrity guard).
  var _sbSum = 0; Object.keys(sizeBreakdown || {}).forEach(function(_k){ _sbSum += safeNum(sizeBreakdown[_k]); });
  if (_sbSum > 0 && _sbSum !== pairs) return { success:false, error:'Size breakdown totals ' + _sbSum + ' but pairs is ' + pairs + ' — they must match.' };

  // S.6: lock acquired BEFORE the rate/cap/size-balance reads below — with the
  // reads outside the lock, two concurrent issues could both pass the cap.
  var jobCardId;
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

  // Resolve approved activity rates for this order + department
  var deptKey = DEPT_KEY[movement] || '';
  var approvedByName = {};
  try {
    var ar = getApprovedActivitiesForArticle(orderRef);
    if (ar && ar.success && Array.isArray(ar.activities)) {
      ar.activities.forEach(function(a) {
        if (!deptKey || safeStr(a.dept).toLowerCase().indexOf(deptKey) === 0)
          approvedByName[safeStr(a.activityName)] = { rate:safeNum(a.rate), comm:safeNum(a.comm), dept:safeStr(a.dept) };
      });
    }
  } catch(e) {}
  if (!Object.keys(approvedByName).length)
    return { success:false, error:'No approved activities for this department on order ' + orderRef + '. Ask Arvind to set up and approve activities first.' };

  // Normalise assignments; every activity must be approved and have a contractor
  var normAssign = [];
  var _ctrIdSet = _validContractorIds_();   // S.9: every assignment must carry a real CTR-ID
  for (var i = 0; i < assignments.length; i++) {
    var an  = safeStr(assignments[i].activityName || assignments[i].activity).trim();
    var cid = safeStr(assignments[i].contractorId).trim();
    if (!an || !cid) continue;
    if (!/^CTR-\d+$/.test(cid) || !_ctrIdSet[cid]) return { success:false, error:'Unknown contractor id "' + cid + '" on activity ' + an + ' — reload the app and pick the contractor again.' };
    var meta = approvedByName[an];
    if (!meta) return { success:false, error:'Activity not approved for this department: ' + an };
    normAssign.push({ activity:an, contractorId:cid, rate:meta.rate, comm:meta.comm });
  }
  if (!normAssign.length) return { success:false, error:'Assign a contractor to at least one activity' };

  // Stage cap — pairs move once for the whole department. Reuse the tested calc.
  try {
    var maxRes = getMaxIssuableForStage(orderRef, movement);
    if (maxRes && maxRes.success && pairs > safeNum(maxRes.maxIssuable)) {
      return { success:false, error:'Cannot issue ' + pairs + ' pairs. Maximum available for this stage: ' +
               safeNum(maxRes.maxIssuable) + ' pairs (' + safeStr(maxRes.source) + ').' };
    }
  } catch(e) {}

  // Per-size cap — can't issue more of a size than remains in the approved size run.
  // Admin (Ayush) may override, but only with the two-step override password (7.1).
  try {
    var _bal = getOrderSizeBalance(orderRef, movement);
    if (_bal && _bal.success && _bal.sizes) {
      var _over = [];
      Object.keys(sizeBreakdown || {}).forEach(function(k) {
        var want = safeNum(sizeBreakdown[k]);
        var rem  = _bal.sizes[k] ? safeNum(_bal.sizes[k].remaining) : 0;
        if (want > rem) _over.push(k + ': ' + want + ' > ' + rem + ' left');
      });
      if (_over.length) {
        if (_user.role !== 'admin')
          return { success:false, error:'Over the approved size run — ' + _over.join(', ') + '. Ask Ayush to override.' };
        var _ov = verifyAdminOverride(data.overridePassword);
        if (!_ov || !_ov.success)
          return { success:false, error:'Admin override of the size run needs the override password. ' + ((_ov && _ov.error) || '') };
      }
    }
  } catch(e) {}

    var ws       = ensureJobCardsSheet();
    var dataRows = Math.max(0, ws.getLastRow() - 1);
    var nextNum  = dataRows + 1;
    var year     = new Date().getFullYear();
    var seq      = String(nextNum); while (seq.length < 3) seq = '0' + seq;
    jobCardId    = 'JC-' + year + '-' + seq;
    var issuedBy = Session.getActiveUser().getEmail();
    var now      = new Date().toISOString();
    var primary  = normAssign[0].contractorId;  // legacy CONTRACTOR_ID col / display fallback
    ws.appendRow([
      jobCardId, orderRef, workOrder, store, movement, primary,
      pairs, 0, JSON.stringify(sizeBreakdown), issuedBy,
      now, expectedReturn, '', 'ISSUED', notes,
      '', JSON.stringify(normAssign)
    ]);
    SpreadsheetApp.flush();
  } catch(e) {
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
  }

  // Resolve current open periodId; fall back to synthetic JC date
  var periodId = 'JC-' + new Date().toISOString().slice(0, 10);
  try {
    var ss2 = SpreadsheetApp.openById(SHEET_ID);
    var pp  = ss2.getSheetByName('PAYMENT_PERIODS');
    if (pp && pp.getLastRow() > 1) {
      var ppV = pp.getRange(2, 1, pp.getLastRow() - 1, 7).getValues();
      var oids = [];
      ppV.forEach(function(r){ if (safeStr(r[6]).trim().toUpperCase() === 'OPEN') oids.push(safeStr(r[0])); });
      oids.sort();
      if (oids.length) periodId = oids[0];
    }
  } catch(pe) {}

  // One IN-side WIP entry for the whole department movement
  var wipWarning;
  try {
    var wipResult = saveWipEntry({
      orderRef:    orderRef,
      workOrder:   workOrder,
      store:       store,
      movement:    movement,
      pairs:       pairs,
      periodId:    periodId,
      notes:       'Job Card ' + jobCardId,
      contractors: normAssign.map(function(a){ return a.contractorId; }),
      jobCardRef:  jobCardId
    });
    if (wipResult && wipResult.success === false) wipWarning = wipResult.error;
    else try { generateDailyReport(); } catch(e) {}
  } catch(wipErr) { wipWarning = wipErr.message; }

  try { CacheService.getScriptCache().remove('storeScreenData_' + CONFIG.ENV); } catch(ce) {}
  try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
  var res = { success:true, jobCardId:jobCardId };
  if (wipWarning) res.warning = 'WIP entry not created: ' + wipWarning;
  return res;
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
      // S.6: locked — batch-tag
      try {
        var tagLock = LockService.getScriptLock();
        try {
          tagLock.waitLock(10000);
          var ws = ensureJobCardsSheet();
          var lastRow = ws.getLastRow();
          var idCol = ws.getRange(2, 1, lastRow-1, 1).getValues();
          for (var r = 0; r < idCol.length; r++) {
            if (safeStr(idCol[r][0]).trim() === singleResult.jobCardId) {
              ws.getRange(r+2, 16).setValue(batchId); // column P
              break;
            }
          }
        } finally {
          tagLock.releaseLock();
        }
      } catch(tagErr) {}
      results.push({ activityName: item.activityName, jobCardId: singleResult.jobCardId, success: true });
    } else {
      anyFailed = true;
      results.push({ activityName: item.activityName, success: false, error: singleResult.error });
    }
  }

  try { CacheService.getScriptCache().remove('storeScreenData_' + CONFIG.ENV); } catch(ce) {}
  try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
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
  var receivedSizeBreakdown = (data.receivedSizeBreakdown && typeof data.receivedSizeBreakdown === 'object') ? data.receivedSizeBreakdown : null;

  if (!jobCardId)                                                                           return { success: false, error: 'jobCardId is required' };
  if (!pairsReceived || pairsReceived <= 0 || Math.floor(pairsReceived) !== pairsReceived) return { success: false, error: 'pairsReceived must be a positive integer' };

  var orderRef, workOrder, store, inMovement, contractorId;
  var pairsIssued, effectivePairs, newReceived, newStatus;

  var lock = LockService.getScriptLock();
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

    // Per-size received breakdown (from the receive form). Must total the pairs
    // credited this event; cumulative per-size receipts can't exceed what was
    // issued. Accumulated into RECEIVED_BREAKDOWN (col 18) across partial receives.
    // 8.B2: a sized card (non-empty SIZE_BREAKDOWN) must carry a breakdown on
    // EVERY receive — silently accepting breakdown-less receives is what let
    // PAIRS_RECEIVED drift from RECEIVED_BREAKDOWN and permanently blocked the
    // final receive on affected cards.
    var _issuedSb = {}; try { _issuedSb = JSON.parse(safeStr(row[8])) || {}; } catch(e) {}
    var _issuedSbHasSizes = Object.keys(_issuedSb).some(function(k){ return safeNum(_issuedSb[k]) > 0; });
    if (_issuedSbHasSizes && !receivedSizeBreakdown) {
      return { success:false, error:'This job card tracks sizes — the receive must include a per-size breakdown totalling ' +
               effectivePairs + ' pairs. If you don\'t see the size panel, reload the app (an old version may be open).' };
    }
    var mergedRcvSb = null;
    if (receivedSizeBreakdown) {
      var _existSb  = {}; try { _existSb  = JSON.parse(safeStr(ws.getRange(sheetRow, 18).getValue())) || {}; } catch(e) {}
      var _rsum = 0; Object.keys(receivedSizeBreakdown).forEach(function(k){ _rsum += safeNum(receivedSizeBreakdown[k]); });
      if (_rsum !== effectivePairs)
        return { success:false, error:'Received size breakdown totals ' + _rsum + ' but pairs received is ' + effectivePairs + ' — they must match.' };
      mergedRcvSb = {};
      Object.keys(_existSb).forEach(function(k){ mergedRcvSb[k] = safeNum(_existSb[k]); });
      var _bad = '';
      Object.keys(receivedSizeBreakdown).forEach(function(k){
        mergedRcvSb[k] = safeNum(mergedRcvSb[k]) + safeNum(receivedSizeBreakdown[k]);
        if (Object.keys(_issuedSb).length && mergedRcvSb[k] > safeNum(_issuedSb[k])) _bad = k;
      });
      if (_bad) return { success:false, error:'Received more of size ' + _bad + ' than was issued.' };
    }

    newReceived = currentRecvd + effectivePairs;
    newStatus   = newReceived >= pairsIssued ? 'COMPLETE' : 'PARTIAL';
    var now     = new Date().toISOString();

    ws.getRange(sheetRow, 8).setValue(newReceived);
    if (!safeStr(row[12]).trim()) ws.getRange(sheetRow, 13).setValue(now);
    ws.getRange(sheetRow, 14).setValue(newStatus);
    if (finalNotes) {
      ws.getRange(sheetRow, 15).setValue(existingNotes ? existingNotes + '; ' + finalNotes : finalNotes);
    }
    if (mergedRcvSb) ws.getRange(sheetRow, 18).setValue(JSON.stringify(mergedRcvSb));
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
  // F.5: 'Upper IN' retired (never used — zero rows on both sheets); every IN
  // movement pairs with an OUT movement of the same prefix.
  {
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

  try { CacheService.getScriptCache().remove('storeScreenData_' + CONFIG.ENV); } catch(ce) {}
  try { CacheService.getScriptCache().remove('dashboardData_' + CONFIG.ENV); } catch(ce) {}
  var rcvResult = { success: true, jobCardId: jobCardId, totalReceived: newReceived, pairsIssued: pairsIssued, status: newStatus };
  if (rcvWarning) rcvResult.warning = 'WIP entry not created: ' + rcvWarning;
  return rcvResult;
}

function getJobCards(filters, ss) {
  try {
    var ws      = ensureJobCardsSheet(ss);
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows   = ws.getRange(2, 1, lastRow - 1, 18).getValues();
    var result = [];
    rows.forEach(function(r) {
      if (!safeStr(r[0]).trim()) return;
      var sd = {};
      try { sd = JSON.parse(safeStr(r[8])) || {}; } catch(e) {}
      var asg = [];
      try { asg = JSON.parse(safeStr(r[16])) || []; } catch(e) {}
      var rsd = {};
      try { rsd = JSON.parse(safeStr(r[17])) || {}; } catch(e) {}
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
        notes:          safeStr(r[14]),
        batchId:        safeStr(r[15]),
        assignments:    Array.isArray(asg) ? asg : [],
        receivedBreakdown: (rsd && typeof rsd === 'object') ? rsd : {}
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

// Admin override (7.1 #3 + #4): correct a job card directly, gated by the override
// password. Reassign contractors (assignments), fix pairs issued/received, or set
// status. Only the fields provided change.
function adminEditJobCard(data) {
  var ov = verifyAdminOverride(data.overridePassword);
  if (!ov.success) return { success:false, error: ov.error || 'Override not verified' };
  var jobCardId = safeStr(data.jobCardId).trim();
  if (!jobCardId) return { success:false, error:'jobCardId is required' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ws = ensureJobCardsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return { success:false, error:'Job card not found' };
    var rows = ws.getRange(2, 1, lastRow-1, 17).getValues();
    var idx = -1;
    for (var i = 0; i < rows.length; i++) { if (safeStr(rows[i][0]).trim() === jobCardId) { idx = i; break; } }
    if (idx < 0) return { success:false, error:'Job card not found: ' + jobCardId };
    var sheetRow = idx + 2;
    var changed = [];

    // Reassign contractors (department card): full assignments array replaces col 17
    var _ctrIdSetAE = _validContractorIds_();   // S.9: overrides can't introduce unknown contractor ids either
    if (Array.isArray(data.assignments) && data.assignments.length) {
      var norm = data.assignments.map(function(a) {
        return { activity: safeStr(a.activity || a.activityName), contractorId: safeStr(a.contractorId).trim(),
                 rate: safeNum(a.rate), comm: safeNum(a.comm) };
      }).filter(function(a){ return a.contractorId; });
      for (var _vi = 0; _vi < norm.length; _vi++) {
        if (!/^CTR-\d+$/.test(norm[_vi].contractorId) || !_ctrIdSetAE[norm[_vi].contractorId])
          return { success:false, error:'Unknown contractor id "' + norm[_vi].contractorId + '" — reload and pick the contractor again.' };
      }
      if (norm.length) {
        ws.getRange(sheetRow, 17).setValue(JSON.stringify(norm));   // ASSIGNMENTS
        ws.getRange(sheetRow, 6).setValue(norm[0].contractorId);    // CONTRACTOR_ID (display fallback)
        changed.push('contractors');
      }
    } else if (data.contractorId !== undefined && safeStr(data.contractorId).trim()) {
      var _aeCid = safeStr(data.contractorId).trim();
      if (!/^CTR-\d+$/.test(_aeCid) || !_ctrIdSetAE[_aeCid])
        return { success:false, error:'Unknown contractor id "' + _aeCid + '" — reload and pick the contractor again.' };
      ws.getRange(sheetRow, 6).setValue(_aeCid);
      changed.push('contractor');
    }

    // Pairs corrections
    if (data.pairsIssued !== undefined && safeStr(data.pairsIssued) !== '') {
      var pi = safeNum(data.pairsIssued);
      if (pi >= 0) { ws.getRange(sheetRow, 7).setValue(pi); changed.push('pairsIssued=' + pi); }   // PAIRS_ISSUED
    }
    if (data.pairsReceived !== undefined && safeStr(data.pairsReceived) !== '') {
      var prc = safeNum(data.pairsReceived);
      if (prc >= 0) {
        // 8.B2: on a sized card the override follows the same rule as a normal
        // receive — the new total must arrive with its absolute per-size split,
        // which REPLACES RECEIVED_BREAKDOWN (col 18).
        var _isb = {}; try { _isb = JSON.parse(safeStr(rows[idx][8])) || {}; } catch(e) {}
        var _isbHasSizes = Object.keys(_isb).some(function(k){ return safeNum(_isb[k]) > 0; });
        if (_isbHasSizes) {
          var rb = (data.receivedBreakdown && typeof data.receivedBreakdown === 'object') ? data.receivedBreakdown : null;
          if (!rb) return { success:false, error:'This card tracks sizes — send receivedBreakdown (received per size, totalling ' + prc + ').' };
          var rbSum = 0, rbBad = '', rbClean = {};
          Object.keys(rb).forEach(function(k){
            var q = safeNum(rb[k]);
            if (q <= 0) return;
            rbClean[k] = q; rbSum += q;
            if (q > safeNum(_isb[k])) rbBad = k;
          });
          if (rbSum !== prc) return { success:false, error:'Received size breakdown totals ' + rbSum + ' but pairs received is ' + prc + ' — they must match.' };
          if (rbBad)         return { success:false, error:'Received more of size ' + rbBad + ' than was issued.' };
          ws.getRange(sheetRow, 18).setValue(JSON.stringify(rbClean));   // RECEIVED_BREAKDOWN
          changed.push('receivedBreakdown');
        }
        ws.getRange(sheetRow, 8).setValue(prc); changed.push('pairsReceived=' + prc);  // PAIRS_RECEIVED
      }
    }

    // Status
    if (data.status !== undefined && safeStr(data.status).trim()) {
      var st = safeStr(data.status).trim().toUpperCase();
      if (['ISSUED','PARTIAL','COMPLETE','PAYMENT_PENDING','PAID','CANCELLED'].indexOf(st) < 0)
        return { success:false, error:'Invalid status: ' + st };
      ws.getRange(sheetRow, 14).setValue(st);   // STATUS
      changed.push('status=' + st);
    }

    if (!changed.length) return { success:false, error:'Nothing to change' };
    SpreadsheetApp.flush();
    ['dashboardData_','storeScreenData_'].forEach(function(k){ try { CacheService.getScriptCache().remove(k + CONFIG.ENV); } catch(e) {} });
    return { success:true, changed:changed };
  } catch(e) {
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
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

function getMaxIssuableForStage(orderRef, movement) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var STAGE_ORDER = ['Cutting','Preparation','Fitter','Lasting','Packing','Dispatch'];
    var STAGE_DEPT_KEY = {
      'Cutting':'cutting','Preparation':'prep','Fitter':'fitter',
      'Lasting':'lasting','Packing':'finish','Dispatch':'dispatch'
    };
    var STAGE_OWN_MOVEMENTS = {
      'Cutting':['Cutting IN'],
      'Preparation':['Preparation IN'],
      'Fitter':['Fitter IN'],
      'Lasting':['Lasting IN'],
      'Packing':['Packing IN'],
      'Dispatch':['Dispatch IN']
    };
    var MOVEMENT_TO_STAGE = {
      'Cutting IN':'Cutting','Preparation IN':'Preparation',
      'Fitter IN':'Fitter','Lasting IN':'Lasting',
      'Packing IN':'Packing','Dispatch IN':'Dispatch'
    };

    var currentStage = MOVEMENT_TO_STAGE[movement] || '';
    if (!currentStage) return { success:true, maxIssuable:0, source:'unknown' };

    var currentStageIdx = STAGE_ORDER.indexOf(currentStage);
    var orderActRes = getApprovedActivitiesForArticle(orderRef, ss);
    var orderActiveDepts = {};
    if (orderActRes && orderActRes.success && Array.isArray(orderActRes.activities)) {
      orderActRes.activities.forEach(function(a) {
        var dk = deptKeyOf(a.dept); // S.7: canonical short key
        Object.keys(STAGE_DEPT_KEY).forEach(function(stageName) {
          if (STAGE_DEPT_KEY[stageName] === dk) orderActiveDepts[stageName] = true;
        });
      });
    }

    var predecessorStage = null;
    for (var si = currentStageIdx - 1; si >= 0; si--) {
      if (orderActiveDepts[STAGE_ORDER[si]]) { predecessorStage = STAGE_ORDER[si]; break; }
    }

    var allJCs = getJobCards({orderRef: orderRef}, ss);
    if (!Array.isArray(allJCs)) allJCs = [];

    var thisStageMovements = STAGE_OWN_MOVEMENTS[currentStage] || [];
    var thisStageAlreadyIssued = 0;
    allJCs.forEach(function(jc) {
      if (thisStageMovements.indexOf(jc.movement) >= 0) {
        var st = safeStr(jc.status).toUpperCase();
        if (st !== 'CANCELLED') thisStageAlreadyIssued += safeNum(jc.pairsIssued);
      }
    });

    if (predecessorStage) {
      var predMovements = STAGE_OWN_MOVEMENTS[predecessorStage] || [];
      var predReceived = 0;
      allJCs.forEach(function(jc) {
        if (predMovements.indexOf(jc.movement) >= 0) {
          var st = safeStr(jc.status).toUpperCase();
          // Count pairs physically received back from the predecessor stage.
          // PARTIAL cards hold real received pairs that are available downstream,
          // so they must be included — excluding them made a partially-received
          // Cutting card show 0 available for Preparation (bug 8.B1).
          if (st === 'PARTIAL' || st === 'COMPLETE' || st === 'PAYMENT_PENDING' || st === 'PAID') {
            predReceived += safeNum(jc.pairsReceived);
          }
        }
      });
      return {
        success:true,
        maxIssuable: Math.max(0, predReceived - thisStageAlreadyIssued),
        source: predecessorStage + ' received',
        predReceived: predReceived,
        alreadyIssued: thisStageAlreadyIssued
      };
    } else {
      var oi = ss.getSheetByName('ORDER_INDEX');
      var orderLotSize = 0;
      if (oi && oi.getLastRow() > 3) {
        var oiRows = oi.getRange(4, 1, oi.getLastRow()-3, 9).getValues();
        for (var oiR = 0; oiR < oiRows.length; oiR++) {
          if (safeStr(oiRows[oiR][1]).trim() === orderRef) {
            orderLotSize = safeNum(oiRows[oiR][8]);
            break;
          }
        }
      }
      return {
        success:true,
        maxIssuable: Math.max(0, orderLotSize - thisStageAlreadyIssued),
        source:'order lot size',
        orderLotSize: orderLotSize,
        alreadyIssued: thisStageAlreadyIssued
      };
    }
  } catch(e) {
    return { success:false, error:e.message, maxIssuable:0 };
  }
}

// Maintenance / migration — bring existing job cards in line with the per-size
// rules. (1) Rescale any SIZE_BREAKDOWN that doesn't total PAIRS_ISSUED (corrupt
// pre-guard data, e.g. 60 vs 30) proportionally to PAIRS_ISSUED. (2) Backfill
// RECEIVED_BREAKDOWN for COMPLETE cards that have none — a complete card received
// every issued pair, so its received sizes equal its (corrected) issued sizes.
// PARTIAL cards are left alone (which sizes returned is unknown; getOrderSizeBalance
// falls back safely). Idempotent; skips CANCELLED. Pass a spreadsheet to target a
// specific environment (see runLiveJobCardMigration); defaults to the ENV sheet.
function repairMismatchedSizeBreakdowns(ss) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var ws = ensureJobCardsSheet(ss);
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return { success:true, rescaled:0, backfilled:0, details:[] };
    var rows = ws.getRange(2, 1, lastRow-1, 18).getValues();
    var details = [], rescaled = 0, backfilled = 0;
    for (var i = 0; i < rows.length; i++) {
      var r  = rows[i];
      var id = safeStr(r[0]).trim();
      if (!id) continue;
      var status = safeStr(r[13]).trim().toUpperCase();
      if (status === 'CANCELLED') continue;
      var pairs = safeNum(r[6]);
      var sb = {};
      try { sb = JSON.parse(safeStr(r[8])) || {}; } catch(e) { sb = {}; }
      var keys = Object.keys(sb);
      var effectiveSb = sb;

      // (1) Rescale a mismatched SIZE_BREAKDOWN to total PAIRS_ISSUED.
      if (keys.length && pairs > 0) {
        var sum = 0; keys.forEach(function(k){ sum += safeNum(sb[k]); });
        if (sum > 0 && sum !== pairs) {
          var ratio = pairs / sum, scaled = {}, running = 0;
          keys.forEach(function(k){ var v = Math.floor(safeNum(sb[k]) * ratio); scaled[k] = v; running += v; });
          var order = keys.slice().sort(function(a,b){ return safeNum(sb[b]) - safeNum(sb[a]); });
          var remN = pairs - running, oi = 0;
          while (remN > 0 && order.length) { scaled[order[oi % order.length]] += 1; remN--; oi++; }
          var clean = {};
          Object.keys(scaled).forEach(function(k){ if (scaled[k] > 0) clean[k] = scaled[k]; });
          ws.getRange(i+2, 9).setValue(JSON.stringify(clean));   // col I = SIZE_BREAKDOWN
          effectiveSb = clean; rescaled++;
          details.push({ jobCardId:id, action:'rescaled', pairsIssued:pairs, oldSum:sum, newBreakdown:clean });
        }
      }

      // (2) Backfill RECEIVED_BREAKDOWN for COMPLETE cards that have none.
      var rsb = {};
      try { rsb = JSON.parse(safeStr(r[17])) || {}; } catch(e) { rsb = {}; }
      if (status === 'COMPLETE' && Object.keys(effectiveSb).length && !Object.keys(rsb).length) {
        ws.getRange(i+2, 18).setValue(JSON.stringify(effectiveSb));   // col R = RECEIVED_BREAKDOWN
        backfilled++;
        details.push({ jobCardId:id, action:'backfilled_received', received:effectiveSb });
      }
    }
    SpreadsheetApp.flush();
    ['dashboardData_DEV','dashboardData_LIVE','storeScreenData_DEV','storeScreenData_LIVE'].forEach(function(k){ try { CacheService.getScriptCache().remove(k); } catch(e) {} });
    Logger.log('repairMismatchedSizeBreakdowns: rescaled ' + rescaled + ', backfilled ' + backfilled + ' — ' + JSON.stringify(details));
    return { success:true, rescaled:rescaled, backfilled:backfilled, details:details };
  } catch(e) {
    return { success:false, error:e.message };
  } finally {
    lock.releaseLock();
  }
}

// Run the job-card migration against the LIVE sheet explicitly (independent of
// CONFIG.ENV) so it's safe to run from the editor while HEAD stays on DEV. Use
// AFTER promoting new code to LIVE. Run manually: Run ▸ runLiveJobCardMigration.
function runLiveJobCardMigration() {
  var liveSs = SpreadsheetApp.openById(CONFIG.LIVE_SHEET_ID);
  var res = repairMismatchedSizeBreakdowns(liveSs);
  Logger.log('runLiveJobCardMigration → ' + JSON.stringify(res));
  return res;
}
