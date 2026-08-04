// Phase B — master data tables. Suppliers/Materials/Rates/Styles.
// All ensure-functions S.8-hardened; all writers locked; IDs monotonic;
// dept keys canonical short-form.

// ── B.4 SUPPLIER MASTER ───────────────────────────────────────────────────────

// S.8: fast path returns with no lock; a wrong header on a sheet WITH data
// throws and touches nothing; headers are bootstrapped only on a missing or
// genuinely empty sheet, inside the script lock with a re-check.
function ensureSuppliersSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var HEADERS = ['SUP_ID','SUPPLIER_NAME','CONTACT_PERSON','PHONE','GST_NO','MATERIALS','PAYMENT_TERMS','STATUS','CREATED_AT'];
  var ws = ss.getSheetByName('MASTER_SUPPLIERS');
  if (ws && safeStr(ws.getRange(1, 1).getValue()) === 'SUP_ID') return ws;   // fast path — no lock
  if (ws && ws.getLastRow() > 1) {
    throw new Error('Supplier sheet header mismatch — the MASTER_SUPPLIERS sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('MASTER_SUPPLIERS');                              // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('MASTER_SUPPLIERS');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);
      ws.setFrozenRows(1);
    } else if (safeStr(ws.getRange(1, 1).getValue()) !== 'SUP_ID') {
      if (ws.getLastRow() > 1) throw new Error('Supplier sheet header mismatch — the MASTER_SUPPLIERS sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);                          // genuinely empty: bootstrap only
      ws.setFrozenRows(1);
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

// Read-only supplier list. No lock. Blank-ID rows are skipped.
function getSuppliers() {
  try {
    var ws = ensureSuppliersSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
    var result = [];
    rows.forEach(function(r) {
      var supId = safeStr(r[0]).trim();
      if (!supId) return;
      result.push({
        supId:         supId,
        name:          safeStr(r[1]).trim(),
        contactPerson: safeStr(r[2]).trim(),
        phone:         safeStr(r[3]).trim(),
        gstNo:         safeStr(r[4]).trim(),
        materials:     safeStr(r[5]).trim(),
        paymentTerms:  safeStr(r[6]).trim(),
        status:        safeStr(r[7]).trim().toUpperCase() || 'ACTIVE'
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Create or edit a supplier. payload: { supId?, name, contactPerson, phone,
// gstNo, materials, paymentTerms, status?, allowDuplicateName? }.
// With supId → EDIT in place (cols B-H only; A and I are never touched).
// S.9-style duplicate-name guard with admin override (see saveContractor).
function saveSupplier(payload) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  // S.8/S.6 convention: ensure* runs BEFORE the writer's lock — it self-locks
  // its one-time mutation path, and nesting script locks is undefined.
  var ws;
  try { ws = ensureSuppliersSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var newName = safeStr(payload && payload.name).trim();
    if (!newName) return { success: false, error: 'Supplier name is required' };
    var editId = safeStr(payload && payload.supId).trim();

    var rows = ws.getLastRow() > 1
      ? ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues()
      : [];

    // Duplicate-name guard: refuse a name matching an existing ACTIVE supplier
    // (renames included, excluding the row being edited) unless the admin
    // deliberately resubmits with allowDuplicateName:true.
    if (!(payload.allowDuplicateName === true)) {
      for (var di = 0; di < rows.length; di++) {
        var dId = safeStr(rows[di][0]).trim();
        var dNm = safeStr(rows[di][1]).trim();
        var dSt = safeStr(rows[di][7]).trim().toUpperCase();
        if (!dId || !dNm) continue;
        if (editId && dId === editId) continue;                              // renames: skip own row
        if (dNm.toLowerCase() === newName.toLowerCase() && dSt !== 'ARCHIVED') {
          return { success: false, duplicateOf: dId,
                   error: 'A supplier named "' + dNm + '" already exists (' + dId + '). Duplicate supplier names confuse material sourcing and payments. Use the existing supplier, or Ayush can deliberately force a duplicate.' };
        }
      }
    }

    var vals = [
      newName,
      safeStr(payload.contactPerson).trim(),
      safeStr(payload.phone).trim(),
      safeStr(payload.gstNo).trim(),
      safeStr(payload.materials).trim(),
      safeStr(payload.paymentTerms).trim()
    ];

    if (editId) {
      // EDIT: locate the row by SUP_ID and update cols B-H in place.
      var found = -1;
      for (var i = 0; i < rows.length; i++) {
        if (safeStr(rows[i][0]).trim() === editId) { found = i + 2; break; }
      }
      if (found < 0) return { success: false, error: 'Supplier ' + editId + ' not found' };
      var newStatus = safeStr(payload.status).trim().toUpperCase()
                   || safeStr(rows[found - 2][7]).trim().toUpperCase() || 'ACTIVE';
      if (newStatus !== 'ACTIVE' && newStatus !== 'ARCHIVED')
        return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
      ws.getRange(found, 2, 1, 7).setValues([vals.concat([newStatus])]);     // B-H only; never A or I
      SpreadsheetApp.flush();
      return { success: true, supId: editId };
    }

    // NEW: monotonic SUP-#### id — max existing numeric suffix + 1, never row-count.
    var maxNum = 0;
    rows.forEach(function(r) {
      var existing = safeStr(r[0]).trim();
      if (/^SUP-\d+$/.test(existing)) {
        var n = parseInt(existing.replace('SUP-', ''), 10);
        if (n > maxNum) maxNum = n;
      }
    });
    var seq = String(maxNum + 1);
    while (seq.length < 4) seq = '0' + seq;
    var supId = 'SUP-' + seq;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    ws.getRange(ws.getLastRow() + 1, 1, 1, 9).setValues([
      [supId].concat(vals).concat(['ACTIVE', now])
    ]);
    SpreadsheetApp.flush();
    return { success: true, supId: supId };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// Archive / reactivate a supplier. Admin only; locked; ACTIVE/ARCHIVED only.
function setSupplierStatus(supId, status) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var id = safeStr(supId).trim();
  if (!id) return { success: false, error: 'supId is required' };
  var st = safeStr(status).trim().toUpperCase();
  if (st !== 'ACTIVE' && st !== 'ARCHIVED')
    return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
  var ws;
  try { ws = ensureSuppliersSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    if (ws.getLastRow() < 2) return { success: false, error: 'Supplier ' + id + ' not found' };
    var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (safeStr(ids[i][0]).trim() === id) {
        ws.getRange(i + 2, 8).setValue(st);                                  // col H = STATUS
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Supplier ' + id + ' not found' };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// ── B.3 MATERIAL MASTER ───────────────────────────────────────────────────────

// S.8: same shape as ensureSuppliersSheet. Note: MASTER_MATERIALS and
// MATERIAL_RATE_LOG both key A1='MAT_ID' — safe because each ensure* looks up
// its own sheet by NAME first, then checks that sheet's A1.
function ensureMaterialsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var HEADERS = ['MAT_ID','MATERIAL_NAME','UNIT','CURRENT_RATE','SUPPLIER_ID','NOTES','STATUS','UPDATED_AT'];
  var ws = ss.getSheetByName('MASTER_MATERIALS');
  if (ws && safeStr(ws.getRange(1, 1).getValue()) === 'MAT_ID') return ws;    // fast path — no lock
  if (ws && ws.getLastRow() > 1) {
    throw new Error('Material sheet header mismatch — the MASTER_MATERIALS sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('MASTER_MATERIALS');                               // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('MASTER_MATERIALS');
      ws.getRange(1, 1, 1, 8).setValues([HEADERS]);
      ws.setFrozenRows(1);
    } else if (safeStr(ws.getRange(1, 1).getValue()) !== 'MAT_ID') {
      if (ws.getLastRow() > 1) throw new Error('Material sheet header mismatch — the MASTER_MATERIALS sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
      ws.getRange(1, 1, 1, 8).setValues([HEADERS]);                           // genuinely empty: bootstrap only
      ws.setFrozenRows(1);
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

// S.8: append-only rate history for materials.
function ensureMaterialRateLogSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var HEADERS = ['MAT_ID','RATE','DATE','UPDATED_BY'];
  var ws = ss.getSheetByName('MATERIAL_RATE_LOG');
  if (ws && safeStr(ws.getRange(1, 1).getValue()) === 'MAT_ID') return ws;    // fast path — no lock
  if (ws && ws.getLastRow() > 1) {
    throw new Error('Rate log header mismatch — the MATERIAL_RATE_LOG sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('MATERIAL_RATE_LOG');                              // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('MATERIAL_RATE_LOG');
      ws.getRange(1, 1, 1, 4).setValues([HEADERS]);
      ws.setFrozenRows(1);
    } else if (safeStr(ws.getRange(1, 1).getValue()) !== 'MAT_ID') {
      if (ws.getLastRow() > 1) throw new Error('Rate log header mismatch — the MATERIAL_RATE_LOG sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
      ws.getRange(1, 1, 1, 4).setValues([HEADERS]);                           // genuinely empty: bootstrap only
      ws.setFrozenRows(1);
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

// Read-only material list. No lock. Blank-ID rows are skipped.
function getMaterials() {
  try {
    var ws = ensureMaterialsSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 8).getValues();
    var result = [];
    rows.forEach(function(r) {
      var matId = safeStr(r[0]).trim();
      if (!matId) return;
      result.push({
        matId:      matId,
        name:       safeStr(r[1]).trim(),
        unit:       safeStr(r[2]).trim(),
        rate:       safeNum(r[3]),
        supplierId: safeStr(r[4]).trim(),
        notes:      safeStr(r[5]).trim(),
        status:     safeStr(r[6]).trim().toUpperCase() || 'ACTIVE'
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Create or edit a material. payload: { matId?, name, unit, rate, supplierId,
// notes, status?, allowDuplicateName? }. With matId → EDIT in place (col A is
// never touched). RATE HISTORY RULE: on create, and on any edit where the rate
// value changes, a row is appended to MATERIAL_RATE_LOG inside the same lock,
// after the master write. Duplicate-name guard as saveSupplier.
function saveMaterial(payload) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var VALID_UNITS = ['mtr','kg','pair','pc','roll'];
  // S.8/S.6 convention: BOTH ensure* run BEFORE the writer's lock — they
  // self-lock their one-time mutation path, and nesting script locks is undefined.
  var ws, logWs;
  try { ws = ensureMaterialsSheet(); logWs = ensureMaterialRateLogSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var newName = safeStr(payload && payload.name).trim();
    if (!newName) return { success: false, error: 'Material name is required' };
    var unit = safeStr(payload && payload.unit).trim().toLowerCase();
    if (VALID_UNITS.indexOf(unit) < 0)
      return { success: false, error: 'Unit must be one of: ' + VALID_UNITS.join(', ') };
    var rate = safeNum(payload && payload.rate);
    var editId = safeStr(payload && payload.matId).trim();

    var rows = ws.getLastRow() > 1
      ? ws.getRange(2, 1, ws.getLastRow() - 1, 8).getValues()
      : [];

    // Duplicate-name guard: refuse a name matching an existing ACTIVE material
    // (renames included, excluding the row being edited) unless the admin
    // deliberately resubmits with allowDuplicateName:true.
    if (!(payload.allowDuplicateName === true)) {
      for (var di = 0; di < rows.length; di++) {
        var dId = safeStr(rows[di][0]).trim();
        var dNm = safeStr(rows[di][1]).trim();
        var dSt = safeStr(rows[di][6]).trim().toUpperCase();
        if (!dId || !dNm) continue;
        if (editId && dId === editId) continue;                              // renames: skip own row
        if (dNm.toLowerCase() === newName.toLowerCase() && dSt !== 'ARCHIVED') {
          return { success: false, duplicateOf: dId,
                   error: 'A material named "' + dNm + '" already exists (' + dId + '). Duplicate material names confuse BOMs and rate history. Use the existing material, or Ayush can deliberately force a duplicate.' };
        }
      }
    }

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var vals = [
      newName,
      unit,
      rate,
      safeStr(payload.supplierId).trim(),
      safeStr(payload.notes).trim()
    ];

    if (editId) {
      // EDIT: locate the row by MAT_ID and update cols B-H in place (A never).
      var found = -1;
      for (var i = 0; i < rows.length; i++) {
        if (safeStr(rows[i][0]).trim() === editId) { found = i + 2; break; }
      }
      if (found < 0) return { success: false, error: 'Material ' + editId + ' not found' };
      var oldRate = safeNum(rows[found - 2][3]);
      var newStatus = safeStr(payload.status).trim().toUpperCase()
                   || safeStr(rows[found - 2][6]).trim().toUpperCase() || 'ACTIVE';
      if (newStatus !== 'ACTIVE' && newStatus !== 'ARCHIVED')
        return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
      ws.getRange(found, 2, 1, 7).setValues([vals.concat([newStatus, now])]); // B-H only; never A
      // RATE HISTORY: log only when the rate VALUE actually changed.
      if (rate !== oldRate) {
        logWs.getRange(logWs.getLastRow() + 1, 1, 1, 4)
             .setValues([[editId, rate, now, user.email]]);
      }
      SpreadsheetApp.flush();
      return { success: true, matId: editId };
    }

    // NEW: monotonic MAT-#### id — max existing numeric suffix + 1, never row-count.
    var maxNum = 0;
    rows.forEach(function(r) {
      var existing = safeStr(r[0]).trim();
      if (/^MAT-\d+$/.test(existing)) {
        var n = parseInt(existing.replace('MAT-', ''), 10);
        if (n > maxNum) maxNum = n;
      }
    });
    var seq = String(maxNum + 1);
    while (seq.length < 4) seq = '0' + seq;
    var matId = 'MAT-' + seq;
    ws.getRange(ws.getLastRow() + 1, 1, 1, 8).setValues([
      [matId].concat(vals).concat(['ACTIVE', now])
    ]);
    // RATE HISTORY: every new material logs its opening rate.
    logWs.getRange(logWs.getLastRow() + 1, 1, 1, 4)
         .setValues([[matId, rate, now, user.email]]);
    SpreadsheetApp.flush();
    return { success: true, matId: matId };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// Archive / reactivate a material. Admin only; locked; ACTIVE/ARCHIVED only.
function setMaterialStatus(matId, status) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var id = safeStr(matId).trim();
  if (!id) return { success: false, error: 'matId is required' };
  var st = safeStr(status).trim().toUpperCase();
  if (st !== 'ACTIVE' && st !== 'ARCHIVED')
    return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
  var ws;
  try { ws = ensureMaterialsSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    if (ws.getLastRow() < 2) return { success: false, error: 'Material ' + id + ' not found' };
    var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (safeStr(ids[i][0]).trim() === id) {
        ws.getRange(i + 2, 7).setValue(st);                                  // col G = STATUS
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Material ' + id + ' not found' };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// ── B.2 LABOUR RATE CARD ──────────────────────────────────────────────────────
// APPEND-ONLY SUPERSEDE SEMANTICS: rates are never edited in place. Every save
// appends a new LR-#### row (STATUS=ACTIVE) and, if an ACTIVE row already
// existed for the same (stage, activity, article) tuple, that old row gets a
// single-cell STATUS='SUPERSEDED' write — no other cell of a historical row is
// ever touched. getStandardRate() is THE single lookup entry point for costing.

var RATE_STAGES = ['cutting','prep','fitter','lasting','finish','dispatch'];

// Normalize an EFFECTIVE_FROM cell (Date object or string) to 'yyyy-MM-dd'.
function _rateDateStr_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]')
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return safeStr(v).trim();
}

// S.8: same shape as ensureSuppliersSheet — fast path returns with no lock; a
// wrong header on a sheet WITH data throws and touches nothing; headers are
// bootstrapped only on a missing or genuinely empty sheet, inside the script
// lock with a re-check.
function ensureRatesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var HEADERS = ['RATE_ID','STAGE','ACTIVITY','ARTICLE_ID','RATE_PER_PAIR','COMMISSION_PER_PAIR','EFFECTIVE_FROM','STATUS','CREATED_AT'];
  var ws = ss.getSheetByName('MASTER_RATES');
  if (ws && safeStr(ws.getRange(1, 1).getValue()) === 'RATE_ID') return ws;   // fast path — no lock
  if (ws && ws.getLastRow() > 1) {
    throw new Error('Rates sheet header mismatch — the MASTER_RATES sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('MASTER_RATES');                                   // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('MASTER_RATES');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);
      ws.setFrozenRows(1);
    } else if (safeStr(ws.getRange(1, 1).getValue()) !== 'RATE_ID') {
      if (ws.getLastRow() > 1) throw new Error('Rates sheet header mismatch — the MASTER_RATES sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);                           // genuinely empty: bootstrap only
      ws.setFrozenRows(1);
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

// Read-only rate list. No lock. Blank-ID rows are skipped. STAGE is stored
// canonical short-form; both stored and filter values go through deptKeyOf so
// comparisons never depend on how a stage was typed.
// filters (all optional): { activity, articleId, stage, activeOnly }.
function getRates(filters) {
  try {
    var ws = ensureRatesSheet();
    if (ws.getLastRow() < 2) return [];
    var f = filters || {};
    var fStage = safeStr(f.stage).trim() ? deptKeyOf(f.stage) : '';
    var fAct   = safeStr(f.activity).trim().toLowerCase();
    var fArt   = safeStr(f.articleId).trim();
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
    var result = [];
    rows.forEach(function(r) {
      var rateId = safeStr(r[0]).trim();
      if (!rateId) return;
      var stage    = deptKeyOf(r[1]);
      var activity = safeStr(r[2]).trim();
      var articleId = safeStr(r[3]).trim();
      var status   = safeStr(r[7]).trim().toUpperCase() || 'ACTIVE';
      if (fStage && stage !== fStage) return;
      if (fAct && activity.toLowerCase() !== fAct) return;
      if (fArt && articleId !== fArt) return;
      if (f.activeOnly === true && status !== 'ACTIVE') return;
      result.push({
        rateId:        rateId,
        stage:         stage,
        activity:      activity,
        articleId:     articleId,
        rate:          safeNum(r[4]),
        comm:          safeNum(r[5]),
        effectiveFrom: _rateDateStr_(r[6]),
        status:        status
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Save a labour rate. Admin only; locked. payload: { stage, activity,
// articleId?, rate, comm?, effectiveFrom? }. APPEND-ONLY: always appends a new
// LR-#### row (max existing suffix + 1, never row-count) with STATUS=ACTIVE.
// If an ACTIVE row already exists for the same (deptKeyOf(stage), activity
// case-insens-trimmed, articleId) tuple, THAT row's STATUS cell — and only
// that cell — is set to 'SUPERSEDED' before the append, inside the same lock.
// articleId '' means the base/standard rate for the activity.
function saveRate(payload) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  // S.8/S.6 convention: ensure* runs BEFORE the writer's lock — it self-locks
  // its one-time mutation path, and nesting script locks is undefined.
  var ws;
  try { ws = ensureRatesSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var stage = deptKeyOf(payload && payload.stage);
    if (RATE_STAGES.indexOf(stage) < 0)
      return { success: false, error: 'Stage must be one of: ' + RATE_STAGES.join(', ') };
    var activity = safeStr(payload && payload.activity).trim();
    if (!activity) return { success: false, error: 'Activity is required' };
    var articleId = safeStr(payload && payload.articleId).trim();
    var rate = safeNum(payload && payload.rate);
    var comm = safeNum(payload && payload.comm);
    var effFrom = safeStr(payload && payload.effectiveFrom).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(effFrom))
      effFrom = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var rows = ws.getLastRow() > 1
      ? ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues()
      : [];

    // One pass: track max LR suffix AND find the ACTIVE row this save supersedes.
    var maxNum = 0, supersededId = '';
    for (var i = 0; i < rows.length; i++) {
      var rId = safeStr(rows[i][0]).trim();
      if (/^LR-\d+$/.test(rId)) {
        var n = parseInt(rId.replace('LR-', ''), 10);
        if (n > maxNum) maxNum = n;
      }
      if (!rId) continue;
      if (safeStr(rows[i][7]).trim().toUpperCase() !== 'ACTIVE') continue;
      if (deptKeyOf(rows[i][1]) !== stage) continue;
      if (safeStr(rows[i][2]).trim().toLowerCase() !== activity.toLowerCase()) continue;
      if (safeStr(rows[i][3]).trim() !== articleId) continue;
      // SUPERSEDE: single-cell STATUS write. NEVER any other cell of a
      // historical row — history stays byte-for-byte intact.
      ws.getRange(i + 2, 8).setValue('SUPERSEDED');                          // col H = STATUS
      supersededId = rId;
    }

    var seq = String(maxNum + 1);
    while (seq.length < 4) seq = '0' + seq;
    var rateId = 'LR-' + seq;
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    ws.getRange(ws.getLastRow() + 1, 1, 1, 9).setValues([
      [rateId, stage, activity, articleId, rate, comm, effFrom, 'ACTIVE', now]
    ]);
    SpreadsheetApp.flush();
    var out = { success: true, rateId: rateId };
    if (supersededId) out.supersededId = supersededId;
    return out;
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

/**
 * getStandardRate(activity, articleId, onDate) — THE single rate-lookup entry
 * point for all future costing. Any code that needs "what does this activity
 * pay per pair" must call this function and nothing else.
 *
 * CONTRACT:
 * (a) considers only STATUS='ACTIVE' rows with matching activity (case-insens trim);
 * (b) rows with EFFECTIVE_FROM <= onDate (onDate optional, default today;
 *     accepts Date or 'yyyy-mm-dd');
 * (c) most specific wins: a row whose ARTICLE_ID matches the passed articleId
 *     beats a blank-ARTICLE_ID base row;
 * (d) among equals, latest EFFECTIVE_FROM wins, ties broken by highest RATE_ID
 *     sequence;
 * (e) returns {found:true, rateId, stage, rate, comm, articleSpecific:bool,
 *     effectiveFrom} or {found:false}.
 *
 * Read-only, no lock.
 */
function getStandardRate(activity, articleId, onDate) {
  try {
    var act = safeStr(activity).trim().toLowerCase();
    if (!act) return { found: false };
    var art = safeStr(articleId).trim();
    var onStr;
    if (Object.prototype.toString.call(onDate) === '[object Date]')
      onStr = Utilities.formatDate(onDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    else
      onStr = safeStr(onDate).trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(onStr))
      onStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var ws = ensureRatesSheet();
    if (ws.getLastRow() < 2) return { found: false };
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
    var best = null, bestSpecific = false, bestEff = '', bestSeq = -1;
    rows.forEach(function(r) {
      var rateId = safeStr(r[0]).trim();
      if (!rateId) return;
      if (safeStr(r[7]).trim().toUpperCase() !== 'ACTIVE') return;           // (a)
      if (safeStr(r[2]).trim().toLowerCase() !== act) return;                // (a)
      var rowArt = safeStr(r[3]).trim();
      var specific = (art !== '' && rowArt === art);
      if (rowArt !== '' && !specific) return;      // article-specific row for a DIFFERENT article never applies
      var eff = _rateDateStr_(r[6]);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(eff)) return;                          // unreadable date: skip
      if (eff > onStr) return;                                               // (b) not yet effective
      var seq = /^LR-\d+$/.test(rateId) ? parseInt(rateId.replace('LR-', ''), 10) : 0;
      // (c) specific beats base; (d) then latest EFFECTIVE_FROM; then highest seq.
      var wins = false;
      if (!best) wins = true;
      else if (specific !== bestSpecific) wins = specific;
      else if (eff !== bestEff) wins = (eff > bestEff);
      else wins = (seq > bestSeq);
      if (wins) { best = r; bestSpecific = specific; bestEff = eff; bestSeq = seq; }
    });
    if (!best) return { found: false };
    return {
      found:           true,
      rateId:          safeStr(best[0]).trim(),
      stage:           deptKeyOf(best[1]),
      rate:            safeNum(best[4]),
      comm:            safeNum(best[5]),
      articleSpecific: bestSpecific,
      effectiveFrom:   bestEff
    };
  } catch(e) {
    return { found: false, error: e.message };
  }
}

// Read-only rate history for one material, newest first. No lock.
function getMaterialRateLog(matId) {
  try {
    var id = safeStr(matId).trim();
    if (!id) return [];
    var ws = ensureMaterialRateLogSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 4).getValues();
    var result = [];
    rows.forEach(function(r) {
      if (safeStr(r[0]).trim() !== id) return;
      var d = r[2];
      var dateStr = (Object.prototype.toString.call(d) === '[object Date]')
        ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm')
        : safeStr(d).trim();
      result.push({
        rate:      safeNum(r[1]),
        date:      dateStr,
        updatedBy: safeStr(r[3]).trim()
      });
    });
    result.reverse();                                                        // appended chronologically → newest first
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ── B.1 STYLE MASTER ──────────────────────────────────────────────────────────
// SIZE_RUN_DEFAULT holds a JSON object whose KEYS are size tokens and whose
// values are all 0 — the style stores WHICH sizes exist, never quantities.

var STYLE_CATEGORIES = ['Men','Ladies','Toddlers','Junior'];

// S.8: same shape as ensureSuppliersSheet — fast path returns with no lock; a
// wrong header on a sheet WITH data throws and touches nothing; headers are
// bootstrapped only on a missing or genuinely empty sheet, inside the script
// lock with a re-check.
function ensureStylesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var HEADERS = ['STYLE_ID','STYLE_NAME','CATEGORY','GRADING_DEFAULT','SIZE_RUN_DEFAULT','CONSTRUCTION_NOTES','STATUS','CREATED_AT','UPDATED_AT'];
  var ws = ss.getSheetByName('MASTER_STYLES');
  if (ws && safeStr(ws.getRange(1, 1).getValue()) === 'STYLE_ID') return ws;  // fast path — no lock
  if (ws && ws.getLastRow() > 1) {
    throw new Error('Style sheet header mismatch — the MASTER_STYLES sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    ws = ss.getSheetByName('MASTER_STYLES');                                  // re-check inside lock
    if (!ws) {
      ws = ss.insertSheet('MASTER_STYLES');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);
      ws.setFrozenRows(1);
    } else if (safeStr(ws.getRange(1, 1).getValue()) !== 'STYLE_ID') {
      if (ws.getLastRow() > 1) throw new Error('Style sheet header mismatch — the MASTER_STYLES sheet has data but its header row is wrong. Nothing was changed. Call Ayush.');
      ws.getRange(1, 1, 1, 9).setValues([HEADERS]);                           // genuinely empty: bootstrap only
      ws.setFrozenRows(1);
    }
    return ws;
  } finally {
    lock.releaseLock();
  }
}

// Read-only style list. No lock. Blank-ID rows are skipped. sizeRun is the
// parsed SIZE_RUN_DEFAULT object; an unreadable cell degrades to {}.
function getStyles() {
  try {
    var ws = ensureStylesSheet();
    if (ws.getLastRow() < 2) return [];
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
    var result = [];
    rows.forEach(function(r) {
      var styleId = safeStr(r[0]).trim();
      if (!styleId) return;
      var sizeRun = {};
      try {
        var parsed = JSON.parse(safeStr(r[4]).trim() || '{}');
        if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) sizeRun = parsed;
      } catch(pe) { sizeRun = {}; }
      result.push({
        styleId:  styleId,
        name:     safeStr(r[1]).trim(),
        category: safeStr(r[2]).trim(),
        grading:  safeStr(r[3]).trim().toUpperCase(),
        sizeRun:  sizeRun,
        notes:    safeStr(r[5]).trim(),
        status:   safeStr(r[6]).trim().toUpperCase() || 'ACTIVE'
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// Create or edit a style. payload: { styleId?, name, category, grading,
// sizeRun, notes, status?, allowDuplicateName? }. sizeRun is an OBJECT whose
// keys are size tokens and whose values are all 0 (the style records sizes,
// never quantities) — stored as JSON. With styleId → EDIT in place (cols B-G
// plus UPDATED_AT col I; A and H are never touched). Duplicate-name guard
// with admin override, same shape as saveSupplier.
function saveStyle(payload) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  // S.8/S.6 convention: ensure* runs BEFORE the writer's lock — it self-locks
  // its one-time mutation path, and nesting script locks is undefined.
  var ws;
  try { ws = ensureStylesSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    var newName = safeStr(payload && payload.name).trim();
    if (!newName) return { success: false, error: 'Style name is required' };

    // Category: canonical-cased match against the fixed list.
    var catIn = safeStr(payload && payload.category).trim().toLowerCase();
    var category = '';
    STYLE_CATEGORIES.forEach(function(c) { if (c.toLowerCase() === catIn) category = c; });
    if (!category)
      return { success: false, error: 'Category must be one of: ' + STYLE_CATEGORIES.join(', ') };

    var grading = safeStr(payload && payload.grading).trim().toUpperCase();
    if (grading !== 'UK' && grading !== 'EU')
      return { success: false, error: 'Grading must be UK or EU' };

    // sizeRun: object of {token: 0}. Sizes only — any non-zero value means a
    // quantity is being smuggled into the style master, which is refused.
    var srIn = payload && payload.sizeRun;
    if (!srIn || typeof srIn !== 'object' || Array.isArray(srIn))
      return { success: false, error: 'Size run must be an object of size tokens' };
    var sizeRun = {};
    var srKeys = Object.keys(srIn);
    for (var si = 0; si < srKeys.length; si++) {
      var tok = safeStr(srKeys[si]).trim();
      if (!tok) continue;
      if (safeNum(srIn[srKeys[si]]) !== 0)
        return { success: false, error: 'Size run holds sizes only — quantities live on orders, not the style. All size values must be 0.' };
      sizeRun[tok] = 0;
    }

    var editId = safeStr(payload && payload.styleId).trim();

    var rows = ws.getLastRow() > 1
      ? ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues()
      : [];

    // Duplicate-name guard: refuse a name matching an existing ACTIVE style
    // (renames included, excluding the row being edited) unless the admin
    // deliberately resubmits with allowDuplicateName:true.
    if (!(payload.allowDuplicateName === true)) {
      for (var di = 0; di < rows.length; di++) {
        var dId = safeStr(rows[di][0]).trim();
        var dNm = safeStr(rows[di][1]).trim();
        var dSt = safeStr(rows[di][6]).trim().toUpperCase();
        if (!dId || !dNm) continue;
        if (editId && dId === editId) continue;                              // renames: skip own row
        if (dNm.toLowerCase() === newName.toLowerCase() && dSt !== 'ARCHIVED') {
          return { success: false, duplicateOf: dId,
                   error: 'A style named "' + dNm + '" already exists (' + dId + '). Duplicate style names confuse orders and costing. Use the existing style, or Ayush can deliberately force a duplicate.' };
        }
      }
    }

    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm');
    var vals = [
      newName,
      category,
      grading,
      JSON.stringify(sizeRun),
      safeStr(payload.notes).trim()
    ];

    if (editId) {
      // EDIT: locate the row by STYLE_ID; update cols B-G in place and refresh
      // UPDATED_AT (col I). A (id) and H (CREATED_AT) are never touched.
      var found = -1;
      for (var i = 0; i < rows.length; i++) {
        if (safeStr(rows[i][0]).trim() === editId) { found = i + 2; break; }
      }
      if (found < 0) return { success: false, error: 'Style ' + editId + ' not found' };
      var newStatus = safeStr(payload.status).trim().toUpperCase()
                   || safeStr(rows[found - 2][6]).trim().toUpperCase() || 'ACTIVE';
      if (newStatus !== 'ACTIVE' && newStatus !== 'ARCHIVED')
        return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
      ws.getRange(found, 2, 1, 6).setValues([vals.concat([newStatus])]);     // B-G only; never A or H
      ws.getRange(found, 9).setValue(now);                                   // col I = UPDATED_AT
      SpreadsheetApp.flush();
      return { success: true, styleId: editId };
    }

    // NEW: monotonic STY-#### id — max existing numeric suffix + 1, never row-count.
    var maxNum = 0;
    rows.forEach(function(r) {
      var existing = safeStr(r[0]).trim();
      if (/^STY-\d+$/.test(existing)) {
        var n = parseInt(existing.replace('STY-', ''), 10);
        if (n > maxNum) maxNum = n;
      }
    });
    var seq = String(maxNum + 1);
    while (seq.length < 4) seq = '0' + seq;
    var styleId = 'STY-' + seq;
    ws.getRange(ws.getLastRow() + 1, 1, 1, 9).setValues([
      [styleId].concat(vals).concat(['ACTIVE', now, now])
    ]);
    SpreadsheetApp.flush();
    return { success: true, styleId: styleId };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// Archive / reactivate a style. Admin only; locked; ACTIVE/ARCHIVED only.
function setStyleStatus(styleId, status) {
  var user = getUserInfo();
  if (user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var id = safeStr(styleId).trim();
  if (!id) return { success: false, error: 'styleId is required' };
  var st = safeStr(status).trim().toUpperCase();
  if (st !== 'ACTIVE' && st !== 'ARCHIVED')
    return { success: false, error: 'Status must be ACTIVE or ARCHIVED' };
  var ws;
  try { ws = ensureStylesSheet(); }
  catch(ee) { return { success: false, error: ee.message }; }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    try {
    if (ws.getLastRow() < 2) return { success: false, error: 'Style ' + id + ' not found' };
    var ids = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (safeStr(ids[i][0]).trim() === id) {
        ws.getRange(i + 2, 7).setValue(st);                                  // col G = STATUS
        ws.getRange(i + 2, 9).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm')); // col I = UPDATED_AT
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: 'Style ' + id + ' not found' };
    } catch(e) { return { success: false, error: e.message }; }
  } finally {
    lock.releaseLock();
  }
}

// B.2b: batch wrapper for the activity-setup screen pre-fill. Preserves the
// single-entry-point contract — every lookup goes through getStandardRate().
// Returns { '<activity name>': getStandardRate result } for unique names.
function getStandardRateBatch(names, articleId) {
  var out = {};
  try {
    (Array.isArray(names) ? names : []).forEach(function(n) {
      var key = safeStr(n).trim();
      if (!key || out[key]) return;
      out[key] = getStandardRate(key, articleId, null);
    });
    return out;
  } catch(e) { return { error: e.message }; }
}
