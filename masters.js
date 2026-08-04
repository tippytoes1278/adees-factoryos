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
