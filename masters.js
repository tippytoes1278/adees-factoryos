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
