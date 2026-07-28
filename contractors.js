function getContractorsData(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    var contractors = [];
    if (mc && mc.getLastRow() > 3) {
      // S.9: include ctrId (col A) so the client never has to join by name.
      mc.getRange(4, 1, mc.getLastRow()-3, 7).getValues().forEach(function(r) {
        if (!r[1]) return;
        contractors.push({
          ctrId: safeStr(r[0]).trim(),
          name: safeStr(r[1]),
          paymentMethod: safeStr(r[2]) || 'Cash',
          status: safeStr(r[3]),
          dept: safeStr(r[4]),
          phone: safeStr(r[5])
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
    // S.9: refuse a name that duplicates an existing ACTIVE contractor — name
    // collisions are how payments get mis-attributed. Admin can deliberately
    // override by resubmitting with allowDuplicateName:true (the client offers
    // this only to the admin role after the rejection).
    var newName = safeStr(payload.name).trim();
    if (!newName) return { success: false, error: 'Contractor name is required' };
    if (mc.getLastRow() > 3) {
      var _dupRows = mc.getRange(4, 1, mc.getLastRow()-3, 4).getValues();
      for (var _di = 0; _di < _dupRows.length; _di++) {
        var _dn = safeStr(_dupRows[_di][1]).trim();
        var _dst = safeStr(_dupRows[_di][3]).trim().toUpperCase();
        if (_dn && _dn.toLowerCase() === newName.toLowerCase() && _dst !== 'INACTIVE') {
          if (!(payload.allowDuplicateName === true && user.role === 'admin')) {
            return { success: false, duplicateOf: safeStr(_dupRows[_di][0]).trim(),
                     error: 'A contractor named "' + _dn + '" already exists (' + safeStr(_dupRows[_di][0]).trim() + '). Use the existing contractor — duplicate names mis-attribute payments.' + (user.role === 'admin' ? '' : ' Only Ayush can override this.') };
          }
          break;
        }
      }
    }
    var nextCtrId = 'CTR-001';
    try {
      var mcRows = mc.getLastRow() > 3
        ? mc.getRange(4, 1, mc.getLastRow()-3, 1).getValues()
        : [];
      var maxNum = 0;
      mcRows.forEach(function(r) {
        var existing = safeStr(r[0]).trim();
        if (/^CTR-\d+$/.test(existing)) {
          var n = parseInt(existing.replace('CTR-', ''), 10);
          if (n > maxNum) maxNum = n;
        }
      });
      var nextNum = maxNum + 1;
      var seq = String(nextNum);
      while (seq.length < 3) seq = '0' + seq;
      nextCtrId = 'CTR-' + seq;
    } catch(cidErr) { Logger.log('CTR-ID gen error: ' + cidErr.message); }
    mc.getRange(mc.getLastRow() + 1, 1, 1, 7).setValues([[
      nextCtrId, safeStr(payload.name), safeStr(payload.paymentMethod) || 'Cash',
      'ACTIVE', safeStr(payload.dept), safeStr(payload.phone), now
    ]]);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
    return { success: true, ctrId: nextCtrId };
  } catch(e) { return { success: false, error: e.message }; }
}

function getContractors(ss) {
  try {
    if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc || mc.getLastRow() < 4) return [];
    var rows = mc.getRange(4, 1, mc.getLastRow()-3, 4).getValues();
    var result = [];
    rows.forEach(function(r) {
      var name   = safeStr(r[1]).trim();
      var status = safeStr(r[3]).trim().toUpperCase();
      if (!name) return;
      if (status === 'INACTIVE') return;
      result.push({
        ctrId:         safeStr(r[0]).trim(),
        name:          name,
        paymentMethod: safeStr(r[2]).trim() || 'Cash',
        status:        status || 'ACTIVE'
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function assignContractorIds() {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc) return { success: false, error: 'MASTER_CONTRACTORS sheet not found' };

    var lastRow = mc.getLastRow();
    if (lastRow < 4) return { success: true, assigned: 0, duplicate: 'none' };

    // Write CTR-ID header into A3
    mc.getRange(3, 1).setValue('CTR-ID');

    // Read all data rows: cols A-G (1-7), starting row 4
    var numRows = lastRow - 3;
    var data = mc.getRange(4, 1, numRows, 7).getValues();

    // Detect duplicate "Jai Prakash Press" — keep first occurrence (lower row index)
    var jpSeen = -1;
    var duplicateRowNum = -1;
    for (var i = 0; i < data.length; i++) {
      var name = safeStr(data[i][1]).trim();
      if (name.toLowerCase() === 'jai prakash press') {
        if (jpSeen < 0) {
          jpSeen = i;
        } else {
          duplicateRowNum = i + 4;
          mc.getRange(duplicateRowNum, 4).setValue('INACTIVE');
          mc.getRange(duplicateRowNum, 5).setValue('DUPLICATE - REMOVED');
          data[i][3] = 'INACTIVE';
          Logger.log('Marked row ' + duplicateRowNum + ' (Jai Prakash Press) as DUPLICATE - REMOVED');
          break;
        }
      }
    }

    // Find highest existing CTR-nnn to avoid collisions on re-run
    var maxExisting = 0;
    for (var i = 0; i < data.length; i++) {
      var existing = safeStr(data[i][0]).trim();
      if (/^CTR-\d+$/.test(existing)) {
        var n = parseInt(existing.replace('CTR-', ''), 10);
        if (n > maxExisting) maxExisting = n;
      }
    }
    var nextId = maxExisting + 1;

    // Assign CTR-IDs to active, non-blank, ID-less rows
    var counter = 0;
    for (var i = 0; i < data.length; i++) {
      var ctrId  = safeStr(data[i][0]).trim();
      var rName  = safeStr(data[i][1]).trim();
      var status = safeStr(data[i][3]).trim().toUpperCase();
      if (!rName) continue;
      if (status === 'INACTIVE') continue;
      if (ctrId) continue;
      var newId = 'CTR-' + (String(nextId).padStart ? String(nextId).padStart(3,'0') : ('00'+nextId).slice(-3));
      mc.getRange(i + 4, 1).setValue(newId);
      Logger.log('Assigned ' + newId + ' → row ' + (i+4) + ': ' + rName);
      nextId++;
      counter++;
    }

    SpreadsheetApp.flush();
    var dupMsg = duplicateRowNum > 0 ? 'row ' + duplicateRowNum + ' marked INACTIVE' : 'none found';
    Logger.log('assignContractorIds done — assigned: ' + counter + ', duplicate: ' + dupMsg);
    return { success: true, assigned: counter, duplicate: dupMsg };
  } catch(e) {
    Logger.log('assignContractorIds error: ' + e.message);
    return { success: false, error: e.message };
  }
}

// One-time migration: rewrite contractor references on JOB_CARDS to CTR-IDs.
// Each contractor field is resolved: already a valid CTR-ID → keep; a name that
// matches MASTER_CONTRACTORS → replaced with its CTR-ID; anything unresolved →
// left untouched and reported for manual review. Pass dryRun=true to preview with
// NO writes. Handles both the primary CONTRACTOR_ID (col F) and the ASSIGNMENTS JSON
// (col Q). Idempotent — safe to re-run (CTR-IDs are skipped).
function migrateJobCardContractors(dryRun, targetEnv) {
  var u = getUserInfo();
  if (u.role !== 'admin') return { success:false, error:'Not authorised' };
  try {
    var _sid = (targetEnv === 'LIVE') ? CONFIG.LIVE_SHEET_ID
             : (targetEnv === 'DEV')  ? CONFIG.DEV_SHEET_ID
             : SHEET_ID;
    var ss = SpreadsheetApp.openById(_sid);
    // name(lowercased) → CTR-ID, and set of valid CTR-IDs
    var nameToId = {}, idSet = {}, _activeByName = {}, _dupActiveNames = [];
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() >= 4) {
      mc.getRange(4, 1, mc.getLastRow()-3, 4).getValues().forEach(function(r){
        var id = safeStr(r[0]).trim(), nm = safeStr(r[1]).trim();
        var st = safeStr(r[3]).trim().toUpperCase();
        if (id) {
          idSet[id] = true;
          if (nm) {
            var key = nm.toLowerCase();
            // S.9 hard guard: the name→id map is last-wins, so TWO ACTIVE rows
            // sharing a name would rewrite cards to the wrong twin's CTR-ID.
            if (st !== 'INACTIVE') {
              if (_activeByName[key]) _dupActiveNames.push(nm + ' (' + _activeByName[key] + ' vs ' + id + ')');
              _activeByName[key] = id;
            }
            nameToId[key] = _activeByName[key] || nameToId[key] || id;
          }
        }
      });
    }
    if (_dupActiveNames.length)
      return { success:false, error:'Refusing to run: duplicate ACTIVE contractor names would make name→ID mapping ambiguous — ' + _dupActiveNames.join('; ') + '. Deactivate or rename the duplicates first.' };
    function resolve(v) {
      var s = safeStr(v).trim();
      if (!s) return { id:s, changed:false, ok:true };
      if (/^CTR-\d+$/.test(s) && idSet[s]) return { id:s, changed:false, ok:true };
      var hit = nameToId[s.toLowerCase()];
      if (hit) return { id:hit, changed:(hit !== s), ok:true };
      return { id:s, changed:false, ok:false };
    }
    var report = { success:true, dryRun:!!dryRun, env:(targetEnv || CONFIG.ENV), updated:0, alreadyOk:0, unmatched:[] };
    var jc = ss.getSheetByName('JOB_CARDS');
    if (jc && jc.getLastRow() > 1) {
      var rows = jc.getRange(2, 1, jc.getLastRow()-1, 17).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i], sheetRow = i + 2, rowChanged = false, jcId = safeStr(r[0]).trim();
        if (!jcId) continue;
        var pr = resolve(r[5]);
        if (!pr.ok && safeStr(r[5]).trim()) report.unmatched.push(jcId + ' (contractor "' + safeStr(r[5]) + '")');
        if (pr.changed) { if (!dryRun) jc.getRange(sheetRow, 6).setValue(pr.id); rowChanged = true; }
        var asg = null;
        try { asg = JSON.parse(safeStr(r[16])); } catch(e) {}
        if (Array.isArray(asg) && asg.length) {
          var asgChanged = false;
          asg.forEach(function(a) {
            var ar = resolve(a.contractorId);
            if (!ar.ok && safeStr(a.contractorId).trim()) report.unmatched.push(jcId + ' (assignment "' + safeStr(a.contractorId) + '")');
            if (ar.changed) { a.contractorId = ar.id; asgChanged = true; }
          });
          if (asgChanged) { if (!dryRun) jc.getRange(sheetRow, 17).setValue(JSON.stringify(asg)); rowChanged = true; }
        }
        if (rowChanged) report.updated++; else report.alreadyOk++;
      }
    }
    var seen = {}, uniq = [];
    report.unmatched.forEach(function(x){ if (!seen[x]) { seen[x] = 1; uniq.push(x); } });
    report.unmatched = uniq;
    if (!dryRun) SpreadsheetApp.flush();
    Logger.log('migrateJobCardContractors ' + (dryRun ? '(dry-run) ' : '') +
               '— updated: ' + report.updated + ', ok: ' + report.alreadyOk + ', unmatched: ' + report.unmatched.length);
    return report;
  } catch(e) {
    return { success:false, error:e.message };
  }
}

// Menu wrappers. Previews are read-only (safe on LIVE); apply changes data.
function _jcMigReport(r, title) {
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, r.success
    ? ('Target sheet: ' + r.env + '\nUpdated: ' + r.updated + ' cards\nAlready OK: ' + r.alreadyOk +
       '\nUnresolved (need manual review): ' + r.unmatched.length +
       (r.unmatched.length ? ('\n\n' + r.unmatched.slice(0,30).join('\n')) : ''))
    : ('Error: ' + r.error), ui.ButtonSet.OK);
}
function migrateJCContractorsPreviewLive() { _jcMigReport(migrateJobCardContractors(true, 'LIVE'), 'Preview — LIVE (no changes made)'); }
function migrateJCContractorsPreviewDev()  { _jcMigReport(migrateJobCardContractors(true, 'DEV'),  'Preview — DEV (no changes made)'); }
function migrateJCContractorsApplyLive() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Apply on LIVE',
      'Back up the LIVE sheet first (File → Make a copy). This rewrites contractor references on LIVE job cards to CTR-IDs. Continue?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  _jcMigReport(migrateJobCardContractors(false, 'LIVE'), 'Applied — LIVE');
}
function migrateJCContractorsApplyDev() { _jcMigReport(migrateJobCardContractors(false, 'DEV'), 'Applied — DEV'); }

function ensureEnrollmentsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('CONTRACTOR_ENROLLMENTS');
  if (!ws) {
    ws = ss.insertSheet('CONTRACTOR_ENROLLMENTS');
    ws.getRange(1, 1, 1, 7).setValues([[
      'ENROLLMENT_ID', 'CONTRACTOR_ID', 'CONTRACTOR_NAME',
      'DEPARTMENT', 'ENROLLED_BY', 'ENROLLED_AT', 'STATUS'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function enrollContractor(data) {
  var VALID_DEPTS = [
    'Cutting', 'Preparation', 'Fitter',
    'Lasting/Pasting', 'Finishing/Packing', 'Dispatch'
  ];
  const lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var contractorId = safeStr(data.contractorId).trim();
      var department   = safeStr(data.department).trim();
      if (!contractorId) throw new Error('contractorId is required');
      if (VALID_DEPTS.indexOf(department) < 0) throw new Error('Invalid department: ' + department);

      var ws = ensureEnrollmentsSheet();
      var lastRow = ws.getLastRow();
      if (lastRow > 1) {
        var existing = ws.getRange(2, 1, lastRow - 1, 7).getValues();
        for (var i = 0; i < existing.length; i++) {
          if (safeStr(existing[i][1]).trim() === contractorId &&
              safeStr(existing[i][3]).trim() === department &&
              safeStr(existing[i][6]).trim().toUpperCase() === 'ACTIVE') {
            return { success: false, error: 'Already enrolled in this department' };
          }
        }
      }

      var contractorName = '';
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var mc = ss.getSheetByName('MASTER_CONTRACTORS');
      if (mc && mc.getLastRow() >= 4) {
        var mcRows = mc.getRange(4, 1, mc.getLastRow() - 3, 2).getValues();
        for (var j = 0; j < mcRows.length; j++) {
          if (safeStr(mcRows[j][0]).trim() === contractorId) {
            contractorName = safeStr(mcRows[j][1]).trim();
            break;
          }
        }
      }

      var dataRows = Math.max(0, lastRow - 1);
      var nextNum  = dataRows + 1;
      var year     = new Date().getFullYear();
      var enrollmentId = 'ENR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3, '0') : ('00' + nextNum).slice(-3));
      var user = getUserInfo();
      var now  = new Date().toISOString();
      ws.appendRow([enrollmentId, contractorId, contractorName, department, user.email, now, 'ACTIVE']);
      SpreadsheetApp.flush();
      try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
      return { success: true, enrollmentId: enrollmentId };
    } catch(e) { return { success: false, error: e.message }; }
  } finally { lock.releaseLock(); }
}

function unenrollContractor(enrollmentId) {
  const lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ws = ensureEnrollmentsSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'Enrollment not found' };
      var colA = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < colA.length; i++) {
        if (safeStr(colA[i][0]).trim() === safeStr(enrollmentId).trim()) {
          ws.getRange(i + 2, 7).setValue('INACTIVE');
          SpreadsheetApp.flush();
          try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
          return { success: true };
        }
      }
      return { success: false, error: 'Enrollment not found' };
    } catch(e) { return { success: false, error: e.message }; }
  } finally { lock.releaseLock(); }
}

function getEnrollments(filters, ss) {
  try {
    var ws = ensureEnrollmentsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows = ws.getRange(2, 1, lastRow - 1, 7).getValues();
    var result = [];
    rows.forEach(function(r) {
      var obj = {
        enrollmentId:   safeStr(r[0]).trim(),
        contractorId:   safeStr(r[1]).trim(),
        contractorName: safeStr(r[2]).trim(),
        department:     safeStr(r[3]).trim(),
        enrolledBy:     safeStr(r[4]).trim(),
        enrolledAt:     safeStr(r[5]).trim(),
        status:         safeStr(r[6]).trim()
      };
      if (!obj.enrollmentId) return;
      if (filters) {
        if (filters.contractorId && obj.contractorId !== filters.contractorId) return;
        if (filters.department   && obj.department   !== filters.department)   return;
        if (filters.status       && obj.status.toUpperCase() !== filters.status.toUpperCase()) return;
      }
      result.push(obj);
    });
    return result;
  } catch(e) { return { success: false, error: e.message }; }
}

function getContractorsScreenData() {
  try {
    var _cc = CacheService.getScriptCache();
    var _cv = _cc.get('contractorsScreen_' + CONFIG.ENV);
    if (_cv) return JSON.parse(_cv);
  } catch(ce) {}
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var result = { ctrs: null, contractors: [], enrollments: [] };
  try { result.ctrs = getContractorsData(ss); } catch(e) {}
  try { result.contractors = getContractors(ss); } catch(e) {}
  try { result.enrollments = getEnrollments({status:'ACTIVE'}, ss); } catch(e) {}
  try {
    CacheService.getScriptCache()
      .put('contractorsScreen_' + CONFIG.ENV, JSON.stringify(result), 300);
  } catch(ce) {}
  return result;
}

/* ══════════════════════════════════════════════════════════════════════════
   PHASE 8.1 — Contractor delete (soft) / reactivate
   Delete = mark status INACTIVE (reversible). getContractors() already hides
   INACTIVE rows; getContractorsData() still returns them so the Contractors
   screen can show a greyed card with a Reactivate action.
   ═══════════════════════════════════════════════════════════════════════════ */

// Locate a contractor row by CTR-ID. Returns 1-based sheet row or -1.
function _findContractorRow(mc, ctrId) {
  if (!mc || mc.getLastRow() < 4) return -1;
  var ids = mc.getRange(4, 1, mc.getLastRow() - 3, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (safeStr(ids[i][0]).trim() === ctrId) return i + 4;
  }
  return -1;
}

function _contractorNameById(ss, ctrId) {
  try {
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() >= 4) {
      var rows = mc.getRange(4, 1, mc.getLastRow() - 3, 2).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (safeStr(rows[i][0]).trim() === ctrId) return safeStr(rows[i][1]).trim();
      }
    }
  } catch(e) {}
  return '';
}

// Soft-delete: set status (col D) to INACTIVE. Reversible via reactivateContractor.
function deleteContractor(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var ctrId = safeStr(payload && payload.ctrId).trim();
  if (!ctrId) return { success: false, error: 'ctrId is required' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc) return { success: false, error: 'MASTER_CONTRACTORS sheet not found' };
    var row = _findContractorRow(mc, ctrId);
    if (row < 0) return { success: false, error: 'Contractor not found' };
    mc.getRange(row, 4).setValue('INACTIVE');
    // Also deactivate the contractor's department enrollments so they drop off
    // the store's issue dropdowns.
    try {
      var ws = ensureEnrollmentsSheet();
      if (ws.getLastRow() > 1) {
        var er = ws.getRange(2, 1, ws.getLastRow() - 1, 7).getValues();
        for (var i = 0; i < er.length; i++) {
          if (safeStr(er[i][1]).trim() === ctrId &&
              safeStr(er[i][6]).trim().toUpperCase() === 'ACTIVE') {
            ws.getRange(i + 2, 7).setValue('INACTIVE');
          }
        }
      }
    } catch(enrErr) { Logger.log('deleteContractor enrollment cleanup: ' + enrErr.message); }
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function reactivateContractor(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var ctrId = safeStr(payload && payload.ctrId).trim();
  if (!ctrId) return { success: false, error: 'ctrId is required' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (!mc) return { success: false, error: 'MASTER_CONTRACTORS sheet not found' };
    var row = _findContractorRow(mc, ctrId);
    if (row < 0) return { success: false, error: 'Contractor not found' };
    mc.getRange(row, 4).setValue('ACTIVE');
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

/* ══════════════════════════════════════════════════════════════════════════
   PHASE 8.2 — Contractor profile + document storage (Google Drive)
   Profile fields live in CONTRACTOR_PROFILE (one row per CTR-ID).
   Documents (Aadhaar, passbook, agreement, etc.) are uploaded into a per-
   contractor Drive folder; the sheet CONTRACTOR_DOCS stores links only.
   Requires the drive.file OAuth scope (added to appsscript.json).
   ═══════════════════════════════════════════════════════════════════════════ */

function ensureContractorProfileSheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('CONTRACTOR_PROFILE');
  if (!ws) {
    ws = ss.insertSheet('CONTRACTOR_PROFILE');
    ws.getRange(1, 1, 1, 11).setValues([[
      'CONTRACTOR_ID', 'ADDRESS', 'AADHAAR_NO', 'PAN_NO', 'BANK_NAME',
      'BANK_ACCOUNT', 'IFSC', 'ALT_PHONE', 'NOTES', 'UPDATED_BY', 'UPDATED_AT'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function ensureContractorDocsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('CONTRACTOR_DOCS');
  if (!ws) {
    ws = ss.insertSheet('CONTRACTOR_DOCS');
    ws.getRange(1, 1, 1, 9).setValues([[
      'DOC_ID', 'CONTRACTOR_ID', 'DOC_TYPE', 'FILE_NAME', 'FILE_URL',
      'FILE_ID', 'UPLOADED_BY', 'UPLOADED_AT', 'STATUS'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

// Root Drive folder for all contractor docs, kept separate per environment.
function _contractorDocsRootFolder() {
  var rootName = 'Adees FactoryOS — Contractor Documents (' + CONFIG.ENV + ')';
  var it = DriveApp.getFoldersByName(rootName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(rootName);
}

// Per-contractor subfolder, e.g. "CTR-004 — Raju Kumar".
function _contractorFolder(ctrId, name) {
  var root  = _contractorDocsRootFolder();
  var fname = ctrId + (name ? ' — ' + name : '');
  var it = root.getFoldersByName(fname);
  if (it.hasNext()) return it.next();
  return root.createFolder(fname);
}

function getContractorProfile(contractorId) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var ctrId = safeStr(contractorId).trim();
  if (!ctrId) return { success: false, error: 'contractorId is required' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var profile = {
      contractorId: ctrId, address: '', aadhaarNo: '', panNo: '',
      bankName: '', bankAccount: '', ifsc: '', altPhone: '', notes: '',
      updatedBy: '', updatedAt: ''
    };
    var ps = ensureContractorProfileSheet(ss);
    if (ps.getLastRow() > 1) {
      var pr = ps.getRange(2, 1, ps.getLastRow() - 1, 11).getValues();
      for (var i = 0; i < pr.length; i++) {
        if (safeStr(pr[i][0]).trim() === ctrId) {
          profile.address     = safeStr(pr[i][1]);
          profile.aadhaarNo   = safeStr(pr[i][2]);
          profile.panNo       = safeStr(pr[i][3]);
          profile.bankName    = safeStr(pr[i][4]);
          profile.bankAccount = safeStr(pr[i][5]);
          profile.ifsc        = safeStr(pr[i][6]);
          profile.altPhone    = safeStr(pr[i][7]);
          profile.notes       = safeStr(pr[i][8]);
          profile.updatedBy   = safeStr(pr[i][9]);
          profile.updatedAt   = safeStr(pr[i][10]);
          break;
        }
      }
    }
    var docs = [];
    var ds = ensureContractorDocsSheet(ss);
    if (ds.getLastRow() > 1) {
      ds.getRange(2, 1, ds.getLastRow() - 1, 9).getValues().forEach(function(r) {
        if (safeStr(r[1]).trim() !== ctrId) return;
        if (safeStr(r[8]).trim().toUpperCase() === 'DELETED') return;
        docs.push({
          docId:      safeStr(r[0]),
          docType:    safeStr(r[2]),
          fileName:   safeStr(r[3]),
          url:        safeStr(r[4]),
          uploadedBy: safeStr(r[6]),
          uploadedAt: safeStr(r[7])
        });
      });
    }
    return { success: true, profile: profile, docs: docs };
  } catch(e) { return { success: false, error: e.message }; }
}

function saveContractorProfile(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var ctrId = safeStr(payload && payload.contractorId).trim();
  if (!ctrId) return { success: false, error: 'contractorId is required' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ps = ensureContractorProfileSheet(ss);
    var now = new Date().toISOString();
    var rowVals = [
      ctrId, safeStr(payload.address), safeStr(payload.aadhaarNo), safeStr(payload.panNo),
      safeStr(payload.bankName), safeStr(payload.bankAccount), safeStr(payload.ifsc),
      safeStr(payload.altPhone), safeStr(payload.notes), user.email, now
    ];
    var found = -1;
    if (ps.getLastRow() > 1) {
      var ids = ps.getRange(2, 1, ps.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (safeStr(ids[i][0]).trim() === ctrId) { found = i + 2; break; }
      }
    }
    if (found > 0) ps.getRange(found, 1, 1, 11).setValues([rowVals]);
    else ps.appendRow(rowVals);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Upload one document. payload: { contractorId, docType, fileName, mimeType, dataBase64 }
function uploadContractorDoc(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  try {
    var ctrId    = safeStr(payload && payload.contractorId).trim();
    var docType  = safeStr(payload && payload.docType).trim() || 'Other';
    var fileName = safeStr(payload && payload.fileName).trim() || 'document';
    var mime     = safeStr(payload && payload.mimeType).trim() || 'application/octet-stream';
    var b64      = safeStr(payload && payload.dataBase64);
    if (!ctrId) return { success: false, error: 'contractorId is required' };
    if (!b64)   return { success: false, error: 'No file data received' };

    var ss   = SpreadsheetApp.openById(SHEET_ID);
    var name = _contractorNameById(ss, ctrId);
    var bytes = Utilities.base64Decode(b64);
    var blob  = Utilities.newBlob(bytes, mime, fileName);
    var folder = _contractorFolder(ctrId, name);
    var file   = folder.createFile(blob);
    file.setName('[' + docType + '] ' + fileName);
    try { file.setDescription('Contractor ' + ctrId + ' (' + name + ') — ' + docType); } catch(de) {}

    var ws = ensureContractorDocsSheet(ss);
    var seq = ws.getLastRow(); // header counts as 1 → first doc becomes seq 1
    var seqStr = String(seq); while (seqStr.length < 3) seqStr = '0' + seqStr;
    var docId  = 'DOC-' + new Date().getFullYear() + '-' + seqStr;
    var nowIso = new Date().toISOString();
    ws.appendRow([docId, ctrId, docType, fileName, file.getUrl(), file.getId(), user.email, nowIso, 'ACTIVE']);
    SpreadsheetApp.flush();
    try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
    return { success: true, doc: {
      docId: docId, contractorId: ctrId, docType: docType, fileName: fileName,
      url: file.getUrl(), uploadedBy: user.email, uploadedAt: nowIso
    }};
  } catch(e) { return { success: false, error: e.message }; }
}

// Soft-delete a document row and move the Drive file to trash.
function deleteContractorDoc(payload) {
  var user = getUserInfo();
  if (user.role !== 'accounts' && user.role !== 'admin')
    return { success: false, error: 'Not authorised' };
  var docId = safeStr(payload && payload.docId).trim();
  if (!docId) return { success: false, error: 'docId is required' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ensureContractorDocsSheet(ss);
    if (ws.getLastRow() < 2) return { success: false, error: 'Document not found' };
    var rows = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (safeStr(rows[i][0]).trim() === docId) {
        ws.getRange(i + 2, 9).setValue('DELETED');
        try { DriveApp.getFileById(safeStr(rows[i][5]).trim()).setTrashed(true); } catch(fe) {}
        SpreadsheetApp.flush();
        try { CacheService.getScriptCache().remove('contractorsScreen_' + CONFIG.ENV); } catch(ce) {}
        return { success: true };
      }
    }
    return { success: false, error: 'Document not found' };
  } catch(e) { return { success: false, error: e.message }; }
}

// S.9: set of every CTR-ID in MASTER_CONTRACTORS (any status). Used by job-card
// writers to refuse a contractorId that isn't a real contractor (names, typos,
// stale clients).
function _validContractorIds_(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var idSet = {};
  try {
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() > 3) {
      mc.getRange(4, 1, mc.getLastRow()-3, 1).getValues().forEach(function(r){
        var id = safeStr(r[0]).trim();
        if (id) idSet[id] = true;
      });
    }
  } catch(e) { Logger.log('_validContractorIds_ error: ' + e.message); }
  return idSet;
}

// S.9 read-only diagnostic. Run from the editor: Run > auditContractorNames.
// Reports duplicate contractor names (grouped, with statuses), plus JOB_CARDS
// contractor values that aren't valid CTR-IDs. Writes nothing.
function auditContractorNames(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  var report = { sheet: ss.getName(), contractors:0, duplicateNames:[], duplicateActiveNames:[], invalidJobCardContractors:[] };
  try {
    var byName = {};
    var mc = ss.getSheetByName('MASTER_CONTRACTORS');
    if (mc && mc.getLastRow() > 3) {
      mc.getRange(4, 1, mc.getLastRow()-3, 4).getValues().forEach(function(r){
        var id = safeStr(r[0]).trim(), nm = safeStr(r[1]).trim();
        var st = safeStr(r[3]).trim().toUpperCase() || 'ACTIVE';
        if (!nm) return;
        report.contractors++;
        var key = nm.toLowerCase();
        if (!byName[key]) byName[key] = [];
        byName[key].push(id + ':' + st);
      });
    }
    Object.keys(byName).forEach(function(k){
      if (byName[k].length > 1) {
        report.duplicateNames.push(k + ' → ' + byName[k].join(', '));
        var actives = byName[k].filter(function(e){ return e.indexOf(':INACTIVE') < 0; });
        if (actives.length > 1) report.duplicateActiveNames.push(k + ' → ' + actives.join(', '));
      }
    });
    var idSet = _validContractorIds_(ss);
    var jc = ss.getSheetByName('JOB_CARDS');
    if (jc && jc.getLastRow() > 1) {
      jc.getRange(2, 1, jc.getLastRow()-1, 6).getValues().forEach(function(r){
        var jcId = safeStr(r[0]).trim(), cid = safeStr(r[5]).trim();
        if (jcId && cid && !idSet[cid]) report.invalidJobCardContractors.push(jcId + ' → "' + cid + '"');
      });
    }
  } catch(e) { report.error = e.message; }
  Logger.log('auditContractorNames → ' + JSON.stringify(report, null, 2));
  return report;
}
function auditContractorNamesLive() { return auditContractorNames(SpreadsheetApp.openById(CONFIG.LIVE_SHEET_ID)); }
