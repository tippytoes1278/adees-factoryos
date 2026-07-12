function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Factory OS Admin')
    .addItem('Run CTR-ID Migration', 'assignContractorIds')
    .addItem('Backfill Order Size Columns', 'backfillOrderSizesMenu')
    .addItem('Migrate WIP_ENTRIES Schema', 'migrateWipEntriesMenu')
    .addToUi();
}

function saveWIP(rowNum, produced) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var wr = ss.getSheetByName('WIP_RECONCILIATION');
    wr.getRange('D'+rowNum).setValue(produced);
    SpreadsheetApp.flush();
    return { success:true,
      status: safeStr(wr.getRange('G'+rowNum).getValue()),
      diff:   safeStr(wr.getRange('F'+rowNum).getValue()) };
  } catch(e) { return { success:false, error:e.message }; }
}

// ── WIP ENTRIES — Phase 5.1 ───────────────────────────────────────────────────

function ensureWipEntriesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('WIP_ENTRIES');
  var NEW_HEADERS = ['WIP_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT','ENTRY_TYPE','PAIRS','SUBMITTED_BY','SUBMITTED_AT','PERIOD_ID','STATUS','NOTES','CONTRACTORS','JOB_CARD_REF'];
  if (!ws) {
    ws = ss.insertSheet('WIP_ENTRIES');
    ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
    ws.setFrozenRows(1);
  } else {
    if (safeStr(ws.getRange(1, 1).getValue()) !== 'WIP_ID') {
      ws.clearContents();
      ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
      ws.setFrozenRows(1);
    }
  }
  return ws;
}

function saveWipEntry(data, status) {
  var STORE_MOVEMENT_MAP = {
    'Upper Store':              ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store':  ['Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':           ['Dispatch IN','Dispatch OUT']
  };
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var orderRef    = safeStr(data.orderRef).trim();
      var workOrder   = safeStr(data.workOrder   || '').trim();
      var store       = safeStr(data.store).trim();
      var movement    = safeStr(data.movement).trim();
      var pairs       = safeNum(data.pairs);
      var periodId    = safeStr(data.periodId).trim();
      var notes       = safeStr(data.notes       || '').trim();
      var contractors = Array.isArray(data.contractors) ? data.contractors : [];
      var jobCardRef  = safeStr(data.jobCardRef  || '').trim();

      var _wipUser = getUserInfo();
      if (_wipUser.role !== 'store' && _wipUser.role !== 'admin')
        return { success:false, error:'Not authorised' };

      if (!orderRef) throw new Error('orderRef is required');
      if (!STORE_MOVEMENT_MAP[store]) throw new Error('Invalid store: ' + store);
      if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0) throw new Error('Invalid movement for store: ' + movement);
      if (!pairs || pairs <= 0 || Math.floor(pairs) !== pairs) throw new Error('pairs must be a positive integer');
      if (!periodId) throw new Error('periodId is required');

      status = status || 'SUBMITTED';
      var entryType = movement.slice(-2) === 'IN' ? 'IN' : 'OUT';

      var ws = ensureWipEntriesSheet();
      var dataRows = Math.max(0, ws.getLastRow() - 1);
      var nextNum  = dataRows + 1;
      var year     = new Date().getFullYear();
      var wipId    = 'WIP-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));

      var user          = getUserInfo();
      var now           = new Date().toISOString();
      var contractorsStr = contractors.join(',');

      ws.appendRow([wipId, orderRef, workOrder, store, movement, entryType, pairs, user.email, now, periodId, status, notes, contractorsStr, jobCardRef]);
      SpreadsheetApp.flush();
      return { success: true, wipId: wipId };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function getWipEntries(filters, ss) {
  try {
    var ws = ensureWipEntriesSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows = ws.getRange(2, 1, lastRow - 1, 14).getValues();
    var entries = rows.map(function(r) {
      return {
        wipId:       safeStr(r[0]),
        orderRef:    safeStr(r[1]),
        workOrder:   safeStr(r[2]),
        store:       safeStr(r[3]),
        movement:    safeStr(r[4]),
        entryType:   safeStr(r[5]),
        pairs:       safeNum(r[6]),
        submittedBy: safeStr(r[7]),
        submittedAt: safeStr(r[8]),
        periodId:    safeStr(r[9]),
        status:      safeStr(r[10]),
        notes:       safeStr(r[11]),
        contractors: safeStr(r[12]),
        jobCardRef:  safeStr(r[13])
      };
    });
    if (filters) {
      if (filters.orderRef)   entries = entries.filter(function(e){ return e.orderRef   === filters.orderRef; });
      if (filters.store)      entries = entries.filter(function(e){ return e.store      === filters.store; });
      if (filters.movement)   entries = entries.filter(function(e){ return e.movement   === filters.movement; });
      if (filters.entryType)  entries = entries.filter(function(e){ return e.entryType  === filters.entryType; });
      if (filters.periodId)   entries = entries.filter(function(e){ return e.periodId   === filters.periodId; });
      if (filters.status)     entries = entries.filter(function(e){ return e.status     === filters.status; });
      if (filters.jobCardRef) entries = entries.filter(function(e){ return e.jobCardRef === filters.jobCardRef; });
    }
    return entries;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function voidWipEntry(wipId) {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ws = ensureWipEntriesSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'WIP entry not found' };
      var ids = ws.getRange(2, 1, lastRow - 1, 1).getValues();
      var targetRow = -1;
      for (var i = 0; i < ids.length; i++) {
        if (safeStr(ids[i][0]) === wipId) { targetRow = i + 2; break; }
      }
      if (targetRow < 0) return { success: false, error: 'WIP entry not found' };
      var status = safeStr(ws.getRange(targetRow, 11).getValue());
      if (status === 'LINKED') return { success: false, error: 'Cannot void a linked entry' };
      ws.getRange(targetRow, 11).setValue('VOID');
      SpreadsheetApp.flush();
      return { success: true };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

// ── WIP GRID — Phase 5.2f ────────────────────────────────────────────────────

function getWipGrid() {
  var ALL_MOVEMENTS = [
    'Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT',
    'Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT',
    'Dispatch IN','Dispatch OUT'
  ];
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var oi = ss.getSheetByName('ORDER_INDEX');
    var orders = [];
    if (oi && oi.getLastRow() >= 4) {
      var oiRows = oi.getRange(4, 1, oi.getLastRow() - 3, 9).getValues();
      oiRows.forEach(function(r) {
        var orderId  = safeStr(r[0]).trim();
        var artSheet = safeStr(r[1]).trim();
        var article  = safeStr(r[2]).trim();
        var color    = safeStr(r[3]).trim();
        var customer = safeStr(r[4]).trim();
        var lotSize  = safeNum(r[8]);
        if (!orderId || lotSize <= 0) return;
        orders.push({ orderId: orderId, artSheet: artSheet, article: article, color: color, customer: customer, lotSize: lotSize });
      });
    }
    var ws = ensureWipEntriesSheet();
    var today = new Date();
    var mm = String(today.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(today.getDate());      if (dd.length < 2) dd = '0' + dd;
    var todayStr = today.getFullYear() + '-' + mm + '-' + dd;
    var map = {};
    if (ws.getLastRow() >= 2) {
      var wipRows = ws.getRange(2, 1, ws.getLastRow() - 1, 14).getValues();
      wipRows.forEach(function(r) {
        var orderRef    = safeStr(r[1]).trim();  // WIP_ENTRIES col B = ART sheet name
        var movement    = safeStr(r[4]).trim();
        var pairs       = safeNum(r[6]);
        var submittedAt = safeStr(r[8]).trim();
        var status      = safeStr(r[10]).trim();
        if (!orderRef || !movement || status === 'VOID') return;
        // Match WIP entry to order via artSheet (col B of ORDER_INDEX)
        var matched = null;
        for (var oi2 = 0; oi2 < orders.length; oi2++) {
          if (orders[oi2].artSheet === orderRef) { matched = orders[oi2]; break; }
        }
        if (!matched) return;
        var key = matched.orderId;
        if (!map[key]) map[key] = {};
        if (!map[key][movement]) map[key][movement] = { today: 0, total: 0 };
        map[key][movement].total += pairs;
        if (submittedAt && submittedAt.indexOf(todayStr) === 0) map[key][movement].today += pairs;
      });
    }
    var entries = [];
    Object.keys(map).forEach(function(orderId) {
      ALL_MOVEMENTS.forEach(function(mv) {
        var m = map[orderId][mv] || { today: 0, total: 0 };
        entries.push({ orderRef: orderId, movement: mv, pairsToday: m.today, pairsTotal: m.total });
      });
    });
    return { orders: orders, entries: entries, movements: ALL_MOVEMENTS };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function submitWipGrid(gridEntries) {
  var STORE_MOVEMENT_MAP = {
    'Upper Store':              ['Cutting IN','Cutting OUT','Preparation IN','Preparation OUT','Fitter IN','Fitter OUT'],
    'Lasting & Packing Store':  ['Upper IN','Lasting IN','Lasting OUT','Packing IN','Packing OUT'],
    'Dispatch Store':           ['Dispatch IN','Dispatch OUT']
  };
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      if (!Array.isArray(gridEntries) || !gridEntries.length) throw new Error('No entries provided');
      var ws = ensureWipEntriesSheet();
      var today = new Date();
      var mm = String(today.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
      var dd = String(today.getDate());      if (dd.length < 2) dd = '0' + dd;
      var defaultPeriodId = 'GRID-' + today.getFullYear() + '-' + mm + '-' + dd;
      var user = getUserInfo();
      var now  = new Date().toISOString();
      var toInsert = [];
      gridEntries.forEach(function(entry) {
        var orderRef  = safeStr(entry.orderRef).trim();
        var workOrder = safeStr(entry.workOrder  || '').trim();
        var store     = safeStr(entry.store).trim();
        var movement  = safeStr(entry.movement).trim();
        var pairs     = safeNum(entry.pairs);
        var periodId  = safeStr(entry.periodId   || defaultPeriodId).trim();
        if (!orderRef) throw new Error('orderRef required');
        if (!STORE_MOVEMENT_MAP[store]) throw new Error('Invalid store: ' + store);
        if (STORE_MOVEMENT_MAP[store].indexOf(movement) < 0) throw new Error('Invalid movement for store: ' + movement);
        if (!pairs || pairs <= 0) throw new Error('pairs must be positive for: ' + orderRef);
        var entryType = movement.slice(-2) === 'IN' ? 'IN' : 'OUT';
        toInsert.push([orderRef, workOrder, store, movement, entryType, pairs, periodId]);
      });
      var dataRows = Math.max(0, ws.getLastRow() - 1);
      var saved = 0;
      toInsert.forEach(function(row, i) {
        var nextNum = dataRows + i + 1;
        var seq = String(nextNum); while (seq.length < 3) seq = '0' + seq;
        var wipId = 'WIP-' + today.getFullYear() + '-' + seq;
        ws.appendRow([wipId, row[0], row[1], row[2], row[3], row[4], row[5], user.email, now, row[6], 'PENDING', 'Grid entry', '', '']);
        saved++;
      });
      SpreadsheetApp.flush();
      return { success: true, saved: saved };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

// ── WIP_ENTRIES SCHEMA MIGRATION — Phase 5.3 ─────────────────────────────────

function migrateWipEntries() {
  var STAGE_MAP = {
    'Cutting':             { store: 'Upper Store',             movement: 'Cutting IN' },
    'Preparation':         { store: 'Upper Store',             movement: 'Preparation IN' },
    'Upper Making':        { store: 'Upper Store',             movement: 'Fitter IN' },
    'Lasting & Pasting':   { store: 'Lasting & Packing Store', movement: 'Lasting IN' },
    'Finishing & Packing': { store: 'Lasting & Packing Store', movement: 'Packing IN' },
    'Dispatch':            { store: 'Dispatch Store',          movement: 'Dispatch IN' }
  };
  var NEW_HEADERS = ['WIP_ID','ORDER_REF','WORK_ORDER','STORE','MOVEMENT','ENTRY_TYPE','PAIRS','SUBMITTED_BY','SUBMITTED_AT','PERIOD_ID','STATUS','NOTES','CONTRACTORS','JOB_CARD_REF'];
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ss = SpreadsheetApp.openById(SHEET_ID);
      var ws = ss.getSheetByName('WIP_ENTRIES');
      if (!ws) {
        ensureWipEntriesSheet();
        return { success: true, message: 'Created WIP_ENTRIES with new schema. No migration needed.' };
      }
      var header3 = safeStr(ws.getRange(1, 3).getValue());
      if (header3 === 'WORK_ORDER') {
        return { success: true, message: 'Already on new schema. No migration needed.' };
      }
      var header1 = safeStr(ws.getRange(1, 1).getValue());
      if (header1 !== 'WIP_ID') {
        ws.clearContents();
        ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
        ws.setFrozenRows(1);
        SpreadsheetApp.flush();
        return { success: true, message: 'Reinitialized blank/corrupt sheet with new schema.' };
      }
      var lastRow = ws.getLastRow();
      var migratedRows = 0;
      var skippedRows  = 0;
      var newRows      = [];
      if (lastRow >= 2) {
        var oldData = ws.getRange(2, 1, lastRow - 1, 12).getValues();
        oldData.forEach(function(r) {
          var wipId       = safeStr(r[0]).trim();
          var orderRef    = safeStr(r[1]).trim();
          var workOrder   = safeStr(r[2]).trim();
          var stage       = safeStr(r[3]).trim();
          var pairs       = safeNum(r[5]);
          var submittedBy = safeStr(r[6]).trim();
          var submittedAt = safeStr(r[7]).trim();
          var periodId    = safeStr(r[8]).trim();
          var status      = safeStr(r[9]).trim();
          var notes       = safeStr(r[10]).trim();
          var contractors = safeStr(r[11]).trim();
          if (!wipId) { skippedRows++; return; }
          var mapped    = STAGE_MAP[stage] || { store: 'Upper Store', movement: 'Cutting IN' };
          newRows.push([wipId, orderRef, workOrder, mapped.store, mapped.movement, 'IN', pairs, submittedBy, submittedAt, periodId, status, notes, contractors, '']);
          migratedRows++;
        });
      }
      ws.clearContents();
      ws.getRange(1, 1, 1, 14).setValues([NEW_HEADERS]);
      if (newRows.length) ws.getRange(2, 1, newRows.length, 14).setValues(newRows);
      ws.setFrozenRows(1);
      SpreadsheetApp.flush();
      return { success: true, migrated: migratedRows, skipped: skippedRows };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function migrateWipEntriesMenu() {
  var result = migrateWipEntries();
  var ui = SpreadsheetApp.getUi();
  var msg = result.success
    ? (result.message || ('Migrated: ' + result.migrated + ' rows. Skipped: ' + result.skipped + ' blank rows.'))
    : ('Error: ' + result.error);
  ui.alert('Migrate WIP_ENTRIES Schema', msg, ui.ButtonSet.OK);
}

// ── DAILY REPORTS — Phase 5.3a ────────────────────────────────────────────────

function ensureDailyReportsSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = ss.getSheetByName('DAILY_REPORTS');
  if (!ws) {
    ws = ss.insertSheet('DAILY_REPORTS');
    ws.getRange(1, 1, 1, 7).setValues([[
      'REPORT_ID', 'DATE', 'SUBMITTED_BY', 'SUBMITTED_AT',
      'ENTRY_COUNT', 'TOTAL_PAIRS', 'STATUS'
    ]]);
    ws.setFrozenRows(1);
  }
  return ws;
}

function submitDay() {
  var lock = LockService.getPublicLock();
  try {
    lock.waitLock(10000);
    try {
      var ss        = SpreadsheetApp.openById(SHEET_ID);
      var email     = Session.getActiveUser().getEmail();
      var todayDate = new Date().toISOString().slice(0, 10);
      var now       = new Date().toISOString();

      var ws      = ensureWipEntriesSheet();
      var lastRow = ws.getLastRow();
      if (lastRow < 2) return { success: false, error: 'No draft entries to submit' };

      var rows        = ws.getRange(2, 1, lastRow - 1, 14).getValues();
      var matchedIdxs = [];
      rows.forEach(function(r, i) {
        var status      = safeStr(r[10]).trim();
        var submittedBy = safeStr(r[7]).trim();
        var submittedAt = safeStr(r[8]).trim();
        if (status === 'DRAFT' && submittedBy === email && submittedAt.slice(0, 10) === todayDate) {
          matchedIdxs.push(i);
        }
      });

      if (!matchedIdxs.length) return { success: false, error: 'No draft entries to submit' };

      matchedIdxs.forEach(function(i) {
        ws.getRange(i + 2, 11).setValue('SUBMITTED');
      });

      var entryCount = matchedIdxs.length;
      var totalPairs = 0;
      matchedIdxs.forEach(function(i) { totalPairs += safeNum(rows[i][6]); });

      var dr       = ensureDailyReportsSheet();
      var drRows   = Math.max(0, dr.getLastRow() - 1);
      var nextNum  = drRows + 1;
      var year     = new Date().getFullYear();
      var reportId = 'DR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));

      dr.appendRow([reportId, todayDate, email, now, entryCount, totalPairs, 'SUBMITTED']);
      SpreadsheetApp.flush();
      return { success: true, reportId: reportId, entryCount: entryCount, totalPairs: totalPairs, date: todayDate };
    } catch(e) {
      return { success: false, error: e.message };
    }
  } finally {
    lock.releaseLock();
  }
}

function generateDailyReport() {
  try {
    var email     = Session.getActiveUser().getEmail();
    var todayDate = new Date().toISOString().slice(0, 10);
    var now       = new Date().toISOString();

    var wipWs   = ensureWipEntriesSheet();
    var lastRow = wipWs.getLastRow();
    var entryCount = 0, totalPairs = 0;
    if (lastRow >= 2) {
      var rows = wipWs.getRange(2, 1, lastRow - 1, 11).getValues();
      rows.forEach(function(r) {
        if (safeStr(r[10]).trim() === 'SUBMITTED' &&
            safeStr(r[7]).trim() === email &&
            safeStr(r[8]).trim().slice(0, 10) === todayDate) {
          entryCount++;
          totalPairs += safeNum(r[6]);
        }
      });
    }
    if (entryCount === 0) return;

    var dr      = ensureDailyReportsSheet();
    var drLast  = dr.getLastRow();
    var foundRow = -1;
    if (drLast >= 2) {
      var drRows = dr.getRange(2, 1, drLast - 1, 4).getValues();
      for (var i = 0; i < drRows.length; i++) {
        if (safeStr(drRows[i][1]).trim() === todayDate &&
            safeStr(drRows[i][2]).trim() === email) {
          foundRow = i + 2; break;
        }
      }
    }
    if (foundRow > 0) {
      dr.getRange(foundRow, 4).setValue(now);
      dr.getRange(foundRow, 5).setValue(entryCount);
      dr.getRange(foundRow, 6).setValue(totalPairs);
    } else {
      var drRowCount = Math.max(0, dr.getLastRow() - 1);
      var nextNum    = drRowCount + 1;
      var year       = new Date().getFullYear();
      var reportId   = 'DR-' + year + '-' + (String(nextNum).padStart ? String(nextNum).padStart(3,'0') : ('00'+nextNum).slice(-3));
      dr.appendRow([reportId, todayDate, email, now, entryCount, totalPairs, 'AUTO']);
    }
    SpreadsheetApp.flush();
  } catch(e) {}
}

function getDailyReports(filters) {
  try {
    var ws      = ensureDailyReportsSheet();
    var lastRow = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows    = ws.getRange(2, 1, lastRow - 1, 7).getValues();
    var result  = [];
    rows.forEach(function(r) {
      var reportId    = safeStr(r[0]).trim();
      if (!reportId) return;
      var date        = safeStr(r[1]).trim();
      var submittedBy = safeStr(r[2]).trim();
      var submittedAt = safeStr(r[3]).trim();
      var entryCount  = safeNum(r[4]);
      var totalPairs  = safeNum(r[5]);
      var status      = safeStr(r[6]).trim();
      if (filters) {
        if (filters.date        && date        !== filters.date)        return;
        if (filters.submittedBy && submittedBy !== filters.submittedBy) return;
      }
      result.push({ reportId: reportId, date: date, submittedBy: submittedBy, submittedAt: submittedAt, entryCount: entryCount, totalPairs: totalPairs, status: status });
    });
    result.sort(function(a, b) { return b.date < a.date ? -1 : b.date > a.date ? 1 : 0; });
    if (filters && filters.limit) result = result.slice(0, filters.limit);
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getTodaysDrafts() {
  try {
    var ss        = SpreadsheetApp.openById(SHEET_ID);
    var todayDate = new Date().toISOString().slice(0, 10);
    var ws        = ensureWipEntriesSheet();
    var lastRow   = ws.getLastRow();
    if (lastRow < 2) return [];
    var rows      = ws.getRange(2, 1, lastRow - 1, 14).getValues();
    var result    = [];
    rows.forEach(function(r) {
      var status      = safeStr(r[10]).trim();
      var submittedAt = safeStr(r[8]).trim();
      if (status !== 'DRAFT' || submittedAt.slice(0, 10) !== todayDate) return;
      result.push({
        wipId:       safeStr(r[0]),
        orderRef:    safeStr(r[1]),
        workOrder:   safeStr(r[2]),
        store:       safeStr(r[3]),
        movement:    safeStr(r[4]),
        entryType:   safeStr(r[5]),
        pairs:       safeNum(r[6]),
        submittedBy: safeStr(r[7]),
        submittedAt: submittedAt,
        periodId:    safeStr(r[9]),
        status:      status,
        notes:       safeStr(r[11]),
        contractors: safeStr(r[12]),
        jobCardRef:  safeStr(r[13])
      });
    });
    return result;
  } catch(e) {
    return { success: false, error: e.message };
  }
}
