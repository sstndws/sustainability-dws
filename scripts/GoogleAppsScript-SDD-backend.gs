/**
 * Google Apps Script — backend SDD / Mill / TTP / Grievance
 * Salin isi file ini ke project Apps Script Anda (gabungkan dengan file lain bila perlu).
 *
 * Perubahan vs versi yang Anda kirim:
 * - Mill sheet: payload dari dashboard memakai QUARTER / YEAR; sheet sering "Quarter"/"Year".
 *   resolveMillQuarterYearKeys_ dipanggil di addRow/updateRow untuk sheet mill.
 *   getAll('mill') memanggil mirrorMillQuarterYearOnRead_ agar JSON juga berisi QUARTER/YEAR.
 * - doPost: action "delete" — jika body.rows berisi array nomor baris, pakai bulkDelete
 *   (urut turun di dalam bulkDelete); jika tidak, fallback deleteRow(body.row).
 *   Ini selaras dengan klien sustain-dashboard yang mengirim { action, sheet, row, rows }.
 * - bulkDelete: deduplikasi nomor baris sebelum hapus (hindari deleteRow ganda pada nomor sama).
 */

const SHEETS = {
  mill: 'Mill Onboarding Profile',
  ttp: 'Monitoring TTP/TTM',
  grievance: 'Grievance Monitoring',
  sdd: 'SDD Data'
};

const TTP_HEADERS = [
  'NO', 'COMPANY CODE', 'GROUP NAME', 'COMPANY NAME', 'MILL NAME', 'UML ID',
  'FFB SUPPLIER GROUP NAME', 'FFB SUPPLIER NAME', 'CATEGORY', 'LAT', 'LONG',
  'VILLAGE ID', 'VILLAGE', 'SUBDISTRICT', 'DISTRICT', 'PROVINCE',
  'CONCESION AREA', 'PLANTED AREA', 'NUMBER OD SMALLHOLDERS', 'TAHUN TANAM',
  'LEGALITAS', 'ISPO (Y/N)', 'RSPO (Y/N)', 'ISCC (Y/N)',
  'FFB SUPPLY to MILL (TON)', 'CONVERSION FFB to PK (5%)', 'PK SUPPLY to KCP',
  'CONVERSION FFB to CPO (20%)', 'CPO SUPPLY to REFINERY',
  '% PK TRACEABLE', '% CPO TRACEABLE',
  'Total PK % Traceable', 'Total CPO % Traceable',
  'MSD', 'PK Traceable Volume', 'CPO Traceable Volume'
];

// ═══════════════════════════════════════════════════════════
//  ENTRY POINTS
// ═══════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const action   = (e && e.parameter && e.parameter.action)  || '';
    const sheetKey = (e && e.parameter && e.parameter.sheet)   || '';

    if (action === 'getAll') {
      if (sheetKey === 'ttp') ensureTtpHeaders_();
      return respond(getData(sheetKey));
    }
    if (action === 'getByMillId') return respond(getByMillId(e.parameter.millId));
    if (action === 'ping')        return respond({ success: true, message: 'Apps Script is alive' });

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body     = parsePostBody_(e);
    const action   = body.action   || '';
    const sheetKey = body.sheet    || '';

    if (action === 'add')           return respond(addRow(sheetKey, body.data || {}));
    if (action === 'update')        return respond(updateRow(sheetKey, body.row, body.data || {}));
    if (action === 'insertSDD')     return respond(insertSddRow(body.data || {}));
    if (action === 'updateSDD')     return respond(updateSddRow(body.row, body.data || {}));
    if (action === 'delete') {
      var multi = body.rows;
      if (Array.isArray(multi) && multi.length) {
        return respond(bulkDelete(sheetKey, multi));
      }
      return respond(deleteRow(sheetKey, body.row));
    }
    if (action === 'bulkDelete')    return respond(bulkDelete(sheetKey, body.rows || []));
    if (action === 'upsertSDD')     return respond(upsertSDD(body.data || {}));
    if (action === 'bulkUpsertSDD') return respond(bulkUpsertSDD(body.rows || []));

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.message });
  }
}

function doOptions() {
  return respond({ success: true });
}

// ═══════════════════════════════════════════════════════════
//  TTP / TTM  ─ AUTO-INIT HEADERS
// ═══════════════════════════════════════════════════════════

/**
 * Pastikan tab "Monitoring TTP/TTM" punya 36 header sesuai TTP_HEADERS.
 * - Sheet kosong  → tulis semua header sekaligus di Row 1.
 * - Header sudah ada sebagian → append kolom yang belum ada di kanan.
 * Dipanggil otomatis setiap getAll('ttp').
 */
function ensureTtpHeaders_() {
  const sheet   = getSheet('ttp');
  const lastCol = sheet.getLastColumn();

  if (lastCol === 0) {
    // Sheet kosong — tulis seluruh header sekaligus
    sheet.getRange(1, 1, 1, TTP_HEADERS.length).setValues([TTP_HEADERS]);
    return;
  }

  const existing    = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
                           .map(function(h) { return String(h || '').trim(); });
  const existingSet = new Set(existing.filter(Boolean));
  const missing     = TTP_HEADERS.filter(function(h) { return !existingSet.has(h); });

  if (!missing.length) return;

  const start = existing.length + 1;
  sheet.insertColumnsAfter(existing.length, missing.length);
  sheet.getRange(1, start, 1, missing.length).setValues([missing]);
}

// ═══════════════════════════════════════════════════════════
//  SDD UPSERT  (single row)
// ═══════════════════════════════════════════════════════════

function upsertSDD(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    sanitizeSddPayload_(data);
    const sheet = getSheet('sdd');
    ensureOtherProductExtensionHeadersOrdered_(sheet);
    ensureMainProductExtensionHeadersOrdered_(sheet);
    ensureSddHeadersForPayloads_(sheet, [data || {}]);

    const rows        = sheet.getDataRange().getValues();
    if (!rows.length) throw new Error('SDD sheet is empty (no headers)');

    const headers     = rows[0];
    const headerIndex = indexByHeader_(headers);

    normalizeSubmittedFlags_(data);
    const matchIdx = findMatchingSddRowIndex_(rows, headers, headerIndex, data);

    if (matchIdx > 0) {
      const existingObj = rowToObject_(headers, rows[matchIdx]);
      assertSubmittedNotDowngraded_(existingObj, data);
      const updatedRow = headers.map(function(h, j) {
        return data[h] !== undefined ? data[h] : rows[matchIdx][j];
      });
      sheet.getRange(matchIdx + 1, 1, 1, updatedRow.length).setValues([updatedRow]);

      const millId = updatedRow[headerIndex['Mill ID']] || updatedRow[headerIndex['Mil ID']] || '';
      return { success: true, action: 'updated', row: matchIdx + 1, millId: millId };
    }

    // Insert baru
    const millIdCol = headerIndex['Mill ID'] !== undefined
      ? headerIndex['Mill ID']
      : headerIndex['Mil ID'];

    if (millIdCol !== undefined && !normalize_(data['Mill ID']) && !normalize_(data['Mil ID'])) {
      const generatedId = generateMillId(rows, millIdCol);
      data['Mill ID'] = generatedId;
      data['Mil ID']  = generatedId;
    }

    if (!data['Date Imported']) {
      data['Date Imported'] = Utilities.formatDate(
        new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
      );
    }

    const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
    sheet.appendRow(newRow);

    return { success: true, action: 'inserted', millId: data['Mill ID'] || data['Mil ID'] || '' };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

function insertSddRow(data) {
  sanitizeSddPayload_(data);
  const sheet = getSheet('sdd');
  ensureOtherProductExtensionHeadersOrdered_(sheet);
  ensureMainProductExtensionHeadersOrdered_(sheet);
  ensureSddHeadersForPayloads_(sheet, [data || {}]);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);
  return { success: true, action: 'inserted', row: sheet.getLastRow() };
}

function updateSddRow(rowNum, data) {
  sanitizeSddPayload_(data);
  const r = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for updateSDD: ' + rowNum);
  const sheet = getSheet('sdd');
  ensureOtherProductExtensionHeadersOrdered_(sheet);
  ensureMainProductExtensionHeadersOrdered_(sheet);
  ensureSddHeadersForPayloads_(sheet, [data || {}]);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];
  const updated = headers.map(function(h, i) {
    return Object.prototype.hasOwnProperty.call(data, h) ? data[h] : current[i];
  });
  sheet.getRange(r, 1, 1, headers.length).setValues([updated]);
  return { success: true, action: 'updated', row: r };
}

// ═══════════════════════════════════════════════════════════
//  SDD BULK UPSERT
// ═══════════════════════════════════════════════════════════

function bulkUpsertSDD(rowsIn) {
  if (!Array.isArray(rowsIn)) throw new Error('bulkUpsertSDD expects an array of rows');
  if (!rowsIn.length) return { success: true, count: 0, results: [] };

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet('sdd');
    ensureOtherProductExtensionHeadersOrdered_(sheet);
    ensureMainProductExtensionHeadersOrdered_(sheet);
    rowsIn.forEach(function(r) { sanitizeSddPayload_(r); });
    ensureSddHeadersForPayloads_(sheet, rowsIn);

    const raw = sheet.getDataRange().getValues();
    if (!raw.length) throw new Error('SDD sheet is empty (no headers)');

    const headers     = raw[0];
    const headerIndex = indexByHeader_(headers);
    const millIdCol   = headerIndex['Mill ID'] !== undefined
      ? headerIndex['Mill ID']
      : headerIndex['Mil ID'];

    const working = raw.map(function(r) { return r.slice(); });

    let maxExistingNum = 0;
    if (millIdCol !== undefined) {
      for (let i = 1; i < working.length; i++) {
        const m = String(working[i][millIdCol] || '').match(/^SDD-(\d+)$/);
        if (m) maxExistingNum = Math.max(maxExistingNum, parseInt(m[1], 10));
      }
    }

    const todayStamp = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'
    );

    const updates = [];
    const inserts = [];
    const results = [];

    rowsIn.forEach(function(data) {
      if (!data || typeof data !== 'object') {
        results.push({ success: false, error: 'Empty row payload' });
        return;
      }
      normalizeSubmittedFlags_(data);
      if (!data['Date Imported']) data['Date Imported'] = todayStamp;

      const matchIdx = findMatchingSddRowIndex_(working, headers, headerIndex, data);

      if (matchIdx > 0) {
        const existingObj = rowToObject_(headers, working[matchIdx]);
        assertSubmittedNotDowngraded_(existingObj, data);
        const updatedRow = headers.map(function(h, j) {
          return data[h] !== undefined ? data[h] : working[matchIdx][j];
        });
        working[matchIdx] = updatedRow;
        updates.push({ sheetRow: matchIdx + 1, values: updatedRow });
        results.push({
          success: true, action: 'updated', row: matchIdx + 1,
          millId: millIdCol !== undefined ? (updatedRow[millIdCol] || '') : ''
        });
      } else {
        if (millIdCol !== undefined && !normalize_(data['Mill ID']) && !normalize_(data['Mil ID'])) {
          maxExistingNum++;
          const generatedId = 'SDD-' + String(maxExistingNum).padStart(4, '0');
          data['Mill ID'] = generatedId;
          data['Mil ID']  = generatedId;
        }
        const newRow = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
        working.push(newRow);
        inserts.push(newRow);
        results.push({
          success: true, action: 'inserted',
          millId: millIdCol !== undefined ? (newRow[millIdCol] || '') : ''
        });
      }
    });

    updates.forEach(function(u) {
      sheet.getRange(u.sheetRow, 1, 1, headers.length).setValues([u.values]);
    });

    if (inserts.length) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, inserts.length, headers.length).setValues(inserts);
    }

    return { success: true, count: results.length, results: results };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  SDD MATCHING  &  HELPERS
// ═══════════════════════════════════════════════════════════

function findMatchingSddRowIndex_(rows, headers, idx, data) {
  const millId    = normalize_(data['Mill ID']) || normalize_(data['Mil ID']);
  const millIdCol = idx['Mill ID'] !== undefined ? idx['Mill ID'] : idx['Mil ID'];
  const payloadKind = classifySddPayload_(data);

  function millIdOk(i) {
    if (!millId || millIdCol === undefined) return true;
    return normalize_(rows[i][millIdCol]) === millId;
  }
  function rowKind(i) { return classifySddPayload_(rowToObject_(headers, rows[i])); }
  function sameKind(i) { return rowKind(i) === payloadKind; }

  // 1) FFB strict: FFB - ID
  const ffbId = normalize_(data['FFB - ID']);
  if (payloadKind === 'FFB' && ffbId && idx['FFB - ID'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['FFB - ID']]) === ffbId && millIdOk(i)) return i;
    }
  }

  // 2) FFB composite: Supplier Name + Mill Name
  const ffbSupplier = normalize_(data['FFB - Supplier Name']);
  const ffbMill     = normalize_(data['FFB - Mill Name']);
  if (payloadKind === 'FFB' && ffbSupplier && ffbMill
      && idx['FFB - Supplier Name'] !== undefined && idx['FFB - Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['FFB - Supplier Name']]) === ffbSupplier
          && normalize_(rows[i][idx['FFB - Mill Name']]) === ffbMill
          && millIdOk(i)) return i;
    }
  }

  // 3) TML strict: UML ID
  const uml = normalize_(data['UML ID']) || normalize_(data['TML - UML ID']);
  if (payloadKind === 'TML' && uml) {
    const candidates = ['UML ID', 'TML - UML ID'];
    for (let c = 0; c < candidates.length; c++) {
      const h = candidates[c];
      if (idx[h] === undefined) continue;
      for (let i = 1; i < rows.length; i++) {
        if (!sameKind(i)) continue;
        if (normalize_(rows[i][idx[h]]) === uml && millIdOk(i)) return i;
      }
    }
  }

  // 4) TML fallback: Mill Name
  const tmlMill = normalize_(data['TML - Mill Name']);
  if (payloadKind === 'TML' && tmlMill && idx['TML - Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['TML - Mill Name']]) === tmlMill && millIdOk(i)) return i;
    }
  }

  // 5) MAIN: Company Name + Mill Name
  const company = normalize_(data['Company Name']);
  const mill    = normalize_(data['Mill Name']);
  if (payloadKind === 'MAIN' && company && mill
      && idx['Company Name'] !== undefined && idx['Mill Name'] !== undefined) {
    for (let i = 1; i < rows.length; i++) {
      if (!sameKind(i)) continue;
      if (normalize_(rows[i][idx['Company Name']]) === company
          && normalize_(rows[i][idx['Mill Name']]) === mill
          && millIdOk(i)) return i;
    }
  }

  // 6) Last resort: Mill ID unique + same kind
  const hasNaturalKey = Boolean(
    (ffbSupplier && ffbMill) || ffbId || uml || tmlMill || (company && mill)
  );
  if (millId && millIdCol !== undefined && !hasNaturalKey) {
    let hit = -1, count = 0;
    for (let i = 1; i < rows.length; i++) {
      if (normalize_(rows[i][millIdCol]) === millId && sameKind(i)) { hit = i; count++; }
    }
    if (count === 1) return hit;
  }

  return -1;
}

function normalize_(v) {
  return String(v === undefined || v === null ? '' : v).trim();
}

function isSubmitted_(rowObj) {
  const st   = normalize_(rowObj['SCR - Screening Status']).toLowerCase();
  const flag = normalize_(rowObj['isSubmitted']).toLowerCase();
  return st === 'submitted' || flag === 'true';
}

function normalizeSubmittedFlags_(payload) {
  if (!payload || typeof payload !== 'object') return;
  const submitted = isSubmitted_(payload);
  if (submitted) {
    payload['SCR - Screening Status'] = 'Submitted';
    payload['isSubmitted'] = 'true';
  } else {
    if (normalize_(payload['SCR - Screening Status']).toLowerCase() === 'draft') {
      payload['SCR - Screening Status'] = 'Draft';
    }
    if (payload['isSubmitted'] === true || payload['isSubmitted'] === false) {
      payload['isSubmitted'] = payload['isSubmitted'] ? 'true' : 'false';
    }
  }
}

function assertSubmittedNotDowngraded_(existingRow, incomingRow) {
  if (isSubmitted_(existingRow) && !isSubmitted_(incomingRow)) {
    throw new Error('Row already submitted and locked; draft overwrite is blocked.');
  }
}

function rowToObject_(headers, rowValues) {
  const out = {};
  for (let i = 0; i < headers.length; i++) out[headers[i]] = rowValues[i];
  return out;
}

function classifySddPayload_(rowObj) {
  const hasFfb = normalize_(rowObj['FFB - Supplier Name']) || normalize_(rowObj['FFB - ID']);
  const hasTml = normalize_(rowObj['TML - Mill Name'])
              || normalize_(rowObj['TML - UML ID'])
              || normalize_(rowObj['UML ID']);
  if (hasFfb) return 'FFB';
  if (hasTml) return 'TML';
  return 'MAIN';
}

function indexByHeader_(headers) {
  const map = {};
  headers.forEach(function(h, i) { map[String(h)] = i; });
  return map;
}

function sanitizeSddPayload_(data) {
  if (!data || typeof data !== 'object') return;
  delete data._row;
  delete data.__row;
}

// ═══════════════════════════════════════════════════════════
//  SDD HEADER AUTO-EXPAND
// ═══════════════════════════════════════════════════════════

function getSddOtherProductExtensionHeaders_() {
  var list = [];
  for (var k = 4; k <= 12; k++) {
    list.push('Other Product ' + k);
    list.push('Other Product ' + k + ' Avg/Month (Ton)');
  }
  return list;
}

function getSddMainProductExtensionHeaders_() {
  var list = [];
  for (var k = 2; k <= 3; k++) {
    list.push('Main Product ' + k);
    list.push('Main Product ' + k + ' Avg Production/Month (Ton)');
    list.push('Main Product ' + k + ' Yield');
  }
  return list;
}

function ensureOtherProductExtensionHeadersOrdered_(sheet) {
  if (!sheet) return;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var want    = getSddOtherProductExtensionHeaders_();
  var present = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (h) present[h] = true;
  }
  var missing = want.filter(function(w) { return !present[w]; });
  if (!missing.length) return;

  var cpoCol = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).trim() === 'CPO Quality - FFA') { cpoCol = j + 1; break; }
  }
  if (cpoCol > 0) {
    sheet.insertColumnsBefore(cpoCol, missing.length);
    sheet.getRange(1, cpoCol, 1, missing.length).setValues([missing]);
  } else {
    sheet.insertColumnsAfter(headers.length, missing.length);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureMainProductExtensionHeadersOrdered_(sheet) {
  if (!sheet) return;
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var want    = getSddMainProductExtensionHeaders_();
  var present = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').trim();
    if (h) present[h] = true;
  }
  var missing = want.filter(function(w) { return !present[w]; });
  if (!missing.length) return;

  var cpoCol = -1;
  for (var j = 0; j < headers.length; j++) {
    if (String(headers[j]).trim() === 'CPO Quality - FFA') { cpoCol = j + 1; break; }
  }
  if (cpoCol > 0) {
    sheet.insertColumnsBefore(cpoCol, missing.length);
    sheet.getRange(1, cpoCol, 1, missing.length).setValues([missing]);
  } else {
    sheet.insertColumnsAfter(headers.length, missing.length);
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureSddHeadersForPayloads_(sheet, payloadRows) {
  if (!sheet) return;
  if (!Array.isArray(payloadRows) || !payloadRows.length) return;

  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const existing = new Set(
    headers.map(function(h) { return String(h || '').trim(); }).filter(Boolean)
  );

  const missing = [];
  payloadRows.forEach(function(row) {
    if (!row || typeof row !== 'object') return;
    Object.keys(row).forEach(function(k) {
      const key = String(k || '').trim();
      if (!key) return;
      if (!existing.has(key)) { existing.add(key); missing.push(key); }
    });
  });
  if (!missing.length) return;

  const start = headers.length + 1;
  sheet.insertColumnsAfter(headers.length, missing.length);
  sheet.getRange(1, start, 1, missing.length).setValues([missing]);
}

function generateMillId(rows, millIdCol) {
  let maxNum = 0;
  for (let i = 1; i < rows.length; i++) {
    const id    = String(rows[i][millIdCol] || '');
    const match = id.match(/^SDD-(\d+)$/);
    if (match) maxNum = Math.max(maxNum, parseInt(match[1], 10));
  }
  return 'SDD-' + String(maxNum + 1).padStart(4, '0');
}

// ═══════════════════════════════════════════════════════════
//  GET BY MILL ID
// ═══════════════════════════════════════════════════════════

function getByMillId(millId) {
  const sheet = getSheet('sdd');
  const rows  = sheet.getDataRange().getValues();
  if (!rows.length) return { success: false, error: 'Empty sheet' };

  const headers   = rows[0];
  const millIdCol = headers.indexOf('Mill ID');
  const milIdCol  = headers.indexOf('Mil ID');
  if (millIdCol < 0 && milIdCol < 0) {
    return { success: false, error: '"Mill ID" or "Mil ID" header not found' };
  }

  const want = String(millId || '').trim();
  for (let i = 1; i < rows.length; i++) {
    const okMill = millIdCol >= 0 && String(rows[i][millIdCol]) === want;
    const okMil  = milIdCol  >= 0 && String(rows[i][milIdCol])  === want;
    if (okMill || okMil) {
      const obj = { success: true, _row: i + 1 };
      headers.forEach(function(h, j) { obj[h] = rows[i][j]; });
      return obj;
    }
  }
  return { success: false, error: 'Not found' };
}

// ═══════════════════════════════════════════════════════════
//  GENERIC CRUD
// ═══════════════════════════════════════════════════════════

/**
 * Dashboard (sustain-dashboard) mengirim QUARTER / YEAR (uppercase).
 * Sheet "Mill Onboarding Profile" sering memakai header "Quarter" / "Year".
 * Salin payload ke nama kolom yang tepat sebelum map headers → baris.
 */
function resolveMillQuarterYearKeys_(data, headers) {
  if (!data || typeof data !== 'object' || !Array.isArray(headers)) return;
  var list = headers.map(function(x) { return String(x || '').trim(); });
  function findQuarterCol() {
    for (var i = 0; i < list.length; i++) {
      if (/^quarter$/i.test(list[i])) return list[i];
    }
    return null;
  }
  function findYearCol() {
    for (var i = 0; i < list.length; i++) {
      var h = list[i];
      if (h === 'Year' || h === 'YEAR') return h;
    }
    for (var j = 0; j < list.length; j++) {
      var y = list[j];
      if (!/^year$/i.test(y)) continue;
      if (/planted|capacity|mill|tanam|ffb|tml|issue|expiry|average/i.test(y)) continue;
      return y;
    }
    return null;
  }
  var qCol = findQuarterCol();
  var yCol = findYearCol();
  if (qCol && data['QUARTER'] !== undefined && String(data['QUARTER']).trim() !== '' &&
      (data[qCol] === undefined || String(data[qCol]).trim() === '')) {
    data[qCol] = data['QUARTER'];
  }
  if (yCol && data['YEAR'] !== undefined && String(data['YEAR']).trim() !== '' &&
      (data[yCol] === undefined || String(data[yCol]).trim() === '')) {
    data[yCol] = data['YEAR'];
  }
}

/** Tambahkan QUARTER / YEAR di response getAll('mill') agar selaras dengan klien. */
function mirrorMillQuarterYearOnRead_(obj) {
  if (!obj || typeof obj !== 'object') return;
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    if (/^quarter$/i.test(String(k).trim()) && obj['QUARTER'] === undefined) obj['QUARTER'] = obj[k];
  });
  Object.keys(obj).forEach(function(k) {
    if (k === '_row') return;
    var t = String(k).trim();
    if ((t === 'Year' || t === 'YEAR') && obj['YEAR'] === undefined) obj['YEAR'] = obj[k];
  });
}

function getData(sheetKey) {
  const sheet = getSheet(sheetKey);
  const rows  = sheet.getDataRange().getValues();
  if (!rows.length) return [];

  const headers = rows[0];
  return rows.slice(1).map(function(row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = row[j]; });
    if (sheetKey === 'mill') mirrorMillQuarterYearOnRead_(obj);
    return obj;
  });
}

function addRow(sheetKey, data) {
  const sheet   = getSheet(sheetKey);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'mill') resolveMillQuarterYearKeys_(data, headers);
  const newRow  = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(newRow);
  return { success: true };
}

function updateRow(sheetKey, rowNum, data) {
  const sheet = getSheet(sheetKey);
  const r     = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for update: ' + rowNum);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (sheetKey === 'mill') resolveMillQuarterYearKeys_(data, headers);
  const current = sheet.getRange(r, 1, 1, headers.length).getValues()[0];

  const updatedRow = headers.map(function(h, j) {
    return data[h] !== undefined ? data[h] : current[j];
  });
  sheet.getRange(r, 1, 1, updatedRow.length).setValues([updatedRow]);
  return { success: true };
}

function deleteRow(sheetKey, rowNum) {
  const sheet = getSheet(sheetKey);
  const r     = Number(rowNum);
  if (!r || r < 2) throw new Error('Invalid row number for delete: ' + rowNum);
  sheet.deleteRow(r);
  return { success: true };
}

function bulkDelete(sheetKey, rowNums) {
  if (!Array.isArray(rowNums) || !rowNums.length) return { success: true, deleted: 0 };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(sheetKey);
    let validRows = rowNums
      .map(function(r) { return Number(r); })
      .filter(function(r) { return r >= 2 && !isNaN(r); });
    const seen = {};
    validRows = validRows.filter(function(r) {
      if (seen[r]) return false;
      seen[r] = true;
      return true;
    });
    if (!validRows.length) return { success: true, deleted: 0 };
    validRows.sort(function(a, b) { return b - a; });
    validRows.forEach(function(r) { sheet.deleteRow(r); });
    return { success: true, deleted: validRows.length };
  } finally {
    try { lock.releaseLock(); } catch (e) { /* ignore */ }
  }
}

// ═══════════════════════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════════════════════

function parsePostBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing request body');
  }
  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    throw new Error('Invalid JSON body: ' + String(e.postData.contents).slice(0, 200));
  }
}

function getSheet(sheetKey) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEETS[sheetKey];
  if (!name) throw new Error('Sheet key not found: ' + sheetKey);
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Tab not found: ' + name);
  return sheet;
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
