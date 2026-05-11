/**
 * PATCH — untuk Apps Script v3 (full backend yang Anda paste).
 *
 * Cara pakai:
 * 1) Salin function resolveMillQuarterYearKeys_ + mirrorMillQuarterYearOnRead_ ke project GAS
 *    (misalnya sebelum section GENERIC CRUD).
 * 2) Ganti function getData, addRow, updateRow di script Anda dengan versi di bawah
 *    (sudah termasuk panggilan helper untuk sheetKey === 'mill').
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
    for (var j = 0; j < list.length; j++) {
      var h = list[j];
      if (h === 'Year' || h === 'YEAR') return h;
    }
    for (var k = 0; k < list.length; k++) {
      var y = list[k];
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

// ─── Ganti seluruh function berikut di script v3 Anda ─────────────────────

function getData(sheetKey) {
  const sheet = getSheet(sheetKey);
  const rows  = sheet.getDataRange().getValues();
  if (!rows.length) return [];
  const headers = rows[0];
  return rows.slice(1).map(function(row, i) {
    const obj = { _row: i + 2 };
    headers.forEach(function(h, j) { obj[h] = rows[i + 1][j]; });
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
  const updated = headers.map(function(h, j) {
    return data[h] !== undefined ? data[h] : current[j];
  });
  sheet.getRange(r, 1, 1, updated.length).setValues([updated]);
  return { success: true };
}
