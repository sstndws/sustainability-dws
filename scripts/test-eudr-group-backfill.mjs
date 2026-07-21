/**
 * Simulates EUDR registry dedup: UML ID primary key, group name backfill.
 */
import assert from 'node:assert/strict';

function eudrNormalizeUmlId_(val) {
  const s = String(val == null ? '' : val).trim();
  if (!s || s === '—' || s === '-') return '';
  if (/^n\/?a$/i.test(s)) return '';
  return s.toLowerCase();
}

function eudrEntityKey_(group, company, mill, umlId) {
  const uml = eudrNormalizeUmlId_(umlId);
  if (uml) return 'uml:' + uml;
  return [group, company, mill].map(function(s) {
    return String(s || '').trim().toLowerCase();
  }).join('|');
}

function pickMillGroupName_(row) {
  const s = String(row['GROUP NAME'] || '').trim();
  return s && s !== '—' && s !== '-' ? s : '';
}

function pickMillCompanyName_(row) {
  return String(row['COMPANY NAME'] || '').trim();
}

function millPickField_(row, keys) {
  for (const k of keys) {
    const v = row[k];
    if (v != null && String(v).trim()) return String(v).trim();
  }
  return '';
}

function eudrPickUmlId_(row) {
  return millPickField_(row, ['UML ID']);
}

function millYearVal(r) { return String(r.YEAR || r['YEAR'] || '').trim(); }
function millMonthVal(r) { return String(r.MONTH || r['MONTH'] || '').trim(); }

function eudrCompanyMillKey_(company, mill) {
  return eudrEntityKey_('', company, mill);
}

function eudrRegistryDedupKey_(company, mill, umlId) {
  const uml = eudrNormalizeUmlId_(umlId);
  if (uml) return 'uml:' + uml;
  return eudrCompanyMillKey_(company, mill);
}

function eudrBestGroupNameForMill_(allData, company, mill, umlId) {
  const uml = eudrNormalizeUmlId_(umlId);
  if (uml) {
    for (let i = 0; i < (allData || []).length; i++) {
      const r = allData[i];
      if (eudrNormalizeUmlId_(eudrPickUmlId_(r)) !== uml) continue;
      const g = pickMillGroupName_(r);
      if (g) return g;
    }
  }
  let best = '';
  (allData || []).forEach(function(r) {
    if (pickMillCompanyName_(r) !== company) return;
    if (millPickField_(r, ['MILL NAME']) !== mill) return;
    const g = pickMillGroupName_(r);
    if (g) best = g;
  });
  return best;
}

function eudrBuildMillRegistry_(allData) {
  const seen = {};
  const groups = {};
  const dedupArr = [];
  (allData || []).forEach(function(r) {
    const company = pickMillCompanyName_(r);
    const mill = millPickField_(r, ['MILL NAME']);
    const umlId = eudrPickUmlId_(r);
    const key = eudrRegistryDedupKey_(company, mill, umlId);
    if (!company || !mill || !key || key === '|') return;
    const group = pickMillGroupName_(r);
    if (group) groups[key] = group;
    if (seen[key] === undefined) {
      seen[key] = dedupArr.length;
      dedupArr.push(r);
    } else {
      const prev = dedupArr[seen[key]];
      const prevYear = parseInt(millYearVal(prev), 10) || 0;
      const curYear = parseInt(millYearVal(r), 10) || 0;
      const prevMonth = parseInt(millMonthVal(prev), 10) || 0;
      const curMonth = parseInt(millMonthVal(r), 10) || 0;
      if (curYear > prevYear || (curYear === prevYear && curMonth > prevMonth)) {
        dedupArr[seen[key]] = r;
      }
    }
  });
  return dedupArr.map(function(r) {
    const company = pickMillCompanyName_(r);
    const mill = millPickField_(r, ['MILL NAME']);
    const umlId = eudrPickUmlId_(r);
    const dedupKey = eudrRegistryDedupKey_(company, mill, umlId);
    const group = pickMillGroupName_(r) || groups[dedupKey] || eudrBestGroupNameForMill_(allData, company, mill, umlId);
    return { 'GROUP NAME': group, 'COMPANY NAME': company, 'MILL NAME': mill, 'UML ID': umlId };
  });
}

// Group backfill: newest period row missing group
const patiware = [
  { YEAR: '2024', MONTH: '6', 'GROUP NAME': 'KPN PLANTATION', 'COMPANY NAME': 'PATIWARE', 'MILL NAME': 'PATIWARE', 'UML ID': 'PO100203923' },
  { YEAR: '2025', MONTH: '12', 'GROUP NAME': '', 'COMPANY NAME': 'PATIWARE', 'MILL NAME': 'PATIWARE', 'UML ID': 'PO100203923' },
];
const patiOut = eudrBuildMillRegistry_(patiware);
assert.equal(patiOut.length, 1);
assert.equal(patiOut[0]['GROUP NAME'], 'KPN PLANTATION');
assert.equal(patiOut[0]['UML ID'], 'PO100203923');

// Same company+mill, different UML → separate entries
const multiUml = [
  { YEAR: '2025', MONTH: '12', 'COMPANY NAME': 'ACME', 'MILL NAME': 'ACME MILL', 'UML ID': 'PO1000000001' },
  { YEAR: '2025', MONTH: '12', 'COMPANY NAME': 'ACME', 'MILL NAME': 'ACME MILL', 'UML ID': 'PO1000000002' },
];
assert.equal(eudrBuildMillRegistry_(multiUml).length, 2);

// Same UML, different mill spelling → one entry
const sameUml = [
  { YEAR: '2024', MONTH: '1', 'COMPANY NAME': 'EKAJAYA', 'MILL NAME': 'EKA JAYA MULTI PERKASA', 'UML ID': 'PO1000004490' },
  { YEAR: '2025', MONTH: '12', 'COMPANY NAME': 'EKAJAYA', 'MILL NAME': 'EKAJAYA MULTI PERKASA', 'UML ID': 'PO1000004490' },
];
assert.equal(eudrBuildMillRegistry_(sameUml).length, 1);

console.log('test-eudr-group-backfill: OK');
