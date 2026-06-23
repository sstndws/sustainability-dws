#!/usr/bin/env node
/** Debug Monthly Report KPI counts against live GAS data. */
import {
  mrdBuildTtpByMillMaps_,
  mrdTtpRowsForMill_,
  mrdTtpRowHasSupplier_,
  mrdResolveTtpSupplierCol_,
} from '../src/monthly-report-labels.js';

const GAS_URL =
  process.env.GAS_WEBAPP_URL ||
  'https://script.google.com/macros/s/AKfycbzxvlLyrztPd5EqCRnesji7yQfvfqzriAGAiKazqTgjGFenFUtI_bm_Zx3u1x-xG_PvoA/exec';

async function gasGet(sheet) {
  const url = GAS_URL + '?' + new URLSearchParams({ action: 'getAll', sheet }).toString();
  const res = await fetch(url, { redirect: 'follow' });
  const text = await res.text();
  const data = JSON.parse(text);
  if (Array.isArray(data)) return data;
  if (Array.isArray(data.data)) return data.data;
  if (Array.isArray(data.rows)) return data.rows;
  return [];
}

function parseYear(v) {
  const n = parseInt(String(v || '').replace(/\D/g, ''), 10);
  return isNaN(n) ? '' : String(n);
}

function millYear(r) {
  return r.YEAR || r['REPORTING YEAR'] || r['Reporting Year'] || '';
}

function millMonth(r) {
  return String(r.MONTH || r['REPORTING MONTH'] || r['Reporting Month'] || '').trim();
}

function mrdDataPeriodFromReport(year, month) {
  const y = parseInt(year, 10);
  const m = parseInt(month, 10);
  if (!y || !m) return { year: String(year || ''), month: String(month || '') };
  if (m === 1) return { year: String(y - 1), month: '12' };
  return { year: String(y), month: String(m - 1) };
}

const reportYear = '2025';
const reportMonth = '12';
const dataPeriod = mrdDataPeriodFromReport(reportYear, reportMonth);

console.log('GAS URL:', GAS_URL.slice(0, 60) + '...');
console.log('Report:', reportYear, 'month', reportMonth, '→ data period', dataPeriod);

const [millData, ttpData] = await Promise.all([gasGet('mill'), gasGet('ttp')]);
console.log('Mill rows:', millData.length, '| TTP rows:', ttpData.length);

const mills = millData.filter(function(r) {
  const y = parseYear(millYear(r));
  if (dataPeriod.year && y && y !== dataPeriod.year) return false;
  if (dataPeriod.month && millMonth) {
    const m = String(millMonth(r) || '').trim();
    if (m && m !== dataPeriod.month) return false;
  }
  return true;
});
console.log('Mills in data period:', mills.length);

const ttpFields = ttpData.length ? Object.keys(ttpData[0]).filter((k) => k !== '_row') : [];
const millCol = ttpFields.find((h) => String(h).toUpperCase() === 'MILL NAME') || 'MILL NAME';
const groupCol = ttpFields.find((h) => String(h).toUpperCase() === 'GROUP NAME') || 'GROUP NAME';
const companyCol = ttpFields.find((h) => String(h).toUpperCase() === 'COMPANY NAME') || 'COMPANY NAME';
const supplierCol = mrdResolveTtpSupplierCol_(ttpFields);
const yearCol = ttpFields.find((h) => String(h).toUpperCase() === 'YEAR') || 'YEAR';

const ttpFiltered = ttpData.filter(function(r) {
  const y = parseYear(r[yearCol] || millYear(r));
  return !dataPeriod.year || !y || y === dataPeriod.year;
});
console.log('TTP filtered year', dataPeriod.year + ':', ttpFiltered.length);
console.log('Supplier col:', supplierCol, '| ttpFields sample:', ttpFields.slice(0, 8).join(', '));

const millKeys = new Map();
mills.forEach(function(r) {
  const key = [r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']]
    .map((x) => String(x || '').trim().toLowerCase())
    .join('|');
  if (!millKeys.has(key)) millKeys.set(key, r);
});

const ttpMillMaps = mrdBuildTtpByMillMaps_(ttpFiltered, millCol, groupCol, companyCol);
let emptyCount = 0;
let withSupplier = 0;
millKeys.forEach(function(millRow) {
  const rows = mrdTtpRowsForMill_(ttpMillMaps, millRow);
  const has = rows.some((r) => mrdTtpRowHasSupplier_(r, supplierCol));
  if (has) withSupplier++;
  else emptyCount++;
});

console.log('\n=== UNTRACEABLE ===');
console.log('With supplier:', withSupplier);
console.log('Untraceable (empty):', emptyCount);
console.log('Expected ~105 untraceable, ~103 with supplier');
