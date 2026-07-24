/**
 * Monthly Report (Detail) — read-only compliance snapshot (fast, in-memory first).
 */

import { buildMonthlyReportPdfPair_, MRD_PDF_SECTIONS, mrdPdfSectionsNoDupHighRisk_ } from './monthly-report-pdf.js';
import {
  MRD_SDD_TABLE_TITLE,
  MRD_SDD_COLS,
  MRD_GRV_SUMMARY_COLS,
  MRD_MONTH_SHORT_,
  facilityPctColLabel,
  facilityTtmColLabel,
  mrdSortSddRows_,
  mrdSortMillItems_,
  mrdSortGrvItemsByDateDesc_,
  mrdSortEudrItems_,
  eudrCpoSupply_,
  eudrPkSupply_,
  mrdSortFacilityBundles_,
  mrdSortFacilityCompanies_,
  mrdSortBundlesByFacility_,
  mrdSortEmptyMillItems_,
  mrdResolveTtpSupplierCol_,
  mrdTtpRowHasSupplier_,
  mrdBuildTtpByMillMaps_,
  mrdTtpRowsForMill_,
  mrdFormatNblRisers_,
  mrdReportHeaderMeta_,
  mrdTraceYearFromReport_,
  grvGroupName_,
  mrdShowInMillOnboarding_,
  normalizeSddCategory,
  sddStatusText,
  sddCompanyName,
  sddDateImport,
  sddLastUpdate,
  mrdBuildEudrPotentialCompanySet_,
  mrdFacilityEudrPotentialCount_,
  mrdUniqueFacilityCompanyCount_,
  mrdCompanyIsEudrPotential_,
  mrdMillProductLabel_,
  mrdEudrPotentialLabel_,
} from './monthly-report-labels.js';
import { millRiskReason_, millRiskReasonTokens_ } from './mill-risk-reason.js';
import {
  dashLoadingHtml_,
  dashSetButtonBusy_,
  dashClearButtonBusy_,
} from './dash-loading.js';
import {
  buildMrdExecutiveData_,
  buildMillPeriodKpis_,
  buildQuarterComparison_,
  renderMrdExecutiveCharts_,
  destroyMrdExecutiveCharts_,
  collectMrdExecutiveChartImages_,
  exportMrdExecutivePdf_,
  mrdExecutiveFilename_,
  mrdExecutiveHeaderMeta_,
  monthToQuarter_,
  previousQuarter_,
  quarterEndMonth_,
  formatDelta_,
  formatQtyDisplay_,
} from './monthly-report-executive.js';

const MRD_ROW_LIMIT = 5000;
const MRD_SDD_LIMIT = 5000;

let _deps = null;
let _bound = false;
let _snapshot = null;
let _search = '';
let _year = String(new Date().getFullYear());
let _month = ''; // will be auto-corrected to latest data month on load
let _monthAutoSelected = false;
let _expanded = new Set();
let _loadGen = 0;
let _eudrPending = false;
let _facilityPending = false;
let _nblByCache = new Map();
let _mrdExportInFlight = false;
let _sddCache = [];
let _facilityBundles = [];
let _facilityBundlesPeriodKey = '';
let _lastTtpCount = 0;
let _lastEudrTtpCount = -1;
let _eudrFetchOk = false;
let _ttpFetchOk = false;
let _renderRaf = 0;
let _viewMode = 'executive';
let _mrdExecChartModule = null;
let _mrdExecDataCache = null;
let _mrdExecExportInFlight = false;

function mrdDataPeriodShortLabel_(dataPeriod) {
  const y = String((dataPeriod && dataPeriod.year) || '').trim();
  const m = dataPeriod && dataPeriod.month ? parseInt(dataPeriod.month, 10) : 0;
  if (y && m >= 1 && m <= 12) {
    return (MRD_MONTH_SHORT_[m] || String(m)) + ' ' + y;
  }
  if (y) return 'Full year ' + y;
  return 'all periods';
}

/** Facility Performance follows the same −1 month data period as the other sections
 * (Supplied CPO/PK data lags one month behind the reporting month). */
function getFacilityPeriod_() {
  return getDataPeriod_();
}

/** Monthly Report facility build: supply sheets use data period; mill rows use as-of report period. */
function getFacilityReportContext_() {
  const report = getReportPeriod_();
  const data = getFacilityPeriod_();
  return {
    year: data.year,
    month: data.month,
    reportYear: report.year,
    reportMonth: report.month,
    millPickMode: 'as-of',
  };
}

function mrdFacilityPeriodKey_(facilityPeriod) {
  const report = getReportPeriod_();
  const dp = facilityPeriod && facilityPeriod.year != null ? facilityPeriod : getFacilityPeriod_();
  return String(report.year || '') + '|' + String(report.month || '')
    + '|data:' + String(dp.year || '') + '|' + String(dp.month || '');
}

function scheduleRenderAll() {
  if (_renderRaf) cancelAnimationFrame(_renderRaf);
  _renderRaf = requestAnimationFrame(function() {
    _renderRaf = 0;
    renderAll();
  });
}

function withTimeout(promise, ms, label) {
  return Promise.race([
    promise,
    new Promise(function(_, reject) {
      setTimeout(function() {
        reject(new Error((label || 'Request') + ' timeout (' + ms + 'ms)'));
      }, ms);
    }),
  ]);
}

function esc(s) {
  return _deps && _deps.escHtml ? _deps.escHtml(s) : String(s == null ? '' : s);
}

function parseYear(v) {
  const n = parseInt(String(v || '').replace(/\D/g, ''), 10);
  return isNaN(n) ? '' : String(n);
}

/** Years for MRD filters — wide default range plus any year found in loaded data. */
function mrdCollectAvailableYears_(extraYears) {
  const years = new Set();
  const now = new Date().getFullYear();
  for (let y = now + 5; y >= 2000; y--) {
    years.add(String(y));
  }
  (_deps.getMillData() || []).forEach(function(r) {
    const y = parseYear(_deps.millYearVal(r));
    if (y) years.add(y);
  });
  (extraYears || []).forEach(function(y) {
    const py = parseYear(y);
    if (py) years.add(py);
  });
  return years;
}

function mrdRenderYearSelectOptions_(selectEl, selectedYear) {
  if (!selectEl) return;
  const sel = parseYear(selectedYear) || String(new Date().getFullYear());
  const years = mrdCollectAvailableYears_([sel]);
  years.add(sel);
  selectEl.innerHTML = Array.from(years).sort(function(a, b) {
    return Number(b) - Number(a);
  }).map(function(y) {
    return '<option value="' + y + '"' + (y === sel ? ' selected' : '') + '>' + y + '</option>';
  }).join('');
}

function mrdReadExportYear_() {
  const yearEl = document.getElementById('mrdExportYearSel');
  const raw = yearEl ? yearEl.value : _year;
  const y = parseYear(raw);
  if (y) return y;
  return String(new Date().getFullYear());
}

/** Sync year/month from toolbar before export, header, or reload. */
function syncPeriodFromUi_() {
  const yearSel = document.getElementById('mrdYearSel');
  const monthSel = document.getElementById('mrdMonthSel');
  if (yearSel && yearSel.value) _year = yearSel.value;
  if (monthSel) _month = monthSel.value;
  return { year: _year, month: _month };
}

/** Find the latest month (1-12) that has actual data for the given year. Returns 0 if none. */
function mrdFindLatestDataMonthForYear_(year) {
  if (!_deps) return 0;
  const y = String(year || _year || '');
  const millYear = _deps.millYearVal;
  const millMonth = _deps.millMonthVal;
  const mainSrc = (_deps.getMillRawRows ? _deps.getMillRawRows() : (_deps.getMillData ? _deps.getMillData() : [])) || [];
  const wasteSrc = (_deps.getMillWasteRawRows ? _deps.getMillWasteRawRows() : []) || [];
  let maxMonth = 0;
  [].concat(mainSrc, wasteSrc).forEach(function(r) {
    if (!r || !mrdShowInMillOnboarding_(r)) return;
    const ry = parseYear(millYear(r));
    if (y && ry && ry !== y) return;
    const rm = parseInt(String(millMonth(r) || '').trim(), 10);
    if (rm >= 1 && rm <= 12 && rm > maxMonth) maxMonth = rm;
  });
  return maxMonth;
}

/** Disable month options with no data and auto-select latest month if current selection is empty. */
function mrdSyncMonthOptions_() {
  const monthSel = document.getElementById('mrdMonthSel');
  if (!monthSel) return;
  const latestMonth = mrdFindLatestDataMonthForYear_(_year);
  Array.from(monthSel.options).forEach(function(opt) {
    const v = parseInt(opt.value, 10);
    if (!v) { opt.disabled = false; return; } // "Full year" always enabled
    opt.disabled = latestMonth > 0 && v > latestMonth;
  });
  // Auto-select latest data month on first load if no selection made yet
  const cur = parseInt(_month, 10);
  const curOpt = cur ? monthSel.querySelector('option[value="' + cur + '"]') : null;
  if (!_month || (curOpt && curOpt.disabled) || !_monthAutoSelected) {
    if (latestMonth > 0) {
      _month = String(latestMonth);
      monthSel.value = _month;
    } else if (!_month) {
      _month = '';
      monthSel.value = '';
    }
    _monthAutoSelected = true;
  }
}

function applyDefaultMonthSelection_() {
  const monthSel = document.getElementById('mrdMonthSel');
  if (!monthSel) return;
  if (_month) {
    monthSel.value = _month;
  } else {
    _month = monthSel.value;
  }
}

/** Reporting period from toolbar (year + month user selected). */
function getReportPeriod_() {
  syncPeriodFromUi_();
  return { year: _year, month: _month };
}

/** @deprecated Header lag only — tables filter by {@link getReportPeriod_}. */
function getDataPeriod_() {
  const report = getReportPeriod_();
  return mrdDataPeriodFromReport_(report.year, report.month);
}

function parseYearFromDate_(raw) {
  if (!raw) return '';
  if (raw instanceof Date && !isNaN(raw.getTime())) return String(raw.getFullYear());
  const s = String(raw).trim();
  const iso = s.match(/^(\d{4})/);
  if (iso) return iso[1];
  const d = new Date(s);
  if (!isNaN(d.getTime())) return String(d.getFullYear());
  return '';
}

function mrdMonthMatchesFilter_(rowMonth, targetMonth) {
  if (!targetMonth) return true;
  const rm = parseInt(String(rowMonth || '').trim(), 10);
  const tm = parseInt(String(targetMonth || '').trim(), 10);
  if (rm >= 1 && rm <= 12 && tm >= 1 && tm <= 12) return rm === tm;
  return String(rowMonth || '').trim() === String(targetMonth || '').trim();
}

function mrdDataPeriodFromReport_(year, month) {
  const meta = mrdReportHeaderMeta_(year, month);
  return { year: meta.dataYear, month: meta.dataMonth };
}

function mrdResolveReportPeriod_(override) {
  if (override && (override.year != null || override.month != null)) {
    return {
      year: override.year != null ? String(override.year) : _year,
      month: override.month != null ? String(override.month) : _month,
    };
  }
  return getReportPeriod_();
}

function parseMonthFromDate(raw) {
  if (!raw) return '';
  if (raw instanceof Date && !isNaN(raw.getTime())) return String(raw.getMonth() + 1);
  const s = String(raw).trim();
  const iso = s.match(/^(\d{4})-(\d{2})/);
  if (iso) return String(parseInt(iso[2], 10));
  const d = new Date(s);
  if (!isNaN(d.getTime())) return String(d.getMonth() + 1);
  return '';
}

function matchesSearch(blob) {
  if (!_search) return true;
  return String(blob || '').toLowerCase().includes(_search);
}

function isHighRisk(val) {
  return String(val || '').toLowerCase().includes('high');
}

/** Same resolved risk as Mill Registry table (RESULT RISK LEVEL → RISK LEVEL). */
function mrdResolvedRisk_(rowOrItem) {
  const item = rowOrItem && rowOrItem.row ? rowOrItem : null;
  const r = item ? item.row : rowOrItem;
  if (item && String(item.risk || '').trim()) return String(item.risk).trim();
  if (_deps && _deps.millResolvedRiskLevel && r) return _deps.millResolvedRiskLevel(r) || '';
  return '';
}

function mrdIsHighRiskItem_(rowOrItem) {
  return isHighRisk(mrdResolvedRisk_(rowOrItem));
}

function mrdRiskReasonTokens_(rowOrItem) {
  const item = rowOrItem && rowOrItem.row ? rowOrItem : null;
  const r = item ? item.row : rowOrItem;
  if (!r) return [];
  const opts = {
    millIsNblYes: _deps && _deps.millIsNblYes_ ? _deps.millIsNblYes_ : undefined,
  };
  if (_deps && _deps.millRiskReasonTokens) return _deps.millRiskReasonTokens(r) || [];
  return millRiskReasonTokens_(r, opts);
}

function mrdRiskReason_(rowOrItem) {
  const tokens = mrdRiskReasonTokens_(rowOrItem);
  if (tokens.length) return tokens.join('; ');
  const item = rowOrItem && rowOrItem.row ? rowOrItem : null;
  const r = item ? item.row : rowOrItem;
  if (!r) return '';
  if (_deps && _deps.millRiskReason) return _deps.millRiskReason(r) || '';
  if (_deps && _deps.millHighRiskReason) return _deps.millHighRiskReason(r) || '';
  return millRiskReason_(r, {
    millIsNblYes: _deps && _deps.millIsNblYes_ ? _deps.millIsNblYes_ : undefined,
  });
}

function mrdRiskReasonPillClass_(label) {
  const s = String(label || '').toLowerCase();
  if (s.includes('no buy')) return 'mill-risk-reason-pill--nbl';
  if (s.includes('deforest')) return 'mill-risk-reason-pill--deforest';
  if (s.includes('grievance')) return 'mill-risk-reason-pill--grievance';
  if (s.includes('coordinate') || s.includes('legality') || s.includes('apl')
    || s.includes('ndpe') || s.includes('certification')) {
    return 'mill-risk-reason-pill--gap';
  }
  return 'mill-risk-reason-pill--other';
}

function mrdRiskReasonCell_(rowOrItem) {
  const tokens = mrdRiskReasonTokens_(rowOrItem);
  if (!tokens.length) return '<span class="mrd-muted">—</span>';
  const title = tokens.join('; ');
  return '<div class="cert-pill-list mill-risk-reason-pill-list" title="' + esc(title) + '">'
    + tokens.map(function(tok) {
      return '<span class="cert-pill mill-risk-reason-pill ' + mrdRiskReasonPillClass_(tok) + '">'
        + esc(tok) + '</span>';
    }).join('')
    + '</div>';
}

function isNblYes(val) {
  return _deps && _deps.millIsNblYes_ ? _deps.millIsNblYes_(val) : /yes|nbl|no buy/i.test(String(val || ''));
}

function riskPill(val) {
  const s = String(val || '').trim();
  if (!s) return '<span class="mrd-muted">—</span>';
  const lower = s.toLowerCase();
  let cls = 'mrd-pill';
  if (lower.includes('high')) cls += ' mrd-pill--high';
  else if (lower.includes('med')) cls += ' mrd-pill--med';
  else if (lower.includes('low')) cls += ' mrd-pill--low';
  return '<span class="' + cls + '">' + esc(s) + '</span>';
}

function statusPill(val, kind) {
  const s = String(val || '').trim() || '—';
  const lower = s.toLowerCase();
  let cls = 'mrd-pill';
  if (kind === 'sdd') {
    if (lower === 'submitted') cls += ' mrd-pill--ok';
    else if (lower === 'draft') cls += ' mrd-pill--warn';
  } else if (kind === 'grv') {
    if (lower === 'open') cls += ' mrd-pill--high';
    else if (lower === 'closed') cls += ' mrd-pill--ok';
  } else if (kind === 'eudr') {
    if (lower === 'potential') cls += ' mrd-pill--ok';
  }
  return '<span class="' + cls + '">' + esc(s) + '</span>';
}

function hasCellValue(val) {
  if (val === undefined || val === null) return false;
  const s = String(val).trim();
  return !!(s && s !== '—' && s !== '-');
}

function isZeroish(val) {
  const s = String(val == null ? '' : val).trim();
  return s === '0' || s === '0%' || s === '—' || s === '';
}

function filterColumns(columns, rows) {
  return columns.filter(function(col) {
    if (col.always) return true;
    return rows.some(function(row) {
      if (col.hasData) return col.hasData(row);
      if (col.raw) return hasCellValue(col.raw(row));
      return hasCellValue(row[col.key]);
    });
  });
}

function renderSmartTable(columns, rows, opts) {
  opts = opts || {};
  if (!rows.length && !opts.alwaysShowTable) return opts.empty || '<p class="mrd-empty">No data.</p>';
  const active = filterColumns(columns, rows);
  const colCount = active.length;
  let html = (opts.note || '') + '<div class="table-scroll mrd-table-scroll"><table class="mrd-table' + (opts.tableClass ? ' ' + opts.tableClass : '') + '"><thead><tr>';
  active.forEach(function(col) {
    const titleAttr = col.title ? ' title="' + esc(col.title) + '"' : '';
    html += '<th' + (col.thCls ? ' class="' + col.thCls + '"' : '') + titleAttr + '>' + esc(col.label) + '</th>';
  });
  html += '</tr></thead><tbody>';
  if (!rows.length && opts.alwaysShowTable) {
    html += '<tr><td colspan="' + colCount + '" class="mrd-empty-row">' + esc(opts.emptyRowText || 'No data for this period.') + '</td></tr>';
  }
  rows.forEach(function(row, idx) {
    if (row._before) html += row._before;
    html += '<tr' + (row._trClass ? ' class="' + row._trClass + '"' : '') + '>';
    active.forEach(function(col) {
      html += '<td' + (col.tdCls ? ' class="' + col.tdCls + '"' : '') + '>';
      html += col.render ? col.render(row, idx, colCount) : esc(col.raw ? col.raw(row) : (row[col.key] || '—'));
      html += '</td>';
    });
    html += '</tr>';
    if (row._after) {
      html += row._after.replace(/colspan="99"/g, 'colspan="' + colCount + '"');
    }
  });
  html += '</tbody></table></div>';
  return html;
}

function sectionHtml(id, title, desc, bodyHtml, num) {
  const expanded = !_expanded.has(id + ':closed');
  return ''
    + '<div class="table-card mrd-section mrd-section--' + esc(id) + (expanded ? ' is-open' : '') + '" data-mrd-section="' + esc(id) + '">'
    + '<div class="table-header mrd-section-head">'
    + '<button type="button" class="mrd-section-head-btn" data-mrd-toggle="' + esc(id) + '" aria-expanded="' + (expanded ? 'true' : 'false') + '">'
    + '<span class="mrd-section-badge">' + esc(num) + '</span>'
    + '<div class="mrd-section-head-text"><h3>' + esc(title) + '</h3><p>' + esc(desc) + '</p></div>'
    + '<span class="mrd-section-chev" aria-hidden="true"></span>'
    + '</button></div>'
    + '<div class="mrd-section-body">' + bodyHtml + '</div></div>';
}

function flatSectionHtml(id, title, desc, bodyHtml, num) {
  return ''
    + '<div class="table-card mrd-section mrd-section--flat mrd-section--' + esc(id) + ' is-open" data-mrd-section="' + esc(id) + '">'
    + '<div class="mrd-section-head mrd-section-head--flat">'
    + '<span class="mrd-section-badge">' + esc(num) + '</span>'
    + '<div class="mrd-section-head-text"><h3>' + esc(title) + '</h3><p>' + esc(desc) + '</p></div>'
    + '</div>'
    + '<div class="mrd-section-body">' + bodyHtml + '</div></div>';
}

function limitNote(total, limit) {
  if (total <= limit) return '';
  return '<p class="mrd-limit-note">Showing ' + limit + ' of ' + total + ' rows — use search to narrow results.</p>';
}

function filterForExport_(list, searchFn) {
  return (list || []).filter(searchFn);
}

async function resolveAllNblForExport_(mills) {
  const lists = await _deps.ensureNblLists();
  const out = [];
  for (let i = 0; i < mills.length; i++) {
    const item = mills[i];
    const copy = Object.assign({}, item, { row: item.row });
    if (millNeedsNblRiserResolve_(copy)) {
      const info = _deps.resolveNblBy(copy.row, lists);
      copy.nblBy = info.label || '';
      copy.nblMatches = info.matches || [];
      _nblByCache.set(copy.cacheKey, info);
    }
    out.push(copy);
  }
  return out;
}

function mrdExportDefaultSections_() {
  return MRD_PDF_SECTIONS.map(function(s) { return s.id; }).filter(function(id) {
    if (id === 'kpi') return true;
    return !_expanded.has(id + ':closed');
  });
}

function mrdMountExportModal_() {
  const modal = document.getElementById('mrd-export-modal');
  if (!modal) return;
  if (modal.parentElement !== document.body) {
    document.body.appendChild(modal);
  }
}

function mrdRefreshExportSectionList_() {
  const listEl = document.getElementById('mrdExportSectionList');
  if (!listEl) return;
  const defaults = new Set(mrdExportDefaultSections_());
  listEl.innerHTML = MRD_PDF_SECTIONS.map(function(sec) {
    const checked = defaults.has(sec.id) ? ' checked' : '';
    return '<label class="pf-export-facility-item mrd-export-section-item">'
      + '<input type="checkbox" name="mrdExportSection" value="' + esc(sec.id) + '"' + checked + ' />'
      + '<span>' + esc(sec.label) + '</span>'
      + '</label>';
  }).join('');
}

function mrdSyncExportPeriodFromPage_() {
  const report = getReportPeriod_();
  const yearEl = document.getElementById('mrdExportYearSel');
  const monthSel = document.getElementById('mrdExportMonthSel');
  if (yearEl) yearEl.value = report.year || String(new Date().getFullYear());
  if (monthSel) {
    const m = report.month || String(new Date().getMonth() + 1);
    monthSel.value = m;
  }
}

function mrdExportReportPeriod_() {
  const monthSel = document.getElementById('mrdExportMonthSel');
  return {
    year: mrdReadExportYear_(),
    month: monthSel ? monthSel.value : _month,
  };
}

function mrdOpenExportModal_() {
  if (_mrdExportInFlight) {
    if (typeof window.showSddToast === 'function') {
      window.showSddToast('PDF export in progress — please wait.', 'info');
    }
    return;
  }
  if (!_snapshot) {
    alert('Data not loaded yet — wait a moment and try again.');
    return;
  }
  if (!_deps.getJsPDF) {
    alert('PDF library is not ready.');
    return;
  }
  const modal = document.getElementById('mrd-export-modal');
  if (!modal) return;
  mrdSyncExportPeriodFromPage_();
  mrdRefreshExportSectionList_();
  modal.style.display = 'flex';
  modal.setAttribute('aria-hidden', 'false');
}

function mrdCloseExportModal_() {
  const modal = document.getElementById('mrd-export-modal');
  if (!modal) return;
  modal.style.display = 'none';
  modal.setAttribute('aria-hidden', 'true');
}

function mrdExportSelectedSections_() {
  const boxes = document.querySelectorAll('input[name="mrdExportSection"]:checked');
  const sections = Array.from(boxes).map(function(el) { return el.value; });
  if (!sections.length) return null;
  const order = MRD_PDF_SECTIONS.map(function(s) { return s.id; });
  return order.filter(function(id) { return sections.indexOf(id) !== -1; });
}

function mrdSetExportButtonsBusy_(busy, activeBtn, prevTxt, label) {
  const ids = ['mrdExportConfirm', 'mrdBtnExport'];
  ids.forEach(function(id) {
    const el = document.getElementById(id);
    if (!el) return;
    if (busy) el.disabled = true;
    else el.disabled = false;
  });
  if (busy && activeBtn) {
    dashSetButtonBusy_(activeBtn, label || 'Loading…');
  } else if (!busy && activeBtn) {
    const fallback = prevTxt || (activeBtn.id === 'mrdExportConfirm' ? 'Generate 2 PDFs' : 'Export PDF');
    dashClearButtonBusy_(activeBtn, fallback);
  }
}

async function exportMonthlyReport_(exportOpts) {
  exportOpts = exportOpts || {};
  if (_mrdExportInFlight) return;
  if (!_snapshot) {
    alert('Data not loaded yet — wait a moment and try again.');
    return;
  }
  if (!_deps.getJsPDF) {
    alert('PDF library is not ready.');
    return;
  }

  const sections = exportOpts.sections;
  if (!sections || !sections.length) {
    alert('Select at least one section to export.');
    return;
  }

  const reportPeriod = exportOpts.reportPeriod || getReportPeriod_();
  if (!reportPeriod.month) {
    alert('Select a reporting month for the PDF header.');
    return;
  }

  const btn = document.getElementById('mrdExportConfirm') || document.getElementById('mrdBtnExport');
  const prevTxt = btn ? btn.textContent : '';
  _mrdExportInFlight = true;
  mrdSetExportButtonsBusy_(true, btn, prevTxt, 'Loading…');

  try {
    _year = String(reportPeriod.year || _year);
    _month = String(reportPeriod.month != null ? reportPeriod.month : _month);
    const yearSel = document.getElementById('mrdYearSel');
    const monthSel = document.getElementById('mrdMonthSel');
    if (yearSel && _year) yearSel.value = _year;
    if (monthSel) monthSel.value = _month;
    const pageDataPeriod = { year: _year, month: _month };
    const facilityPeriod = getFacilityReportContext_();

    if (sections.includes('sdd') && _sddCache.length === 0 && _deps.fetchSddList) {
      try {
        const rows = await withTimeout(_deps.fetchSddList(), 25000, 'SDD');
        _sddCache = Array.isArray(rows) ? rows : [];
      } catch (_) { /* continue without SDD */ }
    }

    if (btn) dashSetButtonBusy_(btn, 'Loading…');
    const extra = await _deps.preparePdfExport({
      sections: sections,
      dataPeriod: pageDataPeriod,
      facilityPeriod: facilityPeriod,
    });

    _snapshot = rebuildSnapshot_({ reportPeriod: pageDataPeriod, sddRows: _sddCache, sddLoading: false });
    await resolveNblMillsForSnapshot_();

    const prevEudr = (_snapshot && _snapshot.eudrPotential) || [];
    const extraEudr = (extra && extra.eudr) || [];
    let bestEudr = prevEudr.length >= extraEudr.length ? prevEudr.slice() : extraEudr.slice();
    if (sections.includes('eudr') && _deps.fetchEudrPotential) {
      try {
        const fetched = await withTimeout(_deps.fetchEudrPotential(), 120000, 'EUDR');
        if (Array.isArray(fetched) && fetched.length) bestEudr = fetched;
      } catch (_) { /* keep bestEudr */ }
    } else if (extraEudr.length > bestEudr.length) {
      bestEudr = extraEudr.slice();
    }
    _snapshot = rebuildSnapshot_({
      reportPeriod: pageDataPeriod,
      eudrPotential: bestEudr.length ? bestEudr : undefined,
    });
    const s = _snapshot;

    if (btn) dashSetButtonBusy_(btn, 'Building PDF…');
    const exportSections = mrdPdfSectionsNoDupHighRisk_(sections.slice());

    // Mill Onboarding PDF export: HIGH RISK mills only (same Result Risk Level as website).
    const allMillsResolved = await resolveAllNblForExport_(mrdSortMillItems_(s.mills || []));
    const highRiskMills = mrdSortMillItems_(allMillsResolved.filter(function(item) {
      return mrdIsHighRiskItem_(item) && mrdShowInMillOnboarding_(item);
    }));
    const millsForPdf = highRiskMills;
    const nblMills = mrdSortMillItems_(allMillsResolved.filter(function(item) {
      return isNblYes(item.nbl) && matchesSearch(item.search);
    }));

    let facilityBundles = [];
    if (sections.includes('facility')) {
      if (_deps.loadFacilityBundlesForReport) {
        facilityBundles = await withTimeout(
          _deps.loadFacilityBundlesForReport(facilityPeriod),
          120000,
          'Facility performance'
        );
      } else if (_deps.getFacilityBundles) {
        if (_deps.preparePfDataForReport) {
          await withTimeout(_deps.preparePfDataForReport(facilityPeriod), 120000, 'Facility performance');
        }
        facilityBundles = _deps.getFacilityBundles(facilityPeriod) || [];
      }
    }
    if (_search) {
      const q = _search;
      facilityBundles = facilityBundles.filter(function(b) {
        const blob = [b.facility, (b.companies || []).map(function(c) { return c.company; }).join(' ')].join(' ').toLowerCase();
        return blob.includes(q);
      });
    }

    // Export always includes every EUDR potential row (no search filter).
    const eudrList = s.eudrPotential || [];
    const exportStats = Object.assign({}, s.stats, {
      highRisk: highRiskMills.length,
      nblMills: nblMills.length,
      eudrPotential: eudrList.length,
    });

    await buildMonthlyReportPdfPair_({
      getJsPDF: _deps.getJsPDF,
      year: reportPeriod.year,
      month: reportPeriod.month,
      dataYear: pageDataPeriod.year,
      dataMonth: pageDataPeriod.month,
      sections: exportSections,
      data: {
        stats: exportStats,
        facility: s.facility,
        sdd: mrdSortSddRows_(filterForExport_(s.sdd, function(r) {
          return matchesSearch([
            r['SCR - Screening Status'], r['Group Name'], r['Grup Name'], r['Mill Name'],
            r['Company Name'], r['Date Imported'], r.supplier_type, r.updated_at,
          ].join(' ').toLowerCase());
        })),
        mills: millsForPdf,
        highRiskMills: highRiskMills,
        emptyMills: s.emptyMills || [],
        traceRows: s.traceRows || [],
        traceTotals: s.traceTotals || {},
        grv: mrdSortGrvItemsByDateDesc_(filterForExport_(s.grv, function(item) { return matchesSearch(item.search); })),
        nblAll: nblMills,
        facilityBundles: mrdSortFacilityBundles_(facilityBundles),
        eudrPotential: mrdSortEudrItems_(eudrList),
      },
    });

    if (typeof window.showSddToast === 'function') {
      window.showSddToast('2 PDF files downloaded successfully.', 'success');
    }
  } catch (err) {
    console.error('[MRD PDF]', err);
    alert('PDF export failed: ' + (err && err.message ? err.message : String(err)));
  } finally {
    _mrdExportInFlight = false;
    mrdSetExportButtonsBusy_(false, btn, prevTxt);
  }
}

function mrdBindExportModalOnce_() {
  mrdMountExportModal_();

  const closeIds = ['mrdExportModalClose', 'mrdExportCancel', 'mrdExportModalBackdrop'];
  closeIds.forEach(function(id) {
    const el = document.getElementById(id);
    if (el && !el._mrdExportBound) {
      el._mrdExportBound = true;
      el.addEventListener('click', mrdCloseExportModal_);
    }
  });

  const selectAll = document.getElementById('mrdExportSelectAll');
  if (selectAll && !selectAll._mrdExportBound) {
    selectAll._mrdExportBound = true;
    selectAll.addEventListener('click', function() {
      document.querySelectorAll('input[name="mrdExportSection"]').forEach(function(el) { el.checked = true; });
    });
  }

  const selectNone = document.getElementById('mrdExportSelectNone');
  if (selectNone && !selectNone._mrdExportBound) {
    selectNone._mrdExportBound = true;
    selectNone.addEventListener('click', function() {
      document.querySelectorAll('input[name="mrdExportSection"]').forEach(function(el) { el.checked = false; });
    });
  }

  const confirm = document.getElementById('mrdExportConfirm');
  if (confirm && !confirm._mrdExportBound) {
    confirm._mrdExportBound = true;
    confirm.addEventListener('click', async function() {
      if (_mrdExportInFlight) return;
      const sections = mrdExportSelectedSections_();
      if (!sections) {
        alert('Select at least one section to export.');
        return;
      }
      const reportPeriod = mrdExportReportPeriod_();
      if (!reportPeriod.month) {
        alert('Select a reporting month for the PDF header.');
        return;
      }
      mrdCloseExportModal_();
      await exportMonthlyReport_({ sections: sections, reportPeriod: reportPeriod });
    });
  }
}

function buildSnapshotSync(opts) {
  opts = opts || {};
  const reportPeriod = mrdResolveReportPeriod_(opts.reportPeriod);
  const periodYear = reportPeriod.year;
  const periodMonth = reportPeriod.month;
  const millData = (_deps.getMillData() || []);
  const ttpData = (_deps.getTtpData() || []);
  const ttpFields = _deps.getTtpFields() || [];
  const grvData = (_deps.getGrvData() || []);
  const nblReg = (_deps.getNblRegistry() || []);
  const nblUni = (_deps.getNblUnilever() || []);
  const sddInput = opts.sddRows != null ? opts.sddRows : [];
  const eudrInput = opts.eudrPotential != null ? opts.eudrPotential : [];

  const millYear = _deps.millYearVal;
  const millMonth = _deps.millMonthVal;

  function filterMillsByPeriod_(targetYear, targetMonth) {
    return millData.filter(function(r) {
      const y = parseYear(millYear(r));
      if (targetYear && y && y !== String(targetYear)) return false;
      if (targetMonth && millMonth) {
        const m = millMonth(r);
        if (m && !mrdMonthMatchesFilter_(m, targetMonth)) return false;
      }
      return true;
    });
  }

  // Mill Onboarding: Main (CPO/PK) + Waste (POME/SHELL) always both — no product picker.
  // Rows stay separate (not General-merged) so export/list show both product types.
  let mills;
  if (_deps.getMillsForReportPeriod && periodYear) {
    const tagRows_ = function(rows, product) {
      return (rows || []).map(function(r) {
        const copy = Object.assign({}, r);
        copy._mrdProduct = product;
        return copy;
      });
    };
    mills = tagRows_(_deps.getMillsForReportPeriod(periodYear, periodMonth, 'main'), 'main')
      .concat(tagRows_(_deps.getMillsForReportPeriod(periodYear, periodMonth, 'waste'), 'waste'));
  } else {
    mills = filterMillsByPeriod_(periodYear, periodMonth);
  }
  const millEffectiveYear = periodYear;
  const millEffectiveMonth = periodMonth;

  const sddFiltered = sddInput.filter(function(r) {
    const st = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
    if (st !== 'draft' && st !== 'submitted') return false;
    if (periodYear) {
      const upd = String(r.updated_at || r['SCR - Last Updated'] || '').slice(0, 4);
      if (upd && upd !== periodYear) return false;
    }
    if (periodMonth) {
      const m = parseMonthFromDate(r.updated_at || r['SCR - Last Updated']);
      if (m && m !== periodMonth) return false;
    }
    return true;
  });

  const millCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'MILL NAME'; }) || 'MILL NAME';
  const groupCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'GROUP NAME'; }) || 'GROUP NAME';
  const companyCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'COMPANY NAME'; }) || 'COMPANY NAME';
  const supplierCol = mrdResolveTtpSupplierCol_(ttpFields);
  const yearCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'YEAR'; }) || 'YEAR';

  // Traceability sheet: year only — always report year − 1 (e.g. Jan 2026 report → 2025 TTP).
  const traceYear = mrdTraceYearFromReport_(periodYear);
  const ttpFiltered = ttpData.filter(function(r) {
    const y = parseYear(r[yearCol] || millYear(r));
    return !traceYear || !y || y === traceYear;
  });

  const millKeys = new Map();
  mills.forEach(function(r) {
    const key = [r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']].map(function(x) {
      return String(x || '').trim().toLowerCase();
    }).join('|');
    if (!millKeys.has(key)) millKeys.set(key, r);
  });

  const ttpMillMaps = mrdBuildTtpByMillMaps_(ttpFiltered, millCol, groupCol, companyCol);

  const ttpByMill = ttpMillMaps.byMillName;

  const emptyMills = [];
  millKeys.forEach(function(millRow) {
    const rows = mrdTtpRowsForMill_(ttpMillMaps, millRow);
    const hasSupplier = rows.some(function(r) {
      return mrdTtpRowHasSupplier_(r, supplierCol);
    });
    if (!hasSupplier) emptyMills.push({ millRow: millRow, rows: rows });
  });
  const emptyMillsSorted = mrdSortEmptyMillItems_(emptyMills);

  const millRows = mills.map(function(r) {
    const product = mrdMillProductLabel_(r);
    const cacheKey = [r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME'], product].join('|');
    const cached = _nblByCache.get(cacheKey);
    return {
      row: r,
      product: product,
      cacheKey: cacheKey,
      risk: _deps.millResolvedRiskLevel(r),
      nbl: r['BUYER NO BUY LIST'],
      nblBy: cached ? cached.label : '',
      nblMatches: cached ? cached.matches : null,
      search: [
        r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME'], r['PROVINCE'],
        product, r['RESULT RISK LEVEL'], r['BUYER NO BUY LIST'],
        r['CERTIFICATION'], r['SUPPLIER STATUS'],
      ].join(' ').toLowerCase(),
    };
  });

  // Grievance: Date Received year = reporting year (not lagged data period).
  const grvRows = mrdSortGrvItemsByDateDesc_(grvData.filter(function(r) {
    const dateRaw = r['Date Received'] || r['DATE RECEIVED'] || r['Date received'] || '';
    if (periodYear) {
      const yr = parseYearFromDate_(dateRaw);
      if (!yr || yr !== String(periodYear)) return false;
    }
    return true;
  }).map(function(r) {
    return { row: r, search: Object.keys(r).map(function(k) { return r[k]; }).join(' ').toLowerCase() };
  }));

  const nblAll = [];
  nblReg.forEach(function(r) {
    nblAll.push({
      source: 'NBL Registry',
      riser: r._nblRiser || r.Riser || '',
      group: r._nblGroup || r['Group Name NBL'] || '',
      company: r._nblCompany || r['Company Name NBL'] || '',
      search: [r._nblRiser, r._nblGroup, r._nblCompany].join(' ').toLowerCase(),
    });
  });
  nblUni.forEach(function(r) {
    nblAll.push({
      source: 'Unilever NBL',
      riser: r._nblRiser || 'Unilever',
      group: '',
      company: r._nblCompany || r['Company Name NBL'] || '',
      search: [r._nblCompany, r._nblRiser].join(' ').toLowerCase(),
    });
  });

  const facility = _deps.buildFacilitySummary(mills, ttpFiltered);
  const traceTotals = _deps.buildTraceTotals
    ? _deps.buildTraceTotals(traceYear, periodYear, periodMonth)
    : {};

  const highRiskMillRows = millRows.filter(mrdIsHighRiskItem_);

  return {
    reportPeriod: { year: String(periodYear || ''), month: String(periodMonth || '') },
    sdd: mrdSortSddRows_(sddFiltered),
    mills: mrdSortMillItems_(millRows),
    highRiskMills: mrdSortMillItems_(highRiskMillRows),
    millsTotal: millRows.length,
    millEffectiveYear: millEffectiveYear,
    millEffectiveMonth: millEffectiveMonth,
    emptyMills: emptyMillsSorted,
    traceRows: mrdSortMillItems_(_deps.buildTraceRows ? _deps.buildTraceRows(mills, ttpFiltered, ttpByMill, supplierCol) : []),
    traceTotals: traceTotals,
    traceYear: traceYear,
    ttpByMill: ttpByMill,
    grv: grvRows,
    nblAll: nblAll,
    facility: facility,
    eudrPotential: mrdSortEudrItems_(eudrInput),
    eudrLoading: !!opts.eudrLoading,
    sddLoading: !!opts.sddLoading,
    facilityBundles: opts.facilityBundles != null ? opts.facilityBundles : _facilityBundles,
    facilityLoading: !!opts.facilityLoading,
    stats: {
      sddTotal: sddFiltered.length,
      sddRequested: sddFiltered.length,
      sddDraft: sddFiltered.filter(function(r) {
        return String(r['SCR - Screening Status'] || '').toLowerCase() === 'draft';
      }).length,
      sddSubmitted: sddFiltered.filter(function(r) {
        return String(r['SCR - Screening Status'] || '').toLowerCase() === 'submitted';
      }).length,
      sddDone: sddFiltered.filter(function(r) {
        return String(r['SCR - Screening Status'] || '').toLowerCase() === 'submitted';
      }).length,
      totalMills: mills.length,
      totalGroups: new Set(mills.map(function(r) { return r['GROUP NAME']; }).filter(Boolean)).size,
      highRisk: mills.filter(function(r) { return mrdIsHighRiskItem_(r); }).length,
      nblMills: mills.filter(function(r) { return isNblYes(r['BUYER NO BUY LIST']); }).length,
      emptyTraceMills: emptyMills.length,
      grievances: grvRows.length,
      nblEntries: nblAll.length,
      eudrPotential: eudrInput.length,
      facilities: facility.cpo.length + facility.pk.length,
      ttmCpoPct: traceTotals.ttmCpoFmt || '—',
      ttmPkPct: traceTotals.ttmPkFmt || '—',
      ttpCpoPct: traceTotals.ttpCpoFmt || '—',
      ttpPkPct: traceTotals.ttpPkFmt || '—',
    },
  };
}

function renderKpis(stats, opts) {
  opts = opts || {};
  const el = document.getElementById('mrdKpiRow');
  if (!el) return;
  const tracePending = !!opts.tracePending;
  const eudrPending = !!opts.eudrPending;
  const untraceN = tracePending ? '…' : stats.emptyTraceMills;
  const eudrN = eudrPending ? '…' : stats.eudrPotential;
  const items = _viewMode === 'executive'
    ? [
      { n: stats.sddRequested != null ? stats.sddRequested : stats.sddTotal, l: 'SDD Requested', s: (stats.sddDone != null ? stats.sddDone : stats.sddSubmitted) + ' submitted' },
      { n: stats.totalMills, l: 'Total Mills', s: stats.totalGroups + ' groups' },
      { n: stats.highRisk, l: 'High Risk', s: 'Result Risk Level = HIGH', hot: stats.highRisk > 0 },
      { n: stats.nblMills, l: 'NBL Mills', s: stats.nblEntries + ' registry entries', hot: stats.nblMills > 0 },
      { n: untraceN, l: 'Untraceable', s: tracePending ? 'loading…' : 'no supplier data', hot: !tracePending && stats.emptyTraceMills > 0 },
      { n: stats.grievances, l: 'Grievances', s: 'Date Received · report year' },
      { n: eudrN, l: 'EUDR Potential', s: eudrPending ? 'loading…' : 'by formula', hot: !eudrPending && stats.eudrPotential > 0 },
      { n: stats.ttmCpoPct || '—', l: 'TTM CPO', s: 'TTM PK ' + (stats.ttmPkPct || '—') },
    ]
    : [
      { n: stats.sddRequested != null ? stats.sddRequested : stats.sddTotal, l: 'SDD Requested', s: (stats.sddDone != null ? stats.sddDone : stats.sddSubmitted) + ' done' },
      { n: stats.totalMills, l: 'Total Mills', s: stats.totalGroups + ' groups' },
      {
        n: untraceN,
        l: 'Untraceable Mills',
        s: tracePending ? 'loading traceability…' : 'mills without suppliers',
        hot: !tracePending && stats.emptyTraceMills > 0,
      },
      {
        n: eudrN,
        l: 'EUDR Potential',
        s: eudrPending ? 'loading EUDR…' : 'by formula',
        hot: !eudrPending && stats.eudrPotential > 0,
      },
    ];
  el.innerHTML = items.map(function(it) {
    return '<div class="stat-card' + (it.hot ? ' mrd-stat-hot' : '') + '">'
      + '<div class="stat-card-num">' + esc(it.n) + '</div>'
      + '<div class="stat-card-label">' + esc(it.l) + '</div>'
      + '<div class="mrd-stat-sub">' + esc(it.s) + '</div></div>';
  }).join('');
}

async function mrdEnsureChartModule_() {
  if (_mrdExecChartModule) return _mrdExecChartModule;
  const mod = await import('chart.js/auto');
  _mrdExecChartModule = mod.default || mod.Chart || mod;
  return _mrdExecChartModule;
}

function mrdExecutiveChartEls_() {
  return {
    qoq: document.getElementById('mrdExecChartQoq'),
    trend: document.getElementById('mrdExecChartTrend'),
    supply: document.getElementById('mrdExecChartSupply'),
    sdd: document.getElementById('mrdExecChartSdd'),
    risk: document.getElementById('mrdExecChartRisk'),
    nbl: document.getElementById('mrdExecChartNbl'),
    traceGap: document.getElementById('mrdExecChartTraceGap'),
    trace: document.getElementById('mrdExecChartTrace'),
    sddCat: document.getElementById('mrdExecChartSddCat'),
    grv: document.getElementById('mrdExecChartGrv'),
  };
}

function mrdExecQtyHelpers_() {
  return {
    cpo: function(r) {
      return _deps && _deps.millSupplyCpoQty ? _deps.millSupplyCpoQty(r) : 0;
    },
    pk: function(r) {
      return _deps && _deps.millSupplyPkQty ? _deps.millSupplyPkQty(r) : 0;
    },
    pomeIscc: function(r) {
      return _deps && _deps.millWasteSupplyQty ? _deps.millWasteSupplyQty(r, 'ISCC') : 0;
    },
    pomeIns: function(r) {
      return _deps && _deps.millWasteSupplyQty ? _deps.millWasteSupplyQty(r, 'INS') : 0;
    },
    shell: function(r) {
      return _deps && _deps.millWasteSupplyQty ? _deps.millWasteSupplyQty(r, 'SHELL') : 0;
    },
  };
}

function mrdExecKpiOpts_() {
  return {
    entityKey: _deps && _deps.millEntityKey
      ? _deps.millEntityKey
      : function(r) {
        return String(r['MILL NAME'] || r['COMPANY NAME'] || '').trim().toUpperCase();
      },
    resolveRisk: function(r) {
      return _deps && _deps.millResolvedRiskLevel
        ? (_deps.millResolvedRiskLevel(r) || '')
        : String(r['RESULT RISK LEVEL'] || r['RISK LEVEL'] || '').trim();
    },
    isNbl: function(r) {
      return isNblYes(r['BUYER NO BUY LIST']);
    },
    pickGroup: function(r) {
      return _deps && _deps.pickMillGroupName
        ? _deps.pickMillGroupName(r)
        : String(r['GROUP NAME'] || '').trim();
    },
  };
}

function mrdCollectMillsForPeriod_(year, month) {
  if (!_deps || !_deps.getMillsForReportPeriod || !year) return [];
  const main = (_deps.getMillsForReportPeriod(year, month, 'main') || []).map(function(r) {
    const copy = Object.assign({}, r);
    copy._mrdProduct = 'main';
    return copy;
  });
  const waste = (_deps.getMillsForReportPeriod(year, month, 'waste') || []).map(function(r) {
    const copy = Object.assign({}, r);
    copy._mrdProduct = 'waste';
    return copy;
  });
  return main.concat(waste);
}

/** Exact month+year rows from sheet (no as-of) — for supply quantity totals. */
function mrdCollectExactMonthRows_(year, month) {
  const y = String(year || '');
  const m = parseInt(String(month || ''), 10);
  if (!y || m < 1 || m > 12) return [];
  const millYear = _deps.millYearVal;
  const millMonth = _deps.millMonthVal;
  const mainSrc = (_deps.getMillRawRows ? _deps.getMillRawRows() : (_deps.getMillData ? _deps.getMillData() : [])) || [];
  const wasteSrc = (_deps.getMillWasteRawRows ? _deps.getMillWasteRawRows() : []) || [];
  function match_(r) {
    if (!r || !mrdShowInMillOnboarding_(r)) return false;
    const ry = parseYear(millYear(r));
    if (ry && ry !== y) return false;
    const rm = millMonth(r);
    return mrdMonthMatchesFilter_(rm, String(m));
  }
  return mainSrc.filter(match_).concat(wasteSrc.filter(match_));
}

function mrdBuildExecutiveExtras_() {
  const report = getReportPeriod_();
  const year = parseInt(String(report.year || ''), 10);
  const month = parseInt(String(report.month || ''), 10);
  const kpiOpts = mrdExecKpiOpts_();
  const qtyHelpers = mrdExecQtyHelpers_();
  const monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  // Supply qty: exact MONTH + YEAR from sheet columns (not as-of carry-forward).
  let supplyRows = [];
  let supplyPeriodLabel = 'selected period';
  if (year && month >= 1 && month <= 12) {
    supplyRows = mrdCollectExactMonthRows_(year, month);
    // Fall back to latest available month if selected month has no data
    if (!supplyRows.length) {
      const fallbackM = mrdFindLatestDataMonthForYear_(year);
      if (fallbackM && fallbackM !== month) {
        supplyRows = mrdCollectExactMonthRows_(year, fallbackM);
        supplyPeriodLabel = (monthNames[fallbackM] || fallbackM) + ' ' + year + ' (latest available)';
      } else {
        supplyPeriodLabel = (monthNames[month] || month) + ' ' + year + ' (no data)';
      }
    } else {
      supplyPeriodLabel = (monthNames[month] || month) + ' ' + year + ' (exact month)';
    }
  } else if (year) {
    // Full year: sum all months in that year from sheet
    const mainSrc = (_deps.getMillRawRows ? _deps.getMillRawRows() : (_deps.getMillData ? _deps.getMillData() : [])) || [];
    const wasteSrc = (_deps.getMillWasteRawRows ? _deps.getMillWasteRawRows() : []) || [];
    const y = String(year);
    function yearMatch_(r) {
      if (!r || !mrdShowInMillOnboarding_(r)) return false;
      const ry = parseYear(_deps.millYearVal(r));
      return !ry || ry === y;
    }
    supplyRows = mainSrc.filter(yearMatch_).concat(wasteSrc.filter(yearMatch_));
    supplyPeriodLabel = 'Full year ' + year + ' (all months)';
  }

  const supplyKpis = buildMillPeriodKpis_(supplyRows, Object.assign({}, kpiOpts, { qtyHelpers: qtyHelpers }));

  // Quarter comparison: use selected month's quarter but cap to latest quarter with actual data.
  const latestDataMonth = mrdFindLatestDataMonthForYear_(year);
  const latestDataQuarter = latestDataMonth ? monthToQuarter_(latestDataMonth) : 4;
  const rawQuarter = month >= 1 && month <= 12 ? monthToQuarter_(month) : (year ? 4 : 0);
  const activeQuarter = Math.min(rawQuarter, latestDataQuarter);
  let quarterComparison = null;
  const quarterlyTrend = [];

  if (year && activeQuarter >= 1 && activeQuarter <= 4) {
    const curEnd = quarterEndMonth_(activeQuarter);
    const curRows = mrdCollectMillsForPeriod_(year, curEnd);
    const curKpis = buildMillPeriodKpis_(curRows, kpiOpts);

    const prev = previousQuarter_(year, activeQuarter);
    let prevKpis = { totalMills: 0, groupCount: 0, highRisk: 0, nbl: 0, entityKeys: new Set() };
    if (prev) {
      const prevEnd = quarterEndMonth_(prev.quarter);
      const prevRows = mrdCollectMillsForPeriod_(prev.year, prevEnd);
      prevKpis = buildMillPeriodKpis_(prevRows, kpiOpts);
    }

    quarterComparison = buildQuarterComparison_(curKpis, prevKpis, {
      currentLabel: 'Q' + activeQuarter + ' ' + year,
      previousLabel: prev ? ('Q' + prev.quarter + ' ' + prev.year) : '—',
    });

    [1, 2, 3, 4].forEach(function(q) {
      const endM = quarterEndMonth_(q);
      const rows = mrdCollectMillsForPeriod_(year, endM);
      const snap = buildMillPeriodKpis_(rows, kpiOpts);
      quarterlyTrend.push({
        label: 'Q' + q,
        totalMills: snap.totalMills,
        active: q === activeQuarter,
      });
    });
  }

  return {
    supply: supplyKpis.supply,
    supplyPeriodLabel: supplyPeriodLabel,
    quarterComparison: quarterComparison,
    quarterlyTrend: quarterlyTrend,
    activeQuarter: activeQuarter,
    year: year,
    month: month,
  };
}

function syncMrdViewModeUi_() {
  document.querySelectorAll('[data-mrd-view]').forEach(function(btn) {
    btn.classList.toggle('active', btn.getAttribute('data-mrd-view') === _viewMode);
  });
  const panel = document.getElementById('panel-monthly-report-detail');
  if (panel) {
    panel.classList.toggle('mrd-panel--executive', _viewMode === 'executive');
    panel.classList.toggle('mrd-panel--detail', _viewMode === 'detail');
  }
  const execBtn = document.getElementById('mrdBtnExecExport');
  const detailBtn = document.getElementById('mrdBtnExport');
  if (execBtn) execBtn.hidden = _viewMode !== 'executive';
  if (detailBtn) detailBtn.classList.toggle('mrd-export-btn--secondary', _viewMode === 'executive');
  const searchWrap = document.querySelector('#panel-monthly-report-detail .mrd-toolbar-search-wrap');
  if (searchWrap) searchWrap.hidden = _viewMode === 'executive';
}

async function renderMrdExecutiveView_() {
  if (!_snapshot) return;
  const execEl = document.getElementById('mrdExecutiveView');
  if (!execEl) return;

  const extras = mrdBuildExecutiveExtras_();
  const data = buildMrdExecutiveData_(_snapshot, extras);
  _mrdExecDataCache = data;

  const qoqHint = document.getElementById('mrdExecQoqHint');
  const qoqCards = document.getElementById('mrdExecQoqCards');
  if (qoqHint) {
    if (data.quarterComparison) {
      const qc = data.quarterComparison;
      qoqHint.textContent = qc.previousLabel + ' → ' + qc.currentLabel
        + ' · as-of end of quarter · unique mills (excl. Trader/Refinery)';
    } else {
      qoqHint.textContent = 'Select a month to derive the active quarter (Q1–Q4).';
    }
  }
  if (qoqCards) {
    if (data.quarterComparison) {
      const qc = data.quarterComparison;
      const d = qc.delta;
      const cards = [
        { n: qc.current.totalMills, l: 'Mills · ' + qc.currentLabel, s: formatDelta_(d.totalMills) + ' vs ' + qc.previousLabel, hot: d.totalMills > 0 },
        { n: d.millsAdded, l: 'Mills added', s: 'new vs previous quarter', hot: d.millsAdded > 0 },
        { n: d.millsRemoved, l: 'Mills removed', s: 'no longer in snapshot', hot: d.millsRemoved > 0 },
        { n: qc.current.groupCount, l: 'Groups', s: formatDelta_(d.groupCount) + ' QoQ' },
        { n: qc.current.highRisk, l: 'High Risk', s: formatDelta_(d.highRisk) + ' QoQ', hot: d.highRisk > 0 },
        { n: qc.current.nbl, l: 'NBL Mills', s: formatDelta_(d.nbl) + ' QoQ', hot: d.nbl > 0 },
      ];
      qoqCards.innerHTML = cards.map(function(it) {
        return '<div class="mrd-exec-qoq-card' + (it.hot ? ' is-hot' : '') + '">'
          + '<div class="mrd-exec-qoq-card__num">' + esc(it.n) + '</div>'
          + '<div class="mrd-exec-qoq-card__label">' + esc(it.l) + '</div>'
          + '<div class="mrd-exec-qoq-card__sub">' + esc(it.s) + '</div></div>';
      }).join('');
    } else {
      qoqCards.innerHTML = '<p class="mrd-empty">Choose Year + Month to see quarter growth.</p>';
    }
  }

  const supplyHint = document.getElementById('mrdExecSupplyHint');
  if (supplyHint) {
    supplyHint.textContent = 'Sheet columns SUPPLY CPO / PK / POME ISCC / INS / SHELL · '
      + (data.supplyPeriodLabel || 'selected period') + ' · ton';
  }

  const kpisEl = document.getElementById('mrdExecutiveKpis');
  if (kpisEl) {
    const supply = data.supply || {};
    const items = [
      { n: formatQtyDisplay_(supply.cpo), l: 'CPO (ton)' },
      { n: formatQtyDisplay_(supply.pk), l: 'PK (ton)' },
      { n: formatQtyDisplay_(supply.pomeIscc), l: 'POME ISCC (ton)' },
      { n: formatQtyDisplay_(supply.pomeIns), l: 'POME INS (ton)' },
      { n: formatQtyDisplay_(supply.shell), l: 'SHELL GGL (ton)' },
      { n: (_snapshot.stats && _snapshot.stats.ttmCpoPct) || '—', l: 'TTM CPO' },
      { n: (_snapshot.stats && _snapshot.stats.ttmPkPct) || '—', l: 'TTM PK' },
      { n: (_snapshot.stats && _snapshot.stats.ttpCpoPct) || '—', l: 'TTP CPO' },
    ];
    kpisEl.innerHTML = items.map(function(it) {
      return '<div class="mrd-exec-kpi"><div class="mrd-exec-kpi__num">' + esc(it.n)
        + '</div><div class="mrd-exec-kpi__label">' + esc(it.l) + '</div></div>';
    }).join('');
  }

  const trendYear = document.getElementById('mrdExecTrendYear');
  if (trendYear) {
    trendYear.textContent = extras.year ? '(' + extras.year + ')' : '';
  }

}

const scheduleRenderMrdExecutive_ = (function() {
  let raf = 0;
  return function() {
    if (raf) cancelAnimationFrame(raf);
    raf = requestAnimationFrame(function() {
      raf = 0;
      renderMrdExecutiveView_();
    });
  };
})();

async function exportMrdExecutiveReport_() {
  if (_mrdExecExportInFlight) return;
  _mrdExecExportInFlight = true;
  const btn = document.getElementById('mrdBtnExecExport');
  dashSetButtonBusy_(btn, 'Generating…');
  if (typeof window.showSddToast === 'function') {
    window.showSddToast('Generating charts for PDF…', 'info');
  }
  let offscreenContainer = null;
  try {
    if (!_snapshot) throw new Error('No report data loaded.');
    if (!_mrdExecDataCache) await renderMrdExecutiveView_();
    const data = _mrdExecDataCache;
    const Chart = await mrdEnsureChartModule_();

    // Offscreen canvases — fixed size so Chart.js renders correctly
    const CHART_SIZES = {
      sdd:      [700, 420],
      risk:     [700, 420],
      nbl:      [700, 420],
      traceGap: [700, 420],
      trace:    [800, 420],
      sddCat:   [820, 460],
      grv:      [700, 420],
      qoq:      [900, 440],
      trend:    [900, 380],
      supply:   [900, 440],
    };
    offscreenContainer = document.createElement('div');
    offscreenContainer.style.cssText = 'position:fixed;left:-9999px;top:0;width:800px;visibility:hidden;pointer-events:none;z-index:-1;';
    const offscreenEls = {};
    Object.keys(CHART_SIZES).forEach(function(key) {
      const sz = CHART_SIZES[key];
      const wrap = document.createElement('div');
      wrap.style.cssText = 'width:' + sz[0] + 'px;height:' + sz[1] + 'px;';
      const canvas = document.createElement('canvas');
      canvas.width = sz[0];
      canvas.height = sz[1];
      wrap.appendChild(canvas);
      offscreenContainer.appendChild(wrap);
      offscreenEls[key] = canvas;
    });
    document.body.appendChild(offscreenContainer);

    renderMrdExecutiveCharts_(Chart, data, offscreenEls);

    // Wait 2 frames for Chart.js to finish rendering
    await new Promise(function(resolve) {
      requestAnimationFrame(function() { requestAnimationFrame(resolve); });
    });

    const chartImages = collectMrdExecutiveChartImages_(offscreenEls);
    destroyMrdExecutiveCharts_();

    const report = getReportPeriod_();
    const header = mrdExecutiveHeaderMeta_(report.year, report.month);
    await exportMrdExecutivePdf_({
      periodLine: header.periodLine,
      dataPeriodLine: header.dataPeriodLine,
      cutoffLine: header.cutoffLine,
      filename: mrdExecutiveFilename_(report.year, report.month),
    }, data, chartImages, _deps && _deps.getJsPDF);
    if (typeof window.showSddToast === 'function') {
      window.showSddToast('Executive report PDF downloaded.', 'success');
    }
  } catch (err) {
    console.error('[MRD Executive] PDF export failed:', err);
    if (typeof window.showSddToast === 'function') {
      window.showSddToast(err && err.message ? err.message : 'PDF export failed.', 'error');
    } else {
      alert(err && err.message ? err.message : 'PDF export failed.');
    }
  } finally {
    if (offscreenContainer && offscreenContainer.parentNode) {
      offscreenContainer.parentNode.removeChild(offscreenContainer);
    }
    _mrdExecExportInFlight = false;
    dashClearButtonBusy_(btn, 'Export Executive PDF');
  }
}

function setMrdViewMode_(mode) {
  const m = String(mode || '').trim().toLowerCase();
  _viewMode = m === 'detail' ? 'detail' : 'executive';
  syncMrdViewModeUi_();
  renderAll();
}

function renderDetailSections_() {
  const sections = document.getElementById('mrdSections');
  if (!sections || !_snapshot) return;
  const s = _snapshot;
  const stats = s.stats || {};
  let html = '';
  const reportPeriod = (s.reportPeriod && (s.reportPeriod.year || s.reportPeriod.month))
    ? s.reportPeriod
    : getReportPeriod_();
  const reportLabel = mrdDataPeriodShortLabel_(reportPeriod);
  const traceYear = mrdTraceYearFromReport_(reportPeriod.year);
  const tracePeriodLabel = traceYear ? ('Year ' + traceYear) : (reportPeriod.year ? ('Full year ' + reportPeriod.year) : 'all periods');
  html += flatSectionHtml('sdd', 'Supplier Due Diligence · ' + reportLabel, stats.sddRequested + ' requested · ' + stats.sddDone + ' done', renderSddSection(s.sdd, s.sddLoading), '01');
  html += sectionHtml('highRisk', 'High Risk Suppliers · ' + reportLabel, stats.highRisk + ' mills · Result Risk Level = HIGH', renderHighRiskSection(s.mills), '02A');
  html += sectionHtml('mill', 'Mill Onboarding · ' + reportLabel, stats.totalMills + ' mills', renderMillSection(s.mills), '02');
  html += sectionHtml('trace', 'Traceability Data · ' + tracePeriodLabel, 'TTM CPO ' + (stats.ttmCpoPct || '—') + ' · TTM PK ' + (stats.ttmPkPct || '—') + ' · TTP CPO ' + (stats.ttpCpoPct || '—') + ' · TTP PK ' + (stats.ttpPkPct || '—'), renderTraceSection(s.traceTotals, stats), '03');
  html += flatSectionHtml('grv', 'Grievance Monitoring · ' + reportLabel, stats.grievances + ' grievances', renderGrvSection(s.grv), '04');
  html += sectionHtml('nbl', 'Active NBL Mills · ' + reportLabel, stats.nblMills + ' mills on No Buy List', renderNblSection(s.mills), '05');
  html += sectionHtml('facility', 'Facility Performance · ' + reportLabel, 'CPO & PK · traceability & ISPO', renderFacilitySection(s.facilityBundles, s.facilityLoading, s.eudrPotential), '06');
  html += sectionHtml('eudr', 'EUDR Potential', stats.eudrPotential + ' potential mills', renderEudrSection(s.eudrPotential, s.eudrLoading), '07');
  sections.innerHTML = html;
}

function renderSddSection(data, loading) {
  if (loading) return '<div class="mrd-empty mrd-empty--loading dash-loading-host">' + dashLoadingHtml_('Loading…', { inline: true }) + '</div>';
  const rows = data.filter(function(r) {
    return matchesSearch([
      r['SCR - Screening Status'], r['Group Name'], r['Grup Name'], r['Mill Name'],
      r['Company Name'], r['Date Imported'], r.supplier_type, r.updated_at,
    ].join(' ').toLowerCase());
  }).slice(0, MRD_SDD_LIMIT);
  const cols = [
    { label: MRD_SDD_COLS[0], always: true, raw: function(r) { return sddDateImport(r); } },
    { label: MRD_SDD_COLS[1], always: true, raw: function(r) { return sddCompanyName(r); } },
    { label: MRD_SDD_COLS[2], always: true, raw: function(r) {
      return normalizeSddCategory(r.supplier_type || r['Supplier Type'] || r['SUPPLIER_TYPE']);
    }},
    { label: MRD_SDD_COLS[3], always: true, raw: function(r) { return sddStatusText(r); } },
    { label: MRD_SDD_COLS[4], always: true, raw: function(r) { return sddLastUpdate(r); } },
  ];
  return renderSmartTable(cols, rows, {
    alwaysShowTable: true,
    emptyRowText: 'No SDD records for this period.',
    note: '<p class="mrd-table-caption">' + esc(MRD_SDD_TABLE_TITLE) + '</p>' + limitNote(data.length, MRD_SDD_LIMIT),
  });
}

function renderHighRiskSection(rows) {
  const filtered = (rows || []).filter(function(item) {
    return mrdIsHighRiskItem_(item) && matchesSearch(item.search);
  });
  const visible = mrdSortMillItems_(filtered).slice(0, MRD_ROW_LIMIT);
  const cols = [
    { label: 'Product', title: 'Main = CPO/PK · Waste = POME/SHELL', thCls: 'mrd-th-wrap', always: true, render: function(row) {
      const p = mrdMillProductLabel_(row._data);
      const cls = p === 'Waste' ? 'mrd-pill mrd-pill--waste' : 'mrd-pill mrd-pill--main';
      return '<span class="' + cls + '">' + esc(p) + '</span>';
    }},
    { label: 'Result Risk Level', title: 'Result Risk Level', thCls: 'mrd-th-wrap', hasData: function(row) { return hasCellValue(row._data.risk); }, render: function(row) { return riskPill(row._data.risk); } },
    { label: 'Risk Reason', title: 'Why this Result Risk Level was assigned (sheet formula)', thCls: 'mrd-th-wrap mrd-th-risk-reason', always: true, render: function(row) { return mrdRiskReasonCell_(row._data); } },
    { label: 'Group Name', title: 'Group Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._data.row['GROUP NAME']; } },
    { label: 'Company Name', title: 'Company Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._data.row['COMPANY NAME']; } },
    { label: 'Mill Name', title: 'Mill Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._data.row['MILL NAME']; } },
    { label: 'Province', thCls: 'mrd-th-wrap', raw: function(row) { return row._data.row['PROVINCE']; } },
    { label: 'No Buy List', title: 'No Buy List', thCls: 'mrd-th-wrap', hasData: function(row) { return isNblYes(row._data.nbl); }, render: function(row) {
      return isNblYes(row._data.nbl) ? '<span class="mrd-pill mrd-pill--high">Yes</span>' : '';
    }},
  ];
  const tableRows = visible.map(function(item) {
    return { _data: item, _trClass: 'mrd-row--data mrd-row--high-risk' };
  });
  return renderSmartTable(cols, tableRows, {
    empty: '<p class="mrd-empty">No high-risk mills for this period.</p>',
    note: limitNote(filtered.length, MRD_ROW_LIMIT),
    tableClass: 'mrd-table--wide mrd-table--high-risk',
  });
}

function renderMillSection(rows) {
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); });
  // Sort: HIGH risk first so they're always visible even when list is capped.
  filtered.sort(function(a, b) {
    const aH = isHighRisk(a.risk) ? 0 : 1;
    const bH = isHighRisk(b.risk) ? 0 : 1;
    return aH - bH;
  });
  const visible = filtered.slice(0, MRD_ROW_LIMIT);
  const tableRows = visible.map(function(item, idx) {
    const r = item.row;
    const expId = 'mill-' + idx;
    const isExp = _expanded.has(expId);
    const out = { _data: item, _trClass: 'mrd-row--data' };
    out._after = isExp
      ? '<tr class="mrd-detail-row"><td colspan="99"><div class="mrd-detail-panel">' + renderMillDetailHtml_(item, r) + '</div></td></tr>'
      : '';
    out._render = { item: item, r: r, expId: expId, isExp: isExp };
    return out;
  });
  const cols = [
    {
      label: '', always: true, thCls: 'mrd-th-expand', tdCls: 'mrd-td-expand',
      render: function(row) {
        const d = row._render;
        return '<button type="button" class="mrd-expand-btn' + (d.isExp ? ' is-open' : '') + '" data-mrd-expand-btn="' + esc(d.expId) + '" data-mrd-mill-key="' + esc(d.item.cacheKey) + '">›</button>';
      },
    },
    { label: 'Product', title: 'Main = CPO/PK · Waste = POME/SHELL', thCls: 'mrd-th-wrap', always: true, render: function(row) {
      const p = mrdMillProductLabel_(row._render.item);
      const cls = p === 'Waste' ? 'mrd-pill mrd-pill--waste' : 'mrd-pill mrd-pill--main';
      return '<span class="' + cls + '">' + esc(p) + '</span>';
    }},
    { label: 'Result Risk Level', title: 'Result Risk Level', thCls: 'mrd-th-wrap', hasData: function(row) { return hasCellValue(row._render.item.risk); }, render: function(row) { return riskPill(row._render.item.risk); } },
    { label: 'Group Name', title: 'Group Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._render.r['GROUP NAME']; } },
    { label: 'Company Name', title: 'Company Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._render.r['COMPANY NAME']; } },
    { label: 'Mill Name', title: 'Mill Name', thCls: 'mrd-th-wrap', raw: function(row) { return row._render.r['MILL NAME']; } },
    { label: 'Province', thCls: 'mrd-th-wrap', raw: function(row) { return row._render.r['PROVINCE']; } },
    { label: 'No Buy List', title: 'No Buy List', thCls: 'mrd-th-wrap', hasData: function(row) { return isNblYes(row._render.item.nbl); }, render: function(row) {
      return isNblYes(row._render.item.nbl) ? '<span class="mrd-pill mrd-pill--high">Yes</span>' : '';
    }},
  ];
  return renderSmartTable(cols, tableRows, {
    empty: '<p class="mrd-empty">No mills for this period.</p>',
    note: limitNote(filtered.length, MRD_ROW_LIMIT),
    tableClass: 'mrd-table--wide',
  });
}

function renderMillDetailHtml_(item, r) {
  const grid = [
    detailItem('Supplier Status', r['SUPPLIER STATUS']),
    detailItem('Certification', r['CERTIFICATION']),
    detailItem('Total Grievances', r['TOTAL GRIEVANCES']),
    detailItem('Facility Name CPO', r['FACILITY NAME CPO']),
    detailItem('Facility Name PK', r['FACILITY NAME PK']),
    item.nblBy ? detailItem('NBL by', item.nblBy) : '',
  ].filter(Boolean).join('');
  let html = grid ? '<div class="mrd-detail-grid">' + grid + '</div>' : '';
  if (item.nblMatches && item.nblMatches.length) {
    html += '<p class="mrd-detail-subhead">NBL source matches</p><ul class="mrd-detail-list">';
    item.nblMatches.forEach(function(m) {
      html += '<li><strong>' + esc(m.source) + '</strong> — ' + esc(m.target) + ' (via ' + esc(m.by) + ')</li>';
    });
    html += '</ul>';
  } else if (isNblYes(item.nbl) && !item.nblBy) {
    html += '<p class="mrd-detail-note">NBL source — click expand again if it has not appeared.</p>';
  }
  if (isHighRisk(item.risk)) {
    const reason = mrdRiskReason_(item);
    if (reason) {
      html += '<p class="mrd-detail-note mrd-detail-note--risk"><strong>Risk Reason:</strong> ' + esc(reason) + '</p>';
    } else {
      html += '<p class="mrd-detail-note mrd-detail-note--risk">High risk supplier — review mitigation plan.</p>';
    }
  }
  return html;
}

function detailItem(label, val) {
  if (!hasCellValue(val)) return '';
  return '<div class="mrd-detail-item"><span class="mrd-detail-lbl">' + esc(label) + '</span><span class="mrd-detail-val">' + esc(val) + '</span></div>';
}

function formatPctNum_(n) {
  if (n == null || isNaN(n)) return '—';
  if (_deps && _deps.formatPct) return _deps.formatPct(n);
  return (Math.round(n * 10) / 10) + '%';
}

function renderTraceSection(traceTotals, stats) {
  const t = traceTotals || {};
  const s = stats || {};
  function card(cls, val, lbl, hint) {
    return '<div class="mrd-trace-total-card ' + cls + '">'
      + '<span class="mrd-trace-total-val">' + esc(val || '—') + '</span>'
      + '<span class="mrd-trace-total-lbl">' + esc(lbl) + '</span>'
      + '<span class="mrd-trace-total-hint">' + esc(hint) + '</span>'
      + '</div>';
  }
  return ''
    + '<div class="mrd-trace-totals">'
    + '<div class="mrd-trace-totals-group mrd-trace-totals-group--ttm">'
    + '<span class="mrd-trace-totals-group__title">TTM · Traceability to Mill</span>'
    + '<div class="mrd-trace-totals-group__cards">'
    + card('mrd-trace-total-card--cpo', t.ttmCpoFmt, 'TTM CPO %', 'Supply-weighted · valid coordinate')
    + card('mrd-trace-total-card--pk', t.ttmPkFmt, 'TTM PK %', 'Supply-weighted · valid coordinate')
    + '</div></div>'
    + '<div class="mrd-trace-totals-group mrd-trace-totals-group--ttp">'
    + '<span class="mrd-trace-totals-group__title">TTP · Traceability to Plantation</span>'
    + '<div class="mrd-trace-totals-group__cards">'
    + card('mrd-trace-total-card--ttp-cpo', t.ttpCpoFmt, 'TTP CPO %', 'Supplier traceability · volume-weighted')
    + card('mrd-trace-total-card--ttp-pk', t.ttpPkFmt, 'TTP PK %', 'Supplier traceability · volume-weighted')
    + '</div></div>'
    + '</div>'
    + '<p class="mrd-trace-total-note">'
    + esc((s.emptyTraceMills || 0) + ' mills without supplier · ' + (s.totalMills || 0) + ' total mills')
    + '</p>';
}

function renderNblSection(millRows) {
  const filtered = mrdSortMillItems_(millRows.filter(function(item) {
    return isNblYes(item.nbl) && matchesSearch(item.search);
  }));
  if (!filtered.length) {
    return '<p class="mrd-empty">No active NBL mills for this period.</p>';
  }
  function colHtml(items) {
    return items.map(function(item) {
      const r = item.row;
      return '<div class="mrd-nbl-item">'
        + '<span class="mrd-nbl-item-grp">' + esc(r['GROUP NAME'] || '—') + '</span>'
        + '<span class="mrd-nbl-item-co">' + esc(r['COMPANY NAME'] || '—')
        + ' <span class="' + (mrdMillProductLabel_(item) === 'Waste' ? 'mrd-pill mrd-pill--waste' : 'mrd-pill mrd-pill--main') + '">'
        + esc(mrdMillProductLabel_(item)) + '</span></span>'
        + '<span class="mrd-nbl-item-riser">' + esc(mrdFormatNblRisers_(item)) + '</span>'
        + '</div>';
    }).join('');
  }
  let html = '';
  if (filtered.length > MRD_ROW_LIMIT) {
    html += limitNote(filtered.length, MRD_ROW_LIMIT);
  }
  const nblHead = '<div class="mrd-nbl-head"><span>Company Group Name</span><span>Company Name</span><span>NBL Riser</span></div>';
  const visible = filtered.slice(0, MRD_ROW_LIMIT);
  const visHalf = Math.ceil(visible.length / 2);
  html += '<div class="mrd-nbl-grid">'
    + '<div class="mrd-nbl-col">' + nblHead + colHtml(visible.slice(0, visHalf)) + '</div>'
    + '<div class="mrd-nbl-col">' + nblHead + colHtml(visible.slice(visHalf)) + '</div>'
    + '</div>';
  return html;
}

function renderGrvSection(rows) {
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const fieldDefs = [
    { key: 'Date Received', label: MRD_GRV_SUMMARY_COLS[0], always: true },
    { key: 'Grievance Category', label: MRD_GRV_SUMMARY_COLS[1], always: true },
    { key: 'Complainant', label: MRD_GRV_SUMMARY_COLS[2], always: true },
    { key: 'Grievance Subject Group', label: MRD_GRV_SUMMARY_COLS[3], always: true, group: true },
    { key: 'Grievance Subject', label: MRD_GRV_SUMMARY_COLS[4], altKey: 'Subject', always: true },
    { key: 'Risk Classification', label: MRD_GRV_SUMMARY_COLS[5], pill: 'risk', always: true },
    { key: 'Grievance Status', label: MRD_GRV_SUMMARY_COLS[6], pill: 'grv', always: true },
  ];
  const cols = fieldDefs.map(function(f) {
    return {
      label: f.label,
      always: true,
      render: function(item) {
        if (f.group) return esc(grvGroupName_(item.row));
        const v = item.row[f.key] || (f.altKey ? item.row[f.altKey] : '');
        if (f.pill === 'grv') return statusPill(v, 'grv');
        if (f.pill === 'risk') return riskPill(v);
        return esc(v || '—');
      },
    };
  });
  return renderSmartTable(cols, filtered, {
    empty: '<p class="mrd-empty">No grievances for this period.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
    tableClass: 'mrd-table--wide',
  });
}

function renderFacilitySection(bundles, loading, eudrPotential) {
  if (loading) return '<div class="mrd-empty mrd-empty--loading dash-loading-host">' + dashLoadingHtml_('Loading…', { inline: true }) + '</div>';
  const eudrCompanySet = mrdBuildEudrPotentialCompanySet_(eudrPotential);
  const active = mrdSortFacilityBundles_((bundles || []).filter(function(b) {
    return (b.companies || []).length > 0 && hasCellValue(b.facility);
  }));
  if (!active.length) return '<p class="mrd-empty">No facility performance data for this period.</p>';

  return active.map(function(bundle) {
    const isPk = bundle.type === 'pk';
    const accent = isPk ? 'pk' : 'cpo';
    const sum = bundle.summary || {};
    const pctLabel = facilityPctColLabel(isPk);
    const ttmLabel = facilityTtmColLabel(isPk);
    const facilityPct = (bundle.traceCalc && bundle.traceCalc.formatted)
      ? bundle.traceCalc.formatted
      : (isPk ? (sum.avgPk || '—') : (sum.avgCpo || '—'));
    const facilityTtm = (bundle.ttmCalc && bundle.ttmCalc.formatted)
      ? bundle.ttmCalc.formatted
      : (sum.avgTtm || '0%');
    const companies = mrdSortFacilityCompanies_(bundle.companies || []);
    const eudrCount = mrdFacilityEudrPotentialCount_(companies, eudrCompanySet);
    const uniqueCount = mrdUniqueFacilityCompanyCount_(companies);
    const visible = companies.filter(function(c) {
      return matchesSearch([bundle.facility, c.company, c.group, c.certification, c.riskLevel].join(' ').toLowerCase());
    }).slice(0, MRD_ROW_LIMIT);

    const kpis = [
      { l: 'No Buy List', v: String(sum.nblYes || 0) },
      { l: 'High Risk', v: String(sum.highRisk || 0), warn: (sum.highRisk || 0) > 0 },
      { l: 'Total Grievance', v: String(sum.grievanceSum || 0) },
      { l: 'Estimated ISPO Supply %', v: sum.ispoPct || '—' },
      { l: ttmLabel, v: facilityTtm, pct: true },
      { l: pctLabel, v: facilityPct, pct: true },
    ];
    const kpiHtml = kpis.map(function(k) {
      return '<div class="mrd-facility-kpi' + (k.pct ? ' mrd-facility-kpi--pct' : '') + (k.warn ? ' mrd-facility-kpi--warn' : '') + '">'
        + '<span class="mrd-facility-kpi-val">' + esc(k.v) + '</span>'
        + '<span class="mrd-facility-kpi-lbl">' + esc(k.l) + '</span></div>';
    }).join('');

    const cols = [
      { label: 'Group Name', always: true, raw: function(c) { return c.group || c['GROUP NAME'] || ''; } },
      { label: 'Company Name', always: true, render: function(c) {
        const name = esc(c.company);
        if (!mrdCompanyIsEudrPotential_(c, eudrCompanySet)) return name;
        return name + ' <span class="mrd-eudr-flag eudr-status-val eudr-status-val--potential">EUDR Potential</span>';
      }},
      { label: 'Certification', raw: function(c) { return c.certification; } },
      { label: 'No Buy List', hasData: function(c) { return isNblYes(c.nbl); }, render: function(c) {
        return isNblYes(c.nbl) ? '<span class="mrd-pill mrd-pill--high">Yes</span>' : esc(c.nbl || '—');
      }},
      { label: 'Result Risk Level', hasData: function(c) { return hasCellValue(c.riskLevel); }, render: function(c) { return riskPill(c.riskLevel); } },
      { label: 'Total Grievance', hasData: function(c) { return !isZeroish(c.grievance); }, raw: function(c) { return c.grievance; }, tdCls: 'mrd-td-num' },
      { label: ttmLabel, always: true, render: function(c) { return formatPctNum_(c.ttmPctNum != null ? c.ttmPctNum : 0); }, tdCls: 'mrd-td-pct' },
      { label: pctLabel, always: true, render: function(c) { return formatPctNum_(c.ttpPctNum); }, tdCls: 'mrd-td-pct' },
    ];
    const table = renderSmartTable(cols, visible, {
      empty: '<p class="mrd-empty">No companies for this facility.</p>',
      note: '<p class="mrd-table-caption">Company breakdown</p>' + limitNote(companies.length, MRD_ROW_LIMIT),
      tableClass: 'mrd-table--wide',
    });

    const eudrMeta = eudrCount > 0
      ? ' · <span class="mrd-facility-eudr">' + esc(mrdEudrPotentialLabel_(eudrCount)) + '</span>'
      : '';
    const companyMeta = eudrCount > 0
      ? esc(uniqueCount + '/' + eudrCount + ' companies')
      : esc(companies.length + ' companies');

    return '<div class="mrd-facility-block mrd-facility-block--' + accent + '">'
      + '<div class="mrd-facility-head">'
      + '<span class="mrd-facility-badge">' + esc(isPk ? 'PK' : 'CPO') + '</span>'
      + '<span class="mrd-facility-name">' + esc(bundle.facility) + '</span>'
      + '<span class="mrd-facility-meta">' + companyMeta
      + eudrMeta + ' · '
      + esc(facilityTtm) + ' ' + esc(ttmLabel.replace('% ', '')) + ' · '
      + esc(facilityPct) + ' ' + esc(pctLabel.replace('% ', '')) + '</span>'
      + '</div>'
      + '<div class="mrd-facility-kpi-row">' + kpiHtml + '</div>'
      + table
      + '</div>';
  }).join('');
}

function renderEudrSection(rows, loading) {
  if (loading) return '<div class="mrd-empty mrd-empty--loading dash-loading-host">' + dashLoadingHtml_('Loading…', { inline: true }) + '</div>';
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const cols = [
    { label: 'Group Name', raw: function(item) { return item.row['GROUP NAME']; } },
    { label: 'Company Name', raw: function(item) { return item.row['COMPANY NAME']; } },
    { label: 'Mill Name', raw: function(item) { return item.row['MILL NAME']; } },
    { label: 'Province', raw: function(item) { return item.row['PROVINCE']; } },
    { label: 'CPO Supply', raw: function(item) { return eudrCpoSupply_(item.row); } },
    { label: 'PK Supply', raw: function(item) { return eudrPkSupply_(item.row); } },
  ];
  return renderSmartTable(cols, filtered, {
    empty: '<p class="mrd-empty">No mills with Potential status.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
  });
}

function renderAll() {
  if (!_snapshot) return;
  syncPeriodFromUi_();
  syncMrdViewModeUi_();
  renderKpis(_snapshot.stats, {
    tracePending: !!_snapshot.supplementalLoading || !_ttpFetchOk,
    eudrPending: !!_snapshot.eudrLoading || _eudrPending || !_eudrFetchOk,
  });

  const execEl = document.getElementById('mrdExecutiveView');
  const sections = document.getElementById('mrdSections');
  if (_viewMode === 'executive') {
    if (sections) sections.hidden = true;
    if (execEl) execEl.hidden = false;
    scheduleRenderMrdExecutive_();
  } else {
    destroyMrdExecutiveCharts_();
    if (execEl) execEl.hidden = true;
    if (sections) {
      sections.hidden = false;
      renderDetailSections_();
    }
  }
}

function updateScopeText(extra) {
  const el = document.getElementById('mrdScopeText');
  if (!el) return;
  syncPeriodFromUi_();
  const report = (_snapshot && _snapshot.reportPeriod && (_snapshot.reportPeriod.year || _snapshot.reportPeriod.month))
    ? _snapshot.reportPeriod
    : getReportPeriod_();
  const meta = mrdReportHeaderMeta_(report.year, report.month);
  let txt = meta.periodLine + ' · ' + meta.dataPeriodLine + ' · ' + meta.cutoffLine;
  txt += ' · Product: Main + Waste';
  if (_snapshot) {
    txt += ' · ' + _snapshot.stats.totalMills + ' mills · ' + _snapshot.stats.sddSubmitted + ' SDD submitted';
  }
  if (extra) txt += ' · ' + extra;
  el.textContent = txt;
}

function setUiLoading(show) {
  const loading = document.getElementById('mrdLoading');
  const content = document.getElementById('mrdContent');
  if (loading) loading.hidden = !show;
  if (content) content.hidden = show;
}

async function loadFacilityInBackground(gen, opts) {
  opts = opts || {};
  const facilityPeriod = getFacilityReportContext_();
  const periodKey = mrdFacilityPeriodKey_(facilityPeriod);
  if (_facilityPending || (!_deps.loadFacilityBundlesForReport && !_deps.getFacilityBundles)) return;
  if (!opts.force && _facilityBundles.length && _facilityBundlesPeriodKey === periodKey) {
    if (_snapshot) {
      _snapshot.facilityBundles = _facilityBundles;
      _snapshot.facilityLoading = false;
      scheduleRenderAll();
    }
    return;
  }
  _facilityPending = true;
  if (_snapshot) {
    _snapshot.facilityLoading = true;
    scheduleRenderAll();
  }
  try {
    if (_deps.loadFacilityBundlesForReport) {
      _facilityBundles = mrdSortFacilityBundles_(await _deps.loadFacilityBundlesForReport(facilityPeriod) || []);
    } else {
      if (_deps.preparePfDataForReport) await _deps.preparePfDataForReport(facilityPeriod);
      if (gen !== _loadGen) return;
      _facilityBundles = mrdSortFacilityBundles_(_deps.getFacilityBundles(facilityPeriod) || []);
    }
    _facilityBundlesPeriodKey = periodKey;
    if (gen !== _loadGen) return;
    if (_snapshot) {
      _snapshot.facilityBundles = _facilityBundles;
      _snapshot.facilityLoading = false;
      scheduleRenderAll();
    }
  } catch (_) {
    if (_snapshot) {
      _snapshot.facilityLoading = false;
      if (gen === _loadGen) {
        _snapshot.facilityBundles = [];
        _facilityBundles = [];
        _facilityBundlesPeriodKey = '';
      }
      scheduleRenderAll();
    }
  } finally {
    _facilityPending = false;
    if (gen === _loadGen) mrdRebuildIfMillsLoaded_();
    // Guard: always clear loading flag so it can't get permanently stuck
    if (_snapshot && _snapshot.facilityLoading) {
      _snapshot.facilityLoading = false;
      scheduleRenderAll();
    }
  }
}

async function loadEudrInBackground(gen, opts) {
  opts = opts || {};
  if (!_deps.fetchEudrPotential) return;
  if (_eudrPending) return;
  const ttpLen = (_deps.getTtpData ? (_deps.getTtpData() || []).length : 0);
  const ttpStable = ttpLen > 0 && ttpLen === _lastEudrTtpCount;
  if (!opts.force && ttpStable && _eudrFetchOk && _snapshot && _snapshot.eudrPotential && _snapshot.eudrPotential.length && !_snapshot.eudrLoading) {
    return;
  }
  _eudrPending = true;
  if (_snapshot) {
    _snapshot.eudrLoading = true;
    scheduleRenderAll();
  }
  try {
    const eudr = await _deps.fetchEudrPotential();
    if (gen !== _loadGen || !_snapshot) return;
    _eudrFetchOk = true;
    _lastEudrTtpCount = (_deps.getTtpData ? (_deps.getTtpData() || []).length : 0);
    _snapshot.eudrPotential = eudr;
    _snapshot.eudrLoading = false;
    _snapshot.stats.eudrPotential = eudr.length;
    scheduleRenderAll();
    updateScopeText();
  } catch (err) {
    if (_snapshot) {
      _snapshot.eudrLoading = false;
      if (gen === _loadGen) {
        _snapshot.eudrPotential = [];
        _snapshot.stats.eudrPotential = 0;
      }
      scheduleRenderAll();
      updateScopeText('EUDR: ' + (err && err.message ? err.message : 'failed to load'));
    }
  } finally {
    _eudrPending = false;
    _eudrFetchOk = true;
    if (_snapshot && _snapshot.eudrLoading) {
      _snapshot.eudrLoading = false;
      scheduleRenderAll();
    }
  }
}

function syncNblCacheToSnapshot_(snapshot) {
  const snap = snapshot || _snapshot;
  if (!snap || !snap.mills) return;
  snap.mills.forEach(function(item) {
    if (!item || !item.cacheKey) return;
    const cached = _nblByCache.get(item.cacheKey);
    if (!cached) return;
    item.nblBy = cached.label || '';
    item.nblMatches = cached.matches || [];
  });
}

function mrdRebuildIfMillsLoaded_() {
  if (!_deps || !_snapshot) return false;
  const millData = _deps.getMillData() || [];
  if (!millData.length) return false;
  const prevTotal = (_snapshot.stats && _snapshot.stats.totalMills) || 0;
  if (prevTotal > 0) return false;
  _snapshot = rebuildSnapshot_({
    facilityBundles: _facilityBundles,
    facilityLoading: false,
    sddRows: _sddCache,
    sddLoading: false,
  });
  scheduleRenderAll();
  updateScopeText();
  scheduleNblRiserResolve_();
  return true;
}

function scheduleNblRiserResolve_() {
  const gen = _loadGen;
  resolveNblMillsForSnapshot_().catch(function() {});
  return gen;
}

function rebuildSnapshot_(opts) {
  opts = opts || {};
  const prev = _snapshot || {};
  const snap = buildSnapshotSync({
    reportPeriod: opts.reportPeriod || getReportPeriod_(),
    sddRows: opts.sddRows != null ? opts.sddRows : _sddCache,
    eudrPotential: opts.eudrPotential != null ? opts.eudrPotential : (prev.eudrPotential || []),
    eudrLoading: opts.eudrLoading != null ? opts.eudrLoading : !!prev.eudrLoading,
    sddLoading: opts.sddLoading != null ? opts.sddLoading : !!prev.sddLoading,
    facilityBundles: opts.facilityBundles != null ? opts.facilityBundles : (prev.facilityBundles || _facilityBundles),
    facilityLoading: opts.facilityLoading != null ? opts.facilityLoading : !!prev.facilityLoading,
  });
  syncNblCacheToSnapshot_(snap);
  _lastTtpCount = (_deps && _deps.getTtpData ? (_deps.getTtpData() || []).length : 0);
  return snap;
}

function rebuildMonthlyReportSnapshot_() {
  if (!_deps) return;
  const ttpLen = (_deps.getTtpData ? (_deps.getTtpData() || []).length : 0);
  _ttpFetchOk = ttpLen > 0;
  if (_eudrFetchOk && ttpLen > 0 && ttpLen !== _lastEudrTtpCount) {
    if (_deps.clearEudrCache) _deps.clearEudrCache();
  }
  _snapshot = rebuildSnapshot_({});
  scheduleRenderAll();
  updateScopeText();
  if (!_eudrFetchOk && !_eudrPending && _deps.fetchEudrPotential) {
    loadEudrInBackground(_loadGen, {});
  }
}

async function loadMillInBackground(gen) {
  const loadMill = _deps.ensureMillData || _deps.ensureCoreData;
  if (!loadMill) return;
  try {
    await withTimeout(loadMill(), 90000, 'Mill data');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    scheduleRenderAll();
    updateScopeText();
    populateYearSelect();
    await resolveNblMillsForSnapshot_();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('mill: ' + (err && err.message ? err.message : String(err)));
  }
}

function millNeedsNblRiserResolve_(item) {
  if (!item || !isNblYes(item.nbl)) return false;
  const matches = item.nblMatches || [];
  if (matches.length && matches.some(function(m) { return String(m && m.riser || '').trim(); })) {
    return false;
  }
  const by = String(item.nblBy || '').trim();
  if (by && !/source unresolved/i.test(by)) {
    const stripped = by.replace(/^NBL by\s+/i, '').trim();
    if (stripped) return false;
  }
  return true;
}

async function resolveNblMillsForSnapshot_(preloadedLists) {
  if (!_snapshot || !_deps.ensureNblLists || !_deps.resolveNblBy) return;
  const loadGen = _loadGen;
  try {
    const lists = preloadedLists || await _deps.ensureNblLists();
    if (loadGen !== _loadGen || !_snapshot) return;
    const pending = (_snapshot.mills || []).filter(millNeedsNblRiserResolve_);
    if (!pending.length) {
      syncNblCacheToSnapshot_();
      return;
    }
    pending.forEach(function(item) {
      const info = _deps.resolveNblBy(item.row, lists);
      if (item.cacheKey) _nblByCache.set(item.cacheKey, info);
    });
    syncNblCacheToSnapshot_();
    scheduleRenderAll();
  } catch (err) {
    console.warn('[MRD] NBL riser resolve:', err && err.message ? err.message : err);
  }
}

async function loadNblRisersInBackground(gen) {
  if (!_deps || !_deps.ensureNblLists) return;
  try {
    const lists = await _deps.ensureNblLists();
    if (gen !== _loadGen) return;
    await resolveNblMillsForSnapshot_(lists);
  } catch (_) { /* optional */ }
}

async function loadSupplementalInBackground(gen) {
  if (!_deps.ensureSupplementalData) return;
  if (_snapshot) {
    _snapshot.supplementalLoading = true;
    scheduleRenderAll();
  }
  try {
    await withTimeout(_deps.ensureSupplementalData(), 120000, 'Traceability & grievance');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    await resolveNblMillsForSnapshot_();
    if (gen !== _loadGen) return;
    scheduleRenderAll();
    updateScopeText();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('supplemental data: timeout — click Refresh to try again');
  } finally {
    _ttpFetchOk = true;
    if (_snapshot) {
      _snapshot.supplementalLoading = false;
      scheduleRenderAll();
    }
  }
}

async function loadSddInBackground(gen, opts) {
  opts = opts || {};
  if (!_deps.fetchSddList) return;
  if (!opts.force && _sddCache.length) {
    if (_snapshot) {
      _snapshot = rebuildSnapshot_({ sddRows: _sddCache, sddLoading: false });
      scheduleRenderAll();
    }
    return;
  }
  _snapshot = rebuildSnapshot_({ sddLoading: true });
  scheduleRenderAll();
  try {
    const rows = await withTimeout(_deps.fetchSddList(), 20000, 'SDD');
    if (gen !== _loadGen) return;
    _sddCache = Array.isArray(rows) ? rows : [];
    _snapshot = rebuildSnapshot_({ sddRows: _sddCache, sddLoading: false });
    scheduleRenderAll();
    updateScopeText();
  } catch (err) {
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({ sddLoading: false });
    scheduleRenderAll();
    updateScopeText('SDD: ' + (err && err.message ? err.message : 'failed to load'));
  }
}

function startBackgroundLoads_(gen, force) {
  const opts = { force: force };
  Promise.all([
    loadMillInBackground(gen),
    loadSupplementalInBackground(gen),
    loadSddInBackground(gen, opts),
    loadNblRisersInBackground(gen),
  ]).then(function() {
    updateScopeText();
    mrdRebuildIfMillsLoaded_();
    scheduleNblRiserResolve_();
    // Use _loadGen (not gen) so facility always loads for the current active generation
    loadEudrInBackground(_loadGen, opts);
    loadFacilityInBackground(_loadGen, opts);
  }).catch(function() {
    updateScopeText();
  });
}

async function loadAndRender(opts) {
  opts = opts || {};
  const force = !!opts.force;
  if (force) {
    _eudrFetchOk = false;
    _lastEudrTtpCount = -1;
    _ttpFetchOk = false;
  }
  const gen = ++_loadGen;
  syncPeriodFromUi_();
  const errEl = document.getElementById('mrdError');
  if (errEl) { errEl.hidden = true; errEl.textContent = ''; }

  try {
    // Selalu tampilkan UI dari cache memori — jangan blokir menunggu network
    // Reset loading flags so they don't stay stuck from a previous generation
    _snapshot = rebuildSnapshot_({ eudrLoading: false, facilityLoading: false });
    mrdSyncMonthOptions_();
    syncPeriodFromUi_(); // re-read _month after auto-correct
    renderAll();
    updateScopeText('loading data…');
    scheduleNblRiserResolve_();
  } catch (err) {
    if (errEl) {
      errEl.hidden = false;
      errEl.textContent = 'Failed to render report: ' + (err && err.message ? err.message : String(err));
    }
  } finally {
    setUiLoading(false);
  }

  startBackgroundLoads_(gen, force);
}

async function resolveMillNblOnExpand(cacheKey, row) {
  if (_nblByCache.has(cacheKey)) return _nblByCache.get(cacheKey);
  const lists = await _deps.ensureNblLists();
  const info = _deps.resolveNblBy(row, lists);
  _nblByCache.set(cacheKey, info);
  return info;
}

function populateYearSelect() {
  const yearSel = document.getElementById('mrdYearSel');
  if (!yearSel) return;
  mrdRenderYearSelectOptions_(yearSel, _year);
  yearSel._mrdYearsPopulated = true;
}

function bindOnce() {
  if (_bound) return;
  const panel = document.getElementById('panel-monthly-report-detail');
  if (!panel) return;

  populateYearSelect();
  applyDefaultMonthSelection_();

  const yearSel = document.getElementById('mrdYearSel');
  const monthSel = document.getElementById('mrdMonthSel');
  const searchEl = document.getElementById('mrdSearch');

  if (yearSel && !yearSel._mrdBound) {
    yearSel._mrdBound = true;
    yearSel.addEventListener('change', function() {
      syncPeriodFromUi_();
      _monthAutoSelected = false; // reset so month re-auto-selects for new year
      _nblByCache.clear();
      _facilityBundles = [];
      _facilityBundlesPeriodKey = '';
      loadAndRender();
    });
  }

  if (monthSel && !monthSel._mrdBound) {
    monthSel._mrdBound = true;
    monthSel.addEventListener('change', function() {
      syncPeriodFromUi_();
      _nblByCache.clear();
      _facilityBundles = [];
      _facilityBundlesPeriodKey = '';
      loadAndRender();
    });
  }

  document.getElementById('mrdBtnRefresh')?.addEventListener('click', function() {
    syncPeriodFromUi_();
    _nblByCache.clear();
    _facilityBundles = [];
    _facilityBundlesPeriodKey = '';
    if (_deps.clearEudrCache) _deps.clearEudrCache();
    loadAndRender({ force: true });
  });

  if (searchEl && !searchEl._mrdBound) {
    searchEl._mrdBound = true;
    let t = null;
    searchEl.addEventListener('input', function() {
      clearTimeout(t);
      t = setTimeout(function() {
        _search = String(searchEl.value || '').trim().toLowerCase();
        renderAll();
      }, 150);
    });
  }

  document.getElementById('mrdBtnExport')?.addEventListener('click', mrdOpenExportModal_);
  document.getElementById('mrdBtnExecExport')?.addEventListener('click', function() {
    exportMrdExecutiveReport_();
  });
  mrdBindExportModalOnce_();

  document.querySelectorAll('[data-mrd-view]').forEach(function(btn) {
    if (btn.dataset.mrdViewBound) return;
    btn.dataset.mrdViewBound = '1';
    btn.addEventListener('click', function() {
      setMrdViewMode_(btn.getAttribute('data-mrd-view'));
    });
  });
  syncMrdViewModeUi_();

  panel.addEventListener('click', function(e) {
    const toggleSec = e.target.closest('[data-mrd-toggle]');
    if (toggleSec) {
      const id = toggleSec.getAttribute('data-mrd-toggle');
      const key = id + ':closed';
      const wasClosed = _expanded.has(key);
      if (wasClosed) _expanded.delete(key);
      else _expanded.add(key);
      renderAll();
      return;
    }
    const expBtn = e.target.closest('[data-mrd-expand-btn]');
    if (expBtn) {
      const id = expBtn.getAttribute('data-mrd-expand-btn');
      const millKey = expBtn.getAttribute('data-mrd-mill-key');
      const opening = !_expanded.has(id);
      if (_expanded.has(id)) _expanded.delete(id);
      else _expanded.add(id);
      renderAll();
      if (opening && millKey && _snapshot) {
        const item = _snapshot.mills.find(function(m) { return m.cacheKey === millKey; });
        if (item && millNeedsNblRiserResolve_(item)) {
          resolveMillNblOnExpand(millKey, item.row).then(function(info) {
            item.nblBy = info.label || '';
            item.nblMatches = info.matches || [];
            renderAll();
          });
        }
      }
    }
  });

  _bound = true;
}

export function initMonthlyReport_(deps) {
  _deps = deps;
  window.refreshMonthlyReport_ = function(opts) {
    opts = opts || {};
    const ttpLen = (_deps.getTtpData ? (_deps.getTtpData() || []).length : 0);
    const eudrStale = !_eudrFetchOk && !_eudrPending;
    const ttpStale = ttpLen !== _lastTtpCount || !_ttpFetchOk;
    const traceSuspicious = _snapshot && _snapshot.stats
      && _snapshot.stats.totalMills > 0
      && _snapshot.stats.emptyTraceMills >= _snapshot.stats.totalMills
      && ttpLen > 0;
    if (!opts.force && _snapshot && !ttpStale && !eudrStale && !traceSuspicious) {
      renderAll();
      updateScopeText();
      return;
    }
    loadAndRender(opts);
  };
  window.rebuildMonthlyReportSnapshot_ = rebuildMonthlyReportSnapshot_;
  window.__mrdResetEudrFetch_ = function() { _eudrFetchOk = false; };
  window.exportMonthlyReport_ = exportMonthlyReport_;
  bindOnce();
  loadAndRender();
}

export function refreshMonthlyReport_(opts) {
  if (_deps) window.refreshMonthlyReport_(opts);
}
