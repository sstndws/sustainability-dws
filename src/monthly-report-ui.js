/**
 * Monthly Report (Detail) — read-only compliance snapshot (fast, in-memory first).
 */

import { buildMonthlyReportPdfPair_, MRD_PDF_SECTIONS } from './monthly-report-pdf.js';
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
  mrdSortFacilityBundles_,
  mrdSortFacilityCompanies_,
  mrdSortBundlesByFacility_,
  mrdSortEmptyMillItems_,
  mrdFormatNblRisers_,
  mrdReportHeaderMeta_,
  normalizeSddCategory,
  sddStatusText,
  sddCompanyName,
  sddDateImport,
  sddLastUpdate,
} from './monthly-report-labels.js';

const MRD_ROW_LIMIT = 5000;
const MRD_SDD_LIMIT = 5000;

let _deps = null;
let _bound = false;
let _snapshot = null;
let _search = '';
let _year = String(new Date().getFullYear());
let _month = String(new Date().getMonth() + 1);
let _expanded = new Set();
let _loadGen = 0;
let _eudrPending = false;
let _facilityPending = false;
let _nblByCache = new Map();
let _sddCache = [];
let _facilityBundles = [];
let _renderRaf = 0;

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

/** Sync year/month from toolbar before export, header, or reload. */
function syncPeriodFromUi_() {
  const yearSel = document.getElementById('mrdYearSel');
  const monthSel = document.getElementById('mrdMonthSel');
  if (yearSel && yearSel.value) _year = yearSel.value;
  if (monthSel) _month = monthSel.value;
  return { year: _year, month: _month };
}

function applyDefaultMonthSelection_() {
  const monthSel = document.getElementById('mrdMonthSel');
  if (!monthSel) return;
  if (!monthSel.value) {
    monthSel.value = _month;
  } else {
    _month = monthSel.value;
  }
}

/** Reporting period (UI) vs data period (lag 1 month). */
function getReportPeriod_() {
  syncPeriodFromUi_();
  return { year: _year, month: _month };
}

function getDataPeriod_() {
  const report = getReportPeriod_();
  const meta = mrdReportHeaderMeta_(report.year, report.month);
  return { year: meta.dataYear, month: meta.dataMonth };
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
    if (isNblYes(copy.nbl) && !copy.nblBy) {
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

function mrdOpenExportModal_() {
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

async function exportMonthlyReport_(exportOpts) {
  exportOpts = exportOpts || {};
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

  const btn = document.getElementById('mrdExportConfirm') || document.getElementById('mrdBtnExport');
  const prevTxt = btn ? btn.textContent : '';
  if (btn) { btn.disabled = true; btn.textContent = 'Preparing…'; }

  try {
    syncPeriodFromUi_();
    if (sections.includes('sdd') && _sddCache.length === 0 && _deps.fetchSddList) {
      try {
        const rows = await withTimeout(_deps.fetchSddList(), 25000, 'SDD');
        _sddCache = Array.isArray(rows) ? rows : [];
        _snapshot = rebuildSnapshot_({ sddRows: _sddCache, sddLoading: false });
      } catch (_) { /* continue without SDD */ }
    }

    if (btn) btn.textContent = 'Building PDF…';
    const extra = await _deps.preparePdfExport({ sections: sections });
    // Merge freshly-fetched extra.eudr into the snapshot so stats are correct even
    // when the background load hadn't completed yet. Use extra.eudr only if non-empty
    // (a truthy [] would otherwise shadow a previously-loaded _snapshot.eudrPotential).
    const freshEudr = extra && extra.eudr && extra.eudr.length ? extra.eudr : null;
    _snapshot = rebuildSnapshot_({ eudrPotential: freshEudr !== null ? freshEudr : undefined });
    const s = _snapshot;

    const mills = await resolveAllNblForExport_(
      mrdSortMillItems_(filterForExport_(s.mills, function(item) { return matchesSearch(item.search); }))
    );
    // highRiskMills must include ALL high-risk mills regardless of search filter.
    // When no search is active, mills already contains every mill — reuse it.
    // When a search IS active, resolve all s.mills in a separate pass.
    const allMillsResolved = _search
      ? await resolveAllNblForExport_(mrdSortMillItems_(s.mills || []))
      : mills;
    const highRiskMills = mrdSortMillItems_(allMillsResolved.filter(function(item) {
      const r = item.row || {};
      const rr = String(
        r['RESULT RISK LEVEL'] != null ? r['RESULT RISK LEVEL'] :
        r['Result Risk Level'] != null ? r['Result Risk Level'] :
        r['RISK LEVEL'] != null ? r['RISK LEVEL'] : item.risk || ''
      ).toLowerCase();
      return rr.includes('high');
    }));
    const nblMills = mrdSortMillItems_(mills.filter(function(item) {
      return isNblYes(item.nbl) && matchesSearch(item.search);
    }));

    let facilityBundles = _deps.getFacilityBundles ? _deps.getFacilityBundles(getDataPeriod_().year) : [];
    if (_search) {
      const q = _search;
      facilityBundles = facilityBundles.filter(function(b) {
        const blob = [b.facility, (b.companies || []).map(function(c) { return c.company; }).join(' ')].join(' ').toLowerCase();
        return blob.includes(q);
      });
    }

    // s.eudrPotential is authoritative: freshEudr was already merged into the
    // snapshot at the rebuildSnapshot_ call above. Using extra.eudr directly
    // could give a stale/partial list that disagrees with stats.eudrPotential.
    const eudrSource = s.eudrPotential || [];
    const eudrList = filterForExport_(eudrSource, function(item) { return matchesSearch(item.search); });

    const report = getReportPeriod_();
    await buildMonthlyReportPdfPair_({
      getJsPDF: _deps.getJsPDF,
      year: report.year,
      month: report.month,
      sections: sections,
      data: {
        stats: s.stats,
        facility: s.facility,
        sdd: mrdSortSddRows_(filterForExport_(s.sdd, function(r) {
          return matchesSearch([
            r['SCR - Screening Status'], r['Group Name'], r['Grup Name'], r['Mill Name'],
            r.supplier_type, r.updated_at,
          ].join(' ').toLowerCase());
        })),
        mills: mills,
        highRiskMills: highRiskMills,
        emptyMills: filterForExport_(s.emptyMills, function(item) {
          return matchesSearch([
            item.millRow['GROUP NAME'], item.millRow['COMPANY NAME'], item.millRow['MILL NAME'],
          ].join(' ').toLowerCase());
        }),
        traceRows: filterForExport_(s.traceRows, function(item) { return matchesSearch(item.search); }),
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
    if (btn) {
      btn.disabled = false;
      btn.textContent = prevTxt || (btn.id === 'mrdExportConfirm' ? 'Generate 2 PDFs' : 'Export PDF');
    }
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
      const sections = mrdExportSelectedSections_();
      if (!sections) {
        alert('Select at least one section to export.');
        return;
      }
      mrdCloseExportModal_();
      await exportMonthlyReport_({ sections: sections });
    });
  }
}

function buildSnapshotSync(opts) {
  opts = opts || {};
  const reportPeriod = getReportPeriod_();
  const dataPeriod = getDataPeriod_();
  const dataYear = dataPeriod.year;
  const dataMonth = dataPeriod.month;
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
      if (targetYear && y && y !== targetYear) return false;
      if (targetMonth && millMonth) {
        const m = String(millMonth(r) || '').trim();
        if (m && m !== targetMonth) return false;
      }
      return true;
    });
  }

  // Mill Onboarding: 1-month lag. If that yields 0 results, fall back to report month.
  let mills = filterMillsByPeriod_(dataYear, dataMonth);
  let millEffectiveYear = dataYear;
  let millEffectiveMonth = dataMonth;
  if (!mills.length && (reportPeriod.year || reportPeriod.month)) {
    const fallback = filterMillsByPeriod_(reportPeriod.year, reportPeriod.month);
    if (fallback.length) {
      mills = fallback;
      millEffectiveYear = reportPeriod.year;
      millEffectiveMonth = reportPeriod.month;
    }
  }

  const sddFiltered = sddInput.filter(function(r) {
    const st = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
    if (st !== 'draft' && st !== 'submitted') return false;
    if (dataYear) {
      const upd = String(r.updated_at || r['SCR - Last Updated'] || '').slice(0, 4);
      if (upd && upd !== dataYear) return false;
    }
    if (dataMonth) {
      const m = parseMonthFromDate(r.updated_at || r['SCR - Last Updated']);
      if (m && m !== dataMonth) return false;
    }
    return true;
  });

  const millCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'MILL NAME'; }) || 'MILL NAME';
  const supplierCol = ttpFields.find(function(h) { return /FFB SUPPLIER NAME/i.test(h); }) || '';
  const yearCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'YEAR'; }) || 'YEAR';

  const ttpFiltered = ttpData.filter(function(r) {
    const y = parseYear(r[yearCol] || millYear(r));
    return !dataYear || !y || y === dataYear;
  });

  const millKeys = new Map();
  mills.forEach(function(r) {
    const key = [r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']].map(function(x) {
      return String(x || '').trim().toLowerCase();
    }).join('|');
    if (!millKeys.has(key)) millKeys.set(key, r);
  });

  const ttpByMill = new Map();
  ttpFiltered.forEach(function(r) {
    const mill = String(r[millCol] || '').trim() || '(No Mill Name)';
    if (!ttpByMill.has(mill)) ttpByMill.set(mill, []);
    ttpByMill.get(mill).push(r);
  });

  const emptyMills = [];
  millKeys.forEach(function(millRow) {
    const millName = String(millRow['MILL NAME'] || '').trim();
    const rows = ttpByMill.get(millName) || [];
    const hasSupplier = rows.some(function(r) {
      return !!(supplierCol && String(r[supplierCol] || '').trim());
    });
    if (!hasSupplier) emptyMills.push({ millRow: millRow, rows: rows });
  });
  const emptyMillsSorted = mrdSortEmptyMillItems_(emptyMills);

  const millRows = mills.map(function(r) {
    const cacheKey = [r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME']].join('|');
    const cached = _nblByCache.get(cacheKey);
    return {
      row: r,
      cacheKey: cacheKey,
      risk: _deps.millResolvedRiskLevel(r),
      nbl: r['BUYER NO BUY LIST'],
      nblBy: cached ? cached.label : '',
      nblMatches: cached ? cached.matches : null,
      search: [
        r['GROUP NAME'], r['COMPANY NAME'], r['MILL NAME'], r['PROVINCE'],
        r['RESULT RISK LEVEL'], r['BUYER NO BUY LIST'],
        r['CERTIFICATION'], r['SUPPLIER STATUS'],
      ].join(' ').toLowerCase(),
    };
  });

  // Grievance: full data year (no month filter — shows entire year's grievances)
  const grvRows = mrdSortGrvItemsByDateDesc_(grvData.filter(function(r) {
    if (dataYear) {
      const dr = String(r['Date Received'] || '').slice(0, 4);
      if (dr && dr !== dataYear) return false;
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
  const traceTotals = _deps.buildTraceTotals ? _deps.buildTraceTotals(dataYear) : {};

  // Build highRiskMills from the same millRows array used for stats.highRisk.
  // Read risk directly from the raw row to avoid any stale item.risk value.
  const highRiskMillRows = millRows.filter(function(item) {
    const r = item.row || {};
    const rr = String(
      r['RESULT RISK LEVEL'] != null ? r['RESULT RISK LEVEL'] :
      r['Result Risk Level'] != null ? r['Result Risk Level'] :
      r['RISK LEVEL'] != null ? r['RISK LEVEL'] :
      item.risk || ''
    ).toLowerCase();
    return rr.includes('high');
  });

  return {
    sdd: mrdSortSddRows_(sddFiltered),
    mills: mrdSortMillItems_(millRows),
    highRiskMills: mrdSortMillItems_(highRiskMillRows),
    millsTotal: millRows.length,
    millEffectiveYear: millEffectiveYear,
    millEffectiveMonth: millEffectiveMonth,
    emptyMills: emptyMillsSorted,
    traceRows: mrdSortMillItems_(_deps.buildTraceRows ? _deps.buildTraceRows(mills, ttpFiltered, ttpByMill, supplierCol) : []),
    traceTotals: traceTotals,
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
      highRisk: mills.filter(function(r) {
        const rr = String(
          r['RESULT RISK LEVEL'] != null ? r['RESULT RISK LEVEL'] :
          r['Result Risk Level'] != null ? r['Result Risk Level'] :
          r['RISK LEVEL'] != null ? r['RISK LEVEL'] : ''
        ).toLowerCase();
        return rr.includes('high');
      }).length,
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

function renderKpis(stats) {
  const el = document.getElementById('mrdKpiRow');
  if (!el) return;
  const items = [
    { n: stats.sddRequested != null ? stats.sddRequested : stats.sddTotal, l: 'SDD Requested', s: (stats.sddDone != null ? stats.sddDone : stats.sddSubmitted) + ' done' },
    { n: stats.totalMills, l: 'Total Mills', s: stats.totalGroups + ' groups' },
    { n: stats.emptyTraceMills, l: 'Untraceable Mills', s: 'mills without suppliers', hot: stats.emptyTraceMills > 0 },
    { n: stats.eudrPotential, l: 'EUDR Potential', s: 'by formula' },
  ];
  el.innerHTML = items.map(function(it) {
    return '<div class="stat-card' + (it.hot ? ' mrd-stat-hot' : '') + '">'
      + '<div class="stat-card-num">' + esc(it.n) + '</div>'
      + '<div class="stat-card-label">' + esc(it.l) + '</div>'
      + '<div class="mrd-stat-sub">' + esc(it.s) + '</div></div>';
  }).join('');
}

function renderSddSection(data, loading) {
  if (loading) return '<p class="mrd-empty mrd-empty--loading">Loading Supplier Due Diligence…</p>';
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
    html += '<p class="mrd-detail-note mrd-detail-note--risk">High risk supplier — review mitigation plan.</p>';
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
        + '<span class="mrd-nbl-item-co">' + esc(r['COMPANY NAME'] || '—') + '</span>'
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
    { key: 'Grievance Subject Group', label: MRD_GRV_SUMMARY_COLS[3], always: true },
    { key: 'Grievance Subject', label: MRD_GRV_SUMMARY_COLS[4], altKey: 'Subject', always: true },
    { key: 'Risk Classification', label: MRD_GRV_SUMMARY_COLS[5], pill: 'risk', always: true },
    { key: 'Grievance Status', label: MRD_GRV_SUMMARY_COLS[6], pill: 'grv', always: true },
  ];
  const cols = fieldDefs.map(function(f) {
    return {
      label: f.label,
      always: true,
      render: function(item) {
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

function renderFacilitySection(bundles, loading) {
  if (loading) return '<p class="mrd-empty mrd-empty--loading">Loading facility performance…</p>';
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
      { label: 'Company Name', always: true, raw: function(c) { return c.company; } },
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

    return '<div class="mrd-facility-block mrd-facility-block--' + accent + '">'
      + '<div class="mrd-facility-head">'
      + '<span class="mrd-facility-badge">' + esc(isPk ? 'PK' : 'CPO') + '</span>'
      + '<span class="mrd-facility-name">' + esc(bundle.facility) + '</span>'
      + '<span class="mrd-facility-meta">' + esc(companies.length) + ' companies · '
      + esc(facilityTtm) + ' ' + esc(ttmLabel.replace('% ', '')) + ' · '
      + esc(facilityPct) + ' ' + esc(pctLabel.replace('% ', '')) + '</span>'
      + '</div>'
      + '<div class="mrd-facility-kpi-row">' + kpiHtml + '</div>'
      + table
      + '</div>';
  }).join('');
}

function renderEudrSection(rows, loading) {
  if (loading) return '<p class="mrd-empty mrd-empty--loading">Loading EUDR Potential…</p>';
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const cols = [
    { label: 'Group Name', raw: function(item) { return item.row['GROUP NAME']; } },
    { label: 'Company Name', raw: function(item) { return item.row['COMPANY NAME']; } },
    { label: 'Mill Name', raw: function(item) { return item.row['MILL NAME']; } },
    { label: 'Province', raw: function(item) { return item.row['PROVINCE']; } },
    { label: 'Supply To', raw: function(item) { return item.row['SUPPLY TO']; } },
    { label: 'Status', always: true, render: function() { return statusPill('Potential', 'eudr'); } },
  ];
  return renderSmartTable(cols, filtered, {
    empty: '<p class="mrd-empty">No mills with Potential status.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
  });
}

function renderAll() {
  if (!_snapshot) return;
  renderKpis(_snapshot.stats);
  const sections = document.getElementById('mrdSections');
  if (!sections) return;
  const s = _snapshot;
  const stats = s.stats || {};
  let html = '';
  const dataPeriod = getDataPeriod_();
  // Mill period may differ from data period when fallback to report month was used
  const millMonthForLabel = s.millEffectiveMonth || dataPeriod.month;
  const millYearForLabel = s.millEffectiveYear || dataPeriod.year;
  const dataMonthLabel = millMonthForLabel ? (MRD_MONTH_SHORT_[parseInt(millMonthForLabel, 10)] || millMonthForLabel) : '';
  const millPeriodLabel = (dataMonthLabel && millYearForLabel) ? (dataMonthLabel + ' ' + millYearForLabel) : (millYearForLabel || 'all periods');
  const fullYearLabel = dataPeriod.year ? ('Full year ' + dataPeriod.year) : 'all periods';
  html += flatSectionHtml('sdd', 'Supplier Due Diligence', stats.sddRequested + ' requested · ' + stats.sddDone + ' done · ' + millPeriodLabel, renderSddSection(s.sdd, s.sddLoading), '01');
  html += sectionHtml('mill', 'Mill Onboarding', stats.totalMills + ' mills · ' + stats.highRisk + ' high risk · ' + millPeriodLabel, renderMillSection(s.mills), '02');
  html += sectionHtml('trace', 'Traceability Data · ' + fullYearLabel, 'TTM CPO ' + (stats.ttmCpoPct || '—') + ' · TTM PK ' + (stats.ttmPkPct || '—') + ' · TTP CPO ' + (stats.ttpCpoPct || '—') + ' · TTP PK ' + (stats.ttpPkPct || '—'), renderTraceSection(s.traceTotals, stats), '03');
  html += flatSectionHtml('grv', 'Grievance Monitoring · ' + fullYearLabel, stats.grievances + ' grievances in ' + fullYearLabel, renderGrvSection(s.grv), '04');
  html += sectionHtml('nbl', 'Active NBL Mills', stats.nblMills + ' mills on No Buy List', renderNblSection(s.mills), '05');
  html += sectionHtml('facility', 'Facility Performance', 'CPO & PK · traceability & ISPO', renderFacilitySection(s.facilityBundles, s.facilityLoading), '06');
  html += sectionHtml('eudr', 'EUDR Potential', stats.eudrPotential + ' potential mills', renderEudrSection(s.eudrPotential, s.eudrLoading), '07');
  sections.innerHTML = html;
}

function updateScopeText(extra) {
  const el = document.getElementById('mrdScopeText');
  if (!el) return;
  const report = getReportPeriod_();
  const meta = mrdReportHeaderMeta_(report.year, report.month);
  let txt = meta.periodLine + ' · ' + meta.dataPeriodLine + ' · ' + meta.cutoffLine;
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
  if (_facilityPending || !_deps.getFacilityBundles) return;
  if (!opts.force && _facilityBundles.length) {
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
    if (_deps.preparePfDataForReport) await _deps.preparePfDataForReport();
    if (gen !== _loadGen) return;
    _facilityBundles = mrdSortFacilityBundles_(_deps.getFacilityBundles(getDataPeriod_().year) || []);
    if (_snapshot) {
      _snapshot.facilityBundles = _facilityBundles;
      _snapshot.facilityLoading = false;
      scheduleRenderAll();
    }
  } catch (_) {
    if (_snapshot) {
      _snapshot.facilityLoading = false;
      if (gen === _loadGen) _snapshot.facilityBundles = [];
      scheduleRenderAll();
    }
  } finally {
    _facilityPending = false;
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
  // Skip only if currently in-flight; a previous run that returned 0 results should
  // be retried (fetchEudrPotential no longer caches empty results).
  if (_eudrPending) return;
  if (!opts.force && _snapshot && _snapshot.eudrPotential && _snapshot.eudrPotential.length && !_snapshot.eudrLoading) {
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
    _snapshot.eudrPotential = eudr;
    _snapshot.eudrLoading = false;
    _snapshot.stats.eudrPotential = eudr.length;
    scheduleRenderAll();
    updateScopeText();
  } catch (_) {
    if (_snapshot) {
      _snapshot.eudrLoading = false;
      if (gen === _loadGen) _snapshot.eudrPotential = [];
      scheduleRenderAll();
    }
  } finally {
    _eudrPending = false;
    if (_snapshot && _snapshot.eudrLoading) {
      _snapshot.eudrLoading = false;
      scheduleRenderAll();
    }
  }
}

function rebuildSnapshot_(opts) {
  const prev = _snapshot || {};
  return buildSnapshotSync({
    sddRows: opts.sddRows != null ? opts.sddRows : _sddCache,
    eudrPotential: opts.eudrPotential != null ? opts.eudrPotential : (prev.eudrPotential || []),
    eudrLoading: opts.eudrLoading != null ? opts.eudrLoading : !!prev.eudrLoading,
    sddLoading: opts.sddLoading != null ? opts.sddLoading : !!prev.sddLoading,
    facilityBundles: opts.facilityBundles != null ? opts.facilityBundles : (prev.facilityBundles || _facilityBundles),
    facilityLoading: opts.facilityLoading != null ? opts.facilityLoading : !!prev.facilityLoading,
  });
}

async function loadMillInBackground(gen) {
  const loadMill = _deps.ensureMillData || _deps.ensureCoreData;
  if (!loadMill) return;
  try {
    await withTimeout(loadMill(), 30000, 'Mill data');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    scheduleRenderAll();
    updateScopeText();
    populateYearSelect();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('mill: ' + (err && err.message ? err.message : String(err)));
  }
}

async function resolveNblMillsForSnapshot_() {
  if (!_snapshot || !_deps.ensureNblLists || !_deps.resolveNblBy) return;
  const pending = (_snapshot.mills || []).filter(function(item) {
    return isNblYes(item.nbl) && !item.nblBy;
  });
  if (!pending.length) return;
  try {
    const lists = await _deps.ensureNblLists();
    pending.forEach(function(item) {
      const info = _deps.resolveNblBy(item.row, lists);
      item.nblBy = info.label || '';
      item.nblMatches = info.matches || [];
      if (item.cacheKey) _nblByCache.set(item.cacheKey, info);
    });
    scheduleRenderAll();
  } catch (_) { /* NBL riser lookup is optional for display */ }
}

async function loadSupplementalInBackground(gen) {
  if (!_deps.ensureSupplementalData) return;
  try {
    await withTimeout(_deps.ensureSupplementalData(), 45000, 'Traceability & grievance');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    await resolveNblMillsForSnapshot_();
    if (gen !== _loadGen) return;
    scheduleRenderAll();
    updateScopeText();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('supplemental data: timeout — refresh to try again');
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
  ]).then(function() {
    // Use _loadGen (not gen) so facility always loads for the current active generation
    loadEudrInBackground(_loadGen, opts);
    loadFacilityInBackground(_loadGen, opts);
  }).catch(function() {});
}

async function loadAndRender(opts) {
  opts = opts || {};
  const force = !!opts.force;
  const gen = ++_loadGen;
  syncPeriodFromUi_();
  const errEl = document.getElementById('mrdError');
  if (errEl) { errEl.hidden = true; errEl.textContent = ''; }

  try {
    // Selalu tampilkan UI dari cache memori — jangan blokir menunggu network
    // Reset loading flags so they don't stay stuck from a previous generation
    _snapshot = rebuildSnapshot_({ eudrLoading: false, facilityLoading: false });
    renderAll();
    updateScopeText('loading data…');
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
  if (!yearSel || yearSel._mrdYearsPopulated) return;
  yearSel._mrdYearsPopulated = true;
  const years = new Set(['2026', '2025', '2024']);
  (_deps.getMillData() || []).forEach(function(r) {
    const y = parseYear(_deps.millYearVal(r));
    if (y) years.add(y);
  });
  yearSel.innerHTML = Array.from(years).sort().reverse().map(function(y) {
    return '<option value="' + y + '"' + (y === _year ? ' selected' : '') + '>' + y + '</option>';
  }).join('');
}

function bindOnce() {
  if (_bound) return;
  _bound = true;
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
      _year = yearSel.value || String(new Date().getFullYear());
      _nblByCache.clear();
      _facilityBundles = [];
      loadAndRender();
    });
  }

  if (monthSel && !monthSel._mrdBound) {
    monthSel._mrdBound = true;
    monthSel.addEventListener('change', function() {
      _month = monthSel.value || '';
      loadAndRender();
    });
  }

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

  document.getElementById('mrdBtnRefresh')?.addEventListener('click', function() {
    _nblByCache.clear();
    _facilityBundles = [];
    if (_deps.clearEudrCache) _deps.clearEudrCache();
    loadAndRender({ force: true });
  });

  document.getElementById('mrdBtnExport')?.addEventListener('click', mrdOpenExportModal_);
  mrdBindExportModalOnce_();

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
        if (item && isNblYes(item.nbl) && !item.nblBy) {
          resolveMillNblOnExpand(millKey, item.row).then(function(info) {
            item.nblBy = info.label || '';
            item.nblMatches = info.matches || [];
            renderAll();
          });
        }
      }
    }
  });
}

export function initMonthlyReport_(deps) {
  _deps = deps;
  window.refreshMonthlyReport_ = function(opts) {
    opts = opts || {};
    if (!opts.force && _snapshot) {
      renderAll();
      updateScopeText();
      return;
    }
    loadAndRender(opts);
  };
  window.exportMonthlyReport_ = exportMonthlyReport_;
  bindOnce();
  loadAndRender();
}

export function refreshMonthlyReport_(opts) {
  if (_deps) window.refreshMonthlyReport_(opts);
}
