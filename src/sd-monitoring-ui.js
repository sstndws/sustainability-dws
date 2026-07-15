/**
 * SD Monitoring — Declaration Monitoring third tab.
 * Sheet: "SD Monitoring". Formula columns are display-only (never written from web).
 */

import { buildBrandedExcelSheet_ } from './excel-brand-header.js';

const SD_SHEET_KEY = 'sdMonitoring';

const SD_EDITABLE_FIELDS = [
  'MONTH', 'YEAR', 'CATEGORY', 'COMPANY NAME', 'DO NUMBER', 'CONTRACT NUMBER',
  'LOADING START DATE', 'LOADING END DATE', 'MILL NETTO',
  'UNLOADING START DATE', 'UNLOADING END DATE', 'Netto Bulking',
  'SD NUMBER', 'SD DATE',
];

const SD_FORMULA_FIELDS = [
  'LAST SD DATE', 'DAY LEFT', 'Status', 'Result', 'Risk Number',
];

const SD_ALL_FIELDS = SD_EDITABLE_FIELDS.concat(SD_FORMULA_FIELDS);

const SD_DATE_FIELDS = new Set([
  'LOADING START DATE', 'LOADING END DATE',
  'UNLOADING START DATE', 'UNLOADING END DATE',
  'LAST SD DATE', 'SD DATE',
]);

const SD_FIELD_LABELS = {
  'MONTH': 'Month',
  'YEAR': 'Year',
  'CATEGORY': 'Category',
  'COMPANY NAME': 'Company Name',
  'DO NUMBER': 'DO Number',
  'CONTRACT NUMBER': 'Contract Number',
  'LOADING START DATE': 'Loading Start Date',
  'LOADING END DATE': 'Loading End Date',
  'MILL NETTO': 'Mill Netto',
  'UNLOADING START DATE': 'Unloading Start Date',
  'UNLOADING END DATE': 'Unloading End Date',
  'Netto Bulking': 'Netto Bulking',
  'LAST SD DATE': 'Last SD Date',
  'DAY LEFT': 'Day Left',
  'Status': 'Status',
  'SD NUMBER': 'SD Number',
  'SD DATE': 'SD Date',
  'Result': 'Result',
  'Risk Number': 'Risk Number',
};

function sdNormKey_(h) {
  return String(h || '').replace(/\s+/g, ' ').trim().toUpperCase();
}

function sdIsFormulaField_(h) {
  const key = sdNormKey_(h);
  return SD_FORMULA_FIELDS.some(function(f) { return sdNormKey_(f) === key; });
}

function sdPick_(row, wanted) {
  if (!row) return '';
  if (row[wanted] != null && String(row[wanted]).trim() !== '') return row[wanted];
  const want = sdNormKey_(wanted);
  const keys = Object.keys(row);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].charAt(0) === '_') continue;
    if (sdNormKey_(keys[i]) === want) return row[keys[i]];
  }
  return '';
}

function sdLabel_(h) {
  return SD_FIELD_LABELS[h] || h;
}

function sdStripFormulaPayload_(data) {
  const out = {};
  Object.keys(data || {}).forEach(function(k) {
    if (!k || k.charAt(0) === '_' || sdIsFormulaField_(k)) return;
    out[k] = data[k];
  });
  return out;
}

export function createSdMonitoringController_(deps) {
  const {
    apiGet,
    apiPost,
    escHtml,
    dashDateFieldHtml,
    dashDateCollectValues,
    getJsPDF,
    openConfirm,
    showToast,
    gasOpts,
    mountOverlay,
    lockScroll,
    unlockScroll,
    resetOverlayScroll,
  } = deps;

  let sdData = [];
  let sdLoaded = false;
  let sdLoadPromise = null;
  let sdSearch = '';
  let sdFormMode = 'add';
  let sdFormRow = null;
  let sdFormEditRowNum_ = 0;
  let sdDetailCurrent = null;
  let sdTableDelegationBound = false;
  let sdActive = false;

  function sdCategoryPills_(raw) {
    const parts = String(raw || '')
      .split(/[;,|/]+/)
      .map(function(s) { return s.trim(); })
      .filter(Boolean);
    if (!parts.length) return '<span class="cert-pill-empty">—</span>';
    return '<div class="cert-pill-list">' + parts.map(function(part) {
      return '<span class="cert-pill sd-cat-pill">' + escHtml(part) + '</span>';
    }).join('') + '</div>';
  }

  function sdStatusBadge_(val) {
    const s = String(val || '').trim();
    if (!s) return '—';
    const low = s.toLowerCase();
    let cls = 'status-badge risk-med';
    if (/done|close|ok|pass|complete/.test(low)) cls = 'status-badge grv-closed';
    else if (/open|progress|pending|wait/.test(low)) cls = 'status-badge grv-open';
    else if (/high|fail|risk/.test(low)) cls = 'status-badge risk-high';
    return '<span class="' + cls + '"><span class="s-dot"></span>' + escHtml(s) + '</span>';
  }

  function sdDayLeftCell_(val) {
    const s = String(val == null ? '' : val).trim();
    if (!s) return '—';
    const n = parseFloat(String(s).replace(/,/g, ''));
    if (!isNaN(n) && n <= 7) {
      return '<span class="sd-day-left sd-day-left--warn">' + escHtml(s) + '</span>';
    }
    return escHtml(s);
  }

  function sdFormatDisplay_(field, val) {
    if (val == null || String(val).trim() === '') return '—';
    return String(val);
  }

  function prepareSdRow_(row) {
    const o = Object.assign({}, row);
    o._row = Number(row._row) || 0;
    const blob = SD_ALL_FIELDS.map(function(f) {
      return String(sdPick_(o, f) || '');
    }).join(' ').toLowerCase();
    o._sdSearchBlob = blob;
    return o;
  }

  function sdRegistryColumns_() {
    return [
      { label: 'Month', minWidth: 70, cell: function(d) { return escHtml(sdFormatDisplay_('MONTH', sdPick_(d, 'MONTH'))); } },
      { label: 'Year', minWidth: 70, cell: function(d) { return escHtml(sdFormatDisplay_('YEAR', sdPick_(d, 'YEAR'))); } },
      { label: 'Category', minWidth: 120, cell: function(d) { return sdCategoryPills_(sdPick_(d, 'CATEGORY')); } },
      { label: 'Company', minWidth: 160, cell: function(d) {
        return '<span class="bl-row-key">' + escHtml(sdFormatDisplay_('COMPANY NAME', sdPick_(d, 'COMPANY NAME'))) + '</span>';
      } },
      { label: 'DO No.', minWidth: 110, cell: function(d) { return escHtml(sdFormatDisplay_('DO NUMBER', sdPick_(d, 'DO NUMBER'))); } },
      { label: 'Contract', minWidth: 110, cell: function(d) { return escHtml(sdFormatDisplay_('CONTRACT NUMBER', sdPick_(d, 'CONTRACT NUMBER'))); } },
      { label: 'Mill Netto', minWidth: 90, cell: function(d) { return escHtml(sdFormatDisplay_('MILL NETTO', sdPick_(d, 'MILL NETTO'))); } },
      { label: 'Netto Bulking', minWidth: 100, cell: function(d) { return escHtml(sdFormatDisplay_('Netto Bulking', sdPick_(d, 'Netto Bulking'))); } },
      { label: 'Last SD Date', minWidth: 110, cell: function(d) { return escHtml(sdFormatDisplay_('LAST SD DATE', sdPick_(d, 'LAST SD DATE'))); } },
      { label: 'Day Left', minWidth: 80, cell: function(d) { return sdDayLeftCell_(sdPick_(d, 'DAY LEFT')); } },
      { label: 'Status', minWidth: 100, cell: function(d) { return sdStatusBadge_(sdPick_(d, 'Status')); } },
      { label: 'SD No.', minWidth: 100, cell: function(d) { return escHtml(sdFormatDisplay_('SD NUMBER', sdPick_(d, 'SD NUMBER'))); } },
      { label: 'Result', minWidth: 90, cell: function(d) { return escHtml(sdFormatDisplay_('Result', sdPick_(d, 'Result'))); } },
      { label: 'Risk No.', minWidth: 80, cell: function(d) { return escHtml(sdFormatDisplay_('Risk Number', sdPick_(d, 'Risk Number'))); } },
      { label: '', minWidth: 110, isActions: true, cell: function(d) {
        return ''
          + '<div class="row-actions">'
          + '<button type="button" class="btn-row btn-edit sd-row-edit" data-row="' + d._row + '">Edit</button>'
          + '<button type="button" class="btn-row btn-delete sd-row-delete" data-rownum="' + d._row + '">Del</button>'
          + '</div>';
      } },
    ];
  }

  function updateSdStats_() {
    const rows = sdData || [];
    const companies = {};
    let openN = 0;
    let dueN = 0;
    rows.forEach(function(r) {
      const co = String(sdPick_(r, 'COMPANY NAME') || '').trim();
      if (co) companies[co.toUpperCase()] = true;
      const st = String(sdPick_(r, 'Status') || '').toLowerCase();
      if (st && !/done|close|complete|ok|pass/.test(st)) openN++;
      const day = parseFloat(String(sdPick_(r, 'DAY LEFT') || '').replace(/,/g, ''));
      if (!isNaN(day) && day <= 7) dueN++;
    });
    const setPair = function(n, value, label) {
      const numEl = document.getElementById('bl-stat-' + n);
      const labelEl = document.getElementById('bl-stat-' + n + '-label');
      if (numEl) numEl.textContent = String(value);
      if (labelEl) labelEl.textContent = label;
    };
    setPair(1, rows.length || '0', 'Total SD');
    setPair(2, openN || '0', 'Open status');
    setPair(3, dueN || '0', 'Day left ≤ 7');
    setPair(4, Object.keys(companies).length || '0', 'Companies');
  }

  function filteredSdRows_() {
    const q = sdSearch;
    return (sdData || []).filter(function(d) {
      return !q || (d._sdSearchBlob || '').includes(q);
    });
  }

  function renderSdTable_() {
    const body = document.getElementById('blTableBody');
    const table = document.getElementById('blTable');
    if (!body) return;
    bindSdTableDelegationOnce_();
    const cols = sdRegistryColumns_();
    const headRow = document.getElementById('blTableHeadRow');
    if (headRow) {
      headRow.innerHTML = cols.map(function(col) {
        const minW = col.minWidth ? ' style="min-width:' + col.minWidth + 'px"' : '';
        return '<th' + minW + '>' + escHtml(col.label) + '</th>';
      }).join('');
    }
    const filtered = filteredSdRows_();
    if (table) table.style.display = '';
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="' + cols.length + '" style="text-align:center;padding:32px;color:#9C8A8A;">No SD records found</td></tr>';
      return;
    }
    body.innerHTML = filtered.map(function(d) {
      return '<tr class="bl-main-row sd-main-row" data-row="' + d._row + '">' + cols.map(function(col) {
        const minW = col.minWidth ? ' style="min-width:' + col.minWidth + 'px"' : '';
        return '<td' + minW + '>' + col.cell(d) + '</td>';
      }).join('') + '</tr>';
    }).join('');
  }

  function bindSdTableDelegationOnce_() {
    if (sdTableDelegationBound) return;
    const body = document.getElementById('blTableBody');
    if (!body) return;
    sdTableDelegationBound = true;
    body.addEventListener('click', function(e) {
      if (!sdActive) return;
      const editBtn = e.target.closest('.sd-row-edit');
      if (editBtn && body.contains(editBtn)) {
        e.stopPropagation();
        const rowNum = parseInt(editBtn.dataset.row, 10);
        const row = sdData.find(function(d) { return Number(d._row) === rowNum; });
        if (row) openSdFormModal_('edit', row);
        return;
      }
      const delBtn = e.target.closest('.sd-row-delete');
      if (delBtn && body.contains(delBtn)) {
        e.stopPropagation();
        openConfirm(SD_SHEET_KEY, parseInt(delBtn.dataset.rownum, 10));
        return;
      }
      const tr = e.target.closest('tr.sd-main-row');
      if (tr && body.contains(tr) && !e.target.closest('.row-actions')) {
        const rowNum = parseInt(tr.dataset.row, 10);
        const row = sdData.find(function(d) { return Number(d._row) === rowNum; });
        if (row) openSdDetailModal_(row);
      }
    });
  }

  async function fetchSdRows_() {
    const rows = await apiGet(SD_SHEET_KEY, gasOpts ? { baseUrl: gasOpts.baseUrl } : undefined);
    return (rows || []).map(prepareSdRow_);
  }

  async function loadSdData_(opts) {
    opts = opts || {};
    if (sdLoadPromise) return sdLoadPromise;
    const loading = document.getElementById('bl-loading');
    const errEl = document.getElementById('bl-error');
    const table = document.getElementById('blTable');
    if (loading && !opts.soft) loading.style.display = '';
    if (errEl) errEl.style.display = 'none';
    if (table && !opts.soft) table.style.display = 'none';
    sdLoadPromise = fetchSdRows_()
      .then(function(rows) {
        sdData = rows;
        sdLoaded = true;
        updateSdStats_();
        if (sdActive) renderSdTable_();
        if (loading) loading.style.display = 'none';
        if (table && sdActive) table.style.display = '';
      })
      .catch(function(err) {
        console.warn('[SD] Load failed:', err);
        if (loading) loading.style.display = 'none';
        if (errEl && sdActive) {
          errEl.style.display = '';
          errEl.textContent = 'Failed to load SD Monitoring: ' + (err.message || err);
        }
        if (typeof showToast === 'function') {
          showToast('Failed to load SD Monitoring: ' + (err.message || err), 'error');
        }
      })
      .finally(function() { sdLoadPromise = null; });
    return sdLoadPromise;
  }

  async function reloadSdDataSoft_() {
    try {
      sdData = await fetchSdRows_();
      sdLoaded = true;
      updateSdStats_();
      if (sdActive) renderSdTable_();
    } catch (err) {
      sdLoaded = false;
      await loadSdData_();
    }
  }

  function buildSdFormFields_(row) {
    const grid = document.getElementById('sdFormFieldsGrid');
    const compGrid = document.getElementById('sdFormComputedGrid');
    const compSec = document.getElementById('sdFormComputedSection');
    if (!grid) return;
    const data = row || {};

    grid.innerHTML = SD_EDITABLE_FIELDS.map(function(f) {
      const label = sdLabel_(f);
      const val = sdPick_(data, f);
      if (SD_DATE_FIELDS.has(f)) {
        return dashDateFieldHtml(f, val, { label: label });
      }
      if (f === 'CATEGORY') {
        return ''
          + '<div class="form-field">'
          + '<label>' + escHtml(label) + '</label>'
          + '<input type="text" data-field="' + escHtml(f) + '" value="' + escHtml(String(val || '')) + '" placeholder="e.g. CPO; PK; POME">'
          + '<p class="bl-form-hint" style="margin-top:6px;">Separate multiple categories with ; or ,</p>'
          + '</div>';
      }
      return ''
        + '<div class="form-field">'
        + '<label>' + escHtml(label) + '</label>'
        + '<input type="text" data-field="' + escHtml(f) + '" value="' + escHtml(String(val || '')) + '" placeholder="' + escHtml(label) + '">'
        + '</div>';
    }).join('');

    if (compGrid && compSec) {
      if (row) {
        compSec.hidden = false;
        compGrid.innerHTML = SD_FORMULA_FIELDS.map(function(f) {
          const val = sdPick_(data, f);
          return ''
            + '<div class="form-field mp-field--computed">'
            + '<label>' + escHtml(sdLabel_(f)) + ' <span class="bl-form-hint">auto</span></label>'
            + '<input type="text" value="' + escHtml(String(val || '—')) + '" readonly disabled>'
            + '</div>';
        }).join('');
      } else {
        compSec.hidden = true;
        compGrid.innerHTML = '';
      }
    }
  }

  function collectSdFormData_() {
    const body = document.getElementById('sdFormBody');
    if (!body) return {};
    const data = {};
    body.querySelectorAll('input[data-field], textarea[data-field], select[data-field]').forEach(function(el) {
      const f = el.getAttribute('data-field');
      if (!f || sdIsFormulaField_(f) || el.disabled) return;
      data[f] = el.value;
    });
    if (typeof dashDateCollectValues === 'function') {
      Object.assign(data, dashDateCollectValues(body));
    }
    return sdStripFormulaPayload_(data);
  }

  async function openSdFormModal_(mode, row) {
    const overlay = mountOverlay ? mountOverlay('sdFormOverlay') : document.getElementById('sdFormOverlay');
    if (!overlay) return;
    sdFormMode = mode || 'add';
    sdFormRow = row || null;
    sdFormEditRowNum_ = (sdFormMode === 'edit' && row && row._row != null)
      ? Number(row._row)
      : 0;
    if (sdFormMode === 'edit' && !(sdFormEditRowNum_ >= 2)) {
      console.warn('[SD] Edit opened without a valid sheet row number', row);
      sdFormMode = 'add';
      sdFormEditRowNum_ = 0;
    }
    const titleEl = document.getElementById('sdFormTitle');
    const saveBtn = document.getElementById('sdFormSave');
    if (titleEl) titleEl.textContent = sdFormMode === 'edit' ? 'Edit SD Record' : 'Add SD Record';
    if (saveBtn) saveBtn.textContent = 'Save SD';
    buildSdFormFields_(row);
    if (typeof lockScroll === 'function') lockScroll();
    document.body.classList.add('bl-form-open');
    overlay.classList.add('active');
    if (typeof resetOverlayScroll === 'function') {
      resetOverlayScroll(overlay, document.getElementById('sdFormBody'));
    }
  }

  function closeSdFormModal_() {
    const overlay = document.getElementById('sdFormOverlay');
    if (overlay) overlay.classList.remove('active');
    document.body.classList.remove('bl-form-open');
    if (typeof unlockScroll === 'function') unlockScroll();
    sdFormRow = null;
    sdFormEditRowNum_ = 0;
    sdFormMode = 'add';
  }

  async function saveSdForm_() {
    const saveBtn = document.getElementById('sdFormSave');
    if (saveBtn && saveBtn.disabled) return;
    const editRowNum = (sdFormEditRowNum_ >= 2)
      ? sdFormEditRowNum_
      : (sdFormMode === 'edit' && sdFormRow && sdFormRow._row != null
        ? Number(sdFormRow._row)
        : 0);
    const payload = collectSdFormData_();
    if (!String(payload['COMPANY NAME'] || '').trim() && !String(payload['DO NUMBER'] || '').trim()) {
      alert('Enter at least Company Name or DO Number.');
      return;
    }
    if (saveBtn) {
      saveBtn.disabled = true;
      saveBtn.textContent = 'Saving…';
    }
    try {
      const opts = gasOpts ? { baseUrl: gasOpts.baseUrl } : undefined;
      if (editRowNum >= 2) {
        await apiPost({
          action: 'update',
          sheet: SD_SHEET_KEY,
          row: editRowNum,
          data: payload,
        }, opts);
      } else {
        await apiPost({
          action: 'add',
          sheet: SD_SHEET_KEY,
          data: payload,
        }, opts);
      }
      closeSdFormModal_();
      await reloadSdDataSoft_();
      if (typeof showToast === 'function') showToast('SD record saved.', 'success');
    } catch (err) {
      alert('Failed to save SD: ' + (err.message || err));
    } finally {
      if (saveBtn) {
        saveBtn.disabled = false;
        saveBtn.textContent = 'Save SD';
      }
    }
  }

  function openSdDetailModal_(row) {
    const overlay = mountOverlay ? mountOverlay('sdDetailOverlay') : document.getElementById('sdDetailOverlay');
    const bodyEl = document.getElementById('sdDetailBody');
    if (!overlay || !bodyEl || !row) return;
    sdDetailCurrent = row;

    const titleEl = document.getElementById('sdDetailTitle');
    const subEl = document.getElementById('sdDetailSubtitle');
    const company = sdPick_(row, 'COMPANY NAME') || '—';
    const sdNo = sdPick_(row, 'SD NUMBER') || sdPick_(row, 'DO NUMBER') || '';
    if (titleEl) titleEl.textContent = 'SD Detail';
    if (subEl) subEl.textContent = company + (sdNo ? ' · ' + sdNo : '');

    function cellHtml(f) {
      const val = sdPick_(row, f);
      if (f === 'CATEGORY') return sdCategoryPills_(val);
      if (f === 'Status') return sdStatusBadge_(val);
      if (f === 'DAY LEFT') return sdDayLeftCell_(val);
      return escHtml(sdFormatDisplay_(f, val));
    }

    const inputGrid = SD_EDITABLE_FIELDS.map(function(f) {
      return ''
        + '<div class="bl-detail-field">'
        + '<div class="bl-detail-label">' + escHtml(sdLabel_(f)) + '</div>'
        + '<div class="bl-detail-value">' + cellHtml(f) + '</div>'
        + '</div>';
    }).join('');

    const formulaGrid = SD_FORMULA_FIELDS.map(function(f) {
      return ''
        + '<div class="bl-detail-field bl-detail-field--computed">'
        + '<div class="bl-detail-label">' + escHtml(sdLabel_(f)) + ' <span class="bl-form-hint">formula</span></div>'
        + '<div class="bl-detail-value">' + cellHtml(f) + '</div>'
        + '</div>';
    }).join('');

    bodyEl.innerHTML = ''
      + '<div class="bl-detail-section"><div class="bl-detail-section-title">Record</div><div class="bl-detail-grid">' + inputGrid + '</div></div>'
      + '<div class="bl-detail-section"><div class="bl-detail-section-title">Sheet formulas</div><div class="bl-detail-grid">' + formulaGrid + '</div></div>';

    document.body.classList.add('bl-detail-open');
    if (typeof lockScroll === 'function') lockScroll();
    overlay.classList.add('active');
    if (typeof resetOverlayScroll === 'function') resetOverlayScroll(overlay, bodyEl);
  }

  function closeSdDetailModal_() {
    const overlay = document.getElementById('sdDetailOverlay');
    if (overlay) overlay.classList.remove('active');
    document.body.classList.remove('bl-detail-open');
    if (typeof unlockScroll === 'function') unlockScroll();
    sdDetailCurrent = null;
    const bodyEl = document.getElementById('sdDetailBody');
    if (bodyEl) bodyEl.innerHTML = '';
  }

  function sdExportRowsToExcel_(rows, filename) {
    const XLSX = window.XLSX;
    if (!XLSX) {
      alert('Excel library not loaded. Refresh the page.');
      return;
    }
    const headers = SD_ALL_FIELDS.map(sdLabel_);
    const bodyRows = (rows || []).map(function(r) {
      return SD_ALL_FIELDS.map(function(f) {
        const v = sdPick_(r, f);
        return v == null ? '' : String(v);
      });
    });
    const ws = buildBrandedExcelSheet_(XLSX, headers, bodyRows, { headerFill: '8B1A1A' });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'SD Monitoring');
    XLSX.writeFile(wb, filename || 'SD-Monitoring.xlsx', { cellStyles: true });
  }

  function exportSdExcelList_() {
    const rows = filteredSdRows_();
    if (!rows.length) {
      alert('No SD records to export.');
      return;
    }
    const stamp = new Date().toISOString().slice(0, 10);
    sdExportRowsToExcel_(rows, 'SD-Monitoring-' + stamp + '.xlsx');
  }

  function exportSdExcelOne_(row) {
    if (!row) return;
    const name = String(sdPick_(row, 'COMPANY NAME') || sdPick_(row, 'SD NUMBER') || 'SD')
      .replace(/[^\w\-]+/g, '_')
      .slice(0, 40);
    sdExportRowsToExcel_([row], 'SD-' + name + '.xlsx');
  }

  function exportSdPdfOne_(row) {
    if (!row) return;
    let JsPDFLib;
    try {
      JsPDFLib = typeof getJsPDF === 'function' ? getJsPDF() : null;
    } catch (e) {
      JsPDFLib = null;
    }
    if (!JsPDFLib) {
      alert('PDF library not loaded. Refresh the page.');
      return;
    }
    const doc = new JsPDFLib({ unit: 'pt', format: 'a4' });
    const margin = 40;
    let y = margin;
    const pageW = doc.internal.pageSize.getWidth();
    const maxW = pageW - margin * 2;

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(16);
    doc.setTextColor(139, 26, 26);
    doc.text('SD Monitoring Detail', margin, y);
    y += 22;
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.setTextColor(80, 60, 60);
    doc.text(
      String(sdPick_(row, 'COMPANY NAME') || '—')
        + (sdPick_(row, 'SD NUMBER') ? '  ·  ' + sdPick_(row, 'SD NUMBER') : ''),
      margin,
      y
    );
    y += 18;
    doc.setDrawColor(200, 180, 180);
    doc.line(margin, y, pageW - margin, y);
    y += 16;

    function writeField(label, value) {
      if (y > doc.internal.pageSize.getHeight() - 50) {
        doc.addPage();
        y = margin;
      }
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(9);
      doc.setTextColor(120, 90, 90);
      doc.text(label, margin, y);
      y += 12;
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(11);
      doc.setTextColor(30, 20, 20);
      const lines = doc.splitTextToSize(String(value == null || value === '' ? '—' : value), maxW);
      doc.text(lines, margin, y);
      y += lines.length * 14 + 8;
    }

    SD_EDITABLE_FIELDS.forEach(function(f) {
      writeField(sdLabel_(f), sdPick_(row, f));
    });
    y += 6;
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.setTextColor(139, 26, 26);
    doc.text('Sheet formulas', margin, y);
    y += 16;
    SD_FORMULA_FIELDS.forEach(function(f) {
      writeField(sdLabel_(f), sdPick_(row, f));
    });

    const name = String(sdPick_(row, 'COMPANY NAME') || sdPick_(row, 'SD NUMBER') || 'SD')
      .replace(/[^\w\-]+/g, '_')
      .slice(0, 40);
    doc.save('SD-' + name + '.pdf');
  }

  function setSdSearch_(q) {
    sdSearch = String(q || '').toLowerCase().trim();
    if (sdActive) renderSdTable_();
  }

  function setSdActive_(on) {
    sdActive = !!on;
    if (sdActive) {
      const body = document.getElementById('blTableBody');
      if (body && !sdLoaded) body.innerHTML = '';
      updateSdStats_();
      if (sdLoaded) renderSdTable_();
      else loadSdData_();
    }
  }

  function bindSdUi_() {
    const formOverlay = document.getElementById('sdFormOverlay');
    const detailOverlay = document.getElementById('sdDetailOverlay');
    document.getElementById('sdFormClose')?.addEventListener('click', closeSdFormModal_);
    document.getElementById('sdFormCancel')?.addEventListener('click', closeSdFormModal_);
    document.getElementById('sdFormSave')?.addEventListener('click', saveSdForm_);
    formOverlay?.addEventListener('click', function(e) {
      if (e.target === formOverlay) closeSdFormModal_();
    });

    document.getElementById('sdDetailClose')?.addEventListener('click', closeSdDetailModal_);
    document.getElementById('sdDetailCloseBtn')?.addEventListener('click', closeSdDetailModal_);
    document.getElementById('sdDetailExportBtn')?.addEventListener('click', function() {
      if (sdDetailCurrent) exportSdExcelOne_(sdDetailCurrent);
    });
    document.getElementById('sdDetailPdfBtn')?.addEventListener('click', function() {
      if (sdDetailCurrent) exportSdPdfOne_(sdDetailCurrent);
    });
    document.getElementById('sdDetailEditBtn')?.addEventListener('click', function() {
      const row = sdDetailCurrent;
      closeSdDetailModal_();
      if (row) openSdFormModal_('edit', row);
    });
    document.getElementById('sdDetailDeleteBtn')?.addEventListener('click', function() {
      const row = sdDetailCurrent;
      if (!row || !row._row) return;
      closeSdDetailModal_();
      openConfirm(SD_SHEET_KEY, row._row);
    });
    detailOverlay?.addEventListener('click', function(e) {
      if (e.target === detailOverlay) closeSdDetailModal_();
    });
  }

  bindSdUi_();

  return {
    isActive: function() { return sdActive; },
    setActive: setSdActive_,
    setSearch: setSdSearch_,
    load: loadSdData_,
    reloadSoft: reloadSdDataSoft_,
    openAdd: function() { openSdFormModal_('add', null); },
    exportExcel: exportSdExcelList_,
    render: renderSdTable_,
    SHEET_KEY: SD_SHEET_KEY,
  };
}
