/**
 * Risk Analysis & Mitigation — list + add/edit (Google Sheet tab).
 */
import './risk-analysis-mitigation.css';

const SHEET_KEY = 'riskAnalysisMitigation';

const F = {
  no: 'No',
  kategori: 'Kategori',
  subIsu: 'Sub-Isu / Topik Risiko',
  level: 'Level Risiko',
  prob: 'Probabilitas',
  sumber: 'Sumber Risiko',
  dampak: 'Dampak ke Perusahaan',
  regulasi: 'Regulasi / Standar Terkait',
  mitigasi: 'Rencana Penanganan / Mitigasi',
  prioritas: 'Prioritas',
  status: 'Status',
};

const DEFAULT_KATEGORI = [
  'NDPE & Deforestasi',
  'EUDR & Regulasi Global',
  'Traceability & Supply Chain',
  'Sertifikasi & Compliance',
  'Sosial & Grievance',
  'Lingkungan & Iklim',
];

const LEVEL_OPTS = ['TINGGI', 'SEDANG', 'RENDAH'];
const PROB_OPTS = ['Sangat Tinggi', 'Tinggi', 'Sedang', 'Rendah', 'Sangat Rendah'];
const PRIO_OPTS = [
  'Segera (1–3 Bln)',
  'Jangka Pendek (3–6 Bln)',
  'Jangka Menengah (6–12 Bln)',
  'Jangka Panjang (>12 Bln)',
];
const STATUS_OPTS = ['Belum Dimulai', 'Monitoring', 'Sedang Berjalan', 'Selesai', 'Ditunda'];

let _deps = null;
let _rows = [];
let _search = '';
let _filterKat = '';
let _filterLevel = '';
let _filterStatus = '';
let _modalMode = 'view';
let _modalRow = null;
let _formFromView = false;
let _loadSeq = 0;
let _uiBound = false;
let _saving = false;

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function toast(msg, kind) {
  if (typeof window.showSddToast === 'function') window.showSddToast(msg, kind || 'info');
}

function normKey_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normLevel_(s) {
  const t = normKey_(s);
  if (t.indexOf('tinggi') !== -1 && t.indexOf('sangat') === -1) return 'TINGGI';
  if (t.indexOf('sedang') !== -1) return 'SEDANG';
  if (t.indexOf('rendah') !== -1) return 'RENDAH';
  return String(s || '').trim().toUpperCase() || '—';
}

function levelPillHtml_(val) {
  const n = normKey_(val);
  let cls = 'ram-level-pill--other';
  if (n.indexOf('tinggi') !== -1) cls = 'ram-level-pill--high';
  else if (n.indexOf('sedang') !== -1) cls = 'ram-level-pill--med';
  else if (n.indexOf('rendah') !== -1) cls = 'ram-level-pill--low';
  const label = String(val || '—').trim() || '—';
  return '<span class="ram-level-pill ' + cls + '">' + esc(label) + '</span>';
}

function prioPillHtml_(val) {
  const t = String(val || '');
  const n = normKey_(t);
  let cls = 'ram-prio-pill--other';
  let icon = '📌';
  if (n.indexOf('segera') !== -1 || n.indexOf('1') !== -1) {
    cls = 'ram-prio-pill--urgent';
    icon = '⚡';
  } else if (n.indexOf('pendek') !== -1 || n.indexOf('3–6') !== -1 || n.indexOf('3-6') !== -1) {
    cls = 'ram-prio-pill--short';
    icon = '📋';
  } else if (n.indexOf('menengah') !== -1 || n.indexOf('6–12') !== -1 || n.indexOf('6-12') !== -1) {
    cls = 'ram-prio-pill--mid';
    icon = '📌';
  }
  const short = t.replace(/^[^\w]+/, '').trim() || '—';
  return '<span class="ram-prio-pill ' + cls + '" title="' + esc(t) + '">' + icon + ' ' + esc(short) + '</span>';
}

function statusPillHtml_(val) {
  const n = normKey_(val);
  let cls = 'ram-status-pill--todo';
  if (n.indexOf('monitor') !== -1) cls = 'ram-status-pill--watch';
  else if (n.indexOf('berjalan') !== -1 || n.indexOf('progress') !== -1) cls = 'ram-status-pill--progress';
  else if (n.indexOf('selesai') !== -1) cls = 'ram-status-pill--done';
  return '<span class="ram-status-pill ' + cls + '">' + esc(String(val || '—').trim() || '—') + '</span>';
}

function regulasiPillsHtml_(raw) {
  const s = String(raw || '').trim();
  if (!s) return '<span class="cert-pill-empty">—</span>';
  const parts = s.split(/[;,|]+/).map(function(p) { return p.trim(); }).filter(Boolean);
  if (!parts.length) return '<span class="cert-pill-empty">—</span>';
  return '<div class="ram-pill-list">' + parts.slice(0, 4).map(function(p) {
    return '<span class="ram-reg-pill cert-pill" title="' + esc(p) + '">' + esc(p) + '</span>';
  }).join('') + (parts.length > 4 ? '<span class="ram-reg-pill">+' + (parts.length - 4) + '</span>' : '') + '</div>';
}

function uniqueSorted_(arr) {
  const set = new Set();
  arr.forEach(function(v) {
    const s = String(v || '').trim();
    if (s) set.add(s);
  });
  return Array.from(set).sort(function(a, b) { return a.localeCompare(b); });
}

function collectFilterOptions_() {
  const kat = uniqueSorted_(DEFAULT_KATEGORI.concat(_rows.map(function(r) { return r[F.kategori]; })));
  const lvl = uniqueSorted_(LEVEL_OPTS.concat(_rows.map(function(r) { return r[F.level]; })));
  const st = uniqueSorted_(STATUS_OPTS.concat(_rows.map(function(r) { return r[F.status]; })));
  return { kat: kat, lvl: lvl, st: st };
}

function fillSelect_(el, options, current, includeEmpty) {
  if (!el) return;
  const cur = String(current || '').trim();
  let html = includeEmpty ? '<option value="">— Select —</option>' : '';
  options.forEach(function(opt) {
    html += '<option value="' + esc(opt) + '"' + (opt === cur ? ' selected' : '') + '>' + esc(opt) + '</option>';
  });
  if (cur && options.indexOf(cur) === -1) {
    html += '<option value="' + esc(cur) + '" selected>' + esc(cur) + ' (existing)</option>';
  }
  el.innerHTML = html;
}

function fillFilterSelect_(el, options, allLabel) {
  if (!el) return;
  let html = '<option value="">' + esc(allLabel || 'All') + '</option>';
  options.forEach(function(opt) {
    html += '<option value="' + esc(opt) + '">' + esc(opt) + '</option>';
  });
  el.innerHTML = html;
}

function getFilteredRows_() {
  const q = normKey_(_search);
  return _rows.filter(function(r) {
    if (_filterKat && String(r[F.kategori] || '').trim() !== _filterKat) return false;
    if (_filterLevel && normLevel_(r[F.level]) !== normLevel_(_filterLevel)) return false;
    if (_filterStatus && String(r[F.status] || '').trim() !== _filterStatus) return false;
    if (!q) return true;
    const blob = [
      r[F.kategori], r[F.subIsu], r[F.level], r[F.prob], r[F.prioritas], r[F.status],
      r[F.regulasi], r[F.sumber], r[F.dampak], r[F.mitigasi],
    ].join(' ');
    return normKey_(blob).indexOf(q) !== -1;
  });
}

function updateStats_() {
  const total = _rows.length;
  let high = 0;
  let urgent = 0;
  let active = 0;
  _rows.forEach(function(r) {
    const lv = normKey_(r[F.level]);
    if (lv.indexOf('tinggi') !== -1) high++;
    const pr = normKey_(r[F.prioritas]);
    if (pr.indexOf('segera') !== -1) urgent++;
    const st = normKey_(r[F.status]);
    if (st.indexOf('berjalan') !== -1) active++;
  });
  const elTotal = document.getElementById('ramStatTotal');
  const elHigh = document.getElementById('ramStatHigh');
  const elUrgent = document.getElementById('ramStatUrgent');
  const elActive = document.getElementById('ramStatActive');
  if (elTotal) elTotal.textContent = String(total);
  if (elHigh) elHigh.textContent = String(high);
  if (elUrgent) elUrgent.textContent = String(urgent);
  if (elActive) elActive.textContent = String(active);
}

function syncFilterDropdowns_() {
  const opts = collectFilterOptions_();
  const fk = document.getElementById('ramFilterKategori');
  const fl = document.getElementById('ramFilterLevel');
  const fs = document.getElementById('ramFilterStatus');
  fillFilterSelect_(fk, opts.kat, 'All categories');
  fillFilterSelect_(fl, opts.lvl, 'All levels');
  fillFilterSelect_(fs, opts.st, 'All statuses');
  if (fk) fk.value = _filterKat;
  if (fl) fl.value = _filterLevel;
  if (fs) fs.value = _filterStatus;
}

function renderTable_() {
  const body = document.getElementById('ramTableBody');
  if (!body) return;
  const filtered = getFilteredRows_();
  if (!filtered.length) {
    body.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:36px;color:#9C8A8A;">'
      + (_rows.length ? 'No rows match filters.' : 'No risk entries yet — click <strong>Add risk</strong>.')
      + '</td></tr>';
    return;
  }
  body.innerHTML = filtered.map(function(r) {
    const rowNum = r._row;
    const snippet = String(r[F.mitigasi] || r[F.dampak] || '').trim();
    return '<tr class="ram-row-click" data-ram-row="' + esc(String(rowNum)) + '">'
      + '<td>' + esc(r[F.no] != null ? r[F.no] : '—') + '</td>'
      + '<td><span class="ram-kategori-pill" title="' + esc(r[F.kategori]) + '">' + esc(r[F.kategori] || '—') + '</span></td>'
      + '<td><div class="ram-issue-title">' + esc(r[F.subIsu] || '—') + '</div>'
      + (snippet ? '<div class="ram-issue-snippet">' + esc(snippet) + '</div>' : '') + '</td>'
      + '<td>' + levelPillHtml_(r[F.level]) + '</td>'
      + '<td><span class="ram-prob-text">' + esc(r[F.prob] || '—') + '</span></td>'
      + '<td>' + prioPillHtml_(r[F.prioritas]) + '</td>'
      + '<td>' + statusPillHtml_(r[F.status]) + '</td>'
      + '<td>' + regulasiPillsHtml_(r[F.regulasi]) + '</td>'
      + '<td class="ram-th-act"><button type="button" class="btn-sm btn-outline ram-btn-view" data-ram-row="' + esc(String(rowNum)) + '">View</button></td>'
      + '</tr>';
  }).join('');
}

function rowBySheetRow_(rowNum) {
  const n = Number(rowNum);
  return _rows.find(function(r) { return Number(r._row) === n; }) || null;
}

function nextNoPreview_() {
  let max = 0;
  _rows.forEach(function(r) {
    const n = Number(r[F.no]);
    if (!isNaN(n) && n > max) max = n;
  });
  return max + 1;
}

function ensureRamModalMounted_() {
  const overlay = document.getElementById('ramModalOverlay');
  if (overlay && overlay.parentElement !== document.body) {
    document.body.appendChild(overlay);
  }
}

function syncModalFooter_(ui) {
  const btnClose = document.getElementById('ramModalBtnClose');
  const btnEdit = document.getElementById('ramModalBtnEdit');
  const btnCancel = document.getElementById('ramModalBtnCancel');
  const btnSave = document.getElementById('ramModalBtnSave');
  const isView = ui === 'view';
  const isForm = ui === 'add' || ui === 'edit';
  if (btnClose) btnClose.hidden = !isView;
  if (btnEdit) btnEdit.hidden = !isView;
  if (btnCancel) btnCancel.hidden = !isForm;
  if (btnSave) btnSave.hidden = !isForm;
  if (ui === 'closed') {
    if (btnClose) btnClose.hidden = true;
    if (btnEdit) btnEdit.hidden = true;
    if (btnCancel) btnCancel.hidden = true;
    if (btnSave) btnSave.hidden = true;
  }
}

function populateFormFields_(row) {
  const opts = collectFilterOptions_();
  fillSelect_(document.getElementById('ramFieldKategori'), opts.kat, row ? row[F.kategori] : '', true);
  fillSelect_(document.getElementById('ramFieldLevel'), LEVEL_OPTS, row ? row[F.level] : 'TINGGI', true);
  fillSelect_(document.getElementById('ramFieldProb'), PROB_OPTS, row ? row[F.prob] : 'Tinggi', true);
  fillSelect_(document.getElementById('ramFieldPrioritas'), PRIO_OPTS, row ? row[F.prioritas] : PRIO_OPTS[0], true);
  fillSelect_(document.getElementById('ramFieldStatus'), STATUS_OPTS, row ? row[F.status] : STATUS_OPTS[0], true);

  const fNo = document.getElementById('ramFieldNo');
  if (fNo) fNo.value = row ? String(row[F.no] != null ? row[F.no] : '') : 'Auto';
  document.getElementById('ramFieldSubIsu').value = row ? (row[F.subIsu] || '') : '';
  document.getElementById('ramFieldSumber').value = row ? (row[F.sumber] || '') : '';
  document.getElementById('ramFieldDampak').value = row ? (row[F.dampak] || '') : '';
  document.getElementById('ramFieldRegulasi').value = row ? (row[F.regulasi] || '') : '';
  document.getElementById('ramFieldMitigasi').value = row ? (row[F.mitigasi] || '') : '';
}

function renderViewPanel_(row) {
  const panel = document.getElementById('ramViewPanel');
  if (!panel || !row) return;
  const no = row[F.no] != null ? row[F.no] : '—';
  panel.innerHTML =
    '<div class="ram-view-head">'
    + '<span class="ram-kategori-pill">' + esc(row[F.kategori] || '—') + '</span>'
    + '<h4>' + esc(row[F.subIsu] || '—') + '</h4>'
    + '<div class="ram-view-meta">'
    + levelPillHtml_(row[F.level])
    + statusPillHtml_(row[F.status])
    + prioPillHtml_(row[F.prioritas])
    + '</div></div>'
    + '<dl class="ram-view-dl">'
    + '<dt>No</dt><dd>' + esc(no) + '</dd>'
    + '<dt>Probability</dt><dd>' + esc(row[F.prob] || '—') + '</dd>'
    + '<dt>Risk source</dt><dd class="ram-view-dd--block">' + esc(row[F.sumber] || '—') + '</dd>'
    + '<dt>Business impact</dt><dd class="ram-view-dd--block">' + esc(row[F.dampak] || '—') + '</dd>'
    + '<dt>Regulation / standard</dt><dd>' + regulasiPillsHtml_(row[F.regulasi]) + '</dd>'
    + '<dt>Mitigation plan</dt><dd class="ram-view-dd--block">' + esc(row[F.mitigasi] || '—') + '</dd>'
    + '</dl>';
}

function showModalPanels_(mode) {
  const viewPanel = document.getElementById('ramViewPanel');
  const formPanel = document.getElementById('ramFormPanel');
  const dialog = document.getElementById('ramModalDialog');
  const isView = mode === 'view';
  const isForm = mode === 'add' || mode === 'edit';
  if (viewPanel) {
    viewPanel.hidden = !isView;
    if (isForm) viewPanel.innerHTML = '';
  }
  if (formPanel) formPanel.hidden = !isForm;
  if (dialog) {
    dialog.classList.toggle('ram-modal--view', isView);
    dialog.classList.toggle('ram-modal--form', isForm);
  }
}

function openModalShell_() {
  ensureRamModalMounted_();
  const overlay = document.getElementById('ramModalOverlay');
  if (overlay) {
    overlay.classList.add('is-open');
    overlay.setAttribute('aria-hidden', 'false');
  }
  document.body.classList.add('ram-modal-open');
}

function openViewModal_(row) {
  if (!row) return;
  _modalMode = 'view';
  _modalRow = row;
  _formFromView = false;
  const title = document.getElementById('ramModalTitle');
  const sub = document.getElementById('ramModalSub');
  if (title) title.textContent = 'View risk entry';
  if (sub) sub.textContent = 'Review details below. Use Edit to change this record.';
  renderViewPanel_(row);
  showModalPanels_('view');
  syncModalFooter_('view');
  openModalShell_();
}

function openFormModal_(mode, row, fromView) {
  _modalMode = mode;
  _modalRow = row || null;
  _formFromView = !!fromView;
  const title = document.getElementById('ramModalTitle');
  const sub = document.getElementById('ramModalSub');
  if (title) {
    title.textContent = mode === 'add' ? 'Add risk entry' : 'Edit risk entry';
  }
  if (sub) {
    sub.textContent = mode === 'add'
      ? 'Sequence number is assigned automatically. Fields marked * are required.'
      : 'Update fields and save. Required fields are marked *.';
  }
  populateFormFields_(row);
  showModalPanels_(mode);
  syncModalFooter_(mode);
  openModalShell_();
  const subIsu = document.getElementById('ramFieldSubIsu');
  if (subIsu) setTimeout(function() { subIsu.focus(); }, 80);
}

function closeModal_() {
  const overlay = document.getElementById('ramModalOverlay');
  if (overlay) {
    overlay.classList.remove('is-open');
    overlay.setAttribute('aria-hidden', 'true');
  }
  document.body.classList.remove('ram-modal-open');
  const viewPanel = document.getElementById('ramViewPanel');
  const formPanel = document.getElementById('ramFormPanel');
  const dialog = document.getElementById('ramModalDialog');
  if (viewPanel) {
    viewPanel.hidden = true;
    viewPanel.innerHTML = '';
  }
  if (formPanel) formPanel.hidden = true;
  if (dialog) {
    dialog.classList.remove('ram-modal--view', 'ram-modal--form');
  }
  syncModalFooter_('closed');
  _modalMode = '';
  _modalRow = null;
  _formFromView = false;
}

function cancelFormModal_() {
  if (_formFromView && _modalRow) {
    openViewModal_(_modalRow);
    return;
  }
  closeModal_();
}

function readFormPayload_() {
  const kat = String(document.getElementById('ramFieldKategori').value || '').trim();
  const sub = String(document.getElementById('ramFieldSubIsu').value || '').trim();
  if (!kat || !sub) {
    throw new Error('Category and sub-issue are required.');
  }
  const payload = {};
  payload[F.kategori] = kat;
  payload[F.subIsu] = sub;
  payload[F.level] = String(document.getElementById('ramFieldLevel').value || '').trim();
  payload[F.prob] = String(document.getElementById('ramFieldProb').value || '').trim();
  payload[F.prioritas] = String(document.getElementById('ramFieldPrioritas').value || '').trim();
  payload[F.status] = String(document.getElementById('ramFieldStatus').value || '').trim();
  payload[F.sumber] = String(document.getElementById('ramFieldSumber').value || '').trim();
  payload[F.dampak] = String(document.getElementById('ramFieldDampak').value || '').trim();
  payload[F.regulasi] = String(document.getElementById('ramFieldRegulasi').value || '').trim();
  payload[F.mitigasi] = String(document.getElementById('ramFieldMitigasi').value || '').trim();
  if (_modalMode === 'edit' && _modalRow) {
    payload[F.no] = _modalRow[F.no];
  }
  return payload;
}

async function saveModal_() {
  if (_saving || !_deps) return;
  let payload;
  try {
    payload = readFormPayload_();
  } catch (err) {
    toast(err.message || String(err), 'error');
    return;
  }
  const saveBtn = document.getElementById('ramModalBtnSave');
  _saving = true;
  if (_deps.dashSetButtonBusy_) _deps.dashSetButtonBusy_(saveBtn, 'Saving…');
  try {
    if (_modalMode === 'edit' && _modalRow && _modalRow._row) {
      await _deps.apiPost({ action: 'update', sheet: SHEET_KEY, row: _modalRow._row, data: payload });
      Object.assign(_modalRow, payload);
      toast('Risk entry updated.', 'success');
    } else {
      const res = await _deps.apiPost({ action: 'add', sheet: SHEET_KEY, data: payload });
      toast('Risk entry added.', 'success');
      await loadRamData_(true);
      if (res && res.row) {
        const hit = rowBySheetRow_(res.row);
        if (hit) _modalRow = hit;
      }
    }
    closeModal_();
    updateStats_();
    syncFilterDropdowns_();
    renderTable_();
  } catch (err) {
    toast('Save failed: ' + (err.message || err), 'error');
  } finally {
    _saving = false;
    if (_deps.dashClearButtonBusy_) _deps.dashClearButtonBusy_(saveBtn);
  }
}

function setLoading_(on) {
  const el = document.getElementById('ramLoading');
  const wrap = document.getElementById('ramTableWrap');
  const err = document.getElementById('ramError');
  if (el) {
    if (on) {
      el.hidden = false;
      el.setAttribute('aria-hidden', 'false');
    } else {
      el.hidden = true;
      el.setAttribute('aria-hidden', 'true');
    }
  }
  if (wrap) wrap.classList.toggle('is-loading', !!on);
  if (on && err) {
    err.hidden = true;
    err.textContent = '';
  }
}

function showError_(msg) {
  const err = document.getElementById('ramError');
  if (!err) return;
  if (!msg) {
    err.hidden = true;
    err.textContent = '';
    return;
  }
  err.hidden = false;
  err.textContent = msg;
}

async function loadRamData_(force) {
  if (!_deps) return;
  const seq = ++_loadSeq;
  setLoading_(true);
  try {
    const raw = await _deps.apiGet(SHEET_KEY, force ? { bustCache: true } : {});
    if (seq !== _loadSeq) return;
    _rows = Array.isArray(raw) ? raw.slice() : [];
    _rows.sort(function(a, b) {
      const na = Number(a[F.no]);
      const nb = Number(b[F.no]);
      if (!isNaN(na) && !isNaN(nb) && na !== nb) return na - nb;
      return Number(a._row || 0) - Number(b._row || 0);
    });
    showError_('');
    updateStats_();
    syncFilterDropdowns_();
    renderTable_();
  } catch (e) {
    if (seq !== _loadSeq) return;
    const msg = e.message || String(e);
    if (/Sheet key not found|Tab not found|Risk Analysis/i.test(msg)) {
      showError_('Could not load the risk register. Check that the backend is deployed and the Risk Analysis & Mitigation data source is available.');
    } else {
      showError_('Failed to load data: ' + msg);
    }
    _rows = [];
    renderTable_();
  } finally {
    if (seq === _loadSeq) setLoading_(false);
  }
}

function bindUiOnce_() {
  if (_uiBound) return;
  _uiBound = true;

  document.getElementById('ramBtnAdd')?.addEventListener('click', function() {
    openFormModal_('add', null, false);
  });
  document.getElementById('ramBtnRefresh')?.addEventListener('click', function() {
    const btn = document.getElementById('ramBtnRefresh');
    if (_deps && _deps.dashSetButtonBusy_) _deps.dashSetButtonBusy_(btn, '…');
    loadRamData_(true).finally(function() {
      if (_deps && _deps.dashClearButtonBusy_) _deps.dashClearButtonBusy_(btn);
    });
  });

  document.getElementById('ramSearch')?.addEventListener('input', function(e) {
    _search = e.target.value || '';
    renderTable_();
  });

  document.getElementById('ramFilterKategori')?.addEventListener('change', function(e) {
    _filterKat = e.target.value || '';
    renderTable_();
  });
  document.getElementById('ramFilterLevel')?.addEventListener('change', function(e) {
    _filterLevel = e.target.value || '';
    renderTable_();
  });
  document.getElementById('ramFilterStatus')?.addEventListener('change', function(e) {
    _filterStatus = e.target.value || '';
    renderTable_();
  });

  document.getElementById('ramModalClose')?.addEventListener('click', closeModal_);
  document.getElementById('ramModalBtnClose')?.addEventListener('click', closeModal_);
  document.getElementById('ramModalBtnCancel')?.addEventListener('click', cancelFormModal_);
  document.getElementById('ramModalBtnEdit')?.addEventListener('click', function() {
    if (_modalRow) openFormModal_('edit', _modalRow, true);
  });
  document.getElementById('ramModalBtnSave')?.addEventListener('click', function() { saveModal_(); });
  document.getElementById('ramModalOverlay')?.addEventListener('click', function(e) {
    if (e.target.id === 'ramModalOverlay') {
      if (_modalMode === 'view') closeModal_();
      else cancelFormModal_();
    }
  });

  document.addEventListener('keydown', function(e) {
    if (e.key === 'Escape') {
      const ov = document.getElementById('ramModalOverlay');
      if (ov && ov.classList.contains('is-open')) {
        if (_modalMode === 'view') closeModal_();
        else cancelFormModal_();
      }
    }
  });

  const body = document.getElementById('ramTableBody');
  if (body) {
    body.addEventListener('click', function(e) {
      const viewBtn = e.target.closest('.ram-btn-view');
      if (viewBtn) {
        e.stopPropagation();
        const row = rowBySheetRow_(viewBtn.getAttribute('data-ram-row'));
        if (row) openViewModal_(row);
        return;
      }
      const tr = e.target.closest('tr.ram-row-click');
      if (!tr || !body.contains(tr)) return;
      const row = rowBySheetRow_(tr.getAttribute('data-ram-row'));
      if (row) openViewModal_(row);
    });
  }
}

/** Wire panel once; call when dashboard boots. */
export function setupRiskAnalysisMitigationPanel_(deps) {
  _deps = deps;
  ensureRamModalMounted_();
  bindUiOnce_();
}

/** Load list when user opens the sidebar panel. */
export function ensureRiskAnalysisMitigationPanel_() {
  if (!_deps) return;
  loadRamData_(false);
}

export { loadRamData_ as reloadRiskAnalysisMitigation_ };
