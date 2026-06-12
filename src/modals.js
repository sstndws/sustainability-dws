// ── TRACEABILITY + FFB SCREENING LOGIC ────────────────────────
// Modal save = in-memory ONLY. Tidak ada localStorage, tidak ada API call.
// Data dikirim ke Google Sheets saat user klik Save as Draft / Submit di main form
// (handleFinalSave → apiCreateSubmission / apiUpdateSubmission).
window._tmlSelectedMill = '';
window._tmlScreeningData = {};
window._tmlYNState = {};
window._tmlScreeningType = ''; // 'traceability' | 'ffb'
window._ffbSelectedSupplier = '';
window._ffbScreeningData = {};
window._ffbYNState = {};

// ── VIEW SCREENING POPUP (submitted-only, read-only, single scroll) ────────────
window.openViewScreeningPopup = async function() {
  var overlayId = 'sdd-view-screening-overlay';
  var existing = document.getElementById(overlayId);
  if (!existing) {
    var ov = document.createElement('div');
    ov.id = overlayId;
    ov.style.cssText = 'position:fixed;inset:0;background:rgba(15,23,42,0.52);z-index:100010;display:none;align-items:flex-start;justify-content:center;padding:24px 12px;overflow-y:auto;';
    ov.innerHTML =
      '<div style="width:min(860px,96vw);background:#fff;border-radius:16px;box-shadow:0 20px 60px rgba(0,0,0,0.3);margin:auto;overflow:hidden;">'
        + '<div style="display:flex;align-items:center;justify-content:space-between;padding:16px 22px;border-bottom:1.5px solid rgba(44,40,40,0.1);background:rgba(30,64,175,0.04);">'
          + '<span style="font-size:13px;font-weight:700;letter-spacing:1.2px;text-transform:uppercase;color:#1e40af;">Screening Results — View Only</span>'
          + '<button id="sdd-view-screening-close" style="background:none;border:none;cursor:pointer;padding:4px;color:#6b7280;font-size:20px;line-height:1;">&#x2715;</button>'
        + '</div>'
        + '<div id="sdd-view-screening-body" style="padding:22px;max-height:78vh;overflow-y:auto;"></div>'
      + '</div>';
    document.body.appendChild(ov);
    ov.addEventListener('click', function(e) { if (e.target === ov) ov.style.display = 'none'; });
    ov.querySelector('#sdd-view-screening-close').addEventListener('click', function() { ov.style.display = 'none'; });
    existing = ov;
  }

  var body = document.getElementById('sdd-view-screening-body');
  if (!body) return;
  body.innerHTML = '<div style="padding:18px;text-align:center;color:#6b7280;font-size:13px;">Loading screening data…</div>';
  existing.style.display = 'flex';

  var sid = window._sddSubmissionId || window._scrLoadedKey || '';
  var cachedGroup = sid && window._scrSavedGroupsByKey && window._scrSavedGroupsByKey[sid];
  var mills = (cachedGroup && Array.isArray(cachedGroup.mills)) ? cachedGroup.mills : [];
  var ffbRows = (cachedGroup && Array.isArray(cachedGroup.ffb_rows)) ? cachedGroup.ffb_rows : [];

  if (!mills.length && !ffbRows.length && sid) {
    try {
      var fetchFn = typeof window.apiGetSubmissionById === 'function' ? window.apiGetSubmissionById : null;
      if (!fetchFn) throw new Error('apiGetSubmissionById tidak tersedia');
      var res = await fetchFn(sid);
      mills   = (res && res.mills) ? res.mills : [];
      ffbRows = (res && res.ffb_rows) ? res.ffb_rows : [];
    } catch (e) {
      body.innerHTML = '<div style="padding:18px;color:#991b1b;font-size:13px;">Failed to load data: ' + (e.message || e) + '</div>';
      return;
    }
  }

  function esc(v) {
    return String(v == null ? '' : v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  }
  function row(label, val) {
    return '<div style="display:flex;gap:0;padding:6px 0;border-bottom:1px solid rgba(74,28,28,0.05);font-size:13px;line-height:1.5;">'
      + '<span style="color:#5F4A48;font-weight:600;width:240px;flex-shrink:0;">' + esc(label) + '</span>'
      + '<span style="color:#9C8080;width:16px;flex-shrink:0;">:</span>'
      + '<span style="color:#1A0A0A;">' + (esc(val) || '<span style="color:#B09A9A;">—</span>') + '</span>'
      + '</div>';
  }
  function section(title, inner) {
    return '<div style="margin-bottom:18px;border:1px solid rgba(44,40,40,0.09);border-radius:10px;overflow:hidden;">'
      + '<div style="padding:10px 16px;background:rgba(44,40,40,0.03);border-bottom:1.5px solid rgba(44,40,40,0.09);font-size:11px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;color:#2C2828;">' + esc(title) + '</div>'
      + '<div style="padding:14px 18px;">' + inner + '</div>'
      + '</div>';
  }

  var html = '';

  if (mills.length) {
    mills.forEach(function(m, idx) {
      var millName = m['TML - Mill Name'] || ('Mill #' + (idx + 1));
      var inner = '';
      inner += row('Mill Name', m['TML - Mill Name']);
      inner += row('Company Name', m['TML - Company Name']);
      inner += row('Valid Coordinate', m['SCR - TML Valid Coordinate']);
      inner += row('Forest Area', m['SCR - TML Forest Area']);
      inner += row('Peatland', m['SCR - TML Peatland']);
      inner += row('Moratorium', m['SCR - TML Moratorium']);
      inner += row('Moratorium (Ha)', m['SCR - TML Moratorium (Ha)']);
      inner += row('Deforestation Buffer 50KM (Ha)', m['SCR - TML Deforestation Buffer 50KM (Ha)']);
      inner += row('Screening Status', m['SCR - TML Screening Status']);
      inner += row('Screening Date', m['SCR - TML Screening Date']);
      html += section('Mill Screening — ' + millName, inner);
    });
  } else {
    html += '<div style="color:#9C8080;font-size:13px;margin-bottom:14px;padding:12px 16px;background:#f9fafb;border-radius:8px;">No Mill List data for this submission.</div>';
  }

  if (ffbRows.length) {
    ffbRows.forEach(function(f, idx) {
      var supplierName = f['FFB - Supplier Name'] || ('Supplier #' + (idx + 1));
      var inner = '';
      inner += row('Mill Name', f['FFB - Mill Name']);
      inner += row('Supplier Name', f['FFB - Supplier Name']);
      inner += row('Category', f['FFB - Supplier Category']);
      inner += row('Village', f['FFB - Village']);
      inner += row('Sub District', f['FFB - Sub District']);
      inner += row('District', f['FFB - District']);
      inner += row('Valid Coordinate', f['FFB - Valid Coordinate']);
      inner += row('Forest Area', f['FFB - Forest Area']);
      inner += row('Peatland', f['FFB - Peatland']);
      inner += row('Moratorium', f['FFB - Moratorium']);
      inner += row('Moratorium (Ha)', f['FFB - Moratorium (Ha)']);
      inner += row('Distance to Mill (Km)', f['FFB - Distance to Mill (Km)']);
      inner += row('Deforestation (Ha)', f['FFB - Deforestation (Ha)']);
      inner += row('Burn Area (Ha)', f['FFB - Burn Area (Ha)']);
      inner += row('Screening Status', f['FFB - Screening Status']);
      inner += row('Screening Date', f['FFB - Screening Date']);
      html += section('FFB Screening — ' + supplierName, inner);
    });
  } else {
    html += '<div style="color:#9C8080;font-size:13px;padding:12px 16px;background:#f9fafb;border-radius:8px;">No FFB Supplier data for this submission.</div>';
  }

  body.innerHTML = html || '<div style="color:#9C8080;font-size:13px;padding:12px;">No screening data.</div>';
};

// ── TYPE SELECTOR ──────────────────────────────────────────────
window.openTmlScreeningPicker = function() {
  // Reset type selection
  window._tmlScreeningType = '';
  window._tmlSelectedMill = '';
  window._ffbSelectedSupplier = '';
  ['trace','ffb'].forEach(t => {
    const card = document.getElementById('type-card-' + t);
    const check = document.getElementById('type-check-' + t);
    if (card) { card.style.background = 'white'; card.style.borderColor = 'rgba(74,28,28,0.15)'; }
    if (check) check.style.display = 'none';
  });
  document.getElementById('tml-type-overlay').style.display = 'flex';
};

window.tmlSelectTypeCard = function(t) {
  ['trace','ffb'].forEach(k => {
    const card = document.getElementById('type-card-' + k);
    const check = document.getElementById('type-check-' + k);
    if (k === t) {
      if (card) { card.style.background = 'rgba(139,26,26,0.04)'; card.style.borderColor = '#8B1A1A'; }
      if (check) check.style.display = 'flex';
    } else {
      if (card) { card.style.background = 'white'; card.style.borderColor = 'rgba(74,28,28,0.15)'; }
      if (check) check.style.display = 'none';
    }
  });
};

window.tmlTypeNext = function() {
  if (!window._tmlScreeningType) { alert('Select screening type first.'); return; }
  document.getElementById('tml-type-overlay').style.display = 'none';
  if (window._tmlScreeningType === 'traceability') {
    window._openTmlMillPicker();
  } else {
    window._openFfbSupplierPicker();
  }
};

/** After save/delete — return to Traceability vs FFB type picker (not main page). */
window.returnToScreeningTypePicker_ = function() {
  ['tml-form-overlay', 'tml-pick-overlay', 'ffb-form-overlay', 'ffb-pick-overlay', 'tml-result-overlay', 'ffb-result-overlay'].forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  var typeOverlay = document.getElementById('tml-type-overlay');
  if (typeOverlay) typeOverlay.style.display = 'flex';
  if (window._tmlScreeningType === 'traceability') window.tmlSelectTypeCard('trace');
  else if (window._tmlScreeningType === 'ffb') window.tmlSelectTypeCard('ffb');
};

// ── TRACEABILITY: MILL PICKER ──────────────────────────────────
window._openTmlMillPicker = function() {
  const mills = window._tmlMillNames || [];
  const list = document.getElementById('tml-mill-list');
  if (!mills.length) {
    list.innerHTML = '<div style="padding:16px;text-align:center;color:#9C8080;font-size:13px;">No mill data. Import an Excel file with traceability data first.</div>';
  } else {
    list.innerHTML = mills.map((m, i) => {
      const statusNote = window._tmlScreeningData[m] ? ' <span style="color:#8B1A1A;font-size:12px;font-weight:600;">(' + window._tmlScreeningData[m].status + ')</span>' : '';
      return `<div onclick="window._tmlSelectedMill='${m.replace(/'/g,"\\'")}';document.querySelectorAll('.tml-mill-item').forEach(el=>el.style.background='white');this.style.background='rgba(139,26,26,0.07)';" class="tml-mill-item" style="padding:11px 16px;cursor:pointer;border-bottom:1px solid rgba(74,28,28,0.07);font-size:13px;font-family:Inter,sans-serif;color:#1A0A0A;transition:background 0.15s;">`
        + '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#8B1A1A" stroke-width="2.5" style="margin-right:8px;vertical-align:middle;"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>'
        + m + statusNote + '</div>';
    }).join('');
  }
  window._tmlSelectedMill = '';
  document.getElementById('tml-pick-overlay').style.display = 'flex';
};

window.openTmlScreeningForm = function() {
  if (!window._tmlSelectedMill) { alert('Select a Mill first.'); return; }
  document.getElementById('tml-pick-overlay').style.display = 'none';
  document.getElementById('tml-form-mill-label').textContent = 'Mill: ' + window._tmlSelectedMill;
  document.getElementById('tml-form-status').textContent = '';

  // Reset form
  ['coord','mora'].forEach(k => {
    window._tmlYNState[k] = '';
    document.querySelectorAll('[name="tml-' + k + '"]').forEach(r => {
      r.closest('label').style.background = 'white';
      r.closest('label').style.borderColor = 'rgba(74,28,28,0.15)';
    });
  });
  ['fa-apl','fa-hpk','fa-hp','fa-hl','fa-ksa','pl-no','pl-prot','pl-cult'].forEach(id => {
    const cb = document.getElementById(id);
    if (cb) { cb.checked = false; window.tmlToggleHa(cb, id + '-ha'); }
  });
  document.getElementById('mora-ha').value = '';
  document.getElementById('mora-ha-wrap').style.display = 'none';
  document.getElementById('defbuf-ha').value = '';
  syncForestAreaGroup_(TML_FOREST_CFG, window.tmlToggleHa);
  syncForestAreaGroup_(TML_PEAT_CFG, window.tmlToggleHa);

  window._populateTmlScreeningForm(window._tmlSelectedMill);
  document.getElementById('tml-form-overlay').style.display = 'flex';
};

window._populateTmlScreeningForm = function(mill) {
  const saved = window._tmlScreeningData && window._tmlScreeningData[mill] ? window._tmlScreeningData[mill] : null;
  const setRadio = function(name, value) {
    document.querySelectorAll('[name="' + name + '"]').forEach(r => {
      const lbl = r.closest('label');
      if (r.value === value) {
        r.checked = true;
        lbl.style.background = 'rgba(139,26,26,0.08)';
        lbl.style.borderColor = '#8B1A1A';
      } else {
        r.checked = false;
        lbl.style.background = 'white';
        lbl.style.borderColor = 'rgba(74,28,28,0.15)';
      }
    });
  };

  const forestItems = [
    {id:'fa-apl', label:'APL'}, {id:'fa-hpk', label:'HPK'},
    {id:'fa-hp', label:'HP'}, {id:'fa-hl', label:'HL'}, {id:'fa-ksa', label:'KSA/KPA'}
  ];
  const peatItems = [
    {id:'pl-no', label:'No'}, {id:'pl-prot', label:'Protected Peat'}, {id:'pl-cult', label:'Cultivated Peat'}
  ];

  if (!saved) {
    setRadio('tml-coord', '');
    setRadio('tml-mora', '');
    forestItems.forEach(item => {
      const cb = document.getElementById(item.id);
      if (cb) { cb.checked = false; window.tmlToggleHa(cb, item.id + '-ha'); }
    });
    peatItems.forEach(item => {
      const cb = document.getElementById(item.id);
      if (cb) { cb.checked = false; window.tmlToggleHa(cb, item.id + '-ha'); }
    });
    document.getElementById('mora-ha-wrap').style.display = 'none';
    document.getElementById('mora-ha').value = '';
    document.getElementById('defbuf-ha').value = '';
    document.getElementById('tml-form-status').textContent = '';
    window._tmlYNState.coord = '';
    window._tmlYNState.mora = '';
    syncForestAreaGroup_(TML_FOREST_CFG, window.tmlToggleHa);
    syncForestAreaGroup_(TML_PEAT_CFG, window.tmlToggleHa);
    return;
  }

  window._tmlYNState.coord = saved.coord || '';
  window._tmlYNState.mora = saved.mora || '';
  setRadio('tml-coord', saved.coord === 'Yes' ? 'Yes' : saved.coord === 'No' ? 'No' : '');
  setRadio('tml-mora', saved.mora === 'Yes' ? 'Yes' : saved.mora === 'No' ? 'No' : '');
  if (saved.mora === 'Yes') {
    document.getElementById('mora-ha-wrap').style.display = 'block';
    document.getElementById('mora-ha').value = saved.moraHa === '—' ? '' : saved.moraHa;
  } else {
    document.getElementById('mora-ha-wrap').style.display = 'none';
    document.getElementById('mora-ha').value = '';
  }
  document.getElementById('defbuf-ha').value = saved.defbuf === '—' ? '' : saved.defbuf;

  forestItems.forEach(item => {
    const cb = document.getElementById(item.id);
    const match = saved.forestItems?.find(fi => fi.label === item.label);
    if (cb) {
      cb.checked = !!match;
      window.tmlToggleHa(cb, item.id + '-ha');
      document.getElementById(item.id + '-ha').value = match && match.ha !== '—' ? match.ha : '';
    }
  });
  peatItems.forEach(item => {
    const cb = document.getElementById(item.id);
    const match = saved.peatItems?.find(fi => fi.label === item.label);
    if (cb) {
      cb.checked = !!match;
      window.tmlToggleHa(cb, item.id + '-ha');
      document.getElementById(item.id + '-ha').value = match && match.ha !== '—' ? match.ha : '';
    }
  });
  document.getElementById('tml-form-status').textContent = saved.status ? 'Status: ' + saved.status : '';
  syncForestAreaGroup_(TML_FOREST_CFG, window.tmlToggleHa);
  syncForestAreaGroup_(TML_PEAT_CFG, window.tmlToggleHa);
};

window.tmlSetYN = function(key, val, labelEl) {
  window._tmlYNState[key] = val;
  document.querySelectorAll('[name="tml-' + key + '"]').forEach(r => {
    const lbl = r.closest('label');
    if (r.value === val) {
      lbl.style.background = 'rgba(139,26,26,0.08)';
      lbl.style.borderColor = '#8B1A1A';
    } else {
      lbl.style.background = 'white';
      lbl.style.borderColor = 'rgba(74,28,28,0.15)';
    }
  });
};

window.tmlToggleHa = function(cb, haId) {
  const inp = document.getElementById(haId);
  if (!inp) return;
  if (cb.checked) {
    inp.disabled = false;
    inp.style.background = 'white';
    inp.style.color = '#1A0A0A';
    inp.focus();
    cb.closest('label').style.borderColor = '#8B1A1A';
    cb.closest('label').style.background = 'rgba(139,26,26,0.04)';
  } else {
    inp.disabled = true;
    inp.value = '';
    inp.style.background = '#f8f6f5';
    inp.style.color = '#9C8080';
    cb.closest('label').style.borderColor = 'rgba(74,28,28,0.12)';
    cb.closest('label').style.background = 'white';
  }
};

const TML_FOREST_CFG = { apl: 'fa-apl', others: ['fa-hpk', 'fa-hp', 'fa-hl', 'fa-ksa'] };
const FFB_FOREST_CFG = { apl: 'ffb-fa-apl', others: ['ffb-fa-hpk', 'ffb-fa-hp', 'ffb-fa-hl', 'ffb-fa-ksa'] };
const TML_PEAT_CFG = { apl: 'pl-no', others: ['pl-prot', 'pl-cult'] };
const FFB_PEAT_CFG = { apl: 'ffb-pl-no', others: ['ffb-pl-prot', 'ffb-pl-cult'] };

function setForestRowVisible_(cbId, visible, toggleHaFn) {
  const cb = document.getElementById(cbId);
  if (!cb) return;
  const row = cb.closest('label');
  if (row) row.style.display = visible ? '' : 'none';
  if (!visible && cb.checked) {
    cb.checked = false;
    toggleHaFn(cb, cbId + '-ha');
  }
}

function syncForestAreaGroup_(cfg, toggleHaFn) {
  const aplCb = document.getElementById(cfg.apl);
  if (!aplCb) return;
  const aplChecked = aplCb.checked;
  const anyOtherChecked = cfg.others.some(function(id) {
    const c = document.getElementById(id);
    return c && c.checked;
  });

  if (aplChecked) {
    setForestRowVisible_(cfg.apl, true, toggleHaFn);
    cfg.others.forEach(function(id) { setForestRowVisible_(id, false, toggleHaFn); });
  } else if (anyOtherChecked) {
    setForestRowVisible_(cfg.apl, false, toggleHaFn);
    cfg.others.forEach(function(id) { setForestRowVisible_(id, true, toggleHaFn); });
  } else {
    setForestRowVisible_(cfg.apl, true, toggleHaFn);
    cfg.others.forEach(function(id) { setForestRowVisible_(id, true, toggleHaFn); });
  }
}

function onForestAreaChange_(cb, cfg, toggleHaFn) {
  toggleHaFn(cb, cb.id + '-ha');
  if (cb.id === cfg.apl && cb.checked) {
    cfg.others.forEach(function(id) {
      const other = document.getElementById(id);
      if (other && other.checked) {
        other.checked = false;
        toggleHaFn(other, id + '-ha');
      }
    });
  } else if (cb.id !== cfg.apl && cb.checked) {
    const apl = document.getElementById(cfg.apl);
    if (apl && apl.checked) {
      apl.checked = false;
      toggleHaFn(apl, cfg.apl + '-ha');
    }
  }
  syncForestAreaGroup_(cfg, toggleHaFn);
}

window.tmlOnForestAreaChange = function(cb) {
  onForestAreaChange_(cb, TML_FOREST_CFG, window.tmlToggleHa);
};

window.ffbOnForestAreaChange = function(cb) {
  onForestAreaChange_(cb, FFB_FOREST_CFG, window.ffbToggleHa);
};

window.tmlOnPeatlandChange = function(cb) {
  onForestAreaChange_(cb, TML_PEAT_CFG, window.tmlToggleHa);
};

window.ffbOnPeatlandChange = function(cb) {
  onForestAreaChange_(cb, FFB_PEAT_CFG, window.ffbToggleHa);
};

/**
 * saveTmlScreening — SYNC. Tulis ke window._tmlScreeningData (in-memory only).
 * Tidak ada API call, tidak ada localStorage. API call terjadi saat user klik
 * Save as Draft / Submit di main form (handleFinalSave → apiCreateSubmission/apiUpdateSubmission).
 */
window.saveTmlScreening = function() {
  const mill = window._tmlSelectedMill;
  if (!mill) {
    alert('Select a Mill first.');
    return;
  }

  const coord = window._tmlYNState['coord'] || '—';
  const mora = window._tmlYNState['mora'] || '—';
  const moraHa = mora === 'Yes' ? (document.getElementById('mora-ha').value || '—') : '—';
  const defbuf = document.getElementById('defbuf-ha').value || '—';

  const forestItems = [
    {id:'fa-apl', label:'APL'}, {id:'fa-hpk', label:'HPK'},
    {id:'fa-hp', label:'HP'}, {id:'fa-hl', label:'HL'}, {id:'fa-ksa', label:'KSA/KPA'}
  ].filter(x => document.getElementById(x.id) && document.getElementById(x.id).checked)
   .map(x => ({ label: x.label, ha: document.getElementById(x.id + '-ha').value || '—' }));

  const peatItems = [
    {id:'pl-no', label:'No'}, {id:'pl-prot', label:'Protected Peat'}, {id:'pl-cult', label:'Cultivated Peat'}
  ].filter(x => document.getElementById(x.id) && document.getElementById(x.id).checked)
   .map(x => ({ label: x.label, ha: document.getElementById(x.id + '-ha').value || '—' }));

  if (!window._tmlScreeningData) window._tmlScreeningData = {};
  window._tmlScreeningData[mill] = {
    coord,
    forestItems,
    peatItems,
    mora,
    moraHa,
    defbuf,
    status: 'Draft',
    date: new Date().toLocaleDateString('id-ID')
  };

  window.returnToScreeningTypePicker_();
  if (typeof window.showSddToast === 'function') {
    window.showSddToast(
      'Saved temporarily. Click Save as Draft on the main form to send to Sheets.',
      'success'
    );
  }
};

/**
 * deleteTmlScreening — SYNC. Hapus mill dari window._tmlScreeningData saja.
 * Tidak ada API call. Penghapusan baru ter-propagate ke Sheets saat handleFinalSave berikutnya.
 */
window.deleteTmlScreening = function() {
  const mill = window._tmlSelectedMill;
  if (!mill) { alert('Select a Mill first.'); return; }
  if (!confirm('Are you sure you want to delete screening for "' + mill + '"?')) return;
  if (window._tmlScreeningData) delete window._tmlScreeningData[mill];
  window.returnToScreeningTypePicker_();
  if (typeof window.showSddToast === 'function') {
    window.showSddToast('TML screening for "' + mill + '" removed from memory.', 'info');
  }
};

// ── FFB SUPPLIER LIST SCREENING ────────────────────────────────
window._openFfbSupplierPicker = function() {
  const suppliers = window._ffbSupplierNames || [];
  const list = document.getElementById('ffb-supplier-list');
  if (!suppliers.length) {
    list.innerHTML = '<div style="padding:16px;text-align:center;color:#9C8080;font-size:13px;">No supplier data. Import an Excel file with FFB Supplier List data first.</div>';
  } else {
    list.innerHTML = suppliers.map(s => {
      const statusNote = window._ffbScreeningData[s] ? ' <span style="color:#8B1A1A;font-size:12px;font-weight:600;">(' + window._ffbScreeningData[s].status + ')</span>' : '';
      return `<div onclick="window._ffbSelectedSupplier='${s.replace(/'/g,"\\'")}';document.querySelectorAll('.ffb-supplier-item').forEach(el=>el.style.background='white');this.style.background='rgba(139,26,26,0.07)';" class="ffb-supplier-item" style="padding:11px 16px;cursor:pointer;border-bottom:1px solid rgba(74,28,28,0.07);font-size:13px;font-family:Inter,sans-serif;color:#1A0A0A;transition:background 0.15s;">`
        + '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#8B1A1A" stroke-width="2.5" style="margin-right:8px;vertical-align:middle;"><path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>'
        + s + statusNote + '</div>';
    }).join('');
  }
  window._ffbSelectedSupplier = '';
  document.getElementById('ffb-pick-overlay').style.display = 'flex';
};

window.openFfbScreeningForm = function() {
  if (!window._ffbSelectedSupplier) { alert('Select a Supplier first.'); return; }
  document.getElementById('ffb-pick-overlay').style.display = 'none';
  document.getElementById('ffb-form-supplier-label').textContent = 'Supplier: ' + window._ffbSelectedSupplier;
  document.getElementById('ffb-form-status').textContent = '';

  // Reset form
  window._ffbYNState = {};
  document.querySelectorAll('[name="ffb-mora"],[name="ffb-coord"]').forEach(r => {
    r.closest('label').style.background = 'white';
    r.closest('label').style.borderColor = 'rgba(74,28,28,0.15)';
  });
  ['ffb-fa-apl','ffb-fa-hpk','ffb-fa-hp','ffb-fa-hl','ffb-fa-ksa','ffb-pl-no','ffb-pl-prot','ffb-pl-cult'].forEach(id => {
    const cb = document.getElementById(id);
    if (cb) { cb.checked = false; window.ffbToggleHa(cb, id + '-ha'); }
  });
  document.getElementById('ffb-mora-ha').value = '';
  document.getElementById('ffb-mora-ha-wrap').style.display = 'none';
  document.getElementById('ffb-dist-km').value = '';
  document.getElementById('ffb-defor').value = '';
  document.getElementById('ffb-burn').value = '';
  syncForestAreaGroup_(FFB_FOREST_CFG, window.ffbToggleHa);
  syncForestAreaGroup_(FFB_PEAT_CFG, window.ffbToggleHa);

  window._populateFfbScreeningForm(window._ffbSelectedSupplier);
  document.getElementById('ffb-form-overlay').style.display = 'flex';
};

window._populateFfbScreeningForm = function(supplier) {
  const saved = window._ffbScreeningData && window._ffbScreeningData[supplier] ? window._ffbScreeningData[supplier] : null;
  const setRadio = function(name, value) {
    document.querySelectorAll('[name="' + name + '"]').forEach(r => {
      const lbl = r.closest('label');
      if (r.value === value) {
        r.checked = true;
        lbl.style.background = 'rgba(139,26,26,0.08)';
        lbl.style.borderColor = '#8B1A1A';
      } else {
        r.checked = false;
        lbl.style.background = 'white';
        lbl.style.borderColor = 'rgba(74,28,28,0.15)';
      }
    });
  };

  const forestItems = [
    {id:'ffb-fa-apl', label:'APL'}, {id:'ffb-fa-hpk', label:'HPK'},
    {id:'ffb-fa-hp', label:'HP'}, {id:'ffb-fa-hl', label:'HL'}, {id:'ffb-fa-ksa', label:'KSA/KPA'}
  ];
  const peatItems = [
    {id:'ffb-pl-no', label:'No'}, {id:'ffb-pl-prot', label:'Protected Peat'}, {id:'ffb-pl-cult', label:'Cultivated Peat'}
  ];

  if (!saved) {
    setRadio('ffb-coord', '');
    setRadio('ffb-mora', '');
    forestItems.forEach(item => {
      const cb = document.getElementById(item.id);
      if (cb) { cb.checked = false; window.ffbToggleHa(cb, item.id + '-ha'); }
    });
    peatItems.forEach(item => {
      const cb = document.getElementById(item.id);
      if (cb) { cb.checked = false; window.ffbToggleHa(cb, item.id + '-ha'); }
    });
    document.getElementById('ffb-mora-ha-wrap').style.display = 'none';
    document.getElementById('ffb-mora-ha').value = '';
    document.getElementById('ffb-dist-km').value = '';
    document.getElementById('ffb-defor').value = '';
    document.getElementById('ffb-burn').value = '';
    document.getElementById('ffb-form-status').textContent = '';
    window._ffbYNState.coord = '';
    window._ffbYNState.mora = '';
    syncForestAreaGroup_(FFB_FOREST_CFG, window.ffbToggleHa);
    syncForestAreaGroup_(FFB_PEAT_CFG, window.ffbToggleHa);
    return;
  }

  window._ffbYNState.coord = saved.coord || '';
  window._ffbYNState.mora = saved.mora || '';
  setRadio('ffb-coord', saved.coord === 'Yes' ? 'Yes' : saved.coord === 'No' ? 'No' : '');
  setRadio('ffb-mora', saved.mora === 'Yes' ? 'Yes' : saved.mora === 'No' ? 'No' : '');
  if (saved.mora === 'Yes') {
    document.getElementById('ffb-mora-ha-wrap').style.display = 'block';
    document.getElementById('ffb-mora-ha').value = saved.moraHa === '—' ? '' : saved.moraHa;
  } else {
    document.getElementById('ffb-mora-ha-wrap').style.display = 'none';
    document.getElementById('ffb-mora-ha').value = '';
  }
  document.getElementById('ffb-dist-km').value = saved.distKm === '—' ? '' : saved.distKm;
  document.getElementById('ffb-defor').value = saved.defor === '—' ? '' : saved.defor;
  document.getElementById('ffb-burn').value = saved.burn === '—' ? '' : saved.burn;

  forestItems.forEach(item => {
    const cb = document.getElementById(item.id);
    const match = saved.forestItems?.find(fi => fi.label === item.label);
    if (cb) {
      cb.checked = !!match;
      window.ffbToggleHa(cb, item.id + '-ha');
      document.getElementById(item.id + '-ha').value = match && match.ha !== '—' ? match.ha : '';
    }
  });
  peatItems.forEach(item => {
    const cb = document.getElementById(item.id);
    const match = saved.peatItems?.find(fi => fi.label === item.label);
    if (cb) {
      cb.checked = !!match;
      window.ffbToggleHa(cb, item.id + '-ha');
      document.getElementById(item.id + '-ha').value = match && match.ha !== '—' ? match.ha : '';
    }
  });
  document.getElementById('ffb-form-status').textContent = saved.status ? 'Status: ' + saved.status : '';
  syncForestAreaGroup_(FFB_FOREST_CFG, window.ffbToggleHa);
  syncForestAreaGroup_(FFB_PEAT_CFG, window.ffbToggleHa);
};

window.ffbSetYN = function(key, val, labelEl) {
  window._ffbYNState[key] = val;
  document.querySelectorAll('[name="ffb-' + key + '"]').forEach(r => {
    const lbl = r.closest('label');
    if (r.value === val) {
      lbl.style.background = 'rgba(139,26,26,0.08)';
      lbl.style.borderColor = '#8B1A1A';
    } else {
      lbl.style.background = 'white';
      lbl.style.borderColor = 'rgba(74,28,28,0.15)';
    }
  });
  if (key === 'mora') {
    document.getElementById('ffb-mora-ha-wrap').style.display = val === 'Yes' ? 'block' : 'none';
    if (val === 'No') document.getElementById('ffb-mora-ha').value = '';
  }
};

window.ffbToggleHa = function(cb, haId) {
  const inp = document.getElementById(haId);
  if (!inp) return;
  if (cb.checked) {
    inp.disabled = false;
    inp.style.background = 'white';
    inp.style.color = '#1A0A0A';
    inp.focus();
    cb.closest('label').style.borderColor = '#8B1A1A';
    cb.closest('label').style.background = 'rgba(139,26,26,0.04)';
  } else {
    inp.disabled = true;
    inp.value = '';
    inp.style.background = '#f8f6f5';
    inp.style.color = '#9C8080';
    cb.closest('label').style.borderColor = 'rgba(74,28,28,0.12)';
    cb.closest('label').style.background = 'white';
  }
};

/**
 * saveFfbScreening — SYNC. Tulis ke window._ffbScreeningData (in-memory only).
 * Tidak ada API call, tidak ada localStorage. API call terjadi saat user klik
 * Save as Draft / Submit di main form (handleFinalSave → apiCreateSubmission/apiUpdateSubmission).
 */
window.saveFfbScreening = function() {
  const supplier = window._ffbSelectedSupplier;
  if (!supplier) {
    alert('Select a supplier first.');
    return;
  }

  const coord = window._ffbYNState['coord'] || '—';
  const mora = window._ffbYNState['mora'] || '—';
  const moraHa = mora === 'Yes' ? (document.getElementById('ffb-mora-ha').value || '—') : '—';
  const distKm = document.getElementById('ffb-dist-km').value || '—';
  const defor = document.getElementById('ffb-defor').value || '—';
  const burn = document.getElementById('ffb-burn').value || '—';

  const forestItems = [
    {id:'ffb-fa-apl', label:'APL'}, {id:'ffb-fa-hpk', label:'HPK'},
    {id:'ffb-fa-hp', label:'HP'}, {id:'ffb-fa-hl', label:'HL'}, {id:'ffb-fa-ksa', label:'KSA/KPA'}
  ].filter(x => document.getElementById(x.id) && document.getElementById(x.id).checked)
   .map(x => ({ label: x.label, ha: document.getElementById(x.id + '-ha').value || '—' }));

  const peatItems = [
    {id:'ffb-pl-no', label:'No'}, {id:'ffb-pl-prot', label:'Protected Peat'}, {id:'ffb-pl-cult', label:'Cultivated Peat'}
  ].filter(x => document.getElementById(x.id) && document.getElementById(x.id).checked)
   .map(x => ({ label: x.label, ha: document.getElementById(x.id + '-ha').value || '—' }));

  if (!window._ffbScreeningData) window._ffbScreeningData = {};
  window._ffbScreeningData[supplier] = {
    coord,
    forestItems,
    peatItems,
    mora,
    moraHa,
    distKm,
    defor,
    burn,
    status: 'Draft',
    date: new Date().toLocaleDateString('id-ID')
  };

  window.returnToScreeningTypePicker_();
  if (typeof window.showSddToast === 'function') {
    window.showSddToast(
      'Saved temporarily. Click Save as Draft on the main form to send to Sheets.',
      'success'
    );
  }
};

/**
 * deleteFfbScreening — SYNC. Hapus supplier dari window._ffbScreeningData saja.
 * Tidak ada API call.
 */
window.deleteFfbScreening = function() {
  const supplier = window._ffbSelectedSupplier;
  if (!supplier) { alert('Select a supplier first.'); return; }
  if (!confirm('Are you sure you want to delete FFB screening for "' + supplier + '"?')) return;
  if (window._ffbScreeningData) delete window._ffbScreeningData[supplier];
  window.returnToScreeningTypePicker_();
  if (typeof window.showSddToast === 'function') {
    window.showSddToast('FFB screening for "' + supplier + '" removed from memory.', 'info');
  }
};
