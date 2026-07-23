/**
 * EUDR Due Diligence Statement (DDS) — list, form, save, PDF export.
 */

import { initDashDateFields, dashDateCollectValues, dashIsoToDisplay } from './dash-date-field.js';
import {
  DDS_LIST_FIELDS,
  DDS_FIELD_SECTIONS,
  DDS_DATE_FIELDS,
  DDS_DOC_CODES,
  DDS_SUPPLIER_TYPES,
  DDS_HS_CODES,
  DDS_HS_CUSTOM_STORAGE_KEY,
  DDS_COUNTRY_RISK,
  DDS_RA_RISK,
  DDS_FIELD_LABELS,
} from './dds-constants.js';

export {
  DDS_LIST_FIELDS,
  DDS_FIELD_SECTIONS,
  DDS_DATE_FIELDS,
  DDS_DOC_CODES,
  DDS_SUPPLIER_TYPES,
  DDS_HS_CODES,
  DDS_HS_CUSTOM_STORAGE_KEY,
  DDS_COUNTRY_RISK,
  DDS_RA_RISK,
} from './dds-constants.js';

function ddsLabel_(field) {
  return DDS_FIELD_LABELS[field] || field;
}

function ddsNormSd_(val) {
  return String(val || '').trim().toLowerCase();
}

function ddsPick_(row, field) {
  if (!row) return '';
  if (row[field] != null && String(row[field]).trim() !== '') return row[field];
  const want = String(field || '').trim().toUpperCase();
  return Object.keys(row).reduce(function(found, k) {
    if (found !== '') return found;
    if (k.charAt(0) === '_') return found;
    if (String(k).trim().toUpperCase() === want) return row[k];
    return found;
  }, '');
}

function ddsDefaultDocuments_() {
  return DDS_DOC_CODES.map(function(def) {
    return {
      'DOC CODE': def.code,
      'DOC NUMBER': '',
      'DOC DATE': '',
      AVAILABLE: '',
      NOTES: '',
    };
  });
}

export function initDdsPanel_(deps) {
  const {
    apiGet,
    apiPost,
    escHtml,
    dashDateFieldHtml,
    dashDateCollectValues,
    getJsPDF,
    showToast,
    debounce,
    mountOverlay,
  } = deps;

  let ddsMaster = [];
  let ddsSuppliers = [];
  let ddsGeo = [];
  let ddsDocuments = [];
  let ddsLoaded = false;
  let ddsLoadPromise = null;
  let ddsSearch = '';
  let ddsFormMode = 'add';
  let ddsFormMasterRow = null;
  let ddsSaving = false;

  function ddsToast_(msg, type) {
    if (typeof showToast === 'function') showToast(msg, type || 'info');
  }

  function ddsGetCustomHsCodes_() {
    try {
      const raw = localStorage.getItem(DDS_HS_CUSTOM_STORAGE_KEY);
      const parsed = raw ? JSON.parse(raw) : [];
      return Array.isArray(parsed) ? parsed.filter(function(entry) {
        return entry && String(entry.code || '').trim();
      }) : [];
    } catch (e) {
      return [];
    }
  }

  function ddsSaveCustomHsCode_(entry) {
    const code = String(entry.code || '').trim();
    if (!code) return false;
    const desc = String(entry.description || '').trim();
    const all = ddsGetAllHsCodes_();
    if (all.some(function(item) { return item.code === code; })) return true;
    const custom = ddsGetCustomHsCodes_();
    custom.push({ code: code, description: desc });
    localStorage.setItem(DDS_HS_CUSTOM_STORAGE_KEY, JSON.stringify(custom));
    return true;
  }

  function ddsGetAllHsCodes_() {
    return DDS_HS_CODES.concat(ddsGetCustomHsCodes_());
  }

  function ddsHsOptionLabel_(entry) {
    const desc = String(entry.description || '').trim();
    if (!desc) return entry.code;
    const short = desc.length > 70 ? desc.slice(0, 67) + '…' : desc;
    return entry.code + ' — ' + short;
  }

  function ddsHsFieldHtml_(value, opts) {
    opts = opts || {};
    const label = ddsLabel_('HS CODE');
    const saved = value != null ? String(value).trim() : '';
    const allCodes = ddsGetAllHsCodes_();
    const matched = allCodes.some(function(entry) { return entry.code === saved; });
    const req = opts.required ? ' required' : '';

    let optsHtml = '<option value="">— Select HS code —</option>';
    allCodes.forEach(function(entry) {
      const selected = saved === entry.code ? ' selected' : '';
      optsHtml += ''
        + '<option value="' + escHtml(entry.code) + '"' + selected
        + ' title="' + escHtml(entry.description || '') + '">'
        + escHtml(ddsHsOptionLabel_(entry)) + '</option>';
    });
    if (saved && !matched) {
      optsHtml += '<option value="' + escHtml(saved) + '" selected>' + escHtml(saved) + ' (saved)</option>';
    }
    optsHtml += '<option value="__custom__"' + (!saved && opts.openCustom ? ' selected' : '') + '>+ Add custom HS code…</option>';

    return ''
      + '<div class="form-field dds-hs-field full">'
      + '<label>' + escHtml(label) + '</label>'
      + '<select class="dds-hs-select"' + req + '>' + optsHtml + '</select>'
      + '<input type="hidden" data-dds-field="HS CODE" value="' + escHtml(saved) + '" />'
      + '<div class="dds-hs-custom"' + (opts.openCustom ? '' : ' style="display:none"') + '>'
      + '<div class="dds-hs-custom-row">'
      + '<input type="text" class="dds-hs-custom-code" placeholder="CN/HS heading (e.g. 1511.10)" value="' + escHtml(saved && !matched ? saved : '') + '" />'
      + '<input type="text" class="dds-hs-custom-desc" placeholder="Description (optional)" />'
      + '</div>'
      + '<div class="dds-hs-custom-actions">'
      + '<button type="button" class="btn-sm btn-outline dds-hs-save-btn">Save to list</button>'
      + '<button type="button" class="btn-sm btn-primary dds-hs-use-btn">Use this code</button>'
      + '</div>'
      + '</div>'
      + '<p class="dds-hs-hint">EUDR Annex I CN/HS headings for palm products. Use <strong>+ Add custom HS code</strong> for sub-codes (e.g. 1511.10).</p>'
      + '</div>';
  }

  function ddsInitHsField_(wrap) {
    if (!wrap) return;
    const select = wrap.querySelector('.dds-hs-select');
    const hidden = wrap.querySelector('[data-dds-field="HS CODE"]');
    const customPanel = wrap.querySelector('.dds-hs-custom');
    if (!select || !hidden) return;
    const saved = String(hidden.value || '').trim();
    if (saved && select.value !== saved && select.value !== '__custom__') {
      const hit = Array.from(select.options).some(function(opt) { return opt.value === saved; });
      if (hit) select.value = saved;
    }
    if (select.value === '__custom__' && customPanel) customPanel.style.display = '';
  }

  function ddsApplyHsCustomCode_(wrap, saveToList) {
    if (!wrap) return;
    const codeInput = wrap.querySelector('.dds-hs-custom-code');
    const descInput = wrap.querySelector('.dds-hs-custom-desc');
    const hidden = wrap.querySelector('[data-dds-field="HS CODE"]');
    const select = wrap.querySelector('.dds-hs-select');
    const code = codeInput ? String(codeInput.value || '').trim() : '';
    if (!code) {
      ddsToast_('Enter a CN/HS heading first.', 'warning');
      if (codeInput) codeInput.focus();
      return;
    }
    const desc = descInput ? String(descInput.value || '').trim() : '';
    if (saveToList) {
      ddsSaveCustomHsCode_({ code: code, description: desc });
      const mainGrid = document.getElementById('ddsFieldsMain');
      const fieldWrap = mainGrid && mainGrid.querySelector('.dds-hs-field');
      if (fieldWrap) {
        fieldWrap.outerHTML = ddsHsFieldHtml_(code, {});
        ddsInitHsField_(mainGrid.querySelector('.dds-hs-field'));
      }
      ddsToast_('HS code added to list.', 'success');
      return;
    }

    if (hidden) hidden.value = code;
    if (select) {
      let opt = Array.from(select.options).find(function(o) { return o.value === code; });
      if (!opt) {
        opt = document.createElement('option');
        opt.value = code;
        opt.textContent = code + (desc ? ' — ' + desc : '');
        select.insertBefore(opt, select.querySelector('option[value="__custom__"]'));
      }
      select.value = code;
    }
    const customPanel = wrap.querySelector('.dds-hs-custom');
    if (customPanel) customPanel.style.display = 'none';
  }

  function ddsBundleBySd_(sdNumber) {
    const key = ddsNormSd_(sdNumber);
    return {
      master: ddsMaster.find(function(r) { return ddsNormSd_(ddsPick_(r, 'SD NUMBER')) === key; }) || null,
      suppliers: ddsSuppliers.filter(function(r) { return ddsNormSd_(ddsPick_(r, 'SD NUMBER')) === key; }),
      geolocation: ddsGeo.filter(function(r) { return ddsNormSd_(ddsPick_(r, 'SD NUMBER')) === key; }),
      documents: ddsDocuments.filter(function(r) { return ddsNormSd_(ddsPick_(r, 'SD NUMBER')) === key; }),
    };
  }

  function ddsPrepareMasterRow_(row) {
    const out = Object.assign({}, row || {});
    out._sdKey = ddsNormSd_(ddsPick_(out, 'SD NUMBER'));
    out._searchBlob = DDS_LIST_FIELDS.map(function(f) { return String(ddsPick_(out, f) || ''); }).join(' ').toLowerCase();
    return out;
  }

  function ddsUpdateStats_() {
    const totalEl = document.getElementById('dds-stat-total');
    const buyersEl = document.getElementById('dds-stat-buyers');
    const plantsEl = document.getElementById('dds-stat-plants');
    if (!totalEl) return;
    const rows = ddsMaster || [];
    totalEl.textContent = String(rows.length);
    const buyers = {};
    const plants = {};
    rows.forEach(function(r) {
      const b = String(ddsPick_(r, 'BUYER NAME') || '').trim();
      const p = String(ddsPick_(r, 'PLANT') || '').trim();
      if (b) buyers[b.toLowerCase()] = true;
      if (p) plants[p.toLowerCase()] = true;
    });
    if (buyersEl) buyersEl.textContent = String(Object.keys(buyers).length);
    if (plantsEl) plantsEl.textContent = String(Object.keys(plants).length);
  }

  function ddsFilteredMaster_() {
    const q = ddsSearch;
    return (ddsMaster || []).filter(function(r) {
      return !q || (r._searchBlob || '').includes(q);
    });
  }

  function renderDdsTable_() {
    const body = document.getElementById('ddsTableBody');
    const table = document.getElementById('ddsTable');
    if (!body) return;
    const filtered = ddsFilteredMaster_();
    if (table) table.style.display = filtered.length ? 'table' : 'none';
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="12" style="text-align:center;padding:32px;color:#9C8A8A;">'
        + (ddsSearch ? 'No results match your search.' : 'No DDS records yet. Click "+ Add DDS" to create one.')
        + '</td></tr>';
      if (table) table.style.display = 'table';
      return;
    }
    body.innerHTML = filtered.map(function(d) {
      const sd = escHtml(String(ddsPick_(d, 'SD NUMBER') || '—'));
      const cells = DDS_LIST_FIELDS.map(function(f) {
        return '<td>' + escHtml(String(ddsPick_(d, f) || '—')) + '</td>';
      }).join('');
      return ''
        + '<tr class="dds-row-clickable" data-sd="' + sd + '" data-row="' + String(d._row || '') + '">'
        + cells
        + '<td class="dds-actions-col"><div class="row-actions dds-row-actions">'
        + '<button type="button" class="btn-row btn-export dds-row-export" data-sd="' + sd + '" title="Export PDF &amp; DOCX">'
        + '<svg class="btn-row-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3v12m0 0 4-4m-4 4-4-4M5 21h14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
        + 'Export</button>'
        + '<button type="button" class="btn-row btn-edit dds-row-edit" data-sd="' + sd + '" title="Edit">'
        + '<svg class="btn-row-icon" viewBox="0 0 24 24" aria-hidden="true"><path d="M4 20h4l10.5-10.5a1.5 1.5 0 0 0-4-4L4 16v4z" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/></svg>'
        + 'Edit</button>'
        + '</div></td>'
        + '</tr>';
    }).join('');
  }

  async function loadDdsData_(force, opts) {
    opts = opts || {};
    const silent = !!opts.silent;
    if (ddsLoadPromise) {
      if (!force) return ddsLoadPromise;
      try { await ddsLoadPromise; } catch (e) { /* superseded reload */ }
    }

    const loading = document.getElementById('dds-loading');
    const errEl = document.getElementById('dds-error');
    const table = document.getElementById('ddsTable');
    if (loading && !silent) loading.style.display = '';
    if (errEl) errEl.style.display = 'none';
    if (table && !force && !silent) table.style.display = 'none';

    function fetchWithTimeout_(promise, ms) {
      return Promise.race([
        promise,
        new Promise(function(_, reject) {
          setTimeout(function() {
            reject(new Error('Request timed out — check GAS URL / network'));
          }, ms);
        }),
      ]);
    }

    ddsLoadPromise = fetchWithTimeout_(Promise.all([
      apiGet('eudrDds'),
      apiGet('eudrDdsSuppliers'),
      apiGet('eudrDdsGeolocation'),
      apiGet('eudrDdsDocuments'),
    ]), 45000).then(function(results) {
      ddsMaster = (results[0] || []).map(ddsPrepareMasterRow_);
      ddsSuppliers = results[1] || [];
      ddsGeo = results[2] || [];
      ddsDocuments = results[3] || [];
      ddsLoaded = true;
      ddsUpdateStats_();
      renderDdsTable_();
      if (errEl) errEl.style.display = 'none';
    }).catch(function(err) {
      console.warn('[DDS] Load failed:', err);
      ddsLoaded = true;
      ddsMaster = ddsMaster || [];
      ddsUpdateStats_();
      renderDdsTable_();
      if (errEl) {
        errEl.style.display = '';
        errEl.textContent = 'Failed to load DDS data: ' + (err.message || err);
      }
      ddsToast_('Failed to load DDS: ' + (err.message || err), 'error');
    }).finally(function() {
      if (loading) loading.style.display = 'none';
      if (table) table.style.display = 'table';
      ddsLoadPromise = null;
    });
    return ddsLoadPromise;
  }

  function ddsFieldHtml_(field, value, opts) {
    opts = opts || {};
    const label = ddsLabel_(field);
    const val = value != null ? String(value) : '';
    const req = opts.required ? ' required' : '';
    const ro = opts.readonly ? ' readonly disabled' : '';
    if (DDS_DATE_FIELDS.has(field) && typeof dashDateFieldHtml === 'function') {
      return dashDateFieldHtml(field, val, { label: label, required: !!opts.required });
    }
    if (field === 'HS CODE') {
      return ddsHsFieldHtml_(val, { required: !!opts.required });
    }
    if (field === 'COUNTRY RISK CATEGORY') {
      const optsHtml = ['<option value="">— Select —</option>'].concat(DDS_COUNTRY_RISK.map(function(c) {
        return '<option value="' + escHtml(c) + '"' + (val === c ? ' selected' : '') + '>' + escHtml(c) + '</option>';
      })).join('');
      return ''
        + '<div class="form-field"><label>' + escHtml(label) + '</label>'
        + '<select data-dds-field="' + escHtml(field) + '"' + req + '>' + optsHtml + '</select></div>';
    }
    if (field === 'RA OVERALL RISK') {
      const optsHtml = ['<option value="">— Select —</option>'].concat(DDS_RA_RISK.map(function(c) {
        return '<option value="' + escHtml(c) + '"' + (val === c ? ' selected' : '') + '>' + escHtml(c) + '</option>';
      })).join('');
      return ''
        + '<div class="form-field"><label>' + escHtml(label) + '</label>'
        + '<select data-dds-field="' + escHtml(field) + '"' + req + '>' + optsHtml + '</select></div>';
    }
    if (field === 'RA METHODOLOGY' || field === 'RA MITIGATION SUMMARY' || field === 'EXPORTER ADDRESS' || field === 'BUYER ADDRESS') {
      return ''
        + '<div class="form-field full"><label>' + escHtml(label) + '</label>'
        + '<textarea data-dds-field="' + escHtml(field) + '" rows="3"' + req + ro + '>' + escHtml(val) + '</textarea></div>';
    }
    return ''
      + '<div class="form-field"><label>' + escHtml(label) + '</label>'
      + '<input type="text" data-dds-field="' + escHtml(field) + '" value="' + escHtml(val) + '"' + req + ro + ' />'
      + '</div>';
  }

  function ddsRenderFieldGrids_(master) {
    const data = master || {};
    const grids = {
      main: document.getElementById('ddsFieldsMain'),
      admin: document.getElementById('ddsFieldsAdmin'),
      identity: document.getElementById('ddsFieldsIdentity'),
      product: document.getElementById('ddsFieldsProduct'),
      shipment: document.getElementById('ddsFieldsShipment'),
      geoMeta: document.getElementById('ddsFieldsGeoMeta'),
      risk: document.getElementById('ddsFieldsRisk'),
      sign: document.getElementById('ddsFieldsSign'),
    };
    Object.keys(DDS_FIELD_SECTIONS).forEach(function(section) {
      const hostKey = section === 'geoMeta' ? 'geoMeta' : section;
      const el = grids[hostKey];
      if (!el) return;
      el.innerHTML = DDS_FIELD_SECTIONS[section].map(function(field) {
        return ddsFieldHtml_(field, ddsPick_(data, field), {
          required: field === 'SD NUMBER',
          readonly: ddsFormMode === 'edit' && field === 'SD NUMBER',
        });
      }).join('');
    });
    ddsInitHsField_(document.querySelector('#ddsFieldsMain .dds-hs-field'));
    const formBody = document.getElementById('ddsFormBody');
    if (formBody) initDashDateFields(formBody);
  }

  function ddsSupplierRowHtml_(row, index) {
    const typeOpts = ['<option value="">—</option>'].concat(DDS_SUPPLIER_TYPES.map(function(t) {
      const sel = String(row['SUPPLIER TYPE'] || '') === t ? ' selected' : '';
      return '<option value="' + escHtml(t) + '"' + sel + '>' + escHtml(t) + '</option>';
    })).join('');
    return ''
      + '<tr data-supplier-row="' + index + '">'
      + '<td>' + (index + 1) + '</td>'
      + '<td><input type="text" data-supplier-field="SUPPLIER NAME" value="' + escHtml(String(row['SUPPLIER NAME'] || '')) + '" /></td>'
      + '<td><input type="text" data-supplier-field="SUPPLIER ADDRESS" value="' + escHtml(String(row['SUPPLIER ADDRESS'] || '')) + '" /></td>'
      + '<td><select data-supplier-field="SUPPLIER TYPE">' + typeOpts + '</select></td>'
      + '<td><input type="text" data-supplier-field="SUPPLIER DDS REF" value="' + escHtml(String(row['SUPPLIER DDS REF'] || '')) + '" /></td>'
      + '<td><input type="text" data-supplier-field="PROOF DOC" value="' + escHtml(String(row['PROOF DOC'] || '')) + '" /></td>'
      + '<td><button type="button" class="btn-icon dds-remove-row" data-remove-supplier="' + index + '" title="Remove">✕</button></td>'
      + '</tr>';
  }

  function ddsGeoRowHtml_(row, index) {
    return ''
      + '<tr data-geo-row="' + index + '">'
      + '<td>' + (index + 1) + '</td>'
      + '<td><input type="text" data-geo-field="PLOT ID" value="' + escHtml(String(row['PLOT ID'] || '')) + '" /></td>'
      + '<td><input type="text" data-geo-field="AREA HA" value="' + escHtml(String(row['AREA HA'] || '')) + '" /></td>'
      + '<td><input type="text" data-geo-field="COORDINATES" value="' + escHtml(String(row.COORDINATES || '')) + '" /></td>'
      + '<td><input type="text" data-geo-field="HARVEST DATE" value="' + escHtml(String(row['HARVEST DATE'] || '')) + '" placeholder="DD/MM/YYYY" /></td>'
      + '<td><input type="text" data-geo-field="NOTES" value="' + escHtml(String(row.NOTES || '')) + '" /></td>'
      + '<td><button type="button" class="btn-icon dds-remove-row" data-remove-geo="' + index + '" title="Remove">✕</button></td>'
      + '</tr>';
  }

  function ddsDocRowHtml_(row) {
    const code = String(row['DOC CODE'] || '').trim();
    const def = DDS_DOC_CODES.find(function(d) { return d.code === code; }) || { label: code };
    const avail = String(row.AVAILABLE || '').trim().toUpperCase();
    return ''
      + '<tr data-doc-code="' + escHtml(code) + '">'
      + '<td><strong>' + escHtml(code) + '</strong></td>'
      + '<td class="dds-doc-label">' + escHtml(def.label) + '</td>'
      + '<td><input type="text" data-doc-field="DOC NUMBER" value="' + escHtml(String(row['DOC NUMBER'] || '')) + '" /></td>'
      + '<td><input type="text" data-doc-field="DOC DATE" value="' + escHtml(String(row['DOC DATE'] || '')) + '" placeholder="DD/MM/YYYY" /></td>'
      + '<td><select data-doc-field="AVAILABLE">'
      + '<option value=""' + (!avail ? ' selected' : '') + '>—</option>'
      + '<option value="Y"' + (avail === 'Y' ? ' selected' : '') + '>Y</option>'
      + '<option value="N"' + (avail === 'N' ? ' selected' : '') + '>N</option>'
      + '</select></td>'
      + '<td><input type="text" data-doc-field="NOTES" value="' + escHtml(String(row.NOTES || '')) + '" /></td>'
      + '</tr>';
  }

  function ddsRenderChildTables_(bundle) {
    const suppliers = (bundle && bundle.suppliers) || [];
    const geo = (bundle && bundle.geolocation) || [];
    let docs = (bundle && bundle.documents) || [];
    if (!docs.length) docs = ddsDefaultDocuments_();
    else {
      docs = DDS_DOC_CODES.map(function(def) {
        const hit = docs.find(function(d) { return String(d['DOC CODE'] || '').trim() === def.code; });
        return hit ? Object.assign({ 'DOC CODE': def.code }, hit) : { 'DOC CODE': def.code };
      });
    }

    const supBody = document.getElementById('ddsSupplierBody');
    const geoBody = document.getElementById('ddsGeoBody');
    const docBody = document.getElementById('ddsDocBody');
    const supRows = suppliers.length ? suppliers.slice().sort(function(a, b) {
      return Number(a['LINE NO'] || 0) - Number(b['LINE NO'] || 0);
    }) : [{ 'SUPPLIER NAME': '', 'SUPPLIER ADDRESS': '', 'SUPPLIER TYPE': '', 'SUPPLIER DDS REF': '', 'PROOF DOC': '' }];
    const geoRows = geo.length ? geo.slice().sort(function(a, b) {
      return Number(a['LINE NO'] || 0) - Number(b['LINE NO'] || 0);
    }) : [];

    if (supBody) supBody.innerHTML = supRows.map(function(r, i) { return ddsSupplierRowHtml_(r, i); }).join('');
    if (geoBody) geoBody.innerHTML = geoRows.length
      ? geoRows.map(function(r, i) { return ddsGeoRowHtml_(r, i); }).join('')
      : ddsGeoRowHtml_({}, 0);
    if (docBody) docBody.innerHTML = docs.map(ddsDocRowHtml_).join('');
  }

  function ddsScrollToSection_(sectionId) {
    const el = document.getElementById(sectionId);
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function openDdsForm_(mode, sdNumber) {
    try {
      const overlay = mountOverlay ? mountOverlay('ddsFormOverlay') : document.getElementById('ddsFormOverlay');
      if (!overlay) {
        ddsToast_('Form overlay not found.', 'error');
        return;
      }
      ddsFormMode = mode || 'add';
      const bundle = sdNumber ? ddsBundleBySd_(sdNumber) : { master: {}, suppliers: [], geolocation: [], documents: ddsDefaultDocuments_() };
      ddsFormMasterRow = bundle.master;
      const master = Object.assign({}, bundle.master || {});

      const titleEl = document.getElementById('ddsFormTitle');
      const subEl = document.getElementById('ddsFormSubtitle');
      if (titleEl) titleEl.textContent = ddsFormMode === 'edit' ? 'Edit Due Diligence Statement' : 'New Due Diligence Statement';
      if (subEl) subEl.textContent = ddsFormMode === 'edit'
        ? ('SD Number: ' + (ddsPick_(master, 'SD NUMBER') || '—'))
        : 'Enter shipment data — saved to Google Sheet. Export PDF & DOCX after Save.';

      ddsRenderFieldGrids_(master);
      ddsRenderChildTables_(bundle);
      overlay.classList.add('active');
      document.body.classList.add('bl-overlay-scroll-lock');
      const bodyEl = document.getElementById('ddsFormBody');
      if (bodyEl) bodyEl.scrollTop = 0;
      overlay.scrollTop = 0;
    } catch (err) {
      console.error('[DDS] openDdsForm_ failed:', err);
      ddsToast_('Could not open form: ' + (err.message || err), 'error');
    }
  }

  function closeDdsForm_() {
    const overlay = document.getElementById('ddsFormOverlay');
    if (overlay) overlay.classList.remove('active');
    document.body.classList.remove('bl-overlay-scroll-lock');
    ddsFormMasterRow = null;
  }

  function ddsCollectMaster_() {
    const body = document.getElementById('ddsFormBody');
    if (!body) return {};
    const data = {};
    dashDateCollectValues(body);
    body.querySelectorAll('[data-dds-field]').forEach(function(el) {
      const f = el.getAttribute('data-dds-field');
      if (!f || el.disabled) return;
      data[f] = el.value != null ? String(el.value).trim() : '';
    });
    body.querySelectorAll('.dash-date-value[data-field]').forEach(function(el) {
      const f = el.getAttribute('data-field');
      if (!f || el.disabled) return;
      const iso = el.value != null ? String(el.value).trim() : '';
      data[f] = iso ? dashIsoToDisplay(iso) : '';
    });
    return data;
  }

  function ddsCollectSuppliers_() {
    const body = document.getElementById('ddsSupplierBody');
    if (!body) return [];
    return Array.from(body.querySelectorAll('tr[data-supplier-row]')).map(function(tr, i) {
      const row = { 'LINE NO': String(i + 1) };
      tr.querySelectorAll('[data-supplier-field]').forEach(function(el) {
        row[el.getAttribute('data-supplier-field')] = el.value != null ? String(el.value).trim() : '';
      });
      return row;
    }).filter(function(r) {
      return String(r['SUPPLIER NAME'] || r['SUPPLIER ADDRESS'] || r['PROOF DOC'] || '').trim();
    });
  }

  function ddsCollectGeo_() {
    const body = document.getElementById('ddsGeoBody');
    if (!body) return [];
    return Array.from(body.querySelectorAll('tr[data-geo-row]')).map(function(tr, i) {
      const row = { 'LINE NO': String(i + 1) };
      tr.querySelectorAll('[data-geo-field]').forEach(function(el) {
        row[el.getAttribute('data-geo-field')] = el.value != null ? String(el.value).trim() : '';
      });
      return row;
    }).filter(function(r) {
      return String(r['PLOT ID'] || r.COORDINATES || r['AREA HA'] || '').trim();
    });
  }

  function ddsCollectDocuments_() {
    const body = document.getElementById('ddsDocBody');
    if (!body) return ddsDefaultDocuments_();
    return Array.from(body.querySelectorAll('tr[data-doc-code]')).map(function(tr) {
      const row = { 'DOC CODE': tr.getAttribute('data-doc-code') || '' };
      tr.querySelectorAll('[data-doc-field]').forEach(function(el) {
        row[el.getAttribute('data-doc-field')] = el.value != null ? String(el.value).trim() : '';
      });
      return row;
    });
  }

  async function saveDdsForm_() {
    if (ddsSaving) return;
    const master = ddsCollectMaster_();
    if (!String(master['SD NUMBER'] || '').trim()) {
      ddsToast_('SD Number is required.', 'warning');
      ddsScrollToSection_('ddsSectionMain');
      return;
    }
    const suppliers = ddsCollectSuppliers_();
    if (!suppliers.length && !String(master['SUPPLIER NAME'] || '').trim()) {
      ddsToast_('At least one supplier is required.', 'warning');
      ddsScrollToSection_('ddsSectionSuppliers');
      return;
    }
    if (suppliers.length && !String(master['SUPPLIER NAME'] || '').trim()) {
      master['SUPPLIER NAME'] = suppliers[0]['SUPPLIER NAME'] || '';
    }
    ddsSaving = true;
    const saveBtn = document.getElementById('ddsFormSave');
    if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'Saving…'; }
    try {
      await apiPost({
        action: 'upsertEudrDds',
        data: {
          master: master,
          suppliers: suppliers,
          geolocation: ddsCollectGeo_(),
          documents: ddsCollectDocuments_(),
        },
      });
      ddsToast_('DDS saved.', 'success');
      closeDdsForm_();
      await loadDdsData_(true, { silent: true });
    } catch (err) {
      ddsToast_('Save failed: ' + (err.message || err), 'error');
    } finally {
      ddsSaving = false;
      if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = 'Save DDS'; }
    }
  }

  async function fetchDdsBundleFromSheet_(sdNumber) {
    const key = ddsNormSd_(sdNumber);
    if (!key) return null;
    const results = await Promise.all([
      apiGet('eudrDds'),
      apiGet('eudrDdsSuppliers'),
      apiGet('eudrDdsGeolocation'),
      apiGet('eudrDdsDocuments'),
    ]);
    const masterRows = (results[0] || []).map(ddsPrepareMasterRow_);
    ddsMaster = masterRows;
    ddsSuppliers = results[1] || [];
    ddsGeo = results[2] || [];
    ddsDocuments = results[3] || [];
    ddsLoaded = true;
    ddsUpdateStats_();
    renderDdsTable_();
    return ddsBundleBySd_(sdNumber);
  }

  async function exportDdsBothFromSheet_(sdNumber) {
    const sd = String(sdNumber || '').trim();
    if (!sd) {
      ddsToast_('SD Number is required.', 'warning');
      return;
    }
    let bundle = null;
    try {
      bundle = await fetchDdsBundleFromSheet_(sd);
    } catch (err) {
      console.warn('[DDS] Sheet reload before export failed:', err);
      bundle = ddsBundleBySd_(sd);
    }
    if (!bundle || !bundle.master) {
      ddsToast_('Record not found in sheet. Save the record first.', 'warning');
      return;
    }
    try {
      ddsToast_('Generating PDF & DOCX from template…', 'info');
      const [{ buildFilledDdsDocxBlob_ }, { exportDdsPdfFromBlob_ }, { downloadFilledDdsDocx_ }] = await Promise.all([
        import('./dds-docx-fill.js'),
        import('./dds-pdf.js'),
        import('./dds-docx.js'),
      ]);
      const filled = await buildFilledDdsDocxBlob_(bundle);

      let pdfOk = false;
      let pdfErr = null;
      try {
        await exportDdsPdfFromBlob_(filled);
        pdfOk = true;
      } catch (err) {
        pdfErr = err;
        console.error('[DDS] PDF export failed:', err);
      }

      downloadFilledDdsDocx_(filled);

      if (pdfOk) {
        ddsToast_('PDF and DOCX exported.', 'success');
      } else {
        ddsToast_('DOCX exported. PDF failed: ' + (pdfErr && pdfErr.message ? pdfErr.message : pdfErr), 'warning');
      }
    } catch (err) {
      ddsToast_('Export failed: ' + (err.message || err), 'error');
    }
  }

  function bindDdsUiOnce_() {
    if (window.__ddsUiBound) return;
    window.__ddsUiBound = true;

    const searchEl = document.getElementById('ddsSearch');
    const clearEl = document.getElementById('ddsSearchClear');
    const btnRefresh = document.getElementById('btn-refresh-dds');
    const btnAdd = document.getElementById('btn-add-dds');
    const tableBody = document.getElementById('ddsTableBody');

    const debouncedRender = typeof debounce === 'function'
      ? debounce(function() { renderDdsTable_(); }, 120)
      : function() { renderDdsTable_(); };

    if (searchEl && clearEl) {
      searchEl.addEventListener('input', function() {
        ddsSearch = this.value.toLowerCase().trim();
        clearEl.classList.toggle('show', !!this.value);
        debouncedRender();
      });
      clearEl.addEventListener('click', function() {
        searchEl.value = '';
        ddsSearch = '';
        this.classList.remove('show');
        renderDdsTable_();
        searchEl.focus();
      });
    }

    if (btnRefresh) btnRefresh.addEventListener('click', function() {
      ddsLoaded = false;
      loadDdsData_(true);
    });

    if (btnAdd) btnAdd.addEventListener('click', function() { openDdsForm_('add'); });

    document.getElementById('ddsFormClose')?.addEventListener('click', closeDdsForm_);
    document.getElementById('ddsFormCancel')?.addEventListener('click', closeDdsForm_);
    document.getElementById('ddsFormSave')?.addEventListener('click', saveDdsForm_);
    document.getElementById('ddsExportBtn')?.addEventListener('click', function() {
      const master = ddsCollectMaster_();
      if (!String(master['SD NUMBER'] || '').trim()) {
        ddsToast_('Enter SD Number first.', 'warning');
        ddsScrollToSection_('ddsSectionMain');
        return;
      }
      exportDdsBothFromSheet_(master['SD NUMBER']);
    });

    document.getElementById('ddsAddSupplierRow')?.addEventListener('click', function() {
      const body = document.getElementById('ddsSupplierBody');
      if (!body) return;
      const idx = body.querySelectorAll('tr[data-supplier-row]').length;
      body.insertAdjacentHTML('beforeend', ddsSupplierRowHtml_({}, idx));
    });

    document.getElementById('ddsAddGeoRow')?.addEventListener('click', function() {
      const body = document.getElementById('ddsGeoBody');
      if (!body) return;
      const idx = body.querySelectorAll('tr[data-geo-row]').length;
      body.insertAdjacentHTML('beforeend', ddsGeoRowHtml_({}, idx));
    });

    document.getElementById('ddsSupplierBody')?.addEventListener('click', function(e) {
      const btn = e.target.closest('[data-remove-supplier]');
      if (!btn) return;
      const tr = btn.closest('tr');
      if (tr) tr.remove();
    });

    document.getElementById('ddsGeoBody')?.addEventListener('click', function(e) {
      const btn = e.target.closest('[data-remove-geo]');
      if (!btn) return;
      const tr = btn.closest('tr');
      if (tr) tr.remove();
    });

    document.getElementById('ddsFormBody')?.addEventListener('change', function(e) {
      if (!e.target.classList.contains('dds-hs-select')) return;
      const wrap = e.target.closest('.dds-hs-field');
      if (!wrap) return;
      const hidden = wrap.querySelector('[data-dds-field="HS CODE"]');
      const customPanel = wrap.querySelector('.dds-hs-custom');
      if (e.target.value === '__custom__') {
        if (customPanel) customPanel.style.display = '';
        if (hidden) hidden.value = '';
        wrap.querySelector('.dds-hs-custom-code')?.focus();
      } else {
        if (customPanel) customPanel.style.display = 'none';
        if (hidden) hidden.value = e.target.value;
      }
    });

    document.getElementById('ddsFormBody')?.addEventListener('click', function(e) {
      const wrap = e.target.closest('.dds-hs-field');
      if (!wrap) return;
      if (e.target.classList.contains('dds-hs-save-btn')) {
        e.preventDefault();
        ddsApplyHsCustomCode_(wrap, true);
      } else if (e.target.classList.contains('dds-hs-use-btn')) {
        e.preventDefault();
        ddsApplyHsCustomCode_(wrap, false);
      }
    });

    if (tableBody) {
      tableBody.addEventListener('click', function(e) {
        const exportBtn = e.target.closest('.dds-row-export');
        if (exportBtn) {
          e.stopPropagation();
          exportDdsBothFromSheet_(exportBtn.getAttribute('data-sd'));
          return;
        }
        const editBtn = e.target.closest('.dds-row-edit');
        if (editBtn) {
          e.stopPropagation();
          openDdsForm_('edit', editBtn.getAttribute('data-sd'));
          return;
        }
        const tr = e.target.closest('.dds-row-clickable');
        if (tr && !e.target.closest('.row-actions')) {
          openDdsForm_('edit', tr.getAttribute('data-sd'));
        }
      });
    }

    const overlay = document.getElementById('ddsFormOverlay');
    if (overlay) {
      overlay.addEventListener('click', function(e) {
        if (e.target === overlay) closeDdsForm_();
      });
    }

    document.addEventListener('keydown', function(e) {
      if (e.key !== 'Escape') return;
      const ov = document.getElementById('ddsFormOverlay');
      if (ov && ov.classList.contains('active')) closeDdsForm_();
    });
  }

  bindDdsUiOnce_();

  return {
    load: loadDdsData_,
    isLoaded: function() { return ddsLoaded; },
    openAdd: function() { openDdsForm_('add'); },
    reload: function() { ddsLoaded = false; return loadDdsData_(true); },
  };
}
