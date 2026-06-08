/**
 * Monthly Report (Detail) — read-only compliance snapshot (fast, in-memory first).
 */

const MRD_ROW_LIMIT = 200;
const MRD_SDD_LIMIT = 150;

let _deps = null;
let _bound = false;
let _snapshot = null;
let _search = '';
let _year = '2025';
let _month = '';
let _expanded = new Set();
let _loadGen = 0;
let _eudrPending = false;
let _nblByCache = new Map();
let _sddCache = [];

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
  if (!rows.length) return opts.empty || '<p class="mrd-empty">Tidak ada data.</p>';
  const active = filterColumns(columns, rows);
  const colCount = active.length;
  let html = (opts.note || '') + '<div class="table-scroll mrd-table-scroll"><table class="mrd-table"><thead><tr>';
  active.forEach(function(col) {
    html += '<th' + (col.thCls ? ' class="' + col.thCls + '"' : '') + '>' + esc(col.label) + '</th>';
  });
  html += '</tr></thead><tbody>';
  rows.forEach(function(row, idx) {
    if (row._before) html += row._before;
    html += '<tr' + (row._trClass ? ' class="' + row._trClass + '"' : '') + '>';
    active.forEach(function(col) {
      html += '<td' + (col.tdCls ? ' class="' + col.tdCls + '"' : '') + '>';
      html += col.render ? col.render(row, idx, colCount) : esc(col.raw ? col.raw(row) : (row[col.key] || '—'));
      html += '</td>';
    });
    html += '</tr>';
    if (row._after) html += row._after;
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

function limitNote(total, limit) {
  if (total <= limit) return '';
  return '<p class="mrd-limit-note">Menampilkan ' + limit + ' dari ' + total + ' baris — gunakan search untuk mempersempit.</p>';
}

function buildSnapshotSync(opts) {
  opts = opts || {};
  const millData = (_deps.getMillData() || []);
  const ttpData = (_deps.getTtpData() || []);
  const ttpFields = _deps.getTtpFields() || [];
  const grvData = (_deps.getGrvData() || []);
  const nblReg = (_deps.getNblRegistry() || []);
  const nblUni = (_deps.getNblUnilever() || []);
  const sddInput = opts.sddRows != null ? opts.sddRows : [];
  const eudrInput = opts.eudrPotential != null ? opts.eudrPotential : [];

  const millYear = _deps.millYearVal;
  const mills = millData.filter(function(r) {
    const y = parseYear(millYear(r));
    return (!_year || !y || y === _year);
  });

  const sddFiltered = sddInput.filter(function(r) {
    const st = String(r['SCR - Screening Status'] || '').trim().toLowerCase();
    if (st !== 'draft' && st !== 'submitted') return false;
    if (_year) {
      const upd = String(r.updated_at || r['SCR - Last Updated'] || '').slice(0, 4);
      if (upd && upd !== _year) return false;
    }
    if (_month) {
      const m = parseMonthFromDate(r.updated_at || r['SCR - Last Updated']);
      if (m && m !== _month) return false;
    }
    return true;
  });

  const millCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'MILL NAME'; }) || 'MILL NAME';
  const supplierCol = ttpFields.find(function(h) { return /FFB SUPPLIER NAME/i.test(h); }) || '';
  const yearCol = ttpFields.find(function(h) { return String(h).toUpperCase() === 'YEAR'; }) || 'YEAR';

  const ttpFiltered = ttpData.filter(function(r) {
    const y = parseYear(r[yearCol] || millYear(r));
    return !_year || !y || y === _year;
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
  emptyMills.sort(function(a, b) {
    return String(a.millRow['MILL NAME'] || '').localeCompare(String(b.millRow['MILL NAME'] || ''));
  });

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

  const grvRows = grvData.filter(function(r) {
    if (_year) {
      const dr = String(r['Date Received'] || '').slice(0, 4);
      if (dr && dr !== _year) return false;
    }
    if (_month) {
      const m = parseMonthFromDate(r['Date Received']);
      if (m && m !== _month) return false;
    }
    return true;
  }).map(function(r) {
    return { row: r, search: Object.keys(r).map(function(k) { return r[k]; }).join(' ').toLowerCase() };
  });

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

  return {
    sdd: sddFiltered,
    mills: millRows,
    millsTotal: millRows.length,
    emptyMills: emptyMills,
    ttpByMill: ttpByMill,
    grv: grvRows,
    nblAll: nblAll,
    facility: facility,
    eudrPotential: eudrInput,
    eudrLoading: !!opts.eudrLoading,
    sddLoading: !!opts.sddLoading,
    stats: {
      sddTotal: sddFiltered.length,
      sddDraft: sddFiltered.filter(function(r) {
        return String(r['SCR - Screening Status'] || '').toLowerCase() === 'draft';
      }).length,
      sddSubmitted: sddFiltered.filter(function(r) {
        return String(r['SCR - Screening Status'] || '').toLowerCase() === 'submitted';
      }).length,
      totalMills: mills.length,
      totalGroups: new Set(mills.map(function(r) { return r['GROUP NAME']; }).filter(Boolean)).size,
      highRisk: mills.filter(function(r) { return isHighRisk(_deps.millResolvedRiskLevel(r)); }).length,
      nblMills: mills.filter(function(r) { return isNblYes(r['BUYER NO BUY LIST']); }).length,
      emptyTraceMills: emptyMills.length,
      grievances: grvRows.length,
      nblEntries: nblAll.length,
      eudrPotential: eudrInput.length,
      facilities: facility.cpo.length + facility.pk.length,
    },
  };
}

function renderKpis(stats) {
  const el = document.getElementById('mrdKpiRow');
  if (!el) return;
  const items = [
    { n: stats.sddSubmitted, l: 'SDD Submitted', s: stats.sddDraft + ' draft' },
    { n: stats.totalMills, l: 'Total Mills', s: stats.totalGroups + ' groups' },
    { n: stats.highRisk, l: 'High Risk', s: stats.nblMills + ' NBL mills', hot: true },
    { n: stats.emptyTraceMills, l: 'Empty Traceability', s: 'mills tanpa supplier' },
    { n: stats.grievances, l: 'Grievances', s: 'in period' },
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
  if (loading) return '<p class="mrd-empty mrd-empty--loading">Memuat Supplier Due Diligence…</p>';
  const rows = data.filter(function(r) {
    return matchesSearch([
      r['SCR - Screening Status'], r['Group Name'], r['Grup Name'], r['Mill Name'],
      r.supplier_type, r.updated_at,
    ].join(' ').toLowerCase());
  }).slice(0, MRD_SDD_LIMIT);
  const cols = [
    { label: 'Status', always: true, render: function(r) { return statusPill(r['SCR - Screening Status'], 'sdd'); } },
    { label: 'Type', raw: function(r) { return r.supplier_type || r['Supplier Type']; } },
    { label: 'Group', raw: function(r) { return r['Group Name'] || r['Grup Name']; } },
    { label: 'Mill', raw: function(r) { return r['Mill Name']; } },
    { label: 'Updated', raw: function(r) { return String(r.updated_at || r['SCR - Last Updated'] || '').slice(0, 10); } },
  ];
  return renderSmartTable(cols, rows, {
    empty: '<p class="mrd-empty">Belum ada data SDD untuk periode ini.</p>',
    note: limitNote(data.length, MRD_SDD_LIMIT),
  });
}

function renderMillSection(rows) {
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); });
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
    { label: 'Group', raw: function(row) { return row._render.r['GROUP NAME']; } },
    { label: 'Company', raw: function(row) { return row._render.r['COMPANY NAME']; } },
    { label: 'Mill', raw: function(row) { return row._render.r['MILL NAME']; } },
    { label: 'Risk', hasData: function(row) { return hasCellValue(row._render.item.risk); }, render: function(row) { return riskPill(row._render.item.risk); } },
    { label: 'NBL', hasData: function(row) { return isNblYes(row._render.item.nbl); }, render: function(row) {
      return isNblYes(row._render.item.nbl) ? '<span class="mrd-pill mrd-pill--high">Yes</span>' : '';
    }},
    { label: 'Province', raw: function(row) { return row._render.r['PROVINCE']; } },
  ];
  return renderSmartTable(cols, tableRows, {
    empty: '<p class="mrd-empty">Tidak ada mill untuk periode ini.</p>',
    note: limitNote(filtered.length, MRD_ROW_LIMIT),
  });
}

function renderMillDetailHtml_(item, r) {
  const grid = [
    detailItem('Supplier Status', r['SUPPLIER STATUS']),
    detailItem('Certification', r['CERTIFICATION']),
    detailItem('Total Grievances', r['TOTAL GRIEVANCES']),
    detailItem('Facility CPO', r['FACILITY NAME CPO']),
    detailItem('Facility PK', r['FACILITY NAME PK']),
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
    html += '<p class="mrd-detail-note">NBL source — klik expand lagi jika belum muncul.</p>';
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

function renderTraceSection(emptyMills, ttpByMill) {
  const items = emptyMills.filter(function(item) {
    return matchesSearch([
      item.millRow['GROUP NAME'], item.millRow['COMPANY NAME'], item.millRow['MILL NAME'],
    ].join(' ').toLowerCase());
  }).slice(0, MRD_ROW_LIMIT);
  const tableRows = items.map(function(item, idx) {
    const r = item.millRow;
    const expId = 'ttp-' + idx;
    const isExp = _expanded.has(expId);
    const millName = String(r['MILL NAME'] || '').trim();
    const subRows = ttpByMill.get(millName) || [];
    const out = { _render: { r: r, expId: expId, isExp: isExp, millName: millName, subRows: subRows } };
    out._after = isExp
      ? '<tr class="mrd-detail-row"><td colspan="99"><div class="mrd-detail-panel mrd-detail-panel--trace">'
        + '<p class="mrd-detail-note">Belum ada supplier FFB (tahun ' + esc(_year) + ').'
        + (subRows.length ? ' · ' + subRows.length + ' baris TTP tanpa supplier.' : '')
        + '</p></div></td></tr>'
      : '';
    return out;
  });
  const cols = [
    { label: '', always: true, thCls: 'mrd-th-expand', tdCls: 'mrd-td-expand',
      render: function(row) {
        const d = row._render;
        return '<button type="button" class="mrd-expand-btn' + (d.isExp ? ' is-open' : '') + '" data-mrd-expand-btn="' + esc(d.expId) + '">›</button>';
      }},
    { label: 'Group', raw: function(row) { return row._render.r['GROUP NAME']; } },
    { label: 'Company', raw: function(row) { return row._render.r['COMPANY NAME']; } },
    { label: 'Mill', raw: function(row) { return row._render.millName; } },
    { label: 'Status', always: true, render: function() { return '<span class="mrd-pill mrd-pill--warn">Kosong</span>'; } },
  ];
  return renderSmartTable(cols, tableRows, {
    empty: '<p class="mrd-empty mrd-empty--ok">Semua mill sudah punya data supplier untuk ' + esc(_year) + '.</p>',
    note: limitNote(emptyMills.length, MRD_ROW_LIMIT),
  });
}

function renderGrvSection(rows) {
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const fieldDefs = [
    { key: 'Grievance ID', label: 'ID' },
    { key: 'Date Received', label: 'Date Received' },
    { key: 'Complainant', label: 'Complainant' },
    { key: 'Grievance Category', label: 'Category' },
    { key: 'Subject', label: 'Subject' },
    { key: 'Risk Classification', label: 'Risk', pill: 'risk' },
    { key: 'Grievance Status', label: 'Status', pill: 'grv' },
  ];
  const activeFields = fieldDefs.filter(function(f) {
    return filtered.some(function(item) { return hasCellValue(item.row[f.key]); });
  });
  const tableRows = filtered.map(function(item, idx) {
    const r = item.row;
    const expId = 'grv-' + idx;
    const isExp = _expanded.has(expId);
    const detailFields = [
      ['Description', r['Grievance Description']],
      ['Verification', r['Verification Findings']],
      ['Corrective Action', r['Corrective Action']],
      ['Preventive Action', r['Preventive Action']],
      ['Action Taken', r['Action Taken']],
      ['Date Closed', r['Date Closed']],
    ].filter(function(pair) { return hasCellValue(pair[1]); });
    const out = { _render: { r: r, expId: expId, isExp: isExp, fields: activeFields } };
    if (isExp && detailFields.length) {
      let grid = detailFields.map(function(pair) { return detailItem(pair[0], pair[1]); }).join('');
      out._after = '<tr class="mrd-detail-row"><td colspan="99"><div class="mrd-detail-panel"><div class="mrd-detail-grid">' + grid + '</div></div></td></tr>';
    }
    return out;
  });
  const cols = [
    { label: '', always: true, thCls: 'mrd-th-expand', tdCls: 'mrd-td-expand',
      render: function(row) {
        const d = row._render;
        return '<button type="button" class="mrd-expand-btn' + (d.isExp ? ' is-open' : '') + '" data-mrd-expand-btn="' + esc(d.expId) + '">›</button>';
      }},
  ].concat(activeFields.map(function(f) {
    return {
      label: f.label,
      render: function(row) {
        const v = row._render.r[f.key];
        if (f.pill === 'grv') return statusPill(v, 'grv');
        if (f.pill === 'risk') return riskPill(v);
        return esc(v || '—');
      },
      hasData: function(row) { return hasCellValue(row._render.r[f.key]); },
    };
  }));
  return renderSmartTable(cols, tableRows, {
    empty: '<p class="mrd-empty">Tidak ada grievance untuk periode ini.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
  });
}

function renderNblSection(rows) {
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const cols = [
    { label: 'Source', always: true, raw: function(r) { return r.source; } },
    { label: 'Riser', raw: function(r) { return r.riser; } },
    { label: 'Group', raw: function(r) { return r.group; } },
    { label: 'Company', raw: function(r) { return r.company; } },
  ];
  return renderSmartTable(cols, filtered, {
    empty: '<p class="mrd-empty">Tidak ada entri NBL.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
  });
}

function renderFacilitySection(facility) {
  function renderBlock(title, groups, accent) {
    const visible = groups.filter(function(g) {
      return matchesSearch([g.facility, g.companies, g.tracePct].join(' ').toLowerCase());
    }).slice(0, MRD_ROW_LIMIT);
    const cols = [
      { label: 'Facility', always: true, raw: function(g) { return g.facility; }, tdCls: 'mrd-td-facility' },
      { label: 'Companies', hasData: function(g) { return !isZeroish(g.companies); }, raw: function(g) { return g.companies; }, tdCls: 'mrd-td-num' },
      { label: 'High Risk', hasData: function(g) { return !isZeroish(g.highRisk); }, raw: function(g) { return g.highRisk; }, tdCls: 'mrd-td-num mrd-td-warn' },
      { label: 'NBL', hasData: function(g) { return !isZeroish(g.nbl); }, raw: function(g) { return g.nbl; }, tdCls: 'mrd-td-num mrd-td-danger' },
      { label: 'Grievance', hasData: function(g) { return !isZeroish(g.grievance); }, raw: function(g) { return g.grievance; }, tdCls: 'mrd-td-num' },
      { label: '% Traceable', hasData: function(g) { return !isZeroish(g.tracePct); }, raw: function(g) { return g.tracePct; }, tdCls: 'mrd-td-pct' },
    ];
    const table = renderSmartTable(cols, visible, {
      empty: '<p class="mrd-empty">Tidak ada data ' + esc(title) + '.</p>',
      note: limitNote(groups.length, MRD_ROW_LIMIT),
    });
    return '<div class="mrd-facility-block mrd-facility-block--' + accent + '"><div class="mrd-facility-title">' + esc(title) + '</div>' + table + '</div>';
  }
  const cpo = renderBlock('CPO Facility Performance', facility.cpo, 'cpo');
  const pk = renderBlock('PK Facility Performance', facility.pk, 'pk');
  if (cpo.indexOf('mrd-empty') !== -1 && pk.indexOf('mrd-empty') !== -1) {
    return '<p class="mrd-empty">Tidak ada data facility performance.</p>';
  }
  return cpo + pk;
}

function renderEudrSection(rows, loading) {
  if (loading) return '<p class="mrd-empty mrd-empty--loading">Memuat EUDR Potential…</p>';
  const filtered = rows.filter(function(item) { return matchesSearch(item.search); }).slice(0, MRD_ROW_LIMIT);
  const cols = [
    { label: 'Status', always: true, render: function() { return statusPill('Potential', 'eudr'); } },
    { label: 'Group', raw: function(item) { return item.row['GROUP NAME']; } },
    { label: 'Company', raw: function(item) { return item.row['COMPANY NAME']; } },
    { label: 'Mill', raw: function(item) { return item.row['MILL NAME']; } },
    { label: 'Province', raw: function(item) { return item.row['PROVINCE']; } },
    { label: 'Supply To', raw: function(item) { return item.row['SUPPLY TO']; } },
  ];
  return renderSmartTable(cols, filtered, {
    empty: '<p class="mrd-empty">Tidak ada mill berstatus Potential.</p>',
    note: limitNote(rows.length, MRD_ROW_LIMIT),
  });
}

function renderAll() {
  if (!_snapshot) return;
  renderKpis(_snapshot.stats);
  const sections = document.getElementById('mrdSections');
  if (!sections) return;
  const s = _snapshot;
  let html = '';
  html += sectionHtml('sdd', 'Supplier Due Diligence', 'Progress — submitted / draft', renderSddSection(s.sdd, s.sddLoading), '01');
  html += sectionHtml('mill', 'Mill Onboarding', 'Mills, groups, high risk & NBL (expand untuk detail)', renderMillSection(s.mills), '02');
  html += sectionHtml('trace', 'Traceability Data ' + _year, 'Mill tanpa supplier — expand per mill', renderTraceSection(s.emptyMills, s.ttpByMill), '03');
  html += sectionHtml('grv', 'Grievance Monitoring', 'Tabel + detail expand', renderGrvSection(s.grv), '04');
  html += sectionHtml('nbl', 'No Buy List', 'Registry + Unilever', renderNblSection(s.nblAll), '05');
  html += sectionHtml('facility', 'Facility Performance', 'Ringkasan CPO & PK', renderFacilitySection(s.facility), '06');
  html += sectionHtml('eudr', 'EUDR Potential', 'Hanya status Potential (formula umum)', renderEudrSection(s.eudrPotential, s.eudrLoading), '07');
  sections.innerHTML = html;
}

function updateScopeText(extra) {
  const el = document.getElementById('mrdScopeText');
  if (!el) return;
  const monthNames = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const mLabel = _month ? monthNames[parseInt(_month, 10)] : 'All months';
  let txt = 'Period: ' + (_year || 'All years') + ' · ' + mLabel;
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

async function loadEudrInBackground(gen) {
  if (_eudrPending || !_deps.fetchEudrPotential) return;
  _eudrPending = true;
  if (_snapshot) {
    _snapshot.eudrLoading = true;
    renderAll();
  }
  try {
    const eudr = await _deps.fetchEudrPotential();
    if (gen !== _loadGen || !_snapshot) return;
    _snapshot.eudrPotential = eudr;
    _snapshot.eudrLoading = false;
    _snapshot.stats.eudrPotential = eudr.length;
    renderAll();
    updateScopeText();
  } catch (_) {
    if (gen === _loadGen && _snapshot) {
      _snapshot.eudrLoading = false;
      _snapshot.eudrPotential = [];
      renderAll();
    }
  } finally {
    _eudrPending = false;
  }
}

function rebuildSnapshot_(opts) {
  const prev = _snapshot || {};
  return buildSnapshotSync({
    sddRows: opts.sddRows != null ? opts.sddRows : _sddCache,
    eudrPotential: opts.eudrPotential != null ? opts.eudrPotential : (prev.eudrPotential || []),
    eudrLoading: opts.eudrLoading != null ? opts.eudrLoading : !!prev.eudrLoading,
    sddLoading: opts.sddLoading != null ? opts.sddLoading : !!prev.sddLoading,
  });
}

async function loadMillInBackground(gen) {
  const loadMill = _deps.ensureMillData || _deps.ensureCoreData;
  if (!loadMill) return;
  try {
    await withTimeout(loadMill(), 30000, 'Mill data');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    renderAll();
    updateScopeText();
    populateYearSelect();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('mill: ' + (err && err.message ? err.message : String(err)));
  }
}

async function loadSupplementalInBackground(gen) {
  if (!_deps.ensureSupplementalData) return;
  try {
    await withTimeout(_deps.ensureSupplementalData(), 45000, 'Traceability & grievance');
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({});
    renderAll();
    updateScopeText();
  } catch (err) {
    if (gen !== _loadGen) return;
    updateScopeText('data tambahan: timeout — refresh untuk coba lagi');
  }
}

async function loadSddInBackground(gen) {
  if (!_deps.fetchSddList) return;
  _snapshot = rebuildSnapshot_({ sddLoading: true });
  renderAll();
  try {
    const rows = await withTimeout(_deps.fetchSddList(), 20000, 'SDD');
    if (gen !== _loadGen) return;
    _sddCache = Array.isArray(rows) ? rows : [];
    _snapshot = rebuildSnapshot_({ sddRows: _sddCache, sddLoading: false });
    renderAll();
    updateScopeText();
  } catch (err) {
    if (gen !== _loadGen) return;
    _snapshot = rebuildSnapshot_({ sddLoading: false });
    renderAll();
    updateScopeText('SDD: ' + (err && err.message ? err.message : 'gagal memuat'));
  }
}

async function loadAndRender() {
  const gen = ++_loadGen;
  const errEl = document.getElementById('mrdError');
  if (errEl) { errEl.hidden = true; errEl.textContent = ''; }

  // Selalu tampilkan UI dari cache memori — jangan blokir menunggu network
  _snapshot = rebuildSnapshot_({ eudrLoading: _eudrPending });
  setUiLoading(false);
  renderAll();
  updateScopeText('memuat data…');

  loadMillInBackground(gen);
  loadSupplementalInBackground(gen);
  loadSddInBackground(gen);
  loadEudrInBackground(gen);
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

  const yearSel = document.getElementById('mrdYearSel');
  const monthSel = document.getElementById('mrdMonthSel');
  const searchEl = document.getElementById('mrdSearch');

  if (yearSel && !yearSel._mrdBound) {
    yearSel._mrdBound = true;
    yearSel.addEventListener('change', function() {
      _year = yearSel.value || '2025';
      _nblByCache.clear();
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
    if (_deps.clearEudrCache) _deps.clearEudrCache();
    loadAndRender();
  });

  panel.addEventListener('click', function(e) {
    const toggleSec = e.target.closest('[data-mrd-toggle]');
    if (toggleSec) {
      const id = toggleSec.getAttribute('data-mrd-toggle');
      const key = id + ':closed';
      if (_expanded.has(key)) _expanded.delete(key);
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
  window.refreshMonthlyReport_ = function() { loadAndRender(); };
  bindOnce();
  loadAndRender();
}

export function refreshMonthlyReport_() {
  if (_deps) loadAndRender();
}
