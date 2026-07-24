/**
 * Monthly Report — Executive view (charts + PDF export).
 * Includes QoQ mill growth and sheet-aligned supply quantity totals.
 */
import {
  classifyRiskBucket_,
  topEntries_,
  quarterEndMonth_,
  quarterMonthRangeLabel_,
} from './mill-executive-report.js';
import {
  normalizeSddCategory,
  mrdReportHeaderMeta_,
  mrdShowInMillOnboarding_,
} from './monthly-report-labels.js';

const CHART_COLORS = {
  High: '#C03030',
  Medium: '#D4A017',
  Low: '#2E7D32',
  Other: '#7A6A6A',
  Unclassified: '#B8A8A8',
  'No Buy List': '#8B1A1A',
  'Non-NBL': '#4A6741',
  Draft: '#D4A017',
  Submitted: '#2E7D32',
  'With Supplier': '#2E7D32',
  Untraceable: '#C03030',
  Potential: '#00838F',
  CPO: '#8B1A1A',
  PK: '#2E7D32',
  'POME ISCC': '#00838F',
  'POME INS': '#1565C0',
  'SHELL GGL': '#6D4C41',
};

let chartInstances = {};

export function destroyMrdExecutiveCharts_() {
  Object.keys(chartInstances).forEach(function(k) {
    try { chartInstances[k].destroy(); } catch (_) { /* noop */ }
  });
  chartInstances = {};
}

export function monthToQuarter_(month) {
  const m = parseInt(String(month || ''), 10);
  if (m >= 1 && m <= 12) return Math.ceil(m / 3);
  return 0;
}

export function previousQuarter_(year, quarter) {
  const y = parseInt(String(year || ''), 10);
  const q = parseInt(String(quarter || ''), 10);
  if (!y || q < 1 || q > 4) return null;
  if (q === 1) return { year: y - 1, quarter: 4 };
  return { year: y, quarter: q - 1 };
}

export function formatDelta_(n) {
  const v = Number(n) || 0;
  if (v > 0) return '+' + v;
  return String(v);
}

function parsePctNum_(raw) {
  const v = parseFloat(String(raw == null ? '' : raw).replace('%', '').replace(',', '.').trim());
  return isNaN(v) ? null : v;
}

function formatQtyDisplay_(n) {
  const v = Number(n) || 0;
  if (!v) return '0';
  return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
}

const PIE_PCT_PLUGIN = {
  id: 'piePct',
  afterDraw: function(chart) {
    if (chart.config.type !== 'pie' && chart.config.type !== 'doughnut') return;
    const ctx = chart.ctx;
    chart.data.datasets.forEach(function(dataset, di) {
      const meta = chart.getDatasetMeta(di);
      const total = dataset.data.reduce(function(a, b) { return a + b; }, 0);
      meta.data.forEach(function(arc, idx) {
        const val = dataset.data[idx];
        if (!val || !total) return;
        const pct = Math.round((val / total) * 100);
        if (pct < 4) return;
        const pos = arc.tooltipPosition();
        const fontSize = Math.max(12, Math.round(chart.width * 0.025));
        ctx.save();
        ctx.font = 'bold ' + fontSize + 'px system-ui, sans-serif';
        ctx.fillStyle = '#ffffff';
        ctx.shadowColor = 'rgba(0,0,0,0.45)';
        ctx.shadowBlur = 3;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(pct + '%', pos.x, pos.y);
        ctx.restore();
      });
    });
  },
};

function pieOptions_(canvas) {
  return {
    responsive: false,
    animation: false,
    width: canvas ? canvas.width : 700,
    height: canvas ? canvas.height : 420,
    plugins: {
      legend: {
        position: 'bottom',
        labels: { boxWidth: 14, padding: 14, font: { size: 14, family: 'system-ui, sans-serif' } },
      },
      tooltip: {
        callbacks: {
          label: function(ctx) {
            const total = ctx.dataset.data.reduce(function(a, b) { return a + b; }, 0);
            const pct = total ? Math.round((ctx.raw / total) * 1000) / 10 : 0;
            return ctx.label + ': ' + ctx.raw + ' (' + pct + '%)';
          },
        },
      },
    },
  };
}

function barOptions_(horizontal, pctMode, canvas) {
  const idxAxis = horizontal ? 'y' : 'x';
  const valAxis = horizontal ? 'x' : 'y';
  return {
    responsive: false,
    animation: false,
    width: canvas ? canvas.width : 800,
    height: canvas ? canvas.height : 480,
    indexAxis: horizontal ? 'y' : 'x',
    plugins: { legend: { display: false } },
    scales: {
      [idxAxis]: {
        grid: { display: false },
        ticks: { font: { size: 12 }, autoSkip: false, maxRotation: horizontal ? 0 : 45, minRotation: 0 },
      },
      [valAxis]: {
        beginAtZero: true,
        suggestedMax: pctMode ? 100 : undefined,
        ticks: {
          stepSize: pctMode ? undefined : 1,
          precision: pctMode ? 1 : 0,
          font: { size: 12 },
          callback: pctMode ? function(v) { return v + '%'; } : undefined,
        },
        grid: { color: 'rgba(139, 26, 26, 0.08)' },
      },
    },
  };
}

function defaultEntityKey_(r) {
  const mill = String(r['MILL NAME'] || r['Mill Name'] || '').trim().toUpperCase();
  if (mill) return mill;
  const co = String(r['COMPANY NAME'] || r['Company Name'] || '').trim().toUpperCase();
  return co ? ('\u0000' + co) : '';
}

/**
 * Build unique-mill KPI snapshot from as-of rows (main and/or waste).
 * opts.qtyHelpers optional — when present, also sum supply columns.
 */
export function buildMillPeriodKpis_(rows, opts) {
  opts = opts || {};
  const entityKey = opts.entityKey || defaultEntityKey_;
  const resolveRisk = opts.resolveRisk || function(r) {
    return String(r['RESULT RISK LEVEL'] || r['RISK LEVEL'] || '').trim();
  };
  const isNbl = opts.isNbl || function(r) {
    return /yes|nbl|no buy/i.test(String(r['BUYER NO BUY LIST'] || ''));
  };
  const pickGroup = opts.pickGroup || function(r) {
    return String(r['GROUP NAME'] || '').trim();
  };
  const qty = opts.qtyHelpers || null;

  const seen = new Set();
  const entities = new Map();
  const groups = new Set();
  let highRisk = 0;
  let nbl = 0;
  const supply = { cpo: 0, pk: 0, pomeIscc: 0, pomeIns: 0, shell: 0 };

  (rows || []).forEach(function(r) {
    if (!r || !mrdShowInMillOnboarding_(r)) return;
    const ek = entityKey(r);
    if (!ek || ek === '\u0001') return;

    if (!seen.has(ek)) {
      seen.add(ek);
      entities.set(ek, r);
      const grp = pickGroup(r);
      if (grp) groups.add(grp);
      if (classifyRiskBucket_(resolveRisk(r)) === 'High') highRisk++;
      if (isNbl(r)) nbl++;
    }

    if (qty) {
      supply.cpo += qty.cpo(r) || 0;
      supply.pk += qty.pk(r) || 0;
      supply.pomeIscc += qty.pomeIscc(r) || 0;
      supply.pomeIns += qty.pomeIns(r) || 0;
      supply.shell += qty.shell(r) || 0;
    }
  });

  return {
    totalMills: entities.size,
    groupCount: groups.size,
    highRisk: highRisk,
    nbl: nbl,
    entityKeys: seen,
    supply: supply,
  };
}

export function buildQuarterComparison_(currentKpis, previousKpis, meta) {
  const cur = currentKpis || {};
  const prev = previousKpis || {};
  const curKeys = cur.entityKeys || new Set();
  const prevKeys = prev.entityKeys || new Set();
  let added = 0;
  let removed = 0;
  curKeys.forEach(function(k) { if (!prevKeys.has(k)) added++; });
  prevKeys.forEach(function(k) { if (!curKeys.has(k)) removed++; });

  return {
    currentLabel: meta && meta.currentLabel ? meta.currentLabel : 'Current',
    previousLabel: meta && meta.previousLabel ? meta.previousLabel : 'Previous',
    current: {
      totalMills: cur.totalMills || 0,
      groupCount: cur.groupCount || 0,
      highRisk: cur.highRisk || 0,
      nbl: cur.nbl || 0,
    },
    previous: {
      totalMills: prev.totalMills || 0,
      groupCount: prev.groupCount || 0,
      highRisk: prev.highRisk || 0,
      nbl: prev.nbl || 0,
    },
    delta: {
      totalMills: (cur.totalMills || 0) - (prev.totalMills || 0),
      groupCount: (cur.groupCount || 0) - (prev.groupCount || 0),
      highRisk: (cur.highRisk || 0) - (prev.highRisk || 0),
      nbl: (cur.nbl || 0) - (prev.nbl || 0),
      millsAdded: added,
      millsRemoved: removed,
    },
  };
}

export function buildMrdExecutiveData_(snapshot, extras) {
  extras = extras || {};
  const s = snapshot || {};
  const stats = s.stats || {};
  const mills = s.mills || [];
  const sdd = s.sdd || [];
  const grv = s.grv || [];

  const sddStatus = {
    Draft: stats.sddDraft || 0,
    Submitted: stats.sddSubmitted || 0,
  };

  const sddCategory = {};
  sdd.forEach(function(r) {
    const cat = normalizeSddCategory(
      r.supplier_type || r['Supplier Type'] || r['SUPPLIER_TYPE'] || ''
    ) || 'Unknown';
    sddCategory[cat] = (sddCategory[cat] || 0) + 1;
  });

  const riskBuckets = { High: 0, Medium: 0, Low: 0, Other: 0, Unclassified: 0 };
  mills.forEach(function(item) {
    const bucket = classifyRiskBucket_(item.risk);
    riskBuckets[bucket] = (riskBuckets[bucket] || 0) + 1;
  });

  const nblBuckets = {
    'No Buy List': stats.nblMills || 0,
    'Non-NBL': Math.max(0, (stats.totalMills || 0) - (stats.nblMills || 0)),
  };

  const traceability = {
    labels: ['TTM CPO', 'TTM PK', 'TTP CPO', 'TTP PK'],
    values: [
      parsePctNum_(stats.ttmCpoPct),
      parsePctNum_(stats.ttmPkPct),
      parsePctNum_(stats.ttpCpoPct),
      parsePctNum_(stats.ttpPkPct),
    ].map(function(v) { return v == null ? 0 : v; }),
  };

  const traceGap = {
    'With Supplier': Math.max(0, (stats.totalMills || 0) - (stats.emptyTraceMills || 0)),
    Untraceable: stats.emptyTraceMills || 0,
  };

  const grvStatus = {};
  grv.forEach(function(item) {
    const row = item.row || item;
    const st = String(
      row['Grievance Status'] || row['GRIEVANCE STATUS'] || row['Status'] || 'Unknown'
    ).trim() || 'Unknown';
    grvStatus[st] = (grvStatus[st] || 0) + 1;
  });

  const supply = extras.supply || { cpo: 0, pk: 0, pomeIscc: 0, pomeIns: 0, shell: 0 };
  const supplyBuckets = {
    CPO: Math.round((supply.cpo || 0) * 100) / 100,
    PK: Math.round((supply.pk || 0) * 100) / 100,
    'POME ISCC': Math.round((supply.pomeIscc || 0) * 100) / 100,
    'POME INS': Math.round((supply.pomeIns || 0) * 100) / 100,
    'SHELL GGL': Math.round((supply.shell || 0) * 100) / 100,
  };

  return {
    stats: stats,
    sddStatus: sddStatus,
    sddCategory: sddCategory,
    riskBuckets: riskBuckets,
    nblBuckets: nblBuckets,
    traceability: traceability,
    traceGap: traceGap,
    grvStatus: grvStatus,
    eudrPotential: stats.eudrPotential || 0,
    nblEntries: stats.nblEntries || 0,
    facilities: stats.facilities || 0,
    supply: supply,
    supplyBuckets: supplyBuckets,
    supplyPeriodLabel: extras.supplyPeriodLabel || '',
    quarterComparison: extras.quarterComparison || null,
    quarterlyTrend: extras.quarterlyTrend || [],
  };
}

function renderPie_(Chart, key, canvas, buckets, colorMap) {
  if (!Chart || !canvas) return;
  const filtered = Object.entries(buckets || {}).filter(function(e) { return e[1] > 0; });
  if (!filtered.length) return;
  chartInstances[key] = new Chart(canvas, {
    type: 'pie',
    data: {
      labels: filtered.map(function(e) { return e[0]; }),
      datasets: [{
        data: filtered.map(function(e) { return e[1]; }),
        backgroundColor: filtered.map(function(e) { return colorMap[e[0]] || '#8B7355'; }),
        borderWidth: 2,
        borderColor: '#fff',
      }],
    },
    options: pieOptions_(canvas),
    plugins: [PIE_PCT_PLUGIN],
  });
}

export function renderMrdExecutiveCharts_(Chart, data, els) {
  destroyMrdExecutiveCharts_();
  if (!Chart || !data || !els) return;

  renderPie_(Chart, 'sdd', els.sdd, data.sddStatus, CHART_COLORS);
  renderPie_(Chart, 'risk', els.risk, data.riskBuckets, CHART_COLORS);
  renderPie_(Chart, 'nbl', els.nbl, data.nblBuckets, CHART_COLORS);
  renderPie_(Chart, 'traceGap', els.traceGap, data.traceGap, CHART_COLORS);

  if (els.grv && Object.keys(data.grvStatus || {}).length) {
    const grvColors = { Open: '#C03030', Closed: '#2E7D32', 'In Progress': '#D4A017' };
    renderPie_(Chart, 'grv', els.grv, data.grvStatus, grvColors);
  }

  if (els.sddCat) {
    const top = topEntries_(data.sddCategory, 8);
    if (top.length) {
      chartInstances.sddCat = new Chart(els.sddCat, {
        type: 'bar',
        data: {
          labels: top.map(function(e) { return e[0]; }),
          datasets: [{
            label: 'Screenings',
            data: top.map(function(e) { return e[1]; }),
            backgroundColor: 'rgba(139, 26, 26, 0.78)',
            borderRadius: 6,
          }],
        },
        options: barOptions_(true, false, els.sddCat),
      });
    }
  }

  if (els.trace && data.traceability) {
    chartInstances.trace = new Chart(els.trace, {
      type: 'bar',
      data: {
        labels: data.traceability.labels,
        datasets: [{
          label: 'Traceability %',
          data: data.traceability.values,
          backgroundColor: [
            'rgba(139, 26, 26, 0.85)',
            'rgba(46, 125, 50, 0.85)',
            'rgba(230, 81, 0, 0.85)',
            'rgba(21, 101, 192, 0.85)',
          ],
          borderRadius: 6,
        }],
      },
      options: barOptions_(false, true, els.trace),
    });
  }

  if (els.qoq && data.quarterComparison) {
    const qc = data.quarterComparison;
    chartInstances.qoq = new Chart(els.qoq, {
      type: 'bar',
      data: {
        labels: ['Total Mills', 'Groups', 'High Risk', 'NBL'],
        datasets: [
          {
            label: qc.previousLabel,
            data: [
              qc.previous.totalMills,
              qc.previous.groupCount,
              qc.previous.highRisk,
              qc.previous.nbl,
            ],
            backgroundColor: 'rgba(156, 128, 128, 0.55)',
            borderRadius: 4,
          },
          {
            label: qc.currentLabel,
            data: [
              qc.current.totalMills,
              qc.current.groupCount,
              qc.current.highRisk,
              qc.current.nbl,
            ],
            backgroundColor: 'rgba(139, 26, 26, 0.85)',
            borderRadius: 4,
          },
        ],
      },
      options: {
        responsive: false,
        animation: false,
        width: els.qoq ? els.qoq.width : 800,
        height: els.qoq ? els.qoq.height : 380,
        plugins: {
          legend: {
            position: 'bottom',
            labels: { boxWidth: 12, padding: 12, font: { size: 13 } },
          },
        },
        scales: {
          x: { grid: { display: false }, ticks: { font: { size: 12 } } },
          y: {
            beginAtZero: true,
            ticks: { stepSize: 1, precision: 0, font: { size: 12 } },
            grid: { color: 'rgba(139, 26, 26, 0.08)' },
          },
        },
      },
    });
  }

  if (els.trend && data.quarterlyTrend && data.quarterlyTrend.length) {
    const qt = data.quarterlyTrend;
    chartInstances.trend = new Chart(els.trend, {
      type: 'bar',
      data: {
        labels: qt.map(function(x) { return x.label; }),
        datasets: [{
          label: 'Unique Mills',
          data: qt.map(function(x) { return x.totalMills; }),
          backgroundColor: qt.map(function(x) {
            return x.active ? 'rgba(139, 26, 26, 0.88)' : 'rgba(139, 26, 26, 0.38)';
          }),
          borderRadius: 6,
        }],
      },
      options: barOptions_(false, false, els.trend),
    });
  }

  if (els.supply && data.supplyBuckets) {
    const entries = Object.entries(data.supplyBuckets).filter(function(e) { return e[1] > 0; });
    if (entries.length) {
      chartInstances.supply = new Chart(els.supply, {
        type: 'bar',
        data: {
          labels: entries.map(function(e) { return e[0]; }),
          datasets: [{
            label: 'Quantity (ton)',
            data: entries.map(function(e) { return e[1]; }),
            backgroundColor: entries.map(function(e) { return CHART_COLORS[e[0]] || '#8B7355'; }),
            borderRadius: 6,
          }],
        },
        options: {
          responsive: false,
          animation: false,
          width: els.supply ? els.supply.width : 800,
          height: els.supply ? els.supply.height : 440,
          plugins: {
            legend: { display: false },
            tooltip: {
              callbacks: {
                label: function(ctx) {
                  return ctx.label + ': ' + formatQtyDisplay_(ctx.raw) + ' ton';
                },
              },
            },
          },
          scales: {
            x: { grid: { display: false }, ticks: { font: { size: 12 } } },
            y: {
              beginAtZero: true,
              ticks: { font: { size: 12 } },
              grid: { color: 'rgba(139, 26, 26, 0.08)' },
            },
          },
        },
      });
    }
  }
}

export function collectMrdExecutiveChartImages_(els) {
  const out = {};
  if (!els) return out;
  Object.keys(els).forEach(function(key) {
    const canvas = els[key];
    if (!canvas || typeof canvas.toDataURL !== 'function') return;
    try {
      out[key] = canvas.toDataURL('image/png', 1);
    } catch (_) { /* noop */ }
  });
  return out;
}

export function mrdExecutiveFilename_(year, month) {
  const names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const m = parseInt(String(month || ''), 10);
  const label = m >= 1 && m <= 12 ? names[m] + ' ' + year : 'Full Year ' + year;
  return 'Monthly Report Executive - ' + label + '.pdf';
}

export async function exportMrdExecutivePdf_(meta, data, chartImages, getJsPDF) {
  const JsPDF = typeof getJsPDF === 'function' ? getJsPDF() : null;
  if (!JsPDF) throw new Error('PDF library not loaded. Refresh the page and try again.');

  const doc = new JsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
  const pw = doc.internal.pageSize.getWidth();
  const ph = doc.internal.pageSize.getHeight();
  const margin = 14;
  const stats = data.stats || {};
  const supply = data.supply || {};
  const qc = data.quarterComparison;

  doc.setFillColor(139, 26, 26);
  doc.rect(0, 0, pw, 22, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFontSize(16);
  doc.text('Monthly Report — Executive Summary', margin, 14);
  doc.setFontSize(10);
  doc.text(meta.periodLine || '', margin, 19);

  doc.setTextColor(30, 10, 10);
  let y = 28;
  doc.setFontSize(9);
  doc.setTextColor(80, 60, 60);
  if (meta.dataPeriodLine) {
    doc.text(meta.dataPeriodLine, margin, y);
    y += 5;
  }
  if (meta.cutoffLine) {
    doc.text(meta.cutoffLine, margin, y);
    y += 7;
  }

  const kpis = [
    ['SDD Requested', stats.sddRequested != null ? stats.sddRequested : stats.sddTotal],
    ['Total Mills', stats.totalMills],
    ['High Risk', stats.highRisk],
    ['No Buy List', stats.nblMills],
    ['Untraceable', stats.emptyTraceMills],
    ['Grievances', stats.grievances],
    ['EUDR Potential', data.eudrPotential],
    ['NBL Entries', data.nblEntries],
  ];
  const kpiW = (pw - margin * 2 - 21) / 8;
  kpis.forEach(function(kpi, i) {
    const x = margin + i * (kpiW + 3);
    doc.setDrawColor(220, 200, 200);
    doc.setFillColor(252, 248, 248);
    doc.roundedRect(x, y, kpiW, 16, 2, 2, 'FD');
    doc.setFontSize(12);
    doc.setTextColor(139, 26, 26);
    doc.text(String(kpi[1] == null ? '—' : kpi[1]), x + 3, y + 9);
    doc.setFontSize(7);
    doc.setTextColor(100, 80, 80);
    doc.text(kpi[0], x + 3, y + 14);
  });
  y += 22;

  if (qc) {
    doc.setFontSize(10);
    doc.setTextColor(26, 10, 10);
    doc.text('Quarter comparison · ' + qc.previousLabel + ' → ' + qc.currentLabel, margin, y);
    y += 4;
    const deltas = [
      ['Mills', qc.current.totalMills, formatDelta_(qc.delta.totalMills)],
      ['Added', qc.delta.millsAdded, 'new'],
      ['Removed', qc.delta.millsRemoved, 'left'],
      ['Groups', qc.current.groupCount, formatDelta_(qc.delta.groupCount)],
      ['High Risk', qc.current.highRisk, formatDelta_(qc.delta.highRisk)],
      ['NBL', qc.current.nbl, formatDelta_(qc.delta.nbl)],
    ];
    const dW = (pw - margin * 2 - 15) / 6;
    deltas.forEach(function(d, i) {
      const x = margin + i * (dW + 3);
      doc.setDrawColor(220, 200, 200);
      doc.setFillColor(252, 248, 248);
      doc.roundedRect(x, y, dW, 16, 2, 2, 'FD');
      doc.setFontSize(11);
      doc.setTextColor(139, 26, 26);
      doc.text(String(d[1]), x + 3, y + 8);
      doc.setFontSize(7);
      doc.setTextColor(100, 80, 80);
      doc.text(d[0] + ' (' + d[2] + ')', x + 3, y + 13);
    });
    y += 22;
  }

  doc.setFontSize(10);
  doc.setTextColor(26, 10, 10);
  doc.text('Supply quantity · ' + (data.supplyPeriodLabel || 'selected period') + ' (ton)', margin, y);
  y += 4;
  const qtyKpis = [
    ['CPO', formatQtyDisplay_(supply.cpo)],
    ['PK', formatQtyDisplay_(supply.pk)],
    ['POME ISCC', formatQtyDisplay_(supply.pomeIscc)],
    ['POME INS', formatQtyDisplay_(supply.pomeIns)],
    ['SHELL GGL', formatQtyDisplay_(supply.shell)],
  ];
  const qW = (pw - margin * 2 - 12) / 5;
  qtyKpis.forEach(function(kpi, i) {
    const x = margin + i * (qW + 3);
    doc.setDrawColor(220, 200, 200);
    doc.setFillColor(252, 248, 248);
    doc.roundedRect(x, y, qW, 14, 2, 2, 'FD');
    doc.setFontSize(10);
    doc.setTextColor(139, 26, 26);
    doc.text(String(kpi[1]), x + 3, y + 7);
    doc.setFontSize(7);
    doc.setTextColor(100, 80, 80);
    doc.text(kpi[0], x + 3, y + 11.5);
  });
  y += 20;

  const imgs = chartImages || {};
  const layout = [
    { key: 'qoq', title: 'Quarter vs Previous', w: 88, h: 55 },
    { key: 'trend', title: 'Quarterly Mill Trend', w: 88, h: 55 },
    { key: 'supply', title: 'Supply Quantity (ton)', w: 88, h: 55 },
    { key: 'sdd', title: 'SDD Status', w: 88, h: 55 },
    { key: 'risk', title: 'Mill Risk Level', w: 88, h: 55 },
    { key: 'nbl', title: 'No Buy List', w: 88, h: 55 },
    { key: 'traceGap', title: 'Supplier Traceability Coverage', w: 88, h: 55 },
    { key: 'trace', title: 'Traceability % (TTM / TTP)', w: 88, h: 55 },
    { key: 'sddCat', title: 'SDD by Category', w: 88, h: 55 },
    { key: 'grv', title: 'Grievance Status', w: 88, h: 55 },
  ];
  let col = 0;
  let rowY = y;
  layout.forEach(function(item) {
    if (!imgs[item.key]) return;
    if (rowY + item.h + 14 > ph - 10) {
      doc.addPage();
      rowY = margin;
      col = 0;
    }
    const x = margin + col * (item.w + 6);
    doc.setFontSize(9);
    doc.setTextColor(26, 10, 10);
    doc.text(item.title, x, rowY);
    doc.addImage(imgs[item.key], 'PNG', x, rowY + 2, item.w, item.h);
    col++;
    if (col >= 2) {
      col = 0;
      rowY += item.h + 12;
    }
  });

  doc.setFontSize(8);
  doc.setTextColor(120, 100, 100);
  doc.text('Generated: ' + new Date().toLocaleString('en-GB', { hour12: false }), margin, ph - 8);
  doc.save(meta.filename);
}

export function mrdExecutiveHeaderMeta_(year, month) {
  const meta = mrdReportHeaderMeta_(year, month);
  return {
    periodLine: meta.periodLine,
    dataPeriodLine: meta.dataPeriodLine,
    cutoffLine: meta.cutoffLine,
  };
}

export { quarterEndMonth_, quarterMonthRangeLabel_, formatQtyDisplay_ };
