/**
 * Mill Onboarding — Executive Report (quarterly charts + PDF export).
 * Pure aggregation helpers; row sourcing/dedupe stays in main.js.
 */
import { mrdShowInMillOnboarding_ } from './monthly-report-labels.js';
import { registerPdfFonts_, setPdfFont_, PDF_FONT_SANS } from './pdf-fonts.js';
import { getMillExecutiveBackgroundDataUrl_ } from './mill-executive-bg.js';

/** Offscreen render size for row-2 horizontal bar charts (PDF-scaled; keep modest to avoid main-thread stalls). */
export const MILL_EXEC_HBAR_CHART_SIZE = { w: 520, h: 240 };


export function quarterEndMonth_(quarter) {
  const q = parseInt(String(quarter || ''), 10);
  if (q >= 1 && q <= 4) return q * 3;
  return 12;
}

export function quarterStartMonth_(quarter) {
  const q = parseInt(String(quarter || ''), 10);
  if (q >= 1 && q <= 4) return (q - 1) * 3 + 1;
  return 1;
}

export function quarterMonthRangeLabel_(quarter) {
  const map = { 1: 'Jan – Mar', 2: 'Apr – Jun', 3: 'Jul – Sep', 4: 'Oct – Dec' };
  return map[parseInt(String(quarter || ''), 10)] || '';
}

export function quarterAsOfLabel_(year, quarter) {
  const endMonth = quarterEndMonth_(quarter);
  const names = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return (names[endMonth - 1] || '') + ' ' + year;
}

export function millExecutivePeriodLabel_(year, quarter) {
  if (String(quarter) === 'full') return 'Full Year ' + year;
  return 'Q' + quarter + ' ' + year + ' (' + quarterMonthRangeLabel_(quarter) + ')';
}

export function millExecutiveFilename_(year, quarter) {
  if (String(quarter) === 'full') {
    return 'Mill Onboarding Executive Report - ' + year + ' Full Year.pdf';
  }
  return 'Mill Onboarding Executive Report - ' + year + ' Q' + quarter + '.pdf';
}

export function classifyRiskBucket_(riskStr) {
  const s = String(riskStr || '').trim().toLowerCase();
  if (!s) return 'Unclassified';
  if (s.includes('high')) return 'High';
  if (s.includes('medium') || s.includes('med')) return 'Medium';
  if (s.includes('low')) return 'Low';
  return 'Other';
}

export function filterMillExecutiveRows_(rows) {
  return (rows || []).filter(function(r) {
    return r && mrdShowInMillOnboarding_(r);
  });
}

export function isExcludedSupplierGroup_(grp) {
  const s = String(grp || '').trim();
  if (!s || s === '—' || s === '-') return true;
  const u = s.toUpperCase();
  if (u === 'UNKNOWN' || u === 'NONE' || u === 'NONE GROUP' || u === 'N/A' || u === 'NA') return true;
  if (/^NONE(\s|$)/.test(u)) return true;
  return false;
}

export function aggregateMillExecutiveSnapshot_(rows, opts) {
  opts = opts || {};
  const entityKey = opts.entityKey || function(r) {
    return String(r['MILL NAME'] || r['COMPANY NAME'] || '').trim();
  };
  const resolveRisk = opts.resolveRisk || function(r) {
    return String(r['RESULT RISK LEVEL'] || r['RISK LEVEL'] || '').trim();
  };
  const isNbl = opts.isNbl || function() { return false; };
  const pickGroup = opts.pickGroup || function(r) {
    return String(r['GROUP NAME'] || '').trim();
  };
  const pickProvince = opts.pickProvince || function(r) {
    const p = String(r['PROVINCE'] || '').trim();
    return p && p !== '—' && p !== '-' ? p : 'Unknown';
  };
  const pickQty = opts.pickQty || function(r) {
    function n_(keys) {
      for (let i = 0; i < keys.length; i++) {
        const v = parseFloat(String(r[keys[i]] || '').replace(/[, ]/g, ''));
        if (!isNaN(v) && v > 0) return v;
      }
      return 0;
    }
    return n_(['SUPPLY CPO', 'Supply CPO', 'SUPPLY_CPO']) +
           n_(['SUPPLY PK', 'Supply PK', 'SUPPLY_PK']);
  };

  const seen = new Set();
  const snapshot = {
    totalMills: 0,
    groups: new Set(),
    highRisk: 0,
    nbl: 0,
    riskBuckets: { High: 0, Medium: 0, Low: 0, Other: 0, Unclassified: 0 },
    nblBuckets: { 'No Buy List': 0, 'Non-NBL': 0 },
    provinces: {},
    facilityQty: {},    // supplier group → total supply ton (top suppliers chart)
    traceability: { Traceable: 0, Untraceable: 0 }, // has qty > 0 or not
    certification: {},
    rawRowCount: (rows || []).length,
    entityCount: 0,
  };

  filterMillExecutiveRows_(rows).forEach(function(r) {
    const ek = entityKey(r);
    if (!ek || ek === '\u0001' || seen.has(ek)) return;
    seen.add(ek);
    snapshot.entityCount++;
    snapshot.totalMills++;

    const grp = pickGroup(r);
    if (grp) snapshot.groups.add(grp);

    const bucket = classifyRiskBucket_(resolveRisk(r));
    snapshot.riskBuckets[bucket] = (snapshot.riskBuckets[bucket] || 0) + 1;
    if (bucket === 'High') snapshot.highRisk++;

    if (isNbl(r)) {
      snapshot.nbl++;
      snapshot.nblBuckets['No Buy List']++;
    } else {
      snapshot.nblBuckets['Non-NBL']++;
    }

    const prov = pickProvince(r);
    snapshot.provinces[prov] = (snapshot.provinces[prov] || 0) + 1;

    const qty = pickQty(r);
    const useQuarterSupplyRows = opts.supplyRows && opts.supplyRows.length;
    if (!useQuarterSupplyRows) {
      if (grp && !isExcludedSupplierGroup_(grp)) {
        snapshot.facilityQty[grp] = (snapshot.facilityQty[grp] || 0) + qty;
      }
    }

    if (qty > 0) {
      snapshot.traceability['Traceable']++;
    } else {
      snapshot.traceability['Untraceable']++;
    }

    const certRaw = String(r['CERTIFICATION'] || '').trim();
    if (certRaw && certRaw !== '—' && certRaw !== '-') {
      certRaw.split(/[,;|\n\r]+/).map(function(p) { return p.trim(); }).filter(Boolean).forEach(function(c) {
        snapshot.certification[c] = (snapshot.certification[c] || 0) + 1;
      });
    } else {
      snapshot.certification['None'] = (snapshot.certification['None'] || 0) + 1;
    }
  });

  snapshot.groupCount = snapshot.groups.size;

  if (opts.supplyRows && opts.supplyRows.length) {
    snapshot.facilityQty = {};
    filterMillExecutiveRows_(opts.supplyRows).forEach(function(r) {
      const grp = pickGroup(r);
      const qty = pickQty(r);
      if (grp && !isExcludedSupplierGroup_(grp)) {
        snapshot.facilityQty[grp] = (snapshot.facilityQty[grp] || 0) + qty;
      }
    });
  }

  return snapshot;
}

export function topEntries_(obj, limit) {
  return Object.entries(obj || {})
    .sort(function(a, b) { return b[1] - a[1] || a[0].localeCompare(b[0]); })
    .slice(0, limit || 8);
}

const CHART_COLORS = {
  High: '#DC2626',
  Medium: '#F59E0B',
  Low: '#38BDF8',
  Other: '#94A3B8',
  Unclassified: '#CBD5E1',
  'No Buy List': '#EA580C',
  'Non-NBL': '#32B0EB',
  Closed: '#0EA5E9',
  Open: '#DC2626',
  Invalid: '#94A3B8',
};

const KPI_STYLES = {
  totalMills: { num: [14, 165, 233] },
  groups:     { num: [29, 78, 216] },
  highRisk:   { num: [220, 38, 38] },
  noBuyList:  { num: [234, 88, 12] },
};

function yieldToBrowser_(ms) {
  const delay = ms == null ? 12 : ms;
  return new Promise(function(resolve) {
    requestAnimationFrame(function() {
      setTimeout(resolve, delay);
    });
  });
}

let chartInstances = {};

export function destroyMillExecutiveCharts_() {
  Object.keys(chartInstances).forEach(function(k) {
    try { chartInstances[k].destroy(); } catch (_) { /* noop */ }
  });
  chartInstances = {};
}

// Inline plugin: draws percentage labels on pie slices
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
        if (pct < 4) return; // skip tiny slices
        const pos = arc.tooltipPosition();
        const fontSize = Math.max(14, Math.round(Math.min(chart.width, chart.height) * 0.052));
        ctx.save();
        ctx.font = 'bold ' + fontSize + 'px ' + PDF_FONT_SANS;
        ctx.fillStyle = '#ffffff';
        ctx.shadowColor = 'rgba(0,0,0,0.55)';
        ctx.shadowBlur = 4;
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(pct + '%', pos.x, pos.y);
        ctx.restore();
      });
    });
  },
};

function pieOptions_(canvas) {
  const sz = canvas ? Math.min(canvas.width, canvas.height) : 500;
  const legendFs = Math.round(sz * 0.036);
  const boxW = Math.round(sz * 0.026);
  return {
    responsive: false,
    animation: false,
    backgroundColor: 'transparent',
    width: canvas ? canvas.width : 500,
    height: canvas ? canvas.height : 500,
    layout: { padding: 20 },
    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          boxWidth: boxW,
          padding: Math.round(sz * 0.022),
          font: { size: legendFs, family: PDF_FONT_SANS },
        },
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

function truncLabel_(s, max) {
  max = max || 18;
  s = String(s || '');
  return s.length > max ? s.substring(0, max - 1) + '…' : s;
}

function barOptions_(horizontal, canvas, maxLabelLen) {
  const idxAxis = horizontal ? 'y' : 'x';
  const valAxis = horizontal ? 'x' : 'y';
  const w = canvas ? canvas.width : 800;
  const h = canvas ? canvas.height : 500;
  const fs = Math.round(h * 0.036);
  const maxLen = maxLabelLen || (horizontal ? 26 : 18);
  const layoutPad = horizontal
    ? { top: 6, right: 8, bottom: 6, left: 6 }
    : { top: 10, right: 20, bottom: 10, left: 10 };
  const valScale = {
    beginAtZero: true,
    ticks: { precision: 0, font: { size: fs, family: PDF_FONT_SANS } },
    grid: { color: 'rgba(139,26,26,0.07)' },
  };
  if (horizontal) valScale.grace = 0;
  return {
    responsive: false,
    animation: false,
    backgroundColor: 'transparent',
    width: w,
    height: h,
    layout: { padding: layoutPad },
    indexAxis: horizontal ? 'y' : 'x',
    plugins: { legend: { display: false } },
    scales: {
      [idxAxis]: {
        grid: { display: false },
        ticks: {
          font: { size: fs, family: PDF_FONT_SANS },
          autoSkip: false,
          maxRotation: horizontal ? 0 : 0,
          minRotation: 0,
          callback: function(val) {
            const lbl = this.getLabelForValue ? this.getLabelForValue(val) : String(val);
            return truncLabel_(lbl, maxLen);
          },
        },
      },
      [valAxis]: valScale,
    },
  };
}

/** Grouped / stacked bar charts for full-year quarter comparison (small PDF panels). */
function quarterCompareBarOptions_(canvas, opts) {
  opts = opts || {};
  const w = canvas ? canvas.width : 320;
  const h = canvas ? canvas.height : 320;
  const fs = Math.round(Math.min(w, h) * 0.028);
  const stacked = opts.stacked === true;
  const yScale = {
    stacked: stacked,
    beginAtZero: true,
    grid: { color: 'rgba(139, 26, 26, 0.08)' },
    ticks: {
      font: { size: Math.max(8, fs - 1), family: PDF_FONT_SANS },
      precision: opts.yPct ? 0 : undefined,
      callback: opts.yPct
        ? function(v) { return v + '%'; }
        : (opts.yTon
          ? function(v) {
            const n = Number(v);
            if (isNaN(n)) return v;
            if (n >= 1000) return (Math.round(n / 100) / 10) + 'k';
            return n;
          }
          : undefined),
    },
  };
  if (opts.yMax != null) yScale.max = opts.yMax;

  return {
    responsive: false,
    animation: false,
    width: w,
    height: h,
    layout: { padding: { top: 4, right: 6, bottom: 2, left: 4 } },
    plugins: {
      legend: {
        display: opts.legend !== false,
        position: 'bottom',
        labels: {
          boxWidth: 10,
          padding: opts.legendCompact ? 4 : 8,
          font: { size: Math.max(7, fs - 1), family: PDF_FONT_SANS },
        },
      },
    },
    scales: {
      x: {
        stacked: stacked,
        grid: { display: false },
        ticks: { font: { size: Math.max(9, fs), family: PDF_FONT_SANS } },
      },
      y: yScale,
    },
  };
}

function quarterCompareGroupedDatasets_(keys, seriesMap, palette) {
  return keys.map(function(name, i) {
    const c = palette[i % palette.length];
    return {
      label: truncLabel_(name, 16),
      data: seriesMap[name],
      backgroundColor: c + (c.length === 7 ? 'D9' : ''),
      borderRadius: 3,
      borderSkipped: false,
    };
  });
}

export function renderMillExecutiveCharts_(Chart, snapshot, els) {
  destroyMillExecutiveCharts_();
  if (!Chart || !snapshot || !els) return;
  renderMillExecutiveChartsBody_(Chart, snapshot, els);
}

function renderMillExecutiveChartsBody_(Chart, snapshot, els) {
  if (els.risk) {
    const filtered = ['High', 'Medium', 'Low'].map(function(l) {
      return { label: l, value: snapshot.riskBuckets[l] || 0, color: CHART_COLORS[l] };
    }).filter(function(x) { return x.value > 0; });
    chartInstances.risk = new Chart(els.risk, {
      type: 'pie',
      data: {
        labels: filtered.map(function(x) { return x.label; }),
        datasets: [{
          data: filtered.map(function(x) { return x.value; }),
          backgroundColor: filtered.map(function(x) { return x.color; }),
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: pieOptions_(els.risk),
      plugins: [PIE_PCT_PLUGIN],
    });
  }

  if (els.nbl) {
    const filtered = Object.entries(snapshot.nblBuckets).filter(function(e) { return e[1] > 0; });
    chartInstances.nbl = new Chart(els.nbl, {
      type: 'pie',
      data: {
        labels: filtered.map(function(e) { return e[0]; }),
        datasets: [{
          data: filtered.map(function(e) { return e[1]; }),
          backgroundColor: filtered.map(function(e) { return CHART_COLORS[e[0]] || '#8B7355'; }),
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: pieOptions_(els.nbl),
      plugins: [PIE_PCT_PLUGIN],
    });
  }

  // Province (horizontal bar, top 10 — easier to read truncated names)
  if (els.province) {
    const top = topEntries_(snapshot.provinces, 10);
    chartInstances.province = new Chart(els.province, {
      type: 'bar',
      data: {
        labels: top.map(function(e) { return e[0]; }),
        datasets: [{
          label: 'Mills',
          data: top.map(function(e) { return e[1]; }),
          backgroundColor: 'rgba(96,165,250,0.85)',
          borderRadius: 6,
        }],
      },
      options: barOptions_(true, els.province, 26),
    });
  }

  // Top suppliers by supply quantity (ton), horizontal bar
  if (els.facilityQty) {
    const palette = ['#991B1B','#B45309','#1D4ED8','#6366F1','#6D28D9','#9D174D','#0369A1','#7C3AED'];
    const top = topEntries_(snapshot.facilityQty, 8).filter(function(e) { return e[1] > 0; });
    const barOpts = barOptions_(true, els.facilityQty, 26);
    barOpts.scales.x.ticks.callback = function(v) {
      const n = Number(v);
      if (isNaN(n)) return v;
      if (n >= 1000) return (Math.round(n / 100) / 10) + 'k';
      return n;
    };
    chartInstances.facilityQty = new Chart(els.facilityQty, {
      type: 'bar',
      data: {
        labels: top.map(function(e) { return e[0]; }),
        datasets: [{
          label: 'Supply qty (ton)',
          data: top.map(function(e) { return e[1]; }),
          backgroundColor: top.map(function(_, i) { return palette[i % palette.length] + 'CC'; }),
          borderRadius: 6,
        }],
      },
      options: barOpts,
    });
  }

  // TTP supply traceability (weighted % from TTP data, not mill row counts)
  if (els.traceability) {
    const ttpPct = snapshot.ttpTrace && !isNaN(snapshot.ttpTrace.pct)
      ? Math.min(100, Math.max(0, snapshot.ttpTrace.pct))
      : null;
    let traceable;
    let untraceable;
    if (ttpPct != null) {
      traceable = ttpPct;
      untraceable = Math.max(0, 100 - ttpPct);
    } else {
      const t = snapshot.traceability || {};
      traceable = t.Traceable || 0;
      untraceable = t.Untraceable || 0;
    }
    const traceTotal = traceable + untraceable;
    const donutOpts = Object.assign({}, pieOptions_(els.traceability), {
      cutout: '58%',
    });
    donutOpts.plugins = donutOpts.plugins || {};
    donutOpts.plugins.legend = {
      position: 'bottom',
      labels: {
        boxWidth: Math.round(els.traceability.width * 0.028),
        padding: Math.round(els.traceability.width * 0.024),
        font: { size: Math.round(els.traceability.height * 0.034), family: PDF_FONT_SANS },
        generateLabels: function(chart) {
          return chart.data.labels.map(function(lbl, i) {
            const val = chart.data.datasets[0].data[i];
            const pct = traceTotal ? Math.round((val / traceTotal) * 100) : 0;
            return {
              text: lbl + ': ' + pct + '%',
              fillStyle: chart.data.datasets[0].backgroundColor[i],
              hidden: false,
              index: i,
            };
          });
        },
      },
    };
    chartInstances.traceability = new Chart(els.traceability, {
      type: 'doughnut',
      data: {
        labels: ['Traceable', 'Untraceable'],
        datasets: [{
          data: [traceable, untraceable],
          backgroundColor: ['#0EA5E9', '#57534E'],
          borderWidth: 3,
          borderColor: '#ffffff',
        }],
      },
      options: donutOpts,
      plugins: [PIE_PCT_PLUGIN],
    });
  }

  if (els.grievance) {
    const grv = snapshot.grievanceProgress || {};
    const buckets = snapshot.grievanceBuckets || {
      Closed: grv.closed || 0,
      Open: grv.open || 0,
      Invalid: grv.invalid || 0,
    };
    let filtered = [
      { label: 'Closed', value: buckets.Closed || 0, color: CHART_COLORS.Closed },
      { label: 'Open', value: buckets.Open || 0, color: CHART_COLORS.Open },
      { label: 'Invalid', value: buckets.Invalid || 0, color: CHART_COLORS.Invalid },
    ].filter(function(x) { return x.value > 0; });
    if (!filtered.length) {
      filtered = [{ label: 'No cases', value: 1, color: '#E2E8F0' }];
    }
    const grvCenterPlugin = {
      id: 'grvCenterNote',
      afterDraw: function(chart) {
        drawGrvCenterNote_(chart, grv);
      },
    };
    chartInstances.grievance = new Chart(els.grievance, {
      type: 'doughnut',
      data: {
        labels: filtered.map(function(x) { return x.label; }),
        datasets: [{
          data: filtered.map(function(x) { return x.value; }),
          backgroundColor: filtered.map(function(x) { return x.color; }),
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: Object.assign({}, pieOptions_(els.grievance), {
        cutout: '58%',
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              font: { size: Math.round(els.grievance.height * 0.034), family: PDF_FONT_SANS },
            },
          },
          title: {
            display: false,
          },
        },
      }),
      plugins: [PIE_PCT_PLUGIN, grvCenterPlugin],
    });
  }
}

/** Center label inside grievance donut (Date Received calendar year). */
function drawGrvCenterNote_(chart, grv) {
  grv = grv || {};
  const year = grv.year || '';
  const ctx = chart.ctx;
  const cx = chart.width / 2;
  const cy = chart.height * 0.505;
  const fsMain = Math.max(11, Math.round(chart.height * 0.032));
  const fsSub = Math.max(9, Math.round(chart.height * 0.024));
  ctx.save();
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';

  if (grv.total > 0) {
    const pct = grv.pct != null ? grv.pct : 0;
    const line1 = grv.closed + '/' + grv.total + ' closed · ' + pct + '% resolved';
    ctx.fillStyle = '#334155';
    ctx.font = '600 ' + fsMain + 'px ' + PDF_FONT_SANS;
    ctx.fillText(line1, cx, cy - fsSub * 0.55);
    ctx.fillStyle = '#64748B';
    ctx.font = '500 ' + fsSub + 'px ' + PDF_FONT_SANS;
    ctx.fillText('Date Received · ' + year, cx, cy + fsMain * 0.55);
  } else {
    ctx.fillStyle = '#64748B';
    ctx.font = '600 ' + fsMain + 'px ' + PDF_FONT_SANS;
    ctx.fillText('No grievances', cx, cy - fsSub * 0.4);
    ctx.font = '500 ' + fsSub + 'px ' + PDF_FONT_SANS;
    ctx.fillText('Date Received · ' + year, cx, cy + fsMain * 0.45);
  }
  ctx.restore();
}

/** Renders charts one-by-one so the main thread can breathe (avoids “Page Unresponsive”). */
export async function renderMillExecutiveChartsAsync_(Chart, snapshot, els) {
  destroyMillExecutiveCharts_();
  if (!Chart || !snapshot || !els) return;

  const order = [
    'risk', 'nbl', 'traceability', 'grievance', 'province', 'facilityQty',
  ];
  for (let i = 0; i < order.length; i++) {
    const key = order[i];
    const partial = {};
    partial[key] = els[key];
    renderMillExecutiveChartsBody_(Chart, snapshot, partial);
    await yieldToBrowser_(32);
  }
}

function lineChartOptions_(canvas, yAxisPct) {
  const w = canvas ? canvas.width : 760;
  const h = canvas ? canvas.height : 380;
  const opts = {
    responsive: false,
    animation: false,
    width: w,
    height: h,
    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          boxWidth: 14,
          padding: h <= 320 ? 8 : 14,
          font: { size: Math.round(h * (h <= 320 ? 0.028 : 0.032)), family: PDF_FONT_SANS },
        },
      },
    },
    scales: {
      x: {
        grid: { display: false },
        ticks: { font: { size: h <= 320 ? 10 : 12, family: PDF_FONT_SANS } },
      },
      y: {
        beginAtZero: true,
        ticks: {
          font: { size: h <= 320 ? 9 : 11, family: PDF_FONT_SANS },
          precision: 0,
          callback: yAxisPct
            ? function(v) { return v + '%'; }
            : undefined,
        },
        grid: { color: 'rgba(139, 26, 26, 0.08)' },
      },
    },
  };
  if (yAxisPct) opts.scales.y.max = 100;
  return opts;
}

function millExecTtpPctFromSnapshot_(snapshot) {
  const ttpPct = snapshot.ttpTrace && !isNaN(snapshot.ttpTrace.pct)
    ? Math.min(100, Math.max(0, snapshot.ttpTrace.pct))
    : null;
  if (ttpPct != null) return Math.round(ttpPct * 10) / 10;
  const t = snapshot.traceability || {};
  const traceTotal = (t.Traceable || 0) + (t.Untraceable || 0);
  if (!traceTotal) return null;
  return Math.round(((t.Traceable || 0) / traceTotal) * 1000) / 10;
}

function millExecGrvPctFromSnapshot_(snapshot) {
  const grv = snapshot.grievanceProgress || {};
  if (!grv.total) return grv.year ? 0 : null;
  return grv.pct != null ? grv.pct : Math.round((grv.closed / grv.total) * 100);
}

/**
 * Build series for full-year comparison charts (same metrics as quarterly executive).
 * @param {Array<{ quarter, label, snapshot }>} quarterRows — ordered Q1…Q4 with data
 */
export function buildMillExecutiveQuarterlyComparison_(quarterRows) {
  const rows = (quarterRows || []).filter(function(r) { return r && r.snapshot; });
  const labels = rows.map(function(r) { return r.label || ('Q' + r.quarter); });

  function pickTopKeys_(key, limit) {
    const totals = {};
    rows.forEach(function(r) {
      const obj = r.snapshot[key] || {};
      Object.keys(obj).forEach(function(k) {
        totals[k] = (totals[k] || 0) + (obj[k] || 0);
      });
    });
    return topEntries_(totals, limit || 4).map(function(e) { return e[0]; });
  }

  const kpis = { totalMills: [], groupCount: [], highRisk: [], nbl: [] };
  const riskHigh = [];
  const riskMedium = [];
  const riskLow = [];
  const nblCount = [];
  const nblPct = [];
  const nonNblCount = [];
  const ttpPct = [];
  const grvClosedPct = [];

  rows.forEach(function(r) {
    const s = r.snapshot;
    const total = s.totalMills || 0;
    kpis.totalMills.push(total);
    kpis.groupCount.push(s.groupCount != null ? s.groupCount : s.groups.size);
    kpis.highRisk.push(s.highRisk || 0);
    kpis.nbl.push(s.nbl || 0);
    riskHigh.push(s.riskBuckets.High || 0);
    riskMedium.push(s.riskBuckets.Medium || 0);
    riskLow.push(s.riskBuckets.Low || 0);
    nblCount.push(s.nbl || 0);
    nonNblCount.push(Math.max(0, total - (s.nbl || 0)));
    nblPct.push(total ? Math.round(((s.nbl || 0) / total) * 100) : 0);
    const tp = millExecTtpPctFromSnapshot_(s);
    ttpPct.push(tp != null ? tp : null);
    const gp = millExecGrvPctFromSnapshot_(s);
    grvClosedPct.push(gp != null ? gp : null);
  });

  const provinceKeys = pickTopKeys_('provinces');
  const supplierKeys = pickTopKeys_('facilityQty');
  const provinceSeries = {};
  const supplierSeries = {};
  provinceKeys.forEach(function(name) {
    provinceSeries[name] = rows.map(function(r) {
      return (r.snapshot.provinces && r.snapshot.provinces[name]) || 0;
    });
  });
  supplierKeys.forEach(function(name) {
    supplierSeries[name] = rows.map(function(r) {
      return (r.snapshot.facilityQty && r.snapshot.facilityQty[name]) || 0;
    });
  });

  return {
    labels: labels,
    quarters: rows.map(function(r) { return r.quarter; }),
    kpis: kpis,
    riskHigh: riskHigh,
    riskMedium: riskMedium,
    riskLow: riskLow,
    nblCount: nblCount,
    nonNblCount: nonNblCount,
    nblPct: nblPct,
    ttpPct: ttpPct,
    grvClosedPct: grvClosedPct,
    provinceSeries: provinceSeries,
    supplierSeries: supplierSeries,
  };
}

function quarterCompareLineOptions_(canvas, opts) {
  opts = opts || {};
  const o = lineChartOptions_(canvas, opts.yPct);
  if (canvas && canvas.height <= 320) {
    o.plugins.legend.labels.font.size = Math.max(7, Math.round(canvas.height * 0.026));
    o.plugins.legend.labels.padding = 6;
    o.plugins.legend.labels.boxWidth = 10;
  }
  if (opts.yPct && o.scales && o.scales.y) {
    // Headroom so 100% points / markers are not clipped at the chart top
    o.scales.y.max = 110;
    o.scales.y.grace = '8%';
    const prevTickCb = o.scales.y.ticks.callback;
    o.scales.y.ticks.callback = function(v) {
      const n = Number(v);
      if (!isNaN(n) && n > 100) return '';
      if (prevTickCb) return prevTickCb.call(this, v);
      return v + '%';
    };
  }
  o.layout = Object.assign({ padding: { top: 14, right: 8, bottom: 2, left: 4 } }, o.layout || {});
  return o;
}

function renderMillExecutiveFullYearComparisonBody_(Chart, bundle, els) {
  if (!Chart || !bundle || !els) return;
  const labels = bundle.labels || [];
  const barPalette = ['#2563EB', '#C03030', '#E67E22', '#7C3AED', '#0D9488', '#B45309'];
  const nblQuarterColors = ['#C2410C', '#EA580C', '#F97316', '#FB923C'];

  // 1) Risk — stacked bar (composition per quarter)
  if (els.risk && labels.length) {
    chartInstances.risk = new Chart(els.risk, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: 'High', data: bundle.riskHigh, backgroundColor: CHART_COLORS.High, stack: 'risk', borderRadius: 2 },
          { label: 'Medium', data: bundle.riskMedium, backgroundColor: CHART_COLORS.Medium, stack: 'risk', borderRadius: 2 },
          { label: 'Low', data: bundle.riskLow, backgroundColor: CHART_COLORS.Low, stack: 'risk', borderRadius: 2 },
        ],
      },
      options: quarterCompareBarOptions_(els.risk, { stacked: true, legendCompact: true }),
    });
  }

  // 2) NBL — pie (share of NBL mills across quarters)
  if (els.nbl && labels.length) {
    const nblData = (bundle.nblCount || []).map(function(v) { return v || 0; });
    const hasNbl = nblData.some(function(v) { return v > 0; });
    chartInstances.nbl = new Chart(els.nbl, {
      type: 'pie',
      data: {
        labels: labels,
        datasets: [{
          data: hasNbl ? nblData : [1],
          backgroundColor: hasNbl
            ? nblQuarterColors.slice(0, labels.length)
            : ['#E2E8F0'],
          borderWidth: 2,
          borderColor: '#fff',
        }],
      },
      options: Object.assign({}, pieOptions_(els.nbl), {
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: Math.round(els.nbl.width * 0.022),
              padding: Math.round(els.nbl.height * 0.018),
              font: { size: Math.round(els.nbl.height * 0.032), family: PDF_FONT_SANS },
            },
          },
        },
      }),
      plugins: hasNbl ? [PIE_PCT_PLUGIN] : [],
    });
  }

  // 3) TTP — area chart (trend % traceable)
  if (els.traceability && labels.length) {
    chartInstances.traceability = new Chart(els.traceability, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: '% traceable',
          data: bundle.ttpPct,
          borderColor: '#2563EB',
          backgroundColor: 'rgba(37, 99, 235, 0.32)',
          borderWidth: 2.5,
          pointRadius: 4,
          pointBackgroundColor: '#2563EB',
          tension: 0.35,
          fill: true,
          clip: false,
        }],
      },
      options: quarterCompareLineOptions_(els.traceability, { yPct: true }),
    });
  }

  // 4) Grievance — line (% closed by quarter)
  if (els.grievance && labels.length) {
    chartInstances.grievance = new Chart(els.grievance, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: '% closed',
          data: bundle.grvClosedPct,
          borderColor: '#0EA5E9',
          backgroundColor: 'rgba(14, 165, 233, 0.15)',
          borderWidth: 2.5,
          pointRadius: 4,
          pointHoverRadius: 4,
          tension: 0.2,
          fill: false,
          clip: false,
        }],
      },
      options: quarterCompareLineOptions_(els.grievance, { yPct: true }),
    });
  }

  // 5) Province — horizontal grouped bar (top 3 provinces)
  if (els.province && labels.length) {
    const provKeys = Object.keys(bundle.provinceSeries || {}).slice(0, 3);
    const hOpts = barOptions_(true, els.province, 12);
    hOpts.plugins.legend = {
      display: true,
      position: 'bottom',
      labels: {
        boxWidth: 10,
        padding: 6,
        font: { size: Math.round(els.province.height * 0.032), family: PDF_FONT_SANS },
      },
    };
    chartInstances.province = new Chart(els.province, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: quarterCompareGroupedDatasets_(provKeys, bundle.provinceSeries, barPalette),
      },
      options: hOpts,
    });
  }

  // 6) Top supplier — vertical grouped bar
  if (els.facilityQty && labels.length) {
    const supKeys = Object.keys(bundle.supplierSeries || {});
    chartInstances.facilityQty = new Chart(els.facilityQty, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: quarterCompareGroupedDatasets_(supKeys, bundle.supplierSeries, barPalette),
      },
      options: quarterCompareBarOptions_(els.facilityQty, { yTon: true, legendCompact: true }),
    });
  }
}

/** Full-year — six slots; mixed chart types (bar / pie / area / line). */
export async function renderMillExecutiveFullYearComparisonChartsAsync_(Chart, bundle, els) {
  destroyMillExecutiveCharts_();
  if (!Chart || !bundle || !els) return;
  const order = ['risk', 'nbl', 'traceability', 'grievance', 'province', 'facilityQty'];
  for (let i = 0; i < order.length; i++) {
    const key = order[i];
    const partial = {};
    partial[key] = els[key];
    renderMillExecutiveFullYearComparisonBody_(Chart, bundle, partial);
    await yieldToBrowser_(32);
  }
}

export function buildMillExecutiveFullYearInsights_(bundle, meta) {
  const insights = [];
  const labels = bundle.labels || [];
  const k = bundle.kpis || {};
  if (labels.length && k.totalMills && k.totalMills.length) {
    const a = k.totalMills[0];
    const b = k.totalMills[k.totalMills.length - 1];
    insights.push({
      label: 'Mills',
      value: labels[0] + ' ' + a + ' → ' + labels[labels.length - 1] + ' ' + b
        + (b !== a ? (' (' + (b >= a ? '+' : '') + (b - a) + ')') : ''),
      tone: 'neutral',
    });
  }
  if (bundle.ttpPct && bundle.ttpPct.filter(function(v) { return v != null; }).length) {
    const vals = bundle.ttpPct.filter(function(v) { return v != null; });
    insights.push({
      label: 'TTP',
      value: 'Traceable ' + vals[0] + '% → ' + vals[vals.length - 1] + '% by quarter',
      tone: 'neutral',
    });
  }
  if (k.highRisk && k.highRisk.length) {
    insights.push({
      label: 'High risk',
      value: labels[0] + ' ' + k.highRisk[0] + ' → ' + labels[labels.length - 1] + ' ' + k.highRisk[k.highRisk.length - 1] + ' mills',
      tone: 'neutral',
    });
  }
  if (k.nbl && k.nbl.length) {
    insights.push({
      label: 'NBL',
      value: labels[0] + ' ' + k.nbl[0] + ' → ' + labels[labels.length - 1] + ' ' + k.nbl[k.nbl.length - 1] + ' mills on NBL',
      tone: 'neutral',
    });
  }
  const provNames = Object.keys(bundle.provinceSeries || {});
  if (provNames.length) {
    const top = provNames[0];
    const ser = bundle.provinceSeries[top];
    insights.push({
      label: 'Top province',
      value: truncLabel_(top, 22) + ' · ' + ser[0] + ' → ' + ser[ser.length - 1] + ' mills',
      tone: 'neutral',
    });
  }
  insights.push({
    label: 'Scope',
    value: 'Full year ' + (meta.year || '') + ' · ' + labels.join(', ') + ' compared',
    tone: 'neutral',
  });
  return insights.slice(0, 6);
}

/** @deprecated — use full-year comparison grid instead */
export function renderMillExecutiveQuarterlyLineChart_(Chart, canvas, trend) {
  if (!Chart || !canvas || !trend || !trend.length) return;
  const labels = trend.map(function(t) { return t.label; });
  function series_(key) {
    return trend.map(function(t) {
      if (t.hasData === false) return null;
      const v = t[key];
      return v == null ? null : v;
    });
  }
  const lineDs = function(label, key, color, bg) {
    return {
      label: label,
      data: series_(key),
      borderColor: color,
      backgroundColor: bg,
      borderWidth: 2.5,
      pointRadius: 5,
      tension: 0.25,
      fill: false,
      spanGaps: false,
    };
  };
  chartInstances.quarterlyLine = new Chart(canvas, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        lineDs('Total mills', 'totalMills', '#2563EB', 'rgba(37, 99, 235, 0.12)'),
        lineDs('High risk', 'highRisk', '#C03030', 'rgba(192, 48, 48, 0.1)'),
        lineDs('No buy list', 'nbl', '#E67E22', 'rgba(230, 126, 34, 0.1)'),
        lineDs('Groups', 'groupCount', '#7C3AED', 'rgba(124, 58, 237, 0.08)'),
      ],
    },
    options: lineChartOptions_(canvas),
  });
}

function pdfDataUrlFormat_(dataUrl) {
  if (!dataUrl || typeof dataUrl !== 'string') return 'JPEG';
  if (dataUrl.indexOf('data:image/png') === 0) return 'PNG';
  return 'JPEG';
}

/** Card tint behind each chart — JPEG flatten matches PDF softCard_ (no dark/light halos). */
export const MILL_EXEC_CHART_CARD_FILLS = {
  risk: [252, 248, 246],
  nbl: [246, 249, 253],
  traceability: [246, 249, 253],
  grievance: [249, 246, 241],
  province: [246, 249, 253],
  facilityQty: [249, 246, 241],
  quarterlyLine: [246, 249, 253],
};

const PDF_CHART_JPEG_Q = 0.62;

function canvasToPdfChartImage_(canvas, cardFill) {
  const fill = cardFill || MILL_EXEC_CHART_CARD_FILLS.nbl;
  if (!canvas || typeof canvas.toDataURL !== 'function') return '';
  const w = canvas.width;
  const h = canvas.height;
  const flat = document.createElement('canvas');
  flat.width = w;
  flat.height = h;
  const ctx = flat.getContext('2d');
  if (!ctx) return canvas.toDataURL('image/jpeg', PDF_CHART_JPEG_Q);
  ctx.fillStyle = 'rgb(' + fill[0] + ',' + fill[1] + ',' + fill[2] + ')';
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(canvas, 0, 0);
  return flat.toDataURL('image/jpeg', PDF_CHART_JPEG_Q);
}

export function collectMillExecutiveChartImages_(els) {
  const out = {};
  if (!els) return out;
  Object.keys(els).forEach(function(key) {
    const canvas = els[key];
    if (!canvas || typeof canvas.toDataURL !== 'function') return;
    try {
      out[key] = canvasToPdfChartImage_(canvas, MILL_EXEC_CHART_CARD_FILLS[key]);
    } catch (_) { /* noop */ }
  });
  return out;
}

export async function collectMillExecutiveChartImagesAsync_(els) {
  const out = {};
  if (!els) return out;
  const keys = Object.keys(els);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const canvas = els[key];
    if (!canvas || typeof canvas.toDataURL !== 'function') continue;
    try {
      out[key] = canvasToPdfChartImage_(canvas, MILL_EXEC_CHART_CARD_FILLS[key]);
    } catch (_) { /* noop */ }
    await yieldToBrowser_(32);
  }
  return out;
}

export function buildMillExecutiveInsights_(snapshot, meta) {
  meta = meta || {};
  const total = snapshot.totalMills || 1;
  const tracePct = snapshot.ttpTrace && !isNaN(snapshot.ttpTrace.pct)
    ? Math.round(snapshot.ttpTrace.pct)
    : (function() {
      const traceTotal = (snapshot.traceability.Traceable || 0) + (snapshot.traceability.Untraceable || 0);
      return traceTotal ? Math.round((snapshot.traceability.Traceable / traceTotal) * 100) : 0;
    })();
  const highPct = Math.round((snapshot.highRisk / total) * 100);
  const nblPct = Math.round((snapshot.nbl / total) * 100);

  const topProv = topEntries_(snapshot.provinces, 1)[0];
  const topFac = topEntries_(snapshot.facilityQty, 1)[0];
  const insights = [];

  insights.push({
    label: 'TTP traceability',
    value: tracePct + '% supply traceable (Facility Performance weighted)'
      + (snapshot.ttpTrace
        ? ' · CPO ' + (snapshot.ttpTrace.cpoFmt || '—') + ' · PK ' + (snapshot.ttpTrace.pkFmt || '—')
        : ''),
    tone: tracePct >= 90 ? 'good' : tracePct >= 70 ? 'warn' : 'bad',
  });

  const grv = snapshot.grievanceProgress;
  if (grv) {
    const grvVal = grv.total > 0
      ? grv.closed + '/' + grv.total + ' closed · ' + grv.pct + '% done (' + grv.year + ')'
      : 'No grievances in ' + grv.year + ' (Date Received)';
    insights.push({
      label: 'Grievance',
      value: grvVal,
      tone: grv.total === 0 ? 'neutral' : grv.pct >= 100 ? 'good' : grv.pct >= 50 ? 'warn' : 'bad',
    });
  }

  insights.push({
    label: 'Risk exposure',
    value: highPct + '% high risk · ' + snapshot.highRisk + ' mills',
    tone: highPct <= 25 ? 'good' : highPct <= 45 ? 'warn' : 'bad',
  });

  insights.push({
    label: 'No Buy List',
    value: nblPct + '% on NBL · ' + snapshot.nbl + ' mills',
    tone: nblPct <= 10 ? 'good' : nblPct <= 20 ? 'warn' : 'bad',
  });

  if (topProv) {
    insights.push({
      label: 'Top province',
      value: truncLabel_(topProv[0], 24) + ' · ' + topProv[1] + ' mills',
      tone: 'neutral',
    });
  }

  if (topFac && topFac[1] > 0) {
    const qty = topFac[1];
    const qtyStr = qty >= 1000 ? (Math.round(qty / 100) / 10) + 'K ton supply' : Math.round(qty) + ' ton supply';
    insights.push({
      label: 'Top supplier',
      value: truncLabel_(topFac[0], 20) + ' · ' + qtyStr,
      tone: 'neutral',
    });
  }

  const quarter = parseInt(String(meta.quarter || ''), 10);
  if (snapshot.quarterlyTrend && quarter > 1) {
    const cur = snapshot.quarterlyTrend[quarter - 1];
    const prev = snapshot.quarterlyTrend[quarter - 2];
    if (cur && prev) {
      const delta = cur.totalMills - prev.totalMills;
      insights.push({
        label: 'Quarter change',
        value: (delta >= 0 ? '+' : '') + delta + ' mills vs Q' + (quarter - 1),
        tone: delta >= 0 ? 'good' : 'warn',
      });
    }
  }

  return insights.slice(0, 6);
}

const MILL_EXEC_PDF_PAL = {
  hero: [252, 250, 246],
  blush: [252, 248, 246],
  mist: [246, 249, 253],
  sand: [249, 246, 241],
};
const MILL_EXEC_PDF_CARD_ALPHA = 0.86;

function millExecPdfWithCardAlpha_(doc, fn) {
  const canAlpha = typeof doc.setGState === 'function' && typeof doc.GState === 'function';
  if (canAlpha) {
    doc.saveGraphicsState();
    doc.setGState(new doc.GState({ opacity: MILL_EXEC_PDF_CARD_ALPHA }));
  }
  fn();
  if (canAlpha) doc.restoreGraphicsState();
}

function millExecPdfSoftCard_(doc, cx, cy, cw, ch, fill, radius) {
  radius = radius || 5;
  millExecPdfWithCardAlpha_(doc, function() {
    const sh = [
      Math.max(0, fill[0] - 14),
      Math.max(0, fill[1] - 14),
      Math.max(0, fill[2] - 14),
    ];
    doc.setFillColor(sh[0], sh[1], sh[2]);
    doc.roundedRect(cx + 0.4, cy + 0.6, cw, ch, radius, radius, 'F');
    doc.setFillColor(fill[0], fill[1], fill[2]);
    doc.roundedRect(cx, cy, cw, ch, radius, radius, 'F');
  });
  doc.setDrawColor(255, 255, 255);
  doc.setLineWidth(0.35);
  doc.roundedRect(cx, cy, cw, ch, radius, radius, 'S');
}

async function millExecPdfApplyPageBackground_(doc, pw, ph, pageBg) {
  if (pageBg) {
    try {
      doc.addImage(pageBg, pdfDataUrlFormat_(pageBg), 0, 0, pw, ph, undefined, 'FAST');
      await yieldToBrowser_(16);
    } catch (err) {
      console.warn('[Mill Executive] Page background skipped:', err);
      doc.setFillColor(250, 247, 242);
      doc.rect(0, 0, pw, ph, 'F');
    }
  } else {
    doc.setFillColor(250, 247, 242);
    doc.rect(0, 0, pw, ph, 'F');
  }
}

/** One landscape executive page (hero KPIs + chart grid + summary). */
async function paintMillExecutiveReportPage_(doc, pw, ph, meta, snapshot, chartImages) {
  const M = 14;
  const W = pw - M * 2;
  const imgs = chartImages || {};
  const PAL = MILL_EXEC_PDF_PAL;

  const FOOTER_H = 5;
  const INSIGHTS_H = 13;
  const BOTTOM_PAD = 3;
  const chartsMaxY = ph - FOOTER_H - BOTTOM_PAD - INSIGHTS_H - BOTTOM_PAD;

  const HERO_Y = 6;
  const HERO_H = 33;
  const HERO_PAD = 6;
  const innerX = M + HERO_PAD;

  millExecPdfSoftCard_(doc, M, HERO_Y, W, HERO_H, PAL.hero, 6);

  const heroInnerW = W - HERO_PAD * 2;
  const kpiAreaW = heroInnerW * 0.44;
  const textAreaW = heroInnerW - kpiAreaW - 5;
  const kpiAreaX = innerX + textAreaW + 5;

  const kpis = [
    { label: 'Total mills', value: snapshot.totalMills, style: KPI_STYLES.totalMills },
    { label: 'Groups', value: snapshot.groupCount, style: KPI_STYLES.groups },
    { label: 'High risk', value: snapshot.highRisk, style: KPI_STYLES.highRisk },
    { label: 'No buy list', value: snapshot.nbl, style: KPI_STYLES.noBuyList },
  ];
  const kpiBoxGap = 2.5;
  const kpiBoxW = (kpiAreaW - kpiBoxGap * 3) / 4;
  const kpiBoxH = HERO_H - HERO_PAD * 2;
  const kpiBoxY = HERO_Y + HERO_PAD;
  const kpiFill = [255, 255, 252];

  const titleLineGap = 7;
  const titleBlockH = 14;
  const titleTop = kpiBoxY + (kpiBoxH - titleBlockH) / 2;
  setPdfFont_(doc, 'serif');
  doc.setFontSize(27);
  doc.setTextColor(15, 12, 10);
  doc.text('Mill Onboarding', innerX, titleTop + 8);

  setPdfFont_(doc, 'sans-bold');
  doc.setFontSize(10);
  doc.setTextColor(55, 48, 45);
  doc.text('Executive report  ·  ' + meta.periodLabel, innerX, titleTop + 8 + titleLineGap);
  if (meta.kpiCaption) {
    setPdfFont_(doc, 'sans');
    doc.setFontSize(7);
    doc.setTextColor(120, 100, 100);
    doc.text(meta.kpiCaption, innerX, titleTop + 8 + titleLineGap + 4);
  }

  kpis.forEach(function(k, i) {
    const boxX = kpiAreaX + i * (kpiBoxW + kpiBoxGap);
    millExecPdfSoftCard_(doc, boxX, kpiBoxY, kpiBoxW, kpiBoxH, kpiFill, 4);

    const st = k.style;
    setPdfFont_(doc, 'sans-bold');
    doc.setFontSize(19);
    doc.setTextColor(st.num[0], st.num[1], st.num[2]);
    doc.text(String(k.value), boxX + kpiBoxW / 2, kpiBoxY + kpiBoxH / 2 - 1, { align: 'center' });

    setPdfFont_(doc, 'sans-bold');
    doc.setFontSize(7);
    doc.setTextColor(87, 83, 78);
    doc.text(k.label, boxX + kpiBoxW / 2, kpiBoxY + kpiBoxH / 2 + 5.5, { align: 'center' });
  });

  const chartsY = HERO_Y + HERO_H + 5;
  const rowGap = 4;
  const TITLE_H = 8;
  const IMG_PAD = 3;
  const col4Gap = 3.5;
  const col4W = (W - col4Gap * 3) / 4;
  const col2Gap = 4;
  const col2W = (W - col2Gap) / 2;
  const chartArea = chartsMaxY - chartsY;
  const row1H = Math.floor(chartArea * 0.48);
  const row2H = chartArea - row1H - rowGap;
  const hBarW = MILL_EXEC_HBAR_CHART_SIZE.w;
  const hBarH = MILL_EXEC_HBAR_CHART_SIZE.h;
  const compare = meta.fullYearComparison === true;
  const chartTitles = compare ? {
    risk: 'Risk level distribution',
    nbl: 'No buy list status',
    traceability: 'TTP supply traceability',
    grievance: 'Grievance progress',
    province: 'Province distribution',
    facilityQty: 'Top supplier (supply qty, ton)',
  } : {
    risk: 'Risk level distribution',
    nbl: 'No buy list status',
    traceability: 'TTP supply traceability',
    grievance: 'Grievance progress',
    province: 'Province distribution',
    facilityQty: 'Top supplier (supply qty, ton)',
  };

  async function chartPanel_(title, imgKey, canvasW, canvasH, cx, cy, cw, ch, fill) {
    millExecPdfSoftCard_(doc, cx, cy, cw, ch, fill || PAL.mist, 5);

    setPdfFont_(doc, 'sans-bold');
    doc.setFontSize(8.5);
    doc.setTextColor(41, 37, 36);
    doc.text(title, cx + 5, cy + 6.5);

    if (!imgs[imgKey]) return;

    const bx = cx + IMG_PAD;
    const by = cy + TITLE_H;
    const bw = cw - IMG_PAD * 2;
    const bh = ch - TITLE_H - IMG_PAD;
    const ratio = canvasW / canvasH;
    const boxR = bw / bh;
    let dw; let dh; let dx; let dy;
    if (ratio > boxR) {
      dw = bw; dh = bw / ratio; dx = bx; dy = by + (bh - dh) / 2;
    } else {
      dh = bh; dw = bh * ratio; dy = by; dx = bx + (bw - dw) / 2;
    }
    doc.addImage(imgs[imgKey], pdfDataUrlFormat_(imgs[imgKey]), dx, dy, dw, dh, undefined, 'FAST');
    await yieldToBrowser_(12);
  }

  const r1y = chartsY;
  await chartPanel_(chartTitles.risk, 'risk', 320, 320, M, r1y, col4W, row1H, PAL.blush);
  await chartPanel_(chartTitles.nbl, 'nbl', 320, 320, M + col4W + col4Gap, r1y, col4W, row1H, PAL.mist);
  await chartPanel_(chartTitles.traceability, 'traceability', 320, 320, M + (col4W + col4Gap) * 2, r1y, col4W, row1H, PAL.mist);
  await chartPanel_(chartTitles.grievance, 'grievance', 320, 320, M + (col4W + col4Gap) * 3, r1y, col4W, row1H, PAL.sand);

  const r2y = r1y + row1H + rowGap;
  await chartPanel_(chartTitles.province, 'province', hBarW, hBarH, M, r2y, col2W, row2H, PAL.mist);
  await chartPanel_(chartTitles.facilityQty, 'facilityQty', hBarW, hBarH, M + col2W + col2Gap, r2y, col2W, row2H, PAL.sand);

  const insightsY = chartsMaxY + BOTTOM_PAD;
  const insights = meta.customInsights
    || buildMillExecutiveInsights_(snapshot, meta);

  setPdfFont_(doc, 'serif');
  doc.setFontSize(9);
  doc.setTextColor(168, 162, 158);
  doc.text('Summary', pw / 2, insightsY + 4.5, { align: 'center' });

  const summaryText = insights.map(function(ins) { return ins.value; }).join('   ·   ');
  setPdfFont_(doc, 'sans-bold');
  doc.setFontSize(8);
  doc.setTextColor(41, 37, 36);
  const summaryLines = doc.splitTextToSize(summaryText, W);
  doc.text(summaryLines.slice(0, 2), pw / 2, insightsY + 11, { align: 'center' });

  setPdfFont_(doc, 'sans');
  doc.setFontSize(6);
  doc.setTextColor(168, 162, 158);
  const footerParts = [
    meta.footerNote || '',
    'Unique mills (excl. Trader / Refinery)',
    'Deduplicated by mill name',
    meta.productView || '',
  ].filter(Boolean);
  doc.text(footerParts.join('  ·  '), pw / 2, ph - 2.5, { align: 'center' });
}

export async function exportMillExecutivePdf_(meta, snapshot, chartImages, getJsPDF) {
  const JsPDF = typeof getJsPDF === 'function' ? getJsPDF() : null;
  if (!JsPDF) throw new Error('PDF library not loaded. Refresh the page and try again.');

  const doc = new JsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
  await registerPdfFonts_(doc);

  const pw = doc.internal.pageSize.getWidth();
  const ph = doc.internal.pageSize.getHeight();
  const pageBg = await getMillExecutiveBackgroundDataUrl_();
  await millExecPdfApplyPageBackground_(doc, pw, ph, pageBg);
  await paintMillExecutiveReportPage_(doc, pw, ph, meta, snapshot, chartImages);

  await yieldToBrowser_(48);
  await new Promise(function(resolve) {
    requestAnimationFrame(function() { setTimeout(resolve, 0); });
  });
  doc.save(meta.filename);
}

/**
 * Full-year PDF — one page, same grid as quarterly; each chart compares Q1–Q4.
 */
export async function exportMillExecutiveFullYearPdf_(meta, snapshot, chartImages, getJsPDF) {
  const JsPDF = typeof getJsPDF === 'function' ? getJsPDF() : null;
  if (!JsPDF) throw new Error('PDF library not loaded. Refresh the page and try again.');
  if (!snapshot) throw new Error('No executive report data available.');

  const doc = new JsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
  await registerPdfFonts_(doc);

  const pw = doc.internal.pageSize.getWidth();
  const ph = doc.internal.pageSize.getHeight();
  const pageBg = await getMillExecutiveBackgroundDataUrl_();
  await millExecPdfApplyPageBackground_(doc, pw, ph, pageBg);

  const pageMeta = Object.assign({}, meta, {
    fullYearComparison: true,
    footerNote: 'Full year ' + (meta.year || '') + ' · quarterly comparison (as-of each quarter end)',
  });
  await paintMillExecutiveReportPage_(doc, pw, ph, pageMeta, snapshot, chartImages);

  await yieldToBrowser_(48);
  await new Promise(function(resolve) {
    requestAnimationFrame(function() { setTimeout(resolve, 0); });
  });
  doc.save(meta.filename);
}
