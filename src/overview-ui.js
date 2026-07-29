import './overview-landing.css';

let refreshBusy = false;
let welcomeName = 'there';
let overviewHubRoot = null;

const OVERVIEW_CACHE_KEY = 'sustain-dashboard.overviewMetrics.v1';

function currentPeriodKey_() {
  const now = new Date();
  const year = now.getFullYear();
  const quarter = Math.ceil((now.getMonth() + 1) / 3);
  return year + '-Q' + quarter;
}

function readOverviewCache_() {
  try {
    const raw = localStorage.getItem(OVERVIEW_CACHE_KEY);
    if (!raw) return null;
    const entry = JSON.parse(raw);
    if (!entry || !entry.periodKey || !entry.payload) return null;
    return entry;
  } catch (_e) {
    return null;
  }
}

function sanitizePayloadForCache_(payload) {
  if (!payload || typeof payload !== 'object') return payload;
  const out = Object.assign({}, payload);
  if (payload.snapshot && typeof payload.snapshot === 'object') {
    const snap = Object.assign({}, payload.snapshot);
    if (snap.groups && typeof snap.groups.size === 'number' && snap.groupCount == null) {
      snap.groupCount = snap.groups.size;
    }
    delete snap.groups;
    out.snapshot = snap;
  }
  return out;
}

function writeOverviewCache_(periodKey, payload) {
  try {
    localStorage.setItem(
      OVERVIEW_CACHE_KEY,
      JSON.stringify({
        periodKey: periodKey,
        savedAt: Date.now(),
        payload: sanitizePayloadForCache_(payload),
      })
    );
  } catch (e) {
    console.warn('[overview] cache write failed', e);
  }
}

function bindQuarterCacheWatch_() {
  if (bindQuarterCacheWatch_.bound) return;
  bindQuarterCacheWatch_.bound = true;
  document.addEventListener('visibilitychange', function() {
    if (document.visibilityState !== 'visible' || !overviewHubRoot) return;
    refreshHub_(overviewHubRoot, { force: false });
  });
}
bindQuarterCacheWatch_.bound = false;

const MODULES = [
  { id: 'ndpe', title: 'NDPE & Forest protection', icon: 'forest' },
  { id: 'trace', title: 'Traceability to source', icon: 'trace' },
  { id: 'eudr', title: 'EUDR readiness', icon: 'eudr' },
  { id: 'grievance', title: 'Grievance & remediation', icon: 'grievance' },
  { id: 'nbl', title: 'No Buy List', icon: 'nbl' },
  { id: 'cert', title: 'Standards & certification', icon: 'cert' },
  { id: 'supply', title: 'Responsible supply base', icon: 'supply' },
  { id: 'report', title: 'Monthly reporting', icon: 'report' },
];

const ICONS = {
  forest: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 22V12"/><path d="M7 12L12 3l5 9"/><path d="M5 12h14"/></svg>',
  trace: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>',
  eudr: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/></svg>',
  grievance: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/></svg>',
  nbl: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/></svg>',
  cert: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="8" r="5"/><path d="M8.5 14.5L7 22l5-2.5L17 22l-1.5-7.5"/></svg>',
  supply: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>',
  report: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>',
};

const PANEL_ICONS = {
  kpi: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/></svg>',
  snapshot: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 11-2.12-9.36L23 10"/></svg>',
};

function formatWelcomeName(emailRaw) {
  const s = (emailRaw || '').trim();
  if (!s) return 'there';
  const local = s.split('@')[0] || s;
  return local
    .replace(/[._-]+/g, ' ')
    .split(' ')
    .filter(Boolean)
    .map(function(w) { return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); })
    .join(' ');
}

function pct_(part, total) {
  const t = Number(total) || 0;
  if (!t) return 0;
  return Math.round((Number(part) / t) * 100);
}

function nextReportLabel_() {
  const d = new Date();
  const next = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return next.toLocaleDateString('en-GB', { month: 'short', day: 'numeric' });
}

function buildHubModel_(payload) {
  const snap = (payload && payload.snapshot) || {};
  const total = snap.totalMills || 0;
  const nbl = snap.nbl || 0;
  const nonNblPct = total ? Math.round(((total - nbl) / total) * 100) : 0;
  const high = snap.highRisk || 0;
  const ndpePct = Math.max(0, Math.min(100, nonNblPct || (100 - pct_(high, total))));

  let tracePct = 0;
  if (snap.ttpTrace && !isNaN(snap.ttpTrace.pct)) tracePct = Math.round(snap.ttpTrace.pct);
  else {
    const tr = snap.traceability || {};
    const trT = (tr.Traceable || 0) + (tr.Untraceable || 0);
    tracePct = trT ? Math.round((tr.Traceable / trT) * 100) : 0;
  }

  const grv = snap.grievanceProgress || {};
  const grvOpen = grv.open || 0;
  const grvResolved = grv.total ? Math.round(grv.pct || (grv.closed / grv.total) * 100) : 100;

  const eudrP = payload && payload.eudr ? payload.eudr.potential : 0;
  const eudrTotal = (payload && payload.eudr ? payload.eudr.potential + payload.eudr.notPotential : 0) || total;
  const eudrPct = eudrTotal ? Math.round((eudrP / eudrTotal) * 100) : 0;

  const certCount = Object.entries(snap.certification || {})
    .filter(function(e) { return e[0] !== 'None'; })
    .reduce(function(sum, e) { return sum + e[1]; }, 0);

  const supplierCount = Object.keys(snap.facilityQty || {}).filter(Boolean).length;

  const pills = {
    ndpe: ndpePct >= 85 ? { text: 'Compliant', tone: 'green' } : { text: 'Review', tone: 'amber' },
    trace: tracePct >= 80 ? { text: 'On track', tone: 'green' } : { text: 'Gap found', tone: 'amber' },
    eudr: eudrPct >= 50 ? { text: 'Potential', tone: 'blue' } : { text: 'In review', tone: 'blue' },
    grievance: grvOpen > 0 ? { text: grvOpen + ' open', tone: 'amber' } : { text: 'Clear', tone: 'green' },
    nbl: nbl > 0 ? { text: nbl + ' blocked', tone: 'amber' } : { text: 'Clear', tone: 'green' },
    cert: { text: 'On track', tone: 'green' },
    supply: { text: 'Monitored', tone: 'green' },
    report: { text: 'Ready', tone: 'blue' },
  };

  const moduleStats = {
    ndpe: { value: String(total), sub: ndpePct + '% compliant · ' + high + ' high risk' },
    trace: { value: tracePct + '%', sub: 'traceable supply' },
    eudr: { value: String(eudrP), sub: eudrPct + '% potential · ' + (eudrTotal - eudrP) + ' not' },
    grievance: { value: String(grvOpen), sub: grvResolved + '% resolved YTD' },
    nbl: { value: String(nbl), sub: nbl ? 'blocked suppliers' : 'none blocked' },
    cert: { value: String(certCount), sub: 'certified mills' },
    supply: { value: String(supplierCount), sub: 'active suppliers' },
    report: { value: nextReportLabel_(), sub: payload && payload.periodLabel ? payload.periodLabel : currentPeriodKey_() },
  };

  return {
    hero: {
      ndpePct: ndpePct,
      grievances: grvOpen,
      nbl: nbl,
    },
    progress: [
      { label: 'NDPE', pct: ndpePct, tone: 'burgundy' },
      { label: 'Traceability', pct: tracePct, tone: 'teal' },
      { label: 'EUDR', pct: eudrPct, tone: 'blue' },
      { label: 'Grievance closed', pct: grvResolved, tone: 'amber' },
    ],
    pills: pills,
    moduleStats: moduleStats,
    activity: buildActivity_(payload, tracePct, nbl, grvOpen, ndpePct),
  };
}

function buildActivity_(payload, tracePct, nbl, grvOpen, ndpePct) {
  const items = [
    { tone: 'green', title: 'NDPE ' + ndpePct + '% · TTP ' + tracePct + '%', time: 'Snapshot' },
  ];
  if (nbl > 0) {
    items.push({ tone: 'red', title: 'NBL: ' + nbl, time: 'Now' });
  }
  if (grvOpen > 0) {
    items.push({ tone: 'blue', title: 'Grievances open: ' + grvOpen, time: 'YTD' });
  }
  if (tracePct > 0 && tracePct < 80) {
    items.push({ tone: 'amber', title: 'TTP below 80% (' + tracePct + '%)', time: 'Flag' });
  }
  if (payload && payload.periodLabel) {
    items.push({ tone: 'green', title: payload.periodLabel, time: String(new Date().getFullYear()) });
  }
  return items.slice(0, 4);
}

function renderHub_(root, model) {
  const welcome = root.querySelector('#ovHubWelcome');
  if (welcome) welcome.textContent = 'Welcome back, ' + welcomeName;

  const heroStats = root.querySelector('#ovHubHeroStats');
  if (heroStats && model.hero) {
    heroStats.innerHTML = [
      { v: model.hero.ndpePct + '%', l: 'NDPE Compliance' },
      { v: String(model.hero.grievances), l: 'Active Grievances' },
      { v: String(model.hero.nbl), l: 'NBL Suppliers' },
    ].map(function(s) {
      return '<div class="ov-hub__glass"><div class="ov-hub__glass-val">' + s.v + '</div><div class="ov-hub__glass-lbl">' + s.l + '</div></div>';
    }).join('');
  }

  const grid = root.querySelector('#ovHubModules');
  if (grid) {
    grid.textContent = '';
    MODULES.forEach(function(m) {
      const pill = model.pills[m.id] || { text: '—', tone: 'blue' };
      const st = (model.moduleStats && model.moduleStats[m.id]) || { value: '—', sub: '' };
      const art = document.createElement('article');
      art.className = 'ov-hub__mod';
      art.innerHTML =
        '<div class="ov-hub__mod-head">'
        + '<span class="ov-hub__mod-icon ov-hub__mod-icon--' + m.icon + '">' + (ICONS[m.icon] || '') + '</span>'
        + '<span class="ov-hub__pill ov-hub__pill--' + pill.tone + '">' + pill.text + '</span>'
        + '</div>'
        + '<h3 class="ov-hub__mod-title">' + m.title + '</h3>'
        + '<div class="ov-hub__mod-stat"><span class="ov-hub__mod-stat-val">' + st.value + '</span>'
        + (st.sub ? '<span class="ov-hub__mod-stat-sub">' + st.sub + '</span>' : '')
        + '</div>';
      grid.appendChild(art);
    });
  }

  const prog = root.querySelector('#ovHubProgress');
  if (prog) {
    prog.innerHTML = model.progress.map(function(p) {
      return '<div class="ov-hub__prog-row">'
        + '<div class="ov-hub__prog-meta"><span>' + p.label + '</span><strong>' + p.pct + '%</strong></div>'
        + '<div class="ov-hub__prog-track"><div class="ov-hub__prog-fill ov-hub__prog-fill--' + p.tone + '" style="width:' + p.pct + '%"></div></div>'
        + '</div>';
    }).join('');
  }

  const act = root.querySelector('#ovHubActivity');
  if (act) {
    act.innerHTML = model.activity.map(function(a) {
      return '<li class="ov-hub__act-item">'
        + '<span class="ov-hub__act-dot ov-hub__act-dot--' + a.tone + '"></span>'
        + '<div class="ov-hub__act-title">' + a.title + '</div>'
        + '<time class="ov-hub__act-time">' + a.time + '</time></li>';
    }).join('');
  }
}

async function refreshHub_(root, options) {
  if (refreshBusy) return;
  const force = !!(options && options.force);
  const periodKey = currentPeriodKey_();
  const btn = root.querySelector('#ovHubRefresh');

  if (!force) {
    const cached = readOverviewCache_();
    if (cached && cached.periodKey === periodKey && cached.payload) {
      renderHub_(root, buildHubModel_(cached.payload));
      return;
    }
  }

  refreshBusy = true;
  if (btn) btn.disabled = true;
  try {
    const fetcher = typeof window.refreshOverviewMetricsData_ === 'function'
      ? window.refreshOverviewMetricsData_
      : null;
    if (!fetcher) {
      const cached = readOverviewCache_();
      const payload = cached && cached.payload ? cached.payload : { snapshot: {} };
      renderHub_(root, buildHubModel_(payload));
      return;
    }
    const payload = await fetcher();
    writeOverviewCache_(periodKey, payload);
    renderHub_(root, buildHubModel_(payload));
  } catch (e) {
    console.warn('[overview]', e);
    const cached = readOverviewCache_();
    if (cached && cached.payload && cached.periodKey === periodKey) {
      renderHub_(root, buildHubModel_(cached.payload));
    } else {
      renderHub_(root, buildHubModel_({ snapshot: {} }));
    }
  } finally {
    refreshBusy = false;
    if (btn) btn.disabled = false;
  }
}

export function updateOverviewWelcomeFromEmail(emailRaw) {
  welcomeName = formatWelcomeName(emailRaw);
  const el = document.getElementById('ovHubWelcome');
  if (el) el.textContent = 'Welcome back, ' + welcomeName;
}

export function mountOverviewLanding(container) {
  if (!container) return;

  container.textContent = '';
  const root = document.createElement('div');
  root.className = 'ov-hub';
  root.innerHTML = `
    <section class="ov-hub__hero">
      <div class="ov-hub__hero-inner">
        <div class="ov-hub__hero-copy">
          <p class="ov-hub__welcome" id="ovHubWelcome">Welcome back, there</p>
          <h1 class="ov-hub__title">Overview</h1>
          <div class="ov-hub__actions">
            <button type="button" class="ov-hub__refresh ov-hub__glass-btn" id="ovHubRefresh">Refresh</button>
          </div>
        </div>
        <div class="ov-hub__hero-stats" id="ovHubHeroStats"></div>
      </div>
    </section>

    <div class="ov-hub__section-head">
      <h2 class="ov-hub__section-title">Modules</h2>
    </div>
    <div class="ov-hub__modules" id="ovHubModules"></div>

    <div class="ov-hub__bottom">
      <section class="ov-hub__panel">
        <h3 class="ov-hub__panel-title"><span class="ov-hub__panel-icon ov-hub__panel-icon--kpi">${PANEL_ICONS.kpi}</span> KPI progress</h3>
        <div class="ov-hub__progress" id="ovHubProgress"></div>
      </section>
      <section class="ov-hub__panel">
        <h3 class="ov-hub__panel-title"><span class="ov-hub__panel-icon ov-hub__panel-icon--snapshot">${PANEL_ICONS.snapshot}</span> Snapshot</h3>
        <ul class="ov-hub__activity" id="ovHubActivity"></ul>
      </section>
    </div>
  `;
  container.appendChild(root);
  overviewHubRoot = root;
  bindQuarterCacheWatch_();

  root.querySelector('#ovHubRefresh').addEventListener('click', function() {
    refreshHub_(root, { force: true });
  });
  window.__overviewMetricsRefresh = function() {
    refreshHub_(root, { force: false });
  };
  refreshHub_(root, { force: false });
}
