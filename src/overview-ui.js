import './overview-landing.css';

let refreshBusy = false;
let welcomeName = 'there';
let overviewHubRoot = null;

const OVERVIEW_CACHE_KEY = 'sustain-dashboard.overviewMetrics.v7';

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

function writeOverviewCache_(payload) {
  const periodKey = (payload && payload.dataPeriodKey) || currentPeriodKey_();
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

function bindDataPeriodWatch_() {
  if (bindDataPeriodWatch_.bound) return;
  bindDataPeriodWatch_.bound = true;
  document.addEventListener('visibilitychange', function() {
    if (document.visibilityState !== 'visible' || !overviewHubRoot) return;
    refreshHub_(overviewHubRoot, { force: false });
  });
}
bindDataPeriodWatch_.bound = false;

/** Module titles match sidebar menu labels exactly. */
const MODULES = [
  { id: 'mill', title: 'Mill Onboarding', icon: 'mill' },
  { id: 'trace', title: 'Traceability Data', icon: 'trace' },
  { id: 'eudr', title: 'EUDR Potential', icon: 'eudr' },
  { id: 'grievance', title: 'Grievance Monitoring', icon: 'grievance' },
  { id: 'nbl', title: 'No Buy List', icon: 'nbl' },
  { id: 'facility', title: 'Facility Performance', icon: 'facility' },
  { id: 'report', title: 'Monthly Report', icon: 'report' },
];

const ICONS = {
  mill: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><rect x="2" y="7" width="20" height="14" rx="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2"/><line x1="12" y1="12" x2="12" y2="16"/><line x1="10" y1="14" x2="14" y2="14"/></svg>',
  trace: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>',
  eudr: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/></svg>',
  grievance: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/></svg>',
  nbl: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="12" cy="12" r="10"/><line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/></svg>',
  facility: '<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M12 20V10"/><path d="M18 20V4"/><path d="M6 20v-4"/></svg>',
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

function overviewAvgTracePct_(a, b) {
  const na = Number(a);
  const nb = Number(b);
  if (!isNaN(na) && !isNaN(nb)) return Math.round((na + nb) / 2);
  if (!isNaN(na)) return Math.round(na);
  if (!isNaN(nb)) return Math.round(nb);
  return NaN;
}

function overviewFmtTracePct_(n) {
  if (n == null || isNaN(n)) return '—';
  return (Math.round(Number(n) * 100) / 100).toFixed(2) + '%';
}

function overviewTracePctNum_(traceTotals, product) {
  if (!traceTotals) return NaN;
  return product === 'pk' ? Number(traceTotals.ttpPkPct) : Number(traceTotals.ttpCpoPct);
}

function buildHubModel_(payload) {
  const snap = (payload && payload.snapshot) || {};
  const traceTotals = (payload && payload.traceTotals) || null;
  const total = snap.totalMills || 0;
  const nbl = snap.nbl || 0;
  const high = snap.highRisk || 0;
  const groupCount = snap.groupCount != null
    ? snap.groupCount
    : (snap.groups && typeof snap.groups.size === 'number' ? snap.groups.size : 0);

  const ttpCpoPct = overviewTracePctNum_(traceTotals, 'cpo');
  const ttpPkPct = overviewTracePctNum_(traceTotals, 'pk');
  const ttpCpoLabel = traceTotals ? (traceTotals.ttpCpoFmt || overviewFmtTracePct_(ttpCpoPct)) : '—';
  const ttpPkLabel = traceTotals ? (traceTotals.ttpPkFmt || overviewFmtTracePct_(ttpPkPct)) : '—';
  const ttpAvg = traceTotals
    ? overviewAvgTracePct_(ttpCpoPct, ttpPkPct)
    : NaN;

  let tracePct = 0;
  if (!isNaN(ttpAvg)) tracePct = ttpAvg;
  else if (snap.ttpTrace && !isNaN(snap.ttpTrace.pct)) tracePct = Math.round(snap.ttpTrace.pct);
  else {
    const tr = snap.traceability || {};
    const trT = (tr.Traceable || 0) + (tr.Untraceable || 0);
    tracePct = trT ? Math.round((tr.Traceable / trT) * 100) : 0;
  }

  const tracePeriodLabel = (payload && payload.tracePeriod && payload.tracePeriod.label) || '';
  const traceDetailSub = tracePeriodLabel || (tracePct + '% traceable');

  const grv = snap.grievanceProgress || {};
  const grvOpen = grv.open || 0;
  const grvClosed = grv.closed || 0;
  const grvTotal = grv.total || 0;
  const grvResolved = grvTotal ? Math.round(grv.pct || (grvClosed / grvTotal) * 100) : 100;

  const eudrP = payload && payload.eudr ? payload.eudr.potential : 0;
  const eudrNot = payload && payload.eudr ? payload.eudr.notPotential : 0;
  const eudrTotal = (payload && payload.eudr && payload.eudr.total)
    ? payload.eudr.total
    : (eudrP + eudrNot || total);
  const eudrPct = eudrTotal ? Math.round((eudrP / eudrTotal) * 100) : 0;

  const nonNblPct = total ? Math.round(((total - nbl) / total) * 100) : 0;
  const ndpePct = Math.max(0, Math.min(100, nonNblPct || (100 - pct_(high, total))));

  const facilityCount = (payload && payload.facilityCount) || 0;
  const ttpMills = (payload && payload.ttpMills) || 0;
  const ttpRecords = (payload && payload.ttpRecords) || 0;

  const periodSub = payload && payload.periodLabel
    ? payload.periodLabel
    : currentPeriodKey_();

  const traceOnTrack = !isNaN(ttpCpoPct) && !isNaN(ttpPkPct) && ttpCpoPct >= 80 && ttpPkPct >= 80;

  const pills = {
    mill: high > 0
      ? { text: high + ' high risk', tone: 'amber' }
      : { text: 'Monitored', tone: 'green' },
    trace: traceOnTrack
      ? { text: 'On track', tone: 'green' }
      : { text: 'Gap found', tone: 'amber' },
    eudr: eudrP > 0
      ? { text: eudrP + ' potential', tone: 'blue' }
      : { text: 'In review', tone: 'blue' },
    grievance: grvOpen > 0
      ? { text: grvOpen + ' open', tone: 'amber' }
      : { text: 'Clear', tone: 'green' },
    nbl: nbl > 0
      ? { text: nbl + ' on NBL', tone: 'amber' }
      : { text: 'Clear', tone: 'green' },
    facility: facilityCount > 0
      ? { text: 'Monitored', tone: 'green' }
      : { text: 'No data', tone: 'blue' },
    report: { text: 'Ready', tone: 'blue' },
  };

  const traceSub = ttpMills > 0
    ? ttpMills + ' mills · ' + ttpRecords + ' records · ' + traceDetailSub
    : traceDetailSub;

  const ttpDualLabel = 'CPO ' + ttpCpoLabel + ' · PK ' + ttpPkLabel;

  const moduleStats = {
    mill: {
      value: String(total),
      sub: high + ' high risk · ' + groupCount + ' groups',
    },
    trace: {
      dual: true,
      cpo: ttpCpoLabel,
      pk: ttpPkLabel,
      sub: traceSub,
    },
    eudr: {
      value: String(eudrP),
      sub: eudrNot + ' not potential · ' + eudrTotal + ' mills',
    },
    grievance: {
      value: String(grvOpen),
      sub: grvClosed + ' closed · ' + grvTotal + ' total',
    },
    nbl: {
      value: String(nbl),
      sub: nbl ? 'mills on NBL' : 'none on NBL',
    },
    facility: {
      value: String(facilityCount),
      sub: ttpDualLabel + ' · ' + facilityCount + ' facilities tracked',
    },
    report: {
      value: nextReportLabel_(),
      sub: periodSub,
    },
  };

  return {
    hero: {
      totalMills: total,
      highRisk: high,
      nbl: nbl,
    },
    progress: [
      { label: 'NDPE', pct: ndpePct, tone: 'burgundy' },
      {
        label: 'TTP CPO traceable',
        pct: isNaN(ttpCpoPct) ? 0 : Math.min(100, Math.max(0, ttpCpoPct)),
        pctFmt: ttpCpoLabel,
        tone: 'teal',
      },
      {
        label: 'TTP PK traceable',
        pct: isNaN(ttpPkPct) ? 0 : Math.min(100, Math.max(0, ttpPkPct)),
        pctFmt: ttpPkLabel,
        tone: 'teal',
      },
      { label: 'EUDR potential', pct: eudrPct, tone: 'blue' },
      { label: 'Grievance closed', pct: grvResolved, tone: 'amber' },
    ],
    pills: pills,
    moduleStats: moduleStats,
    activity: buildActivity_(payload, ttpCpoLabel, ttpPkLabel, nbl, grvOpen, ndpePct, high, periodSub),
  };
}

function buildActivity_(payload, ttpCpoLabel, ttpPkLabel, nbl, grvOpen, ndpePct, highRisk, periodSub) {
  const items = [
    { tone: 'green', title: 'Mill onboarding · ' + (highRisk || 0) + ' high risk', time: 'Snapshot' },
    { tone: 'green', title: 'TTP CPO ' + ttpCpoLabel + ' · PK ' + ttpPkLabel + ' · NDPE ' + ndpePct + '%', time: 'Snapshot' },
  ];
  if (nbl > 0) {
    items.push({ tone: 'red', title: 'No Buy List: ' + nbl, time: 'Now' });
  }
  if (grvOpen > 0) {
    items.push({ tone: 'blue', title: 'Grievances open: ' + grvOpen, time: 'YTD' });
  }
  if (periodSub) {
    items.push({ tone: 'amber', title: periodSub, time: 'Latest data' });
  }
  return items.slice(0, 4);
}

function renderHub_(root, model) {
  const welcome = root.querySelector('#ovHubWelcome');
  if (welcome) welcome.textContent = 'Welcome back, ' + welcomeName;

  const heroStats = root.querySelector('#ovHubHeroStats');
  if (heroStats && model.hero) {
    heroStats.innerHTML = [
      { v: String(model.hero.totalMills), l: 'Total Mills' },
      { v: String(model.hero.highRisk), l: 'High Risk' },
      { v: String(model.hero.nbl), l: 'No Buy List' },
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
      const statHtml = st.dual
        ? '<div class="ov-hub__mod-stat ov-hub__mod-stat--dual">'
          + '<div class="ov-hub__trace-split">'
          + '<div class="ov-hub__trace-split-item"><span class="ov-hub__trace-split-lbl">CPO</span>'
          + '<span class="ov-hub__trace-split-val">' + (st.cpo || '—') + '</span></div>'
          + '<div class="ov-hub__trace-split-item"><span class="ov-hub__trace-split-lbl">PK</span>'
          + '<span class="ov-hub__trace-split-val">' + (st.pk || '—') + '</span></div>'
          + '</div>'
          + (st.sub ? '<span class="ov-hub__mod-stat-sub">' + st.sub + '</span>' : '')
          + '</div>'
        : '<div class="ov-hub__mod-stat"><span class="ov-hub__mod-stat-val">' + (st.value || '—') + '</span>'
          + (st.sub ? '<span class="ov-hub__mod-stat-sub">' + st.sub + '</span>' : '')
          + '</div>';
      const art = document.createElement('article');
      art.className = 'ov-hub__mod';
      art.innerHTML =
        '<div class="ov-hub__mod-head">'
        + '<span class="ov-hub__mod-icon ov-hub__mod-icon--' + m.icon + '">' + (ICONS[m.icon] || '') + '</span>'
        + '<span class="ov-hub__pill ov-hub__pill--' + pill.tone + '">' + pill.text + '</span>'
        + '</div>'
        + '<h3 class="ov-hub__mod-title">' + m.title + '</h3>'
        + statHtml;
      grid.appendChild(art);
    });
  }

  const prog = root.querySelector('#ovHubProgress');
  if (prog) {
    prog.innerHTML = model.progress.map(function(p) {
      const pctLabel = p.pctFmt || (p.pct + '%');
      const barPct = Math.min(100, Math.max(0, Number(p.pct) || 0));
      return '<div class="ov-hub__prog-row">'
        + '<div class="ov-hub__prog-meta"><span>' + p.label + '</span><strong>' + pctLabel + '</strong></div>'
        + '<div class="ov-hub__prog-track"><div class="ov-hub__prog-fill ov-hub__prog-fill--' + p.tone + '" style="width:' + barPct + '%"></div></div>'
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

function renderHubSkeleton_(root) {
  const grid = root.querySelector('#ovHubModules');
  if (grid && !grid.childElementCount) {
    grid.innerHTML = MODULES.map(function(m) {
      return '<article class="ov-hub__mod ov-hub__mod--skeleton">'
        + '<div class="ov-hub__mod-head"><span class="ov-hub__sk-line ov-hub__sk-line--icon"></span>'
        + '<span class="ov-hub__sk-line ov-hub__sk-line--pill"></span></div>'
        + '<div class="ov-hub__sk-line ov-hub__sk-line--title"></div>'
        + '<div class="ov-hub__sk-line ov-hub__sk-line--stat"></div>'
        + '<div class="ov-hub__sk-line ov-hub__sk-line--sub"></div>'
        + '</article>';
    }).join('');
  }
  const prog = root.querySelector('#ovHubProgress');
  if (prog && !prog.innerHTML) {
    prog.innerHTML = [1, 2, 3, 4].map(function() {
      return '<div class="ov-hub__prog-row ov-hub__prog-row--skeleton">'
        + '<div class="ov-hub__sk-line ov-hub__sk-line--prog"></div>'
        + '</div>';
    }).join('');
  }
}

async function refreshHub_(root, options) {
  const force = !!(options && options.force);
  if (refreshBusy) {
    if (force) refreshHub_.queuedForce = true;
    return;
  }
  const fetcher = typeof window.refreshOverviewMetricsData_ === 'function'
    ? window.refreshOverviewMetricsData_
    : null;
  const fastFetcher = typeof window.refreshOverviewMetricsDataFast_ === 'function'
    ? window.refreshOverviewMetricsDataFast_
    : null;

  if (!fetcher) {
    scheduleOverviewFetcherRetry_(root);
    renderHubSkeleton_(root);
    setOverviewLoading_(root, true);
    return;
  }

  const cached = readOverviewCache_();
  const hasGoodCache = !!(cached && cached.payload && cached.payload.snapshot && cached.payload.snapshot.totalMills > 0);
  if (hasGoodCache) {
    renderHub_(root, buildHubModel_(cached.payload));
  } else {
    renderHubSkeleton_(root);
  }

  if (hasGoodCache && !force) {
    refreshBusy = true;
    fetcher().then(function(payload) {
      const newKey = (payload && payload.dataPeriodKey) || currentPeriodKey_();
      const hasData = payload && payload.snapshot && payload.snapshot.totalMills > 0;
      if (hasData && cached.periodKey !== newKey) {
        writeOverviewCache_(payload);
        renderHub_(root, buildHubModel_(payload));
      } else if (hasData) {
        writeOverviewCache_(payload);
        renderHub_(root, buildHubModel_(payload));
      }
    }).catch(function(e) {
      console.warn('[overview] background refresh', e);
    }).finally(function() {
      refreshBusy = false;
      if (refreshHub_.queuedForce) {
        refreshHub_.queuedForce = false;
        refreshHub_(root, { force: true });
      }
    });
    return;
  }

  refreshBusy = true;
  const btn = root.querySelector('#ovHubRefresh');
  if (btn) btn.disabled = true;
  setOverviewLoading_(root, true);
  let partialRendered = false;
  try {
    if (fastFetcher && !hasGoodCache) {
      try {
        const fastPayload = await fastFetcher();
        if (fastPayload && fastPayload.snapshot && fastPayload.snapshot.totalMills > 0) {
          renderHub_(root, buildHubModel_(fastPayload));
          writeOverviewCache_(fastPayload);
          partialRendered = true;
          setOverviewLoading_(root, false);
          if (btn) {
            btn.disabled = false;
            btn.textContent = 'Updating…';
          }
        }
      } catch (fastErr) {
        console.warn('[overview] fast metrics failed', fastErr);
      }
    }
    const payload = await fetcher();
    const newKey = (payload && payload.dataPeriodKey) || currentPeriodKey_();
    const prevCached = readOverviewCache_();
    const hasData = payload && payload.snapshot && payload.snapshot.totalMills > 0;
    if (!force && prevCached && prevCached.periodKey === newKey && prevCached.payload && hasData) {
      renderHub_(root, buildHubModel_(prevCached.payload));
      return;
    }
    if (hasData || force) {
      writeOverviewCache_(payload);
      renderHub_(root, buildHubModel_(payload));
    } else if (prevCached && prevCached.payload) {
      renderHub_(root, buildHubModel_(prevCached.payload));
    } else if (!partialRendered) {
      renderHub_(root, buildHubModel_(payload || { snapshot: {} }));
    }
  } catch (e) {
    console.warn('[overview]', e);
    const failCached = readOverviewCache_();
    if (failCached && failCached.payload) {
      renderHub_(root, buildHubModel_(failCached.payload));
    } else if (!partialRendered) {
      renderHub_(root, buildHubModel_({ snapshot: {} }));
    }
  } finally {
    refreshBusy = false;
    setOverviewLoading_(root, false);
    if (btn) {
      btn.disabled = false;
      btn.textContent = 'Refresh';
    }
    if (refreshHub_.queuedForce) {
      refreshHub_.queuedForce = false;
      refreshHub_(root, { force: true });
    }
  }
}
refreshHub_.queuedForce = false;

function setOverviewLoading_(root, on) {
  if (!root) return;
  root.classList.toggle('ov-hub--loading', !!on);
  const btn = root.querySelector('#ovHubRefresh');
  if (btn) btn.textContent = on ? 'Loading…' : 'Refresh';
}

function scheduleOverviewFetcherRetry_(root) {
  if (scheduleOverviewFetcherRetry_.timer) return;
  let tries = 0;
  scheduleOverviewFetcherRetry_.timer = setInterval(function() {
    tries += 1;
    if (typeof window.refreshOverviewMetricsData_ === 'function') {
      clearInterval(scheduleOverviewFetcherRetry_.timer);
      scheduleOverviewFetcherRetry_.timer = null;
      refreshHub_(root, { force: true });
      return;
    }
    if (tries >= 120) {
      clearInterval(scheduleOverviewFetcherRetry_.timer);
      scheduleOverviewFetcherRetry_.timer = null;
      setOverviewLoading_(root, false);
    }
  }, 500);
}
scheduleOverviewFetcherRetry_.timer = null;

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
  bindDataPeriodWatch_();

  root.querySelector('#ovHubRefresh').addEventListener('click', function() {
    refreshHub_(root, { force: true });
  });
  window.__overviewMetricsRefresh = function(opts) {
    refreshHub_(root, { force: !!(opts && opts.force) });
  };
  refreshHub_(root, { force: false });
}
