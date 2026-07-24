// ── Vite entry point ──────────────────────────────────────────────────────────
// Global error hooks: log only (no modal, no disabling UI — production observability).
(function installGlobalErrorHooks() {
  if (typeof window === 'undefined') return;
  function log(label, detail) {
    if (typeof console !== 'undefined' && console.error) console.error(label, detail);
  }
  window.addEventListener('error', function(ev) {
    log('[app:error]', ev.error || ev.message || ev);
  });
  window.addEventListener('unhandledrejection', function(ev) {
    log('[app:unhandledrejection]', ev.reason);
  });
})();

// CSS + shell first, then heavy app (so #login / overview DOM exists before main binds listeners)
import './style.css';
import './mobile.css';
import './ui-polish.css';

import * as XLSX from 'xlsx-js-style';
window.XLSX = XLSX;

import { mountLoginPage } from './login-ui.js';
import { mountOverviewLanding, updateOverviewWelcomeFromEmail } from './overview-ui.js';
import { getSupabase } from './supabase-client.js';
import {
  allowLocalLogin,
  clearBridgeParamsFromUrl,
  extractBridgeTokens,
  isAuthBridgePath,
  isAuthGateEnabled,
  redirectToHubLogin,
} from './hub-sso.js';

function showHubGateShell_() {
  const dash = document.getElementById('dashboard');
  const login = document.getElementById('login');
  if (dash) {
    dash.classList.remove('active');
    dash.inert = true;
  }
  if (login) {
    login.classList.add('active');
    login.inert = false;
  }
}

function showAppLoadError_(err) {
  console.error('[entry] Failed to load app scripts', err);
  var d = document.createElement('div');
  d.setAttribute('role', 'alert');
  d.style.cssText =
    'position:fixed;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;background:#fef2f2;z-index:2147483647;font:14px system-ui,sans-serif;color:#991b1b;';
  d.textContent =
    'Failed to load application script: ' + (err && err.message ? err.message : String(err));
  document.body.appendChild(d);
}

async function loadAppScripts_() {
  await import('./main.js');
  await import('./modals.js');
}

/**
 * Hub SSO gate runs BEFORE main.js so the dashboard shell cannot open
 * while the heavy bundle is still downloading.
 */
async function boot_() {
  const authGate = isAuthGateEnabled();

  if (authGate) {
    showHubGateShell_();
  }

  // Capture bridge tokens BEFORE any other URL mutation.
  const bridgeTokens = extractBridgeTokens();
  const onBridgePath = isAuthBridgePath();

  // Synchronous mounts — login/overview DOM before main binds listeners.
  mountLoginPage(document.getElementById('login'));
  mountOverviewLanding(document.getElementById('overview-root'));
  window.updateOverviewWelcome = updateOverviewWelcomeFromEmail;

  const SECURE_GAS = import.meta.env.VITE_SECURE_GAS === 'true';
  window.SDD_SECURE_MODE = SECURE_GAS;

  if (!SECURE_GAS) {
    // Dev / legacy only — production must use VITE_SECURE_GAS=true (no GAS URL in browser).
    const SDD_LATEST_WEBAPP_URL =
      'https://script.google.com/macros/s/AKfycbyn7QsagneVRVhfTCls2U1jq5YwRolVXxuE4i9X8vHKuxlzQwwbGAuMjJ8klwnBGidmrQ/exec';
    try {
      localStorage.setItem('SDD_WEBAPP_URL', SDD_LATEST_WEBAPP_URL);
    } catch (e) {
      /* private mode / blocked storage */
    }
    window.SDD_LATEST_WEBAPP_URL = SDD_LATEST_WEBAPP_URL;
    window.SDD_WEBAPP_URL = SDD_LATEST_WEBAPP_URL;
  }

  if (!authGate) {
    if (bridgeTokens || onBridgePath) clearBridgeParamsFromUrl();
    await loadAppScripts_();
    return;
  }

  const sb = getSupabase();
  if (!sb) {
    if (allowLocalLogin()) {
      await loadAppScripts_();
      return;
    }
    redirectToHubLogin();
    return;
  }

  if (bridgeTokens) {
    const statusEl = document.getElementById('hubSsoStatus');
    if (statusEl) statusEl.textContent = 'Signing you in from Hub Portal…';
    const { error } = await sb.auth.setSession({
      access_token: bridgeTokens.access_token,
      refresh_token: bridgeTokens.refresh_token,
    });
    clearBridgeParamsFromUrl();
    if (error) {
      if (statusEl) statusEl.textContent = 'Could not verify Hub session. Redirecting…';
      console.warn('[hub-sso] setSession failed:', error.message || error);
      redirectToHubLogin();
      return;
    }
  } else if (onBridgePath) {
    // Opened /auth-bridge without tokens (plain link / failed Hub handoff).
    clearBridgeParamsFromUrl();
  }

  const res = await sb.auth.getSession();
  if (res.data.session && res.data.session.user) {
    await loadAppScripts_();
    return;
  }

  if (allowLocalLogin()) {
    await loadAppScripts_();
    return;
  }

  const statusEl = document.getElementById('hubSsoStatus');
  if (statusEl) {
    statusEl.textContent = onBridgePath || bridgeTokens
      ? 'Could not verify Hub session. Redirecting…'
      : 'Redirecting to Hub Portal…';
  }
  redirectToHubLogin();
}

boot_().catch(showAppLoadError_);
