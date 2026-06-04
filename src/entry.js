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

import * as XLSX from 'xlsx-js-style';
window.XLSX = XLSX;

import { mountLoginPage } from './login-ui.js';
import { mountOverviewLanding, updateOverviewWelcomeFromEmail } from './overview-ui.js';

// Synchronous — must run before main.js binds #btn-login-submit (no async gap before dynamic import).
mountLoginPage(document.getElementById('login'));
mountOverviewLanding(document.getElementById('overview-root'));
window.updateOverviewWelcome = updateOverviewWelcomeFromEmail;

// Selalu pakai deployment terbaru (BL Monitoring butuh script v3-bl-monitoring)
const SDD_LATEST_WEBAPP_URL =
  'https://script.google.com/macros/s/AKfycbyJNQ4wp7AK6a4CdF1EwamWBhQX_hJliuS4g0s08skD-wjWkM_aCHCrr7_YgMJ7du6QhQ/exec';
try {
  localStorage.setItem('SDD_WEBAPP_URL', SDD_LATEST_WEBAPP_URL);
} catch (e) {
  /* private mode / blocked storage */
}
window.SDD_LATEST_WEBAPP_URL = SDD_LATEST_WEBAPP_URL;
window.SDD_WEBAPP_URL = SDD_LATEST_WEBAPP_URL;

import('./main.js')
  .then(() => import('./modals.js'))
  .catch((err) => {
    console.error('[entry] Failed to load app scripts', err);
    var d = document.createElement('div');
    d.setAttribute('role', 'alert');
    d.style.cssText = 'position:fixed;inset:0;display:flex;align-items:center;justify-content:center;padding:24px;background:#fef2f2;z-index:2147483647;font:14px system-ui,sans-serif;color:#991b1b;';
    d.textContent = 'Gagal memuat skrip aplikasi: ' + (err && err.message ? err.message : String(err));
    document.body.appendChild(d);
  });
