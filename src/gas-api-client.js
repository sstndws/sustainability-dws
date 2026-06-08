/**
 * Secure GAS access — production routes all sheet API calls through /api/gas-proxy
 * with Supabase JWT. GAS URL + API secret never ship to the browser bundle.
 */
import { getSupabase, hasSupabaseConfig } from './supabase-client.js';

export function isSecureGasEnabled() {
  return import.meta.env.VITE_SECURE_GAS === 'true';
}

export function requireSupabaseAuth_() {
  return isSecureGasEnabled() || import.meta.env.VITE_REQUIRE_SUPABASE_AUTH === 'true';
}

export async function getSupabaseAccessToken_() {
  const sb = getSupabase();
  if (!sb) return null;
  const { data, error } = await sb.auth.getSession();
  if (error) return null;
  return data && data.session ? data.session.access_token : null;
}

export async function assertLoggedInForApi_() {
  if (!requireSupabaseAuth_()) return;
  const token = await getSupabaseAccessToken_();
  if (!token) {
    throw new Error('Sesi habis atau belum login. Silakan login ulang untuk mengakses data.');
  }
}

/**
 * @param {{ method?: 'GET'|'POST', params?: object, body?: object, timeoutMs?: number }} opts
 */
export async function gasSecureRequest_(opts) {
  await assertLoggedInForApi_();
  const accessToken = await getSupabaseAccessToken_();
  if (!accessToken) {
    throw new Error('Sesi habis. Silakan login ulang.');
  }

  const timeoutMs = opts.timeoutMs || 120000;
  const controller = new AbortController();
  const tid = setTimeout(function() { controller.abort(); }, timeoutMs);

  let res;
  try {
    res = await fetch('/api/gas-proxy', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + accessToken,
      },
      body: JSON.stringify({
        method: opts.method || 'GET',
        params: opts.params || null,
        body: opts.body || null,
      }),
      signal: controller.signal,
      credentials: 'same-origin',
    });
  } catch (e) {
    clearTimeout(tid);
    if (e && e.name === 'AbortError') {
      throw new Error('Request timeout (' + Math.round(timeoutMs / 1000) + 's).');
    }
    throw new Error('Secure API failed: ' + (e && e.message ? e.message : String(e)));
  }
  clearTimeout(tid);

  const text = await res.text();
  let data;
  try {
    data = text ? JSON.parse(text) : {};
  } catch (parseErr) {
    throw new Error('Secure API: invalid JSON (HTTP ' + res.status + '). ' + text.slice(0, 180));
  }
  if (!res.ok || (data && data.error && !Array.isArray(data))) {
    throw new Error((data && data.error) || ('Secure API HTTP ' + res.status));
  }
  return data;
}
