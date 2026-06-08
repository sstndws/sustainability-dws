/**
 * Secure GAS access — production memanggil /api/gas-proxy (tanpa Supabase).
 * URL GAS + secret hanya di server Vercel (GAS_WEBAPP_URL, GAS_API_SECRET).
 */

export function isSecureGasEnabled() {
  if (import.meta.env.VITE_SECURE_GAS === 'false') return false;
  if (import.meta.env.VITE_SECURE_GAS === 'true') return true;
  if (typeof location !== 'undefined') {
    const h = String(location.hostname || '').toLowerCase();
    if (h.endsWith('.vercel.app')) return true;
  }
  return false;
}

export function requireSupabaseAuth_() {
  return import.meta.env.VITE_REQUIRE_SUPABASE_AUTH === 'true';
}

/**
 * @param {{ method?: 'GET'|'POST', params?: object, body?: object, timeoutMs?: number }} opts
 */
export async function gasSecureRequest_(opts) {
  const timeoutMs = opts.timeoutMs || 120000;
  const controller = new AbortController();
  const tid = setTimeout(function() { controller.abort(); }, timeoutMs);

  let res;
  try {
    res = await fetch('/api/gas-proxy', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
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
