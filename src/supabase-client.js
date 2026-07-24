import { createClient } from '@supabase/supabase-js';

const url = import.meta.env.VITE_SUPABASE_URL;
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

/** @type {import('@supabase/supabase-js').SupabaseClient | null} */
let _client = null;

/**
 * Shared browser client (singleton). Returns null if env is missing — app falls
 * back to demo login in main.js.
 */
export function getSupabase() {
  if (_client) return _client;
  if (!url || !anonKey) return null;
  _client = createClient(String(url).trim(), String(anonKey).trim(), {
    auth: {
      persistSession: true,
      autoRefreshToken: true,
      // Hub SSO passes tokens via /auth-bridge hash/query — we apply them manually.
      // If true, the client may clear the hash before our bridge code runs.
      detectSessionInUrl: false,
    },
  });
  return _client;
}

export function hasSupabaseConfig() {
  return !!(String(url || '').trim() && String(anonKey || '').trim());
}
