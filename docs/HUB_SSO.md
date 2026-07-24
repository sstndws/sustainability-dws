# Hub Portal SSO

Sustain Dashboard does **not** own login. Users sign in at the Hub Portal, then open Downstream with a one-time session bridge.

## Flow

1. User opens `https://sustainability-dws-psi.vercel.app/` without a session → redirect to Hub login.
2. User signs in at `https://sustainability-hub-portal-eight.vercel.app/login`.
3. Hub menu opens Sustain with tokens on `/auth-bridge`.
4. Sustain calls `supabase.auth.setSession(...)`, clears the URL, and shows the dashboard.

## Vercel env (Sustain project)

| Variable | Value |
| --- | --- |
| `VITE_AUTH_ENABLED` | `true` |
| `VITE_SUPABASE_URL` | Same project as Hub |
| `VITE_SUPABASE_ANON_KEY` | Same anon key as Hub |
| `VITE_HUB_PORTAL_URL` | `https://sustainability-hub-portal-eight.vercel.app` |
| `VITE_HUB_LOGIN_PATH` | `/login` (optional) |

Redeploy after changing `VITE_*` vars (they are baked in at build time).

## Hub Portal — open Sustain with session

Use the **same** Supabase project. Preferred handoff (server-side):

```text
https://sustainability-hub-portal-eight.vercel.app/launch/sustainability-dashboard
```

Hub reads the logged-in cookie session and redirects to:

```text
https://sustainability-dws-psi.vercel.app/auth-bridge?access_token=...&refresh_token=...
```

App cards with `ssoBridge: true` already link to `/launch/<appId>`.

## Supabase URL Configuration

- **Site URL:** `https://sustainability-hub-portal-eight.vercel.app`
- **Redirect URLs:** Hub + Sustain (+ localhost for dev). Prefer `/**` wildcards on each origin if your Supabase project allows them.

## Local dev

```bash
# .env.local
VITE_AUTH_ENABLED=true
VITE_SUPABASE_URL=...
VITE_SUPABASE_ANON_KEY=...
VITE_HUB_PORTAL_URL=https://sustainability-hub-portal-eight.vercel.app
# Optional: test without Hub
# VITE_ALLOW_LOCAL_LOGIN=true
```
