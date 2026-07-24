# Sustainability Dashboard

Internal dashboard for mill onboarding, supply task lists, supplier due diligence (SDD), declaration / BL monitoring, grievances, monthly reports, and related sustainability workflows.

**Stack:** Vite frontend · Google Sheets via Google Apps Script (GAS) · Vercel (`/api/gas-proxy`) · optional Supabase auth via **Hub Portal SSO**

---

## Features

| Area | Notes |
| --- | --- |
| Mill Onboarding | Mill registry, profiles, waste / CPO / PK supply |
| Supply Task List | Excel import → draft batches → match → submit to Mill (per-row atomic submit + **Repair status**) |
| Supplier DD / Screening | SDD forms, screening, TML / FFB flows |
| Declaration Monitoring | Shipping / BL monitoring, Excel export |
| Monthly Report | Detail / summary PDF |
| Other panels | Grievance, No Buy List, TTM/TTP, EUDR potential, questionnaire, contacts, facility performance |

---

## Quick start

```bash
npm install
cp .env.example .env.local
# Edit .env.local — at least GAS_WEBAPP_URL for local proxy
npm run dev
```

Or: `./dev.sh`

Open the Vite URL (usually `http://127.0.0.1:5173`).

---

## Environment

See [`.env.example`](.env.example).

| Variable | Where | Purpose |
| --- | --- | --- |
| `GAS_WEBAPP_URL` | Vercel + `.env.local` | Apps Script web app `…/exec` URL (**required**) |
| `GAS_API_SECRET` | Vercel (optional) | Must match GAS Script Property `API_SECRET` if enabled |
| `VITE_SECURE_GAS` | Build | `true` in production — browser talks only to `/api/gas-proxy` |
| `VITE_AUTH_ENABLED` | Build | `true` = require Hub SSO session (blocks direct open) |
| `VITE_HUB_PORTAL_URL` | Build | Hub origin, e.g. `https://sustainability-hub-portal-eight.vercel.app` |
| `VITE_HUB_LOGIN_PATH` | Build | Hub login path (default `/login`) |
| `VITE_ALLOW_LOCAL_LOGIN` | Build | Dev only — show local email/password form |
| `VITE_SUPABASE_URL` / `VITE_SUPABASE_ANON_KEY` | Build | Same Supabase project as Hub |

Hub SSO setup: [docs/HUB_SSO.md](docs/HUB_SSO.md).

After a new Apps Script **Deploy → New deployment**, update `GAS_WEBAPP_URL` in Vercel **and** redeploy the project.

---

## Backend (Apps Script)

Canonical script:

```text
scripts/GoogleAppsScript-backend-v3-full.gs
```

1. Paste / sync into the bound Apps Script project for the spreadsheet.
2. **Deploy → Manage deployments → Edit → New version** (or New deployment).
3. Copy the `/exec` URL into `GAS_WEBAPP_URL`.
4. Confirm sheets used by the app exist (`mill`, `millWaste`, `supplyDraft`, SDD sheets, etc.).

Dev tip: without `results[]` from submit, the UI **will not** mark a whole chunk Submitted on partial success — redeploy GAS so per-row results are available.

---

## Deploy (Vercel)

- Root: Vite build (`npm run build` → `dist`)
- Serverless proxy: [`api/gas-proxy.js`](api/gas-proxy.js)
- Config: [`vercel.json`](vercel.json)

Production should set `VITE_SECURE_GAS=true` so the GAS URL stays server-side.

---

## Scripts

```bash
npm run build          # production build
npm run test:supply    # supply merge / routing / mill views
npm run test:ttp       # TTP mill sync
npm run test:period    # period filtering
npm run test:all       # all of the above + build
```

---

## Repo layout

```text
src/                 Frontend (entry, main app, panels UI, PDF/Excel helpers)
partials/            HTML panel shells included into index.html
api/gas-proxy.js     Vercel → GAS proxy
scripts/             Apps Script source + Node test helpers
vite-plugins/        Local GAS proxy for Vite
docs/                Field-mapping notes
```

---

## Supply Task List — important behaviour

- Submit marks **only** rows that succeed (`results[].ok` / `draft_id`).
- Partial failures stay draft; use **Retry failed**.
- If Task List status and Mill Onboarding disagree, use **Repair status** on the batch, then Submit/Retry reopened rows.

---

## License

Private / internal use. Do not publish this repository or redeploy credentials publicly.
