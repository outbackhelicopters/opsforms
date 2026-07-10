# Heli App — Current State (July 2026)

One repo (`outbackhelicopters/opsforms`), one DigitalOcean app, two components:
static site serves the PWA + admin at `/`; a Node service (`server.js`) handles
`/api/*` (the `/api` prefix is stripped before it reaches Express — every route
is registered under both shapes).

## Pieces
- **flight-ops.html** — pilot iPad PWA. Offline-first (sw.js), SWMS, W&B,
  pax briefing, job sheets, job calendar (week grid: hold a slot to create,
  tap a job to edit). Settings are PIN-protected and hold the device token.
- **admin.html** — office desktop app. Per-user logins (auth.js: provider /
  admin / office roles, invites by email, resets), Today, Calendar
  (click-a-day to create), Paperwork (filters + OneDrive PDFs), Fleet,
  Pilots, Clients, Settings/Users. Dark mode follows the OS.
- **server.js** — PDF build + email + OneDrive filing, calendar with SMS via
  Twilio (2-min supersede buffer, 1-hour reminders, 6pm job-sheet sweep,
  OneDrive scheduler lock), reporting API, auth wiring, device-token gate.
- **auth.js** — users in OneDrive `_system/users.json`; scrypt; signed cookies.
- **config.json** — pilots / aircraft / clients; edit in GitHub (admin UI editing is the next phase).

## Security model
- Office: session cookie (14 days), roles enforced server-side.
- iPads: shared `DEVICE_TOKEN` header on config/send/calendar/drafts —
  enforced only once the env var is set in DigitalOcean.
- Legacy `/reports` + `REPORTS_PASSWORD` still work; retire once the office
  is fully on `/admin.html`.

## Versioning
`APP_VERSION` in flight-ops.html and `CACHE_NAME` in sw.js — bump both
together on every deploy (currently v56).

## Deploying
Upload changed files via GitHub web (Add file → Upload files → commit to main).
Both components auto-deploy. iPads pick up HTML changes on next online open
(network-first); a CACHE_NAME bump purges cached assets.

## Roadmap (agreed)
1. Paperwork ↔ calendar-job matching → "missing paperwork" flags (phase 2)
2. Fleet/pilots/clients editors in admin + config to OneDrive + W&B sign-off
   with audit trail (phase 3 — the CASA-responsibility mechanics)
3. Setup wizard + white-label branding from config + provisioning checklist
   (phase 4 — makes it sellable per-customer; see PLAN-admin-and-commercial.md)
