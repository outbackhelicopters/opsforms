'use strict';
/* ============================================================
   Outback Helicopter Airwork NT — Flight Paperwork API
   POST /api/send  →  generate PDF, send via Office 365, file to OneDrive
   GET  /reports   →  password-protected reporting dashboard
   GET  /api/jobs  →  job records from OneDrive (auth required)
   All via Microsoft Graph API — one set of credentials for everything
   ============================================================ */

const express     = require('express');
const PDFDocument = require('pdfkit');
const fs          = require('fs');
const path        = require('path');
const crypto      = require('crypto');
const { ClientSecretCredential } = require('@azure/identity');

/* ── Branding (per-customer) ──────────────────────────────────
   Pulled from config.json so this app can be cloned for another
   company by editing config.json only — no code changes. Falls
   back to Outback's own details if config.json has no branding
   block yet, so existing behaviour is unchanged. */
function loadConfigFile() {
  // CONFIG_FILE env var lets a separate deployment (e.g. a sales demo
  // instance) point at a different data file — same code, same repo,
  // just a different JSON of branding/pilots/aircraft/clients. Unset
  // in production, so production always loads config.json unchanged.
  const configName = process.env.CONFIG_FILE || 'config.json';
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, configName), 'utf8')); }
  catch (e) { console.error(`${configName} load failed:`, e.message); return {}; }
}
const APP_CONFIG = loadConfigFile();
const BRAND = Object.assign({
  companyName: 'OUTBACK HELICOPTER AIRWORK NT PTY LTD',
  shortName:   'Outback Helicopter Airwork NT',
  shortest:    'Outback Helicopter',
  abn:         '80 137 947 687',
  acn:         '137 947 687',
  address:     'PO Box 37819 Winnellie NT 0821',
  phone:       'Ph: 8941 6811 | Mob: 0427 222 670',
  location:    'Darwin, NT'
}, APP_CONFIG.branding || {});

/* ── Reports auth ─────────────────────────────────────────── */
const REPORTS_PWD    = process.env.REPORTS_PASSWORD || '';
const REPORTS_SECRET = process.env.REPORTS_SECRET   || ('ohant-' + (process.env.MS_CLIENT_ID || 'key').slice(0,12));

function parseCookies(req) {
  const out = {};
  (req.headers.cookie || '').split(';').forEach(c => {
    const i = c.indexOf('=');
    if (i < 0) return;
    out[c.slice(0,i).trim()] = decodeURIComponent(c.slice(i+1).trim());
  });
  return out;
}
function makeToken(pwd) {
  return crypto.createHmac('sha256', REPORTS_SECRET).update(pwd).digest('hex').slice(0,32);
}
let authApi = null; // set below once helpers exist — per-user auth for the admin app
async function requireReportsAuth(req, res, next) {
  const cookies = parseCookies(req);
  if (REPORTS_PWD && cookies.rpt_auth === makeToken(REPORTS_PWD)) return next();
  // Admin-app session (per-user) also grants access to reports/jobs
  if (authApi) {
    try { if (await authApi.sessionUser(req)) return next(); } catch (e) { console.error('session check:', e.message); }
  }
  if (req.path.startsWith('/api/')) return res.status(401).json({ ok:false, error:'Unauthorised' });
  res.redirect('/reports/login');
}

/* ── Graph token helper ───────────────────────────────────── */
async function getGraphToken() {
  const cred = new ClientSecretCredential(
    process.env.MS_TENANT_ID,
    process.env.MS_CLIENT_ID,
    process.env.MS_CLIENT_SECRET
  );
  const { token } = await cred.getToken('https://graph.microsoft.com/.default');
  return token;
}

/* ── In-memory jobs cache (5 min) ────────────────────────── */
let _jobsCache = null;
let _jobsCacheAt = 0;
const CACHE_TTL = 5 * 60 * 1000;

/* ── ACST helpers (Australia/Darwin = UTC+9:30, no DST) ─────── */
const ACST_TZ = 'Australia/Darwin';
function acstDate(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, day: '2-digit', month: 'long', year: 'numeric' }).format(new Date(ts)); }
function acstTime(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(ts)); }
function acstFull(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(ts)); }

const app = express();
app.set('trust proxy', true); // DigitalOcean sits in front of this — needed so req.ip is the real caller, not the load balancer
app.use(express.json({ limit: '20mb' }));

/* ── Device token (iPads) ─────────────────────────────────────
   When DEVICE_TOKEN is set, the endpoints the iPads use require
   either that token (x-device-token header) or a signed-in admin
   session. While DEVICE_TOKEN is unset, requests pass untouched —
   so the fleet can have the token entered in Settings BEFORE
   enforcement is switched on in DigitalOcean. */
async function requireDevice(req, res, next) {
  const expected = process.env.DEVICE_TOKEN;
  if (!expected) return next();
  const got = String(req.headers['x-device-token'] || '');
  if (got.length === expected.length &&
      crypto.timingSafeEqual(Buffer.from(got), Buffer.from(expected))) return next();
  if (authApi) {
    try { if (await authApi.sessionUser(req)) return next(); } catch (e) { console.error('device auth session check:', e.message); }
  }
  res.status(401).json({ ok: false, error: 'Device token required — enter it in Settings on this iPad' });
}

/* ── Rate limiter ──────────────────────────────────────────────
   Calendar writes fan out to SMS, which costs money and can be
   disruptive if it fires a lot — this caps how often one IP can
   hit those routes so a bug or abuse can't run up a Twilio bill
   or spam every pilot. In-memory only, fine at this app's scale. */
const RATE_LIMIT_WINDOW_MS = 60 * 1000;
const RATE_LIMIT_MAX = 20; // per IP per minute — generous for real dispatch use, tight enough to stop a runaway loop
const _rateHits = new Map();

function rateLimit(req, res, next) {
  const ip = req.ip || 'unknown';
  const now = Date.now();
  const hits = (_rateHits.get(ip) || []).filter(t => now - t < RATE_LIMIT_WINDOW_MS);
  if (hits.length >= RATE_LIMIT_MAX) {
    return res.status(429).json({ ok: false, error: 'Too many requests — please wait a minute and try again' });
  }
  hits.push(now);
  _rateHits.set(ip, hits);
  next();
}

/* ── Serve static app files (flight-ops.html, sw.js, etc.) ──
   The deployed repo (opsforms) is flat — server.js sits next to
   flight-ops.html, admin.html etc. Local dev copies of this project
   sometimes nest the API under api/, one level below those files.
   Detect which layout is in play the same way the /admin route
   below already does, instead of assuming one or the other. */
const STATIC_DIR = fs.existsSync(path.join(__dirname, 'flight-ops.html'))
  ? __dirname
  : path.join(__dirname, '..');
app.use(express.static(STATIC_DIR));

/* ── Shared config (pilots, aircraft, clients) ─────────────────
   Served from OneDrive-backed LIVE_OPS once it's loaded (see the
   "Live operational data" block below) — falls back to the repo's
   api/config.json only until that first load completes. */
app.get(['/config', '/api/config'], requireDevice, (_req, res) => {
  res.json({ ...APP_CONFIG, ...LIVE_OPS, branding: { ...BRAND } });
});

/* ── Health check ─────────────────────────────────────────── */
app.get('/health', (_req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

/* ── Reports: login page ──────────────────────────────────── */
app.get('/reports/login', (_req, res) => {
  res.sendFile(path.join(__dirname, 'reports.html'));
});
app.post('/reports/login', express.json(), (req, res) => {
  const { password } = req.body || {};
  if (!REPORTS_PWD) return res.json({ ok:false, error:'Not configured' });
  if (password !== REPORTS_PWD) return res.json({ ok:false, error:'Wrong password' });
  const token = makeToken(REPORTS_PWD);
  res.setHeader('Set-Cookie', `rpt_auth=${token}; Path=/; HttpOnly; Max-Age=2592000; SameSite=Strict`);
  res.json({ ok: true });
});

/* ── Reports: dashboard ───────────────────────────────────── */
app.get('/reports', requireReportsAuth, (_req, res) => {
  res.sendFile(path.join(__dirname, 'reports.html'));
});

/* ── Jobs API (for reporting dashboard) ───────────────────── */
app.get(['/jobs', '/api/jobs'], requireReportsAuth, async (req, res) => {
  // Prevent any upstream proxy/CDN from caching this response by URL —
  // confirmed via testing that identical-URL GETs to this route were being
  // served a stale cached body even though our own in-memory cache logic
  // was correct and a query-varied URL (e.g. ?refresh=1&_=timestamp)
  // reliably returned fresh, correct data every time.
  res.set('Cache-Control', 'no-store');
  try {
    if (!hasGraphCreds()) {
      const jobs = readLocalJson('records', []);
      return res.json({ jobs, total: jobs.length });
    }

    const forceRefresh = req.query.refresh === '1';
    const now = Date.now();
    if (!forceRefresh && _jobsCache && now - _jobsCacheAt < CACHE_TTL) {
      return res.json({ jobs: _jobsCache, cached: true });
    }

    const token = await getGraphToken();
    const { jobs, fileCount } = await fetchAllJobRecords(token);

    // Only cache the result if it looks trustworthy: either the folder
    // genuinely has no files, or we actually got records back. If Graph
    // listed files but every download failed (transient throttling/cold
    // start), don't poison the cache with a false empty result — let the
    // next request try again live instead of serving 0 for up to 5 min.
    if (fileCount === 0 || jobs.length > 0) {
      _jobsCache   = jobs;
      _jobsCacheAt = now;
    }
    res.json({ jobs, total: jobs.length });
  } catch (err) {
    console.error('Jobs fetch error:', err.message);
    res.status(500).json({ ok:false, error: err.message });
  }
});

/* ── Fetch every job record from OneDrive _records/ ─────────────
   Shared by the /api/jobs reporting endpoint and by the per-pilot
   job-counter history seeding below — both need the exact same
   real submission history, fetched the same reliable way. */
async function fetchAllJobRecords(token) {
  const driveUser  = process.env.OPS_EMAIL;
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const recPath    = encodeURIComponent(`${folderName}/_records`);

  // NOTE: we deliberately do NOT $select=@microsoft.graph.downloadUrl here.
  // On OneDrive-for-Business/SharePoint-backed drives, Graph silently omits
  // that field from /children listings even when explicitly selected — it
  // only reliably appears on a per-item GET. Confirmed by direct testing:
  // every listed record here had @odata.etag + name only. So instead we
  // grab each item's id from the listing and download its content via
  // /drive/items/{id}/content, which works regardless of drive type.
  let files = [];
  let url = `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${recPath}:/children`
          + `?$select=id,name&$top=1000`;

  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (r.status === 404) break;
    if (!r.ok) throw new Error(`List records: ${r.status}`);
    const d = await r.json();
    files.push(...(d.value || []).filter(f => f.name && f.name.endsWith('.json')));
    url = d['@odata.nextLink'] || null;
  }

  // Download all records in parallel batches of 20
  const jobs = [];
  for (let i = 0; i < files.length; i += 20) {
    const batch = files.slice(i, i + 20);
    const results = await Promise.all(batch.map(async f => {
      for (let attempt = 0; attempt < 2; attempt++) {
        try {
          const r = await fetch(
            `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/items/${f.id}/content`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          if (r.ok) return await r.json();
        } catch { /* retry */ }
      }
      return null;
    }));
    jobs.push(...results.filter(Boolean));
  }
  return { jobs, fileCount: files.length };
}

/* ── Generic OneDrive JSON helpers (used by calendar + drafts) ─ */
async function listOneDriveJsonFiles(token, folderPath) {
  const driveUser = process.env.OPS_EMAIL;
  const encPath = encodeURIComponent(folderPath);
  let files = [];
  let url = `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${encPath}:/children`
          + `?$select=name,@microsoft.graph.downloadUrl&$top=1000`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (r.status === 404) break;
    if (!r.ok) throw new Error(`List ${folderPath}: ${r.status}`);
    const d = await r.json();
    files.push(...(d.value || []).filter(f => f.name && f.name.endsWith('.json')));
    url = d['@odata.nextLink'] || null;
  }
  const out = [];
  for (let i = 0; i < files.length; i += 20) {
    const batch = files.slice(i, i + 20);
    const results = await Promise.all(batch.map(async f => {
      try {
        const r = await fetch(f['@microsoft.graph.downloadUrl']);
        return r.ok ? await r.json() : null;
      } catch { return null; }
    }));
    out.push(...results.filter(Boolean));
  }
  return out;
}

async function putOneDriveJson(token, filePath, data) {
  const driveUser = process.env.OPS_EMAIL;
  const r = await fetch(
    `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${encodeURIComponent(filePath)}:/content`,
    { method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(data) }
  );
  if (!r.ok) throw new Error(`Save ${filePath}: ${r.status} ${await r.text()}`);
}

async function getOneDriveJson(token, filePath) {
  const driveUser = process.env.OPS_EMAIL;
  const r = await fetch(
    `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${encodeURIComponent(filePath)}:/content`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  if (r.status === 404) return null;
  if (!r.ok) throw new Error(`Read ${filePath}: ${r.status}`);
  return r.json();
}

/* ── Admin app: per-user auth (see auth.js) ───────────────── */
authApi = require('./auth')(app, { getGraphToken, getOneDriveJson, putOneDriveJson, parseCookies, rateLimit, BRAND });

/* ── Live branding (white-label) ──────────────────────────────
   Branding is layered, later wins:
     1. hardcoded defaults (BRAND above)
     2. config.json "branding" block (merged at boot)
     3. OneDrive _system/branding.json — editable from the admin
        setup wizard / Settings with no redeploy.
   The logo works the same way: repo logo.png is the fallback,
   an uploaded one lives in OneDrive and is cached in memory. */
const BRANDING_LOCAL = path.join(__dirname, '_branding.local.json');
const LOGO_LOCAL     = path.join(__dirname, '_brand-logo.local.png');
const hasGraphCreds  = () => !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
const BRAND_FOLDER   = () => process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
let _brandLogoBuf = null;

async function loadLiveBranding() {
  try {
    let data = null;
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      data = await getOneDriveJson(token, `${BRAND_FOLDER()}/_system/branding.json`);
    } else if (fs.existsSync(BRANDING_LOCAL)) {
      data = JSON.parse(fs.readFileSync(BRANDING_LOCAL, 'utf8'));
    }
    if (data && typeof data === 'object') Object.assign(BRAND, data);
  } catch (e) { console.error('branding load failed (using defaults):', e.message); }
}
async function loadBrandLogo() {
  try {
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      const r = await fetch(
        `https://graph.microsoft.com/v1.0/users/${process.env.OPS_EMAIL}/drive/root:/${encodeURIComponent(BRAND_FOLDER() + '/_system/brand-logo.png')}:/content`,
        { headers: { Authorization: `Bearer ${token}` } });
      if (r.ok) _brandLogoBuf = Buffer.from(await r.arrayBuffer());
    } else if (fs.existsSync(LOGO_LOCAL)) {
      _brandLogoBuf = fs.readFileSync(LOGO_LOCAL);
    }
  } catch (e) { console.error('brand logo load failed (using repo logo):', e.message); }
}
loadLiveBranding();
loadBrandLogo();

async function saveLiveBranding(patch) {
  Object.assign(BRAND, patch);
  const snapshot = { ...BRAND };
  if (hasGraphCreds()) {
    const token = await getGraphToken();
    await putOneDriveJson(token, `${BRAND_FOLDER()}/_system/branding.json`, snapshot);
  } else {
    fs.writeFileSync(BRANDING_LOCAL, JSON.stringify(snapshot, null, 2));
  }
}

/* ── Live operational data (pilots / aircraft / clients) ───────
   Same layering as branding above:
     1. api/config.json in the repo — a seed, read only the very
        first time a deployment boots against an empty OneDrive.
     2. OneDrive _system/config.json — the live, per-customer copy.
   Once OneDrive holds a copy, the repo file is never consulted
   again for this data. That means every deployment's pilots,
   aircraft and client list live entirely inside that customer's
   own OneDrive — fully separate from every other deployment and
   from the repo itself, and editable with no redeploy. */
const OPCONFIG_LOCAL = path.join(__dirname, '_opconfig.local.json');

/* ── Generic local JSON store ──────────────────────────────────
   Same fallback pattern as branding/ops-config/audit above, extended
   to calendar jobs, drafts, the job-number counter and job records —
   the pieces that previously called getGraphToken() unconditionally
   and would 500 with no Graph creds set. Used automatically whenever
   hasGraphCreds() is false (e.g. a sales-demo deployment with
   CONFIG_FILE=config.demo.json and no MS_* env vars): data lives in
   flat JSON files next to server.js instead of OneDrive, so the app
   is genuinely interactive with no external integrations, at the
   cost of not surviving a redeploy — fine for a demo, never used in
   a real customer deployment since those always carry Graph creds. */
function localJsonPath(name) { return path.join(__dirname, `_${name}.local.json`); }
function readLocalJson(name, fallback) {
  try { return JSON.parse(fs.readFileSync(localJsonPath(name), 'utf8')); }
  catch (e) { return fallback; }
}
function writeLocalJson(name, data) {
  try { fs.writeFileSync(localJsonPath(name), JSON.stringify(data, null, 2)); }
  catch (e) { console.error(`local store write failed (${name}):`, e.message); }
}

let LIVE_OPS = {
  pilots:   Array.isArray(APP_CONFIG.pilots)   ? APP_CONFIG.pilots   : [],
  aircraft: Array.isArray(APP_CONFIG.aircraft) ? APP_CONFIG.aircraft : [],
  clients:  Array.isArray(APP_CONFIG.clients)  ? APP_CONFIG.clients  : [],
};
async function loadLiveOpsConfig() {
  try {
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      const opPath = `${BRAND_FOLDER()}/_system/config.json`;
      const data = await getOneDriveJson(token, opPath);
      if (data && typeof data === 'object') {
        const rawPilots = Array.isArray(data.pilots) ? data.pilots : [];
        const codedPilots = assignPilotCodes(rawPilots);
        LIVE_OPS = {
          pilots:   codedPilots,
          aircraft: Array.isArray(data.aircraft) ? data.aircraft : [],
          clients:  Array.isArray(data.clients)  ? data.clients  : [],
        };
        // A pilot without a code just got one for the first time (e.g. an
        // existing customer upgrading into per-pilot job numbering) — save
        // it back so the code is stable from here on, not recomputed.
        if (codedPilots.some((p, i) => p.code !== (rawPilots[i] && rawPilots[i].code))) {
          await putOneDriveJson(token, opPath, LIVE_OPS);
        }
      } else {
        // First boot for this deployment — seed OneDrive from the repo
        // copy so nothing breaks, then the repo file stops mattering.
        LIVE_OPS.pilots = assignPilotCodes(LIVE_OPS.pilots);
        await putOneDriveJson(token, opPath, LIVE_OPS);
      }
    } else if (fs.existsSync(OPCONFIG_LOCAL)) {
      const data = JSON.parse(fs.readFileSync(OPCONFIG_LOCAL, 'utf8'));
      const rawPilots = Array.isArray(data.pilots) ? data.pilots : LIVE_OPS.pilots;
      const codedPilots = assignPilotCodes(rawPilots);
      LIVE_OPS = {
        pilots:   codedPilots,
        aircraft: Array.isArray(data.aircraft) ? data.aircraft : LIVE_OPS.aircraft,
        clients:  Array.isArray(data.clients)  ? data.clients  : LIVE_OPS.clients,
      };
      if (codedPilots.some((p, i) => p.code !== (rawPilots[i] && rawPilots[i].code))) {
        fs.writeFileSync(OPCONFIG_LOCAL, JSON.stringify(LIVE_OPS, null, 2));
      }
    } else {
      LIVE_OPS.pilots = assignPilotCodes(LIVE_OPS.pilots);
    }
  } catch (e) { console.error('live config load failed (using repo config.json):', e.message); }
}
/* Aircraft carry a nested W&B block. Numbers only — POH sign-off (name,
   ARN, date, drawn signature) is a deliberate separate step a customer's
   own admin/chief pilot takes via POST /setup/aircraft-signoff (see
   PLAN-admin-and-commercial.md). Nothing entered here or via the setup
   wizard/Fleet editor is ever auto-verified: editing any of these numbers
   carries a prior sign-off forward ONLY if every figure is byte-identical
   to what was last signed — any real change voids it and a fresh sign-off
   is required, so the audit trail always matches what pilots are flying
   on. */
function numOrZero(v) { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; }
function wbNumbersMatch(a, b) {
  const eq = (x, y) => Math.abs((+x || 0) - (+y || 0)) < 1e-9;
  return eq(a.emptyWeight, b.emptyWeight) && eq(a.emptyLongArm, b.emptyLongArm) &&
    eq(a.emptyLatArm, b.emptyLatArm) && eq(a.mtow, b.mtow) && eq(a.fuelDensity, b.fuelDensity) &&
    eq(a.maxFuelL, b.maxFuelL) &&
    (a.source || '') === (b.source || '') && (a.cgEnvKey || '') === (b.cgEnvKey || '') &&
    JSON.stringify(a.accessories || []) === JSON.stringify(b.accessories || []);
}
function normalizeAircraft(list, prevList) {
  const prevByReg = new Map((Array.isArray(prevList) ? prevList : []).map(a => [a.reg, a]));
  return (Array.isArray(list) ? list : []).map(a => {
    const wIn = (a && a.wb) || {};
    const reg = String((a && a.reg) || '').trim().toUpperCase();
    const wb = {
      source:       String(wIn.source || '').trim(),
      emptyWeight:  numOrZero(wIn.emptyWeight),
      emptyLongArm: numOrZero(wIn.emptyLongArm),
      emptyLatArm:  numOrZero(wIn.emptyLatArm),
      mtow:         numOrZero(wIn.mtow),
      fuelDensity:  numOrZero(wIn.fuelDensity) || 0.720,
      maxFuelL:     numOrZero(wIn.maxFuelL), // usable fuel capacity for this specific tail number — overrides the pilot app's generic per-type table when set (0 = not entered, falls back to the type default)
      cgEnvKey:     String(wIn.cgEnvKey || '').trim(), // matches a built-in CG envelope preset, if any
      accessories: (Array.isArray(wIn.accessories) ? wIn.accessories : [])
        .map(x => ({ name: String((x && x.name) || '').trim(), weight: numOrZero(x && x.weight), arm: numOrZero(x && x.arm) }))
        .filter(x => x.name),
    };
    const prev = prevByReg.get(reg);
    const carriesOver = !!(prev && prev.wb && prev.wb.verified && wbNumbersMatch(prev.wb, wb));
    wb.verified = carriesOver;
    wb.signOff  = carriesOver ? prev.wb.signOff : null;
    return { reg, type: String((a && a.type) || '').trim(), wb };
  }).filter(a => a.reg);
}

/* Append-only audit log — who changed what, when. Never blocks a save if
   the write itself fails (e.g. OneDrive hiccup); logged, not thrown. */
const AUDIT_LOCAL = path.join(__dirname, '_audit.local.json');
async function appendAudit(entry) {
  try {
    let list = [];
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      list = (await getOneDriveJson(token, `${BRAND_FOLDER()}/_system/audit-log.json`)) || [];
      if (!Array.isArray(list)) list = [];
      list.push(entry);
      if (list.length > 5000) list = list.slice(-5000);
      await putOneDriveJson(token, `${BRAND_FOLDER()}/_system/audit-log.json`, list);
    } else {
      if (fs.existsSync(AUDIT_LOCAL)) { try { list = JSON.parse(fs.readFileSync(AUDIT_LOCAL, 'utf8')); } catch (_) { list = []; } }
      if (!Array.isArray(list)) list = [];
      list.push(entry);
      if (list.length > 5000) list = list.slice(-5000);
      fs.writeFileSync(AUDIT_LOCAL, JSON.stringify(list, null, 2));
    }
  } catch (e) { console.error('audit log write failed (non-fatal):', e.message); }
}

async function saveLiveOpsConfig(patch, who) {
  if (Array.isArray(patch.pilots))   LIVE_OPS.pilots   = assignPilotCodes(normalizeRosterPilots(patch.pilots));
  if (Array.isArray(patch.aircraft)) {
    const prevAircraft = LIVE_OPS.aircraft;
    const nextAircraft = normalizeAircraft(patch.aircraft, prevAircraft);
    nextAircraft.forEach(next => {
      const prev = prevAircraft.find(p => p.reg === next.reg);
      if (prev && prev.wb && prev.wb.verified && !next.wb.verified) {
        appendAudit({ ts: new Date().toISOString(), who: who || null, action: 'wb-edit-voided-signoff', reg: next.reg });
      } else if (!prev) {
        appendAudit({ ts: new Date().toISOString(), who: who || null, action: 'aircraft-added', reg: next.reg });
      }
    });
    LIVE_OPS.aircraft = nextAircraft;
  }
  if (Array.isArray(patch.clients))  LIVE_OPS.clients  = patch.clients
    .map(c => String(c || '').trim()).filter(Boolean);
  const snapshot = { ...LIVE_OPS };
  if (hasGraphCreds()) {
    const token = await getGraphToken();
    await putOneDriveJson(token, `${BRAND_FOLDER()}/_system/config.json`, snapshot);
  } else {
    fs.writeFileSync(OPCONFIG_LOCAL, JSON.stringify(snapshot, null, 2));
  }
}
loadLiveOpsConfig();

/* Setup endpoints are open until the first account exists, then admin-only */
async function requireSetupAuth(req, res, next) {
  try {
    if (!(await authApi.usersExist())) return next();
    const u = await authApi.sessionUser(req);
    if (u && (u.role === 'provider' || u.role === 'admin')) return next();
  } catch (e) { console.error('setup auth check:', e.message); }
  res.status(401).json({ ok: false, error: 'Admin sign-in required' });
}

app.get(['/setup/status', '/api/setup/status'], async (_req, res) => {
  const out = {
    env: {
      microsoft:     hasGraphCreds(),
      senderEmail:   !!process.env.SENDER_EMAIL,
      opsEmail:      !!process.env.OPS_EMAIL,
      sessionSecret: !!process.env.SESSION_SECRET,
      twilio:        !!(process.env.TWILIO_ACCOUNT_SID && process.env.TWILIO_AUTH_TOKEN && process.env.TWILIO_FROM_NUMBER),
      deviceToken:   !!process.env.DEVICE_TOKEN,
      anthropic:     !!process.env.ANTHROPIC_API_KEY,
    },
    graphOk: false, oneDriveOk: false, graphError: null,
  };
  if (out.env.microsoft) {
    try {
      const token = await getGraphToken();
      out.graphOk = true;
      try {
        await putOneDriveJson(token, `${BRAND_FOLDER()}/_system/setup-ping.json`, { ts: new Date().toISOString() });
        out.oneDriveOk = true;
      } catch (e) { out.graphError = 'OneDrive write failed: ' + e.message; }
    } catch (e) { out.graphError = e.message; }
  }
  res.json({ ok: true, ...out, branding: { ...BRAND } });
});

app.post(['/setup/branding', '/api/setup/branding'], rateLimit, requireSetupAuth, async (req, res) => {
  try {
    const b = req.body || {};
    const patch = {};
    ['companyName','shortName','shortest','brandLine1','brandLine2','brandLine3',
     'logoAlt','abn','acn','address','phone','location','opsEmail'].forEach(k => {
      if (typeof b[k] === 'string' && b[k].trim()) patch[k] = b[k].trim();
    });
    if (b.colors && typeof b.colors === 'object') {
      const hex = v => typeof v === 'string' && /^#[0-9a-fA-F]{6}$/.test(v.trim());
      const colors = {};
      if (hex(b.colors.primary)) colors.primary = b.colors.primary.trim();
      if (hex(b.colors.accent))  colors.accent  = b.colors.accent.trim();
      if (Object.keys(colors).length) patch.colors = { ...(BRAND.colors || {}), ...colors };
    }
    if (typeof b.logoDataUrl === 'string' && b.logoDataUrl) {
      const m = b.logoDataUrl.match(/^data:image\/(png|jpe?g);base64,([A-Za-z0-9+/=]+)$/);
      if (!m) return res.status(400).json({ ok: false, error: 'Logo must be a PNG or JPEG image' });
      const buf = Buffer.from(m[2], 'base64');
      if (buf.length > 1.5 * 1024 * 1024) return res.status(400).json({ ok: false, error: 'Logo too big — keep it under 1.5 MB' });
      _brandLogoBuf = buf;
      if (hasGraphCreds()) {
        const token = await getGraphToken();
        const r = await fetch(
          `https://graph.microsoft.com/v1.0/users/${process.env.OPS_EMAIL}/drive/root:/${encodeURIComponent(BRAND_FOLDER() + '/_system/brand-logo.png')}:/content`,
          { method: 'PUT', headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'image/png' }, body: buf });
        if (!r.ok) throw new Error(`Logo save failed: ${r.status}`);
      } else {
        fs.writeFileSync(LOGO_LOCAL, buf);
      }
    }
    await saveLiveBranding(patch);
    res.json({ ok: true, branding: { ...BRAND } });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});

/* Edit pilots / aircraft / clients — writes straight to this
   customer's OneDrive, no redeploy, no touching the repo. Send any
   one or more of pilots/aircraft/clients; anything omitted is left
   untouched. Same gating as branding: open only until the first
   account exists, then provider/admin only. */
app.put(['/setup/config', '/api/setup/config'], rateLimit, requireSetupAuth, async (req, res) => {
  try {
    const b = req.body || {};
    if (!Array.isArray(b.pilots) && !Array.isArray(b.aircraft) && !Array.isArray(b.clients))
      return res.status(400).json({ ok: false, error: 'Send pilots, aircraft and/or clients as arrays' });
    const u = await authApi.sessionUser(req);
    await saveLiveOpsConfig(b, u ? { name: u.name, email: u.email, role: u.role } : null);
    res.json({ ok: true, pilots: LIVE_OPS.pilots, aircraft: LIVE_OPS.aircraft, clients: LIVE_OPS.clients });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});

/* ── W&B sign-off ─────────────────────────────────────────────
   The ONLY place `verified` is ever set to true. Deliberately
   restricted to the customer's own admin (chief pilot/HOFO) —
   never the provider — so CASA/regulatory responsibility for the
   figures being flown on stays on the customer's side, per
   PLAN-admin-and-commercial.md. Requires a name, an ARN, a date and
   a drawn signature; all four are stored with the sign-off and
   logged to the audit trail. */
async function requireAdminSignOffAuth(req, res, next) {
  try {
    const u = await authApi.sessionUser(req);
    if (u && u.role === 'admin') { req._signOffUser = u; return next(); }
    if (u && u.role === 'provider') return res.status(403).json({ ok: false, error: "Providers can't sign off W&B data — this has to come from the customer's own admin/chief pilot" });
    return res.status(401).json({ ok: false, error: 'Admin sign-in required' });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
}
app.post(['/setup/aircraft-signoff', '/api/setup/aircraft-signoff'], rateLimit, requireAdminSignOffAuth, async (req, res) => {
  try {
    const { reg, name, arn, date, sigDataUrl } = req.body || {};
    const R = String(reg || '').trim().toUpperCase();
    const ac = LIVE_OPS.aircraft.find(a => a.reg === R);
    if (!ac) return res.status(404).json({ ok: false, error: 'Aircraft not found — save the fleet first' });
    if (!String(name || '').trim())      return res.status(400).json({ ok: false, error: 'Name is required' });
    if (!String(arn || '').trim())       return res.status(400).json({ ok: false, error: 'ARN is required' });
    if (!String(date || '').trim())      return res.status(400).json({ ok: false, error: 'Date is required' });
    if (!sigDataUrl || !/^data:image\//.test(sigDataUrl)) return res.status(400).json({ ok: false, error: 'Signature is required' });
    ac.wb.verified = true;
    ac.wb.signOff = {
      name: String(name).trim(), arn: String(arn).trim(), date: String(date).trim(),
      sigDataUrl, signedAt: new Date().toISOString(),
    };
    const snapshot = { ...LIVE_OPS };
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      await putOneDriveJson(token, `${BRAND_FOLDER()}/_system/config.json`, snapshot);
    } else {
      fs.writeFileSync(OPCONFIG_LOCAL, JSON.stringify(snapshot, null, 2));
    }
    await appendAudit({
      ts: new Date().toISOString(),
      who: { name: req._signOffUser.name, email: req._signOffUser.email, role: req._signOffUser.role },
      action: 'wb-signoff', reg: R, arn: String(arn).trim(),
    });
    res.json({ ok: true, aircraft: LIVE_OPS.aircraft });
  } catch (e) { res.status(500).json({ ok: false, error: e.message }); }
});

/* ── Scan an aircraft weighing report ────────────────────────────
   Office uploads a photo or PDF of the weighing report; Claude reads it
   and hands back the numbers so the form fills itself in — the office
   still has to check every value before saving, and W&B still only
   goes live once someone signs it off against the POH, same as always.
   Nothing here bypasses that; it just removes the retyping. */
app.post(['/setup/scan-aircraft', '/api/setup/scan-aircraft'], rateLimit, requireSetupAuth, async (req, res) => {
  try {
    const { dataUrl, mimeType } = req.body || {};
    if (!dataUrl) return res.status(400).json({ ok: false, error: 'No file received' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(400).json({ ok: false, error: "Document scanning isn't set up yet — ANTHROPIC_API_KEY is missing in DigitalOcean" });

    const m = /^data:([^;]+);base64,(.+)$/.exec(dataUrl);
    if (!m) return res.status(400).json({ ok: false, error: 'Could not read that file' });
    const mediaType = mimeType || m[1];
    const base64 = m[2];
    const isPdf = mediaType === 'application/pdf';
    if (!isPdf && !mediaType.startsWith('image/')) {
      return res.status(400).json({ ok: false, error: 'Upload a PDF or a photo (JPG/PNG)' });
    }

    const content = [
      { type: isPdf ? 'document' : 'image', source: { type: 'base64', media_type: mediaType, data: base64 } },
      {
        type: 'text',
        text: 'This is an aircraft weight & balance / weighing report. Read it and return ONLY a JSON object ' +
          '(no other text, no markdown fences) with these exact keys: reg (registration, string or null), ' +
          'type (aircraft type, string or null), emptyWeight (kg, number or null), emptyLongArm (mm, number or null), ' +
          'emptyLatArm (mm, number or null), mtow (kg, number or null), fuelDensity (kg/L, number or null), ' +
          'maxFuelL (usable fuel capacity in litres, number or null). ' +
          "If a value isn't clearly on the document, use null — never guess or estimate a number.",
      },
    ];

    const r = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      body: JSON.stringify({ model: 'claude-sonnet-4-5', max_tokens: 1024, messages: [{ role: 'user', content }] }),
    });
    if (!r.ok) {
      console.error('Claude scan failed:', r.status, await r.text().catch(() => ''));
      return res.status(502).json({ ok: false, error: 'Scan failed — try a clearer photo or a PDF' });
    }
    const data = await r.json();
    const text = (data.content || []).map(b => b.text || '').join('').trim();
    let parsed;
    try {
      const jsonMatch = text.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(jsonMatch ? jsonMatch[0] : text);
    } catch (e) {
      console.error('Could not parse scan response:', text);
      return res.status(502).json({ ok: false, error: "Couldn't read numbers from that document — try again or enter manually" });
    }
    res.json({ ok: true, data: parsed });
  } catch (err) {
    console.error('scan-aircraft error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.get(['/brand-logo', '/api/brand-logo'], (_req, res) => {
  res.setHeader('Cache-Control', 'public, max-age=300');
  if (_brandLogoBuf) return res.type('png').send(_brandLogoBuf);
  res.sendFile(path.join(__dirname, 'logo.png'), err => { if (err) res.status(404).end(); });
});

app.get(['/admin', '/api/admin'], (_req, res) => {
  const flat = path.join(__dirname, 'admin.html');           // deployed layout (flat repo)
  res.sendFile(fs.existsSync(flat) ? flat : path.join(__dirname, '..', 'admin.html'));
});

/* ── Calendar jobs cache ──────────────────────────────────── */
let _calCache = null, _calCacheAt = 0;
const CAL_CACHE_TTL = 60 * 1000; // 1 min — calendar should feel close to live

async function loadCalendarJobs(force) {
  const now = Date.now();
  if (!force && _calCache && now - _calCacheAt < CAL_CACHE_TTL) return _calCache;
  let jobs;
  if (hasGraphCreds()) {
    const token = await getGraphToken();
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    jobs = await listOneDriveJsonFiles(token, `${folderName}/_calendar`);
  } else {
    jobs = readLocalJson('calendar', []);
  }
  _calCache = jobs; _calCacheAt = now;
  return jobs;
}

/* ── Phone number normalizer ─────────────────────────────────
   Pilots can be entered in config.json in whatever format is
   natural — "0412 345 678", "(0412) 345-678", "61412345678",
   "+61 412 345 678" — and this converts it to the E.164 shape
   Twilio needs (+61412345678) before it's ever stored or sent.
   Defaults to Australian numbers (leading 0 → +61); anything
   already starting with + is assumed correct and just cleaned up.
   ============================================================ */
function normalizeAuPhone(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const hasPlus = s.startsWith('+');
  const digits = s.replace(/\D/g, '');
  if (!digits) return '';
  if (hasPlus)              return '+' + digits;          // already international
  if (digits.startsWith('61') && digits.length > 9) return '+' + digits; // "61412345678"
  if (digits.startsWith('0')) return '+61' + digits.slice(1); // "0412 345 678"
  if (digits.length === 9)   return '+61' + digits;        // "412 345 678"
  return '+' + digits;                                      // best effort fallback
}

/* ── Pilot list helpers ───────────────────────────────────────
   Calendar jobs store pilots as an array: [{ name, phone, email }].
   normalizePilots() cleans up incoming request bodies (and accepts
   the old single-pilot shape for backward compatibility), converting
   every phone number to E.164 on the way in. jobPilots() reads
   pilots off any job record, old or new shape. ─────────────── */
function normalizePilots(input, legacyBody) {
  let list = Array.isArray(input) ? input : [];
  if (!list.length && legacyBody && legacyBody.pilotName) {
    list = [{ name: legacyBody.pilotName, phone: legacyBody.pilotPhone, email: legacyBody.pilotEmail }];
  }
  return list
    .map(p => ({ name: String((p && p.name) || '').trim(), phone: normalizeAuPhone((p && p.phone) || ''), email: String((p && p.email) || '').trim() }))
    .filter(p => p.name);
}

/* ── Pilot codes for per-pilot job numbering ─────────────────────
   normalizeRosterPilots() is like normalizePilots() above but for the
   master Fleet/Pilot roster specifically (not a one-off job's pilot
   list) — it keeps an existing `code` field intact so a pilot's short
   code never changes just because the office re-saves Settings.
   assignPilotCodes() then fills in a code for any pilot that doesn't
   have one yet: first letter of first name + first letter of last
   name, uppercased. If two pilots would collide (e.g. Angus Watson
   and Anton Williams both "AW"), whoever doesn't already have that
   code gets "AW2", "AW3", etc. Existing codes are never recomputed,
   so a pilot's prefix — and therefore their run of job numbers —
   stays stable even if the roster is edited or reordered later. */
function normalizeRosterPilots(input) {
  return normalizePilots(input).map((p, i) => {
    const raw = (Array.isArray(input) ? input[i] : null) || {};
    const code = String(raw.code || '').trim().toUpperCase();
    return code ? { ...p, code } : p;
  });
}
function assignPilotCodes(pilots) {
  const list = Array.isArray(pilots) ? pilots : [];
  const used = new Set(list.filter(p => p.code).map(p => p.code));
  return list.map(p => {
    if (p.code) return p;
    const parts = String(p.name || '').trim().split(/\s+/).filter(Boolean);
    const base = (((parts[0] || '')[0] || 'X') + ((parts[parts.length - 1] || '')[0] || 'X')).toUpperCase();
    let code = base, n = 2;
    while (used.has(code)) { code = base + n; n++; }
    used.add(code);
    return { ...p, code };
  });
}
/* Looks up a pilot's stable code from the current roster by name (case/
   whitespace-insensitive). Falls back to deriving one on the fly for a
   name that isn't in the roster (typo, pilot removed after their jobs
   were logged, etc.) so job-number issuance never hard-fails — just
   isn't guaranteed collision-free against the real roster in that rare
   case, same tradeoff the client's offline fallback already accepts. */
function pilotCodeForName(name) {
  const key = String(name || '').trim().toLowerCase();
  const match = (LIVE_OPS.pilots || []).find(p => String(p.name || '').trim().toLowerCase() === key);
  if (match && match.code) return match.code;
  const parts = key.split(/\s+/).filter(Boolean);
  return (((parts[0] || '')[0] || 'X') + ((parts[parts.length - 1] || '')[0] || 'X')).toUpperCase();
}
function jobPilots(job) {
  if (Array.isArray(job.pilots) && job.pilots.length) return job.pilots;
  if (job.pilotName) return [{ name: job.pilotName, phone: job.pilotPhone || '', email: job.pilotEmail || '' }];
  return [];
}

/* ── 2-minute buffer before any calendar-job text goes out ─────
   Created / changed / cancelled all schedule their text 2 minutes
   out instead of sending straight away. If the same pilot on the
   same job gets another notice-worthy edit inside that window (a
   typo fix, a second change of mind, etc.), the earlier pending
   text is dropped and only the latest one survives — so a pilot
   never gets a burst of texts for a job still being fiddled with.
   In-memory only: a redeploy inside the 2-minute window drops any
   text still pending, same as the rest of this app's fire-and-
   forget notifications.
   ============================================================ */
const NOTICE_DELAY_MS = 2 * 60 * 1000;
const _pendingNotices = new Map(); // "<jobId>::<pilotName>" → setTimeout handle

function scheduleNotice(jobId, pilotName, sendFn) {
  const key = `${jobId}::${pilotName}`;
  const prior = _pendingNotices.get(key);
  if (prior) clearTimeout(prior);
  const timer = setTimeout(async () => {
    _pendingNotices.delete(key);
    try { await sendFn(); } catch (e) { console.error('Delayed notice failed:', key, e.message); }
  }, NOTICE_DELAY_MS);
  _pendingNotices.set(key, timer);
}

/* ── Draft job sheets ────────────────────────────────────────
   Once a calendar job is logged (the 6pm sweep has run for it),
   each pilot has a draft job sheet waiting in _drafts/. If the
   job is then edited or cancelled, these keep that draft in step:
     - createDraftForPilot()  same shape the 6pm sweep creates —
       reused here for a pilot added to an already-logged job.
     - findJobDrafts()        drafts for a job, optionally one pilot.
     - updateJobDrafts()      patch date/client on drafts NOT YET
       pulled onto a device — once pulled, the app has no way to
       reach into a specific iPad's storage, so this only helps
       for drafts still sitting on the server.
     - cancelJobDrafts()      same limit — marks not-yet-pulled
       drafts cancelled so they're never handed to a device. A
       pilot who already has the draft is told by SMS instead.
   ============================================================ */
async function createDraftForPilot(token, job, pilot) {
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const draftId = crypto.randomUUID();
  const draft = {
    id:            draftId,
    calendarJobId: job.id,
    jobNo:         'CAL-' + draftId.slice(0, 5).toUpperCase(),
    date:          job.date,
    client:        job.client || '',
    pilotName:     pilot.name,
    hireType:      'wet',
    lines:         [],
    totalHours:    0,
    notes:         job.description || job.notes || '',
    status:        'draft',
    createdAt:     Date.now(),
    pulled:        false,
  };
  await putOneDriveJson(token, `${folderName}/_drafts/${draftId}.json`, draft);
  return draft;
}

async function findJobDrafts(token, jobId, pilotName) {
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const drafts = await listOneDriveJsonFiles(token, `${folderName}/_drafts`);
  return drafts.filter(d => d.calendarJobId === jobId && d.status === 'draft' && (!pilotName || d.pilotName === pilotName));
}

async function updateJobDrafts(token, jobId, patch, onlyPilotNames) {
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const drafts = await findJobDrafts(token, jobId);
  for (const d of drafts) {
    if (d.pulled) continue;
    if (onlyPilotNames && !onlyPilotNames.includes(d.pilotName)) continue;
    try { await putOneDriveJson(token, `${folderName}/_drafts/${d.id}.json`, { ...d, ...patch }); }
    catch (e) { console.error('Draft sync failed:', d.id, e.message); }
  }
}

async function cancelJobDrafts(token, jobId, pilotName) {
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const drafts = await findJobDrafts(token, jobId, pilotName);
  for (const d of drafts) {
    if (d.pulled) continue;
    try { await putOneDriveJson(token, `${folderName}/_drafts/${d.id}.json`, { ...d, status: 'cancelled' }); }
    catch (e) { console.error('Draft cancel failed:', d.id, e.message); }
  }
}

/* ── Calendar jobs API (pilot scheduling — allocate pilots, not aircraft) ─
   GET    /api/calendar-jobs         list (optionally ?from=&to=&all=1)
   POST   /api/calendar-jobs         create { date, startTime, pilots:[{name,phone,email}], client, description, notes }
   PATCH  /api/calendar-jobs/:id     edit
   DELETE /api/calendar-jobs/:id     cancel (soft delete)
   ============================================================ */
app.get(['/calendar-jobs', '/api/calendar-jobs'], requireDevice, async (req, res) => {
  try {
    const force = req.query.refresh === '1';
    let jobs = await loadCalendarJobs(force);
    if (req.query.from) jobs = jobs.filter(j => j.date >= req.query.from);
    if (req.query.to)   jobs = jobs.filter(j => j.date <= req.query.to);
    if (req.query.all !== '1') jobs = jobs.filter(j => j.status !== 'cancelled');
    jobs.sort((a, b) => (a.date + (a.startTime || '')).localeCompare(b.date + (b.startTime || '')));
    res.json({ jobs });
  } catch (err) {
    console.error('calendar-jobs list error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.post(['/calendar-jobs', '/api/calendar-jobs'], requireDevice, rateLimit, async (req, res) => {
  try {
    const b = req.body || {};
    const pilots = normalizePilots(b.pilots, b);
    if (!b.date || !pilots.length) return res.status(400).json({ ok: false, error: 'date and at least one pilot are required' });
    const id = crypto.randomUUID();
    const record = {
      id,
      date:            b.date,
      startTime:       b.startTime || '',
      endTime:         b.endTime   || '',
      pilots,
      client:          b.client || '',
      description:     b.description || '',
      notes:           b.notes || '',
      createdBy:       b.createdBy || '',
      status:          'scheduled',   // scheduled → logged | cancelled
      remindedPilots:  [],            // names already sent the 1-hour reminder
      loggedAt:        null,
      createdAt:       new Date().toISOString(),
    };
    if (hasGraphCreds()) {
      const token = await getGraphToken();
      const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
      await putOneDriveJson(token, `${folderName}/_calendar/${id}.json`, record);
    } else {
      const jobs = readLocalJson('calendar', []);
      jobs.push(record);
      writeLocalJson('calendar', jobs);
    }
    _calCache = null;
    res.json({ ok: true, job: record });

    // Text every allocated pilot — held for 2 minutes in case this gets fixed or cancelled right after
    for (const p of pilots) scheduleNotice(id, p.name, () => notifyJobAssigned(record, p));
  } catch (err) {
    console.error('calendar-jobs create error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.patch(['/calendar-jobs/:id', '/api/calendar-jobs/:id'], requireDevice, rateLimit, async (req, res) => {
  try {
    const token = hasGraphCreds() ? await getGraphToken() : null;
    const jobs = await loadCalendarJobs(true);
    const existing = jobs.find(j => j.id === req.params.id);
    if (!existing) return res.status(404).json({ ok: false, error: 'Not found' });
    const b = req.body || {};
    const oldPilots = jobPilots(existing);
    const newPilots = b.pilots !== undefined ? normalizePilots(b.pilots) : oldPilots;
    const updated = {
      ...existing,
      ...(b.date        !== undefined ? { date: b.date }               : {}),
      ...(b.startTime   !== undefined ? { startTime: b.startTime }      : {}),
      ...(b.endTime      !== undefined ? { endTime: b.endTime }          : {}),
      ...(b.pilots       !== undefined ? { pilots: newPilots }           : {}),
      ...(b.client       !== undefined ? { client: b.client }            : {}),
      ...(b.description  !== undefined ? { description: b.description } : {}),
      ...(b.notes        !== undefined ? { notes: b.notes }              : {}),
      ...(b.status       !== undefined ? { status: b.status }            : {}),
      updatedAt: new Date().toISOString(),
    };
    delete updated.pilotName; delete updated.pilotPhone; delete updated.pilotEmail; // migrated to pilots[]

    const timeChanged = (b.date !== undefined && b.date !== existing.date) ||
                        (b.startTime !== undefined && b.startTime !== existing.startTime);
    // Editing the date/time of a still-scheduled job re-arms the 1-hour reminder
    if (timeChanged && updated.status === 'scheduled') updated.remindedPilots = [];

    const oldNames = new Set(oldPilots.map(p => p.name));
    const newNames = new Set(newPilots.map(p => p.name));
    const removedPilots = oldPilots.filter(p => !newNames.has(p.name));
    const addedPilots   = newPilots.filter(p => !oldNames.has(p.name));
    const keptPilots    = newPilots.filter(p => oldNames.has(p.name));

    if (hasGraphCreds()) {
      const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
      await putOneDriveJson(token, `${folderName}/_calendar/${req.params.id}.json`, updated);
    } else {
      const local = readLocalJson('calendar', []);
      const idx = local.findIndex(j => j.id === req.params.id);
      if (idx >= 0) local[idx] = updated; else local.push(updated);
      writeLocalJson('calendar', local);
    }
    _calCache = null;
    res.json({ ok: true, job: updated });

    // A logged job already has a draft job sheet per pilot sitting in _drafts/ —
    // keep it in step with the edit, wherever the draft still is on the server
    const wasLogged = existing.status === 'logged';

    // Notify pilot(s) — held for 2 minutes so a quick follow-up edit can supersede it
    if (updated.status === 'scheduled' || wasLogged) {
      for (const p of removedPilots) scheduleNotice(req.params.id, p.name, () => notifyJobRemoved(existing, p, wasLogged));   // off the job
      for (const p of addedPilots)   scheduleNotice(req.params.id, p.name, () => notifyJobAssigned(updated, p));              // newly on the job
      if (timeChanged) for (const p of keptPilots) scheduleNotice(req.params.id, p.name, () => notifyJobChanged(existing, updated, p, wasLogged)); // same pilots, date/start time moved
    }

    // Draft job-sheet sync only applies when the drafts pipeline is Graph-backed
    // (the 6pm sweep that creates them never runs without Graph creds either — see runScheduler)
    if (wasLogged && hasGraphCreds()) {
      // Off the job — cancel their draft if it's still sitting on the server, unpulled
      for (const p of removedPilots) cancelJobDrafts(token, req.params.id, p.name).catch(e => console.error('draft cancel failed:', p.name, e.message));
      // Newly added to an already-logged job — they missed the 6pm sweep, so start their draft now
      for (const p of addedPilots) createDraftForPilot(token, updated, p).catch(e => console.error('late draft create failed:', p.name, e.message));
      // Date/client changed on kept pilots — sync onto their draft if it hasn't been pulled yet
      if (timeChanged || (b.client !== undefined && b.client !== existing.client)) {
        updateJobDrafts(token, req.params.id, {
          ...(timeChanged ? { date: updated.date } : {}),
          ...(b.client !== undefined ? { client: updated.client } : {}),
        }, keptPilots.map(p => p.name)).catch(e => console.error('draft sync failed:', e.message));
      }
    }
  } catch (err) {
    console.error('calendar-jobs update error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

app.delete(['/calendar-jobs/:id', '/api/calendar-jobs/:id'], requireDevice, rateLimit, async (req, res) => {
  try {
    const token = hasGraphCreds() ? await getGraphToken() : null;
    const jobs = await loadCalendarJobs(true);
    const existing = jobs.find(j => j.id === req.params.id);
    if (!existing) return res.status(404).json({ ok: false, error: 'Not found' });
    const wasScheduled = existing.status === 'scheduled';
    const wasLogged    = existing.status === 'logged';
    const updated = { ...existing, status: 'cancelled', cancelledAt: new Date().toISOString() };
    if (hasGraphCreds()) {
      const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
      await putOneDriveJson(token, `${folderName}/_calendar/${req.params.id}.json`, updated);
    } else {
      const local = readLocalJson('calendar', []);
      const idx = local.findIndex(j => j.id === req.params.id);
      if (idx >= 0) local[idx] = updated; else local.push(updated);
      writeLocalJson('calendar', local);
    }
    _calCache = null;
    res.json({ ok: true });

    // Only notify if the job was still upcoming or just logged — a job cancelled twice, or one
    // that's already cancelled/done, doesn't need another text
    if (wasScheduled || wasLogged) for (const p of jobPilots(updated)) scheduleNotice(req.params.id, p.name, () => notifyJobCancelled(updated, p, wasLogged));

    // If a draft job sheet already exists for this job, cancel any copy still sitting on the
    // server (unpulled) — a pilot who already has it on their iPad is told by SMS not to submit it
    if (wasLogged && hasGraphCreds()) cancelJobDrafts(token, req.params.id).catch(e => console.error('draft cancel failed:', e.message));
  } catch (err) {
    console.error('calendar-jobs delete error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── Per-pilot Job Advice Sheet number counters ──────────────────
   Each pilot gets their own consecutive sequence (DC-0001, DC-0002...)
   instead of sharing one pool with every other pilot — previously a
   single shared counter meant Danny's and Chris's numbers were
   interleaved and looked random from either one's point of view.
   Counters are keyed by the pilot's stable code (see assignPilotCodes
   above), stored together in one OneDrive file, cached in memory once
   loaded, incremented under an in-process lock so two submits landing
   in the same instant still come out as two different numbers. This
   only serializes writes within this one running instance — fine at
   this app's scale, same tradeoff already accepted by the drafts lock
   below.
   The first time this runs for a deployment with no counters file yet,
   each pilot's starting count is seeded from their real submission
   history in _records/ (existing jobNo values) so numbers continue on
   sensibly instead of resetting everyone to 1. ── */
let JOB_COUNTERS = null; // { counters: { CODE: n } } — null until first loaded
let _jobNoLockChain = Promise.resolve();
function withJobNoLock(fn) {
  const result = _jobNoLockChain.then(fn, fn);
  _jobNoLockChain = result.then(() => {}, () => {});
  return result;
}
async function seedJobCountersFromHistory(token) {
  const counters = {};
  try {
    const { jobs } = await fetchAllJobRecords(token);
    for (const rec of jobs) {
      if (!rec || !rec.jobNo) continue; // only real Job Advice Sheet submissions carry a jobNo
      const code = rec.jobCode || pilotCodeForName(rec.pilotName);
      counters[code] = (counters[code] || 0) + 1;
    }
  } catch (e) { console.error('job counter history scan failed (starting from zero):', e.message); }
  return counters;
}
async function loadJobCounters(token) {
  if (JOB_COUNTERS !== null) return JOB_COUNTERS;
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  try {
    const data = await getOneDriveJson(token, `${folderName}/_system/job-counters.json`);
    if (data && data.counters && typeof data.counters === 'object') {
      JOB_COUNTERS = { counters: data.counters };
      return JOB_COUNTERS;
    }
  } catch (e) { /* fall through to seed */ }
  JOB_COUNTERS = { counters: await seedJobCountersFromHistory(token) };
  return JOB_COUNTERS;
}

/* POST /api/job-number/next — atomically hands out the next Job Advice
   Sheet number for the given pilot. Only call this once, right at the
   moment a pilot actually submits (not while they're still filling the
   form in) — every call consumes a number, even if the submit is later
   abandoned. Body: { pilotName }. */
app.post(['/job-number/next', '/api/job-number/next'], requireDevice, rateLimit, async (req, res) => {
  try {
    const pilotName = String((req.body && req.body.pilotName) || '').trim();
    const result = await withJobNoLock(async () => {
      const code = pilotCodeForName(pilotName);
      if (hasGraphCreds()) {
        const token = await getGraphToken();
        const counters = await loadJobCounters(token);
        counters.counters[code] = (counters.counters[code] || 0) + 1;
        const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
        await putOneDriveJson(token, `${folderName}/_system/job-counters.json`, { counters: counters.counters, updatedAt: new Date().toISOString() });
        return { number: counters.counters[code], code };
      }
      const stored = readLocalJson('jobcounters', { counters: {} });
      stored.counters[code] = (stored.counters[code] || 0) + 1;
      writeLocalJson('jobcounters', stored);
      return { number: stored.counters[code], code };
    });
    res.json({ ok: true, number: result.number, code: result.code });
  } catch (err) {
    console.error('job-number/next error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── Serialize draft claims within this process ───────────────
   Two iPads opening the Jobs tab at almost the same instant could
   both read the same "unclaimed" draft before either had written
   pulled:true, and both walk away thinking they own it. Chaining
   every /api/job-drafts request through one promise queue means
   the second request's read always happens after the first
   request's writes have landed, so it correctly sees the draft as
   already claimed. This only covers a single process — if this
   app is ever run as more than one instance, pair it with the
   scheduler lock below or move claiming to a real datastore. ── */
let _draftsClaimChain = Promise.resolve();
function withDraftsLock(fn) {
  const result = _draftsClaimChain.then(fn, fn);
  _draftsClaimChain = result.then(() => {}, () => {});
  return result;
}

/* ── Pending job-sheet drafts (created by the 6pm daily sweep) ─
   GET /api/job-drafts[?pilot=Name]  — fetch-and-claim pending drafts
   ============================================================ */
app.get(['/job-drafts', '/api/job-drafts'], requireDevice, rateLimit, async (req, res) => {
  try {
    const pending = await withDraftsLock(async () => {
      const pilot = req.query.pilot;
      if (hasGraphCreds()) {
        const token = await getGraphToken();
        const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
        const drafts = await listOneDriveJsonFiles(token, `${folderName}/_drafts`);
        const claim = drafts.filter(d => !d.pulled && d.status === 'draft' && (!pilot || d.pilotName === pilot));
        for (const d of claim) {
          try {
            await putOneDriveJson(token, `${folderName}/_drafts/${d.id}.json`, { ...d, pulled: true, pulledAt: new Date().toISOString() });
          } catch (e) { console.error('draft pull-flag failed:', e.message); }
        }
        return claim;
      }
      // Local mode: the 6pm sweep that creates drafts never runs without Graph creds
      // (see runScheduler), so this is always an empty list — kept for API shape parity.
      const drafts = readLocalJson('drafts', []);
      const claim = drafts.filter(d => !d.pulled && d.status === 'draft' && (!pilot || d.pilotName === pilot));
      claim.forEach(d => { d.pulled = true; d.pulledAt = new Date().toISOString(); });
      if (claim.length) writeLocalJson('drafts', drafts);
      return claim;
    });
    res.json({ drafts: pending });
  } catch (err) {
    console.error('job-drafts error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── Local PDF serving (demo mode only — real deployments file to OneDrive) ─
   /send saves the built PDF to _local-pdfs/ and points the job record's
   oneDriveUrl at this route instead, so the admin Paperwork screen's
   "View filed PDF" link works the same way it does against a real OneDrive URL. */
app.get(['/local-pdf/:filename', '/api/local-pdf/:filename'], requireReportsAuth, (req, res) => {
  const safe = String(req.params.filename || '').replace(/[^a-zA-Z0-9_.-]/g, '');
  const filePath = path.join(__dirname, '_local-pdfs', safe);
  res.sendFile(filePath, err => { if (err) res.status(404).json({ ok: false, error: 'Not found' }); });
});

/* ── Send bundle ──────────────────────────────────────────── */
app.post(['/send', '/api/send'], requireDevice, async (req, res) => {
  const bundle = req.body;
  if (!bundle || !bundle.callsign) return res.status(400).json({ ok: false, error: 'Invalid bundle' });

  try {
    /* 1 — Generate PDF */
    const pdfBuffer = await buildPDF(bundle);

    const safeForm = (bundle.formName || 'form').replace(/[^a-zA-Z0-9_-]/g, '_');
    const dateStr  = new Date(bundle.queuedAt || Date.now()).toISOString().slice(0, 10);

    /* Job Advice Sheet PDFs (identified by bundle.client, same gate used in
       buildPDF's JOB DETAILS section) are named:
         <Jobsheet #> <Pilot Initials> <Rego suffix> <Date DD.MM.YYYY>.pdf
       e.g. "A009 PB RYT 20.08.2026.pdf" — office staff asked for this exact
       layout so filed PDFs sort and scan the same way their paper job sheets
       always did. Other forms (SWMS etc, no job number) keep the original
       Rego_Form_Date name. */
    let filename;
    if (bundle.client) {
      const jobNoLabel = 'A' + String(bundle.jobNo || 0).padStart(3, '0');
      const pilotCode  = bundle.jobCode || pilotCodeForName(bundle.sms?.values?.pilotName || bundle.pilotName || '');
      const regoSuffix = String(bundle.callsign || 'UNK').replace(/^VH-?/i, '') || 'UNK';
      const fdSource   = bundle.flightDate || dateStr; // YYYY-MM-DD
      const [fy, fm, fd] = String(fdSource).split('-');
      const dateForName = (fy && fm && fd) ? `${fd}.${fm}.${fy}` : fdSource;
      filename = `${jobNoLabel} ${pilotCode} ${regoSuffix} ${dateForName}`.replace(/[\\/:*?"<>|]/g, '_') + '.pdf';
    } else {
      filename = `${bundle.callsign}_${safeForm}_${dateStr}.pdf`;
    }

    /* No Graph creds (e.g. a demo deployment) → skip OneDrive/email entirely and
       keep the PDF + job record on local disk instead. Same fallback pattern used
       everywhere else in this file when hasGraphCreds() is false. */
    const demoMode = !hasGraphCreds();

    /* Get Microsoft Graph access token (shared for email + OneDrive) */
    const token = demoMode ? null : await getGraphToken();

    /* 2 — File to OneDrive (or a local folder, served back via /api/local-pdf) */
    let oneDriveUrl = null;
    if (demoMode) {
      try {
        const localDir = path.join(__dirname, '_local-pdfs');
        if (!fs.existsSync(localDir)) fs.mkdirSync(localDir, { recursive: true });
        const safeName = `${Date.now()}_${filename}`.replace(/[^a-zA-Z0-9_.-]/g, '_');
        fs.writeFileSync(path.join(localDir, safeName), pdfBuffer);
        oneDriveUrl = `/api/local-pdf/${safeName}`;
      } catch (err) {
        console.error('Local PDF save failed (non-fatal):', err.message);
      }
    } else {
      try {
        oneDriveUrl = await uploadToOneDrive(token, pdfBuffer, filename, bundle.callsign, dateStr.slice(0, 7));
        console.log('OneDrive:', oneDriveUrl);
      } catch (err) {
        console.error('OneDrive upload failed (non-fatal):', err.message);
      }
    }

    /* 3 — Send email via Office 365 */
    const pilot   = bundle.sms?.values?.pilotName || 'Unknown pilot';
    const trainer = bundle.sms?.values?.trainerName || '';
    const formName = bundle.formName || bundle.sms?.formId || 'Flight Operations Form';
    const formNo   = bundle.formNo   || '';
    /* Defect reports (red ! button on the SWMS, or the standalone MDR form) */
    const isDefect = bundle.kind === 'defect' || bundle.sms?.formId === 'form-mdr';
    const dv = bundle.sms?.values || {};
    const defStatus = dv.grounded ? 'AIRCRAFT GROUNDED' : (dv.deferred ? 'Deferred' : 'Unserviceability');
    const subject = isDefect
      ? `⚠ DEFECT REPORT — ${bundle.callsign} — ${defStatus} — ${dv.systemAffected || 'Defect'} — ${dateStr}`
      : `Flight Paperwork — ${bundle.callsign} — ${formName} — ${dateStr}`;
    const sender  = process.env.SENDER_EMAIL;
    const opsTo   = process.env.OPS_EMAIL;

    /* ── W&B ── */
    const wb = bundle.wb || {};
    const wbResult = wb.result || {};
    const fuelDens = wb.fuelDensity || 0.720;
    const fuelType = fuelDens >= 0.79 ? 'Jet A-1' : 'AvGas 100LL';
    const fuelL    = wb.fuelL || 0;
    const fuelKg   = wb.kg?.fuel || +(fuelL * fuelDens).toFixed(1);

    /* ── Shared inline style strings ── */
    const S = {
      wrap:     'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;max-width:620px;margin:0 auto;background:#ffffff;',
      hdr:      'background:linear-gradient(135deg,#0E1835 0%,#1E2F5C 60%,#243669 100%);padding:28px 36px;',
      coName:   'font-size:18px;font-weight:800;color:#ffffff;letter-spacing:.2px;margin:0;',
      coSub:    'font-size:12px;color:rgba(255,255,255,.5);margin:3px 0 0;font-weight:500;',
      badge:    'display:inline-block;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.9);border-radius:5px;padding:4px 11px;font-size:11px;font-weight:700;letter-spacing:.5px;text-transform:uppercase;float:right;margin-top:4px;',
      body:     'padding:28px 36px 32px;',
      title:    'font-size:20px;font-weight:800;color:#18224A;margin:0 0 3px;',
      sub:      'font-size:13px;color:#888;font-weight:500;margin:0 0 24px;',
      sec:      'border:1px solid #E8EBF2;border-radius:9px;overflow:hidden;margin-bottom:20px;',
      secHd:    'background:#F2F4FA;border-bottom:1px solid #E8EBF2;padding:8px 16px;font-size:10.5px;font-weight:800;color:#18224A;letter-spacing:.6px;text-transform:uppercase;',
      kv:       'display:table;width:100%;padding:7px 16px;border-bottom:1px solid #F2F4FA;box-sizing:border-box;',
      kvLast:   'display:table;width:100%;padding:7px 16px;box-sizing:border-box;',
      k:        'display:table-cell;font-size:12px;font-weight:600;color:#888;width:150px;padding-right:12px;vertical-align:top;padding-top:1px;',
      v:        'display:table-cell;font-size:13px;color:#1C1F28;font-weight:500;vertical-align:top;',
      ackOk:    'background:#F0FDF4;border-radius:6px;padding:7px 11px;margin:3px 0;display:table;width:100%;box-sizing:border-box;',
      ackFail:  'background:#FEF2F2;border-radius:6px;padding:7px 11px;margin:3px 0;display:table;width:100%;box-sizing:border-box;',
      ackIcon:  'display:table-cell;width:24px;font-size:13px;vertical-align:middle;',
      ackNum:   'display:table-cell;width:24px;font-size:12px;font-weight:700;vertical-align:middle;',
      ackLab:   'display:table-cell;font-size:12.5px;font-weight:500;vertical-align:middle;',
      sumOk:    'background:#DCFCE7;border:1px solid #BBF7D0;border-radius:6px;padding:8px 13px;font-size:12px;font-weight:700;color:#15803D;margin:8px 0 4px;',
      sumWarn:  'background:#FEF3C7;border:1px solid #FDE68A;border-radius:6px;padding:8px 13px;font-size:12px;font-weight:700;color:#92400E;margin:8px 0 4px;',
      footer:   'background:#F7F8FA;border-top:1px solid #E8EBF2;padding:14px 36px;font-size:11.5px;color:#aaa;',
    };

    /* ── Builder helpers ── */
    const kv = (k, v, last) =>
      `<div style="${last ? S.kvLast : S.kv}"><span style="${S.k}">${k}</span><span style="${S.v}">${v}</span></div>`;
    const section = (icon, title, inner) =>
      `<div style="${S.sec}"><div style="${S.secHd}">${icon}&nbsp; ${title}</div>${inner}</div>`;
    const ackItem = (ok, num, label) =>
      `<div style="${ok ? S.ackOk : S.ackFail};color:${ok ? '#15803D' : '#B91C1C'};">` +
      `<span style="${S.ackIcon}">${ok ? '✅' : '❌'}</span>` +
      `<span style="${S.ackNum}">${num}.</span>` +
      `<span style="${S.ackLab}">${label}</span></div>`;

    /* ── W&B result pill ── */
    const wbPill = wbResult.pass != null
      ? (wbResult.pass
        ? `<span style="display:inline-block;background:#DCFCE7;color:#15803D;border:1px solid #BBF7D0;border-radius:20px;padding:3px 12px;font-size:12.5px;font-weight:700;">✅ PASS &nbsp;·&nbsp; ${wbResult.total} kg / ${wbResult.mtow} kg MTOW</span>`
        : `<span style="display:inline-block;background:#FEE2E2;color:#B91C1C;border:1px solid #FECACA;border-radius:20px;padding:3px 12px;font-size:12.5px;font-weight:700;">❌ OVER MTOW — ${wbResult.total} kg exceeds ${wbResult.mtow} kg MTOW</span>`)
      : 'Not recorded';

    /* ── SWMS acknowledgments block ── */
    const ackLabels = bundle.smsAckLabels || [];
    const acks      = bundle.sms?.acks || {};
    let acksSection = '';
    if (ackLabels.length) {
      const allAcked = ackLabels.every(a => a.acked);
      const ackItems = ackLabels.map((a, i) => ackItem(a.acked, i + 1, esc(a.step))).join('');
      const summary  = allAcked
        ? `<div style="${S.sumOk}">✅ All ${ackLabels.length} sections read, understood and acknowledged by pilot — signed.</div>`
        : `<div style="${S.sumWarn}">⚠ Not all sections acknowledged — review required before filing.</div>`;
      acksSection = section('📋', 'Safety Management — Section Acknowledgments',
        `<div style="padding:10px 16px 14px;">` +
        `<div style="font-size:12px;color:#666;margin-bottom:8px;">Pilot confirms they have read, understood and will comply with each section below.</div>` +
        ackItems + summary + `</div>`);
    } else if (Object.keys(acks).length) {
      const ackItems = Object.entries(acks).map(([i, v]) => ackItem(v, parseInt(i) + 1, `Section ${parseInt(i) + 1}`)).join('');
      acksSection = section('📋', 'Safety Management — Section Acknowledgments',
        `<div style="padding:10px 16px 14px;">${ackItems}</div>`);
    }

    /* ── Passengers block ── */
    const pax = bundle.pax || [];
    const paxInner = pax.length
      ? pax.map((p, i) => {
          const sigImg = p.sig
            ? `<div style="margin-top:6px;"><img src="${p.sig}" style="height:46px;border:1px solid #E8EBF2;border-radius:5px;display:block;" alt="Passenger signature"></div>`
            : `<div style="font-size:11.5px;color:#B91C1C;font-weight:600;margin-top:4px;">❌ No signature recorded</div>`;
          return `<div style="${i === pax.length - 1 ? S.kvLast : S.kv}">` +
            `<span style="${S.k}"><b>${i + 1}. ${esc(p.name)}</b></span>` +
            `<span style="${S.v}">` +
            (p.weight ? `<span style="color:#555;">${p.weight} kg</span>` : '') +
            (p.date ? `<span style="color:#aaa;font-size:11.5px;"> &nbsp;·&nbsp; Briefed ${esc(p.date)} ${esc(p.time || '')}</span>` : '') +
            `<br><span style="font-size:12px;font-weight:700;color:${p.sig ? '#15803D' : '#B91C1C'};">${p.sig ? '✅ Signed' : '❌ Not signed'}</span>` +
            sigImg +
            `</span></div>`;
        }).join('')
      : `<div style="padding:12px 16px;font-size:13px;color:#aaa;font-style:italic;">No passengers carried this flight</div>`;

    /* ── Defect blocks ── */
    const nl2br = t => esc(t || '').replace(/\n/g, '<br>');
    let defectSection = '';
    if (isDefect) {
      const pill = (bg, fg, bd, t) => `<span style="display:inline-block;background:${bg};color:${fg};border:1px solid ${bd};border-radius:20px;padding:3px 12px;font-size:12.5px;font-weight:700;">${t}</span>`;
      const statusPill = dv.grounded ? pill('#FEE2E2', '#B91C1C', '#FECACA', '⛔ AIRCRAFT GROUNDED — NOT AIRWORTHY')
        : dv.deferred ? pill('#FEF3C7', '#92400E', '#FDE68A', 'Deferred — MEL / CDL ' + esc(dv.melReference || '—'))
        : pill('#F2F4FA', '#18224A', '#E8EBF2', 'Reported — not grounded');
      const ctx = bundle.defectContext || {};
      defectSection = section('⚠️', 'Defect / Unserviceability',
        kv('Status', statusPill) +
        kv('System affected', esc(dv.systemAffected || '—')) +
        (dv.aircraftHours ? kv('Aircraft hours', esc(dv.aircraftHours)) : '') +
        kv('Description', nl2br(dv.defectDescription)) +
        (dv.actionTaken ? kv('Action taken', nl2br(dv.actionTaken)) : '') +
        (dv.maintenanceProvider ? kv('Maintenance notified', esc(dv.maintenanceProvider) + (dv.notifiedAt ? ' — ' + esc(String(dv.notifiedAt).replace('T', ' ')) : '')) : '') +
        (ctx.step ? kv('Raised during', esc((ctx.swmsFormNo ? ctx.swmsFormNo + ' ' : '') + (ctx.swmsForm || 'SWMS') + ' — step: ' + ctx.step)) : '') +
        kv('Reported by', esc(pilot), true));
    }
    let defectsListSection = '';
    const swmsDefects = !isDefect && Array.isArray(bundle.defects) ? bundle.defects : [];
    if (swmsDefects.length) {
      defectsListSection = section('⚠️', `Defects reported during this SWMS (${swmsDefects.length})`,
        swmsDefects.map((d, i) => kv(esc(d.step || 'General'),
          `<b>${esc(d.systemAffected || '')}</b>` + (d.grounded ? ' — <b style="color:#B91C1C;">GROUNDED</b>' : d.deferred ? ' — Deferred' + (d.melReference ? ' (' + esc(d.melReference) + ')' : '') : '') +
          `<br><span style="color:#555;">${esc(d.desc || '')}</span><br><span style="color:#aaa;font-size:11.5px;">Sent separately as its own defect report</span>`,
          i === swmsDefects.length - 1)).join(''));
    }

    const submittedTime = new Date().toLocaleString('en-AU', { timeZone: 'Australia/Darwin', hour:'2-digit', minute:'2-digit', hour12:false });

    const html = `
<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:24px 12px 48px;background:#ECEEF3;">
<div style="${S.wrap}border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,.10);overflow:hidden;">

  <!-- Header -->
  <div style="${S.hdr}">
    <span style="${S.badge}">${isDefect ? 'Defect Report' : 'Flight Paperwork'}</span>
    <div style="${S.coName}">${BRAND.shortName}</div>
    <div style="${S.coSub}">ABN ${BRAND.abn} &nbsp;·&nbsp; ${BRAND.location}</div>
  </div>

  <!-- Body -->
  <div style="${S.body}">
    <div style="${S.title}">${isDefect ? 'Defect Report Received' : 'Flight Paperwork Received'}</div>
    <div style="${S.sub}">Submitted ${dateStr} at ${submittedTime} ACST — PDF attached</div>

    ${section('✈', 'Flight Details',
        kv('Aircraft',   `${esc(bundle.callsign)} — ${esc(bundle.aircraftType || '')}`) +
        kv('Form',       esc((formNo ? formNo + ' — ' : '') + formName)) +
        kv('Pilot',      esc(pilot)) +
        (trainer ? kv('Trainer / Supervisor', esc(trainer)) : '') +
        kv('Date',       dateStr) +
        kv('PDF file',   esc(filename)) +
        (oneDriveUrl ? kv('OneDrive', `<a href="${oneDriveUrl}" style="color:#18224A;font-weight:600;">View filed PDF ↗</a>`, true) : kv('PDF file', esc(filename), true))
    )}

    ${defectSection}${defectsListSection}

    ${isDefect ? '' : section('⚖️', 'Weight &amp; Balance',
        kv('Result',        wbPill) +
        kv('Aircraft (BEW)', (wb.emptyWeight || '—') + ' kg') +
        kv('Pilot',          (wb.kg?.pilot   || 0) + ' kg') +
        (wb.kg?.paxList && wb.kg.paxList.length > 1
          ? wb.kg.paxList.map((w, i) => kv(`Passenger ${i + 1}`, (w || 0) + ' kg')).join('')
          : kv('Passenger(s)', (wb.kg?.pax || 0) + ' kg')) +
        kv('Baggage',        (wb.kg?.baggage || 0) + ' kg') +
        kv('Fuel',           `${fuelL} L (${fuelKg} kg) <span style="color:#888;font-size:11.5px;">— ${fuelType} @ ${fuelDens} kg/L</span>`) +
        kv('Total weight',   `<strong>${wbResult.total || '—'} kg</strong>`) +
        kv('MTOW',           (wbResult.mtow  || '—') + ' kg') +
        kv('CG arm',         wbResult.cgArm != null ? Math.round(wbResult.cgArm) + ' mm' : '—', true)
    )}

    ${acksSection}

    ${isDefect ? '' : section('👤', 'Passengers &amp; Safety Briefing', paxInner)}

    ${(() => {
      const pilotSig   = bundle.sms?.sigs?.pilotSig;
      const trainerSig = bundle.sms?.sigs?.trainerSig;
      if (!pilotSig && !trainerSig) return '';
      let inner = '';
      if (pilotSig) {
        inner += `<div style="${S.kv}"><span style="${S.k}">Pilot</span><span style="${S.v}"><div style="font-size:12px;font-weight:600;color:#555;margin-bottom:4px;">${esc(bundle.sms?.values?.pilotName || '')}</div><img src="${pilotSig}" style="height:60px;border:1px solid #E8EBF2;border-radius:6px;display:block;" alt="Pilot signature"></span></div>`;
      }
      if (trainerSig) {
        inner += `<div style="${S.kvLast}"><span style="${S.k}">Trainer / Examiner</span><span style="${S.v}"><div style="font-size:12px;font-weight:600;color:#555;margin-bottom:4px;">${esc(bundle.sms?.values?.trainerName || '')}</div><img src="${trainerSig}" style="height:60px;border:1px solid #E8EBF2;border-radius:6px;display:block;" alt="Trainer signature"></span></div>`;
      }
      return section('✍️', 'Signatures', inner);
    })()}

  </div>

  <!-- Footer -->
  <div style="${S.footer}">
    PDF attached · Sent automatically by the ${BRAND.shortest} flight paperwork app
  </div>

</div>
</body></html>
    `;

    if (!demoMode) {
      await graphSendMail(token, sender, opsTo, subject, html, pdfBuffer, filename);
      console.log('Email sent:', subject);
    } else {
      console.log('Demo mode — email skipped:', subject);
    }

    /* 4 — Save structured job record (OneDrive, or the local store in demo mode) */
    try { await saveJobRecord(token, bundle, oneDriveUrl); } catch (e) { console.error('Record save failed (non-fatal):', e.message); }

    res.json({ ok: true, filename, oneDriveUrl });

  } catch (err) {
    console.error('Send error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── Save job record to OneDrive _records/ ────────────────── */
async function saveJobRecord(token, bundle, oneDriveUrl) {
  _jobsCache = null; // invalidate cache so next report load is fresh
  const record = {
    submittedAt:  new Date().toISOString(),
    flightDate:   bundle.flightDate || new Date().toISOString().slice(0,10),
    flightTime:   bundle.flightTime || '',
    jobNo:        bundle.jobNo      || null,
    jobCode:      bundle.jobCode    || '',
    calendarJobId: bundle.calendarJobId || null,
    aircraftReg:  bundle.callsign   || '',
    aircraftType: bundle.aircraftType || '',
    pilotName:    bundle.sms?.values?.pilotName  || bundle.pilotName  || '',
    pilotArn:     bundle.sms?.values?.pilotArn   || bundle.pilotArn   || '',
    pilot2Name:   bundle.sms?.values?.trainerName || bundle.pilot2Name || '',
    crew:         Array.isArray(bundle.crew) ? bundle.crew : [],
    client:       bundle.client     || '',
    hireType:     bundle.hireType   || 'wet',
    totalHours:   bundle.totalHours || 0,
    lines:        bundle.lines      || [],
    fuelUplift:   bundle.fuelUplift || 0,
    notes:        bundle.notes      || '',
    paxCount:     (bundle.pax || []).length,
    wbPass:       bundle.wb?.result?.pass ?? null,
    oneDriveUrl:  oneDriveUrl || '',
    formName:     bundle.formName || '',
    formNo:       bundle.formNo   || '',
    kind:         (bundle.kind === 'defect' || bundle.sms?.formId === 'form-mdr') ? 'defect' : (bundle.kind || 'flight'),
    defectsReported: Array.isArray(bundle.defects) ? bundle.defects.length : 0,
  };
  if (record.kind === 'defect') {
    const dv = bundle.sms?.values || {};
    record.defect = {
      systemAffected: dv.systemAffected || '', description: dv.defectDescription || '',
      grounded: !!dv.grounded, deferred: !!dv.deferred, melReference: dv.melReference || '',
      aircraftHours: dv.aircraftHours || '', actionTaken: dv.actionTaken || '',
      maintenanceProvider: dv.maintenanceProvider || '', swmsStep: dv.swmsStep || '',
      parentFlightId: bundle.parentFlightId || null,
    };
  }

  const ts  = new Date().toISOString().replace(/[:.]/g,'-').slice(0,19);
  const reg = (record.aircraftReg).replace(/[^A-Z0-9]/gi,'');
  const filename = `${ts}_${reg || 'UNK'}.json`;

  if (!hasGraphCreds()) {
    const records = readLocalJson('records', []);
    records.push(record);
    writeLocalJson('records', records);
    console.log('Job record saved locally:', filename);
    return;
  }

  const driveUser  = process.env.OPS_EMAIL;
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const uploadPath = `${folderName}/_records/${filename}`;

  const r = await fetch(
    `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${encodeURIComponent(uploadPath)}:/content`,
    {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(record),
    }
  );
  if (!r.ok) throw new Error(`Record upload ${r.status}: ${await r.text()}`);
  console.log('Job record saved:', filename);
}

/* ── Send email via Microsoft Graph ──────────────────────── */
async function graphSendMail(token, from, to, subject, html, pdfBuffer, filename) {
  const body = {
    message: {
      subject,
      body:        { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: to } }],
      attachments: [{
        '@odata.type':  '#microsoft.graph.fileAttachment',
        name:           filename,
        contentType:    'application/pdf',
        contentBytes:   pdfBuffer.toString('base64'),
      }],
    },
    saveToSentItems: true,
  };

  const res = await fetch(`https://graph.microsoft.com/v1.0/users/${from}/sendMail`, {
    method:  'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body:    JSON.stringify(body),
  });

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Graph sendMail ${res.status}: ${txt}`);
  }
}

/* ── Upload to OneDrive via Microsoft Graph ───────────────── */
async function uploadToOneDrive(token, pdfBuffer, filename, callsign, month) {
  const driveUser   = process.env.OPS_EMAIL;
  const folderName  = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const uploadPath  = `${folderName}/${callsign}/${month}/${filename}`;
  const baseUrl     = `https://graph.microsoft.com/v1.0/users/${driveUser}/drive`;

  const res = await fetch(
    `${baseUrl}/root:/${encodeURIComponent(uploadPath)}:/content`,
    {
      method:  'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/pdf' },
      body:    pdfBuffer,
    }
  );

  if (!res.ok) {
    const txt = await res.text();
    throw new Error(`Graph upload ${res.status}: ${txt}`);
  }

  const item = await res.json();
  return item.webUrl || null;
}

/* ── PDF builder ──────────────────────────────────────────── */
async function buildPDF(bundle) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDocument({ margin: 50, size: 'A4', bufferPages: true });
    const chunks = [];
    doc.on('data',  c => chunks.push(c));
    doc.on('end',   () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const NAVY   = '#18224A';
    const ORANGE = '#E8750E';
    const MUT    = '#6B7280';
    const W      = 495;

    /* ── Header ── */
    const HDR_H = 72;
    doc.rect(50, 50, W, HDR_H).fill(NAVY);

    const logoPath = path.join(__dirname, 'logo.png');
    let textX = 68;
    if (_brandLogoBuf) {
      try { doc.image(_brandLogoBuf, 64, 57, { height: 52, width: 52 }); textX = 126; } catch (_) {}
    } else if (fs.existsSync(logoPath)) {
      doc.image(logoPath, 64, 57, { height: 52, width: 52 });
      textX = 126;
    }

    doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(16)
       .text(BRAND.shortName.toUpperCase(), textX, 64, { width: W - (textX - 50) - 16, lineBreak: false });
    doc.font('Helvetica').fontSize(9.5).fillColor('#9AA3C7')
       .text(`Flight Paperwork Bundle — ${BRAND.companyName}`, textX, 85, { width: W - (textX - 50) - 16 });

    doc.rect(50, 50 + HDR_H, W, 3).fill(ORANGE);
    doc.y = 50 + HDR_H + 3 + 18;

    /* ── Page overflow guard ── */
    const checkY = (needed = 30) => {
      if (doc.y + needed > 758) doc.addPage();
    };

    /* ── Section heading ── */
    const secHead = title => {
      checkY(56);
      doc.moveDown(0.3);
      const y = doc.y;
      doc.rect(50, y, W, 20).fill('#E8EBF5');
      doc.fillColor(NAVY).font('Helvetica-Bold').fontSize(9.5)
         .text(title, 58, y + 5, { width: W - 16, lineBreak: false });
      doc.y = y + 26;
      doc.fillColor('#1C1F28');
    };

    /* ── Key/value row ── */
    const kv = (k, v, color) => {
      checkY(18);
      const y = doc.y;
      doc.font('Helvetica-Bold').fontSize(9.5).fillColor(MUT)
         .text(k, 50, y, { width: 128, lineBreak: false });
      doc.font('Helvetica').fontSize(9.5).fillColor(color || '#1C1F28')
         .text(String(v === null || v === undefined ? '—' : v), 185, y, { width: W - 135, lineBreak: false });
      doc.y = y + 17;
    };

    /* ── Flight details ── */
    const ts = bundle.queuedAt || bundle.createdAt || Date.now();
    secHead('FLIGHT DETAILS');
    kv('Aircraft', bundle.callsign);
    kv('Form',     bundle.formName || '—');
    kv('Pilot',    bundle.sms?.values?.pilotName);
    /* Use pilot-set flight date/time if available, fall back to submission time */
    const flightDateStr = bundle.flightDate
      ? new Date(bundle.flightDate + 'T12:00:00').toLocaleDateString('en-AU', { day: '2-digit', month: 'long', year: 'numeric' })
      : acstDate(ts);
    const flightTimeStr = bundle.flightTime || acstTime(ts);
    kv('Date', flightDateStr);
    kv('Time', flightTimeStr);
    if (bundle.sms?.values?.trainerName) kv('Trainer', bundle.sms.values.trainerName);

    /* ── Defect report (red ! on SWMS / standalone MDR) ── */
    const para = (k, v, color) => {
      checkY(34);
      const y = doc.y;
      doc.font('Helvetica-Bold').fontSize(9.5).fillColor(MUT).text(k, 50, y, { width: 128 });
      const yk = doc.y;
      doc.font('Helvetica').fontSize(9.5).fillColor(color || '#1C1F28')
         .text(String(v || '—'), 185, y, { width: W - 135 });
      doc.y = Math.max(doc.y, yk, y + 17) + 3;
    };
    const pdv = bundle.sms?.values || {};
    const pIsDefect = bundle.kind === 'defect' || bundle.sms?.formId === 'form-mdr';
    if (pIsDefect) {
      secHead('DEFECT / UNSERVICEABILITY');
      kv('Status', pdv.grounded ? 'AIRCRAFT GROUNDED — NOT AIRWORTHY' : pdv.deferred ? 'DEFERRED — MEL / CDL ' + (pdv.melReference || '—') : 'Reported — not grounded',
         pdv.grounded ? '#B91C1C' : pdv.deferred ? '#92400E' : null);
      kv('System affected', pdv.systemAffected || '—');
      if (pdv.aircraftHours) kv('Aircraft hours', pdv.aircraftHours);
      para('Description', pdv.defectDescription);
      if (pdv.actionTaken) para('Action taken', pdv.actionTaken);
      if (pdv.maintenanceProvider) kv('Maintenance notified', pdv.maintenanceProvider + (pdv.notifiedAt ? ' — ' + String(pdv.notifiedAt).replace('T', ' ') : ''));
      const ctx = bundle.defectContext || {};
      if (ctx.step) para('Raised during', (ctx.swmsFormNo ? ctx.swmsFormNo + ' ' : '') + (ctx.swmsForm || 'SWMS') + ' — step: ' + ctx.step);
    }
    if (!pIsDefect && Array.isArray(bundle.defects) && bundle.defects.length) {
      secHead(`DEFECTS REPORTED DURING THIS SWMS (${bundle.defects.length})`);
      bundle.defects.forEach(d => para(d.step || 'General',
        (d.grounded ? 'GROUNDED — ' : d.deferred ? 'DEFERRED' + (d.melReference ? ' (' + d.melReference + ')' : '') + ' — ' : '') +
        (d.systemAffected || '') + ': ' + (d.desc || '') + '  [sent separately as its own defect report]',
        d.grounded ? '#B91C1C' : null));
    }

    /* ── Job Advice: Client & hire type ── */
    if (bundle.client) {
      secHead('JOB DETAILS');
      kv('Client',    bundle.client);
      kv('Job No',    bundle.jobNoFormatted || bundle.jobNo || '—');
      kv('Hire Type', bundle.hireType === 'dry' ? 'Dry Hire' : bundle.hireType === 'dual' ? 'Dual Flight' : 'Wet Hire');
      const crewList = Array.isArray(bundle.crew) && bundle.crew.length ? bundle.crew : (bundle.pilot2Name ? [{ pilotName: bundle.pilot2Name, aircraftReg: bundle.aircraft2Reg }] : []);
      crewList.forEach((c, i) => {
        if (c.pilotName) kv(`Pilot ${i + 2}`, c.pilotName + (c.aircraftReg ? ` — ${c.aircraftReg}` : ''));
      });
    }

    /* ── Job Advice: Flight hour lines ── */
    if (Array.isArray(bundle.lines) && bundle.lines.length) {
      const fmtLineDate = d => {
        if (!d) return '—';
        try { return new Date(d + 'T12:00:00').toLocaleDateString('en-AU', { day: '2-digit', month: 'short' }); }
        catch { return d; }
      };
      secHead('FLIGHT HOURS');
      const hY = doc.y;
      doc.font('Helvetica-Bold').fontSize(8.5).fillColor(MUT);
      doc.text('DATE',        50,  hY, { width: 65,  lineBreak: false });
      doc.text('DESCRIPTION', 125, hY, { width: 270, lineBreak: false });
      doc.text('HOURS',       405, hY, { width: 90,  lineBreak: false });
      doc.y = hY + 14;
      doc.rect(50, doc.y, W, 0.5).fill('#D1D5DB'); doc.y += 6;
      bundle.lines.forEach((l, i) => {
        checkY(18);
        const y = doc.y;
        if (i % 2 === 0) { doc.rect(50, y-2, W, 17).fill('#F9FAFB'); }
        doc.font('Helvetica').fontSize(9.5).fillColor('#1C1F28')
           .text(fmtLineDate(l.date), 50, y, { width: 65, lineBreak: false });
        doc.text(l.desc || '—', 125, y, { width: 270, lineBreak: false });
        doc.text((l.hours||0).toFixed(1) + ' hrs', 405, y, { width: 90, lineBreak: false });
        doc.y = y + 17;
      });
      doc.moveDown(0.3);
      doc.rect(50, doc.y, W, 0.5).fill('#D1D5DB'); doc.y += 8;
      const totY = doc.y;
      doc.font('Helvetica-Bold').fontSize(10).fillColor(NAVY)
         .text('TOTAL FLIGHT HOURS', 50, totY, { width: 340, lineBreak: false });
      doc.font('Helvetica-Bold').fontSize(10).fillColor(ORANGE)
         .text((bundle.totalHours||0).toFixed(1) + ' hrs', 395, totY, { width: 100, lineBreak: false });
      doc.y = totY + 20;
      if (bundle.fuelUplift) kv('Fuel Uplift', bundle.fuelUplift + ' L');
      if (bundle.notes)      kv('Notes',       bundle.notes);
    }

    /* ── W&B ── */
    if (bundle.wb?.result) {
      const wb = bundle.wb.result;
      secHead('WEIGHT & BALANCE');
      kv('Result',       wb.pass ? 'PASS' : 'FAIL — OVER MTOW', wb.pass ? '#15803D' : '#B91C1C');
      kv('Total weight', wb.total + ' kg');
      kv('MTOW',         wb.mtow  + ' kg');
      if (bundle.wb.kg) {
        const kg = bundle.wb.kg;
        kv('Pilot',   (kg.pilot   || 0) + ' kg');
        if (kg.paxList && kg.paxList.length > 1) {
          kg.paxList.forEach((w, i) => kv('Passenger ' + (i + 1), (w || 0) + ' kg'));
        } else {
          kv('Pax',   (kg.pax || 0) + ' kg');
        }
        kv('Fuel',    (kg.fuel    || 0) + ' kg  (' + (bundle.wb.fuelL || 0) + ' L)');
        kv('Baggage', (kg.baggage || 0) + ' kg');
      }
      if (bundle.wb.cgEnv && wb.cgArm) {
        drawCgChart(doc, bundle.wb.cgEnv, wb.cgArm, wb.total,
                    bundle.wb.emptyLongArm || 0, bundle.wb.emptyWeight || 0);
      }
    }

    /* ── SWMS acknowledgments ── */
    if (bundle.sms?.acks && Object.keys(bundle.sms.acks).length) {
      secHead('SAFETY MANAGEMENT — SECTION ACKNOWLEDGMENTS');
      doc.font('Helvetica').fontSize(8.5).fillColor(MUT)
         .text('Pilot confirms they have read, understood and will comply with each section below.', 50, doc.y, { width: W });
      doc.y += 6;
      const ackLabels = bundle.smsAckLabels || [];
      Object.entries(bundle.sms.acks).forEach(([i, v]) => {
        checkY(18);
        const y = doc.y;
        const idx = parseInt(i);
        const label = ackLabels[idx]?.step || ('Section ' + (idx + 1));
        doc.font('Helvetica-Bold').fontSize(9).fillColor(v ? '#15803D' : '#B91C1C')
           .text((v ? '✓' : '✗'), 50, y, { width: 16, lineBreak: false });
        doc.font('Helvetica').fontSize(9).fillColor(v ? '#15803D' : '#B91C1C')
           .text((idx + 1) + '. ' + label, 68, y, { width: W - 18, lineBreak: false });
        doc.y = y + 16;
      });
    }

    /* ── Passengers ── */
    if ((bundle.pax || []).length) {
      secHead('PASSENGERS');
      checkY(54);
      const hY = doc.y;
      doc.font('Helvetica-Bold').fontSize(8.5).fillColor(MUT);
      doc.text('NAME',    50,  hY, { width: 190, lineBreak: false });
      doc.text('WEIGHT', 245,  hY, { width:  70, lineBreak: false });
      doc.text('BRIEFED', 320, hY, { width:  70, lineBreak: false });
      doc.text('SIGNED',  400, hY, { width:  70, lineBreak: false });
      doc.y = hY + 14;
      doc.rect(50, doc.y, W, 0.5).fill('#D1D5DB');
      doc.y += 6;
      bundle.pax.forEach((p, i) => {
        checkY(18);
        const y = doc.y;
        if (i % 2 === 0) { doc.rect(50, y - 2, W, 17).fill('#F9FAFB'); doc.fillColor('#1C1F28'); }
        doc.font('Helvetica').fontSize(9.5).fillColor('#1C1F28').text(p.name || '—', 50, y, { width: 190, lineBreak: false });
        doc.text((p.weight || '—') + ' kg',    245, y, { width: 70, lineBreak: false });
        doc.fillColor(p.briefed ? '#15803D' : MUT).text(p.briefed ? 'Yes' : 'No', 320, y, { width: 70, lineBreak: false });
        doc.fillColor(p.sig     ? '#15803D' : MUT).text(p.sig     ? 'Yes' : 'No', 400, y, { width: 70, lineBreak: false });
        doc.y = y + 17;
      });
    }

    /* ── Passenger signatures ── */
    const signedPax = (bundle.pax || []).filter(p => p.sig);
    if (signedPax.length) {
      secHead('PASSENGER SIGNATURES');
      signedPax.forEach(p => {
        checkY(110);
        doc.font('Helvetica-Bold').fontSize(9.5).fillColor(NAVY)
           .text(p.name || '', 50, doc.y, { width: W, lineBreak: false });
        doc.y += 14;
        try {
          const raw = p.sig.replace(/^data:image\/png;base64,/, '');
          doc.image(Buffer.from(raw, 'base64'), 50, doc.y, { width: 200, height: 65 });
          doc.y += 72;
        } catch (_) {}
        doc.rect(50, doc.y, 200, 0.5).fill('#1C1F28');
        doc.font('Helvetica').fontSize(8).fillColor(MUT)
           .text((p.date || '') + (p.time ? ' ' + p.time : ''), 50, doc.y + 4, { width: 200, lineBreak: false });
        doc.y += 22;
      });
    }

    /* ── Pilot & trainer signatures ── */
    const sigDefs = [
      { id: 'pilotSig',   label: 'PILOT SIGNATURE',              nameKey: 'pilotName'   },
      { id: 'trainerSig', label: 'TRAINER / EXAMINER SIGNATURE', nameKey: 'trainerName' },
    ];
    sigDefs.forEach(({ id, label, nameKey }) => {
      const data = bundle.sms?.sigs?.[id];
      if (!data) return;
      checkY(130);
      secHead(label);
      try {
        const raw = data.replace(/^data:image\/png;base64,/, '');
        doc.image(Buffer.from(raw, 'base64'), 50, doc.y, { width: 240, height: 80 });
        doc.y += 88;
      } catch (_) {}
      doc.rect(50, doc.y, 240, 0.5).fill('#1C1F28');
      doc.y += 5;
      doc.font('Helvetica').fontSize(9.5).fillColor('#1C1F28')
         .text(bundle.sms?.values?.[nameKey] || '', 50, doc.y, { width: 240, lineBreak: false });
      doc.y += 20;
    });

    /* ── Footer on every page ── */
    const pages = doc.bufferedPageRange();
    for (let i = 0; i < pages.count; i++) {
      doc.switchToPage(pages.start + i);
      doc.rect(50, 778, W, 0.5).fill('#D1D5DB');
      doc.font('Helvetica').fontSize(7.5).fillColor(MUT)
         .text(
           `${BRAND.companyName}  ·  Page ${i + 1} of ${pages.count}  ·  Generated ${acstFull(Date.now())} ACST`,
           50, 783, { width: W, align: 'center' }
         );
    }

    doc.end();
  });
}

/* ── CG balance map (PDFKit vector drawing) ───────────────── */
function drawCgChart(doc, E, cgArm, weight, bewArm, bewWeight) {
  const L = 70, T = doc.y + 10, CW = 340, CH = 180;
  const R = L + CW, B = T + CH;

  /* axis range */
  const aMin = E.armMin, aMax = E.armMax, wMin = E.wMin, wMax = E.wMax;
  const toX = a => L + (a - aMin) / (aMax - aMin) * CW;
  const toY = w => T + (1 - (w - wMin) / (wMax - wMin)) * CH;

  /* check height — add page if needed */
  if (T + CH + 40 > 750) { doc.addPage(); return drawCgChart(doc, E, cgArm, weight, bewArm, bewWeight); }

  const NAVY = '#18224A', GREEN = '#16a34a', RED = '#dc2626', MUT = '#73778A';

  /* title */
  doc.font('Helvetica-Bold').fontSize(9).fillColor(NAVY)
     .text('LONGITUDINAL BALANCE MAP', L, T - 14, { width: CW, align: 'center' });

  /* grid */
  const aRange = aMax - aMin, wRange = wMax - wMin;
  const aStep = aRange <= 300 ? 50 : aRange <= 500 ? 100 : 200;
  const wStep = wRange <= 300 ? 50 : wRange <= 600 ? 100 : 200;
  doc.save();
  for (let a = Math.ceil(aMin / aStep) * aStep; a <= aMax; a += aStep) {
    const x = toX(a);
    doc.moveTo(x, T).lineTo(x, B).stroke('#EAEAEA');
    doc.font('Helvetica').fontSize(7).fillColor(MUT).text(String(a), x - 14, B + 3, { width: 28, align: 'center' });
  }
  for (let w = Math.ceil(wMin / wStep) * wStep; w <= wMax; w += wStep) {
    const y = toY(w);
    doc.moveTo(L, y).lineTo(R, y).stroke('#EAEAEA');
    doc.font('Helvetica').fontSize(7).fillColor(MUT).text(String(w), L - 30, y - 4, { width: 26, align: 'right' });
  }
  doc.restore();

  /* border */
  doc.rect(L, T, CW, CH).lineWidth(0.5).stroke('#CCCCCC');

  /* envelope polygon (filled) */
  if (E.poly && E.poly.length >= 3) {
    const pts = E.poly.map(([a, w]) => ({ x: toX(a), y: toY(w) }));
    doc.save();
    doc.moveTo(pts[0].x, pts[0].y);
    pts.slice(1).forEach(p => doc.lineTo(p.x, p.y));
    doc.closePath().fillColor('#DCFCE7').fill();
    doc.moveTo(pts[0].x, pts[0].y);
    pts.slice(1).forEach(p => doc.lineTo(p.x, p.y));
    doc.closePath().lineWidth(1.5).stroke(GREEN);
    doc.restore();
    doc.font('Helvetica-Bold').fontSize(7).fillColor(GREEN).text('SAFE ZONE', L + 4, T + 4);
  }

  /* axis labels */
  doc.font('Helvetica').fontSize(7).fillColor(MUT)
     .text('kg', L - 28, T - 4)
     .text('Longitudinal arm (mm from datum)', L, B + 13, { width: CW, align: 'center' })
     .text('FWD →', L + 4, B - 12)
     .text('← AFT', R - 30, B - 12);

  /* BEW dot */
  if (bewArm && bewWeight) {
    const bx = toX(bewArm), by = toY(bewWeight);
    doc.circle(bx, by, 3).fillColor('#9CA3AF').fill();
    doc.font('Helvetica').fontSize(6).fillColor('#9CA3AF').text('BEW', bx + 4, by - 3);
  }

  /* CG dot */
  const inside = cgInPolyPdf(cgArm, weight, E.poly);
  const dotColor = inside === true ? '#15803d' : inside === false ? RED : '#9CA3AF';
  const cx = toX(cgArm), cy = toY(weight);
  if (cx >= L && cx <= R && cy >= T && cy <= B) {
    doc.circle(cx, cy, 6).fillColor(dotColor).fill();
    doc.circle(cx, cy, 6).lineWidth(1.5).stroke('#FFFFFF');
  }

  /* status line */
  const statusTxt = inside === true
    ? `✓ CG WITHIN LIMITS — ${cgArm} mm / ${weight} kg`
    : inside === false
    ? `✗ CG OUTSIDE LIMITS — ${cgArm} mm / ${weight} kg`
    : `CG: ${cgArm} mm / ${weight} kg`;
  doc.font('Helvetica-Bold').fontSize(8).fillColor(dotColor)
     .text(statusTxt, L, B + 22, { width: CW, align: 'center' });

  doc.y = B + 36;
  doc.fillColor('#1C1F28');
}

function cgInPolyPdf(arm, kg, poly) {
  if (!poly || poly.length < 3) return null;
  let inside = false;
  for (let i = 0, j = poly.length - 1; i < poly.length; j = i++) {
    const [xi, yi] = poly[i], [xj, yj] = poly[j];
    if ((yi > kg) !== (yj > kg) && arm < (xj - xi) * (kg - yi) / (yj - yi) + xi) inside = !inside;
  }
  return inside;
}

/* ── Escape HTML ──────────────────────────────────────────── */
function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ── Scheduler: SWMS/briefing reminders + daily 6pm job logging ─
   Runs inside this always-on API process — no external cron needed.
   1. Every 2 minutes: text (SMS) any pilot whose scheduled job starts
      within the next hour (or started up to 15 min ago, to catch
      jobs allocated with too little notice) and hasn't been
      reminded yet.
   2. Once per day, from 6:00pm ACST onward: mark that day's
      scheduled jobs as "logged" and create a draft job sheet
      (main details pre-filled) under each pilot's name.
   ============================================================ */
let _lastDailySweepDate = null;

/* ── Scheduler lock ────────────────────────────────────────────
   This app currently runs as a single DigitalOcean instance, so
   there's only ever one copy of runScheduler() ticking. If that
   ever changes — more instances added for uptime or traffic —
   every instance would run its own copy of this loop with no
   coordination, and pilots would get duplicate texts. This lock
   is a best-effort guard against that: each instance ID's claim
   is written to OneDrive, and any instance that sees a fresh claim
   from a different ID skips its tick. It's not a true atomic lock
   (OneDrive's plain content PUT has no compare-and-swap), so a
   same-millisecond race between two instances starting up at once
   is still possible in theory — cheap insurance for the normal
   case, not a guarantee for high-concurrency deployments.
   ============================================================ */
const INSTANCE_ID = crypto.randomUUID();
const SCHEDULER_LOCK_STALE_MS = 5 * 60 * 1000; // longer than the 2-min tick — a lock older than this means its owner died

async function acquireSchedulerLock(token, folderName) {
  const lockPath = `${folderName}/_calendar/_scheduler-lock.json`;
  try {
    const lock = await getOneDriveJson(token, lockPath);
    if (lock && lock.ownerId !== INSTANCE_ID) {
      const age = Date.now() - new Date(lock.lockedAt).getTime();
      if (age < SCHEDULER_LOCK_STALE_MS) return false; // another instance is active
    }
  } catch (e) { console.error('Scheduler lock read failed (continuing):', e.message); }
  try {
    await putOneDriveJson(token, lockPath, { ownerId: INSTANCE_ID, lockedAt: new Date().toISOString() });
  } catch (e) { console.error('Scheduler lock write failed (continuing anyway):', e.message); }
  return true;
}

function acstDateKey(ts = Date.now()) {
  return new Intl.DateTimeFormat('en-CA', { timeZone: ACST_TZ, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date(ts));
}
function acstHourNow(ts = Date.now()) {
  const parts = new Intl.DateTimeFormat('en-GB', { timeZone: ACST_TZ, hour: '2-digit', hour12: false }).formatToParts(new Date(ts));
  return parseInt(parts.find(p => p.type === 'hour').value, 10);
}

/* ── SMS via Twilio ────────────────────────────────────────────
   Needs three env vars in DigitalOcean (see SETUP.md):
     TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN, TWILIO_FROM_NUMBER
   Uses the plain Twilio REST API over fetch — no extra npm package.
   Used for: the 1-hour-before reminder, and change/cancellation
   notices when a calendar job is edited or cancelled.
   ============================================================ */
async function sendTwilioSMS(to, body) {
  const sid  = process.env.TWILIO_ACCOUNT_SID;
  const auth = process.env.TWILIO_AUTH_TOKEN;
  const from = process.env.TWILIO_FROM_NUMBER;
  if (!sid || !auth || !from) { console.warn('SMS skipped — TWILIO_ACCOUNT_SID/AUTH_TOKEN/FROM_NUMBER not set'); return; }
  const toNormalized = normalizeAuPhone(to);
  if (!toNormalized) { console.warn('SMS skipped — no phone number on file'); return; }

  const params = new URLSearchParams({ To: toNormalized, From: from, Body: body });
  const res = await fetch(`https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`, {
    method: 'POST',
    headers: {
      Authorization: 'Basic ' + Buffer.from(`${sid}:${auth}`).toString('base64'),
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: params.toString(),
  });
  if (!res.ok) throw new Error(`Twilio send ${res.status}: ${await res.text()}`);
}

function fmtJobWhen(job) {
  return job.date + (job.startTime ? ' at ' + job.startTime : '');
}

async function sendReminderSMS(job, pilot, minsAway) {
  if (!pilot || !pilot.phone) { console.warn('SMS reminder skipped — no phone for', pilot && pilot.name, 'on job', job.id); return; }
  const urgency = (minsAway != null && minsAway <= 0)
    ? `has already started${job.startTime ? ' (' + job.startTime + ')' : ''}`
    : `starts in about 1 hour${job.startTime ? ' (' + job.startTime + ')' : ''}`;
  const body =
    `${BRAND.shortest}: reminder — your job` +
    (job.client ? ` for ${job.client}` : '') +
    ` ${urgency}. ` +
    `Please complete your SWMS and flight briefing before you fly.`;
  await sendTwilioSMS(pilot.phone, body);
}

/* ── Per-pilot notices: new assignment, change, removal, cancellation ─
   hadDraft means the job had already been logged (6pm sweep has run
   for it), so a draft job sheet may exist — the wording flags that
   without asserting it was fixed remotely, since a draft already
   pulled onto a device can't be reached or edited from here. ──── */
async function notifyJobChanged(oldJob, job, pilot, hadDraft) {
  if (!pilot || !pilot.phone) return;
  const dateChanged = oldJob.date !== job.date;
  const timeChanged  = (oldJob.startTime || '') !== (job.startTime || '');
  let what;
  if (dateChanged && timeChanged) {
    what = `has moved to ${fmtJobWhen(job)} (was ${fmtJobWhen(oldJob)})`;
  } else if (timeChanged) {
    what = `start time has moved to ${job.startTime || 'unset'} (was ${oldJob.startTime || 'unset'}), still on ${job.date}`;
  } else if (dateChanged) {
    what = `has moved to ${job.date}${job.startTime ? ' at ' + job.startTime : ''} (was ${oldJob.date})`;
  } else {
    what = `has been updated — now ${fmtJobWhen(job)}`;
  }
  const draftNote = hadDraft ? ' If you already have a job sheet started for this, double-check the details before you submit it.' : '';
  const body = `${BRAND.shortest}: your job${job.client ? ` for ${job.client}` : ''} ${what}.${draftNote} Check the app for details.`;
  try { await sendTwilioSMS(pilot.phone, body); } catch (e) { console.error('Change SMS failed:', job.id, pilot.name, e.message); }
}

async function notifyJobAssigned(job, pilot) {
  if (!pilot || !pilot.phone) return;
  const body = `${BRAND.shortest}: you've been allocated a job${job.client ? ` for ${job.client}` : ''} on ${fmtJobWhen(job)}. Check the app for details.`;
  try { await sendTwilioSMS(pilot.phone, body); } catch (e) { console.error('New-assignment SMS failed:', job.id, pilot.name, e.message); }
}

async function notifyJobRemoved(oldJob, pilot, hadDraft) {
  if (!pilot || !pilot.phone) return;
  const draftNote = hadDraft ? ' If a job sheet was already started for this in the app, please don’t submit it.' : ' No action needed.';
  const body = `${BRAND.shortest}: you've been taken off the job on ${fmtJobWhen(oldJob)}${oldJob.client ? ` (${oldJob.client})` : ''}.${draftNote}`;
  try { await sendTwilioSMS(pilot.phone, body); } catch (e) { console.error('Reassignment-removed SMS failed:', oldJob.id, pilot.name, e.message); }
}

async function notifyJobCancelled(job, pilot, hadDraft) {
  if (!pilot || !pilot.phone) return;
  const draftNote = hadDraft ? ' If a job sheet was already started for this in the app, please don’t submit it.' : ' No action needed.';
  const body = `${BRAND.shortest}: your job on ${fmtJobWhen(job)}${job.client ? ` (${job.client})` : ''} has been CANCELLED.${draftNote}`;
  try { await sendTwilioSMS(pilot.phone, body); } catch (e) { console.error('Cancellation SMS failed:', job.id, pilot.name, e.message); }
}

async function runScheduler() {
  // No Graph creds → no OneDrive-backed calendar to sweep and no Twilio SMS to
  // send (e.g. a demo deployment). Skip quietly instead of logging an error
  // every 2 minutes forever.
  if (!hasGraphCreds()) return;
  try {
    const token = await getGraphToken();
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';

    const gotLock = await acquireSchedulerLock(token, folderName);
    if (!gotLock) { console.log('Scheduler: another instance holds the lock — skipping this tick'); return; }

    const jobs = await loadCalendarJobs(true);
    if (!jobs.length) return;
    const now = Date.now();

    /* 1 — 1-hour-before SWMS/flight briefing reminder, per pilot on the job */
    for (const job of jobs) {
      if (job.status !== 'scheduled' || !job.startTime) continue;
      const startMs = new Date(`${job.date}T${job.startTime}:00+09:30`).getTime(); // Australia/Darwin, fixed UTC+9:30
      if (Number.isNaN(startMs)) continue;
      const minsAway = (startMs - now) / 60000;
      // Window runs from 60 min before start to 15 min after — the "after" side catches jobs
      // created or edited with too little notice for a tick to have caught the 60-min mark cleanly
      if (minsAway <= -15 || minsAway > 60) continue;

      const pilots  = jobPilots(job);
      const already = new Set(job.remindedPilots || []);
      const due     = pilots.filter(p => !already.has(p.name));
      if (!due.length) continue;

      for (const p of due) {
        try {
          await sendReminderSMS(job, p, minsAway);
          already.add(p.name);
          console.log('SMS reminder sent for job', job.id, p.name);
        } catch (e) { console.error('Reminder send failed for job', job.id, p.name, e.message); }
      }
      try {
        await putOneDriveJson(token, `${folderName}/_calendar/${job.id}.json`, { ...job, remindedPilots: [...already] });
      } catch (e) { console.error('Failed to save remindedPilots for job', job.id, e.message); }
    }

    /* 2 — Once-daily 6pm ACST sweep: log today's jobs + create a draft job sheet per pilot */
    const todayKey = acstDateKey(now);
    if (acstHourNow(now) >= 18 && _lastDailySweepDate !== todayKey) {
      _lastDailySweepDate = todayKey;
      const todays = jobs.filter(j => j.date === todayKey && j.status === 'scheduled');
      for (const job of todays) {
        try {
          const loggedJob = { ...job, status: 'logged', loggedAt: new Date().toISOString() };
          await putOneDriveJson(token, `${folderName}/_calendar/${job.id}.json`, loggedJob);

          for (const p of jobPilots(job)) {
            const draft = await createDraftForPilot(token, job, p);
            console.log('Job logged + draft sheet created for', p.name, '→', draft.id);
          }
        } catch (e) { console.error('Daily sweep failed for job', job.id, e.message); }
      }
      _calCache = null;
    }
  } catch (err) {
    console.error('Scheduler run failed (non-fatal, will retry in 5 min):', err.message);
  }
}

setInterval(runScheduler, 2 * 60 * 1000);
setTimeout(runScheduler, 15 * 1000); // also run shortly after boot/deploy

/* ── Start ────────────────────────────────────────────────── */
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Heli API listening on :${PORT}`));
