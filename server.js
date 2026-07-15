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
  try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8')); }
  catch (e) { console.error('config.json load failed:', e.message); return {}; }
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
  try {
    const forceRefresh = req.query.refresh === '1';
    const now = Date.now();
    if (!forceRefresh && _jobsCache && now - _jobsCacheAt < CACHE_TTL) {
      return res.json({ jobs: _jobsCache, cached: true });
    }

    const token      = await getGraphToken();
    const driveUser  = process.env.OPS_EMAIL;
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    const recPath    = encodeURIComponent(`${folderName}/_records`);

    let files = [];
    let url = `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${recPath}:/children`
            + `?$select=name,@microsoft.graph.downloadUrl&$top=1000`;

    const _debug = { steps: [] }; // TEMP diagnostic — returned inline in the JSON response, remove after root cause confirmed
    while (url) {
      const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (r.status === 404) { _debug.steps.push({ list: '404' }); break; }
      if (!r.ok) { _debug.steps.push({ list: 'non-ok', status: r.status, body: (await r.text()).slice(0, 300) }); throw new Error(`List records: ${r.status}`); }
      const d = await r.json();
      _debug.steps.push({ list: 'ok', count: (d.value || []).length, sample: d.value && d.value[0] ? { name: d.value[0].name, hasDownloadUrl: !!d.value[0]['@microsoft.graph.downloadUrl'], keys: Object.keys(d.value[0]) } : null });
      files.push(...(d.value || []).filter(f => f.name && f.name.endsWith('.json')));
      url = d['@odata.nextLink'] || null;
    }

    // Download all records in parallel batches of 20
    _debug.filesLength = files.length;
    const jobs = [];
    const downloadNotes = [];
    for (let i = 0; i < files.length; i += 20) {
      const batch = files.slice(i, i + 20);
      const results = await Promise.all(batch.map(async f => {
        try {
          const dlUrl = f['@microsoft.graph.downloadUrl'];
          if (!dlUrl) { downloadNotes.push({ name: f.name, issue: 'no-download-url' }); return null; }
          const r = await fetch(dlUrl);
          if (!r.ok) { downloadNotes.push({ name: f.name, issue: 'fetch-not-ok', status: r.status }); return null; }
          return await r.json();
        } catch (e) { downloadNotes.push({ name: f.name, issue: 'exception', message: e.message }); return null; }
      }));
      jobs.push(...results.filter(Boolean));
    }
    _debug.jobsAfterDownload = jobs.length;
    _debug.downloadNotesSample = downloadNotes.slice(0, 5);

    _jobsCache   = jobs;
    _jobsCacheAt = now;
    res.json({ jobs, total: jobs.length, _debug: forceRefresh ? _debug : undefined });
  } catch (err) {
    console.error('Jobs fetch error:', err.message);
    res.status(500).json({ ok:false, error: err.message });
  }
});

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
        LIVE_OPS = {
          pilots:   Array.isArray(data.pilots)   ? data.pilots   : [],
          aircraft: Array.isArray(data.aircraft) ? data.aircraft : [],
          clients:  Array.isArray(data.clients)  ? data.clients  : [],
        };
      } else {
        // First boot for this deployment — seed OneDrive from the repo
        // copy so nothing breaks, then the repo file stops mattering.
        await putOneDriveJson(token, opPath, LIVE_OPS);
      }
    } else if (fs.existsSync(OPCONFIG_LOCAL)) {
      const data = JSON.parse(fs.readFileSync(OPCONFIG_LOCAL, 'utf8'));
      LIVE_OPS = {
        pilots:   Array.isArray(data.pilots)   ? data.pilots   : LIVE_OPS.pilots,
        aircraft: Array.isArray(data.aircraft) ? data.aircraft : LIVE_OPS.aircraft,
        clients:  Array.isArray(data.clients)  ? data.clients  : LIVE_OPS.clients,
      };
    }
  } catch (e) { console.error('live config load failed (using repo config.json):', e.message); }
}
/* Aircraft carry a nested W&B block. Numbers only, always saved as
   unverified — POH sign-off is a deliberate step a customer's own
   admin/chief pilot takes later (see PLAN-admin-and-commercial.md);
   nothing entered here or via the setup wizard is ever auto-verified. */
function numOrZero(v) { const n = parseFloat(v); return Number.isFinite(n) ? n : 0; }
function normalizeAircraft(list) {
  return (Array.isArray(list) ? list : []).map(a => {
    const wIn = (a && a.wb) || {};
    const wb = {
      verified:     false,
      source:       String(wIn.source || '').trim(),
      emptyWeight:  numOrZero(wIn.emptyWeight),
      emptyLongArm: numOrZero(wIn.emptyLongArm),
      emptyLatArm:  numOrZero(wIn.emptyLatArm),
      mtow:         numOrZero(wIn.mtow),
      fuelDensity:  numOrZero(wIn.fuelDensity) || 0.720,
      cgEnvKey:     String(wIn.cgEnvKey || '').trim(), // matches a built-in CG envelope preset, if any
      accessories: (Array.isArray(wIn.accessories) ? wIn.accessories : [])
        .map(x => ({ name: String((x && x.name) || '').trim(), weight: numOrZero(x && x.weight), arm: numOrZero(x && x.arm) }))
        .filter(x => x.name),
    };
    return { reg: String((a && a.reg) || '').trim().toUpperCase(), type: String((a && a.type) || '').trim(), wb };
  }).filter(a => a.reg);
}
async function saveLiveOpsConfig(patch) {
  if (Array.isArray(patch.pilots))   LIVE_OPS.pilots   = normalizePilots(patch.pilots);
  if (Array.isArray(patch.aircraft)) LIVE_OPS.aircraft  = normalizeAircraft(patch.aircraft);
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
    await saveLiveOpsConfig(b);
    res.json({ ok: true, pilots: LIVE_OPS.pilots, aircraft: LIVE_OPS.aircraft, clients: LIVE_OPS.clients });
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
          'emptyLatArm (mm, number or null), mtow (kg, number or null), fuelDensity (kg/L, number or null). ' +
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
  const token = await getGraphToken();
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const jobs = await listOneDriveJsonFiles(token, `${folderName}/_calendar`);
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
    const token = await getGraphToken();
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
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    await putOneDriveJson(token, `${folderName}/_calendar/${id}.json`, record);
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
    const token = await getGraphToken();
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

    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    await putOneDriveJson(token, `${folderName}/_calendar/${req.params.id}.json`, updated);
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

    if (wasLogged) {
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
    const token = await getGraphToken();
    const jobs = await loadCalendarJobs(true);
    const existing = jobs.find(j => j.id === req.params.id);
    if (!existing) return res.status(404).json({ ok: false, error: 'Not found' });
    const wasScheduled = existing.status === 'scheduled';
    const wasLogged    = existing.status === 'logged';
    const updated = { ...existing, status: 'cancelled', cancelledAt: new Date().toISOString() };
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    await putOneDriveJson(token, `${folderName}/_calendar/${req.params.id}.json`, updated);
    _calCache = null;
    res.json({ ok: true });

    // Only notify if the job was still upcoming or just logged — a job cancelled twice, or one
    // that's already cancelled/done, doesn't need another text
    if (wasScheduled || wasLogged) for (const p of jobPilots(updated)) scheduleNotice(req.params.id, p.name, () => notifyJobCancelled(updated, p, wasLogged));

    // If a draft job sheet already exists for this job, cancel any copy still sitting on the
    // server (unpulled) — a pilot who already has it on their iPad is told by SMS not to submit it
    if (wasLogged) cancelJobDrafts(token, req.params.id).catch(e => console.error('draft cancel failed:', e.message));
  } catch (err) {
    console.error('calendar-jobs delete error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

/* ── Shared Job Advice Sheet number counter ─────────────────────
   Job numbers need to go up the same way no matter which pilot's
   device submits next — a per-device counter (the old approach)
   drifts the moment two devices are in use. Single number stored
   in OneDrive, cached in memory once loaded, incremented under an
   in-process lock so two submits landing in the same instant still
   come out as two different numbers. This only serializes writes
   within this one running instance — fine at this app's scale,
   same tradeoff already accepted by the drafts lock below. ── */
let JOB_COUNTER = null; // { n } — null until first loaded from OneDrive
let _jobNoLockChain = Promise.resolve();
function withJobNoLock(fn) {
  const result = _jobNoLockChain.then(fn, fn);
  _jobNoLockChain = result.then(() => {}, () => {});
  return result;
}
async function loadJobCounter(token) {
  if (JOB_COUNTER !== null) return JOB_COUNTER;
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  try {
    const data = await getOneDriveJson(token, `${folderName}/_system/job-counter.json`);
    JOB_COUNTER = { n: (data && Number.isFinite(data.n)) ? data.n : 0 };
  } catch (e) {
    JOB_COUNTER = { n: 0 }; // nothing saved yet — first job of this deployment
  }
  return JOB_COUNTER;
}

/* POST /api/job-number/next — atomically hands out the next Job Advice
   Sheet number. Only call this once, right at the moment a pilot actually
   submits (not while they're still filling the form in) — every call
   consumes a number, even if the submit is later abandoned. */
app.post(['/job-number/next', '/api/job-number/next'], requireDevice, rateLimit, async (req, res) => {
  try {
    const next = await withJobNoLock(async () => {
      const token = await getGraphToken();
      const counter = await loadJobCounter(token);
      counter.n += 1;
      const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
      await putOneDriveJson(token, `${folderName}/_system/job-counter.json`, { n: counter.n, updatedAt: new Date().toISOString() });
      return counter.n;
    });
    res.json({ ok: true, number: next });
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
      const token = await getGraphToken();
      const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
      const drafts = await listOneDriveJsonFiles(token, `${folderName}/_drafts`);
      const pilot = req.query.pilot;
      const claim = drafts.filter(d => !d.pulled && d.status === 'draft' && (!pilot || d.pilotName === pilot));
      for (const d of claim) {
        try {
          await putOneDriveJson(token, `${folderName}/_drafts/${d.id}.json`, { ...d, pulled: true, pulledAt: new Date().toISOString() });
        } catch (e) { console.error('draft pull-flag failed:', e.message); }
      }
      return claim;
    });
    res.json({ drafts: pending });
  } catch (err) {
    console.error('job-drafts error:', err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
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
    const filename = `${bundle.callsign}_${safeForm}_${dateStr}.pdf`;

    /* Get Microsoft Graph access token (shared for email + OneDrive) */
    const token = await getGraphToken();

    /* 2 — File to OneDrive */
    let oneDriveUrl = null;
    try {
      oneDriveUrl = await uploadToOneDrive(token, pdfBuffer, filename, bundle.callsign, dateStr.slice(0, 7));
      console.log('OneDrive:', oneDriveUrl);
    } catch (err) {
      console.error('OneDrive upload failed (non-fatal):', err.message);
    }

    /* 3 — Send email via Office 365 */
    const pilot   = bundle.sms?.values?.pilotName || 'Unknown pilot';
    const trainer = bundle.sms?.values?.trainerName || '';
    const formName = bundle.formName || bundle.sms?.formId || 'Flight Operations Form';
    const formNo   = bundle.formNo   || '';
    const subject = `Flight Paperwork — ${bundle.callsign} — ${formName} — ${dateStr}`;
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

    const submittedTime = new Date().toLocaleString('en-AU', { timeZone: 'Australia/Darwin', hour:'2-digit', minute:'2-digit', hour12:false });

    const html = `
<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:24px 12px 48px;background:#ECEEF3;">
<div style="${S.wrap}border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,.10);overflow:hidden;">

  <!-- Header -->
  <div style="${S.hdr}">
    <span style="${S.badge}">Flight Paperwork</span>
    <div style="${S.coName}">${BRAND.shortName}</div>
    <div style="${S.coSub}">ABN ${BRAND.abn} &nbsp;·&nbsp; ${BRAND.location}</div>
  </div>

  <!-- Body -->
  <div style="${S.body}">
    <div style="${S.title}">Flight Paperwork Received</div>
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

    ${section('⚖️', 'Weight &amp; Balance',
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

    ${section('👤', 'Passengers &amp; Safety Briefing', paxInner)}

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

    await graphSendMail(token, sender, opsTo, subject, html, pdfBuffer, filename);
    console.log('Email sent:', subject);

    /* 4 — Save structured job record to OneDrive for reporting */
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
  };

  const driveUser  = process.env.OPS_EMAIL;
  const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const ts  = new Date().toISOString().replace(/[:.]/g,'-').slice(0,19);
  const reg = (record.aircraftReg).replace(/[^A-Z0-9]/gi,'');
  const filename = `${ts}_${reg || 'UNK'}.json`;
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
       .text('OUTBACK HELICOPTER AIRWORK NT', textX, 64, { width: W - (textX - 50) - 16, lineBreak: false });
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

    /* ── Job Advice: Client & hire type ── */
    if (bundle.client) {
      secHead('JOB DETAILS');
      kv('Client',    bundle.client);
      kv('Job No',    bundle.jobNo   || '—');
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
