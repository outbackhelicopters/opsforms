'use strict';
/* ============================================================
   Per-user authentication for the office admin app.

   Users live in OneDrive at  <ONEDRIVE_FOLDER>/_system/users.json
   (same storage pattern as the calendar). When Microsoft creds
   aren't configured (local dev), falls back to api/_users.local.json.

   Roles:
     provider — the platform operator (John's team). Present on every
                deployment; cannot be removed or downgraded by customers.
     admin    — the customer's responsible person (chief pilot / HOFO).
     office   — day-to-day scheduling and paperwork; no settings.

   Passwords: scrypt (Node built-in), per-user salt.
   Sessions:  HMAC-signed HTTP-only cookie, 14 days.
   Invites:   one-time token emailed via Graph; also returned to the
              inviter as a link in case email isn't configured yet.
   ============================================================ */

const crypto = require('crypto');
const fs     = require('fs');
const path   = require('path');

const ROLES        = ['provider', 'admin', 'office'];
const COOKIE       = 'adm_sess';
const SESSION_DAYS = 14;
const INVITE_DAYS  = 7;
const STORE_TTL_MS = 60 * 1000;

module.exports = function initAuth(app, deps) {
  const { getGraphToken, getOneDriveJson, putOneDriveJson, parseCookies, rateLimit } = deps;

  const FOLDER     = () => process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
  const USERS_PATH = () => `${FOLDER()}/_system/users.json`;
  const LOCAL_FILE = path.join(__dirname, '_users.local.json');
  const hasGraph   = () => !!(process.env.MS_TENANT_ID && process.env.MS_CLIENT_ID && process.env.MS_CLIENT_SECRET);
  const SECRET     = process.env.SESSION_SECRET
    || crypto.createHash('sha256').update('adm-' + (process.env.MS_CLIENT_SECRET || 'local-dev')).digest('hex');

  /* ── Store (cached, write-through) ─────────────────────── */
  let _store = null, _storeAt = 0;

  async function loadStore(force) {
    const now = Date.now();
    if (!force && _store && now - _storeAt < STORE_TTL_MS) return _store;
    let data = null;
    if (hasGraph()) {
      const token = await getGraphToken();
      data = await getOneDriveJson(token, USERS_PATH());
    } else if (fs.existsSync(LOCAL_FILE)) {
      try { data = JSON.parse(fs.readFileSync(LOCAL_FILE, 'utf8')); } catch { data = null; }
    }
    _store = data && Array.isArray(data.users) ? data : { users: [], invites: [] };
    _store.invites = (_store.invites || []).filter(i => i.exp > now); // drop expired
    _storeAt = now;
    return _store;
  }

  async function saveStore(store) {
    store.updatedAt = new Date().toISOString();
    if (hasGraph()) {
      const token = await getGraphToken();
      await putOneDriveJson(token, USERS_PATH(), store);
    } else {
      fs.writeFileSync(LOCAL_FILE, JSON.stringify(store, null, 2));
    }
    _store = store; _storeAt = Date.now();
  }

  /* ── Passwords ─────────────────────────────────────────── */
  function hashPassword(pw, salt) {
    salt = salt || crypto.randomBytes(16).toString('hex');
    const hash = crypto.scryptSync(String(pw), salt, 64).toString('hex');
    return { salt, hash };
  }
  function verifyPassword(pw, user) {
    if (!user || !user.hash || !user.salt) return false;
    const test = crypto.scryptSync(String(pw), user.salt, 64);
    const real = Buffer.from(user.hash, 'hex');
    return test.length === real.length && crypto.timingSafeEqual(test, real);
  }

  /* ── Sessions ──────────────────────────────────────────── */
  function sign(s) { return crypto.createHmac('sha256', SECRET).update(s).digest('base64url'); }

  function makeSessionCookie(email) {
    const payload = Buffer.from(JSON.stringify({ e: email, x: Date.now() + SESSION_DAYS * 864e5 })).toString('base64url');
    const value   = `${payload}.${sign(payload)}`;
    return `${COOKIE}=${value}; Path=/; HttpOnly; Max-Age=${SESSION_DAYS * 86400}; SameSite=Strict`;
  }

  async function sessionUser(req) {
    const raw = parseCookies(req)[COOKIE];
    if (!raw) return null;
    const dot = raw.lastIndexOf('.');
    if (dot < 0) return null;
    const payload = raw.slice(0, dot), sig = raw.slice(dot + 1);
    const expect = sign(payload);
    if (sig.length !== expect.length || !crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(expect))) return null;
    let data;
    try { data = JSON.parse(Buffer.from(payload, 'base64url').toString()); } catch { return null; }
    if (!data.e || Date.now() > data.x) return null;
    const store = await loadStore();
    const user = store.users.find(u => u.email === data.e && u.active !== false);
    return user || null;
  }

  function requireAuth(roles) {
    return async (req, res, next) => {
      try {
        const user = await sessionUser(req);
        if (!user) return res.status(401).json({ ok: false, error: 'Not signed in' });
        if (roles && !roles.includes(user.role)) return res.status(403).json({ ok: false, error: 'Not permitted for your role' });
        req.user = user;
        next();
      } catch (err) {
        console.error('auth check error:', err.message);
        res.status(500).json({ ok: false, error: 'Auth check failed' });
      }
    };
  }

  /* ── Login attempt limiter (per IP, 10 per 15 min) ─────── */
  const _attempts = new Map();
  function loginLimiter(req, res, next) {
    const ip = req.ip || 'unknown';
    const now = Date.now();
    const hits = (_attempts.get(ip) || []).filter(t => now - t < 15 * 60 * 1000);
    if (hits.length >= 10) return res.status(429).json({ ok: false, error: 'Too many attempts — wait 15 minutes' });
    hits.push(now);
    _attempts.set(ip, hits);
    next();
  }

  /* ── Plain email (no attachment) via Graph ─────────────── */
  async function sendPlainMail(to, subject, html) {
    if (!hasGraph() || !process.env.SENDER_EMAIL) throw new Error('Email not configured');
    const token = await getGraphToken();
    const r = await fetch(`https://graph.microsoft.com/v1.0/users/${process.env.SENDER_EMAIL}/sendMail`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: {
        subject,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: to } }],
      }, saveToSentItems: true }),
    });
    if (!r.ok) throw new Error(`sendMail ${r.status}: ${await r.text()}`);
  }

  /* DO ingress strips the /api prefix in production; local dev keeps it.
     Register every route under both shapes so both environments work. */
  const both = p => [p, '/api' + p];

  function publicUser(u) {
    return { email: u.email, name: u.name, role: u.role, active: u.active !== false, createdAt: u.createdAt, lastLogin: u.lastLogin || null };
  }
  function baseUrl(req) {
    const proto = req.headers['x-forwarded-proto'] || req.protocol || 'https';
    return `${proto}://${req.headers.host}`;
  }

  /* ══ ROUTES ══════════════════════════════════════════════ */

  /* First-run: does any account exist yet? */
  app.get(both('/auth/status'), async (_req, res) => {
    try {
      const store = await loadStore();
      res.json({ ok: true, bootstrapped: store.users.length > 0 });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  /* Create the very first account (provider). Only works while zero users exist. */
  app.post(both('/auth/bootstrap'), rateLimit, async (req, res) => {
    try {
      const store = await loadStore(true);
      if (store.users.length > 0) return res.status(403).json({ ok: false, error: 'Already set up' });
      const { name, email, password } = req.body || {};
      if (!name || !email || !password || String(password).length < 8)
        return res.status(400).json({ ok: false, error: 'Name, email and a password of 8+ characters required' });
      const { salt, hash } = hashPassword(password);
      store.users.push({ email: String(email).toLowerCase().trim(), name: String(name).trim(),
        role: 'provider', salt, hash, active: true, createdAt: new Date().toISOString() });
      await saveStore(store);
      res.setHeader('Set-Cookie', makeSessionCookie(store.users[0].email));
      res.json({ ok: true, user: publicUser(store.users[0]) });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  app.post(both('/auth/login'), loginLimiter, async (req, res) => {
    try {
      const { email, password } = req.body || {};
      const store = await loadStore();
      const user = store.users.find(u => u.email === String(email || '').toLowerCase().trim());
      if (!user || user.active === false || !verifyPassword(password, user))
        return res.status(401).json({ ok: false, error: 'Wrong email or password' });
      user.lastLogin = new Date().toISOString();
      saveStore(store).catch(e => console.error('lastLogin save failed:', e.message));
      res.setHeader('Set-Cookie', makeSessionCookie(user.email));
      res.json({ ok: true, user: publicUser(user) });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  app.post(both('/auth/logout'), (_req, res) => {
    res.setHeader('Set-Cookie', `${COOKIE}=; Path=/; HttpOnly; Max-Age=0; SameSite=Strict`);
    res.json({ ok: true });
  });

  app.get(both('/auth/me'), requireAuth(), (req, res) => res.json({ ok: true, user: publicUser(req.user) }));

  /* ── User management (admin + provider) ────────────────── */

  app.get(both('/users'), requireAuth(['provider', 'admin']), async (_req, res) => {
    try {
      const store = await loadStore();
      res.json({ ok: true,
        users: store.users.map(publicUser),
        invites: store.invites.map(i => ({ email: i.email, name: i.name, role: i.role, kind: i.kind, exp: i.exp })) });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  /* Invite a new user (or send a password reset with kind:'reset') */
  app.post(both('/users/invite'), requireAuth(['provider', 'admin']), rateLimit, async (req, res) => {
    try {
      const { name, email, role, kind } = req.body || {};
      const em = String(email || '').toLowerCase().trim();
      const store = await loadStore(true);
      let entry;
      if (kind === 'reset') {
        const target = store.users.find(u => u.email === em);
        if (!target) return res.status(404).json({ ok: false, error: 'No such user' });
        if (target.role === 'provider' && req.user.role !== 'provider')
          return res.status(403).json({ ok: false, error: 'Only a provider can reset a provider account' });
        entry = { kind: 'reset', email: em, name: target.name, role: target.role };
      } else {
        if (!name || !em) return res.status(400).json({ ok: false, error: 'Name and email required' });
        if (!ROLES.includes(role)) return res.status(400).json({ ok: false, error: 'Bad role' });
        if (role === 'provider' && req.user.role !== 'provider')
          return res.status(403).json({ ok: false, error: 'Only a provider can create provider accounts' });
        if (store.users.some(u => u.email === em)) return res.status(400).json({ ok: false, error: 'User already exists' });
        entry = { kind: 'invite', email: em, name: String(name).trim(), role };
      }
      entry.token = crypto.randomBytes(24).toString('hex');
      entry.exp = Date.now() + INVITE_DAYS * 864e5;
      entry.by = req.user.email;
      store.invites = store.invites.filter(i => !(i.email === em && i.kind === entry.kind));
      store.invites.push(entry);
      await saveStore(store);

      const link = `${baseUrl(req)}/admin.html#invite=${entry.token}`;
      let emailed = false;
      try {
        const what = entry.kind === 'reset' ? 'reset your password' : 'set up your account';
        await sendPlainMail(em, `Flight Ops — ${entry.kind === 'reset' ? 'password reset' : 'your account'}`,
          `<p>Hi ${entry.name},</p><p>Use the link below to ${what} (valid ${INVITE_DAYS} days):</p>` +
          `<p><a href="${link}">${link}</a></p>`);
        emailed = true;
      } catch (e) { console.error('invite email failed:', e.message); }
      res.json({ ok: true, emailed, link });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  /* Accept an invite / reset: set password, sign in */
  app.post(both('/auth/accept-invite'), rateLimit, async (req, res) => {
    try {
      const { token, password } = req.body || {};
      if (!password || String(password).length < 8)
        return res.status(400).json({ ok: false, error: 'Password of 8+ characters required' });
      const store = await loadStore(true);
      const inv = store.invites.find(i => i.token === token && i.exp > Date.now());
      if (!inv) return res.status(400).json({ ok: false, error: 'Invalid or expired link — ask for a new one' });
      const { salt, hash } = hashPassword(password);
      let user = store.users.find(u => u.email === inv.email);
      if (user) { user.salt = salt; user.hash = hash; user.active = true; }
      else {
        user = { email: inv.email, name: inv.name, role: inv.role, salt, hash, active: true, createdAt: new Date().toISOString() };
        store.users.push(user);
      }
      store.invites = store.invites.filter(i => i.token !== token);
      await saveStore(store);
      res.setHeader('Set-Cookie', makeSessionCookie(user.email));
      res.json({ ok: true, user: publicUser(user) });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  /* Update a user: role and/or active. Guarded. */
  app.put(both('/users/:email'), requireAuth(['provider', 'admin']), async (req, res) => {
    try {
      const em = String(req.params.email).toLowerCase();
      const store = await loadStore(true);
      const target = store.users.find(u => u.email === em);
      if (!target) return res.status(404).json({ ok: false, error: 'No such user' });
      if (target.email === req.user.email) return res.status(400).json({ ok: false, error: "You can't change your own account here" });
      if (target.role === 'provider' && req.user.role !== 'provider')
        return res.status(403).json({ ok: false, error: 'Provider accounts can only be changed by a provider' });
      const { role, active } = req.body || {};
      if (role !== undefined) {
        if (!ROLES.includes(role)) return res.status(400).json({ ok: false, error: 'Bad role' });
        if (role === 'provider' && req.user.role !== 'provider')
          return res.status(403).json({ ok: false, error: 'Only a provider can grant the provider role' });
        target.role = role;
      }
      if (active !== undefined) {
        if (target.role === 'provider' && active === false &&
            store.users.filter(u => u.role === 'provider' && u.active !== false).length <= 1)
          return res.status(400).json({ ok: false, error: 'Cannot deactivate the last provider account' });
        target.active = !!active;
      }
      await saveStore(store);
      res.json({ ok: true, user: publicUser(target) });
    } catch (err) { res.status(500).json({ ok: false, error: err.message }); }
  });

  return { sessionUser, requireAuth };
};
