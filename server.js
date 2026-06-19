'use strict';
/* ============================================================
   Outback Helicopter Airwork NT — Flight Paperwork API
   POST /api/send  →  generate PDF, send via Office 365, file to OneDrive
   GET  /reports   →  password-protected reporting dashboard
   GET  /api/jobs  →  job records from OneDrive (auth required)
   All via Microsoft Graph API — one set of credentials for everything
   ============================================================ */

const express     = require('express');
const cors        = require('cors');
const PDFDocument = require('pdfkit');
const fs          = require('fs');
const path        = require('path');
const crypto      = require('crypto');
const { ClientSecretCredential } = require('@azure/identity');

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
function requireReportsAuth(req, res, next) {
  if (!REPORTS_PWD) return res.status(503).send('Set REPORTS_PASSWORD environment variable in DigitalOcean.');
  const cookies = parseCookies(req);
  if (cookies.rpt_auth === makeToken(REPORTS_PWD)) return next();
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
app.use(cors());
app.use(express.json({ limit: '20mb' }));

/* ── Shared config (pilots, aircraft, clients) ────────────── */
/* Edit /api/config.json in GitHub to update all devices */
app.get(['/config', '/api/config'], (_req, res) => {
  res.sendFile(path.join(__dirname, 'config.json'));
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
app.get('/api/jobs', requireReportsAuth, async (req, res) => {
  try {
    const forceRefresh = req.query.refresh === '1';
    const now = Date.now();
    if (!forceRefresh && _jobsCache && now - _jobsCacheAt < CACHE_TTL) {
      return res.json({ jobs: _jobsCache, cached: true });
    }

    const token      = await getGraphToken();
    const driveUser  = process.env.SENDER_EMAIL;
    const folderName = process.env.ONEDRIVE_FOLDER || 'Helicopter Paperwork';
    const recPath    = encodeURIComponent(`${folderName}/_records`);

    let files = [];
    let url = `https://graph.microsoft.com/v1.0/users/${driveUser}/drive/root:/${recPath}:/children`
            + `?$select=name,@microsoft.graph.downloadUrl&$top=1000`;

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
        try {
          const r = await fetch(f['@microsoft.graph.downloadUrl']);
          return r.ok ? await r.json() : null;
        } catch { return null; }
      }));
      jobs.push(...results.filter(Boolean));
    }

    _jobsCache   = jobs;
    _jobsCacheAt = now;
    res.json({ jobs, total: jobs.length });
  } catch (err) {
    console.error('Jobs fetch error:', err.message);
    res.status(500).json({ ok:false, error: err.message });
  }
});

/* ── Send bundle ──────────────────────────────────────────── */
app.post('/send', async (req, res) => {
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
    const subject = `Flight Paperwork — ${bundle.callsign} — ${bundle.formName || 'Form'} — ${dateStr}`;
    const sender  = process.env.SENDER_EMAIL;   // e.g. ops@outbackhelicopters.com.au
    const opsTo   = process.env.OPS_EMAIL;      // who receives the paperwork

    const html = `
      <h2 style="font-family:sans-serif;color:#18224A;">Outback Helicopter Airwork NT</h2>
      <h3 style="font-family:sans-serif;margin-top:0;">Flight Paperwork Received</h3>
      <table style="font-family:sans-serif;font-size:14px;border-collapse:collapse;">
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;white-space:nowrap;">Aircraft</td><td>${esc(bundle.callsign)}</td></tr>
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">Form</td><td>${esc(bundle.formName || '—')}</td></tr>
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">Pilot</td><td>${esc(pilot)}</td></tr>
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">Date</td><td>${dateStr}</td></tr>
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">W&amp;B</td><td>${
          bundle.wb?.result?.pass
            ? '✅ PASS — ' + bundle.wb.result.total + ' kg / ' + bundle.wb.result.mtow + ' kg MTOW'
            : bundle.wb?.result ? '❌ OVER MTOW' : 'Not recorded'
        }</td></tr>
        <tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">Passengers</td><td>${(bundle.pax || []).length}</td></tr>
        ${oneDriveUrl ? `<tr><td style="padding:5px 16px 5px 0;font-weight:700;color:#555;">OneDrive</td><td><a href="${oneDriveUrl}">View filed PDF</a></td></tr>` : ''}
      </table>
      <p style="font-family:sans-serif;font-size:12px;color:#999;margin-top:24px;">
        PDF attached · Sent automatically by the Outback Helicopter flight paperwork app
      </p>
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
    aircraftReg:  bundle.callsign   || '',
    aircraftType: bundle.aircraftType || '',
    pilotName:    bundle.sms?.values?.pilotName  || bundle.pilotName  || '',
    pilotArn:     bundle.sms?.values?.pilotArn   || bundle.pilotArn   || '',
    pilot2Name:   bundle.sms?.values?.trainerName || bundle.pilot2Name || '',
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

  const driveUser  = process.env.SENDER_EMAIL;
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
  const driveUser   = process.env.SENDER_EMAIL;
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
    if (fs.existsSync(logoPath)) {
      doc.image(logoPath, 64, 57, { height: 52, width: 52 });
      textX = 126;
    }

    doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(16)
       .text('OUTBACK HELICOPTER AIRWORK NT', textX, 64, { width: W - (textX - 50) - 16, lineBreak: false });
    doc.font('Helvetica').fontSize(9.5).fillColor('#9AA3C7')
       .text('Flight Paperwork Bundle — Outback Helicopter Airwork NT Pty Ltd', textX, 85, { width: W - (textX - 50) - 16 });

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
      if (bundle.pilot2Name) kv('2nd Pilot', bundle.pilot2Name);
    }

    /* ── Job Advice: Flight hour lines ── */
    if (Array.isArray(bundle.lines) && bundle.lines.length) {
      secHead('FLIGHT HOURS');
      const hY = doc.y;
      doc.font('Helvetica-Bold').fontSize(8.5).fillColor(MUT);
      doc.text('DESCRIPTION', 50, hY, { width: 340, lineBreak: false });
      doc.text('HOURS',       395, hY, { width: 100, lineBreak: false });
      doc.y = hY + 14;
      doc.rect(50, doc.y, W, 0.5).fill('#D1D5DB'); doc.y += 6;
      bundle.lines.forEach((l, i) => {
        checkY(18);
        const y = doc.y;
        if (i % 2 === 0) { doc.rect(50, y-2, W, 17).fill('#F9FAFB'); }
        doc.font('Helvetica').fontSize(9.5).fillColor('#1C1F28')
           .text(l.desc || '—', 50, y, { width: 340, lineBreak: false });
        doc.text((l.hours||0).toFixed(1) + ' hrs', 395, y, { width: 100, lineBreak: false });
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
        kv('Pax',     (kg.pax     || 0) + ' kg');
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
      Object.entries(bundle.sms.acks).forEach(([i, v]) => {
        checkY(18);
        const y = doc.y;
        doc.font('Helvetica').fontSize(9.5).fillColor(v ? '#15803D' : '#B91C1C')
           .text((v ? '✓  ' : '✗  ') + 'Section ' + (parseInt(i) + 1), 50, y, { width: W, lineBreak: false });
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

    /* ── Signatures ── */
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
           `Outback Helicopter Airwork NT Pty Ltd  ·  Page ${i + 1} of ${pages.count}  ·  Generated ${acstFull(Date.now())} ACST`,
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

/* ── Start ────────────────────────────────────────────────── */
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Heli API listening on :${PORT}`));
