'use strict';
/* ============================================================
   Outback Helicopter Airwork NT — Flight Paperwork API
   POST /api/send  →  generate PDF, send via Office 365, file to OneDrive
   All via Microsoft Graph API — one set of credentials for everything
   ============================================================ */

const express     = require('express');
const cors        = require('cors');
const PDFDocument = require('pdfkit');
const { ClientSecretCredential } = require('@azure/identity');

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));

/* ── Health check ─────────────────────────────────────────── */
app.get('/api/health', (_req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

/* ── Send bundle ──────────────────────────────────────────── */
app.post('/api/send', async (req, res) => {
  const bundle = req.body;
  if (!bundle || !bundle.callsign) return res.status(400).json({ ok: false, error: 'Invalid bundle' });

  try {
    /* 1 — Generate PDF */
    const pdfBuffer = await buildPDF(bundle);

    const safeForm = (bundle.formName || 'form').replace(/[^a-zA-Z0-9_-]/g, '_');
    const dateStr  = new Date(bundle.queuedAt || Date.now()).toISOString().slice(0, 10);
    const filename = `${bundle.callsign}_${safeForm}_${dateStr}.pdf`;

    /* Get Microsoft Graph access token (shared for email + OneDrive) */
    const credential = new ClientSecretCredential(
      process.env.MS_TENANT_ID,
      process.env.MS_CLIENT_ID,
      process.env.MS_CLIENT_SECRET
    );
    const { token } = await credential.getToken('https://graph.microsoft.com/.default');

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

    res.json({ ok: true, filename, oneDriveUrl });

  } catch (err) {
    console.error('Send error:', err);
    res.status(500).json({ ok: false, error: err.message });
  }
});

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
function buildPDF(bundle) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDocument({ margin: 50, size: 'A4', bufferPages: true });
    const chunks = [];
    doc.on('data',  c => chunks.push(c));
    doc.on('end',   () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    const NAVY   = '#18224A';
    const ORANGE = '#E8750E';
    const MUT    = '#73778A';
    const W      = 495;

    /* Header */
    doc.rect(50, 50, W, 64).fill(NAVY);
    doc.fillColor('#fff').font('Helvetica-Bold').fontSize(15)
       .text('OUTBACK HELICOPTER AIRWORK NT', 66, 63, { width: W - 32 });
    doc.font('Helvetica').fontSize(10).fillColor('#9AA3C7')
       .text('Flight Paperwork Bundle', 66, 83, { width: W - 32 });
    doc.rect(50, 114, W, 2).fill(ORANGE);
    doc.moveDown(3);

    const secHead = title => {
      doc.moveDown(0.5);
      doc.rect(50, doc.y, W, 22).fill('#F0F1F5');
      doc.fillColor(NAVY).font('Helvetica-Bold').fontSize(10)
         .text(title, 58, doc.y - 17, { width: W - 16 });
      doc.moveDown(0.8);
      doc.fillColor('#1C1F28');
    };

    const kv = (k, v, color) => {
      doc.font('Helvetica-Bold').fontSize(10).fillColor(MUT)
         .text(k, 50, doc.y, { width: 130 });
      doc.font('Helvetica').fillColor(color || '#1C1F28')
         .text(String(v || '—'), 185, doc.y - doc.currentLineHeight(), { width: W - 135 });
      doc.moveDown(0.15);
    };

    /* Flight details */
    secHead('FLIGHT DETAILS');
    const dt = new Date(bundle.queuedAt || bundle.createdAt || Date.now());
    kv('Aircraft', bundle.callsign);
    kv('Form',     bundle.formName);
    kv('Pilot',    bundle.sms?.values?.pilotName);
    kv('ARN',      bundle.sms?.values?.pilotArn);
    kv('Date',     dt.toLocaleDateString('en-AU', { dateStyle: 'long' }));
    kv('Time',     dt.toLocaleTimeString('en-AU', { timeStyle: 'short' }));
    if (bundle.sms?.values?.trainerName) kv('Trainer', bundle.sms.values.trainerName);

    /* W&B */
    if (bundle.wb?.result) {
      const wb = bundle.wb.result;
      secHead('WEIGHT & BALANCE');
      kv('Result',       wb.pass ? 'PASS' : 'FAIL — OVER MTOW', wb.pass ? '#1E7A3C' : '#C0241C');
      kv('Total weight', wb.total + ' kg');
      kv('MTOW',         wb.mtow  + ' kg');
      if (bundle.wb.kg) {
        const kg = bundle.wb.kg;
        kv('Pilot',   (kg.pilot   || 0) + ' kg');
        kv('Pax',     (kg.pax     || 0) + ' kg');
        kv('Fuel',    (kg.fuel    || 0) + ' kg  (' + (bundle.wb.fuelL || 0) + ' L)');
        kv('Baggage', (kg.baggage || 0) + ' kg');
      }
    }

    /* SWMS acknowledgments */
    if (bundle.sms?.acks && Object.keys(bundle.sms.acks).length) {
      secHead('SAFETY MANAGEMENT — SECTION ACKNOWLEDGMENTS');
      Object.entries(bundle.sms.acks).forEach(([i, v]) => {
        doc.font('Helvetica').fontSize(10).fillColor(v ? '#1E7A3C' : '#C0241C')
           .text((v ? '✓' : '✗') + '  Section ' + (parseInt(i) + 1), 50, doc.y, { width: W });
        doc.moveDown(0.15);
      });
    }

    /* Passengers */
    if ((bundle.pax || []).length) {
      secHead('PASSENGERS');
      doc.font('Helvetica-Bold').fontSize(9).fillColor(MUT);
      doc.text('NAME', 50, doc.y, { width: 200 });
      doc.text('WEIGHT', 255, doc.y - doc.currentLineHeight(), { width: 70 });
      doc.text('BRIEFED', 330, doc.y - doc.currentLineHeight(), { width: 70 });
      doc.text('SIGNED',  405, doc.y - doc.currentLineHeight(), { width: 70 });
      doc.moveDown(0.4);
      doc.rect(50, doc.y, W, 0.5).fill('#E6E5E0');
      doc.moveDown(0.4);
      bundle.pax.forEach((p, i) => {
        const y = doc.y;
        if (i % 2 === 0) doc.rect(50, y - 2, W, 16).fill('#FAFAF8');
        doc.font('Helvetica').fontSize(10).fillColor('#1C1F28')
           .text(p.name || '—', 50, y, { width: 200 });
        doc.text((p.weight || '—') + ' kg', 255, y, { width: 70 });
        doc.fillColor(p.briefed ? '#1E7A3C' : MUT).text(p.briefed ? 'Yes' : 'No', 330, y, { width: 70 });
        doc.fillColor(p.sig    ? '#1E7A3C' : MUT).text(p.sig    ? 'Yes' : 'No', 405, y, { width: 70 });
        doc.moveDown(0.5);
      });
    }

    /* Signatures */
    const sigDefs = [
      { id: 'pilotSig',   label: 'PILOT SIGNATURE',             nameKey: 'pilotName'   },
      { id: 'trainerSig', label: 'TRAINER / EXAMINER SIGNATURE', nameKey: 'trainerName' },
    ];
    sigDefs.forEach(({ id, label, nameKey }) => {
      const data = bundle.sms?.sigs?.[id];
      if (!data) return;
      secHead(label);
      try {
        const raw = data.replace(/^data:image\/png;base64,/, '');
        doc.image(Buffer.from(raw, 'base64'), 50, doc.y, { width: 220, height: 80 });
        doc.moveDown(5.5);
      } catch (_) {}
      doc.rect(50, doc.y, 220, 0.5).fill('#1C1F28');
      doc.moveDown(0.3);
      doc.font('Helvetica').fontSize(10).fillColor('#1C1F28')
         .text(bundle.sms?.values?.[nameKey] || '', 50, doc.y, { width: 220 });
      doc.moveDown(0.8);
    });

    /* Footer on every page */
    const pages = doc.bufferedPageRange();
    for (let i = 0; i < pages.count; i++) {
      doc.switchToPage(pages.start + i);
      doc.rect(50, 780, W, 0.5).fill('#E6E5E0');
      doc.font('Helvetica').fontSize(8).fillColor(MUT)
         .text(
           `Outback Helicopter Airwork NT · Page ${i + 1} of ${pages.count} · Generated ${new Date().toLocaleString('en-AU')}`,
           50, 787, { width: W, align: 'center' }
         );
    }

    doc.end();
  });
}

/* ── Escape HTML ──────────────────────────────────────────── */
function esc(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ── Start ────────────────────────────────────────────────── */
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Heli API listening on :${PORT}`));
