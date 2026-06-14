'use strict';
/* ============================================================
   Outback Helicopter Airwork NT — Flight Paperwork API
   POST /api/send  →  generate PDF, send via Office 365, file to OneDrive
   All via Microsoft Graph API — one set of credentials for everything
   ============================================================ */

const express     = require('express');
const cors        = require('cors');
const PDFDocument = require('pdfkit');
const fs          = require('fs');
const path        = require('path');
const { ClientSecretCredential } = require('@azure/identity');

/* ── ACST helpers (Australia/Darwin = UTC+9:30, no DST) ─────── */
const ACST_TZ = 'Australia/Darwin';
function acstDate(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, day: '2-digit', month: 'long', year: 'numeric' }).format(new Date(ts)); }
function acstTime(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(ts)); }
function acstFull(ts)  { return new Intl.DateTimeFormat('en-AU', { timeZone: ACST_TZ, day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit', hour12: true }).format(new Date(ts)); }

const app = express();
app.use(cors());
app.use(express.json({ limit: '20mb' }));

/* ── Health check ─────────────────────────────────────────── */
app.get('/health', (_req, res) => res.json({ ok: true, ts: new Date().toISOString() }));

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
async function buildPDF(bundle) {
  /* Pre-fetch satellite map image if location + API key available */
  let mapImageBuffer = null;
  if (bundle.location?.lat && process.env.GOOGLE_MAPS_KEY) {
    try {
      const { lat, lng } = bundle.location;
      const mapUrl = `https://maps.googleapis.com/maps/api/staticmap?center=${lat},${lng}&zoom=15&size=480x280&maptype=satellite&markers=color:red%7C${lat},${lng}&key=${process.env.GOOGLE_MAPS_KEY}`;
      const mapRes = await fetch(mapUrl);
      if (mapRes.ok) mapImageBuffer = Buffer.from(await mapRes.arrayBuffer());
    } catch (e) { console.error('Map image fetch failed (non-fatal):', e.message); }
  }

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

    /* ── Section heading ── */
    const secHead = title => {
      doc.moveDown(0.6);
      const y = doc.y;
      doc.rect(50, y, W, 20).fill('#E8EBF5');
      doc.fillColor(NAVY).font('Helvetica-Bold').fontSize(9.5)
         .text(title, 58, y + 5, { width: W - 16, lineBreak: false });
      doc.y = y + 26;
      doc.fillColor('#1C1F28');
    };

    /* ── Key/value row ── */
    const kv = (k, v, color) => {
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
    kv('ARN',      bundle.sms?.values?.pilotArn);
    kv('Date',     acstDate(ts) + ' (ACST)');
    kv('Time',     acstTime(ts) + ' (ACST)');
    if (bundle.sms?.values?.trainerName) kv('Trainer', bundle.sms.values.trainerName);

    /* ── Location ── */
    if (bundle.location?.lat) {
      const { lat, lng, accuracy } = bundle.location;
      secHead('LOCATION AT SUBMISSION');
      kv('Latitude',  lat.toFixed(6) + '°');
      kv('Longitude', lng.toFixed(6) + '°');
      if (accuracy) kv('Accuracy', accuracy + ' m');
      if (mapImageBuffer) {
        doc.moveDown(0.4);
        const imgY = doc.y;
        if (imgY + 180 > 750) doc.addPage();
        doc.image(mapImageBuffer, 50, doc.y, { width: W, height: 180 });
        doc.y += 188;
      }
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
        const y = doc.y;
        doc.font('Helvetica').fontSize(9.5).fillColor(v ? '#15803D' : '#B91C1C')
           .text((v ? '✓  ' : '✗  ') + 'Section ' + (parseInt(i) + 1), 50, y, { width: W, lineBreak: false });
        doc.y = y + 16;
      });
    }

    /* ── Passengers ── */
    if ((bundle.pax || []).length) {
      secHead('PASSENGERS');
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
