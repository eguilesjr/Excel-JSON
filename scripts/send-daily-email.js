'use strict';

const nodemailer = require('nodemailer');
const https = require('https');

const GMAIL_USER = process.env.GMAIL_USER || 'eguilesjr@gmail.com';
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;
const TO_EMAIL = process.env.TO_EMAIL || 'eguilesjr@gmail.com';
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const REPO = 'eguilesjr/excel-json';

function httpsGet(options) {
  return new Promise((resolve, reject) => {
    const req = https.get(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body }));
    });
    req.on('error', reject);
  });
}

async function fetchTrackerData() {
  const result = await httpsGet({
    hostname: 'api.github.com',
    path: `/repos/${REPO}/contents/data/tracker.json`,
    headers: {
      'Authorization': `token ${GITHUB_TOKEN}`,
      'Accept': 'application/vnd.github.v3+json',
      'User-Agent': 'middlesex-tracker-action'
    }
  });

  if (result.status !== 200) {
    throw new Error(`GitHub API returned ${result.status}: ${result.body}`);
  }

  const file = JSON.parse(result.body);
  const content = Buffer.from(file.content, 'base64').toString('utf8');
  return JSON.parse(content);
}

function fmtDate(d) {
  if (!d) return '—';
  const [y, m, day] = String(d).split('T')[0].split('-');
  return `${m}/${day}/${y}`;
}

function getStatus(frame) {
  const today = new Date().toISOString().split('T')[0];
  if (frame.installCompleted) return 'complete';
  if (frame.installStarted)   return 'started';
  if (frame.shipmentReceived) return 'received';
  if (
    frame.deliveryDate &&
    frame.deliveryDate !== 'N/A' &&
    frame.deliveryDate < today &&
    frame.colfabStatus !== 'Not by Colfab'
  ) return 'overdue';
  return 'pending';
}

function applyUpdates(data) {
  const fu = data.frameUpdates || {};
  const cu = data.cassetteUpdates || {};

  const frames = (data.frames || []).map(f => ({ ...f, ...(fu[f.id] || {}) }));
  const cassettes = (data.cassettes || []).map(c => ({ ...c, ...(cu[c.drawing] || {}) }));
  return { frames, cassettes, lastUpdated: data.lastUpdated };
}

function buildEmail({ frames, cassettes, lastUpdated }) {
  const today = new Date().toLocaleDateString('en-US', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });

  const total    = frames.length;
  const complete = frames.filter(f => getStatus(f) === 'complete').length;
  const started  = frames.filter(f => getStatus(f) === 'started').length;
  const received = frames.filter(f => getStatus(f) === 'received').length;
  const overdue  = frames.filter(f => getStatus(f) === 'overdue');
  const inProg   = frames.filter(f => getStatus(f) === 'started');
  const pct      = Math.round((complete / total) * 100);

  const statusColor = { complete:'#2d8a4e', started:'#1a6fb5', received:'#7c3aed', overdue:'#cc3300', pending:'#6b7280' };
  const statusLabel = { complete:'Complete', started:'In Progress', received:'Received', overdue:'Overdue', pending:'Pending' };

  const rowBg = (f, i) => {
    const s = getStatus(f);
    if (s === 'complete') return '#f0faf4';
    if (s === 'overdue')  return '#fff5f5';
    return i % 2 === 0 ? '#ffffff' : '#f9fafb';
  };

  const overdueSection = overdue.length === 0 ? '' : `
  <div style="padding:20px;border-bottom:1px solid #e5e7eb;background:#fffbfb;">
    <div style="font-size:15px;font-weight:700;color:#cc3300;margin-bottom:10px;">
      Overdue — Delivery Passed, Not Yet Received (${overdue.length})
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:11px;font-weight:700;text-transform:uppercase;">
        <td style="padding:6px 8px;">#</td><td style="padding:6px 8px;">Type</td>
        <td style="padding:6px 8px;">Area</td><td style="padding:6px 8px;">Location</td>
        <td style="padding:6px 8px;">Delivery</td><td style="padding:6px 8px;">Notes</td>
      </tr>
      ${overdue.map((f, i) => `
      <tr style="background:${i%2===0?'#fff':'#fafafa'};">
        <td style="padding:7px 8px;font-weight:600;">${f.id}</td>
        <td style="padding:7px 8px;font-weight:700;">${f.frameType}</td>
        <td style="padding:7px 8px;">Area ${f.area}</td>
        <td style="padding:7px 8px;">${f.location}</td>
        <td style="padding:7px 8px;color:#cc3300;font-weight:600;">${fmtDate(f.deliveryDate)}</td>
        <td style="padding:7px 8px;color:#6b7280;font-size:12px;">${f.graboyesNotes || f.colfabNotes || '—'}</td>
      </tr>`).join('')}
    </table>
  </div>`;

  const inProgSection = inProg.length === 0 ? '' : `
  <div style="padding:20px;border-bottom:1px solid #e5e7eb;">
    <div style="font-size:15px;font-weight:700;color:#1a6fb5;margin-bottom:10px;">
      In Progress (${inProg.length})
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:11px;font-weight:700;text-transform:uppercase;">
        <td style="padding:6px 8px;">#</td><td style="padding:6px 8px;">Type</td>
        <td style="padding:6px 8px;">Area</td><td style="padding:6px 8px;">Started</td>
        <td style="padding:6px 8px;">Notes</td>
      </tr>
      ${inProg.map((f, i) => `
      <tr style="background:${i%2===0?'#fff':'#fafafa'};">
        <td style="padding:7px 8px;font-weight:600;">${f.id}</td>
        <td style="padding:7px 8px;font-weight:700;">${f.frameType}</td>
        <td style="padding:7px 8px;">Area ${f.area}</td>
        <td style="padding:7px 8px;">${fmtDate(f.installStarted)}</td>
        <td style="padding:7px 8px;color:#6b7280;font-size:12px;">${f.graboyesNotes || '—'}</td>
      </tr>`).join('')}
    </table>
  </div>`;

  const cassetteRows = cassettes.map((c, i) => {
    const s = c.dateInstalled ? 'Installed' : c.dateShipped ? 'Shipped' : c.qtyFabricated ? 'Fabricated' : 'Pending';
    const sc = c.dateInstalled ? '#2d8a4e' : c.dateShipped ? '#1a6fb5' : c.qtyFabricated ? '#7c3aed' : '#6b7280';
    return `<tr style="background:${i%2===0?'#fff':'#fafafa'};">
      <td style="padding:5px 6px;font-weight:700;">${c.drawing}</td>
      <td style="padding:5px 6px;">${c.sheetNum}</td>
      <td style="padding:5px 6px;">${c.cassetteQty}</td>
      <td style="padding:5px 6px;">${c.glassPanels + c.doors}</td>
      <td style="padding:5px 6px;">${c.qtyFabricated || '—'}</td>
      <td style="padding:5px 6px;">${fmtDate(c.dateShipped)}</td>
      <td style="padding:5px 6px;">${fmtDate(c.dateInstalled)}</td>
      <td style="padding:5px 6px;color:${sc};font-weight:600;font-size:11px;">${s}</td>
    </tr>`;
  }).join('');

  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
</head>
<body style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f2f5;margin:0;padding:16px;">
<div style="max-width:680px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1);">

  <div style="background:#1a6fb5;color:#fff;padding:20px 24px;">
    <div style="font-size:20px;font-weight:700;">Middlesex Frame Tracker</div>
    <div style="font-size:13px;opacity:.85;margin-top:4px;">${today}</div>
  </div>

  <div style="padding:20px 24px;border-bottom:1px solid #e5e7eb;">
    <div style="font-size:16px;font-weight:700;margin-bottom:12px;">Daily Summary</div>
    <table style="width:100%;border-collapse:separate;border-spacing:8px 0;">
      <tr>
        <td style="background:#e8f5ec;border-radius:8px;padding:14px;text-align:center;width:25%;">
          <div style="font-size:28px;font-weight:700;color:#2d8a4e;">${complete}</div>
          <div style="font-size:11px;color:#2d8a4e;font-weight:700;margin-top:2px;">COMPLETE</div>
        </td>
        <td style="background:#e8f0f8;border-radius:8px;padding:14px;text-align:center;width:25%;">
          <div style="font-size:28px;font-weight:700;color:#1a6fb5;">${started}</div>
          <div style="font-size:11px;color:#1a6fb5;font-weight:700;margin-top:2px;">IN PROGRESS</div>
        </td>
        <td style="background:#f3e8ff;border-radius:8px;padding:14px;text-align:center;width:25%;">
          <div style="font-size:28px;font-weight:700;color:#7c3aed;">${received}</div>
          <div style="font-size:11px;color:#7c3aed;font-weight:700;margin-top:2px;">RECEIVED</div>
        </td>
        <td style="background:${overdue.length > 0 ? '#fff0ed' : '#f3f4f6'};border-radius:8px;padding:14px;text-align:center;width:25%;">
          <div style="font-size:28px;font-weight:700;color:${overdue.length > 0 ? '#cc3300' : '#6b7280'};">${overdue.length}</div>
          <div style="font-size:11px;color:${overdue.length > 0 ? '#cc3300' : '#6b7280'};font-weight:700;margin-top:2px;">OVERDUE</div>
        </td>
      </tr>
    </table>
    <div style="margin-top:12px;background:#f3f4f6;border-radius:6px;padding:10px 14px;font-size:13px;color:#374151;">
      Overall: <strong>${complete} of ${total} lines complete (${pct}%)</strong>
      ${lastUpdated ? ` &nbsp;·&nbsp; Last updated: ${new Date(lastUpdated).toLocaleString('en-US')}` : ' &nbsp;·&nbsp; No updates recorded yet'}
    </div>
  </div>

  ${overdueSection}
  ${inProgSection}

  <div style="padding:20px 24px;border-bottom:1px solid #e5e7eb;">
    <div style="font-size:15px;font-weight:700;margin-bottom:10px;">All Frames — Full Status</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
        <td style="padding:5px 6px;">#</td>
        <td style="padding:5px 6px;">Type</td>
        <td style="padding:5px 6px;">Area</td>
        <td style="padding:5px 6px;">Qty</td>
        <td style="padding:5px 6px;">Colfab Status</td>
        <td style="padding:5px 6px;">Rcvd</td>
        <td style="padding:5px 6px;">Started</td>
        <td style="padding:5px 6px;">Done</td>
        <td style="padding:5px 6px;">Notes</td>
      </tr>
      ${frames.map((f, i) => {
        const s = getStatus(f);
        const color = statusColor[s];
        return `<tr style="background:${rowBg(f, i)};">
          <td style="padding:5px 6px;font-weight:600;">${f.id}</td>
          <td style="padding:5px 6px;font-weight:700;">${f.frameType}</td>
          <td style="padding:5px 6px;">Area ${f.area}</td>
          <td style="padding:5px 6px;">${f.qty}</td>
          <td style="padding:5px 6px;font-size:10px;">${f.colfabStatus}</td>
          <td style="padding:5px 6px;">${fmtDate(f.shipmentReceived)}</td>
          <td style="padding:5px 6px;">${fmtDate(f.installStarted)}</td>
          <td style="padding:5px 6px;font-weight:${f.installCompleted?'700':'400'};color:${f.installCompleted?'#2d8a4e':'inherit'};">${fmtDate(f.installCompleted)}</td>
          <td style="padding:5px 6px;color:#6b7280;font-size:11px;">${f.graboyesNotes || '—'}</td>
        </tr>`;
      }).join('')}
    </table>
  </div>

  <div style="padding:20px 24px;">
    <div style="font-size:15px;font-weight:700;margin-bottom:10px;">Cassettes — Status</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
        <td style="padding:5px 6px;">Drawing</td>
        <td style="padding:5px 6px;">Sheet</td>
        <td style="padding:5px 6px;">Qty</td>
        <td style="padding:5px 6px;">Items</td>
        <td style="padding:5px 6px;">Fabricated</td>
        <td style="padding:5px 6px;">Shipped</td>
        <td style="padding:5px 6px;">Installed</td>
        <td style="padding:5px 6px;">Status</td>
      </tr>
      ${cassetteRows}
    </table>
  </div>

  <div style="padding:14px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;text-align:center;">
    Middlesex Curtainwall &nbsp;·&nbsp; Automated daily report &nbsp;·&nbsp; eguilesjr/excel-json
  </div>
</div>
</body></html>`;
}

async function main() {
  if (!GMAIL_APP_PASSWORD) {
    throw new Error('GMAIL_APP_PASSWORD secret is not set. See SETUP.md for instructions.');
  }

  console.log('Fetching tracker data from GitHub...');
  let data;
  try {
    data = await fetchTrackerData();
    console.log(`Loaded: ${data.frames?.length ?? 0} frames, ${data.cassettes?.length ?? 0} cassettes`);
  } catch (err) {
    console.error('Could not fetch tracker data:', err.message);
    console.log('Sending email with empty data...');
    data = { frames: [], cassettes: [], frameUpdates: {}, cassetteUpdates: {}, lastUpdated: null };
  }

  const { frames, cassettes, lastUpdated } = applyUpdates(data);
  const html = buildEmail({ frames, cassettes, lastUpdated });

  const todayShort = new Date().toLocaleDateString('en-US', {
    weekday: 'short', month: 'short', day: 'numeric'
  });

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: GMAIL_USER, pass: GMAIL_APP_PASSWORD }
  });

  await transporter.sendMail({
    from: `"Middlesex Tracker" <${GMAIL_USER}>`,
    to: TO_EMAIL,
    subject: `Frame Tracker – ${todayShort}`,
    html
  });

  console.log(`Daily report sent to ${TO_EMAIL}`);
}

main().catch(err => {
  console.error(err.message);
  process.exit(1);
});
