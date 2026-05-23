'use strict';

const nodemailer = require('nodemailer');
const https = require('https');

const GMAIL_USER         = process.env.GMAIL_USER || 'eguilesjr@gmail.com';
const GMAIL_APP_PASSWORD = process.env.GMAIL_APP_PASSWORD;
const TO_EMAIL           = process.env.TO_EMAIL   || 'eguilesjr@gmail.com';
const SUPABASE_URL       = process.env.SUPABASE_URL;
const SUPABASE_ANON_KEY  = process.env.SUPABASE_ANON_KEY;

// ── Reference data (from Excel — never changes) ──────────────────────────────
const REF_FRAMES = [
  {id:1,area:"D",location:"Press",floor:"3rd",frameType:"CW-A",qty:4,colfabStatus:"In Production",deliveryDate:"2026-05-15",colfabNotes:null},
  {id:2,area:"D",location:"Press",floor:"3rd",frameType:"CW-B1",qty:6,colfabStatus:"In Production",deliveryDate:"2026-05-18",colfabNotes:null},
  {id:3,area:"D",location:"Press",floor:"3rd",frameType:"CW-C1",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-18",colfabNotes:null},
  {id:4,area:"D",location:"Press",floor:"3rd",frameType:"CW-F",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-18",colfabNotes:null},
  {id:5,area:"D",location:"Press",floor:"3rd",frameType:"SF-F",qty:1,colfabStatus:"Delivered",deliveryDate:"2026-05-11",colfabNotes:"Without Door Frames"},
  {id:6,area:"D",location:"Press",floor:"3rd",frameType:"SF-E",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:7,area:"D",location:"Press",floor:"3rd",frameType:"SF-H",qty:1,colfabStatus:"Not by Colfab",deliveryDate:"N/A",colfabNotes:"Door Only"},
  {id:8,area:"C",location:"Press",floor:"3rd",frameType:"CW-B1",qty:3,colfabStatus:"In Production",deliveryDate:"2026-05-20",colfabNotes:null},
  {id:9,area:"C",location:"Press",floor:"3rd",frameType:"CW-B2",qty:2,colfabStatus:"In Production",deliveryDate:"2026-05-20",colfabNotes:null},
  {id:10,area:"C",location:"Press",floor:"3rd",frameType:"CW-B3",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-20",colfabNotes:null},
  {id:11,area:"C",location:"Press",floor:"3rd",frameType:"CW-C",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-22",colfabNotes:null},
  {id:12,area:"C",location:"Press",floor:"3rd",frameType:"CW-A",qty:5,colfabStatus:"In Production",deliveryDate:"2026-05-22",colfabNotes:null},
  {id:13,area:"C",location:"Press",floor:"3rd",frameType:"CW-E2",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-27",colfabNotes:null},
  {id:14,area:"C",location:"Press",floor:"3rd",frameType:"CW-E1",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-27",colfabNotes:null},
  {id:15,area:"C",location:"Press",floor:"3rd",frameType:"CW-D",qty:1,colfabStatus:"In Production",deliveryDate:"2026-05-29",colfabNotes:null},
  {id:16,area:"C",location:"Concourse",floor:"2nd",frameType:"SF-C3",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:17,area:"C",location:"Concourse",floor:"2nd",frameType:"SF-C2",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:18,area:"C",location:"Concourse",floor:"2nd",frameType:"SF-C1",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:19,area:"C",location:"Ground",floor:"Ground",frameType:"SF-I",qty:1,colfabStatus:"Not by Colfab",deliveryDate:"N/A",colfabNotes:"Door Only"},
  {id:20,area:"C",location:"Ground",floor:"Ground",frameType:"SF-G",qty:1,colfabStatus:"Not by Colfab",deliveryDate:"N/A",colfabNotes:"Door Only"},
  {id:21,area:"B",location:"Press",floor:"3rd",frameType:"CW-B1",qty:9,colfabStatus:"In Production",deliveryDate:"2026-06-03",colfabNotes:null},
  {id:22,area:"B",location:"Press",floor:"3rd",frameType:"SF-D3",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:23,area:"B",location:"Press",floor:"3rd",frameType:"SF-D2",qty:3,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:24,area:"B",location:"Press",floor:"3rd",frameType:"SF-D1",qty:1,colfabStatus:"Complete",deliveryDate:"2026-05-14",colfabNotes:"Without Door Frames"},
  {id:25,area:"B",location:"Concourse",floor:"2nd",frameType:"SF-B",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-05",colfabNotes:null},
  {id:26,area:"A",location:"Concourse",floor:"2nd",frameType:"CW-A",qty:6,colfabStatus:"In Production",deliveryDate:"2026-06-08",colfabNotes:null},
  {id:27,area:"A",location:"Concourse",floor:"2nd",frameType:"CW-B1",qty:2,colfabStatus:"In Production",deliveryDate:"2026-06-08",colfabNotes:null},
  {id:28,area:"A",location:"Concourse",floor:"2nd",frameType:"SF-A",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-08",colfabNotes:null},
  {id:29,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-K",qty:2,colfabStatus:"In Production",deliveryDate:"2026-06-10",colfabNotes:null},
  {id:30,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-L",qty:2,colfabStatus:"In Production",deliveryDate:"2026-06-10",colfabNotes:null},
  {id:31,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-G",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-10",colfabNotes:null},
  {id:32,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-J",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-12",colfabNotes:null},
  {id:33,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-H",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-12",colfabNotes:null},
  {id:34,area:"A",location:"Gateway",floor:"Gateway",frameType:"CW-I",qty:1,colfabStatus:"In Production",deliveryDate:"2026-06-12",colfabNotes:null}
];

const REF_CASSETTES = [
  {drawing:"CW-1",sheetNum:"201",cassetteQty:1,glassPanels:15,doors:2,notes:"Full-height clearwall. 14 fixed glass + 1 transom + 1 door pair (X100A)."},
  {drawing:"CW-2",sheetNum:"202",cassetteQty:1,glassPanels:12,doors:0,notes:"Long horizontal clearwall. All GLA-01A. No doors."},
  {drawing:"CW-3",sheetNum:"203",cassetteQty:10,glassPanels:1,doors:0,notes:"Single-panel cassette GLA-01A, built 10 times."},
  {drawing:"CW-4",sheetNum:"203",cassetteQty:1,glassPanels:1,doors:0,notes:"Single-panel cassette, GLA-01A."},
  {drawing:"CW-5",sheetNum:"203",cassetteQty:1,glassPanels:1,doors:0,notes:"Single-panel cassette, GLA-01A."},
  {drawing:"CW-6",sheetNum:"204",cassetteQty:1,glassPanels:15,doors:2,notes:"Mirror of CW-1. Door pair X100B in bottom row."},
  {drawing:"CW-7",sheetNum:"205",cassetteQty:1,glassPanels:24,doors:4,notes:"Two-section clearwall. Complex elevation — verify count."},
  {drawing:"CW-8",sheetNum:"206",cassetteQty:1,glassPanels:1,doors:0,notes:"Single-panel cassette, GLA-01A."},
  {drawing:"CW-9",sheetNum:"206",cassetteQty:1,glassPanels:3,doors:0,notes:"Three-panel angled clearwall. All GLA-01."},
  {drawing:"CW-10",sheetNum:"207",cassetteQty:1,glassPanels:12,doors:2,notes:"Two-section clearwall. Top + bottom sections, 2 single doors."},
  {drawing:"CW-11",sheetNum:"208",cassetteQty:1,glassPanels:3,doors:0,notes:"Three-panel angled clearwall. All GLA-01."}
];

// ── Supabase REST fetch ───────────────────────────────────────────────────────
function httpsGet(options) {
  return new Promise((resolve, reject) => {
    https.get(options, (res) => {
      let body = '';
      res.on('data', chunk => body += chunk);
      res.on('end', () => resolve({ status: res.statusCode, body }));
    }).on('error', reject);
  });
}

async function fetchTable(table) {
  const base   = new URL(SUPABASE_URL);
  const result = await httpsGet({
    hostname: base.hostname,
    path:     `/rest/v1/${table}?select=*`,
    headers:  {
      'apikey':        SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
    }
  });
  if (result.status !== 200) {
    throw new Error(`Supabase /${table} returned ${result.status}: ${result.body}`);
  }
  return JSON.parse(result.body);
}

// ── Merge Supabase rows with reference data ───────────────────────────────────
async function getTrackerData() {
  const [frameRows, cassetteRows] = await Promise.all([
    fetchTable('frame_updates'),
    fetchTable('cassette_updates')
  ]);

  const frameMap = {};
  frameRows.forEach(r => {
    frameMap[r.id] = {
      shipmentReceived: r.shipment_received,
      installStarted:   r.install_started,
      installCompleted: r.install_completed,
      graboyesNotes:    r.graboyes_notes
    };
  });

  const cassetteMap = {};
  cassetteRows.forEach(r => {
    cassetteMap[r.drawing] = {
      qtyFabricated: r.qty_fabricated,
      dateShipped:   r.date_shipped,
      dateInstalled: r.date_installed
    };
  });

  const timestamps = frameRows.map(r => r.updated_at).filter(Boolean);
  const lastUpdated = timestamps.length
    ? new Date(Math.max(...timestamps.map(t => new Date(t).getTime()))).toISOString()
    : null;

  return {
    frames:      REF_FRAMES.map(f    => ({ ...f,    ...(frameMap[f.id]           || {}) })),
    cassettes:   REF_CASSETTES.map(c => ({ ...c,    ...(cassetteMap[c.drawing]   || {}) })),
    lastUpdated
  };
}

// ── Helpers ───────────────────────────────────────────────────────────────────
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
  if (frame.deliveryDate && frame.deliveryDate !== 'N/A' &&
      frame.deliveryDate < today && frame.colfabStatus !== 'Not by Colfab') return 'overdue';
  return 'pending';
}

// ── Build HTML email ──────────────────────────────────────────────────────────
function buildEmail({ frames, cassettes, lastUpdated }) {
  const today    = new Date().toLocaleDateString('en-US', { weekday:'long', year:'numeric', month:'long', day:'numeric' });
  const total    = frames.length;
  const complete = frames.filter(f => getStatus(f) === 'complete').length;
  const started  = frames.filter(f => getStatus(f) === 'started').length;
  const received = frames.filter(f => getStatus(f) === 'received').length;
  const overdue  = frames.filter(f => getStatus(f) === 'overdue');
  const inProg   = frames.filter(f => getStatus(f) === 'started');
  const pct      = Math.round((complete / total) * 100);

  const rowBg = (f, i) => {
    const s = getStatus(f);
    if (s === 'complete') return '#f0faf4';
    if (s === 'overdue')  return '#fff5f5';
    return i % 2 === 0 ? '#ffffff' : '#f9fafb';
  };

  const overdueSection = overdue.length === 0 ? '' : `
  <div style="padding:20px 24px;border-bottom:1px solid #e5e7eb;background:#fffbfb;">
    <div style="font-size:15px;font-weight:700;color:#cc3300;margin-bottom:10px;">
      Overdue — Delivery Passed, Not Yet Received (${overdue.length})
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
        <td style="padding:6px 8px;">#</td><td style="padding:6px 8px;">Type</td>
        <td style="padding:6px 8px;">Area</td><td style="padding:6px 8px;">Delivery</td>
        <td style="padding:6px 8px;">Notes</td>
      </tr>
      ${overdue.map((f, i) => `
      <tr style="background:${i%2===0?'#fff':'#fafafa'};">
        <td style="padding:7px 8px;font-weight:600;">${f.id}</td>
        <td style="padding:7px 8px;font-weight:700;">${f.frameType}</td>
        <td style="padding:7px 8px;">Area ${f.area}</td>
        <td style="padding:7px 8px;color:#cc3300;font-weight:600;">${fmtDate(f.deliveryDate)}</td>
        <td style="padding:7px 8px;color:#6b7280;font-size:12px;">${f.graboyesNotes || f.colfabNotes || '—'}</td>
      </tr>`).join('')}
    </table>
  </div>`;

  const inProgSection = inProg.length === 0 ? '' : `
  <div style="padding:20px 24px;border-bottom:1px solid #e5e7eb;">
    <div style="font-size:15px;font-weight:700;color:#1a6fb5;margin-bottom:10px;">In Progress (${inProg.length})</div>
    <table style="width:100%;border-collapse:collapse;font-size:13px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
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
    const s  = c.dateInstalled ? 'Installed' : c.dateShipped ? 'Shipped' : c.qtyFabricated ? 'Fabricated' : 'Pending';
    const sc = c.dateInstalled ? '#2d8a4e'   : c.dateShipped ? '#1a6fb5' : c.qtyFabricated ? '#7c3aed'    : '#6b7280';
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
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
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
        <td style="background:${overdue.length>0?'#fff0ed':'#f3f4f6'};border-radius:8px;padding:14px;text-align:center;width:25%;">
          <div style="font-size:28px;font-weight:700;color:${overdue.length>0?'#cc3300':'#6b7280'};">${overdue.length}</div>
          <div style="font-size:11px;color:${overdue.length>0?'#cc3300':'#6b7280'};font-weight:700;margin-top:2px;">OVERDUE</div>
        </td>
      </tr>
    </table>
    <div style="margin-top:12px;background:#f3f4f6;border-radius:6px;padding:10px 14px;font-size:13px;color:#374151;">
      Overall: <strong>${complete} of ${total} lines complete (${pct}%)</strong>
      ${lastUpdated ? ` &nbsp;&middot;&nbsp; Last update: ${new Date(lastUpdated).toLocaleString('en-US')}` : ''}
    </div>
  </div>

  ${overdueSection}
  ${inProgSection}

  <div style="padding:20px 24px;border-bottom:1px solid #e5e7eb;">
    <div style="font-size:15px;font-weight:700;margin-bottom:10px;">All Frames</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
        <td style="padding:5px 6px;">#</td><td style="padding:5px 6px;">Type</td>
        <td style="padding:5px 6px;">Area</td><td style="padding:5px 6px;">Qty</td>
        <td style="padding:5px 6px;">Colfab</td><td style="padding:5px 6px;">Rcvd</td>
        <td style="padding:5px 6px;">Started</td><td style="padding:5px 6px;">Done</td>
        <td style="padding:5px 6px;">Notes</td>
      </tr>
      ${frames.map((f, i) => `<tr style="background:${rowBg(f,i)};">
        <td style="padding:5px 6px;font-weight:600;">${f.id}</td>
        <td style="padding:5px 6px;font-weight:700;">${f.frameType}</td>
        <td style="padding:5px 6px;">Area ${f.area}</td>
        <td style="padding:5px 6px;">${f.qty}</td>
        <td style="padding:5px 6px;font-size:10px;">${f.colfabStatus}</td>
        <td style="padding:5px 6px;">${fmtDate(f.shipmentReceived)}</td>
        <td style="padding:5px 6px;">${fmtDate(f.installStarted)}</td>
        <td style="padding:5px 6px;font-weight:${f.installCompleted?'700':'400'};color:${f.installCompleted?'#2d8a4e':'inherit'};">${fmtDate(f.installCompleted)}</td>
        <td style="padding:5px 6px;color:#6b7280;font-size:11px;">${f.graboyesNotes||'—'}</td>
      </tr>`).join('')}
    </table>
  </div>

  <div style="padding:20px 24px;">
    <div style="font-size:15px;font-weight:700;margin-bottom:10px;">Cassettes</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px;">
      <tr style="background:#f3f4f6;color:#6b7280;font-size:10px;font-weight:700;text-transform:uppercase;">
        <td style="padding:5px 6px;">Drawing</td><td style="padding:5px 6px;">Sheet</td>
        <td style="padding:5px 6px;">Qty</td><td style="padding:5px 6px;">Items</td>
        <td style="padding:5px 6px;">Fabricated</td><td style="padding:5px 6px;">Shipped</td>
        <td style="padding:5px 6px;">Installed</td><td style="padding:5px 6px;">Status</td>
      </tr>
      ${cassetteRows}
    </table>
  </div>

  <div style="padding:14px 24px;background:#f9fafb;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af;text-align:center;">
    Middlesex Curtainwall &nbsp;&middot;&nbsp; Automated daily report &nbsp;&middot;&nbsp; eguilesjr/excel-json
  </div>
</div>
</body></html>`;
}

// ── Main ──────────────────────────────────────────────────────────────────────
async function main() {
  if (!GMAIL_APP_PASSWORD) throw new Error('GMAIL_APP_PASSWORD secret not set.');
  if (!SUPABASE_URL)       throw new Error('SUPABASE_URL secret not set.');
  if (!SUPABASE_ANON_KEY)  throw new Error('SUPABASE_ANON_KEY secret not set.');

  console.log('Fetching data from Supabase…');
  const data = await getTrackerData();
  console.log(`Loaded ${data.frames.length} frames, ${data.cassettes.length} cassettes`);

  const html      = buildEmail(data);
  const todayShort = new Date().toLocaleDateString('en-US', { weekday:'short', month:'short', day:'numeric' });

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: { user: GMAIL_USER, pass: GMAIL_APP_PASSWORD }
  });

  await transporter.sendMail({
    from:    `"Middlesex Tracker" <${GMAIL_USER}>`,
    to:      TO_EMAIL,
    subject: `Frame Tracker – ${todayShort}`,
    html
  });

  console.log(`Daily report sent to ${TO_EMAIL}`);
}

main().catch(err => { console.error(err.message); process.exit(1); });
