'use strict';

const path = require('path');
const express = require('express');
const { google } = require('googleapis');

const PORT = process.env.PORT || 3000;
const SHEET_ID = (process.env.GOOGLE_SHEET_ID || '').trim().replace(/^["']|["']$/g, '');
const SHEET_TAB = (process.env.GOOGLE_SHEET_TAB || 'Records').trim();
const DRY_RUN = !process.env.GOOGLE_CREDENTIALS_JSON || !SHEET_ID;

// Columns 1–17 follow the office ledger order exactly (two blank spacer columns kept).
// Columns 18–20 are extra details collected by the app.
// MATTER = "Payment For" from the receipt (one input, used on the receipt and in the ledger).
const HEADERS = [
  'DATE', 'CLIENT', 'SUBJECT', 'O/R DAT', 'O/R', 'MATTER', 'AMOUNT', 'DIS/INC', 'NETT',
  'SHARE(%)', 'SALARY', 'GOVIN', 'RAJ', '', 'RENT', '', 'TEL/IT',
  'LAWYER', 'MODE OF PAYMENT', 'TIMESTAMP (SGT)'
];
const HEADER_RANGE = 'A1:V1';                  // 2 wider than HEADERS so old extra headers are cleared
const HIGHLIGHT_COLS = [14, 16];               // RENT, TEL/IT (0-based) – yellow like the ledger
const MONEY_KEYS = ['disInc', 'share', 'salary', 'govin', 'raj', 'rent', 'telIt'];

const LAWYERS = ['Dora Raj', 'Govinda Raju'];
const LEAD_PCT = 70;                           // lead lawyer's share of AMOUNT; the other gets the rest
const MODES = ['PayNow', 'Cash', 'Others'];

/* ---------- Google Sheets client ---------- */
let sheets = null;
let headerChecked = false;

function getSheets() {
  if (sheets) return sheets;
  const creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON);
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: creds.client_email,
      private_key: String(creds.private_key).replace(/\\n/g, '\n')
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });
  sheets = google.sheets({ version: 'v4', auth });
  return sheets;
}

async function ensureHeader() {
  if (headerChecked) return;
  const api = getSheets();
  const meta = await api.spreadsheets.get({ spreadsheetId: SHEET_ID });
  let tab = meta.data.sheets.find(s => s.properties.title === SHEET_TAB);
  if (!tab) {
    const added = await api.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_TAB } } }] }
    });
    tab = { properties: added.data.replies[0].addSheet.properties };
  }
  const tabId = tab.properties.sheetId;

  const res = await api.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!${HEADER_RANGE}`
  });
  const current = (res.data.values && res.data.values[0]) || [];
  const norm = arr => HEADERS.map((_, i) => String(arr[i] ?? '')).join('|');
  if (norm(current) !== norm(HEADERS)) {
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!${HEADER_RANGE}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[...HEADERS, '', '']] }
    });
    // Bold header, freeze row 1, yellow RENT / TEL/IT headers
    await api.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: tabId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: HEADERS.length },
              cell: { userEnteredFormat: { textFormat: { bold: true } } },
              fields: 'userEnteredFormat.textFormat.bold'
            }
          },
          ...HIGHLIGHT_COLS.map(c => ({
            repeatCell: {
              range: { sheetId: tabId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: c, endColumnIndex: c + 1 },
              cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 0 } } },
              fields: 'userEnteredFormat.backgroundColor'
            }
          })),
          {
            updateSheetProperties: {
              properties: { sheetId: tabId, gridProperties: { frozenRowCount: 1 } },
              fields: 'gridProperties.frozenRowCount'
            }
          }
        ]
      }
    });
  }
  headerChecked = true;
}

async function appendRow(row) {
  if (DRY_RUN) {
    console.log('[DRY RUN] Row not sent to Google Sheets:', row);
    return;
  }
  await ensureHeader();
  await getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `'${SHEET_TAB}'!A1`,
    valueInputOption: 'RAW',          // text stored as-is (no formula execution); numbers stay numeric
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
}

/* ---------- Helpers ---------- */
const clean = (v, max = 300) => String(v ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);

// '' -> '' (blank cell); valid amount -> Number with 2 dp; invalid -> NaN
function money(v) {
  const s = clean(v, 20).replace(/[,\s$]/g, '');
  if (!s) return '';
  if (!/^\d+(\.\d{1,2})?$/.test(s)) return NaN;
  const n = Math.round(parseFloat(s) * 100) / 100;
  return n <= 999999999.99 ? n : NaN;
}

function sgtNow() {
  return new Date().toLocaleString('en-SG', {
    timeZone: 'Asia/Singapore', day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });
}

/* ---------- App ---------- */
const app = express();
app.use(express.json({ limit: '20kb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Real end-to-end check: reads the sheet and tests write access.
// Link for the EXCEL button on the receipt page
app.get('/api/config', (_req, res) => {
  res.json({ sheetUrl: SHEET_ID ? `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit` : null });
});

app.get('/healthz', async (_req, res) => {
  if (DRY_RUN) return res.json({ ok: true, sheets: 'dry-run', reason: 'GOOGLE_CREDENTIALS_JSON or GOOGLE_SHEET_ID not set' });
  let serviceAccount = 'unreadable';
  try { serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON).client_email; } catch (_) {}
  try {
    await ensureHeader();                       // needs read + write
    res.json({ ok: true, sheets: 'connected', serviceAccount, sheetId: SHEET_ID, tab: SHEET_TAB });
  } catch (err) {
    res.status(500).json({
      ok: false, sheets: 'error', serviceAccount, sheetId: SHEET_ID,
      error: err.message,
      fix: /permission/i.test(err.message)
        ? `Share the sheet with ${serviceAccount} as Editor`
        : /not found/i.test(err.message) ? 'Check GOOGLE_SHEET_ID' : 'Check credentials and that the Google Sheets API is enabled'
    });
  }
});

app.post('/api/records', async (req, res) => {
  const { case: c = {}, accounts: a = {}, receipt: r = {} } = req.body || {};

  const lawyer = clean(c.lawyer, 50);
  const mode = clean(r.mode, 20);
  const amount = money(r.amount);
  const m = Object.fromEntries(MONEY_KEYS.map(k => [k, money(a[k])]));
  const errors = [];

  if (!LAWYERS.includes(lawyer)) errors.push('Lawyer');
  if (!clean(c.slNo)) errors.push('Bill No');
  if (!clean(r.receivedFrom)) errors.push('Received from');
  if (amount === '' || Number.isNaN(amount) || amount <= 0) errors.push('Amount');
  if (!clean(r.paymentFor)) errors.push('Payment For');
  if (!MODES.includes(mode)) errors.push('Mode of Payment');
  if (mode === 'Others' && !clean(r.modeOther)) errors.push('Mode of Payment (Others)');
  MONEY_KEYS.forEach(k => { if (Number.isNaN(m[k])) errors.push(k); });
  if (typeof m.share === 'number' && m.share > 100) errors.push('Share (%) must be 0–100');
  if (errors.length) return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });

  const nett = Math.round((amount - (m.disInc || 0)) * 100) / 100;

  // Default GOVIN / RAJ split (only where left blank):
  //   /G (Govinda Raju) -> GOVIN 70%, RAJ 30%   |   /D (Dora Raj) -> RAJ 70%, GOVIN 30%
  // Worked in cents so the two parts always add up exactly to AMOUNT.
  const billNo = clean(c.slNo, 40);
  const govinLead = /\/G\s*$/i.test(billNo) || (!/\/D\s*$/i.test(billNo) && lawyer === 'Govinda Raju');
  const amountCents = Math.round(amount * 100);
  const leadCents = Math.round(amountCents * LEAD_PCT / 100);
  const autoGovin = (govinLead ? leadCents : amountCents - leadCents) / 100;
  const autoRaj = (govinLead ? amountCents - leadCents : leadCents) / 100;
  const cents = v => Math.round(v * 100);
  if (m.govin === '' && m.raj === '') {          // both blank -> 70/30 split
    m.govin = autoGovin; m.raj = autoRaj;
  } else if (m.govin === '') {                   // only Raj typed -> Govin gets the rest
    m.govin = Math.max(0, amountCents - cents(m.raj)) / 100;
  } else if (m.raj === '') {                     // only Govin typed -> Raj gets the rest
    m.raj = Math.max(0, amountCents - cents(m.govin)) / 100;
  }                                              // both typed -> keep as entered
  const receiptDate = clean(r.date, 30);

  const row = [
    receiptDate,                       // DATE
    clean(r.receivedFrom),             // CLIENT
    clean(c.caseCategory),             // SUBJECT
    receiptDate,                       // O/R DAT
    billNo,                            // O/R (Bill No)
    clean(r.paymentFor, 500),          // MATTER (= Payment For on the receipt)
    amount,                            // AMOUNT
    m.disInc,                          // DIS/INC
    nett,                              // NETT
    m.share,                           // SHARE(%)
    m.salary,                          // SALARY
    m.govin,                           // GOVIN
    m.raj,                             // RAJ
    '',                                // (spacer)
    m.rent,                            // RENT
    '',                                // (spacer)
    m.telIt,                           // TEL/IT
    lawyer,                            // LAWYER
    mode === 'Others' ? `Others – ${clean(r.modeOther, 100)}` : mode,   // MODE OF PAYMENT
    sgtNow()                           // TIMESTAMP (SGT)
  ];

  try {
    await appendRow(row);
    res.json({ ok: true });
  } catch (err) {
    console.error('Sheets append failed:', err.message);
    res.status(502).json({ ok: false, error: 'Could not save to Google Sheet. Please try again.' });
  }
});

app.listen(PORT, () => {
  console.log(`Raj Govin receipts running on port ${PORT}${DRY_RUN ? ' (DRY RUN – no Google credentials set)' : ''}`);
});
