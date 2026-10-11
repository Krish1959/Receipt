'use strict';

const path = require('path');
const express = require('express');
const { google } = require('googleapis');

const PORT = process.env.PORT || 3000;
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_TAB = process.env.GOOGLE_SHEET_TAB || 'Records';
const DRY_RUN = !process.env.GOOGLE_CREDENTIALS_JSON || !SHEET_ID;

const HEADERS = [
  'Timestamp (SGT)', 'Lawyer', 'S/L No', 'Case Category', 'Item 1', 'Item 2', 'Remarks',
  'Receipt Date', 'Bill No', 'Received From', 'Amount (S$)', 'Amount in Words',
  'Payment For', 'Mode of Payment'
];

const LAWYERS = ['Dora Raj', 'Govinda Raju'];
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
  const exists = meta.data.sheets.some(s => s.properties.title === SHEET_TAB);
  if (!exists) {
    await api.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_TAB } } }] }
    });
  }
  const res = await api.spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!A1:N1`
  });
  if (!res.data.values || !res.data.values.length) {
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!A1`,
      valueInputOption: 'RAW',
      requestBody: { values: [HEADERS] }
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
    valueInputOption: 'RAW',          // stored as plain text – no formula execution
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
}

/* ---------- Helpers ---------- */
const clean = (v, max = 300) => String(v ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);

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

app.get('/healthz', (_req, res) => res.json({ ok: true, sheets: DRY_RUN ? 'dry-run' : 'connected' }));

app.post('/api/records', async (req, res) => {
  const { case: c = {}, receipt: r = {} } = req.body || {};

  const lawyer = clean(c.lawyer, 50);
  const mode = clean(r.mode, 20);
  const amount = Number(r.amount);
  const errors = [];

  if (!LAWYERS.includes(lawyer)) errors.push('Lawyer');
  if (!clean(c.slNo)) errors.push('S/L No');
  if (!clean(r.receivedFrom)) errors.push('Received from');
  if (!Number.isFinite(amount) || amount <= 0 || amount > 999999999.99) errors.push('Amount');
  if (!clean(r.paymentFor)) errors.push('Payment For');
  if (!MODES.includes(mode)) errors.push('Mode of Payment');
  if (mode === 'Others' && !clean(r.modeOther)) errors.push('Mode of Payment (Others)');
  if (errors.length) return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });

  const row = [
    sgtNow(),
    lawyer,
    clean(c.slNo, 40),
    clean(c.caseCategory),
    clean(c.item1, 1000),
    clean(c.item2, 1000),
    clean(c.remarks, 1000),
    clean(r.date, 30),
    clean(r.billNo, 10),
    clean(r.receivedFrom),
    amount.toFixed(2),
    clean(r.amountWords, 500),
    clean(r.paymentFor, 500),
    mode === 'Others' ? `Others – ${clean(r.modeOther, 100)}` : mode
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
