'use strict';

const path = require('path');
const express = require('express');

const g = require('./lib/google');
const auth = require('./lib/auth');
const invoiceRoutes = require('./lib/invoice-routes');
const pettycashRoutes = require('./lib/pettycash-routes');

const PORT = process.env.PORT || 3000;
const SHEET_ID = (process.env.GOOGLE_SHEET_ID || '').trim().replace(/^["']|["']$/g, '');
const SHEET_TAB = (process.env.GOOGLE_SHEET_TAB || 'Records').trim();
const DRY_RUN = g.authMode() === 'none' || !SHEET_ID;

// Columns 1–17 follow the office ledger order exactly (two blank spacer columns kept).
// Columns 18–20 are extra details collected by the app.
// MATTER = "Payment For" from the receipt (one input, used on the receipt and in the ledger).
//
// The row is written in two passes, matching the order of the pages:
//   1. POST /api/records          – the printed receipt. The accounts columns are left
//                                   at their automatic values.
//   2. POST /api/records/accounts – the office accounts columns H–Q, filled in afterwards.
const HEADERS = [
  'DATE', 'CLIENT', 'SUBJECT', 'O/R DAT', 'O/R', 'MATTER', 'AMOUNT', 'DIS/INC', 'NETT',
  'SHARE(%)', 'SALARY', 'GOVIN', 'RAJ', '', 'RENT', '', 'TEL/IT',
  'LAWYER', 'MODE OF PAYMENT', 'TIMESTAMP (SGT)'
];
const HEADER_WIDTH = 22;                       // A..V – wider than HEADERS so old extra headers (incl. the retired TALLY column) are cleared
const HEADER_RANGE = 'A1:V1';
const HIGHLIGHT_COLS = [14, 16];               // RENT, TEL/IT (0-based) – yellow like the ledger
const MONEY_KEYS = ['disInc', 'share', 'salary', 'govin', 'raj', 'rent', 'telIt'];
const COL = { date: 0, client: 1, billNo: 4, matter: 5, amount: 6, lawyer: 17, mode: 18 };
const LAST_COL_LETTER = 'T';
const ACCOUNTS_RANGE = (rowNum) => `H${rowNum}:Q${rowNum}`;   // DIS/INC … TEL/IT

const LAWYERS = ['Dora Raj', 'Govinda Raju'];
const LEAD_PCT = 70;                           // lead lawyer's share of AMOUNT; the other gets the rest
const MODES = ['PayNow', 'Cash', 'Others'];

/* ---------- Google Sheets client ---------- */
// Shared with the invoice module, so one set of credentials serves both.
const getSheets = g.getSheets;
let headerChecked = false;

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
  const extras = current.slice(HEADERS.length).some(v => String(v ?? '').trim() !== '');
  if (norm(current) !== norm(HEADERS) || extras) {
    const padded = [...HEADERS, ...Array(HEADER_WIDTH - HEADERS.length).fill('')];
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!${HEADER_RANGE}`,
      valueInputOption: 'RAW',
      requestBody: { values: [padded] }
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

// Appends the row and returns its sheet row number (null in dry run).
async function appendRow(row) {
  if (DRY_RUN) {
    console.log('[DRY RUN] Row not sent to Google Sheets:', row);
    return null;
  }
  await ensureHeader();
  const res = await getSheets().spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `'${SHEET_TAB}'!A1`,
    valueInputOption: 'RAW',          // text stored as-is (no formula execution); numbers stay numeric
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [row] }
  });
  const m = /![A-Z]+(\d+)/.exec((res.data.updates && res.data.updates.updatedRange) || '');
  return m ? Number(m[1]) : null;
}

// Reads the data rows (A2:T) once, newest last.
async function readRows() {
  const got = await getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!A2:${LAST_COL_LETTER}`,
    valueRenderOption: 'UNFORMATTED_VALUE'
  });
  return got.data.values || [];
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

// Default GOVIN / RAJ split (only where left blank):
//   /G (Govinda Raju) -> GOVIN 70%, RAJ 30%   |   /D (Dora Raj) -> RAJ 70%, GOVIN 30%
// Worked in cents so the two parts always add up exactly to AMOUNT.
function splitAmount(amount, billNo, lawyer, m) {
  const govinLead = /\/G\s*$/i.test(billNo) || (!/\/D\s*$/i.test(billNo) && lawyer === 'Govinda Raju');
  const amountCents = Math.round(amount * 100);
  const leadCents = Math.round(amountCents * LEAD_PCT / 100);
  const cents = v => Math.round(v * 100);
  let govin = m.govin, raj = m.raj;
  if (govin === '' && raj === '') {              // both blank -> 70/30 split
    govin = (govinLead ? leadCents : amountCents - leadCents) / 100;
    raj = (govinLead ? amountCents - leadCents : leadCents) / 100;
  } else if (govin === '') {                     // only Raj typed -> Govin gets the rest
    govin = Math.max(0, amountCents - cents(raj)) / 100;
  } else if (raj === '') {                       // only Govin typed -> Raj gets the rest
    raj = Math.max(0, amountCents - cents(govin)) / 100;
  }                                              // both typed -> keep as entered
  return { govin, raj };
}

function sgtNow() {
  return new Date().toLocaleString('en-SG', {
    timeZone: 'Asia/Singapore', day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });
}

/* ---------- App ---------- */
const app = express();
// Word documents arrive from the Invoice page as base64 in JSON, and scanned
// PDF vouchers arrive the same way from the Petty-cash page.
app.use(express.json({ limit: '12mb' }));

// Everything below this line needs the office password. /healthz, the login
// page and the login call itself stay open.
auth.mount(app);

app.use(express.static(path.join(__dirname, 'public')));
app.use('/api/invoice', invoiceRoutes.router);
app.use('/api/pettycash', pettycashRoutes.router);

// Link for the "Open Google Sheet" buttons
app.get('/api/config', (_req, res) => {
  res.json({
    sheetUrl: SHEET_ID ? `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit` : null,
    authMode: g.authMode()
  });
});

// Real end-to-end check: reads the sheet and tests write access.
app.get('/healthz', async (_req, res) => {
  if (DRY_RUN) return res.json({ ok: true, sheets: 'dry-run', authMode: g.authMode(), reason: 'Google credentials or GOOGLE_SHEET_ID not set' });
  let serviceAccount = g.authMode();
  try { serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON).client_email; } catch (_) {}
  try {
    await ensureHeader();                       // needs read + write
    // The invoice module reports honestly: "connected" only after a real
    // read/write of the Invoices sheet, never because a check was skipped.
    const inv = invoiceRoutes.status();
    let invoices;
    if (!inv.sheetConfigured) invoices = 'NOT CONFIGURED - INVOICES_SHEET_ID is not set';
    else if (inv.dryRun) invoices = 'dry-run';
    else {
      try { await invoiceRoutes.ensureHeader(); invoices = 'connected'; }
      catch (e) { invoices = `error: ${e.message}`; }
    }
    const petty = pettycashRoutes.status();
    res.json({
      ok: true, sheets: 'connected', authMode: g.authMode(), serviceAccount,
      sheetId: SHEET_ID, tab: SHEET_TAB, invoices,
      pettycash: [
        petty.ai ? `AI ready (${petty.model})` : 'NO AI KEY - GEMINI_API_KEY is not set',
        petty.drive ? 'Drive ready' : 'NO DRIVE FOLDER - DRIVE_FOLDER_CLAIMS is not set',
        petty.sheet ? `sheet ready (tab "${petty.sheetTab}")` : 'NO CLAIMS SHEET'
      ].join(' | '),
      env: {
        INVOICES_SHEET_ID: inv.sheetConfigured,
        DRIVE_FOLDER_INVOICES: inv.folderInvoices,
        DRIVE_FOLDER_CASH_RECEIPTS: inv.folderCashReceipts,
        INVOICE_TEMPLATE_FILE_ID: inv.templateFile,
        GEMINI_API_KEY: petty.ai,
        DRIVE_FOLDER_CLAIMS: petty.folderClaims,
        CLAIMS_SHEET: petty.sheetConfigured,
        GOOGLE_OAUTH_CLIENT_ID: !!process.env.GOOGLE_OAUTH_CLIENT_ID,
        GOOGLE_OAUTH_CLIENT_SECRET: !!process.env.GOOGLE_OAUTH_CLIENT_SECRET,
        GOOGLE_OAUTH_REFRESH_TOKEN: !!process.env.GOOGLE_OAUTH_REFRESH_TOKEN,
        APP_PASSWORD: !!process.env.APP_PASSWORD,
        APP_SECRET: !!process.env.APP_SECRET
      }
    });
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

// Step 1 of the save: the receipt the client is given.
// The accounts fields are optional here – they are normally sent empty and filled in
// afterwards by /api/records/accounts, so a printed receipt is never waiting on office data.
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

  const billNo = clean(c.slNo, 40);
  const nett = Math.round((amount - (m.disInc || 0)) * 100) / 100;
  const split = splitAmount(amount, billNo, lawyer, m);
  const receiptDate = clean(r.date, 30);
  const client = clean(r.receivedFrom);
  const matter = clean(r.paymentFor, 500);
  const modeText = mode === 'Others' ? `Others – ${clean(r.modeOther, 100)}` : mode;

  const row = [
    receiptDate,                       // DATE
    client,                            // CLIENT
    clean(c.caseCategory),             // SUBJECT
    receiptDate,                       // O/R DAT
    billNo,                            // O/R (Bill No)
    matter,                            // MATTER (= Payment For on the receipt)
    amount,                            // AMOUNT
    m.disInc,                          // DIS/INC
    nett,                              // NETT
    m.share,                           // SHARE(%)
    m.salary,                          // SALARY
    split.govin,                       // GOVIN
    split.raj,                         // RAJ
    '',                                // (spacer)
    m.rent,                            // RENT
    '',                                // (spacer)
    m.telIt,                           // TEL/IT
    lawyer,                            // LAWYER
    modeText,                          // MODE OF PAYMENT
    sgtNow()                           // TIMESTAMP (SGT)
  ];

  try {
    await appendRow(row);
  } catch (err) {
    console.error('Sheets append failed:', err.message);
    return res.status(502).json({ ok: false, error: 'Could not save to Google Sheet. Please try again.' });
  }

  res.json({ ok: true });
});

// Step 2 of the save: the office accounts columns, typed after the receipt has been printed.
// Updates columns H–Q of the row that already holds this Bill No. NETT and the GOVIN / RAJ
// split are recalculated from the AMOUNT already in the sheet, so the receipt is never touched.
app.post('/api/records/accounts', async (req, res) => {
  const body = req.body || {};
  const billNo = clean(body.billNo, 40);
  const a = body.accounts || {};
  const m = Object.fromEntries(MONEY_KEYS.map(k => [k, money(a[k])]));
  const errors = [];

  if (!billNo) errors.push('Bill No');
  MONEY_KEYS.forEach(k => { if (Number.isNaN(m[k])) errors.push(k); });
  if (typeof m.share === 'number' && m.share > 100) errors.push('Share (%) must be 0–100');
  if (errors.length) return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });

  if (DRY_RUN) {
    console.log('[DRY RUN] Accounts not sent to Google Sheets:', billNo, m);
    return res.json({ ok: true, dryRun: true });
  }

  try {
    await ensureHeader();
    const rows = await readRows();
    let idx = -1;
    for (let i = rows.length - 1; i >= 0; i--) {          // newest row with this Bill No wins
      if (clean((rows[i] || [])[COL.billNo], 40) === billNo) { idx = i; break; }
    }
    if (idx < 0) {
      return res.status(404).json({ ok: false, error: `Bill No ${billNo} is not in the records sheet. Submit the receipt first.` });
    }

    const rowNum = idx + 2;
    const found = rows[idx] || [];
    const amount = Number(found[COL.amount]);
    if (!(amount > 0)) {
      return res.status(409).json({ ok: false, error: 'The saved row has no amount, so NETT and the split cannot be worked out.' });
    }
    const nett = Math.round((amount - (m.disInc || 0)) * 100) / 100;
    const split = splitAmount(amount, billNo, clean(found[COL.lawyer], 50), m);

    await getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!${ACCOUNTS_RANGE(rowNum)}`,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[
          m.disInc,        // H  DIS/INC
          nett,            // I  NETT
          m.share,         // J  SHARE(%)
          m.salary,        // K  SALARY
          split.govin,     // L  GOVIN
          split.raj,       // M  RAJ
          '',              // N  (spacer)
          m.rent,          // O  RENT
          '',              // P  (spacer)
          m.telIt          // Q  TEL/IT
        ]]
      }
    });
    res.json({ ok: true, row: rowNum, amount, nett, govin: split.govin, raj: split.raj });
  } catch (err) {
    console.error('Accounts update failed:', err.message);
    res.status(502).json({ ok: false, error: 'Could not update the Google Sheet. Please try again.' });
  }
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Raj Govin digital records running on port ${PORT}`
      + ` | auth: ${g.authMode()}`
      + `${DRY_RUN ? ' (DRY RUN – no Google credentials set)' : ''}`);
  });
}

module.exports = { app, splitAmount };
