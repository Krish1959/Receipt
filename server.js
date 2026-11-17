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
// Column 21 (U) is the Tally sync status, written by the app after each save.
// MATTER = "Payment For" from the receipt (one input, used on the receipt and in the ledger).
//
// The row is written in two passes, matching the order of the pages:
//   1. POST /api/records          – the printed receipt (and the Tally voucher). The accounts
//                                   columns are left at their automatic values.
//   2. POST /api/records/accounts – the office accounts columns H–Q, filled in afterwards.
const HEADERS = [
  'DATE', 'CLIENT', 'SUBJECT', 'O/R DAT', 'O/R', 'MATTER', 'AMOUNT', 'DIS/INC', 'NETT',
  'SHARE(%)', 'SALARY', 'GOVIN', 'RAJ', '', 'RENT', '', 'TEL/IT',
  'LAWYER', 'MODE OF PAYMENT', 'TIMESTAMP (SGT)', 'TALLY'
];
const HEADER_WIDTH = 22;                       // A..V – wider than HEADERS so old extra headers are cleared
const HEADER_RANGE = 'A1:V1';
const HIGHLIGHT_COLS = [14, 16];               // RENT, TEL/IT (0-based) – yellow like the ledger
const MONEY_KEYS = ['disInc', 'share', 'salary', 'govin', 'raj', 'rent', 'telIt'];
const COL = { date: 0, client: 1, billNo: 4, matter: 5, amount: 6, lawyer: 17, mode: 18, tally: 20 };
const TALLY_COL_LETTER = 'U';
const ACCOUNTS_RANGE = (rowNum) => `H${rowNum}:Q${rowNum}`;   // DIS/INC … TEL/IT

const LAWYERS = ['Dora Raj', 'Govinda Raju'];
const LEAD_PCT = 70;                           // lead lawyer's share of AMOUNT; the other gets the rest
const MODES = ['PayNow', 'Cash', 'Others'];

/* ---------- Tally settings ---------- */
// TALLY_URL is the address of TallyPrime's HTTP gateway (port 9000) as reachable from this server.
// On Render that must be a secure tunnel to the office PC – see README section 5.
const TALLY_URL = (process.env.TALLY_URL || '').trim().replace(/\/+$/, '');
const TALLY_ENABLED = !!TALLY_URL;
const TALLY_COMPANY = (process.env.TALLY_COMPANY || '').trim();      // blank = company currently open in Tally
const TALLY_VOUCHER_TYPE = (process.env.TALLY_VOUCHER_TYPE || 'Receipt').trim();
const TALLY_TIMEOUT_MS = Number(process.env.TALLY_TIMEOUT_MS) || 10000;
const TALLY_LEDGERS = {
  income: (process.env.TALLY_LEDGER_INCOME || 'Legal Fees').trim(),          // credited
  PayNow: (process.env.TALLY_LEDGER_PAYNOW || 'Bank Account').trim(),        // debited
  Cash: (process.env.TALLY_LEDGER_CASH || 'Cash').trim(),
  Others: (process.env.TALLY_LEDGER_OTHERS || process.env.TALLY_LEDGER_PAYNOW || 'Bank Account').trim()
};
let TALLY_EXTRA_HEADERS = {};                  // e.g. tunnel access token: {"CF-Access-Client-Id":"…","CF-Access-Client-Secret":"…"}
try { TALLY_EXTRA_HEADERS = JSON.parse(process.env.TALLY_EXTRA_HEADERS || '{}'); }
catch (_) { console.error('TALLY_EXTRA_HEADERS is not valid JSON – ignored'); }

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

async function setTallyCell(rowNum, text) {
  if (DRY_RUN || !rowNum) return;
  try {
    await getSheets().spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!${TALLY_COL_LETTER}${rowNum}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[text]] }
    });
  } catch (err) {
    console.error(`Could not write Tally status to row ${rowNum}:`, err.message);
  }
}

// Reads the data rows (A2:U) once, newest last.
async function readRows() {
  const got = await getSheets().spreadsheets.values.get({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!A2:${TALLY_COL_LETTER}`,
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

/* ---------- Tally ---------- */
const xmlEsc = s => String(s ?? '').replace(/[&<>"']/g, ch =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[ch]));
const xmlUnesc = s => String(s).replace(/&(amp|lt|gt|quot|apos);/g, (_, e) =>
  ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[e]));
const xmlTag = (xml, name) => {
  const m = new RegExp(`<${name}>([^<]*)</${name}>`, 'i').exec(xml);
  return m ? xmlUnesc(m[1].trim()) : '';
};

const MONTHS = { jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06', jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12' };
// "16 Nov 2026" -> "20261116"; anything unreadable -> today's date in Singapore
function tallyDate(text) {
  const m = /^(\d{1,2})\s+([A-Za-z]{3})[A-Za-z]*\.?,?\s+(\d{4})$/.exec(String(text || '').trim());
  if (m && MONTHS[m[2].toLowerCase()]) return `${m[3]}${MONTHS[m[2].toLowerCase()]}${m[1].padStart(2, '0')}`;
  return new Date().toLocaleDateString('en-CA', { timeZone: 'Asia/Singapore' }).replace(/-/g, '');
}

async function tallyRequest(xml) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), TALLY_TIMEOUT_MS);
  try {
    const res = await fetch(TALLY_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/xml; charset=utf-8', ...TALLY_EXTRA_HEADERS },
      body: xml,
      signal: ctrl.signal
    });
    const text = await res.text();
    if (!res.ok) throw new Error(`Tally gateway answered HTTP ${res.status}`);
    return text;
  } catch (err) {
    if (err.name === 'AbortError') throw new Error('Tally did not answer in time (is the office PC on and TallyPrime open?)');
    if (/fetch failed/i.test(err.message)) throw new Error('Tally cannot be reached (is the office PC on and TallyPrime open?)');
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

const companyVar = () => (TALLY_COMPANY ? `<SVCURRENTCOMPANY>${xmlEsc(TALLY_COMPANY)}</SVCURRENTCOMPANY>` : '');

// One Receipt voucher: Dr bank/cash ledger (by mode of payment), Cr income ledger, for AMOUNT.
// REMOTEID is built from the Bill No so the same receipt is never entered twice.
function buildVoucherXml(rec) {
  const modeKey = /^Others/i.test(rec.mode) ? 'Others' : (TALLY_LEDGERS[rec.mode] ? rec.mode : 'Others');
  const amt = Number(rec.amount).toFixed(2);
  const narration = [`Bill ${rec.billNo}`, rec.client, rec.matter, rec.mode].filter(Boolean).join(' | ');
  return `<ENVELOPE>
 <HEADER><TALLYREQUEST>Import Data</TALLYREQUEST></HEADER>
 <BODY><IMPORTDATA>
  <REQUESTDESC><REPORTNAME>Vouchers</REPORTNAME><STATICVARIABLES>${companyVar()}</STATICVARIABLES></REQUESTDESC>
  <REQUESTDATA><TALLYMESSAGE xmlns:UDF="TallyUDF">
   <VOUCHER REMOTEID="${xmlEsc('rajgovin-' + rec.billNo)}" VCHTYPE="${xmlEsc(TALLY_VOUCHER_TYPE)}" ACTION="Create">
    <DATE>${tallyDate(rec.date)}</DATE>
    <VOUCHERTYPENAME>${xmlEsc(TALLY_VOUCHER_TYPE)}</VOUCHERTYPENAME>
    <REFERENCE>${xmlEsc(rec.billNo)}</REFERENCE>
    <NARRATION>${xmlEsc(narration)}</NARRATION>
    <ALLLEDGERENTRIES.LIST>
     <LEDGERNAME>${xmlEsc(TALLY_LEDGERS.income)}</LEDGERNAME>
     <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>
     <AMOUNT>${amt}</AMOUNT>
    </ALLLEDGERENTRIES.LIST>
    <ALLLEDGERENTRIES.LIST>
     <LEDGERNAME>${xmlEsc(TALLY_LEDGERS[modeKey])}</LEDGERNAME>
     <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>
     <AMOUNT>-${amt}</AMOUNT>
    </ALLLEDGERENTRIES.LIST>
   </VOUCHER>
  </TALLYMESSAGE></REQUESTDATA>
 </IMPORTDATA></BODY>
</ENVELOPE>`;
}

// Throws with a readable reason if Tally did not accept the voucher.
async function postReceiptToTally(rec) {
  const reply = await tallyRequest(buildVoucherXml(rec));
  const created = Number(xmlTag(reply, 'CREATED')) || 0;
  const altered = Number(xmlTag(reply, 'ALTERED')) || 0;
  const errors = Number(xmlTag(reply, 'ERRORS')) || 0;
  if (errors > 0 || created + altered < 1) {
    const why = xmlTag(reply, 'LINEERROR') || reply.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200);
    throw new Error(why || 'Tally rejected the voucher');
  }
  return { voucherId: xmlTag(reply, 'LASTVCHID'), altered: created === 0 };
}

// Posts one receipt and records the outcome in the sheet's TALLY column. Never throws.
async function syncTally(rec, rowNum) {
  try {
    const out = await postReceiptToTally(rec);
    await setTallyCell(rowNum, `POSTED${out.voucherId ? ' #' + out.voucherId : ''} ${sgtNow()}`);
    return { status: 'posted', voucherId: out.voucherId || null };
  } catch (err) {
    console.error(`Tally post failed for ${rec.billNo}:`, err.message);
    await setTallyCell(rowNum, `FAILED: ${clean(err.message, 200)}`);
    return { status: 'failed', error: clean(err.message, 200) };
  }
}

async function tallyLedgerNames() {
  const reply = await tallyRequest(`<ENVELOPE>
 <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST><TYPE>Collection</TYPE><ID>RG Ledgers</ID></HEADER>
 <BODY><DESC>
  <STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>${companyVar()}</STATICVARIABLES>
  <TDL><TDLMESSAGE><COLLECTION NAME="RG Ledgers" ISMODIFY="No"><TYPE>Ledger</TYPE><NATIVEMETHOD>Name</NATIVEMETHOD></COLLECTION></TDLMESSAGE></TDL>
 </DESC></BODY>
</ENVELOPE>`);
  const names = [];
  const re = /<LEDGER\b[^>]*\bNAME="([^"]*)"/gi;
  let m;
  while ((m = re.exec(reply))) names.push(xmlUnesc(m[1]));
  return names;
}

async function tallyHealth() {
  if (!TALLY_ENABLED) return { tally: 'disabled', tallyNote: 'TALLY_URL not set' };
  try {
    const have = new Set((await tallyLedgerNames()).map(n => n.toLowerCase()));
    const needed = [...new Set(Object.values(TALLY_LEDGERS))];
    const missing = needed.filter(n => !have.has(n.toLowerCase()));
    return {
      tally: missing.length ? 'ledgers-missing' : 'connected',
      company: TALLY_COMPANY || '(company open in Tally)',
      ledgers: TALLY_LEDGERS,
      ...(missing.length ? { missingLedgers: missing, fix: 'Create these ledgers in Tally or change the TALLY_LEDGER_* variables' } : {})
    };
  } catch (err) {
    return { tally: 'error', tallyError: err.message };
  }
}

/* ---------- App ---------- */
const app = express();
app.use(express.json({ limit: '20kb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Link for the EXCEL button on the receipt page
app.get('/api/config', (_req, res) => {
  res.json({
    sheetUrl: SHEET_ID ? `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit` : null,
    tally: TALLY_ENABLED
  });
});

// Real end-to-end check: reads the sheet, tests write access, and checks Tally and its ledgers.
app.get('/healthz', async (_req, res) => {
  const tally = await tallyHealth();
  if (DRY_RUN) return res.json({ ok: true, sheets: 'dry-run', reason: 'GOOGLE_CREDENTIALS_JSON or GOOGLE_SHEET_ID not set', ...tally });
  let serviceAccount = 'unreadable';
  try { serviceAccount = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON).client_email; } catch (_) {}
  try {
    await ensureHeader();                       // needs read + write
    res.json({ ok: true, sheets: 'connected', serviceAccount, sheetId: SHEET_ID, tab: SHEET_TAB, ...tally });
  } catch (err) {
    res.status(500).json({
      ok: false, sheets: 'error', serviceAccount, sheetId: SHEET_ID,
      error: err.message,
      fix: /permission/i.test(err.message)
        ? `Share the sheet with ${serviceAccount} as Editor`
        : /not found/i.test(err.message) ? 'Check GOOGLE_SHEET_ID' : 'Check credentials and that the Google Sheets API is enabled',
      ...tally
    });
  }
});

// Step 1 of the save: the receipt the client is given, plus the Tally voucher.
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
    sgtNow(),                          // TIMESTAMP (SGT)
    TALLY_ENABLED ? 'PENDING' : ''     // TALLY (updated below once Tally answers)
  ];

  let rowNum;
  try {
    rowNum = await appendRow(row);
  } catch (err) {
    console.error('Sheets append failed:', err.message);
    return res.status(502).json({ ok: false, error: 'Could not save to Google Sheet. Please try again.' });
  }

  // The sheet row is saved. Tally comes second; a Tally problem never loses the receipt.
  const tally = TALLY_ENABLED
    ? await syncTally({ billNo, date: receiptDate, client, matter, amount, mode: modeText }, rowNum)
    : { status: 'disabled' };
  res.json({ ok: true, tally });
});

// Step 2 of the save: the office accounts columns, typed after the receipt has been printed.
// Updates columns H–Q of the row that already holds this Bill No. NETT and the GOVIN / RAJ
// split are recalculated from the AMOUNT already in the sheet, so the receipt is never touched
// and nothing is re-sent to Tally (the voucher is for AMOUNT, which does not change here).
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

// Re-send receipts whose TALLY column is PENDING or FAILED (e.g. the office PC was off).
// Body { billNo } limits it to one receipt; without it, up to 50 waiting rows are sent.
app.post('/api/tally/retry', async (req, res) => {
  if (!TALLY_ENABLED) return res.status(400).json({ ok: false, error: 'Tally is not configured (TALLY_URL not set).' });
  if (DRY_RUN) return res.status(400).json({ ok: false, error: 'Retry needs the Google Sheet (not available in dry run).' });
  const only = clean((req.body || {}).billNo, 40);
  try {
    await ensureHeader();
    const rows = await readRows();
    const results = [];
    for (let i = 0; i < rows.length && results.length < 50; i++) {
      const v = rows[i];
      const billNo = clean(v[COL.billNo], 40);
      const status = String(v[COL.tally] ?? '');
      const amount = Number(v[COL.amount]);
      if (!billNo || !(amount > 0)) continue;
      if (only ? billNo !== only : !/^(PENDING|FAILED)/.test(status)) continue;
      if (/^POSTED/.test(status)) { results.push({ billNo, status: 'posted', note: 'already in Tally' }); continue; }
      const out = await syncTally({
        billNo, date: clean(v[COL.date], 30), client: clean(v[COL.client]),
        matter: clean(v[COL.matter], 500), amount, mode: clean(v[COL.mode], 130)
      }, i + 2);
      results.push({ billNo, ...out });
    }
    res.json({
      ok: true,
      posted: results.filter(x => x.status === 'posted').length,
      failed: results.filter(x => x.status === 'failed').length,
      results
    });
  } catch (err) {
    console.error('Tally retry failed:', err.message);
    res.status(502).json({ ok: false, error: 'Could not read the Google Sheet for retry.' });
  }
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`Raj Govin receipts running on port ${PORT}`
      + `${DRY_RUN ? ' (DRY RUN – no Google credentials set)' : ''}`
      + ` | Tally: ${TALLY_ENABLED ? TALLY_URL : 'off (TALLY_URL not set)'}`);
  });
}

module.exports = { app, buildVoucherXml, tallyDate, splitAmount };
