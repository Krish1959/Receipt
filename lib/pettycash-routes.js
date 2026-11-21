'use strict';

// Module 1 – the Petty-cash page.
//
// The clerk's round trip:
//   1. POST /api/pettycash/extract  – one scanned voucher (PDF), read by AI
//                                     into a single line: description, amount, GST.
//                                     The PDF itself is kept against the job.
//   2. POST /api/pettycash/build    – the checked lines come back as a Word
//                                     "Cash Purchase Claim", ready to print
//   3. POST /api/pettycash/save     – one Drive folder named after the claim,
//                                     holding the form as PDF and every voucher
//                                     scan, and one row in the Claims sheet
//                                     carrying the link to that folder
//
// The AI only READS. Every line it returns is shown on screen and can be
// corrected before the claim form is made, so a misread amount is never
// filed silently.
//
// WHICH AI MODEL: the names live in models.json at the top of the repo, not
// in this file, so the office can change them without touching any code.
// Google retires models every few months; the app walks that list from the
// top and uses the first name that answers, so a retirement is a slower
// read, never a dead page.

const fs = require('fs');
const path = require('path');
const express = require('express');

const claim = require('./claimdocx');
const g = require('./google');
const { sheetDate } = require('./sheetdate');

const router = express.Router();

/* ---------- settings ---------- */
const GEMINI_KEY = (process.env.GEMINI_API_KEY || '').trim();
const GEMINI_URL = m => `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(m)}:generateContent`;

const MODELS_FILE = path.join(__dirname, '..', 'models.json');
const BUILT_IN_MODELS = ['gemini-3.8-flash', 'gemini-3.7-flash'];   // only if models.json is missing or unreadable

// Where the claims are filed. One folder per claim is created inside this one.
const FOLDER_CLAIMS = (process.env.DRIVE_FOLDER_CLAIMS || '').trim();

// The Claims tab of the Accounts spreadsheet. server.js hands ACCOUNTS_SHEET_ID
// to CLAIMS_SHEET_ID before this file loads; the rest of the chain keeps an
// older deployment working.
const SHEET_ID = (process.env.CLAIMS_SHEET_ID || process.env.INVOICES_SHEET_ID || process.env.GOOGLE_SHEET_ID || '')
  .trim().replace(/^["']|["']$/g, '');
const SHEET_TAB = (process.env.CLAIMS_SHEET_TAB || 'Claims').trim();

const MAX_PDF_BYTES = 6 * 1024 * 1024;         // one voucher
const AI_TIMEOUT_MS = 90 * 1000;
const JOB_TTL_MS = 2 * 60 * 60 * 1000;         // a claim in progress is forgotten after 2 hours
const JOB_BYTES_CAP = 120 * 1024 * 1024;       // scans held in memory across all claims

const DRIVE_READY = g.authMode() !== 'none' && !!FOLDER_CLAIMS;
const SHEET_READY = g.authMode() !== 'none' && !!SHEET_ID;

// TALLY STATUS is deliberately the LAST column and is always left empty here.
// Module 1 never writes to it. Module 2 (the Tally push on the office PC)
// fills it in when the claim has been posted, so a blank cell means "not yet
// in Tally" and nothing in Module 1 can overwrite that answer.
const HEADERS = [
  'CLAIM NO', 'DATE', 'PAY TO', 'LAWYER', 'VOUCHER TYPE', 'DESCRIPTION',
  'VOUCHERS', 'TOTAL (S$)', 'GST (S$)', 'FOLDER', 'TIMESTAMP (SGT)', 'TALLY STATUS'
];
// Module 1 manages EXACTLY its own columns, A1:L1, and never writes past L.
//
// This matters. Module 2 records "posted to Tally" marks in the sheet, and it
// deliberately puts them outside the range Module 1 rewrites on every cold
// start — a mark inside that range was wiped once before, the rows read as
// unposted, and two receipts went into Tally twice. Claims keeps the pattern:
// column L is Module 1's own TALLY STATUS, written empty and never touched
// again, and M onwards is free for Module 2's timestamp.
const HEADER_WIDTH = 12;                       // A..L – Module 1's own columns, no more
const HEADER_RANGE = 'A1:L1';
const COL_FOLDER_LINK = 'J';
const FOLDER_MIME = 'application/vnd.google-apps.folder';

/* ---------- the model list ---------- */

// Read once at start-up. Render restarts on every deploy, so editing
// models.json and pushing it is all that is needed to change the list.
function loadModels() {
  let list = [];
  let from = 'models.json';
  try {
    const raw = JSON.parse(fs.readFileSync(MODELS_FILE, 'utf8'));
    list = (Array.isArray(raw) ? raw : raw.models || [])
      .map(m => (typeof m === 'string' ? m : (m && m.id)))
      .map(s => String(s || '').trim())
      .filter(Boolean);
    if (!list.length) throw new Error('the "models" list is empty');
  } catch (err) {
    list = BUILT_IN_MODELS.slice();
    from = `built-in list (models.json not used: ${err.message})`;
  }

  // GEMINI_MODEL, if set on Render, is tried before anything in the file.
  const forced = (process.env.GEMINI_MODEL || '').trim();
  if (forced) list = [forced, ...list.filter(m => m !== forced)];

  console.log(`Petty-cash AI models from ${from}: ${list.join(', ')}`);
  return list;
}

const MODELS = loadModels();
let active = null;        // the model that last answered; tried first from then on

// Google's way of saying "that model is gone". Anything else — a bad key,
// no quota, an unreadable page — is a real failure and must NOT silently
// burn through every model in the list.
const RETIRED = /no longer available|not found|does not exist|unsupported model|is not supported for|deprecat/i;
const isRetired = (status, why) => status === 404 || (status === 400 && RETIRED.test(why));

/* ---------- helpers ---------- */
const clean = (v, max = 300) => String(v ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
const safeName = s => clean(s, 80).replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();

function sgtNow() {
  return new Date().toLocaleString('en-SG', {
    timeZone: 'Asia/Singapore', day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });
}

// '' -> ''; a valid amount -> Number with 2 dp; anything else -> NaN
function money(v) {
  const s = clean(v, 20).replace(/[,\s$]/g, '');
  if (!s) return '';
  if (!/^\d+(\.\d{1,2})?$/.test(s)) return NaN;
  const n = Math.round(parseFloat(s) * 100) / 100;
  return n <= 9999999.99 ? n : NaN;
}

function decodePdf(base64) {
  const buf = Buffer.from(String(base64 || '').replace(/^data:[^;]+;base64,/, ''), 'base64');
  if (!buf.length) throw new Error('That file came through empty. Please try the drag and drop again.');
  if (buf.length > MAX_PDF_BYTES) throw new Error('That PDF is larger than 6 MB. Please scan it at a lower resolution.');
  if (buf.slice(0, 5).toString('latin1') !== '%PDF-') throw new Error('That is not a PDF file.');
  return buf;
}

/* ---------- claims in flight ---------- */
// One claim being worked on: the voucher scans, and the Word form once it is
// made. Held in memory only — a restart loses claims that have not been
// saved, and the clerk drops the vouchers again.
const jobs = new Map();
const newJobId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

function sweep() {
  const now = Date.now();
  for (const [id, j] of jobs) if (now - j.at > JOB_TTL_MS) jobs.delete(id);

  // Scans are the only bulky thing here. If too much is being held, the
  // oldest claims go first — they have long since been printed or abandoned.
  let total = 0;
  for (const j of jobs.values()) total += j.bytes;
  if (total <= JOB_BYTES_CAP) return;
  const oldest = [...jobs.values()].sort((a, b) => a.at - b.at);
  for (const j of oldest) {
    if (total <= JOB_BYTES_CAP) break;
    total -= j.bytes;
    jobs.delete(j.id);
    console.warn(`Petty-cash: dropped claim ${j.id} to stay inside the memory cap`);
  }
}

function touchJob(id) {
  sweep();
  const existing = jobs.get(clean(id, 40));
  if (existing) { existing.at = Date.now(); return existing; }
  const job = { id: newJobId(), at: Date.now(), files: new Map(), bytes: 0, docx: null, claim: null, saved: null };
  jobs.set(job.id, job);
  return job;
}

function getJob(id) {
  const job = jobs.get(clean(id, 40));
  if (!job) throw new Error('This claim is no longer open — it may have timed out. Please start it again.');
  job.at = Date.now();
  return job;
}

/* ---------- the AI read ---------- */

const PROMPT = [
  'You are reading ONE scanned expense voucher or receipt for a Singapore law office petty-cash claim.',
  'The page may be handwritten, printed, or a photograph of a till receipt.',
  '',
  'Return JSON only:',
  '- description: ONE short line, at most 60 characters, saying what the money was spent on, and where or on whom if that is clear on the page. Examples: "Stationery - Popular Bookstore", "Taxi to Supreme Court", "Refreshments for client meeting". No amounts, no dates, no line breaks.',
  '- amount: the GRAND TOTAL actually paid, as a number, for example 42.70. If several totals appear, take the final amount payable with GST included.',
  '- gst: the GST amount printed on the voucher, as a number. If no GST is shown anywhere on the page, return 0. Never calculate or estimate a GST figure that is not printed.',
  '- date: the date on the voucher as dd-mm-yyyy, or "" if it cannot be read.',
  '- confidence: "high" if the page is printed and clear, "medium" if parts are unclear, "low" if handwriting or image quality leaves you unsure of the amount.',
  '',
  'Read only what is on the page. Never invent a vendor, an amount or a date.'
].join('\n');

const SCHEMA = {
  type: 'OBJECT',
  properties: {
    description: { type: 'STRING' },
    amount: { type: 'NUMBER' },
    gst: { type: 'NUMBER' },
    date: { type: 'STRING' },
    confidence: { type: 'STRING', enum: ['high', 'medium', 'low'] }
  },
  required: ['description', 'amount', 'gst', 'confidence']
};

// One attempt against one model. Rejects with { retired: true } when this
// model is gone and the next should be tried.
async function askModel(model, pdfBase64) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), AI_TIMEOUT_MS);
  let res;
  try {
    res = await fetch(GEMINI_URL(model), {
      method: 'POST',
      signal: controller.signal,
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': GEMINI_KEY },
      body: JSON.stringify({
        contents: [{
          role: 'user',
          parts: [
            { inline_data: { mime_type: 'application/pdf', data: pdfBase64 } },
            { text: PROMPT }
          ]
        }],
        generationConfig: {
          temperature: 0,
          responseMimeType: 'application/json',
          responseSchema: SCHEMA
        }
      })
    });
  } catch (err) {
    throw new Error(err.name === 'AbortError'
      ? 'The AI took too long to read that voucher. Try it again, or type the line in by hand.'
      : `The AI service could not be reached: ${clean(err.message, 120)}`);
  } finally {
    clearTimeout(timer);
  }

  const raw = await res.text();
  if (!res.ok) {
    let why = clean(raw, 300);
    try { why = clean(JSON.parse(raw).error.message, 300); } catch (_) {}
    if (isRetired(res.status, why)) {
      const gone = new Error(why);
      gone.retired = true;
      return Promise.reject(gone);
    }
    if (res.status === 400 && /API key/i.test(why)) why = 'the AI key was refused — check GEMINI_API_KEY on Render';
    if (res.status === 429) why = 'the AI quota for this key is used up for now';
    if (res.status === 403) why = 'this key is not allowed to use the Gemini API — check the key restrictions';
    throw new Error(`The AI could not read that voucher (${clean(why, 200)}).`);
  }

  let parsed;
  try {
    const body = JSON.parse(raw);
    const text = (((body.candidates || [])[0] || {}).content || {}).parts
      ?.map(p => p.text).filter(Boolean).join('') || '';
    parsed = JSON.parse(text);
  } catch (_) {
    throw new Error('The AI reply could not be understood. Please type the line in by hand.');
  }

  const amount = Number(parsed.amount);
  const gst = Number(parsed.gst);
  return {
    description: clean(parsed.description, 120).replace(/\s+/g, ' '),
    amount: Number.isFinite(amount) && amount > 0 ? Math.round(amount * 100) / 100 : '',
    gst: Number.isFinite(gst) && gst > 0 ? Math.round(gst * 100) / 100 : 0,
    date: /^\d{2}-\d{2}-\d{4}$/.test(String(parsed.date || '')) ? parsed.date : '',
    confidence: ['high', 'medium', 'low'].includes(parsed.confidence) ? parsed.confidence : 'low'
  };
}

// Walks the list until one model answers. A retired model is stepped over;
// any other failure stops here, because trying four models against a bad key
// only wastes the clerk's time.
async function readVoucher(pdfBase64) {
  if (!GEMINI_KEY) {
    throw new Error('No AI key is set on the server. Add GEMINI_API_KEY, or type the line in by hand.');
  }
  const order = [active, ...MODELS].filter((m, i, a) => m && a.indexOf(m) === i);
  const skipped = [];

  for (const model of order) {
    try {
      const line = await askModel(model, pdfBase64);
      if (active !== model) {
        console.log(`Petty-cash AI now using ${model}${skipped.length ? ` (${skipped.join(', ')} retired)` : ''}`);
        active = model;
      }
      return { line, model };
    } catch (err) {
      if (!err.retired) throw err;
      console.error(`Petty-cash AI model ${model} is retired: ${err.message}`);
      skipped.push(model);
      if (active === model) active = null;
    }
  }

  throw new Error(
    `None of the AI models in models.json are available any more (${skipped.join(', ')}). `
    + 'Edit models.json with a current model name from ai.google.dev/gemini-api/docs/models, '
    + 'or type the lines in by hand.'
  );
}

/* ---------- Drive ---------- */

// One folder per claim, named after the claim number.
//
// A folder this same claim already made is re-used, so a second attempt
// after a half-finished save does not leave two folders holding one claim.
// A folder left by a DIFFERENT claim is never written into: that would mean
// two claims sharing one number, and the vouchers of one would sit in the
// other's folder. The clerk is asked for a new number instead.
async function claimFolder(name, job) {
  const drive = g.getDrive();
  const q = [
    `name = '${name.replace(/'/g, "\\'")}'`,
    `mimeType = '${FOLDER_MIME}'`,
    `'${FOLDER_CLAIMS}' in parents`,
    'trashed = false'
  ].join(' and ');

  const found = await drive.files.list({
    q, fields: 'files(id,webViewLink)', pageSize: 1,
    supportsAllDrives: true, includeItemsFromAllDrives: true
  });
  if (found.data.files && found.data.files.length) {
    const f = found.data.files[0];
    if (job.folderId !== f.id) {
      const clash = new Error(
        `A claim folder named ${name} is already in Google Drive, from an earlier claim. `
        + 'Change the Claim Form No, make the claim form again, then save.'
      );
      clash.collision = true;
      throw clash;
    }
    return { id: f.id, url: f.webViewLink || `https://drive.google.com/drive/folders/${f.id}`, reused: true };
  }

  const made = await drive.files.create({
    requestBody: { name, mimeType: FOLDER_MIME, parents: [FOLDER_CLAIMS] },
    fields: 'id,webViewLink',
    supportsAllDrives: true
  });
  job.folderId = made.data.id;
  return {
    id: made.data.id,
    url: made.data.webViewLink || `https://drive.google.com/drive/folders/${made.data.id}`,
    reused: false
  };
}

/* ---------- the Claims sheet ---------- */
let headerChecked = false;

async function ensureHeader() {
  if (headerChecked || !SHEET_READY) return;
  const api = g.getSheets();
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
  const extras = current.slice(HEADERS.length).some(v => String(v ?? '').trim() !== '');
  if (extras || HEADERS.join('|') !== HEADERS.map((_, i) => String(current[i] ?? '')).join('|')) {
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!${HEADER_RANGE}`,
      valueInputOption: 'RAW',
      requestBody: { values: [[...HEADERS, ...Array(HEADER_WIDTH - HEADERS.length).fill('')]] }
    });
    await api.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          {
            repeatCell: {
              range: { sheetId: tabId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: HEADERS.length },
              cell: { userEnteredFormat: { textFormat: { bold: true }, backgroundColor: { red: .937, green: .957, blue: .984 } } },
              fields: 'userEnteredFormat(textFormat.bold,backgroundColor)'
            }
          },
          {
            updateSheetProperties: {
              properties: { sheetId: tabId, gridProperties: { frozenRowCount: 1, frozenColumnCount: 1 } },
              fields: 'gridProperties(frozenRowCount,frozenColumnCount)'
            }
          },
          ...[[0, 150], [1, 100], [2, 160], [3, 130], [4, 150], [5, 320], [6, 90], [7, 100], [8, 90], [9, 110], [10, 160], [11, 130]].map(([i, px]) => ({
            updateDimensionProperties: {
              range: { sheetId: tabId, dimension: 'COLUMNS', startIndex: i, endIndex: i + 1 },
              properties: { pixelSize: px }, fields: 'pixelSize'
            }
          }))
        ]
      }
    });
  }
  headerChecked = true;
}

// Plain values first, then the folder link as a formula. Nothing the clerk
// typed is ever evaluated as a formula.
async function appendClaim(rec) {
  await ensureHeader();
  const api = g.getSheets();
  const res = await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!A1`,
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[
        rec.claimNo,       // A
        sheetDate(rec.date), // B  "02 Nov 2026" — never a bare 02-11-2026
        rec.payTo,         // C
        rec.lawyer,        // D
        rec.voucherType,   // E
        rec.description,   // F
        rec.vouchers,      // G
        rec.total,         // H
        rec.gst,           // I
        '',                // J  folder link, written below
        sgtNow(),          // K
        ''                 // L  TALLY STATUS – left for Module 2
      ]]
    }
  });
  const m = /![A-Z]+(\d+)/.exec((res.data.updates && res.data.updates.updatedRange) || '');
  const rowNum = m ? Number(m[1]) : null;
  if (rowNum && rec.folderUrl) {
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `'${SHEET_TAB}'!${COL_FOLDER_LINK}${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[`=HYPERLINK("${rec.folderUrl}","Open folder")`]] }
    });
  }
  return rowNum;
}

/* ================= routes ================= */

router.get('/config', (_req, res) => {
  res.json({
    ok: true,
    ai: !!GEMINI_KEY,
    model: GEMINI_KEY ? (active || MODELS[0]) : null,
    models: MODELS,
    drive: DRIVE_READY,
    sheet: SHEET_READY,
    maxFiles: claim.MAX_ROWS,
    voucherTypes: claim.VOUCHER_TYPES,
    claimsFolder: FOLDER_CLAIMS ? `https://drive.google.com/drive/folders/${FOLDER_CLAIMS}` : null
  });
});

// 1. One voucher in, one claim line out. The scan is kept against the job
//    so it can be filed later without the browser sending it twice.
router.post('/extract', async (req, res) => {
  const body = req.body || {};
  const name = clean(body.name, 120) || 'voucher.pdf';
  try {
    const buf = decodePdf(body.base64);
    const job = touchJob(body.jobId);

    if (job.files.size >= claim.MAX_ROWS) {
      throw new Error(`${claim.MAX_ROWS} vouchers is the limit for one claim form.`);
    }

    const { line, model } = await readVoucher(buf.toString('base64'));
    if (!line.description && line.amount === '') {
      throw new Error('Nothing readable was found on that page. Please check the scan, or type the line in by hand.');
    }

    const fileId = newJobId();
    job.files.set(fileId, { name, buffer: buf });
    job.bytes += buf.length;

    res.json({ ok: true, jobId: job.id, fileId, name, line, model });
  } catch (err) {
    console.error(`Petty-cash extract failed for ${name}:`, err.message);
    res.status(400).json({ ok: false, name, error: clean(err.message, 300) });
  }
});

// 2. The checked lines become the Word claim form.
router.post('/build', async (req, res) => {
  const body = req.body || {};
  const claimNo = clean(body.claimNo, 40);
  const date = clean(body.date, 10);
  const payTo = clean(body.payTo, 120);
  const lawyer = clean(body.lawyer, 50);
  const voucherType = clean(body.voucherType, 40);
  const rows = Array.isArray(body.lines) ? body.lines : [];

  const errors = [];
  if (!claimNo) errors.push('Claim Form No');
  if (!/^\d{2}-\d{2}-\d{4}$/.test(date)) errors.push('Date (dd-mm-yyyy)');
  if (!payTo) errors.push('Pay To');
  if (!rows.length) errors.push('at least one claim line');
  if (rows.length > claim.MAX_ROWS) errors.push(`no more than ${claim.MAX_ROWS} claim lines`);

  const lines = rows.slice(0, claim.MAX_ROWS).map((r, i) => {
    const description = clean(r.description, 120);
    const amount = money(r.amount);
    const gst = money(r.gst);
    if (!description) errors.push(`line ${i + 1} description`);
    if (amount === '' || Number.isNaN(amount) || amount <= 0) errors.push(`line ${i + 1} amount`);
    if (Number.isNaN(gst)) errors.push(`line ${i + 1} GST`);
    else if (typeof gst === 'number' && typeof amount === 'number' && gst > amount) {
      errors.push(`line ${i + 1} GST is larger than the amount`);
    }
    return { description, amount, gst: gst === '' ? 0 : gst, fileId: clean(r.fileId, 40) || null };
  });

  if (errors.length) {
    return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });
  }

  try {
    const made = await claim.buildClaim({ claimNo, date, payTo, voucherType, lines });
    const job = touchJob(body.jobId);
    const type = claim.VOUCHER_TYPES.includes(voucherType) ? voucherType : claim.VOUCHER_TYPES[0];

    job.docx = made.buffer;
    job.saved = null;                         // a rebuilt form has not been filed yet
    job.claim = {
      claimNo, date, payTo, lawyer, voucherType: type,
      total: made.total, gst: made.totalGst,
      lines: lines.map((l, i) => ({ ...l, slNo: i + 1 }))
    };

    res.json({
      ok: true,
      jobId: job.id,
      claimNo, date, payTo, voucherType: type,
      fileName: `Cash Purchase Claim ${safeName(claimNo)}.docx`,
      total: made.total,
      totalGst: made.totalGst,
      lines: made.lines,
      docx: made.buffer.toString('base64')
    });
  } catch (err) {
    console.error('Claim form build failed:', err.message);
    res.status(500).json({ ok: false, error: `The claim form could not be made: ${clean(err.message, 200)}` });
  }
});

// 3. File the claim: one Drive folder named after the claim number holding
//    the form as a PDF and every voucher scan, and one row in the Claims
//    sheet pointing at that folder.
router.post('/save', async (req, res) => {
  const body = req.body || {};
  let job;
  try { job = getJob(body.jobId); }
  catch (err) { return res.status(404).json({ ok: false, error: clean(err.message, 200) }); }

  if (!job.docx || !job.claim) {
    return res.status(409).json({ ok: false, error: 'Make the claim form first, then save it.' });
  }
  // Pressing "Next Claim" after "Save & Exit" (or the other way about) must
  // not file the same claim twice.
  if (job.saved) return res.json({ ok: true, ...job.saved, already: true });

  const c = job.claim;
  const folderName = safeName(c.claimNo);

  if (!DRIVE_READY && !SHEET_READY) {
    return res.json({
      ok: true, dryRun: true, folderUrl: null, row: null,
      warning: g.authMode() === 'none'
        ? 'Nothing was filed: this server has no Google credentials. The claim form downloaded to your computer is the only copy.'
        : 'Nothing was filed: DRIVE_FOLDER_CLAIMS and the Claims sheet are not set on Render. The claim form downloaded to your computer is the only copy.'
    });
  }

  const warnings = [];
  let folder = null;
  let formUrl = null;
  let filed = 0;

  try {
    if (DRIVE_READY) {
      folder = await claimFolder(folderName, job);

      // The form itself, as a PDF. If the conversion fails the Word file is
      // filed instead — a claim must never end up with no form at all.
      const base = `Cash Purchase Claim ${folderName}`;
      const pdf = await g.docxToPdf(job.docx, base, folder.id);
      if (pdf.ok) {
        formUrl = pdf.url;
      } else {
        const doc = await g.uploadFile(job.docx, `${base}.docx`, folder.id);
        formUrl = doc.url;
        warnings.push(`The form was filed as Word, not PDF: ${clean(pdf.error, 120)}`);
      }

      // The scans, numbered to match the lines on the form.
      for (const l of c.lines) {
        if (!l.fileId) continue;
        const f = job.files.get(l.fileId);
        if (!f) { warnings.push(`The scan for line ${l.slNo} was no longer held and could not be filed.`); continue; }
        const stem = safeName(String(f.name).replace(/\.pdf$/i, '')) || `voucher ${l.slNo}`;
        try {
          await g.uploadFile(f.buffer, `Voucher ${l.slNo} - ${stem}.pdf`, folder.id, g.PDF_MIME);
          filed++;
        } catch (err) {
          warnings.push(`Voucher ${l.slNo} could not be filed: ${clean(err.message, 120)}`);
        }
      }
    } else {
      warnings.push('DRIVE_FOLDER_CLAIMS is not set, so nothing was filed to Google Drive.');
    }

    let row = null;
    if (SHEET_READY) {
      row = await appendClaim({
        claimNo: c.claimNo,
        date: c.date,
        payTo: c.payTo,
        lawyer: c.lawyer,
        voucherType: c.voucherType,
        description: clean(c.lines.map(l => l.description).join('; '), 250),
        vouchers: filed,
        total: c.total,
        gst: c.gst,
        folderUrl: folder ? folder.url : null
      });
    } else {
      warnings.push('No Claims sheet is set, so no row was written.');
    }

    job.saved = {
      folderName,
      folderUrl: folder ? folder.url : null,
      formUrl,
      vouchers: filed,
      row,
      warning: warnings.length ? warnings.join(' ') : null
    };
    // The scans are filed; stop holding them.
    job.files.clear();
    job.bytes = 0;

    console.log(`Petty-cash claim ${c.claimNo} filed: folder ${folderName}, ${filed} voucher(s), sheet row ${row || '-'}`);
    res.json({ ok: true, ...job.saved });
  } catch (err) {
    console.error('Petty-cash save failed:', err.message);
    if (err.collision) return res.status(409).json({ ok: false, error: clean(err.message, 300) });
    const why = /permission|forbidden/i.test(err.message)
      ? 'the Claim Forms folder or the Claims sheet is not shared with this app as Editor'
      : /not found/i.test(err.message) ? 'check DRIVE_FOLDER_CLAIMS and the Claims sheet ID'
      : clean(err.message, 160);
    res.status(502).json({ ok: false, error: `The claim could not be filed (${why}). Your downloaded copy is safe — try Save again.` });
  }
});

// What this module can actually see. No key value ever leaves here.
function status() {
  return {
    ai: !!GEMINI_KEY,
    model: GEMINI_KEY ? (active || MODELS[0]) : null,
    models: MODELS,
    folderClaims: !!FOLDER_CLAIMS,
    sheetConfigured: !!SHEET_ID,
    sheetTab: SHEET_TAB,
    drive: DRIVE_READY,
    sheet: SHEET_READY
  };
}

module.exports = { router, status, readVoucher, loadModels, ensureHeader, HEADERS };
