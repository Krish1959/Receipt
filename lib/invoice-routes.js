'use strict';

// Module 1 – the Invoice page.
//
// The lawyer's round trip:
//   1. GET  /api/invoice/template   – a Word invoice stamped with serial + date
//   2. POST /api/invoice/invoice    – the completed invoice comes back; it is
//                                     filed, and its receipt is generated
//   3. GET  /api/invoice/receipt    – that receipt, to check in Word
//   4. POST /api/invoice/receipt    – the checked receipt is filed
//   5. POST /api/invoice/finalise   – one row in the Invoices sheet
//
// The practice draws the money three days after the invoice is issued, so the
// receipt carries the same serial and totals with every date moved on by three.
//
// DATES: the Word documents keep dd-mm-yyyy — that is the printed layout and
// what the +3 day rule matches on. Only the two date cells written to the
// sheet are converted, by lib/sheetdate.js, to "02 Nov 2026".

const fs = require('fs');
const path = require('path');
const express = require('express');

const docx = require('./docxtool');
const g = require('./google');
const { sheetDate } = require('./sheetdate');

const router = express.Router();

/* ---------- settings ---------- */
const FOLDER_INVOICES = (process.env.DRIVE_FOLDER_INVOICES || '').trim();
const FOLDER_RECEIPTS = (process.env.DRIVE_FOLDER_CASH_RECEIPTS || '').trim();
const TEMPLATE_FILE_ID = (process.env.INVOICE_TEMPLATE_FILE_ID || '').trim();
const SHEET_ID = (process.env.INVOICES_SHEET_ID || '').trim().replace(/^["']|["']$/g, '');
const SHEET_TAB = (process.env.INVOICES_SHEET_TAB || 'Invoices').trim();
const RECEIPT_OFFSET_DAYS = Number(process.env.RECEIPT_OFFSET_DAYS) || 3;
const LOCAL_TEMPLATE = path.join(__dirname, '..', 'templates', 'invoice-template.docx');

const DRY_RUN = g.authMode() === 'none';
const JOB_TTL_MS = 2 * 60 * 60 * 1000;          // a job is forgotten after 2 hours
const MAX_DOC_BYTES = 8 * 1024 * 1024;

// One row per invoice/receipt pair. INVOICE NO and RECEIPT NO both carry the
// same serial — the pair shares one number by design — but they are kept as
// two columns so a standalone receipt could be filed here later.
const HEADERS = [
  'INVOICE NO', 'RECEIPT NO', 'INVOICE DATE', 'RECEIPT DATE', 'CLIENT',
  'DESCRIPTION', 'TOTAL (S$)', 'INVOICE', 'RECEIPT', 'REMARKS', 'TIMESTAMP (SGT)'
];
const HEADER_RANGE = 'A1:N1';                  // wider than HEADERS so older columns are cleared
const COL_INVOICE_LINK = 'H';
const COL_RECEIPT_LINK = 'I';

/* ---------- helpers ---------- */
const clean = (v, max = 300) => String(v ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
const safeName = s => clean(s, 80).replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();

function sgtNow() {
  return new Date().toLocaleString('en-SG', {
    timeZone: 'Asia/Singapore', day: '2-digit', month: 'short', year: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false
  });
}

function amount(v) {
  const s = clean(v, 20).replace(/[,\s$]/g, '');
  if (!s) return '';
  if (!/^\d+(\.\d{1,2})?$/.test(s)) return NaN;
  return Math.round(parseFloat(s) * 100) / 100;
}

// Word files are sent as base64 in JSON, which keeps the upload dependency-free.
function decodeDocx(base64) {
  const buf = Buffer.from(String(base64 || '').replace(/^data:[^;]+;base64,/, ''), 'base64');
  if (!buf.length) throw new Error('The file came through empty. Please try the drag and drop again.');
  if (buf.length > MAX_DOC_BYTES) throw new Error('That file is larger than 8 MB.');
  if (buf[0] !== 0x50 || buf[1] !== 0x4b) throw new Error('That is not a Word .docx file.');
  return buf;
}

/* ---------- jobs in flight ---------- */
// One invoice being worked on. Held in memory only: a restart loses jobs that
// have not been finalised, and the lawyer starts that invoice again.
const jobs = new Map();

function putJob(job) {
  jobs.set(job.id, { ...job, at: Date.now() });
  for (const [id, j] of jobs) if (Date.now() - j.at > JOB_TTL_MS) jobs.delete(id);
  return job.id;
}
function getJob(id) {
  const job = jobs.get(clean(id, 40));
  if (!job) throw new Error('This invoice is no longer open — it may have timed out. Please start it again.');
  return job;
}
const newJobId = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);

/* ---------- the blank template ---------- */
let templateCache = null;

// Drive first (so the practice can update the template without a redeploy),
// then the copy bundled in the repo. If BOTH are unavailable the error says
// which one failed and why — a blank "could not build" tells nobody anything.
async function loadTemplate() {
  if (templateCache) return templateCache;
  const tried = [];

  if (!TEMPLATE_FILE_ID) tried.push('INVOICE_TEMPLATE_FILE_ID is not set');
  else if (DRY_RUN) tried.push('Drive skipped (no Google credentials set)');
  else {
    try {
      templateCache = await g.downloadDoc(TEMPLATE_FILE_ID);
      return templateCache;
    } catch (err) {
      const why = /404|not found/i.test(err.message)
        ? `not found, or not shared with this app's Google identity (${g.authMode()})`
        : err.message;
      tried.push(`Drive file ${TEMPLATE_FILE_ID}: ${why}`);
      console.error('Template not readable from Drive:', err.message);
    }
  }

  try {
    templateCache = fs.readFileSync(LOCAL_TEMPLATE);
    console.log('Using the bundled invoice template.');
    return templateCache;
  } catch (err) {
    tried.push(`bundled templates/invoice-template.docx: ${err.code === 'ENOENT' ? 'missing from the deployment' : err.message}`);
  }

  throw new Error(`No invoice template available — ${tried.join('; ')}`);
}

/* ---------- Drive ---------- */
// Files the Word document and its PDF side by side. In dry run nothing leaves
// the server, so the page can be tested without credentials.
async function fileDocument(buffer, baseName, folderId) {
  if (DRY_RUN || !folderId) {
    console.log(`[DRY RUN] ${baseName}.docx not sent to Drive (${buffer.length} bytes)`);
    return { docUrl: null, pdfUrl: null, docId: null, dryRun: true };
  }
  const doc = await g.uploadFile(buffer, `${baseName}.docx`, folderId);
  const pdf = await g.docxToPdf(buffer, baseName, folderId);
  return { docUrl: doc.url, docId: doc.id, pdfUrl: pdf.ok ? pdf.url : null, pdfError: pdf.ok ? null : pdf.error };
}

/* ---------- the Invoices sheet ---------- */
let headerChecked = false;

async function ensureHeader() {
  if (headerChecked || DRY_RUN || !SHEET_ID) return;
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
    const padded = [...HEADERS, '', '', ''];
    await api.spreadsheets.values.update({
      spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!${HEADER_RANGE}`,
      valueInputOption: 'RAW', requestBody: { values: [padded] }
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
              properties: { sheetId: tabId, gridProperties: { frozenRowCount: 1, frozenColumnCount: 2 } },
              fields: 'gridProperties(frozenRowCount,frozenColumnCount)'
            }
          },
          ...[[0, 150], [1, 150], [2, 110], [3, 110], [4, 210], [5, 300], [6, 100], [7, 80], [8, 80], [9, 200], [10, 160]].map(([i, px]) => ({
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

// The row is written as plain values first, then the two link cells are set as
// formulas. Nothing the lawyer typed is ever evaluated as a formula.
//
// The two dates go in as "02 Nov 2026". They arrive here as dd-mm-yyyy, which
// is what the Word documents use and what the form validates — the conversion
// happens at this boundary and nowhere else.
async function appendPair(rec) {
  if (DRY_RUN || !SHEET_ID) {
    console.log('[DRY RUN] Invoice row not sent to Google Sheets:', rec);
    return null;
  }
  await ensureHeader();
  const api = g.getSheets();
  const res = await api.spreadsheets.values.append({
    spreadsheetId: SHEET_ID, range: `'${SHEET_TAB}'!A1`,
    valueInputOption: 'RAW', insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[
        rec.slNo,                    // A  INVOICE NO
        rec.slNo,                    // B  RECEIPT NO — the pair shares one serial
        sheetDate(rec.date),         // C  INVOICE DATE
        sheetDate(rec.receiptDate),  // D  RECEIPT DATE
        rec.client,                  // E
        rec.description,             // F
        rec.total,                   // G
        '',                          // H  INVOICE link, written below
        '',                          // I  RECEIPT link, written below
        rec.remarks,                 // J
        sgtNow()                     // K
      ]]
    }
  });
  const m = /![A-Z]+(\d+)/.exec((res.data.updates && res.data.updates.updatedRange) || '');
  const rowNum = m ? Number(m[1]) : null;
  if (rowNum) {
    const link = (url, label) => (url ? [`=HYPERLINK("${url}","${label}")`] : ['']);
    await api.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          { range: `'${SHEET_TAB}'!${COL_INVOICE_LINK}${rowNum}`, values: [link(rec.invoiceUrl, 'Invoice')] },
          { range: `'${SHEET_TAB}'!${COL_RECEIPT_LINK}${rowNum}`, values: [link(rec.receiptUrl, 'Receipt')] }
        ]
      }
    });
  }
  return rowNum;
}

/* ================= routes ================= */

router.get('/config', (_req, res) => {
  res.json({
    ok: true,
    dryRun: DRY_RUN,
    authMode: g.authMode(),
    receiptOffsetDays: RECEIPT_OFFSET_DAYS,
    sheetUrl: SHEET_ID ? `https://docs.google.com/spreadsheets/d/${SHEET_ID}/edit` : null,
    invoicesFolder: FOLDER_INVOICES ? `https://drive.google.com/drive/folders/${FOLDER_INVOICES}` : null,
    receiptsFolder: FOLDER_RECEIPTS ? `https://drive.google.com/drive/folders/${FOLDER_RECEIPTS}` : null
  });
});

// 1. The blank invoice, stamped with this serial and date.
router.get('/template', async (req, res) => {
  const slNo = clean(req.query.slNo, 40);
  const issueDate = clean(req.query.date, 10);
  if (!slNo) return res.status(400).json({ ok: false, error: 'Sl No is required.' });
  if (!docx.parseDmy(issueDate)) return res.status(400).json({ ok: false, error: 'Date must be dd-mm-yyyy.' });
  try {
    const out = await docx.buildInvoice(await loadTemplate(), { slNo, issueDate, dueDate: issueDate });
    res.setHeader('Content-Type', g.DOCX_MIME);
    res.setHeader('Content-Disposition', `attachment; filename="Invoice ${safeName(slNo)}.docx"`);
    res.send(out);
  } catch (err) {
    console.error('Template build failed:', err.message);
    res.status(500).json({ ok: false, error: clean(err.message, 400) });
  }
});

// 2. The completed invoice comes back: file it, then build its receipt.
router.post('/invoice', async (req, res) => {
  const body = req.body || {};
  try {
    const buf = decodeDocx(body.base64);
    const found = await docx.extractFromBuffer(buf);
    const slNo = clean(body.slNo, 40) || found.serial;
    if (!slNo) return res.status(400).json({ ok: false, error: 'No Sl No found in that invoice.' });

    const baseName = `Invoice ${safeName(slNo)}`;
    const filed = await fileDocument(buf, baseName, FOLDER_INVOICES);

    const made = await docx.invoiceToReceipt(buf, RECEIPT_OFFSET_DAYS);
    const receiptDate = made.receiptDate || docx.shiftDmy(found.issueDate, RECEIPT_OFFSET_DAYS);

    const id = putJob({
      id: newJobId(), slNo, receipt: made.buffer, receiptDate,
      invoiceUrl: filed.pdfUrl || filed.docUrl, fields: found,
      receiptFiled: null
    });

    res.json({
      ok: true, jobId: id, slNo, receiptDate,
      fields: { ...found, slNo },
      invoiceUrl: filed.pdfUrl || filed.docUrl,
      pdf: !!filed.pdfUrl,
      warning: filed.pdfError ? `Filed as Word. The PDF could not be made: ${clean(filed.pdfError, 160)}` : null,
      dryRun: !!filed.dryRun
    });
  } catch (err) {
    console.error('Invoice upload failed:', err.message);
    res.status(400).json({ ok: false, error: clean(err.message, 200) });
  }
});

// 3. The generated receipt, for the lawyer to check in Word.
router.get('/receipt', async (req, res) => {
  try {
    const job = getJob(req.query.jobId);
    res.setHeader('Content-Type', g.DOCX_MIME);
    res.setHeader('Content-Disposition', `attachment; filename="Receipt ${safeName(job.slNo)}.docx"`);
    res.send(job.receipt);
  } catch (err) {
    res.status(404).json({ ok: false, error: clean(err.message, 200) });
  }
});

// 4. The checked receipt comes back and is filed. Dropping it again replaces
//    the copy already in Drive, so one invoice never leaves two receipts.
router.post('/receipt', async (req, res) => {
  const body = req.body || {};
  try {
    const job = getJob(body.jobId);
    const buf = decodeDocx(body.base64);
    const found = await docx.extractFromBuffer(buf);

    // Two checks, and only these two. What the receipt SAYS — the wording, the
    // line items, the arithmetic — is the lawyer's business, not the app's.
    //
    //  1. the receipt must carry the invoice's own number, so the pair holds
    //  2. it must be dated at least RECEIPT_OFFSET_DAYS after the invoice,
    //     which is the rule the practice draws the money under
    const given = clean(found.serial, 40);
    if (!given) {
      throw new Error('No Receipt No. found in that file. Check the receipt still carries the invoice number.');
    }
    if (given.toUpperCase() !== String(job.slNo).toUpperCase()) {
      throw new Error(`Receipt No. ${given} does not match invoice ${job.slNo}. The pair must share one number.`);
    }

    const invoiceDate = docx.parseDmy((job.fields || {}).issueDate);
    const receiptDate = docx.parseDmy(found.issueDate);
    if (!receiptDate) {
      throw new Error('No readable Issue Date (dd-mm-yyyy) found in that receipt.');
    }
    if (invoiceDate) {
      const days = Math.round((receiptDate - invoiceDate) / 86400000);
      if (days < RECEIPT_OFFSET_DAYS) {
        throw new Error(
          `The receipt is dated ${found.issueDate}, ${days} day${days === 1 ? '' : 's'} after the invoice `
          + `(${(job.fields || {}).issueDate}). At least ${RECEIPT_OFFSET_DAYS} days are required.`
        );
      }
    }

    const baseName = `Receipt ${safeName(job.slNo)}`;

    let filed;
    if (job.receiptFiled && job.receiptFiled.docId && !DRY_RUN) {
      await g.replaceFile(job.receiptFiled.docId, buf);
      const pdf = await g.docxToPdf(buf, baseName, FOLDER_RECEIPTS);
      filed = { ...job.receiptFiled, pdfUrl: pdf.ok ? pdf.url : job.receiptFiled.pdfUrl };
    } else {
      filed = await fileDocument(buf, baseName, FOLDER_RECEIPTS);
    }

    job.receiptFiled = filed;
    job.receiptUrl = filed.pdfUrl || filed.docUrl;
    if (found.issueDate) job.receiptDate = found.issueDate;
    jobs.set(job.id, { ...job, at: Date.now() });

    res.json({
      ok: true, receiptUrl: job.receiptUrl, receiptDate: job.receiptDate,
      pdf: !!filed.pdfUrl, dryRun: !!filed.dryRun,
      warning: filed.pdfError ? `Filed as Word. The PDF could not be made: ${clean(filed.pdfError, 160)}` : null
    });
  } catch (err) {
    console.error('Receipt upload failed:', err.message);
    res.status(400).json({ ok: false, error: clean(err.message, 200) });
  }
});

// 5. One row for the pair, then the lawyer is done.
router.post('/finalise', async (req, res) => {
  const body = req.body || {};
  try {
    const job = getJob(body.jobId);
    const total = amount(body.total);
    const date = clean(body.date, 10);
    const errors = [];
    if (!docx.parseDmy(date)) errors.push('Date (dd-mm-yyyy)');
    if (!clean(body.client)) errors.push('Client');
    if (total === '' || Number.isNaN(total) || total <= 0) errors.push('Total');
    if (errors.length) return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });

    const rec = {
      date,
      slNo: clean(body.slNo, 40) || job.slNo,
      client: clean(body.client, 120),
      description: clean(body.description, 160),      // the short version – the file holds the full text
      total,
      receiptDate: clean(body.receiptDate, 10) || job.receiptDate,
      remarks: clean(body.remarks, 200),
      invoiceUrl: job.invoiceUrl,
      receiptUrl: job.receiptUrl || null
    };

    const row = await appendPair(rec);
    jobs.delete(job.id);
    res.json({ ok: true, row, dryRun: DRY_RUN });
  } catch (err) {
    console.error('Finalise failed:', err.message);
    res.status(502).json({ ok: false, error: clean(err.message, 200) });
  }
});

// What the invoice module can actually see. Booleans only — no values leave here.
function status() {
  return {
    sheetConfigured: !!SHEET_ID,
    sheetTab: SHEET_TAB,
    folderInvoices: !!FOLDER_INVOICES,
    folderCashReceipts: !!FOLDER_RECEIPTS,
    templateFile: !!TEMPLATE_FILE_ID,
    dryRun: DRY_RUN
  };
}

module.exports = { router, appendPair, ensureHeader, status, HEADERS, DRY_RUN };
