'use strict';

// Module 1 – the Petty-cash page.
//
// The clerk's round trip:
//   1. POST /api/pettycash/extract  – one scanned voucher (PDF), read by AI
//                                     into a single line: description, amount, GST
//   2. POST /api/pettycash/build    – the checked lines come back as a Word
//                                     "Cash Purchase Claim", ready to print
//
// The AI only READS. Every line it returns is shown on screen and can be
// corrected before the claim form is made, so a misread amount is never
// filed silently. Nothing here writes to Drive or to a sheet — that is
// Part 4.

const express = require('express');
const claim = require('./claimdocx');

const router = express.Router();

/* ---------- settings ---------- */
const GEMINI_KEY = (process.env.GEMINI_API_KEY || '').trim();
const GEMINI_MODEL = (process.env.GEMINI_MODEL || 'gemini-2.5-flash').trim();
const GEMINI_URL = m => `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(m)}:generateContent`;

const MAX_PDF_BYTES = 6 * 1024 * 1024;        // one voucher
const AI_TIMEOUT_MS = 90 * 1000;

/* ---------- helpers ---------- */
const clean = (v, max = 300) => String(v ?? '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
const safeName = s => clean(s, 80).replace(/[\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();

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

async function readVoucher(pdfBase64) {
  if (!GEMINI_KEY) {
    throw new Error('No AI key is set on the server. Add GEMINI_API_KEY, or type the line in by hand.');
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), AI_TIMEOUT_MS);
  let res;
  try {
    res = await fetch(GEMINI_URL(GEMINI_MODEL), {
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
    let why = clean(raw, 200);
    try { why = clean(JSON.parse(raw).error.message, 200); } catch (_) {}
    if (res.status === 400 && /API key/i.test(why)) why = 'the AI key was refused — check GEMINI_API_KEY';
    if (res.status === 429) why = 'the AI quota for this key is used up for now';
    throw new Error(`The AI could not read that voucher (${why}).`);
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

/* ================= routes ================= */

router.get('/config', (_req, res) => {
  res.json({
    ok: true,
    ai: !!GEMINI_KEY,
    model: GEMINI_KEY ? GEMINI_MODEL : null,
    maxFiles: claim.MAX_ROWS,
    voucherTypes: claim.VOUCHER_TYPES
  });
});

// 1. One voucher in, one claim line out.
router.post('/extract', async (req, res) => {
  const body = req.body || {};
  const name = clean(body.name, 120) || 'voucher.pdf';
  try {
    const buf = decodePdf(body.base64);
    const line = await readVoucher(buf.toString('base64'));
    if (!line.description && line.amount === '') {
      throw new Error('Nothing readable was found on that page. Please check the scan, or type the line in by hand.');
    }
    res.json({ ok: true, name, line });
  } catch (err) {
    console.error(`Petty-cash extract failed for ${name}:`, err.message);
    res.status(400).json({ ok: false, name, error: clean(err.message, 240) });
  }
});

// 2. The checked lines become the Word claim form.
router.post('/build', async (req, res) => {
  const body = req.body || {};
  const claimNo = clean(body.claimNo, 40);
  const date = clean(body.date, 10);
  const payTo = clean(body.payTo, 120);
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
    return { description, amount, gst: gst === '' ? 0 : gst };
  });

  if (errors.length) {
    return res.status(400).json({ ok: false, error: `Please check: ${errors.join(', ')}` });
  }

  try {
    const made = await claim.buildClaim({ claimNo, date, payTo, voucherType, lines });
    res.json({
      ok: true,
      claimNo,
      date,
      payTo,
      voucherType: claim.VOUCHER_TYPES.includes(voucherType) ? voucherType : claim.VOUCHER_TYPES[0],
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

// What this module can actually see. Booleans only.
function status() {
  return { ai: !!GEMINI_KEY, model: GEMINI_KEY ? GEMINI_MODEL : null };
}

module.exports = { router, status, readVoucher };
