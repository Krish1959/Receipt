'use strict';

// Minimal .docx handling for the invoice / receipt pair.
//
// The practice's template keeps every value we touch inside a single Word run,
// so plain string replacement on word/document.xml is safe here. Nothing else
// in the file is altered: styles, tables, fonts and the letterhead survive
// exactly as the lawyer left them.

const JSZip = require('jszip');

const DOC_XML = 'word/document.xml';
const PLACEHOLDER_SERIAL = 'ORG-010-26';
const PLACEHOLDER_DATE = '27-10-2026';
const SAMPLE_WATERMARK = 'SAMPLE TEMPLATE - NOT FOR PAYMENT';
const DATE_RE = /\b(\d{2})-(\d{2})-(\d{4})\b/g;

const xmlEsc = s => String(s ?? '').replace(/[&<>"']/g, ch =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[ch]));
const xmlUnesc = s => String(s).replace(/&(amp|lt|gt|quot|apos|#39);/g, (_, e) =>
  ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", '#39': "'" }[e]));

const splitAll = (hay, needle, rep) => hay.split(needle).join(rep);

/* ---------- dates ---------- */

// "19-11-2026" -> Date (UTC midnight); null if unreadable
function parseDmy(text) {
  const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(String(text || '').trim());
  if (!m) return null;
  const d = new Date(Date.UTC(+m[3], +m[2] - 1, +m[1]));
  return Number.isNaN(d.getTime()) ? null : d;
}
const fmtDmy = d => `${String(d.getUTCDate()).padStart(2, '0')}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${d.getUTCFullYear()}`;
const addDays = (d, n) => new Date(d.getTime() + n * 86400000);

// "19-11-2026" + 3 -> "22-11-2026"; unreadable input is returned unchanged
function shiftDmy(text, days) {
  const d = parseDmy(text);
  return d ? fmtDmy(addDays(d, days)) : text;
}

/* ---------- zip in / out ---------- */

async function readDocumentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file(DOC_XML);
  if (!file) throw new Error('That file is not a Word document (word/document.xml is missing).');
  return { zip, xml: await file.async('string') };
}

async function writeDocumentXml(zip, xml) {
  zip.file(DOC_XML, xml);
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

/* ---------- plain text ---------- */

// Paragraph-per-line plain text, used only for reading values back out.
function plainText(xml) {
  // Runs of spaces and tabs are KEPT: the template uses them to separate the
  // client's name from the address inside one paragraph, and the split below
  // depends on them.
  let t = xml.replace(/<w:tab[^>]*\/>/g, '\t')
             .replace(/<\/w:p>/g, '\n')
             .replace(/<\/w:tc>/g, '  |  ')
             .replace(/<[^>]+>/g, '');
  return xmlUnesc(t).split('\n').map(l => l.replace(/[ \t]+$/, '').trim()).filter(Boolean).join('\n');
}

/* ---------- build the blank invoice the lawyer downloads ---------- */

// Stamps the serial number and issue date into the template and drops the
// sample watermark. Everything else is left for the lawyer to fill in Word.
async function buildInvoice(templateBuffer, { slNo, issueDate, dueDate }) {
  const { zip, xml } = await readDocumentXml(templateBuffer);
  let out = splitAll(xml, PLACEHOLDER_SERIAL, xmlEsc(slNo));
  // The template carries the same date twice (Issue Date, Due Date).
  const parts = out.split(PLACEHOLDER_DATE);
  out = parts.length === 3
    ? parts[0] + xmlEsc(issueDate) + parts[1] + xmlEsc(dueDate || issueDate) + parts[2]
    : splitAll(out, PLACEHOLDER_DATE, xmlEsc(issueDate));
  out = splitAll(out, SAMPLE_WATERMARK, '');
  return writeDocumentXml(zip, out);
}

/* ---------- turn the completed invoice into its receipt ---------- */

// Same document, same serial, same amounts. Only the label changes and every
// date moves on by `days` (the practice draws the money 3 days after issue).
async function invoiceToReceipt(invoiceBuffer, days = 3) {
  const { zip, xml } = await readDocumentXml(invoiceBuffer);
  let out = xml;
  const dates = [];
  out = out.replace(DATE_RE, m => { dates.push(m); return shiftDmy(m, days); });
  out = splitAll(out, 'INVOICE', 'RECEIPT');
  out = splitAll(out, 'Invoice No.', 'Receipt No.');
  out = splitAll(out, 'Invoice no.', 'Receipt no.');
  out = splitAll(out, 'Invoice', 'Receipt');
  out = splitAll(out, SAMPLE_WATERMARK, '');
  return {
    buffer: await writeDocumentXml(zip, out),
    datesShifted: dates.length,
    receiptDate: dates.length ? shiftDmy(dates[0], days) : ''
  };
}

/* ---------- read the key fields back out for the sheet ---------- */

const after = (line) => {
  const i = line.indexOf(':');
  return i < 0 ? '' : line.slice(i + 1).trim();
};

// Best-effort only. Whatever this misses, the lawyer corrects on screen before
// the row is written, so a parsing miss is never silently saved.
function extractFields(xml) {
  const text = plainText(xml);
  const lines = text.split('\n');
  const find = re => lines.find(l => re.test(l)) || '';

  // Client: the "To :" line, cut at the first run of spaces that starts the address
  let client = after(find(/^To\b/i));
  client = client.split(/\t|\s{2,}|\s*\|\s*/)[0].trim();
  if (!client) client = after(find(/^Att\b/i)).split(/\t|\s{2,}/)[0].trim();

  const serial = after(find(/(Invoice|Receipt)\s*No\.?/i)).split(/\t|\s{2,}|\s*\|\s*/)[0].trim();
  const issueDate = (/(\d{2}-\d{2}-\d{4})/.exec(find(/Issue\s*Date/i)) || [])[1] || '';
  const dueDate = (/(\d{2}-\d{2}-\d{4})/.exec(find(/Due\s*Date/i)) || [])[1] || '';

  // Description: the "Re :" matter line, falling back to the "Being …" narrative
  let description = after(find(/^Re\b/i)).split(/\s*\|\s*/)[0].replace(/\s+/g, ' ').trim();
  if (!description) description = (find(/^Being\b/i) || '').replace(/\s*\|\s*/g, ' ').replace(/\s+/g, ' ').trim();

  // Total: the last "Total" line wins (Sub-Total comes first in the template)
  const totalLines = lines.filter(l => /total/i.test(l) && /[\d,]+\.\d{2}/.test(l));
  const pickTotal = l => {
    const all = [...String(l).matchAll(/([\d,]+\.\d{2})/g)].map(m => m[1]);
    return all.length ? all[all.length - 1] : '';
  };
  let total = '';
  for (const l of totalLines) if (!/sub-?total/i.test(l)) total = pickTotal(l) || total;
  if (!total && totalLines.length) total = pickTotal(totalLines[totalLines.length - 1]);

  return { client, serial, issueDate, dueDate, description, total: total.replace(/,/g, '') };
}

async function extractFromBuffer(buffer) {
  const { xml } = await readDocumentXml(buffer);
  return extractFields(xml);
}

module.exports = {
  buildInvoice, invoiceToReceipt, extractFromBuffer, extractFields,
  plainText, parseDmy, fmtDmy, addDays, shiftDmy
};
