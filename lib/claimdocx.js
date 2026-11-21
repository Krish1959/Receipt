'use strict';

// Builds the practice's "Cash Purchase Claim" as a Word .docx, from scratch.
//
// There is no .docx template for this form — the office works from a printed
// sheet — so the document is assembled here as WordprocessingML and zipped
// with JSZip, which the app already carries for the invoice module. Nothing
// new has to be installed, and no template file has to live in Drive.
//
// The layout follows the office sheet: title, Ref No / Pay To, the three
// voucher tick-boxes, a five-row claim table with a TOTAL line, and the
// Submitted / Checked / Approved sign-off block.

const JSZip = require('jszip');

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

const VOUCHER_TYPES = ['Petty Cash Voucher', 'Payment Voucher', 'Cash Voucher'];
const MIN_ROWS = 5;            // the printed sheet has five lines; blanks are kept
const MAX_ROWS = 5;            // and never more than five vouchers per claim

/* ---------- small XML helpers ---------- */

const esc = s => String(s ?? '').replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]));

// A run of text. Line breaks inside the text become real Word line breaks.
function run(text, { bold, italic, size, color } = {}) {
  const rpr =
    (bold ? '<w:b/>' : '') +
    (italic ? '<w:i/>' : '') +
    (size ? `<w:sz w:val="${size}"/><w:szCs w:val="${size}"/>` : '') +
    (color ? `<w:color w:val="${color}"/>` : '');
  const body = String(text ?? '').split('\n')
    .map(line => `<w:t xml:space="preserve">${esc(line)}</w:t>`)
    .join('<w:br/>');
  return `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}${body}</w:r>`;
}

// A Wingdings tick-box: reliable in every Word build, unlike a Unicode ☐.
const box = ticked => `<w:r><w:sym w:font="Wingdings" w:char="${ticked ? 'F0FE' : 'F0A8'}"/></w:r>`;

function para(runs, { align, spaceAfter = 0, spaceBefore = 0 } = {}) {
  const ppr =
    '<w:pPr>' +
    (align ? `<w:jc w:val="${align}"/>` : '') +
    `<w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}" w:line="240" w:lineRule="auto"/>` +
    '</w:pPr>';
  return `<w:p>${ppr}${Array.isArray(runs) ? runs.join('') : runs}</w:p>`;
}

const text = (s, opt = {}) => para(run(s, opt), opt);
const blank = (size = 18) => para(run('', { size }));

// One table cell. `span` merges it across following grid columns.
function cell(width, content, { shade, span, valign = 'center', borders = true } = {}) {
  const tcpr =
    '<w:tcPr>' +
    `<w:tcW w:w="${width}" w:type="dxa"/>` +
    (span ? `<w:gridSpan w:val="${span}"/>` : '') +
    (shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${shade}"/>` : '') +
    (borders ? '' : '<w:tcBorders>' +
      ['top', 'left', 'bottom', 'right'].map(s => `<w:${s} w:val="nil"/>`).join('') +
      '</w:tcBorders>') +
    `<w:vAlign w:val="${valign}"/>` +
    '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:left w:w="100" w:type="dxa"/>' +
    '<w:bottom w:w="60" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>' +
    '</w:tcPr>';
  return `<w:tc>${tcpr}${content}</w:tc>`;
}

function row(cells, { height } = {}) {
  const trpr = height ? `<w:trPr><w:trHeight w:val="${height}"/></w:trPr>` : '';
  return `<w:tr>${trpr}${cells.join('')}</w:tr>`;
}

function table(grid, rows, { align = 'center', borders = true } = {}) {
  const width = grid.reduce((a, b) => a + b, 0);
  const edge = borders
    ? '<w:tblBorders>' +
      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        .map(s => `<w:${s} w:val="single" w:sz="8" w:space="0" w:color="000000"/>`).join('') +
      '</w:tblBorders>'
    : '';
  return '<w:tbl><w:tblPr>' +
    `<w:tblW w:w="${width}" w:type="dxa"/>` +
    `<w:jc w:val="${align}"/>` +
    edge +
    '<w:tblLayout w:type="fixed"/>' +
    '</w:tblPr>' +
    `<w:tblGrid>${grid.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>` +
    rows.join('') +
    '</w:tbl>';
}

/* ---------- money ---------- */

const num = v => {
  const n = Number(String(v ?? '').replace(/[,\s$]/g, ''));
  return Number.isFinite(n) ? n : 0;
};
// Worked in cents throughout so the TOTAL always equals the lines exactly.
const cents = v => Math.round(num(v) * 100);
const fmt = c => (c / 100).toFixed(2);

/* ---------- the claim form ---------- */

// lines: [{ description, amount, gst }]  — up to five, already checked on screen
// Returns { buffer, total, totalGst, lines } with the totals the document shows.
async function buildClaim({ claimNo, date, payTo, voucherType, lines = [], preparedBy = '' }) {
  const kept = lines.slice(0, MAX_ROWS).map(l => ({
    description: String(l.description ?? '').replace(/\s+/g, ' ').trim().slice(0, 120),
    amountCents: cents(l.amount),
    gstCents: cents(l.gst)
  }));
  const totalCents = kept.reduce((a, l) => a + l.amountCents, 0);
  const gstCents = kept.reduce((a, l) => a + l.gstCents, 0);
  const type = VOUCHER_TYPES.includes(voucherType) ? voucherType : VOUCHER_TYPES[0];

  const GRID = [850, 4710, 1700, 2100];          // Sl No | Description | Amount | Remark
  const SIGN = [1560, 1900, 1900, 1900];         // sign-off block

  /* header */
  const head = [
    text('RAJ GOVIN LAW PRACTICE', { bold: true, size: 26, align: 'center' }),
    text("People's Park Centre, #05-48, Singapore 058357", { size: 17, align: 'center', spaceAfter: 140 }),
    text('Cash Purchase Claim', { bold: true, size: 30, align: 'center', spaceAfter: 160 })
  ].join('');

  /* Ref No / Pay To — a borderless two-column table keeps them aligned */
  const refRow = table([4680, 4680], [row([
    cell(4680, para([run('Ref No:  ', { bold: true }), run(claimNo)]), { borders: false, valign: 'top' }),
    cell(4680, para([run('Pay To:  ', { bold: true }), run(payTo || '')]), { borders: false, valign: 'top' })
  ])], { align: 'left', borders: false });

  const dateRow = table([4680, 4680], [row([
    cell(4680, para([run('Date:  ', { bold: true }), run(date || '')]), { borders: false, valign: 'top' }),
    cell(4680, para(run('')), { borders: false, valign: 'top' })
  ])], { align: 'left', borders: false });

  /* the three tick-boxes */
  const ticks = para(
    VOUCHER_TYPES.flatMap(t => [box(t === type), run(` ${t}` + (t === VOUCHER_TYPES[2] ? '' : '     '))]),
    { align: 'center', spaceBefore: 120, spaceAfter: 160 }
  );

  /* the claim table */
  const header = row([
    cell(GRID[0], text('Sl No', { bold: true, size: 20 }), { shade: 'EDF2FA' }),
    cell(GRID[1], text('Description', { bold: true, size: 20 }), { shade: 'EDF2FA' }),
    cell(GRID[2], text('Amount (S$)', { bold: true, size: 20, align: 'right' }), { shade: 'EDF2FA' }),
    cell(GRID[3], text('Remark', { bold: true, size: 20 }), { shade: 'EDF2FA' })
  ], { height: 420 });

  const body = [];
  for (let i = 0; i < Math.max(MIN_ROWS, kept.length); i++) {
    const l = kept[i];
    body.push(row([
      cell(GRID[0], text(String(i + 1), { size: 20 })),
      cell(GRID[1], text(l ? l.description : '', { size: 20 }), { valign: 'top' }),
      cell(GRID[2], text(l ? fmt(l.amountCents) : '', { size: 20, align: 'right' })),
      cell(GRID[3], text(l ? `GST $${fmt(l.gstCents)}` : '', { size: 20 }))
    ], { height: 460 }));
  }

  const totalRow = row([
    cell(GRID[0] + GRID[1], text('TOTAL', { bold: true, size: 20, align: 'right' }), { span: 2, shade: 'F5F7FA' }),
    cell(GRID[2], text(fmt(totalCents), { bold: true, size: 20, align: 'right' }), { shade: 'F5F7FA' }),
    cell(GRID[3], text(`GST $${fmt(gstCents)}`, { bold: true, size: 20 }), { shade: 'F5F7FA' })
  ], { height: 440 });

  const claimTable = table(GRID, [header, ...body, totalRow]);

  /* sign-off */
  const signRow = (label, cellsText) => row([
    cell(SIGN[0], text(label, { bold: true, size: 20 }), { valign: 'center' }),
    ...cellsText.map((t, i) => cell(SIGN[i + 1], text(t, { size: 20 }), { valign: 'center' }))
  ], { height: label === 'Signature' ? 900 : 420 });

  const signOff = table(SIGN, [
    row([
      cell(SIGN[0], text('', { size: 20 }), { shade: 'F5F7FA' }),
      cell(SIGN[1], text('Submitted by', { bold: true, size: 20 }), { shade: 'F5F7FA' }),
      cell(SIGN[2], text('Checked by', { bold: true, size: 20 }), { shade: 'F5F7FA' }),
      cell(SIGN[3], text('Approved', { bold: true, size: 20 }), { shade: 'F5F7FA' })
    ], { height: 420 }),
    signRow('Name', [preparedBy || payTo || '', '', '']),
    signRow('Date', [date || '', '', '']),
    signRow('Signature', ['', '', ''])
  ], { align: 'left' });

  const note = text(
    'The practice is not GST registered. GST shown above is the amount printed on each voucher and is claimed as part of the cost.',
    { size: 16, italic: true, color: '6B7380', spaceBefore: 200 }
  );

  const sectPr =
    '<w:sectPr>' +
    '<w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>';

  const document =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:body>' +
    head + refRow + dateRow + ticks + blank(10) +
    claimTable + blank(24) + signOff + note +
    sectPr +
    '</w:body></w:document>';

  const zip = new JSZip();
  zip.file('[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
    '</Types>');
  zip.folder('_rels').file('.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
    '</Relationships>');
  const word = zip.folder('word');
  word.file('document.xml', document);
  word.folder('_rels').file('document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
    '</Relationships>');
  word.file('styles.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:docDefaults><w:rPrDefault><w:rPr>' +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' +
    '<w:sz w:val="22"/><w:szCs w:val="22"/>' +
    '</w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
    '</w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
    '<w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '</w:styles>');

  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });

  return {
    buffer,
    total: totalCents / 100,
    totalGst: gstCents / 100,
    lines: kept.map((l, i) => ({
      slNo: i + 1,
      description: l.description,
      amount: l.amountCents / 100,
      gst: l.gstCents / 100,
      remark: `GST $${fmt(l.gstCents)}`
    }))
  };
}

module.exports = { buildClaim, VOUCHER_TYPES, MAX_ROWS, MIN_ROWS, DOCX_MIME };
