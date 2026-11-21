'use strict';

// Every date written into the Accounts spreadsheet goes through here.
//
// Dates go into the sheet carrying their own month name: "02 Nov 2026".
// A bare 02-11-2026 is day-first only by convention; a human who clicks that
// cell and retypes it has the value reparsed by the spreadsheet's locale,
// which is how two dates reached TallyPrime as February instead of November.
// "02 Nov 2026" survives that. No locale can reinterpret it, no number format
// can redisplay it, and nobody reading the sheet can misread it.
//
// Used at the point of writing to the sheet and NOWHERE else. The forms still
// accept dd-mm-yyyy, the Word invoice and receipt still carry dd-mm-yyyy in
// their printed body, and docxtool's +3 day rule still matches that pattern.
// This converts at the sheet boundary only.
//
// It must also be paired with valueInputOption: 'RAW' on the write. RAW is
// what keeps the text as text; USER_ENTERED would have Sheets parse the
// string back into a date value and undo the whole point.

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

const build = (d, m, y) => `${String(d).padStart(2, '0')} ${MONTHS[m - 1]} ${y}`;
const sane = (d, m) => m >= 1 && m <= 12 && d >= 1 && d <= 31;

// "02-11-2026" / "2-11-2026" / "2026-11-02"  ->  "02 Nov 2026"
// "02 Nov 2026"                              ->  unchanged
// anything else                              ->  unchanged
function sheetDate(value) {
  const s = String(value ?? '').trim();
  if (!s) return '';

  const dmy = /^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/.exec(s);          // day first
  if (dmy) {
    const d = Number(dmy[1]), m = Number(dmy[2]), y = Number(dmy[3]);
    if (sane(d, m)) return build(d, m, y);
  }

  const ymd = /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/.exec(s);          // year first
  if (ymd) {
    const y = Number(ymd[1]), m = Number(ymd[2]), d = Number(ymd[3]);
    if (sane(d, m)) return build(d, m, y);
  }

  // Already "02 Nov 2026", or something unexpected. Pass it through rather
  // than throwing: a date that cannot be parsed must still reach the sheet,
  // where it is visible and can be corrected, instead of failing a lawyer's
  // submission at the desk.
  return s;
}

module.exports = { sheetDate, MONTHS };
