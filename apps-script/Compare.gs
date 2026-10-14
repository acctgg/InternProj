/**
 * Logbook Scanner: Spreadsheet 2 records for the panel, and the photo review.
 *
 * The panel lists Spreadsheet 2's records for the submission's branch, the
 * logbook date and an optional search. A person compares them with the photo
 * and marks values that don't match; the result is saved in the "Review" tab.
 */

const REVIEW_HEADERS = ['Submission', 'Response row', 'Branch', 'Logbook date', 'Spreadsheet 2 row',
  'Column', 'Header', 'Spreadsheet 2 value', 'Status', 'Photo says', 'Reviewed by', 'Reviewed at'];
const REVIEW = { MATCH: 'Match', NOT_MATCH: 'Not match', NOT_IN_PHOTO: 'Not in photo', ONLY_IN_PHOTO: 'Only in photo' };
const REVIEW_COLORS = { 'Match': '#e6f6ef', 'Not match': '#fdecee', 'Not in photo': '#fdecee', 'Only in photo': '#fff4e0' };
const RECORD_LIMIT = 300;

/**
 * Spreadsheet 2 records matching filter {date, branch, q}.
 * date: the date written on the logbook page (blank = any date, newest first).
 * q: words that must all appear in the row's SHOW columns or date.
 */
function getRecords(filter) {
  filter = filter || {};
  const s = getSettings_();
  const mapping = readMapping_(false);
  const dateMap = mapping.filter(m => m.field === DATE)[0];
  if (!dateMap) throw new Error('In "' + SHEETS.CONFIG + '", set one Spreadsheet 2 column to DATE.');
  const branchMap = mapping.filter(m => m.field === BRANCH)[0];
  const shown = mapping.filter(m => m.field === SHOW);
  const out = {
    columns: shown.map(m => ({ letter: m.letter, header: m.header, numeric: false })),
    rows: [], total: 0, duplicates: 0, truncated: false, branchFiltered: !!(branchMap && filter.branch),
  };

  const d = filter.date ? normDate_(filter.date) : '';
  if (filter.date && !d) throw new Error('Could not read the date "' + filter.date + '". Use mm/dd/yy.');
  const words = String(filter.q || '').toLowerCase().split(/\s+/).filter(String);

  const src = openSource_(s);
  const lastRow = src.getLastRow();
  if (lastRow <= s.headerRow) return out;
  const lastCol = Math.max.apply(null, mapping.map(m => m.col));
  const range = src.getRange(s.headerRow + 1, 1, lastRow - s.headerRow, lastCol);
  const vals = range.getValues();
  const disp = range.getDisplayValues();

  let idx = [];
  for (let i = 0; i < vals.length; i++) {
    const rowDate = normDate_(vals[i][dateMap.col - 1]);
    if (!rowDate || (d && rowDate !== d)) continue;
    if (out.branchFiltered && normText_(vals[i][branchMap.col - 1]) !== normText_(filter.branch)) continue;
    if (words.length) {
      const hay = (disp[i][dateMap.col - 1] + ' ' + shown.map(m => disp[i][m.col - 1]).join(' ')).toLowerCase();
      if (!words.every(w => hay.indexOf(w) >= 0)) continue;
    }
    idx.push(i);
  }
  if (s.dedupe) {
    // Spreadsheet 2 can list the same sale several times; count rows identical in every used column once.
    const seen = {};
    const before = idx.length;
    idx = idx.filter(i => {
      const sig = mapping.map(m => {
        const v = vals[i][m.col - 1];
        return v instanceof Date ? normDate_(v) : typeof v === 'number' ? String(round2_(v)) : normText_(v);
      }).join('\u0001');
      if (seen[sig]) return false;
      seen[sig] = true;
      return true;
    });
    out.duplicates = before - idx.length;
  }
  if (!d) {
    // Any date: newest first, sheet order within a day.
    const key = i => normDate_(vals[i][dateMap.col - 1]);
    idx.sort((a, b) => key(a) < key(b) ? 1 : key(a) > key(b) ? -1 : a - b);
  }
  out.total = idx.length;
  out.truncated = idx.length > RECORD_LIMIT;
  idx = idx.slice(0, RECORD_LIMIT);

  // A column is numeric when most of its filled cells hold numbers (right-aligned and totalled in the panel).
  out.columns.forEach((c, k) => {
    const col = shown[k].col - 1;
    const filled = idx.filter(i => vals[i][col] !== '');
    c.numeric = filled.length > 0 && filled.filter(i => typeof vals[i][col] === 'number').length * 2 > filled.length;
  });
  out.rows = idx.map(i => ({
    row: s.headerRow + 1 + i,
    date: shortDate_(vals[i][dateMap.col - 1]),
    cells: shown.map(m => disp[i][m.col - 1]),
    nums: shown.map(m => typeof vals[i][m.col - 1] === 'number' ? vals[i][m.col - 1] : null),
  }));
  return out;
}

/** What was saved for one submission: {date, marks: {row: {status, cells: {letter: {photo, value, header}}}}, extras: [photo text]} */
function getReview(key) {
  const out = { date: '', marks: {}, extras: [] };
  readReview_().filter(r => r[0] === key).forEach(r => {
    out.date = out.date || shortDate_(r[3]);
    const row = String(r[4]), letter = String(r[5]), status = String(r[8]), photo = String(r[9]);
    if (status === REVIEW.ONLY_IN_PHOTO) { out.extras.push(photo); return; }
    const m = out.marks[row] = out.marks[row] || { status: status, cells: {} };
    if (status === REVIEW.NOT_MATCH) { m.status = status; if (letter) m.cells[letter] = { photo: photo, value: String(r[7]), header: String(r[6]) }; }
    else if (m.status !== REVIEW.NOT_MATCH) m.status = status;
  });
  return out;
}

/**
 * Replaces the saved review of one submission.
 * lines: [{row, letter, header, value, status, photo}]; row/letter are blank for "Only in photo".
 */
function saveReview(key, responseRow, branch, date, lines) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const d = normDate_(date);
    if (!d) throw new Error('Type the date written on the logbook page before saving.');
    const sh = ensureSheet_(SHEETS.REVIEW, REVIEW_HEADERS);
    const last = sh.getLastRow();
    if (last > 1) {
      const keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = keys.length - 1; i >= 0; i--) if (keys[i][0] === key) sh.deleteRow(i + 2);
    }
    const who = Session.getActiveUser().getEmail();
    const out = lines.map(l => [key, responseRow, branch || '', d, l.row || '', l.letter || '', l.header || '',
      l.value == null ? '' : l.value, l.status, l.photo || '', who]
      .map(String).concat([new Date()]));
    if (out.length) {
      const start = sh.getLastRow() + 1;
      const rng = sh.getRange(start, 1, out.length, REVIEW_HEADERS.length);
      // Plain text so values like "06497" and "1,338.00" are stored exactly as shown.
      sh.getRange(start, 1, out.length, REVIEW_HEADERS.length - 1).setNumberFormat('@');
      rng.setValues(out);
      rng.setBackgrounds(out.map(r => REVIEW_HEADERS.map(() => REVIEW_COLORS[r[8]] || null)));
    }
    const count = status => lines.filter(l => l.status === status).length;
    return {
      match: count(REVIEW.MATCH), notMatch: count(REVIEW.NOT_MATCH), notInPhoto: count(REVIEW.NOT_IN_PHOTO),
      onlyInPhoto: count(REVIEW.ONLY_IN_PHOTO),
    };
  } finally {
    lock.releaseLock();
  }
}

/** All Review rows (display values), or [] if the tab doesn't exist yet. */
function readReview_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.REVIEW);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, REVIEW_HEADERS.length).getDisplayValues();
}

/** Any date-ish value → 'MM/dd/yy', the way logbook dates are written. */
function shortDate_(v) {
  const d = normDate_(v);
  return d ? d.slice(5, 7) + '/' + d.slice(8, 10) + '/' + d.slice(2, 4) : String(v);
}

function normText_(v) {
  return String(v).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}
