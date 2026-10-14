/**
 * Logbook Scanner: settings, sheet setup and helpers.
 *
 * All settings live in the "Scanner Config" tab so nobody has to edit code.
 */

const SHEETS = {
  CONFIG: 'Scanner Config',
  REVIEW: 'Review',
};

// How each Spreadsheet 2 column is used. Any other value left over from an older
// version (CASH, NAME, ...) counts as SHOW.
const DATE = 'DATE';
const BRANCH = 'BRANCH';
const SHOW = 'SHOW';
const IGNORE = '(ignore)';
const MAP_CHOICES = [DATE, BRANCH, SHOW, IGNORE];

const DEFAULTS = [
  ['Spreadsheet 2 (URL or ID)', 'https://docs.google.com/spreadsheets/d/17ly2V1qu4DmwuT3efKzALchxsyJ5adj8AG8u2lBBe60/edit?gid=147027864'],
  ['Spreadsheet 2 tab name (blank = use the gid in the URL)', ''],
  ['Spreadsheet 2 header row', 1],
  ['Spreadsheet 2 columns to compare', 'C,D,X:AF'],
  ['Spreadsheet 1: Form responses (URL or ID; blank = this spreadsheet)', 'https://docs.google.com/spreadsheets/d/1_jx8q2Y-fKi8T04ZCsJfORbCOPev3lh_XjeG07QL8WY/edit?gid=921021716'],
  ['Spreadsheet 1 tab name (blank = use the gid in the URL)', ''],
  ['Sales Journal photo column', 'F'],
  ['Receipt / extra page photo column', 'G'],
  ['Branch column in Spreadsheet 1 (blank = find the "Branch" header)', ''],
  ['Count identical Spreadsheet 2 rows once (yes/no)', 'yes'],
];
const MAP_TITLE = 'Column matching: how is each Spreadsheet 2 column used? (DATE and BRANCH filter the records; SHOW columns are checked against the photo)';

/** Settings are looked up by label, so tabs set up by an older version keep working. */
function getSettings_() {
  const sh = getConfigSheet_();
  const byLabel = {};
  if (sh.getLastRow()) sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(r => { byLabel[String(r[0]).trim()] = r[1]; });
  const v = DEFAULTS.map(d => d[0] in byLabel ? byLabel[d[0]] : d[1]);
  return {
    source: String(v[0]).trim(),
    sourceTab: String(v[1]).trim(),
    headerRow: Number(v[2]) || 1,
    columns: String(v[3]),
    responses: String(v[4]).trim(),
    responsesTab: String(v[5]).trim(),
    journalCol: String(v[6]).trim() || 'F',
    receiptCol: String(v[7]).trim() || 'G',
    branchCol: String(v[8]).trim(),
    dedupe: !/^(no|n|false|0)$/i.test(String(v[9]).trim()),
  };
}

function getConfigSheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG);
  if (!sh) throw new Error('Run "Scanner → Set up / refresh column matching" first.');
  return sh;
}

/** Creates the Config, Journal Entries and Comparison tabs if missing, then reads Spreadsheet 2's headers. */
function setupScanner() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEETS.CONFIG);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.CONFIG);
    sh.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
    sh.getRange(2, 1, DEFAULTS.length, 2).setValues(DEFAULTS);
    sh.setColumnWidth(1, 340);
    sh.setColumnWidth(2, 420);
    sh.setColumnWidth(3, 220);
  } else {
    addMissingSettings_(sh);
  }
  ensureSheet_(SHEETS.REVIEW, REVIEW_HEADERS);
  refreshMapping_();
  SpreadsheetApp.getUi().alert(
    'Setup done.\n\nOpen the "' + SHEETS.CONFIG + '" tab and check the column-matching table: ' +
    'one column must be DATE, the branch column should be BRANCH, and set every column ' +
    'you want to check against the photo to SHOW.');
}

/** Adds settings introduced after the Config tab was created, just below the existing ones. */
function addMissingSettings_(sh) {
  const labels = () => sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  DEFAULTS.forEach((d, i) => {
    const have = labels();
    if (have.indexOf(d[0]) >= 0) return;
    const before = DEFAULTS.slice(0, i).map(x => have.indexOf(x[0])).filter(r => r >= 0);
    const after = before.length ? Math.max.apply(null, before) + 1 : 1;
    sh.insertRowAfter(after);
    sh.getRange(after + 1, 1, 1, 2).setValues([d]);
  });
}

/** Row of the column-matching title in the Config tab, or 0. */
function mapTitleRow_(sh) {
  const n = sh.getLastRow();
  if (!n) return 0;
  const col = sh.getRange(1, 1, n, 1).getValues();
  for (let i = 0; i < n; i++) if (/^Column matching/.test(String(col[i][0]))) return i + 1;
  return 0;
}

/** Creates the tab if missing; on an existing tab, refreshes the header row (columns are only ever appended). */
function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  return sh;
}

/** Writes the column-matching table, keeping choices already made and guessing new ones from the headers. */
function refreshMapping_() {
  const sh = getConfigSheet_();
  const s = getSettings_();
  const src = openSource_(s);
  const letters = expandColumns_(s.columns);
  const lastCol = Math.max.apply(null, letters.map(colToNum_));
  const headers = src.getRange(s.headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const previous = {};
  readMapping_(true).forEach(m => { previous[m.letter] = m.field; });

  const rows = letters.map(l => {
    const header = headers[colToNum_(l) - 1];
    const prev = previous[l];
    return [l, header, prev ? (MAP_CHOICES.indexOf(prev) >= 0 ? prev : SHOW) : guessField_(header)];
  });

  const titleRow = mapTitleRow_(sh) || DEFAULTS.length + 3;
  const firstRow = titleRow + 2;
  const old = sh.getLastRow() - firstRow + 1;
  if (old > 0) sh.getRange(firstRow, 1, old, 3).clearContent().clearDataValidations();
  sh.getRange(titleRow, 1).setValue(MAP_TITLE).setFontWeight('bold');
  sh.getRange(titleRow + 1, 1, 1, 3)
    .setValues([['Spreadsheet 2 column', 'Header in Spreadsheet 2', 'Use as']]).setFontWeight('bold');
  sh.getRange(firstRow, 1, rows.length, 3).setValues(rows);
  sh.getRange(firstRow, 3, rows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MAP_CHOICES, true).build());
}

function guessField_(header) {
  const h = String(header).toLowerCase();
  if (/date/.test(h)) return DATE;
  if (/branch|store|outlet/.test(h)) return BRANCH;
  if (/deposit|receivable|cheque|transfer|categor|type|other/.test(h)) return IGNORE;
  if (/name|customer|supplier|\bci\b|ci#|ci no|invoice|inv|or no|receipt|cash|card|credit|debit|vat|tin|total|gross|amount/.test(h)) return SHOW;
  return IGNORE;
}

/** Returns [{letter, col, header, field}] from the matching table. */
function readMapping_(includeIgnored) {
  const sh = getConfigSheet_();
  const titleRow = mapTitleRow_(sh);
  if (!titleRow) return [];
  const n = sh.getLastRow() - titleRow - 1;
  if (n <= 0) return [];
  return sh.getRange(titleRow + 2, 1, n, 3).getValues()
    .filter(r => r[0] && (includeIgnored || (r[2] && r[2] !== IGNORE)))
    .map(r => ({
      letter: String(r[0]).trim(), col: colToNum_(String(r[0]).trim()), header: String(r[1]),
      field: MAP_CHOICES.indexOf(r[2]) >= 0 || !r[2] ? r[2] : SHOW,
    }));
}

/** The reference tab in Spreadsheet 2. */
function openSource_(s) {
  return openTab_(s.source, s.sourceTab, 'Spreadsheet 2');
}

/** The tab where the Google Form writes its responses (Spreadsheet 1). */
function getResponsesSheet_(s) {
  if (s.responses) return openTab_(s.responses, s.responsesTab, 'Spreadsheet 1');
  const ss = SpreadsheetApp.getActive();
  if (s.responsesTab) return openTab_(ss.getId(), s.responsesTab, 'this spreadsheet');
  const sheets = ss.getSheets();
  return sheets.filter(t => t.getFormUrl())[0] ||
    sheets.filter(t => /^form responses/i.test(t.getName()))[0] ||
    sheets[0];
}

/** Opens a tab by name, else by the gid in the URL, else the first tab. */
function openTab_(urlOrId, tabName, label) {
  const idMatch = urlOrId.match(/\/d\/([-\w]{25,})/) || urlOrId.match(/^([-\w]{25,})$/);
  if (!idMatch) throw new Error('"' + label + '" in ' + SHEETS.CONFIG + ' is not a valid sheet URL or ID.');
  const ss = SpreadsheetApp.openById(idMatch[1]);
  if (tabName) {
    const tab = ss.getSheetByName(tabName);
    if (!tab) throw new Error('Tab "' + tabName + '" not found in ' + label + '.');
    return tab;
  }
  const gid = urlOrId.match(/gid=(\d+)/);
  if (gid) {
    const tab = ss.getSheets().filter(t => String(t.getSheetId()) === gid[1])[0];
    if (tab) return tab;
  }
  return ss.getSheets()[0];
}

// ---- small helpers ----

function colToNum_(letter) {
  return String(letter).toUpperCase().split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
}

function numToCol_(n) {
  let s = '';
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s;
  return s;
}

/** Spreadsheet 1 column holding the branch (1-based), or 0 if there is none. */
function branchColumn_(s, headers) {
  if (s.branchCol) return colToNum_(s.branchCol);
  return headers.findIndex(h => /branch/i.test(h)) + 1;
}

/** Unique, trimmed branch names, A–Z (case-insensitive, "Branch 2" before "Branch 10"). */
function sortBranches_(names) {
  const seen = {};
  return names.map(n => String(n).trim()).filter(n => {
    const k = n.toLowerCase();
    if (!n || seen[k]) return false;
    seen[k] = true;
    return true;
  }).sort((a, b) => a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true }));
}

/** 'C,D,X:AF' → ['C','D','X','Y',...,'AF'] */
function expandColumns_(spec) {
  const out = [];
  String(spec).toUpperCase().split(',').map(p => p.trim()).filter(String).forEach(p => {
    const [a, b] = p.split(/[:\-]/).map(x => x.trim());
    for (let i = colToNum_(a); i <= colToNum_(b || a); i++) out.push(numToCol_(i));
  });
  return out;
}

/** Any date-ish value → 'yyyy-MM-dd' ('' if not a date). Text dates are read month-first. */
function normDate_(v) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const t = String(v).trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + pad2_(m[1]) + '-' + pad2_(m[2]);
  return '';
}

function pad2_(x) { return ('0' + x).slice(-2); }
