/**
 * Logbook Scanner: settings, sheet setup and helpers.
 *
 * All settings live in the "Scanner Config" tab so nobody has to edit code.
 */

const SHEETS = {
  CONFIG: 'Scanner Config',
  ENTRIES: 'Journal Entries',
  RESULTS: 'Comparison',
};

// Columns of the handwritten Sales Journal, in logbook order.
const FIELDS = ['DATE', 'NAME', 'CI#', 'CASH', 'CARD', 'VAT SALE', '12% VAT'];
const NUMERIC_FIELDS = ['CASH', 'CARD', 'VAT SALE', '12% VAT'];
const TOTAL_FIELD = 'TOTAL (CASH+CARD)';
const IGNORE = '(ignore)';
const MAP_CHOICES = FIELDS.concat([TOTAL_FIELD, IGNORE]);

const DEFAULTS = [
  ['Spreadsheet 2 (URL or ID)', 'https://docs.google.com/spreadsheets/d/17ly2V1qu4DmwuT3efKzALchxsyJ5adj8AG8u2lBBe60/edit?gid=147027864'],
  ['Spreadsheet 2 tab name (blank = use the gid in the URL)', ''],
  ['Spreadsheet 2 header row', 1],
  ['Spreadsheet 2 columns to compare', 'C,D,X:AF'],
  ['Spreadsheet 1: Form responses (URL or ID; blank = this spreadsheet)', 'https://docs.google.com/spreadsheets/d/1_jx8q2Y-fKi8T04ZCsJfORbCOPev3lh_XjeG07QL8WY/edit?gid=921021716'],
  ['Spreadsheet 1 tab name (blank = use the gid in the URL)', ''],
  ['Sales Journal photo column', 'F'],
  ['Receipt / extra page photo column', 'G'],
  ['Tolerance for amounts (±)', 1],
  ['VAT rate', 0.12],
];
const MAP_TITLE_ROW = DEFAULTS.length + 3;
const MAP_FIRST_ROW = MAP_TITLE_ROW + 2;

const ENTRY_HEADERS = ['Submission', 'Response row'].concat(FIELDS, ['Row check', 'Entered by', 'Entered at']);

function getSettings_() {
  const v = getConfigSheet_().getRange(2, 2, DEFAULTS.length, 1).getValues().map(r => r[0]);
  return {
    source: String(v[0]).trim(),
    sourceTab: String(v[1]).trim(),
    headerRow: Number(v[2]) || 1,
    columns: String(v[3]),
    responses: String(v[4]).trim(),
    responsesTab: String(v[5]).trim(),
    journalCol: String(v[6]).trim() || 'F',
    receiptCol: String(v[7]).trim() || 'G',
    tolerance: Number(v[8]) || 0,
    vatRate: Number(v[9]) || 0.12,
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
  }
  ensureSheet_(SHEETS.ENTRIES, ENTRY_HEADERS);
  ensureSheet_(SHEETS.RESULTS, RESULT_HEADERS);
  refreshMapping_();
  SpreadsheetApp.getUi().alert(
    'Setup done.\n\nOpen the "' + SHEETS.CONFIG + '" tab and check the column-matching table: ' +
    'each Spreadsheet 2 column needs the logbook field it holds. One column must be DATE.');
}

function ensureSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
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
    return [l, header, previous[l] || guessField_(header)];
  });

  const old = sh.getLastRow() - MAP_FIRST_ROW + 1;
  if (old > 0) sh.getRange(MAP_FIRST_ROW, 1, old, 3).clearContent().clearDataValidations();
  sh.getRange(MAP_TITLE_ROW, 1).setValue('Column matching: which logbook field is in each Spreadsheet 2 column?').setFontWeight('bold');
  sh.getRange(MAP_TITLE_ROW + 1, 1, 1, 3)
    .setValues([['Spreadsheet 2 column', 'Header in Spreadsheet 2', 'Logbook field']]).setFontWeight('bold');
  sh.getRange(MAP_FIRST_ROW, 1, rows.length, 3).setValues(rows);
  sh.getRange(MAP_FIRST_ROW, 3, rows.length, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(MAP_CHOICES, true).build());
}

function guessField_(header) {
  const h = String(header).toLowerCase();
  if (/date/.test(h)) return 'DATE';
  if (/name|customer/.test(h)) return 'NAME';
  if (/\bci\b|ci#|ci no|invoice|or no|receipt/.test(h)) return 'CI#';
  if (/cash/.test(h)) return 'CASH';
  if (/card|credit|debit/.test(h)) return 'CARD';
  if (/vat/.test(h) && /sale|able/.test(h)) return 'VAT SALE';
  if (/vat/.test(h)) return '12% VAT';
  if (/total|gross|amount/.test(h)) return TOTAL_FIELD;
  return IGNORE;
}

/** Returns [{letter, col, header, field}] from the matching table. */
function readMapping_(includeIgnored) {
  const sh = getConfigSheet_();
  const n = sh.getLastRow() - MAP_FIRST_ROW + 1;
  if (n <= 0) return [];
  return sh.getRange(MAP_FIRST_ROW, 1, n, 3).getValues()
    .filter(r => r[0] && (includeIgnored || (r[2] && r[2] !== IGNORE)))
    .map(r => ({ letter: String(r[0]).trim(), col: colToNum_(String(r[0]).trim()), header: String(r[1]), field: r[2] }));
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

/** '₱1,234.50' → 1234.5; blank → 0. */
function toNumber_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}
