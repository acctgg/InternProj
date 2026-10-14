/**
 * Logbook Scanner: menu and the server functions the Scanner panel calls.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scanner')
    .addItem('Open scanner', 'openScanner')
    .addItem('Re-run comparison', 'runComparisonFromMenu')
    .addSeparator()
    .addItem('Set up / refresh column matching', 'setupScanner')
    .addToUi();
}

function openScanner() {
  const html = HtmlService.createHtmlOutputFromFile('Panel').setWidth(1280).setHeight(780);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Logbook Scanner');
}

function runComparisonFromMenu() {
  const n = runComparison_().length;
  SpreadsheetApp.getActive().getSheetByName(SHEETS.RESULTS).activate();
  SpreadsheetApp.getActive().toast(n + ' comparison rows written.', 'Scanner');
}

/** Form submissions that have at least one photo (newest first), plus the settings the panel needs. */
function getPanelData() {
  const s = getSettings_();
  const sh = getResponsesSheet_(s);
  const last = sh.getLastRow();
  const result = { vatRate: s.vatRate, tolerance: s.tolerance, submissions: [], branches: [] };
  if (last < 2) return result;
  const jc = colToNum_(s.journalCol), rc = colToNum_(s.receiptCol);
  const bc = branchColumn_(s, sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]);
  const vals = sh.getRange(2, 1, last - 1, Math.max(jc, rc, bc, 2)).getDisplayValues();
  const done = {};
  readEntries_().forEach(e => { done[e.submission] = true; });
  result.submissions = vals.map((r, i) => {
    const key = (i + 2) + '|' + r[0];
    return {
      key: key,
      row: i + 2,
      submitted: r[0],
      who: r[1],
      journal: fileIds_(r[jc - 1]),
      receipt: fileIds_(r[rc - 1]),
      branch: bc ? r[bc - 1].trim() : '',
      done: !!done[key],
    };
  }).filter(x => x.journal.length || x.receipt.length).reverse();
  result.branches = sortBranches_(result.submissions.map(x => x.branch));
  return result;
}

/** Drive file IDs from a Form upload cell ("https://drive.google.com/open?id=..., ..."). */
function fileIds_(cell) {
  return String(cell || '').match(/[-\w]{25,}/g) || [];
}

// Long side of the preview sent to the panel. Full-size phone photos are too big for
// google.script.run (it fails with "NetworkError: Connection failure due to HTTP 503").
const PREVIEW_SIZE = 1600;
const MAX_RAW_BYTES = 3e6;

/**
 * Returns a photo as a data URL the panel can draw: a resized copy from Drive's
 * thumbnail renderer (also handles HEIC and PDF), or the original if it is small.
 */
function getImage(fileId) {
  const file = DriveApp.getFileById(fileId);
  const info = { name: file.getName(), url: file.getUrl(), dataUrl: null };
  try {
    const auth = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    const meta = JSON.parse(UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=thumbnailLink', auth).getContentText());
    if (meta.thumbnailLink) {
      const res = UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s' + PREVIEW_SIZE), auth);
      if (res.getResponseCode() === 200) {
        const thumb = res.getBlob();
        info.dataUrl = 'data:' + thumb.getContentType() + ';base64,' + Utilities.base64Encode(thumb.getBytes());
        return info;
      }
    }
  } catch (err) {
    // fall through to the original file
  }
  const type = file.getMimeType();
  if (/^image\/(jpeg|png|webp|gif)$/.test(type) && file.getSize() <= MAX_RAW_BYTES) {
    info.dataUrl = 'data:' + type + ';base64,' + Utilities.base64Encode(file.getBlob().getBytes());
  } else {
    info.error = 'Preview not ready yet (Drive may still be processing a new upload). ' +
      'Try again in a minute, or use "Open in Drive".';
  }
  return info;
}

/** Rows already typed in for one submission. */
function getEntries(key) {
  return readEntries_().filter(e => e.submission === key).map(e => e.values);
}

/**
 * Replaces the typed rows for one submission, re-runs the comparison and
 * returns the comparison rows for the dates on this page.
 * unchecked[i] is true when row i still has values auto-filled from Spreadsheet 2
 * that nobody confirmed against the photo; that is recorded in "Row check".
 */
function saveEntries(key, responseRow, rows, branch, unchecked) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const s = getSettings_();
    const who = Session.getActiveUser().getEmail();
    const now = new Date();
    const out = rows
      .map((r, n) => ({ r: r, unchecked: !!(unchecked && unchecked[n]) }))
      .filter(x => x.r.some(v => String(v).trim() !== ''))
      .map(x => {
        const values = FIELDS.map((f, i) => NUMERIC_FIELDS.indexOf(f) >= 0 && x.r[i] !== '' ? toNumber_(x.r[i]) : x.r[i]);
        const check = rowCheck_(values, s) + (x.unchecked ? ' · filled from Spreadsheet 2, not checked' : '');
        return [key, responseRow].concat(values, [check, who, now, branch || '']);
      });
    if (out.some(r => !normDate_(r[2]))) throw new Error('Every row needs a DATE.');

    const sh = ensureSheet_(SHEETS.ENTRIES, ENTRY_HEADERS);
    const last = sh.getLastRow();
    if (last > 1) {
      const keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
      for (let i = keys.length - 1; i >= 0; i--) if (keys[i][0] === key) sh.deleteRow(i + 2);
    }
    if (out.length) sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    SpreadsheetApp.flush();
    const dates = {};
    out.forEach(r => { dates[normDate_(r[2])] = true; });
    return runComparison_().filter(r => dates[r[0]]);
  } finally {
    lock.releaseLock();
  }
}

/** Arithmetic check on one logbook row: VAT ≈ rate × VAT SALE, and CASH + CARD ≈ VAT SALE + VAT. */
function rowCheck_(values, s) {
  const get = f => values[FIELDS.indexOf(f)];
  const vatSale = get('VAT SALE'), vat = get('12% VAT');
  if (vatSale === '' || vatSale === null) return 'OK';
  const problems = [];
  const vatExpected = toNumber_(vatSale) * s.vatRate;
  if (vat !== '' && Math.abs(toNumber_(vat) - vatExpected) > s.tolerance) {
    problems.push('VAT should be ≈ ' + vatExpected.toFixed(2));
  }
  const paid = toNumber_(get('CASH')) + toNumber_(get('CARD'));
  const gross = toNumber_(vatSale) + toNumber_(vat);
  if (paid && Math.abs(paid - gross) > s.tolerance) {
    problems.push('CASH+CARD (' + paid.toFixed(2) + ') ≠ VAT SALE+VAT (' + gross.toFixed(2) + ')');
  }
  return problems.length ? '⚠ ' + problems.join('; ') : 'OK';
}

/** [{submission, branch, values:[DATE, NAME, ...]}] from the Journal Entries tab. */
function readEntries_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.ENTRIES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ENTRY_HEADERS.length).getValues().map(r => ({
    submission: String(r[0]),
    branch: String(r[ENTRY_HEADERS.length - 1]).trim(),
    values: r.slice(2, 2 + FIELDS.length).map((v, i) => i === 0 ? normDate_(v) : v),
  }));
}
