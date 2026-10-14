/**
 * Logbook Scanner: menu and the server functions the Scanner panel calls.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Scanner')
    .addItem('Open scanner', 'openScanner')
    .addSeparator()
    .addItem('Set up / refresh column matching', 'setupScanner')
    .addToUi();
}

function openScanner() {
  const html = HtmlService.createHtmlOutputFromFile('Panel').setWidth(1280).setHeight(780);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Logbook Scanner');
}

/**
 * Recent Form submissions that have at least one photo (newest first), the branch names (A–Z)
 * and whether the sheets write dates day first or month first.
 */
function getPanelData() {
  const s = getSettings_();
  const sh = getResponsesSheet_(s);
  const last = sh.getLastRow();
  const result = { submissions: [], branches: [], dateOrder: 'MDY', hidden: 0, recentDays: s.recentDays };
  if (last < 2) return result;
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const jc = colToNum_(s.journalCol), rc = colToNum_(s.receiptCol);
  const bc = branchColumn_(s, headers);
  const nc = s.nameCol ? colToNum_(s.nameCol) : findHeader_(headers, /name/i, /branch|file|photo|upload/i);
  const dc = s.dateCol ? colToNum_(s.dateCol) : findHeader_(headers, /date/i, /time ?stamp/i);
  const disp = sh.getRange(2, 1, last - 1, Math.max(jc, rc, bc, nc, dc, 2)).getDisplayValues();
  const stamps = sh.getRange(2, 1, last - 1, 1).getValues().map(r => r[0]);
  const logDates = dc ? sh.getRange(2, dc, last - 1, 1).getValues().map(r => r[0]) : null;
  result.dateOrder = detectDateOrder_(stamps, disp.map(r => r[0]));
  setDateOrder_(result.dateOrder);

  const cutoff = s.recentDays > 0 ? Date.now() - s.recentDays * 86400000 : 0;
  const done = {};
  readReview_().forEach(r => { done[r[0]] = true; });
  for (let i = disp.length - 1; i >= 0; i--) {
    const r = disp[i];
    const journal = fileIds_(r[jc - 1]), receipt = fileIds_(r[rc - 1]);
    if (!journal.length && !receipt.length) continue;
    if (cutoff && stamps[i] instanceof Date && stamps[i].getTime() < cutoff) { result.hidden++; continue; }
    const key = (i + 2) + '|' + r[0];
    result.submissions.push({
      key: key,
      row: i + 2,
      submitted: r[0],
      who: nc ? r[nc - 1] : r[1],
      logDate: dc && (logDates[i] || r[dc - 1]) ? shortDate_(logDates[i] || r[dc - 1]) : '',
      journal: journal,
      receipt: receipt,
      branch: bc ? r[bc - 1].trim() : '',
      done: !!done[key],
    });
  }
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
