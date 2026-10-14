/**
 * Logbook Scanner: compares typed logbook rows with Spreadsheet 2, by date.
 *
 * For every date in "Journal Entries", the logbook rows for that date are
 * compared with all Spreadsheet 2 rows that have the same date:
 *   - amount fields: the totals for the day must agree within the tolerance
 *   - text fields (NAME, CI#): every value must appear on both sides
 *     (near-misses are flagged "Check" so typos can be told from real gaps)
 */

const RESULT_HEADERS = ['Date', 'Spreadsheet 2 column', 'Header', 'Logbook field',
  'Logbook value', 'Spreadsheet 2 value', 'Difference', 'Status', 'Note'];
const STATUS = { OK: '✅ Match', CHECK: '⚠️ Check', BAD: '❌ Mismatch', MISSING: '❌ Date missing' };
const STATUS_COLORS = { '✅': '#d9f2d9', '⚠️': '#fff2cc', '❌': '#f8d7da' };

/** Rebuilds the Comparison tab and returns its rows. */
function runComparison_() {
  const s = getSettings_();
  const mapping = readMapping_(false);
  const dateMap = mapping.filter(m => m.field === 'DATE')[0];
  if (!dateMap) throw new Error('In "' + SHEETS.CONFIG + '", set one Spreadsheet 2 column to DATE.');
  const others = mapping.filter(m => m !== dateMap);

  const book = {};
  readEntries_().forEach(e => {
    const d = e.values[0];
    if (d) (book[d] = book[d] || []).push(e.values);
  });

  const src = openSource_(s);
  const lastRow = src.getLastRow();
  const lastCol = Math.max.apply(null, [dateMap.col].concat(others.map(m => m.col)));
  const data = lastRow > s.headerRow
    ? src.getRange(s.headerRow + 1, 1, lastRow - s.headerRow, lastCol).getValues() : [];
  const ref = {};
  data.forEach(r => {
    const d = normDate_(r[dateMap.col - 1]);
    if (d) (ref[d] = ref[d] || []).push(r);
  });

  const out = [];
  Object.keys(book).sort().forEach(d => {
    const rows = book[d];
    const refRows = ref[d] || [];
    if (!refRows.length) {
      out.push([d, dateMap.letter, dateMap.header, 'DATE', rows.length + ' logbook row(s)', '', '',
        STATUS.MISSING, 'This date is not in Spreadsheet 2']);
      return;
    }
    others.forEach(m => {
      const bookVals = rows.map(r => bookValue_(r, m.field));
      const refVals = refRows.map(r => r[m.col - 1]);
      out.push([d, m.letter, m.header, m.field].concat(compareField_(m.field, bookVals, refVals, s.tolerance)));
    });
  });

  writeResults_(out);
  return out;
}

function bookValue_(row, field) {
  if (field === TOTAL_FIELD) {
    return toNumber_(row[FIELDS.indexOf('CASH')]) + toNumber_(row[FIELDS.indexOf('CARD')]);
  }
  return row[FIELDS.indexOf(field)];
}

/** → [logbook value, Spreadsheet 2 value, difference, status, note] */
function compareField_(field, bookVals, refVals, tolerance) {
  if (NUMERIC_FIELDS.indexOf(field) >= 0 || field === TOTAL_FIELD) {
    const a = round2_(bookVals.reduce((t, v) => t + toNumber_(v), 0));
    const b = round2_(refVals.reduce((t, v) => t + toNumber_(v), 0));
    const diff = round2_(a - b);
    const ok = Math.abs(diff) <= tolerance + 1e-9;
    return [a, b, diff, ok ? STATUS.OK : STATUS.BAD, ok ? '' : 'Day totals differ by ' + diff];
  }

  const bookSet = uniqueText_(bookVals);
  const refSet = uniqueText_(refVals);
  const refKeys = refSet.map(normText_);
  const notes = [];
  const used = {};
  let bad = false, close = false;
  bookSet.forEach(v => {
    const k = normText_(v);
    const exact = refKeys.indexOf(k);
    if (exact >= 0) { used[exact] = true; return; }
    let best = -1, bestScore = 0;
    refKeys.forEach((r, i) => {
      const score = similarity_(k, r);
      if (score > bestScore) { best = i; bestScore = score; }
    });
    if (bestScore >= 0.8) {
      used[best] = true;
      close = true;
      notes.push('"' + v + '" ≈ "' + refSet[best] + '"');
    } else {
      bad = true;
      notes.push('"' + v + '" not in Spreadsheet 2');
    }
  });
  refSet.forEach((v, i) => {
    if (!used[i]) { bad = true; notes.push('"' + v + '" not in logbook'); }
  });
  const status = bad ? STATUS.BAD : close ? STATUS.CHECK : STATUS.OK;
  return [bookSet.join(', '), refSet.join(', '), '', status, notes.join('; ')];
}

function writeResults_(rows) {
  const sh = ensureSheet_(SHEETS.RESULTS, RESULT_HEADERS);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, RESULT_HEADERS.length).clearContent().setBackground(null);
  if (!rows.length) return;
  sh.getRange(2, 1, rows.length, RESULT_HEADERS.length).setValues(rows);
  sh.getRange(2, 1, rows.length, RESULT_HEADERS.length)
    .setBackgrounds(rows.map(r => {
      const color = STATUS_COLORS[Object.keys(STATUS_COLORS).filter(k => String(r[7]).indexOf(k) === 0)[0]] || null;
      return RESULT_HEADERS.map(() => color);
    }));
}

function uniqueText_(vals) {
  const seen = {};
  return vals.map(v => String(v).trim()).filter(v => {
    const k = normText_(v);
    if (!k || seen[k]) return false;
    seen[k] = true;
    return true;
  });
}

function normText_(v) {
  return String(v).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

/** 1 = identical, 0 = nothing alike (Levenshtein-based). */
function similarity_(a, b) {
  if (!a.length || !b.length) return 0;
  let prev = [];
  for (let j = 0; j <= b.length; j++) prev[j] = j;
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur;
  }
  return 1 - prev[b.length] / Math.max(a.length, b.length);
}
