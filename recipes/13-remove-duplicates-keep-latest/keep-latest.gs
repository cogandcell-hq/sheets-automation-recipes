/**
 * Remove duplicates but keep the LATEST row for each key, judged by a date column
 * (e.g. a form "Timestamp"), not by position in the sheet. Adds a "Cleanup" menu item.
 */
const LATEST_CONFIG = {
  sheetName: 'Form Responses 1', // tab to clean
  headerRows: 1,
  keyColumn: 2,                  // 1-based column that identifies a person/record (2 = B, e.g. email)
  dateColumn: 1,                 // 1-based column holding the date/time to compare (1 = A, "Timestamp")
  caseInsensitive: true
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cleanup')
    .addItem('Keep latest row per key', 'keepLatestPerKey')
    .addToUi();
}

/** Turns a cell value into a sortable number. Real dates win; blanks sort as oldest. */
function toTime(v) {
  if (v instanceof Date) return v.getTime();
  if (v === '' || v === null || v === undefined) return -Infinity;
  const t = new Date(v).getTime();
  return isNaN(t) ? -Infinity : t;
}

/** Pure function: one row per key (the newest), original row order preserved. Blank keys are kept. */
function latestRows(rows, keyColumn, dateColumn, caseInsensitive) {
  const best = new Map(); // key -> index of newest row so far
  rows.forEach((row, i) => {
    let key = String(row[keyColumn - 1] === undefined ? '' : row[keyColumn - 1]).trim();
    if (caseInsensitive) key = key.toLowerCase();
    if (key === '') return;
    const prev = best.get(key);
    // ">=" means that on a tie the lower row (added later) wins
    if (prev === undefined || toTime(row[dateColumn - 1]) >= toTime(rows[prev][dateColumn - 1])) best.set(key, i);
  });
  const keep = new Set(best.values());
  return rows.filter((row, i) => {
    const key = String(row[keyColumn - 1] === undefined ? '' : row[keyColumn - 1]).trim();
    return key === '' || keep.has(i);
  });
}

function keepLatestPerKey() {
  const cfg = LATEST_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('No tab named ' + cfg.sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= cfg.headerRows) return;
  const range = sheet.getRange(cfg.headerRows + 1, 1, lastRow - cfg.headerRows, lastCol);
  const rows = range.getValues();
  const kept = latestRows(rows, cfg.keyColumn, cfg.dateColumn, cfg.caseInsensitive);
  const removed = rows.length - kept.length;
  if (removed === 0) { ss.toast('No older duplicates found.'); return; }
  range.clearContent();
  sheet.getRange(cfg.headerRows + 1, 1, kept.length, lastCol).setValues(kept);
  ss.toast('Removed ' + removed + ' older duplicate row(s).');
}
