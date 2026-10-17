/**
 * Remove duplicate rows from a sheet, matching on one or more key columns.
 * Adds a "Cleanup" menu so you can run it from the spreadsheet.
 */
const DEDUPE_CONFIG = {
  sheetName: 'Sheet1',   // tab to clean
  headerRows: 1,         // rows at the top to leave alone
  keyColumns: [2],       // 1-based columns that define a duplicate (e.g. [2] = column B, [2, 3] = B+C)
  caseInsensitive: true, // treat "ANN@X.COM" and "ann@x.com" as the same
  keep: 'first'          // 'first' keeps the oldest row, 'last' keeps the newest
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cleanup')
    .addItem('Remove duplicate rows', 'removeDuplicateRows')
    .addToUi();
}

/** Pure function: returns the rows to keep. Safe to test outside Sheets. */
function dedupeRows(rows, keyColumns, caseInsensitive, keep) {
  const keyOf = (row) => keyColumns
    .map((c) => String(row[c - 1] === undefined ? '' : row[c - 1]).trim())
    .map((v) => (caseInsensitive ? v.toLowerCase() : v))
    .join('\u0001');
  const ordered = keep === 'last' ? rows.slice().reverse() : rows;
  const seen = new Set();
  const kept = [];
  ordered.forEach((row) => {
    const key = keyOf(row);
    const blank = key.replace(/\u0001/g, '') === '';
    if (blank || !seen.has(key)) { // never merge rows whose key is completely empty
      seen.add(key);
      kept.push(row);
    }
  });
  return keep === 'last' ? kept.reverse() : kept;
}

function removeDuplicateRows() {
  const cfg = DEDUPE_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) throw new Error('No tab named ' + cfg.sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= cfg.headerRows) return;

  const range = sheet.getRange(cfg.headerRows + 1, 1, lastRow - cfg.headerRows, lastCol);
  const rows = range.getValues();
  const kept = dedupeRows(rows, cfg.keyColumns, cfg.caseInsensitive, cfg.keep);
  const removed = rows.length - kept.length;
  if (removed === 0) {
    ss.toast('No duplicates found.');
    return;
  }
  range.clearContent();
  sheet.getRange(cfg.headerRows + 1, 1, kept.length, lastCol).setValues(kept);
  ss.toast('Removed ' + removed + ' duplicate row(s).');
}
