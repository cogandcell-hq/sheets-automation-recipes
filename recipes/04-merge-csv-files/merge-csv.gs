/**
 * Merge every CSV in a Google Drive folder into one tab.
 * Columns are matched by header name, so files with different column orders still line up.
 */
const MERGE_CONFIG = {
  folderId: 'PASTE_FOLDER_ID_HERE', // the part after /folders/ in the folder URL
  targetSheet: 'Merged',
  addSourceColumn: true             // adds a "Source file" column so you can trace each row
};

/** Pure function: merge [{name, rows}] tables by header. First row of each table is its header. */
function mergeTables(tables, addSourceColumn) {
  const headers = [];
  const index = {};
  tables.forEach((t) => {
    (t.rows[0] || []).forEach((h) => {
      const key = String(h).replace(/^\uFEFF/, '').trim();
      if (key && !(key in index)) { index[key] = headers.length; headers.push(key); }
    });
  });
  const out = [];
  tables.forEach((t) => {
    const head = (t.rows[0] || []).map((h) => String(h).replace(/^\uFEFF/, '').trim());
    t.rows.slice(1).forEach((row) => {
      if (row.every((v) => String(v).trim() === '')) return; // skip blank lines
      const merged = new Array(headers.length).fill('');
      head.forEach((h, i) => { if (h in index) merged[index[h]] = row[i] === undefined ? '' : row[i]; });
      out.push(addSourceColumn ? merged.concat([t.name]) : merged);
    });
  });
  return [addSourceColumn ? headers.concat(['Source file']) : headers].concat(out);
}

function mergeCsvsFromFolder() {
  const cfg = MERGE_CONFIG;
  const folder = DriveApp.getFolderById(cfg.folderId);
  const files = folder.getFilesByType(MimeType.CSV);
  const tables = [];
  while (files.hasNext()) {
    const file = files.next();
    const text = file.getBlob().getDataAsString('UTF-8');
    tables.push({ name: file.getName(), rows: Utilities.parseCsv(text) });
  }
  if (tables.length === 0) throw new Error('No CSV files found in that folder.');
  tables.sort((a, b) => a.name.localeCompare(b.name));

  const merged = mergeTables(tables, cfg.addSourceColumn);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(cfg.targetSheet) || ss.insertSheet(cfg.targetSheet);
  sheet.clearContents();
  sheet.getRange(1, 1, merged.length, merged[0].length).setValues(merged);
  sheet.setFrozenRows(1);
  ss.toast('Merged ' + tables.length + ' file(s), ' + (merged.length - 1) + ' row(s).');
}
