/**
 * Write a timestamp in a column whenever someone edits a watched column on that row.
 * Uses the simple onEdit trigger: nothing to install, just save the script.
 */
const STAMP_CONFIG = {
  sheetName: 'Tasks',      // tab to watch
  headerRows: 1,
  watchColumns: [1, 2, 3], // edits in A, B or C...
  stampColumn: 5,          // ...write the time into E (keep it out of watchColumns)
  onlyIfEmpty: false,      // true = "created at" (stamp once), false = "last updated"
  numberFormat: 'yyyy-mm-dd hh:mm:ss'
};

function onEdit(e) {
  stampOnEdit(e); // if you also use the "archive done rows" recipe, call archiveOnEdit(e) here too
}

function stampOnEdit(e) {
  const cfg = STAMP_CONFIG;
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== cfg.sheetName) return;

  const firstCol = range.getColumn();
  const lastCol = range.getLastColumn();
  const touchesWatched = cfg.watchColumns.some((c) => c >= firstCol && c <= lastCol);
  if (!touchesWatched) return;

  const startRow = Math.max(range.getRow(), cfg.headerRows + 1);
  const endRow = range.getLastRow(); // pastes can cover many rows
  if (startRow > endRow) return;

  const count = endRow - startRow + 1;
  const stampRange = sheet.getRange(startRow, cfg.stampColumn, count, 1);
  const now = new Date();
  const current = stampRange.getValues();
  const next = current.map(([value]) => [cfg.onlyIfEmpty && value !== '' ? value : now]);
  stampRange.setValues(next).setNumberFormat(cfg.numberFormat);
}
