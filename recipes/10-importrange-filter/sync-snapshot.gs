/**
 * Optional: copy a range from another spreadsheet as plain values on a schedule.
 * Use this instead of IMPORTRANGE when you want a stable snapshot (faster, no "Loading...").
 */
const SYNC_CONFIG = {
  sourceId: 'PASTE_SOURCE_SPREADSHEET_ID', // the long id in the source sheet's URL
  sourceRange: 'Orders!A1:F',
  targetSheet: 'Orders snapshot'
};

/** Pure function: drop fully empty trailing rows that open-ended ranges return. */
function trimEmptyRows(values) {
  let end = values.length;
  while (end > 0 && values[end - 1].every((v) => v === '' || v === null)) end--;
  return values.slice(0, end);
}

function syncSnapshot() {
  const cfg = SYNC_CONFIG;
  const values = trimEmptyRows(SpreadsheetApp.openById(cfg.sourceId).getRange(cfg.sourceRange).getValues());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const target = ss.getSheetByName(cfg.targetSheet) || ss.insertSheet(cfg.targetSheet);
  target.clearContents();
  if (values.length) target.getRange(1, 1, values.length, values[0].length).setValues(values);
}

function installHourlySync() {
  ScriptApp.newTrigger('syncSnapshot').timeBased().everyHours(1).create();
}
