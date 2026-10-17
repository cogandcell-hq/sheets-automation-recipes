/**
 * Email someone when a cell in a watched column changes (optionally only to a specific value,
 * e.g. Status -> "Approved"). Simple onEdit triggers can't send email, so this uses an
 * installable trigger: run installEditTrigger once.
 */
const CHANGE_CONFIG = {
  sheetName: 'Tasks',
  headerRows: 1,
  watchColumn: 4,             // 1-based column to watch (4 = D, e.g. Status)
  triggerValues: ['Approved'],// only these new values send email; [] = any change
  notifyTo: 'you@example.com',// who gets the email
  emailColumn: 0,             // or: 1-based column holding a per-row recipient (0 = use notifyTo)
  labelColumn: 1              // column used to name the row in the subject (1 = A)
};

function installEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers()
    .some((t) => t.getHandlerFunction() === 'emailOnChange');
  if (!exists) ScriptApp.newTrigger('emailOnChange').forSpreadsheet(ss).onEdit().create();
}

/** Pure function: should this edit send an email? */
function shouldNotify(sheetName, row, col, numRows, numCols, newValue, cfg) {
  if (sheetName !== cfg.sheetName) return false;
  if (numRows !== 1 || numCols !== 1) return false; // ignore pastes and multi-cell edits
  if (row <= cfg.headerRows || col !== cfg.watchColumn) return false;
  const v = String(newValue === undefined ? '' : newValue).trim().toLowerCase();
  if (v === '') return false;
  if (!cfg.triggerValues.length) return true;
  return cfg.triggerValues.some((t) => String(t).trim().toLowerCase() === v);
}

function emailOnChange(e) {
  const cfg = CHANGE_CONFIG;
  const range = e.range;
  const sheet = range.getSheet();
  const newValue = e.value !== undefined ? e.value : range.getValue();
  if (!shouldNotify(sheet.getName(), range.getRow(), range.getColumn(),
      range.getNumRows(), range.getNumColumns(), newValue, cfg)) return;
  const rowValues = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).getValues()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const to = cfg.emailColumn ? String(rowValues[cfg.emailColumn - 1]).trim() : cfg.notifyTo;
  if (!to) return;
  const label = rowValues[cfg.labelColumn - 1] || ('row ' + range.getRow());
  const details = headers.map((h, i) => (h || 'Column ' + (i + 1)) + ': ' + rowValues[i]).join('\n');
  MailApp.sendEmail(to,
    label + ' changed to "' + newValue + '"',
    'Row ' + range.getRow() + ' on "' + sheet.getName() + '" was updated' +
    (e.oldValue !== undefined ? ' (was "' + e.oldValue + '")' : '') + '.\n\n' + details +
    '\n\nOpen the sheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}
