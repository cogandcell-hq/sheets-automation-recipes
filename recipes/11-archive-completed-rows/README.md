# Automatically move rows to a 'Done' tab when the status changes

Keep your task, order or lead list short: rows marked Done move to an archive tab with the date, instantly.

- **Setup time:** 5 minutes
- **You'll need:** A tab with a status column (a dropdown works best).
- **Full page with copy buttons:** https://cogandcell.com/recipes/archive-completed-rows/

## Steps

1. Add a dropdown to your status column (Insert → Dropdown) with values like *Open, In progress, Done*.
2. **Extensions → Apps Script**, paste the script, and set the tab names and `statusColumn` (4 = column D).
3. Save. Change any status to *Done* and the row moves to the **Done** tab (created automatically).
4. Already have lots of done rows? Run `archiveAllDone` once to sweep them.

**Code.gs** (also saved as [`archive.gs`](archive.gs)):

```javascript
/**
 * Move a row to an archive tab as soon as its status is set to "Done".
 * Uses the simple onEdit trigger. Also includes archiveAllDone() to sweep existing rows.
 */
const ARCHIVE_CONFIG = {
  sourceSheet: 'Tasks',
  archiveSheet: 'Done',
  headerRows: 1,
  statusColumn: 4,                  // column D holds the status dropdown
  doneValues: ['done', 'complete'], // case-insensitive
  addArchivedAt: true               // append the date it was archived
};

function onEdit(e) {
  archiveOnEdit(e); // if you also use the timestamp recipe, call stampOnEdit(e) here too
}

function isDone(value) {
  return ARCHIVE_CONFIG.doneValues.indexOf(String(value).trim().toLowerCase()) !== -1;
}

function archiveTarget() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ARCHIVE_CONFIG.archiveSheet) || ss.insertSheet(ARCHIVE_CONFIG.archiveSheet);
}

function archiveOnEdit(e) {
  const cfg = ARCHIVE_CONFIG;
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() !== cfg.sourceSheet) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return; // single-cell edits only
  if (range.getColumn() !== cfg.statusColumn || range.getRow() <= cfg.headerRows) return;
  if (!isDone(range.getValue())) return;

  const row = range.getRow();
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  archiveTarget().appendRow(cfg.addArchivedAt ? values.concat([new Date()]) : values);
  sheet.deleteRow(row);
}

/** Pure function: returns 0-based indexes of rows whose status is done. */
function doneRowIndexes(rows, statusColumn) {
  return rows.reduce((acc, row, i) => (isDone(row[statusColumn - 1]) ? acc.concat([i]) : acc), []);
}

function archiveAllDone() {
  const cfg = ARCHIVE_CONFIG;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sourceSheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= cfg.headerRows) return;
  const rows = sheet.getRange(cfg.headerRows + 1, 1, lastRow - cfg.headerRows, sheet.getLastColumn()).getValues();
  const idx = doneRowIndexes(rows, cfg.statusColumn);
  const target = archiveTarget();
  const now = new Date();
  idx.forEach((i) => target.appendRow(cfg.addArchivedAt ? rows[i].concat([now]) : rows[i]));
  idx.slice().reverse().forEach((i) => sheet.deleteRow(i + cfg.headerRows + 1)); // bottom-up so row numbers stay valid
}
```

## Good to know

- Only single-cell edits trigger a move, so pasting a block of statuses won't shuffle rows unexpectedly; use `archiveAllDone` for bulk changes.
- Rows are deleted from the source tab after copying. Formulas in other tabs that point at specific rows may shift; reference whole columns instead.
- Combining with the timestamp recipe? Keep one `onEdit` that calls both `stampOnEdit(e)` and `archiveOnEdit(e)`.

---
**Want it set up for you?** Get a free quote: https://cogandcell.com/#request (fixed price, see https://cogandcell.com/pricing)
