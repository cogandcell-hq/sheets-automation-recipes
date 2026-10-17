# Add an automatic timestamp when a row is edited

Record 'last updated' (or 'created at') automatically whenever someone changes a row. No trigger setup needed.

- **Setup time:** 5 minutes
- **You'll need:** A Google Sheet. Works for anyone who edits the sheet in the browser or app.
- **Full page with copy buttons:** https://cogandcell.com/recipes/timestamp-on-edit/

## Steps

1. Open **Extensions → Apps Script** and paste the script below.
2. Set `sheetName`, the columns to watch, and the column for the timestamp (keep the timestamp column out of `watchColumns`).
3. Want a “created at” stamp that never changes? Set `onlyIfEmpty: true`.
4. Save. That's it: `onEdit` is a *simple trigger*, so it runs automatically. Edit a watched cell to test it.

**Code.gs** (also saved as [`timestamp.gs`](timestamp.gs)):

```javascript
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
```

## Good to know

- Why not `=NOW()`? NOW recalculates constantly, so every row would show the current time. A script writes a fixed value.
- Pasting a block of cells stamps every row in the paste.
- Simple triggers only fire for edits made by people, not for changes made by formulas, other scripts, or Google Form submissions.
- Already have an `onEdit` function in this project? Keep one `onEdit` and call `stampOnEdit(e)` from it.

---
Need it set up for you? $49 fixed-price fix at https://cogandcell.com
