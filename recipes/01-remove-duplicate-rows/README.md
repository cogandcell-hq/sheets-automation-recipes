# Remove duplicate rows in Google Sheets (by email, ID or any column)

Keep one row per customer, order or email: a no-code formula view, plus a one-click script that cleans the sheet in place.

- **Setup time:** 5 minutes
- **You'll need:** A Google Sheet with a header row.
- **Full page with copy buttons:** https://cogandcell.com/recipes/remove-duplicate-rows/

## Option A: a de-duplicated view with a formula (non-destructive)

Make a new tab and put one of these in cell A1. Your original data is untouched; the new tab updates automatically.

**Exact duplicate rows only (every column identical)**

```
=UNIQUE(Sheet1!A:D)
```

**One row per email (column B), keeping the first one**

```
={Sheet1!A1:D1; FILTER(Sheet1!A2:D, Sheet1!B2:B<>"", COUNTIFS(Sheet1!B2:B, Sheet1!B2:B, ROW(Sheet1!B2:B), "<="&ROW(Sheet1!B2:B))=1)}
```

The second formula counts how many times each email has appeared *up to that row* and keeps only the first appearance. On very large sheets (50k+ rows) it can be slow; use Option B instead. Rows with an empty email are left out of this view.

## Option B: clean the sheet in place with a script

1. Make a copy of your sheet first (File → Make a copy). This script overwrites the tab.
2. Open **Extensions → Apps Script**, delete the sample code, and paste the script below.
3. Edit `DEDUPE_CONFIG`: the tab name, and `keyColumns` (e.g. `[2]` for column B, or `[2, 3]` to match on B and C together).
4. Click **Save**, then reload the spreadsheet. A new **Cleanup** menu appears.
5. Choose **Cleanup → Remove duplicate rows** and approve the permission prompt the first time.

**Code.gs** (also saved as [`dedupe.gs`](dedupe.gs)):

```javascript
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
```

## Good to know

- Matching ignores leading/trailing spaces and (by default) upper/lower case, so `Ann@X.com`  and `ann@x.com` count as the same.
- Rows whose key cells are completely empty are never merged together.
- The script writes values back, so formulas in the cleaned range become plain values. Keep formulas in a separate tab if you need them.
- Google Sheets also has **Data → Data cleanup → Remove duplicates** for one-off manual cleanups.

---
**Want it set up for you?** Get a free quote: https://cogandcell.com/#request (fixed price, see https://cogandcell.com/pricing)
