# Google Sheets: remove duplicates but keep the latest entry (by timestamp)

When someone submits a form twice or a record is updated, keep only the newest row per email or ID: a SORTN formula view, an Excel 365 version, and a tested in-place script.

- **Setup time:** 5 minutes
- **You'll need:** A sheet with a key column (e.g. email in B) and a date/time column (e.g. a form Timestamp in A).
- **Full page with copy buttons:** https://cogandcell.com/recipes/remove-duplicates-keep-latest/

## Option A: a live 'latest only' view with a formula

Put this in A2 of a new tab (copy the headers into row 1). It sorts newest-first by the timestamp in column A, then keeps the first row it sees for each value in column B. Your responses tab is untouched.

**Google Sheets: newest row per email (A = Timestamp, B = Email)**

```
=SORTN(SORT(FILTER('Form Responses 1'!A2:E, 'Form Responses 1'!B2:B<>""), 1, FALSE), 9^9, 2, 2, TRUE)
```

How it works: `SORT(..., 1, FALSE)` puts the newest timestamps on top; `SORTN` with tie mode `2` on column `2` returns only the first row for each distinct email, which is now the newest. `9^9` just means “no row limit”.

**Excel 365: same result with SORTBY + XMATCH**

```
=LET(d, SORTBY(A2:E500, A2:A500, -1), k, INDEX(d, , 2), FILTER(d, (XMATCH(k, k)=SEQUENCE(ROWS(k)))*(k<>"")))
```

## Option B: delete the older rows in place with a script

1. Make a copy of the sheet first. This script rewrites the tab.
2. **Extensions → Apps Script**, paste the script, and set `sheetName`, `keyColumn` and `dateColumn`.
3. Save, reload the sheet, then choose **Cleanup → Keep latest row per key**.

**Code.gs** (also saved as [`keep-latest.gs`](keep-latest.gs)):

```javascript
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
```

## Good to know

- The script compares real dates, so it works even if rows are out of order. If two rows have the identical timestamp, the lower one wins.
- SORTN compares keys as exact text, so `Ann@x.com` and `ann@x.com` are different. Wrap the key in a cleaned helper column (`=LOWER(TRIM(B2))`) if your data is messy; the script already ignores case and spaces.
- Want to keep the *first* entry instead? See [remove duplicate rows](../remove-duplicate-rows/).
- Don't run the in-place script on a live Form responses tab you never want altered; use the formula view instead.

---
**Want it set up for you?** Get a free quote: https://cogandcell.com/#request (fixed price, see https://cogandcell.com/pricing)
