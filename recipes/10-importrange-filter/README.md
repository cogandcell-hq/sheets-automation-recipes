# Pull and filter data from another spreadsheet (IMPORTRANGE + QUERY)

Show only the rows you need from another file, combine several files into one, or take a stable scheduled snapshot.

- **Setup time:** 5 minutes
- **You'll need:** View access to the source spreadsheet. Examples assume an `Orders` tab with Status in column E.
- **Full page with copy buttons:** https://cogandcell.com/recipes/importrange-filter/

## Formulas

**Import a range (first time: click the cell and press Allow access)**

```
=IMPORTRANGE("https://docs.google.com/spreadsheets/d/SOURCE_ID/edit", "Orders!A:F")
```

**Only open orders, selected columns (note Col1, Col2... not letters)**

```
=QUERY(IMPORTRANGE("https://docs.google.com/spreadsheets/d/SOURCE_ID/edit", "Orders!A:F"), "select Col1, Col3, Col6 where Col5 = 'Open'", 1)
```

**Same thing with FILTER (import once, filter by column 5)**

```
=LET(d, IMPORTRANGE("https://docs.google.com/spreadsheets/d/SOURCE_ID/edit", "Orders!A2:F"), FILTER(d, CHOOSECOLS(d, 5)="Open"))
```

**Stack two files with the same columns and drop blank rows**

```
=QUERY({IMPORTRANGE("URL_1", "Orders!A2:F"); IMPORTRANGE("URL_2", "Orders!A2:F")}, "select * where Col1 is not null", 0)
```

## Optional: a scheduled snapshot instead of a live link

Lots of IMPORTRANGEs make a sheet slow and can show “Loading…”. This script copies plain values every hour instead.

1. **Extensions → Apps Script**, paste, set the source ID and range.
2. Run `syncSnapshot` once (approve access), then `installHourlySync`.

**Code.gs** (also saved as [`sync-snapshot.gs`](sync-snapshot.gs)):

```javascript
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
```

## Good to know

- Both ranges in a stacked `{ ; }` array must have the same number of columns, or you'll get an error.
- Some locales use `\` instead of `,` or `;` inside array braces. If `{A;B}` errors, check File → Settings → Locale.
- Anyone who can see your sheet can see the imported data, even without access to the source file. Import only what's needed.

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
