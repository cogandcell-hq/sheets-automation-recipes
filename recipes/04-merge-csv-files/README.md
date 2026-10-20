# Merge many CSV files into one sheet (even when the columns differ)

Drop exports from your shop, bank or CRM into one Drive folder and combine them in one click. Columns are matched by header name.

- **Setup time:** 10 minutes
- **You'll need:** A Google Drive folder of .csv files with header rows. (Or: Python 3 on your computer.)
- **Full page with copy buttons:** https://cogandcell.com/recipes/merge-csv-files/

## Option A: in Google Sheets

1. Put all the CSV files in one Google Drive folder. Open the folder and copy the ID from the URL (the part after `/folders/`).
2. In the spreadsheet where you want the result, open **Extensions → Apps Script** and paste the script.
3. Paste the folder ID into `folderId`.
4. Select `mergeCsvsFromFolder` and click **Run**. A **Merged** tab is created (or refreshed) with every row, plus a “Source file” column.

**Code.gs** (also saved as [`merge-csv.gs`](merge-csv.gs)):

```javascript
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
```

## Option B: on your computer with Python (no installs)

Save this as `merge_csvs.py`, then run `python3 merge_csvs.py ./exports merged.csv`. It uses only the standard library and handles the invisible “BOM” character Excel adds to CSVs.

**merge_csvs.py** (also saved as [`merge_csvs.py`](merge_csvs.py)):

```python
#!/usr/bin/env python3
"""Merge every CSV in a folder into one file, matching columns by header name.

Usage:  python3 merge_csvs.py ./exports merged.csv
Standard library only, no installs needed.
"""
import csv
import sys
from pathlib import Path


def merge(folder: Path, out_path: Path) -> int:
    files = sorted(p for p in folder.glob("*.csv") if p.resolve() != out_path.resolve())
    headers, rows = [], []
    for path in files:
        with path.open(newline="", encoding="utf-8-sig") as f:  # utf-8-sig strips Excel's BOM
            for row in csv.DictReader(f):
                row = {(k or "").strip(): v for k, v in row.items() if k}
                if not any((v or "").strip() for v in row.values()):
                    continue  # skip blank lines
                for k in row:
                    if k not in headers:
                        headers.append(k)
                row["Source file"] = path.name
                rows.append(row)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=headers + ["Source file"], restval="")
        writer.writeheader()
        writer.writerows(rows)
    return len(rows)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("usage: merge_csvs.py <folder> <output.csv>")
    n = merge(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"wrote {n} rows to {sys.argv[2]}")
```

## Good to know

- Headers are matched exactly after trimming spaces, so `Email` and `E-mail` become two columns. Rename headers first if your sources disagree.
- Files are processed in name order; prefixing exports with dates (`2026-10-01-orders.csv`) keeps rows in date order.
- Apps Script can process roughly 6 minutes of work per run. That's plenty for tens of thousands of rows; bigger jobs should use the Python version.

---
**Want it set up for you?** Get a free quote: https://cogandcell.com/#request (fixed price, see https://cogandcell.com/pricing)
