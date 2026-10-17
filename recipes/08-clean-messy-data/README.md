# Clean messy data: spaces, names, phone numbers, emails and 'numbers stored as text'

Formulas to normalize the data you get from forms and exports, plus a one-click script that strips the invisible characters that break lookups.

- **Setup time:** 5 minutes
- **You'll need:** Google Sheets (formulas also work in Excel 365 except REGEX functions).
- **Full page with copy buttons:** https://cogandcell.com/recipes/clean-messy-data/

## Formulas (put them in a helper column)

**Trim spaces (also collapses double spaces inside text in Sheets)**

```
=TRIM(SUBSTITUTE(A2, CHAR(160), " "))
```

**Email: trimmed and lower-case**

```
=LOWER(TRIM(A2))
```

**Name in proper case**

```
=PROPER(TRIM(A2))
```

**First name / last name from a full name**

```
=REGEXEXTRACT(TRIM(A2), "^\S+")
=REGEXEXTRACT(TRIM(A2), "\S+$")
```

**US phone numbers to (555) 123-4567, flagging anything odd**

```
=LET(d, REGEXREPLACE(TO_TEXT(A2), "\D", ""), d10, IF(AND(LEN(d)=11, LEFT(d,1)="1"), RIGHT(d, 10), d), IF(LEN(d10)=10, REGEXREPLACE(d10, "(\d{3})(\d{3})(\d{4})", "($1) $2-$3"), "CHECK: "&A2))
```

**"$1,234.50" stored as text to a real number**

```
=IFERROR(VALUE(REGEXREPLACE(TO_TEXT(A2), "[^0-9.\-]", "")), "")
```

**Clean a whole column at once**

```
=ARRAYFORMULA(IF(A2:A="", , LOWER(TRIM(A2:A))))
```

## One-click clean-up script

Trims every selected cell in place and removes non-breaking spaces and zero-width characters (common in data pasted from websites and PDFs). Formulas are left alone.

1. **Extensions → Apps Script**, paste, save, reload the sheet.
2. Select the cells to clean, then **Cleanup → Clean selected cells**.

**Code.gs** (also saved as [`clean.gs`](clean.gs)):

```javascript
/**
 * Clean the selected cells in place: trims spaces, removes non-breaking and zero-width
 * characters that break lookups, and collapses double spaces. Formulas are left untouched.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Cleanup')
    .addItem('Clean selected cells', 'cleanSelection')
    .addToUi();
}

/** Pure function: clean one value. Non-text values (numbers, dates) pass through. */
function cleanValue(value) {
  if (typeof value !== 'string') return value;
  return value
    .replace(/[\u200B-\u200D\uFEFF]/g, '') // zero-width characters
    .replace(/\u00A0/g, ' ')               // non-breaking spaces
    .replace(/[ \t]+/g, ' ')               // runs of spaces/tabs
    .replace(/ *\n */g, '\n')              // spaces around line breaks
    .trim();
}

function cleanSelection() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) return;
  const values = range.getValues();
  const formulas = range.getFormulas();
  let changed = 0;
  const next = values.map((row, r) => row.map((v, c) => {
    if (formulas[r][c]) return formulas[r][c]; // keep formulas as formulas
    const cleaned = cleanValue(v);
    if (cleaned !== v) changed++;
    return cleaned;
  }));
  // setValues treats strings starting with "=" as formulas, which restores them unchanged.
  range.setValues(next);
  SpreadsheetApp.getActiveSpreadsheet().toast('Cleaned ' + changed + ' cell(s).');
}
```

## Good to know

- When you're happy with a helper column, copy it and use **Paste special → Values only** over the original.
- The phone formula is for US/Canada numbers. International numbers need country-specific rules; flagging them with “CHECK” is safer than guessing.
- If you use both this script and the dedupe script in one project, keep just one `onOpen` and add both menu items to it.

---
Need it set up for you? $49 fixed-price fix at https://cogandcell.com
