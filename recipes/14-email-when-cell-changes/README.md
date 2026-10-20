# Google Sheets: send an email when a cell changes (Apps Script)

Email yourself or the person on that row when a status column changes, for example to 'Approved'. Uses an installable trigger, because simple onEdit can't send mail.

- **Setup time:** 10 minutes
- **You'll need:** A Google Sheet you own or can edit, with a column to watch (a dropdown works best).
- **Full page with copy buttons:** https://cogandcell.com/recipes/email-when-cell-changes/

## No-code option first

If you only need to know *that* something changed, Google Sheets has this built in: **Tools → Notification settings → Edit notifications**. It emails you about any change (immediately or as a daily digest), but it can't filter by column or value, and it can't email someone else. For that, use the script.

## Script: email when a specific column changes to a specific value

1. **Extensions → Apps Script**, paste the script.
2. Set `sheetName`, `watchColumn` (4 = column D) and `triggerValues` (leave it as `[]` to email on any change).
3. Set `notifyTo`, or set `emailColumn` to the column holding each row's email to notify that person instead.
4. Select `installEditTrigger` and click **Run** once. Approve the permissions (it needs to send email as you).
5. Change a watched cell to a trigger value. The email arrives within a minute.

**Code.gs** (also saved as [`email-on-change.gs`](email-on-change.gs)):

```javascript
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
```

## Good to know

- Why not just name the function `onEdit`? Simple triggers run without authorization, so they can't call MailApp. That's the most common reason “my onEdit email doesn't send”.
- Edit triggers fire for changes people make. They don't fire when a formula result changes, or when rows arrive from a Google Form or another script. For form rows, see [email on form submit](../email-on-form-submit/).
- Multi-cell pastes are ignored on purpose, so pasting 200 rows doesn't send 200 emails.
- Emails come from the account that ran `installEditTrigger`, and count against its daily Apps Script email quota.

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
