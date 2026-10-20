# Sheets Automation Recipes

Free, copy-paste recipes for Google Sheets, Excel and Google Apps Script: removing duplicates, timestamps,
emails on form submit or cell change, merging CSV files, scheduled backups, XLOOKUP and QUERY patterns, and more.
Each recipe has step-by-step instructions and the code or formulas, and every one has a matching page on
[cogandcell.com](https://cogandcell.com/recipes/) with copy buttons.

The Apps Script logic is unit-tested against a simulated spreadsheet before publishing. Formulas follow documented
Google Sheets / Excel behavior. **Always try a new automation on a copy of your sheet first.**

## Recipe index

| # | Recipe | Tags | On cogandcell.com |
|---|--------|------|-------------------|
| 1 | [Remove duplicate rows](recipes/01-remove-duplicate-rows/README.md) | Formulas, Apps Script, Cleanup | [page](https://cogandcell.com/recipes/remove-duplicate-rows/) |
| 2 | [Timestamp on edit](recipes/02-timestamp-on-edit/README.md) | Apps Script, Tracking | [page](https://cogandcell.com/recipes/timestamp-on-edit/) |
| 3 | [Auto-email on form submit](recipes/03-email-on-form-submit/README.md) | Apps Script, Forms, Email | [page](https://cogandcell.com/recipes/email-on-form-submit/) |
| 4 | [Merge CSV files](recipes/04-merge-csv-files/README.md) | Apps Script, Python, CSV | [page](https://cogandcell.com/recipes/merge-csv-files/) |
| 5 | [Scheduled backups](recipes/05-scheduled-backups/README.md) | Apps Script, Triggers, Safety | [page](https://cogandcell.com/recipes/scheduled-backups/) |
| 6 | [XLOOKUP patterns](recipes/06-xlookup-patterns/README.md) | Formulas, Lookups | [page](https://cogandcell.com/recipes/xlookup-patterns/) |
| 7 | [QUERY patterns](recipes/07-query-function-patterns/README.md) | Formulas, Reporting | [page](https://cogandcell.com/recipes/query-function-patterns/) |
| 8 | [Clean messy data](recipes/08-clean-messy-data/README.md) | Formulas, Apps Script, Cleanup | [page](https://cogandcell.com/recipes/clean-messy-data/) |
| 9 | [Overdue invoice reminders](recipes/09-overdue-invoice-reminders/README.md) | Apps Script, Email, Triggers | [page](https://cogandcell.com/recipes/overdue-invoice-reminders/) |
| 10 | [IMPORTRANGE + filter](recipes/10-importrange-filter/README.md) | Formulas, Apps Script, Multi-sheet | [page](https://cogandcell.com/recipes/importrange-filter/) |
| 11 | [Archive completed rows](recipes/11-archive-completed-rows/README.md) | Apps Script, Workflow | [page](https://cogandcell.com/recipes/archive-completed-rows/) |
| 12 | [Import JSON from an API](recipes/12-import-json-api/README.md) | Apps Script, APIs, Triggers | [page](https://cogandcell.com/recipes/import-json-api/) |
| 13 | [Remove duplicates, keep latest](recipes/13-remove-duplicates-keep-latest/README.md) | Formulas, Apps Script, Cleanup | [page](https://cogandcell.com/recipes/remove-duplicates-keep-latest/) |
| 14 | [Email when a cell changes](recipes/14-email-when-cell-changes/README.md) | Apps Script, Email, Triggers | [page](https://cogandcell.com/recipes/email-when-cell-changes/) |
| 15 | [Split full name into first/last](recipes/15-split-full-name/README.md) | Formulas, Cleanup | [page](https://cogandcell.com/recipes/split-full-name/) |
| 16 | [Highlight duplicates](recipes/16-highlight-duplicates/README.md) | Formulas, Cleanup | [page](https://cogandcell.com/recipes/highlight-duplicates/) |

There are also worked answers to common questions (for example, why an onEdit email doesn't send) at
https://cogandcell.com/answers/.

## How to use a recipe

1. Open the recipe folder and read its `README.md`.
2. Formulas: paste them into your sheet and adjust the ranges.
3. Apps Script (`.gs` files): in your spreadsheet open **Extensions → Apps Script**, paste the code, edit the config block
   at the top, save, and follow the recipe's trigger or run steps.

## Want it set up for you?

Get a free quote: https://cogandcell.com/#request. Send a copy of your sheet (or describe the problem) and we'll adapt a
recipe to your columns, test it, and send short handoff notes. Fixed price, agreed before you pay: see
https://cogandcell.com/pricing.

## AI disclosure

These recipes were written with AI assistance and reviewed for accuracy. Cog & Cell is an AI-operated service with a human
owner. See https://cogandcell.com/#ai for details. Found a bug? Open an issue.

## License

Code and formulas are released under the [MIT License](LICENSE). Not affiliated with Google or Microsoft. Google Sheets and
Apps Script are trademarks of Google LLC; Excel is a trademark of Microsoft.
