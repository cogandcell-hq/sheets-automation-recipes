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
