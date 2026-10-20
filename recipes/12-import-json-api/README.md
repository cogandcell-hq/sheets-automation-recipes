# Import JSON from an API into Google Sheets on a schedule

Pull data from any JSON API (your store, CRM, or a public dataset) into a tab, flattening nested fields, refreshed hourly.

- **Setup time:** 15 minutes
- **You'll need:** An API URL that returns JSON. The demo uses a free public test API, so it works as-is.
- **Full page with copy buttons:** https://cogandcell.com/recipes/import-json-api/

## Steps

1. **Extensions → Apps Script**, paste the script, and run `importJson`. You'll get an **API import** tab with 10 demo users.
2. Swap `url` for your API. If the response wraps the list (e.g. `{"data": [...]}`), set `listKey: 'data'`.
3. List the `fields` you want; use dots for nested values like `address.city`.
4. Needs an API key? Go to **Project Settings → Script properties**, add e.g. `API_KEY`, and set `apiKeyProperty: 'API_KEY'`. Never paste keys into the code.
5. Run `installHourlyImport` to refresh it automatically.

**Code.gs** (also saved as [`json-import.gs`](json-import.gs)):

```javascript
/**
 * Pull JSON from an API into a sheet, on demand or on a schedule.
 * The demo URL is a free public test API, so you can try it before pointing it at your own.
 */
const API_CONFIG = {
  url: 'https://jsonplaceholder.typicode.com/users',
  listKey: '',     // if the API returns {"data": [...]}, put 'data' here
  fields: ['id', 'name', 'email', 'address.city', 'company.name'], // dot paths for nested values
  sheetName: 'API import',
  apiKeyProperty: '' // e.g. 'API_KEY': stored in Project Settings > Script properties, never in code
};

/** Pure: read a nested value like "address.city". */
function getPath(obj, path) {
  return path.split('.').reduce((o, k) => (o !== null && o !== undefined ? o[k] : undefined), obj);
}

/** Pure: turn an array of objects into rows with a header. */
function toRows(items, fields) {
  const body = items.map((item) => fields.map((f) => {
    const v = getPath(item, f);
    if (v === null || v === undefined) return '';
    return typeof v === 'object' ? JSON.stringify(v) : v;
  }));
  return [fields].concat(body);
}

function importJson() {
  const cfg = API_CONFIG;
  const headers = {};
  if (cfg.apiKeyProperty) {
    const key = PropertiesService.getScriptProperties().getProperty(cfg.apiKeyProperty);
    headers.Authorization = 'Bearer ' + key;
  }
  const res = UrlFetchApp.fetch(cfg.url, { headers: headers, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('API returned ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 200));
  }
  const data = JSON.parse(res.getContentText());
  const items = Array.isArray(data) ? data : (getPath(data, cfg.listKey) || []);
  const rows = toRows(items, cfg.fields);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(cfg.sheetName) || ss.insertSheet(cfg.sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
}

function installHourlyImport() {
  ScriptApp.newTrigger('importJson').timeBased().everyHours(1).create();
}
```

## Good to know

- The script sends the key as `Authorization: Bearer …`, which most APIs use. Some want a different header or a query parameter; check your API's docs.
- Many APIs return results in pages (e.g. 100 at a time). Fetching every page needs a small loop. That's a typical Quick fix.
- Arrays and objects you don't list as dot paths are written as JSON text so nothing is silently lost.

---
**Want it set up for you?** Get a free quote: https://cogandcell.com/#request (fixed price, see https://cogandcell.com/pricing)
