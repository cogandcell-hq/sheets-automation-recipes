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
