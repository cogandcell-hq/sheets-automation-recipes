/**
 * Nightly backup: copies this spreadsheet into a Drive folder and keeps only the newest N copies.
 * Run installDailyBackup once; after that it runs on its own.
 */
const BACKUP_CONFIG = {
  folderId: 'PASTE_BACKUP_FOLDER_ID', // a folder you own, ideally not shared widely
  keep: 30,                           // how many backups to keep
  hour: 2                             // hour of day (script time zone) to run
};

function installDailyBackup() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'backupSpreadsheet')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('backupSpreadsheet').timeBased().everyDays(1).atHour(BACKUP_CONFIG.hour).create();
}

/** Pure function: given [{id, created}] returns the ids to delete, keeping the newest `keep`. */
function pickBackupsToDelete(files, keep) {
  return files
    .slice()
    .sort((a, b) => b.created - a.created)
    .slice(keep)
    .map((f) => f.id);
}

function backupSpreadsheet() {
  const cfg = BACKUP_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folder = DriveApp.getFolderById(cfg.folderId);
  const prefix = 'Backup - ' + ss.getName() + ' - ';
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  DriveApp.getFileById(ss.getId()).makeCopy(prefix + stamp, folder);

  const existing = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf(prefix) === 0) existing.push({ id: f.getId(), created: f.getDateCreated().getTime() });
  }
  pickBackupsToDelete(existing, cfg.keep).forEach((id) => DriveApp.getFileById(id).setTrashed(true));
}
