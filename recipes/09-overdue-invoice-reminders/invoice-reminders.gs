/**
 * Daily overdue-invoice reminders from a sheet.
 * Columns (row 1 = headers): Invoice # | Client | Email | Amount | Due date | Status | Last reminder
 * Starts in dryRun mode: it only logs who WOULD be emailed until you set dryRun to false.
 */
const REMINDER_CONFIG = {
  sheetName: 'Invoices',
  headerRows: 1,
  cols: { invoice: 1, client: 2, email: 3, amount: 4, due: 5, status: 6, lastReminder: 7 },
  paidValues: ['paid', 'void', 'cancelled'], // statuses that mean "don't chase"
  graceDays: 3,          // wait this many days after the due date before the first reminder
  minDaysBetween: 7,     // then at most one reminder per week per invoice
  dryRun: true,          // set to false when the log looks right
  businessName: 'Your Business',
  currency: '$'
};

const DAY_MS = 24 * 60 * 60 * 1000;

function installDailyReminder() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'sendOverdueReminders')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendOverdueReminders').timeBased().everyDays(1).atHour(9).create();
}

/** Pure function: which rows need a reminder today? Returns [{rowNumber, ...fields}]. */
function findDueReminders(rows, today, cfg) {
  const c = cfg.cols;
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  const due = [];
  rows.forEach((row, i) => {
    const status = String(row[c.status - 1] || '').trim().toLowerCase();
    const email = String(row[c.email - 1] || '').trim();
    const dueDate = row[c.due - 1];
    const last = row[c.lastReminder - 1];
    if (cfg.paidValues.indexOf(status) !== -1) return;
    if (!email || !(dueDate instanceof Date)) return;
    const daysOverdue = Math.floor((startOfToday - dueDate.getTime()) / DAY_MS);
    if (daysOverdue < cfg.graceDays) return;
    if (last instanceof Date && (startOfToday - last.getTime()) / DAY_MS < cfg.minDaysBetween) return;
    due.push({
      rowNumber: i + cfg.headerRows + 1,
      invoice: row[c.invoice - 1],
      client: row[c.client - 1],
      email: email,
      amount: row[c.amount - 1],
      daysOverdue: daysOverdue
    });
  });
  return due;
}

function sendOverdueReminders() {
  const cfg = REMINDER_CONFIG;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow <= cfg.headerRows) return;
  const width = Math.max.apply(null, Object.keys(cfg.cols).map((k) => cfg.cols[k]));
  const rows = sheet.getRange(cfg.headerRows + 1, 1, lastRow - cfg.headerRows, width).getValues();
  const today = new Date();

  findDueReminders(rows, today, cfg).forEach((r) => {
    const subject = 'Reminder: invoice ' + r.invoice + ' is ' + r.daysOverdue + ' days overdue';
    const body = 'Hi ' + r.client + ',\n\n' +
      'A friendly reminder that invoice ' + r.invoice + ' for ' + cfg.currency + r.amount +
      ' was due ' + r.daysOverdue + ' days ago. If you have already paid, thank you, and please ignore this note.\n\n' +
      'Thanks,\n' + cfg.businessName;
    if (cfg.dryRun) {
      Logger.log('[dry run] would email ' + r.email + ': ' + subject);
      return;
    }
    MailApp.sendEmail(r.email, subject, body);
    sheet.getRange(r.rowNumber, cfg.cols.lastReminder).setValue(today);
  });
}
