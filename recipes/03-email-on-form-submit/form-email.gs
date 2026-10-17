/**
 * Email yourself (and optionally the respondent) every time a Google Form is submitted.
 * Put this in the spreadsheet that collects the form responses, then run installFormTrigger once.
 */
const FORM_EMAIL_CONFIG = {
  notifyTo: 'you@example.com',     // where new submissions go (comma-separate several)
  subjectPrefix: 'New form response',
  emailQuestion: 'Email',          // exact question title that holds the respondent's email
  nameQuestion: 'Name',            // exact question title for their name (optional)
  sendConfirmation: true,          // also send the respondent a "we got it" email
  confirmationSubject: 'Thanks, we received your request',
  confirmationBody: 'Hi {name},\n\nThanks for getting in touch. We received your request and will reply within one business day.\n\nYour answers:\n{answers}\n'
};

/** Run this once from the editor (Run > installFormTrigger) and approve the permissions. */
function installFormTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exists = ScriptApp.getProjectTriggers()
    .some((t) => t.getHandlerFunction() === 'onFormSubmitHandler');
  if (exists) return;
  ScriptApp.newTrigger('onFormSubmitHandler').forSpreadsheet(ss).onFormSubmit().create();
}

/** Pure function: turns e.namedValues into readable "Question: answer" lines. */
function formatAnswers(namedValues) {
  return Object.keys(namedValues)
    .filter((q) => q !== 'Timestamp')
    .map((q) => q + ': ' + [].concat(namedValues[q]).join(', '))
    .join('\n');
}

function firstAnswer(namedValues, question) {
  const v = namedValues[question];
  return v ? String([].concat(v)[0]).trim() : '';
}

function onFormSubmitHandler(e) {
  const cfg = FORM_EMAIL_CONFIG;
  const nv = e.namedValues || {};
  const answers = formatAnswers(nv);
  const email = firstAnswer(nv, cfg.emailQuestion);
  const name = firstAnswer(nv, cfg.nameQuestion) || 'there';
  const validEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);

  const notice = {
    to: cfg.notifyTo,
    subject: cfg.subjectPrefix + (name !== 'there' ? ': ' + name : ''),
    body: answers + '\n\nSheet: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl()
  };
  if (validEmail) notice.replyTo = email; // hit Reply to answer the respondent directly
  MailApp.sendEmail(notice);

  if (cfg.sendConfirmation && validEmail) {
    MailApp.sendEmail({
      to: email,
      subject: cfg.confirmationSubject,
      body: cfg.confirmationBody.replace('{name}', name).replace('{answers}', answers)
    });
  }
}
