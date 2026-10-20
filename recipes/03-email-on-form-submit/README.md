# Send an email automatically when a Google Form is submitted

Get every response in your inbox (with Reply going straight to the customer) and send respondents an instant confirmation.

- **Setup time:** 10 minutes
- **You'll need:** A Google Form linked to a response spreadsheet (Form → Responses → Link to Sheets).
- **Full page with copy buttons:** https://cogandcell.com/recipes/email-on-form-submit/

## Steps

1. Open the **response spreadsheet** (not the form), then **Extensions → Apps Script**.
2. Paste the script and set `notifyTo` to your address.
3. Set `emailQuestion` and `nameQuestion` to the *exact* question titles in your form (e.g. “Email address”). If your form collects emails automatically, use `'Email Address'`.
4. Choose the function `installFormTrigger` in the toolbar and click **Run**. Approve the permissions (it needs to send email as you).
5. Submit a test response. You should get the notification within a minute.

**Code.gs** (also saved as [`form-email.gs`](form-email.gs)):

```javascript
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
```

## Good to know

- Emails are sent from your Google account. Free Gmail accounts can send to roughly 100 recipients per day from scripts; Workspace accounts get more.
- Edit the confirmation text in `confirmationBody`; `{name}` and `{answers}` are filled in for you.
- Running `installFormTrigger` twice is safe; it won't create a duplicate trigger.
- To stop it: Apps Script → Triggers (clock icon) → delete the trigger.

---
Want it done for you? Fixed price, free quote: see pricing at https://cogandcell.com/pricing
