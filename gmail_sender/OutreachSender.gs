/**
 * Outreach sender for Google Sheets + Gmail (runs inside YOUR Google account only).
 * Sheet tab name: "Queue"  Columns: To | Company | Country | Buying Intent | Subject | Body | Status | Date Sent | Notes
 * Status values you can type yourself: "Replied", "Do not contact", "Skip" -> those rows are never emailed.
 */
const SHEET_NAME   = 'Queue';
const SENDER_NAME  = 'Jacobs Roland';
const DAILY_LIMIT  = 20;      // emails per run (one run per weekday). Start at 10 in week 1.
const MAX_BOUNCES  = 3;       // stop sending if more bounces than this arrived in the last 24 hours
const PAUSED       = false;   // set to true to stop all sending without deleting anything
const TEST_MODE    = true;    // true = sends only to yourself so you can check formatting. Set false to go live.

function sendDailyBatch() {
  if (PAUSED) { Logger.log('Paused.'); return; }
  const day = new Date().getDay();
  if (day === 0 || day === 6) { Logger.log('Weekend - no sending.'); return; }

  const bounces = GmailApp.search('from:mailer-daemon newer_than:1d').length;
  if (bounces > MAX_BOUNCES) { Logger.log('Too many bounces today (' + bounces + '). Stopping.'); return; }

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const c = n => head.indexOf(n);
  const me = Session.getActiveUser().getEmail();
  let sent = 0;

  for (let i = 1; i < data.length && sent < DAILY_LIMIT; i++) {
    const row = data[i];
    if (String(row[c('Status')]).trim() !== '') continue;           // already sent / replied / skipped
    const to = String(row[c('To')]).trim();
    if (!to) continue;
    const recipient = TEST_MODE ? me : to;
    try {
      GmailApp.sendEmail(recipient, row[c('Subject')], row[c('Body')], { name: SENDER_NAME });
      if (!TEST_MODE) {
        sh.getRange(i + 1, c('Status') + 1).setValue('Sent');
        sh.getRange(i + 1, c('Date Sent') + 1).setValue(new Date());
      }
      sent++;
      Utilities.sleep(5000 + Math.floor(Math.random() * 7000));    // 5-12 s gap between emails
    } catch (e) {
      sh.getRange(i + 1, c('Notes') + 1).setValue('Error: ' + e.message);
    }
    if (TEST_MODE && sent >= 2) break;                               // test mode: only 2 emails to yourself
  }
  Logger.log((TEST_MODE ? 'TEST: ' : '') + 'Sent ' + sent + ' emails.');
}

/** Run once: creates a daily trigger at about 9-10 am in your account's time zone. */
function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailyBatch').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Daily trigger created.');
}

/** Run to stop the automation completely. */
function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('Trigger removed.');
}
