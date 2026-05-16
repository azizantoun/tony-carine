/**
 * Tony & Carine RSVP — Google Apps Script web app
 *
 * SETUP (one-time, ~3 minutes):
 *
 * 1. Open Google Sheets → create a new sheet → name it e.g. "Tony & Carine RSVPs".
 *    Leave row 1 blank — the script will write headers automatically.
 *
 * 2. In that sheet: Extensions → Apps Script. Delete any placeholder code,
 *    paste this entire file, and click the Save (disk) icon.
 *
 * 3. Top-right → "Deploy" → "New deployment".
 *      - Type: Web app
 *      - Description: anything (e.g. "RSVP v1")
 *      - Execute as: Me (your account)
 *      - Who has access: Anyone
 *    Click Deploy. Authorize when prompted.
 *
 * 4. Copy the "Web app URL" it gives you (ends in /exec).
 *
 * 5. In index.html, replace the placeholder string in the line:
 *      const RSVP_ENDPOINT = 'PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
 *    with that URL. Commit & push.
 *
 * Each RSVP becomes a new row in the sheet AND emails azizantoun@gmail.com.
 */

const NOTIFY_EMAIL = 'azizantoun@gmail.com';

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const params = e.parameter || {};

  const name        = params.name        || '';
  const attendance  = params.attendance  || '';
  const guestCount  = params.guest_count || '';
  const wishes      = params.wishes      || '';
  const submittedAt = params.submitted_at || new Date().toISOString();

  // Write headers on first run
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Submitted At', 'Name', 'Attending', 'Guests', 'Wishes']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  }

  sheet.appendRow([submittedAt, name, attendance, guestCount, wishes]);

  const attendingLabel =
    attendance === 'yes' ? 'Joyfully accepts' :
    attendance === 'no'  ? 'Regretfully declines' :
    attendance;

  const subject = `Tony & Carine RSVP — ${name} (${attendingLabel})`;
  const body =
`New RSVP received

Name:     ${name}
Status:   ${attendingLabel}
Guests:   ${guestCount}
Wishes:   ${wishes || '—'}

Submitted: ${submittedAt}
`;

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput('Tony & Carine RSVP endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
