/**
 * RSVP receiver for branubeginnin.love
 *
 * Setup:
 *  1. Create a Google Sheet → Extensions → Apps Script.
 *  2. Replace the editor contents with this file and save.
 *  3. Deploy → New deployment → type "Web app"
 *       Execute as: Me
 *       Who has access: Anyone
 *  4. Copy the Web app URL (ends in /exec) into index.html (form action).
 *
 * After editing this script, use Deploy → Manage deployments → Edit →
 * Version: "New version" so the same URL serves the new code.
 */

const SHEET_NAME = 'RSVPs';
const FIELDS  = ['name', 'email', 'phone', 'attending', 'guests', 'message', 'language'];
const HEADERS = ['Received', 'Name', 'Email', 'Phone / WhatsApp', 'Attending', 'Guests', 'Message', 'Language'];

function doPost(e) {
  const p = (e && e.parameter) || {};

  // Spam trap filled → pretend success, store nothing
  if (p.website) return json({ ok: true });

  if (!p.name || !p.email || !p.attending) {
    return json({ ok: false, error: 'missing_fields' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    }

    sheet.appendRow([new Date()].concat(FIELDS.map(f => clean(p[f]))));
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// Lets you open the URL in a browser to confirm the deployment is live
function doGet() {
  return json({ ok: true, service: 'rsvp' });
}

// Trim, cap length, and stop values like "=HYPERLINK(...)" running as formulas
function clean(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
