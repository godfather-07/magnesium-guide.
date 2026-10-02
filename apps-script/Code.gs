/**
 * Magnesium Supplement Guide: sign-up pop-up -> Google Sheet.
 *
 * Setup:
 * 1. Open the "Magnesium Supplement Checklist" sheet.
 * 2. Extensions -> Apps Script. Select ALL the old code, delete it, paste this file.
 * 3. Save (disk icon).
 * 4. Deploy -> Manage deployments -> pencil (Edit) -> Version: "New version" -> Deploy.
 *    Keep "Execute as: Me" and "Who has access: Anyone".
 *    Editing the existing deployment keeps the same URL, so the website needs no change.
 */

const SHEET_NAME = 'Sheet1';
const TIMEZONE = 'Asia/Kolkata';
const HEADERS = ['Timestamp (IST)', 'Name', 'Email', 'Phone', 'Consent'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
    const p = (e && e.parameter) || {};

    // Ignore bots that fill the hidden honeypot field.
    if (p['bot-field']) return json_({ result: 'ignored' });

    // Keep the header row in the current layout.
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);

    sheet.appendRow([
      Utilities.formatDate(new Date(), TIMEZONE, 'dd MMM yyyy, HH:mm'),
      safe_(p.name),
      safe_(p.email),
      p.phone ? "'" + String(p.phone).trim() : '',   // apostrophe keeps "+91" and leading zeros
      p.consent ? 'Yes (agreed by continuing)' : 'No'
    ]);

    return json_({ result: 'success' });
  } finally {
    lock.releaseLock();
  }
}

// Stop entries such as "=SUM(...)" from being treated as formulas.
function safe_(v) {
  const s = String(v || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor (Run button) to check that a row is written. */
function testAppend() {
  doPost({ parameter: { name: 'Test', email: 'test@example.com', phone: '+91 98765 43210', consent: 'yes' } });
}
