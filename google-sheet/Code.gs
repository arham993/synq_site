/**
 * SYNQ enquiries -> Google Sheet
 *
 * Paste this into Extensions > Apps Script of the Google Sheet that should
 * collect enquiries, then Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone). Put the web app URL in
 * SHEET_URL inside index.html.
 */
const SHEET_NAME = 'Enquiries';
const HEADERS = ['Submitted at', 'Name', 'Phone', 'Email', 'Interested in', 'Preferred visit date', 'Message', 'Page'];

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return reply_({ ok: true }); // spam bot filled the hidden field

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      clean_(p.name), clean_(p.phone), clean_(p.email), clean_(p.unit),
      clean_(p.visit), clean_(p.message), clean_(p.page)
    ]);
    return reply_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return reply_({ ok: true, info: 'SYNQ enquiry endpoint is running.' });
}

// Keeps values as plain text so a cell never runs as a formula,
// and keeps phone numbers like +91... from being read as numbers.
function clean_(v) {
  const s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function reply_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
