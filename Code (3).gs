/**
 * Form backend: saves name / email / phone to the "Responses" tab.
 *
 * Setup:
 * 1. Open your Google Sheet -> Extensions -> Apps Script (so the script is bound to the sheet).
 * 2. Paste this file, save.
 * 3. Deploy -> New deployment -> Type: Web app
 *    Execute as: Me | Who has access: Anyone
 * 4. Copy the Web app URL (ends with /exec).
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

function handleRequest_(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var params = (e && e.parameter) ? e.parameter : {};
    var name = String(params.name || '').trim();
    var email = String(params.email || '').trim();
    var phone = String(params.phone || '').trim();

    if (!name && !email && !phone) {
      return json_({ success: false, error: 'No data received' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return json_({ success: false, error: 'Script is not bound to a spreadsheet' });
    }

    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Auto-create bold headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS.length)
        .setValues([HEADERS])
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Store phone as text so leading zeros / "+" are preserved
    var nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 4).setNumberFormat('@');
    sheet.getRange(nextRow, 1, 1, HEADERS.length)
      .setValues([[new Date(), name, email, phone]]);

    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
