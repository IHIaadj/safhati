/**
 * Safhati contact form → Google Sheets
 * 1. Create a Google Sheet and paste its ID below.
 * 2. In Extensions > Apps Script, paste this file and deploy as a Web App.
 */
const SPREADSHEET_ID = 'PASTE_YOUR_SPREADSHEET_ID_HERE';
const SHEET_NAME = 'Contacts';
const HEADERS = ['Timestamp', 'Name', 'Company', 'Email', 'Phone', 'Language', 'Message', 'Source'];

function doGet(event) {
  const data = (event && event.parameter) || {};
  const callback = data.cb || data.callback;

  if (callback) {
    return jsonp_(callback, saveContact_(data));
  }

  return output_({ ok: true, service: 'Safhati contact form' });
}

function doPost(event) {
  const data = (event && event.parameter) || {};
  return output_(saveContact_(data));
}

function saveContact_(data) {
  // Honeypot: silently ignore obvious bot submissions.
  if (data.website) return { ok: true };

  if (!data.name || !data.email || !data.message) {
    return { ok: false, error: 'Missing required fields' };
  }

  let lock;
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.company),
      clean_(data.email),
      clean_(data.phone),
      clean_(data.language),
      clean_(data.message),
      clean_(data.source),
    ]);

    return { ok: true };
  } catch (error) {
    return { ok: false, error: String(error && error.message ? error.message : error) };
  } finally {
    if (lock) {
      try {
        lock.releaseLock();
      } catch (error) {}
    }
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange('A1:H1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function clean_(value) {
  return String(value || '').trim().slice(0, 5000);
}

function output_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(callback, payload) {
  const safeCallback = String(callback || '').replace(/[^\w.$]/g, '');
  return ContentService
    .createTextOutput(safeCallback + '(' + JSON.stringify(payload) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
