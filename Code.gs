/**
 * Google Apps Script for wedding RSVP.
 * 1. Create a Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Replace this entire file with this code.
 * 4. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 5. Copy the /exec URL into config.js.
 */

const SHEET_NAME = "RSVP";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Full Name", "Attendance", "Guests", "Message"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function clean_(value, max) {
  return String(value || "").replace(/[<>]/g, "").trim().slice(0, max);
}

function doPost(e) {
  const p = e && e.parameter ? e.parameter : {};
  const name = clean_(p.name, 100);
  const attendance = clean_(p.attendance, 30);
  const guests = Math.max(0, Math.min(10, Number(p.guests || 0)));
  const message = clean_(p.message, 500);

  if (!name || !message) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:"Missing required fields"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  getSheet_().appendRow([
    new Date(),
    name,
    attendance,
    guests,
    message
  ]);

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const action = p.action || "wishes";
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  const items = values.slice(1).reverse()
    .filter(r => r[1] && r[4])
    .slice(0, 100)
    .map(r => ({
      name: clean_(r[1], 100),
      message: clean_(r[4], 500)
    }));

  const payload = JSON.stringify({ok:true, items});
  if (p.callback) {
    const safe = String(p.callback).replace(/[^\w.$]/g, "");
    return ContentService.createTextOutput(`${safe}(${payload})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}