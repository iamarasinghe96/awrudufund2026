// ─────────────────────────────────────────────────────────────────
// Albury Wodonga Payment Receipt Tool — Google Apps Script
// ─────────────────────────────────────────────────────────────────
//
// COLUMN MAPPING:
//   A  Timestamp
//   B  Name
//   C  Contact
//   D  Adults
//   E  Kids 12+
//   F  Kids Under 12
//   G  Remark / Receipt#
//   H  Vegetarian (number)
//   I  Chicken (number)
//   J  Status
//   K  Amount
//   L  Payment Date
// ─────────────────────────────────────────────────────────────────
const SHEET_NAME = 'Sheet1'; // main data sheet
// getSummary() and searchByName() live in appscript_website.gs
// getCategoryBreakdown() lives in appscript_sheet3.gs
function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}
// ── Handle CORS preflight ─────────────────────────────────────────
function doOptions(e) {
  return ContentService.createTextOutput('');
}
// ── GET: lookup contact + meal by remark / receipt# ──────────────
function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    // ── NEW: route to summary or name-search when ?action= is present ──
    const action = (e.parameter.action || '').trim().toLowerCase();
    if (action === 'summary') {
      output.setContent(JSON.stringify(getSummary()));
      return output;
    }
    if (action === 'search') {
      output.setContent(JSON.stringify(searchByName((e.parameter.name || '').trim())));
      return output;
    }
    // ── ORIGINAL code below — not modified ───────────────────────
    const remark = (e.parameter.remark || '').trim().toLowerCase();
    if (!remark) {
      output.setContent(JSON.stringify({ error: 'No remark provided' }));
      return output;
    }
    const data = getSheet().getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowRemark = String(data[i][6]).trim().toLowerCase(); // col G
      if (rowRemark === remark) {
        output.setContent(JSON.stringify({
          contact     : String(data[i][2]).trim(),  // C
          name        : String(data[i][1]).trim(),  // B
          meal_veg    : String(data[i][7]).trim(),  // H
          meal_chicken: String(data[i][8]).trim(),  // I
          status      : String(data[i][9]).trim(),  // J
        }));
        return output;
      }
    }
    output.setContent(JSON.stringify({ error: 'Not found' }));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }
  return output;
}
// ── POST: append a new payment record ────────────────────────────
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  try {
    const b = JSON.parse(e.postData.contents);
    getSheet().appendRow([
      new Date(),            // A  Timestamp
      b.name        || '',   // B  Name
      b.contact     || '',   // C  Contact
      b.adults      || 0,    // D  Adults
      b.kids12      || 0,    // E  Kids 12+
      b.kidsU12     || 0,    // F  Kids Under 12
      b.remark      || '',   // G  Remark / Receipt#
      b.meal_veg    || 0,    // H  Vegetarian
      b.meal_chicken|| 0,    // I  Chicken
      b.status      || '',   // J  Status
      b.amount      || '',   // K  Amount
      b.date        || '',   // L  Payment Date
    ]);
    output.setContent(JSON.stringify({ success: true }));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }
  return output;
}
