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
const SHEET_NAME = 'Sheet1'; // change to match your sheet tab name
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

// ═════════════════════════════════════════════════════════════════
// NEW functions — only called when ?action=summary or ?action=search
// Nothing above this line was changed.
// ═════════════════════════════════════════════════════════════════

// ── Aggregate totals for the fund-balance website ─────────────────
function getSummary() {
  const data         = getSheet().getDataRange().getValues();
  let totalCollected = 0;
  let donorCount     = 0;
  let totalAdults    = 0;
  let totalKids      = 0;
  let totalChicken   = 0;
  let totalVeg       = 0;

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const name   = String(row[1]).trim();
    const amount = parseFloat(row[10]) || 0; // col K

    if (!name && amount === 0) continue; // skip empty rows

    donorCount++;
    totalAdults  += parseInt(row[3]) || 0;                              // D
    totalKids    += (parseInt(row[4]) || 0) + (parseInt(row[5]) || 0); // E + F
    totalVeg     += parseInt(row[7]) || 0;                              // H
    totalChicken += parseInt(row[8]) || 0;                              // I
    totalCollected += amount; // sum all rows regardless of status (Confirmed, Paid, etc.)
  }

  return {
    totalCollected: totalCollected.toFixed(2),
    donorCount,
    totalAdults,
    totalKids,
    totalChicken,
    totalVeg,
  };
}

// ── Search rows by name (partial, case-insensitive) ───────────────
function searchByName(query) {
  if (!query) return { error: 'No name provided' };

  const q       = query.toLowerCase();
  const data    = getSheet().getDataRange().getValues();
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row  = data[i];
    const name = String(row[1]).trim();
    if (!name) continue;

    if (name.toLowerCase().includes(q)) {
      results.push({
        name    : name,
        contact : String(row[2]).trim(),  // C
        adults  : parseInt(row[3]) || 0, // D
        kids12  : parseInt(row[4]) || 0, // E
        kidsU12 : parseInt(row[5]) || 0, // F
        veg     : parseInt(row[7]) || 0, // H
        chicken : parseInt(row[8]) || 0, // I
        status  : String(row[9]).trim(), // J
        amount  : String(row[10]).trim(),// K
      });
    }
  }

  return { results };
}
