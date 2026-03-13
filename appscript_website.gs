// ─────────────────────────────────────────────────────────────────
// Fund-Balance Website — Google Apps Script
// Handles ?action=summary and ?action=search for index.html
// Relies on getCategoryBreakdown() in appscript_sheet3.gs
// ─────────────────────────────────────────────────────────────────

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
    totalCollected += amount;
  }

  return {
    totalCollected: totalCollected.toFixed(2),
    donorCount,
    totalAdults,
    totalKids,
    totalChicken,
    totalVeg,
    categories: getCategoryBreakdown(), // defined in appscript_sheet3.gs
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
