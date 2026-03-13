// ─────────────────────────────────────────────────────────────────
// Sheet 3 — Category Breakdown
// ─────────────────────────────────────────────────────────────────
//
// Sheet 3 structure:
//   A  Category label   (e.g. "Total Families", "1 Adult", …)
//   B  Total count      (static or formula-driven)
//   C  Paid count       (COUNTIF where Sheet1 col J = "Paid")
//
// This file is a separate Apps Script file in the same project.
// It shares the global scope with appscript.gs, so getSummary()
// in appscript.gs can call getCategoryBreakdown() defined here.
// ─────────────────────────────────────────────────────────────────

const SHEET3_NAME = 'Sheet3';

/**
 * Reads Sheet 3 and returns an array of category objects:
 *   { label: string, total: number, paid: number }
 * Called by getSummary() in appscript.gs.
 */
function getCategoryBreakdown() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET3_NAME);
  if (!sheet) return [];

  const data       = sheet.getDataRange().getValues();
  const categories = [];

  for (let i = 0; i < data.length; i++) {
    const label = String(data[i][0]).trim();
    const total = parseInt(data[i][1]) || 0;
    const paid  = parseInt(data[i][2]) || 0;
    if (label) categories.push({ label, total, paid });
  }

  return categories;
}
