/**
 * 
 * Focuses on a specific named range (week) while keeping time columns A:C visible.
 * This solves the printing overlap issue by hiding irrelevant columns.
 * @param {string} rangeName The name of the named range to display.
 */
function focusWeek(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const targetRange = ss.getRangeByName(rangeName);
  
  if (!targetRange) {
    throw new Error("Named range not found: " + rangeName);
  }

  // 1. Reset: Show all columns to recalculate visibility
  const maxCols = sheet.getMaxColumns();
  sheet.showColumns(1, maxCols);
  
  // 2. Define boundaries (A:C are columns 1, 2, and 3)
  const startCol = targetRange.getColumn(); 
  const numCols = targetRange.getNumColumns();
  const endCol = startCol + numCols - 1;

  // 3. Hide the "gap" between Column C (3) and the start of the selected Week
  // Only hides if the week starts after column D (4)
  if (startCol > 4) {
    sheet.hideColumns(4, startCol - 4);
  }

  // 4. Hide everything AFTER the selected week range until the end of the sheet
  if (endCol < maxCols) {
    sheet.hideColumns(endCol + 1, maxCols - endCol);
  }
  
  // 5. Safety check: Ensure A:C time columns remain visible
  sheet.showColumns(1, 3);
  
  // 6. Focus the view on the selected week
  targetRange.activate();
  
  return "Overview Activated: A:C + " + rangeName;
}

/**
 * Resets the sheet view to show all columns.
 */
function showAllColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.showColumns(1, sheet.getMaxColumns());
}