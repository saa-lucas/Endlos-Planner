/**
 * Finds the very last 18-column block, deletes its columns, 
 * and removes its associated Named Range.
 * Includes a safety check to prevent deleting the original base block.
 */
function deleteLast18ColBlock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  
  const blockWidth = 18;
  const dataStartCol = 4; // Column D (Start of the first block)
  
  const currentMaxCols = sheet.getMaxColumns();
  let lastBlockCol = 0;
  
  // 1. Locate the last existing block
  for (let col = dataStartCol; col <= currentMaxCols; col += blockWidth) {
    if (col + blockWidth - 1 <= currentMaxCols) {
      const merges = sheet.getRange(1, col, 5, blockWidth).getMergedRanges();
      if (merges.length > 0) {
        lastBlockCol = col;
      }
    }
  }
  
  // 2. Safety Checks
  if (lastBlockCol === 0) {
    SpreadsheetApp.getUi().alert("No blocks found.");
    return;
  }
  
  if (lastBlockCol === dataStartCol) {
    SpreadsheetApp.getUi().alert("⚠️ Safety Lock: You cannot delete the base template block.");
    return;
  }
  
  // 3. Confirmation Dialog (Important for destructive actions)
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Delete Last Week?', 
    'Are you sure you want to delete the last created week? This action cannot be undone.', 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return; // User canceled
  }

  // 4. Find and remove the associated Named Range
  const namedRanges = ss.getNamedRanges();
  for (let i = 0; i < namedRanges.length; i++) {
    const range = namedRanges[i].getRange();
    // If the named range starts at the exact same column as our last block, remove it
    if (range.getSheet().getName() === sheet.getName() && range.getColumn() === lastBlockCol) {
      namedRanges[i].remove();
      break; 
    }
  }
  
  // 5. Delete the physical columns
  sheet.deleteColumns(lastBlockCol, blockWidth);
  
  // 6. Success message
  ui.alert("Success", "The last week block has been deleted.", ui.ButtonSet.OK);
}