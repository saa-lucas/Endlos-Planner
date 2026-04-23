function getWeekNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  let weeks = [];

  for (let i = 0; i < namedRanges.length; i++) {
    let nr = namedRanges[i];
    let name = nr.getName();
    
    if (name.startsWith("Week")) {
      weeks.push({
        name: name,
        col: nr.getRange().getColumn()
      });
    }
  }

  weeks.sort((a, b) => b.col - a.col);
  return weeks.map(w => w.name);
}

function focusWeek(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const targetRange = ss.getRangeByName(rangeName);
  
  if (!targetRange) return;

  const maxCols = sheet.getMaxColumns();
  sheet.showColumns(1, maxCols);
  
  const startCol = targetRange.getColumn();
  const numCols = targetRange.getNumColumns();
  const endCol = startCol + numCols - 1;

  if (startCol > 4) {
    sheet.hideColumns(4, startCol - 4);
  }

  if (endCol < maxCols) {
    sheet.hideColumns(endCol + 1, maxCols - endCol);
  }
  
  sheet.showColumns(1, 3);
  targetRange.activate();
}

function showAllColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.showColumns(1, sheet.getMaxColumns());
}

function clearOldWeeks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ranges = ss.getNamedRanges();
  let removedCount = 0;
  
  ranges.forEach(range => {
    if (range.getName().startsWith("Week")) {
      range.remove();
      removedCount++;
    }
  });
  
  SpreadsheetApp.getUi().alert(removedCount + " old intervals have been deleted.");
}