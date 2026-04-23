function applySmartLock(lockOldWeeks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Organisieren");
  
  if (!sheet) return;

  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    if (p.getDescription() === 'LockedOldWeeks') p.remove();
  });

  if (lockOldWeeks) {
    const maxCols = sheet.getLastColumn();
    if (maxCols > 20) { 
      const rangeToLock = sheet.getRange(1, 1, sheet.getMaxRows(), maxCols - 18);
      const protection = rangeToLock.protect().setDescription('LockedOldWeeks');
      protection.setWarningOnly(true);
    }
  }
}