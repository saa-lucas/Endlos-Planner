// =====================================================
// MACROS - SYSTEM AUTOMATION LAYER
// =====================================================


// =====================================================
// SMART LOCK SYSTEM
// =====================================================

function applySmartLock(lockOldWeeks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Organisieren");

  if (!sheet) return;

  // Remove locks antigos
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  protections.forEach(p => {
    if (p.getDescription() === 'LockedOldWeeks') {
      p.remove();
    }
  });

  // Aplica lock se necessário
  if (lockOldWeeks) {
    const maxCols = sheet.getLastColumn();

    if (maxCols > 20) {
      const range = sheet.getRange(1, 1, sheet.getMaxRows(), maxCols - 18);

      const protection = range.protect();
      protection.setDescription('LockedOldWeeks');
      protection.setWarningOnly(true);
    }
  }
}


// =====================================================
// SYSTEM PROMPT
// =====================================================

function promptFullSync() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    '⚠️ Confirm Rebuild Daten',
    'Are you sure you want to rebuild the entire history? This might take a few seconds.',
    ui.ButtonSet.YES_NO
  );

  return response === ui.Button.YES;
}