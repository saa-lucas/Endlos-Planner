// SMART LOCK INVISÍVEL E AUTOMÁTICO
function applySmartLock(lockOldWeeks) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Organisieren");
  
  if (!sheet) return;

  // 1. Limpa o cadeado anterior
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === 'LockedOldWeeks') {
      protections[i].remove();
    }
  }

  // 2. Se a ordem for trancar (botão ALL clicado)
  if (lockOldWeeks) {
    const maxCols = sheet.getLastColumn();
    if (maxCols > 20) { 
      const rangeToLock = sheet.getRange(1, 1, sheet.getMaxRows(), maxCols - 18);
      const protection = rangeToLock.protect().setDescription('LockedOldWeeks');
      protection.setWarningOnly(true);
    }
  }
}

// GERA O POP-UP DE CONFIRMAÇÃO NATIVO E LIMPO
function promptFullSync() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Rebuild Daten',
    'Are you sure you want to rebuild the entire history? This might take a few seconds.',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}