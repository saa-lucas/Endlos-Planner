function promptFullSync() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirm Rebuild Daten',
    'Are you sure you want to rebuild the entire history?',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}