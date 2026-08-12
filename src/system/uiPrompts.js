function promptFullSync() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ Confirmar reconstrução da Daten',
    'Tem certeza de que deseja reconstruir todo o histórico?',
    ui.ButtonSet.YES_NO
  );
  return response === ui.Button.YES;
}