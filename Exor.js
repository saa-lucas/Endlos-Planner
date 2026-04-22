function EXORCIZAR_FANTASMAS() {
  // Limpa o banco de dados sujo e recria a linha do tempo do zero (ignorando o futuro)
  updateHoursDashboard("FULL_SYNC");
  SpreadsheetApp.getActiveSpreadsheet().toast("Fantasmas exorcizados e painel cravado!", "Sucesso", 5);
}