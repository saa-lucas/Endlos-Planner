function resetAllThemeMemory() {
  const ps = PropertiesService.getDocumentProperties();
  
  // Limpa as "sombras" dos temas padrão e do custom
  ps.deleteProperty("LIGHT_THEME_CUSTOM_UI");
  ps.deleteProperty("DARK_THEME_CUSTOM_UI");
  ps.deleteProperty("CUSTOM_CUSTOM_UI");
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Memória limpa! Agora o código manda.", "🧹 Reset Concluído", 5);
}