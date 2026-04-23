function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Overview')
      .addItem('Abrir Seletor de Semanas', 'openSidebar')
      .addSeparator()
      .addItem('Mostrar Tudo (Reset)', 'showAllColumns')
      .addToUi();
      
  ui.createMenu('⚡ Organisieren')
      .addItem('✂️ Mesclar / Dividir Bloco', 'toggleSmartMerge')
      .addSeparator()
      .addItem('📈 Abrir Analytics', 'openDashboardModal')
      .addToUi();

  openSidebar();
}