// =====================================================
// OVERVIEW - NAVIGATION LAYER (SIDEBAR + WEEK CONTROL)
// =====================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // MENU PRINCIPAL - NAVIGATION
  ui.createMenu('Overview')
    .addItem('Abrir Seletor de Semanas', 'openSidebar')
    .addSeparator()
    .addItem('Mostrar Tudo (Reset)', 'showAllColumns')
    .addToUi();

  // MENU DO SISTEMA (delegado)
  ui.createMenu('⚡ Organisieren')
    .addItem('✂️ Mesclar / Dividir Bloco', 'toggleSmartMerge')
    .addSeparator()
    .addItem('📈 Abrir Analytics', 'openDashboardModal')
    .addToUi();

  openSidebar();
}


// =====================================================
// SIDEBAR
// =====================================================

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui/sidebar')
    .setTitle('Schedule Selector')
    .setWidth(300);

  SpreadsheetApp.getUi().showSidebar(html);
}


// =====================================================
// WEEK SYSTEM
// =====================================================

function getWeekNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();

  const weeks = [];

  namedRanges.forEach(nr => {
    const name = nr.getName();

    if (name.startsWith("Week")) {
      weeks.push({
        name,
        col: nr.getRange().getColumn()
      });
    }
  });

  weeks.sort((a, b) => b.col - a.col);

  return weeks.map(w => w.name);
}


// =====================================================
// FOCUS SYSTEM
// =====================================================

function focusWeek(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const targetRange = ss.getRangeByName(rangeName);

  if (!targetRange) return;

  const maxCols = sheet.getMaxColumns();
  sheet.showColumns(1, maxCols);

  const startCol = targetRange.getColumn();
  const endCol = startCol + targetRange.getNumColumns() - 1;

  if (startCol > 4) {
    sheet.hideColumns(4, startCol - 4);
  }

  if (endCol < maxCols) {
    sheet.hideColumns(endCol + 1, maxCols - endCol);
  }

  sheet.showColumns(1, 3);
  targetRange.activate();
}


// =====================================================
// UTILS
// =====================================================

function showAllColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.showColumns(1, sheet.getMaxColumns());
}