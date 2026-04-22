/**
 * Creates the custom menus when the spreadsheet opens.
 * The menu labels are in Portuguese for better local UX.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1º MENU: O seu seletor de semanas original
  ui.createMenu('Overview')
      .addItem('Abrir Seletor de Semanas', 'openSidebar')
      .addSeparator()
      .addItem('Mostrar Tudo (Reset)', 'showAllColumns')
      // .addItem('Limpar Semanas Antigas', 'clearOldWeeks')
      .addToUi();
      
  // 2º MENU: As nossas ferramentas da rotina
  ui.createMenu('⚡ Organisieren')
      .addItem('✂️ Mesclar / Dividir Bloco', 'toggleSmartMerge')
      .addSeparator()
      .addItem('📈 Abrir Analytics', 'openDashboardModal')
      .addToUi();

  // 🔥 A MÁGICA ACONTECE AQUI: 
  // Chama a função que abre a aba lateral automaticamente ao carregar a página!
  openSidebar();
}

/**
 * Opens the sidebar with the week selector.
 */
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Schedule Selector')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns an array of named ranges starting with "Week", 
 * mathematically ordered from newest (right) to oldest (left).
 */
function getWeekNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  let weeks = [];

  // 1. Puxa todos os nomes que começam com "Week" e anota em qual coluna eles estão
  for (let i = 0; i < namedRanges.length; i++) {
    let nr = namedRanges[i];
    let name = nr.getName();
    
    if (name.startsWith("Week")) {
      weeks.push({
        name: name,
        col: nr.getRange().getColumn() // Pega a posição real na planilha
      });
    }
  }

  // 2. ORDENAÇÃO MATEMÁTICA: Da coluna MAIOR para a MENOR
  weeks.sort((a, b) => b.col - a.col);

  // 3. Devolve só a lista de nomes já ordenadinha para o HTML
  return weeks.map(w => w.name);
}

/**
 * Focuses on a specific week, keeping time columns A:C visible.
 * @param {string} rangeName
 */
function focusWeek(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const targetRange = ss.getRangeByName(rangeName);
  
  if (!targetRange) return;

  // 1. Reset visibility to recalculate
  const maxCols = sheet.getMaxColumns();
  sheet.showColumns(1, maxCols);
  
  // 2. Define boundaries
  const startCol = targetRange.getColumn();
  const numCols = targetRange.getNumColumns();
  const endCol = startCol + numCols - 1;

  // 3. Hide gap between Column C (3) and the start of the week
  if (startCol > 4) {
    sheet.hideColumns(4, startCol - 4);
  }

  // 4. Hide columns after the target week
  if (endCol < maxCols) {
    sheet.hideColumns(endCol + 1, maxCols - endCol);
  }
  
  // 5. Ensure time columns A:C are always visible
  sheet.showColumns(1, 3);
  
  targetRange.activate();
}

/**
 * Reveals all columns in the active sheet.
 */
function showAllColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.showColumns(1, sheet.getMaxColumns());
}

/**
 * Deletes all named ranges that start with "Week".
 */
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