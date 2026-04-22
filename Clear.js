// ==========================================
// 🧠 FUNÇÃO AUXILIAR: CÉREBRO DE CORES (LEITURA DIRETA DA PALETA)
// ==========================================
function getActiveThemeColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paletteSheet = ss.getSheetByName("Palette Entry");

  // 1. Cores de segurança (Amarelo clássico)
  let colors = { side: '#ffe599', fill: '#fff2cc', text: '#000000' }; 

  // 2. TENTA LER AS CORES REAIS DIRETAMENTE DA PLANILHA (Igual o seu applyUIColorsOnly)
  if (paletteSheet) {
    const lastRow = paletteSheet.getLastRow();
    if (lastRow >= 1) {
      const pValues = paletteSheet.getRange(1, 1, lastRow, 17).getValues(); 
      for (let i = 0; i < pValues.length; i++) {
        const nameVal = pValues[i][14] ? pValues[i][14].toString().toLowerCase().trim() : ""; 
        
        if (nameVal.includes("pattern")) {
          let valP = pValues[i][15] ? pValues[i][15].toString().trim() : ""; // Coluna P (Side)
          let valQ = pValues[i][16] ? pValues[i][16].toString().trim() : ""; // Coluna Q (Fill)
          
          if (valP.startsWith("#")) colors.side = valP;
          if (valQ.startsWith("#")) colors.fill = valQ;
          break; // Achou as cores, pode parar a busca!
        }
      }
    }
  }

  // 3. DEFINE A COR DO TEXTO PARA NÃO SUMIR NO FUNDO
  try {
    const props = PropertiesService.getDocumentProperties();
    const stateStr = props.getProperty("CURRENT_UI_STATE");
    if (stateStr) {
      const state = JSON.parse(stateStr);
      if (state.ui && state.ui.timeText) {
        colors.text = state.ui.timeText;
      } else if (state.theme && state.theme.includes("DARK")) {
        colors.text = "#ffffff";
      }
    }
  } catch (e) {}

  return colors;
}

// ==========================================
// 🧹 VASSOURONA: RESET BLOCK (BLOQUEIO SILENCIOSO NO FRONTEND)
// ==========================================
function resetBlock() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Organisieren");
    if (!sheet) return;

    const activeRange = sheet.getActiveRange();
    let startRow = activeRange.getRow();
    let endRow = activeRange.getLastRow();
    let startCol = activeRange.getColumn();
    let endCol = activeRange.getLastColumn();

    // =========================================================
    // 🛡️ IDENTIFICAÇÃO DA SEMANA ATUAL (ÚLTIMA FÍSICA)
    // =========================================================
    const lastCol = sheet.getLastColumn();
    const currentWeekStart = 4 + Math.floor((lastCol - 4) / 18) * 18;
    const currentWeekEnd = currentWeekStart + 17;

    // 🕵️ DETECTOR DE FOCUS MODE
    // Se o início da última semana física estiver oculto, significa que o 
    // usuário rodou o focusWeek() para isolar a visão no passado.
    const isFocusMode = sheet.isColumnHiddenByUser(currentWeekStart);

    // =========================================================
    // 🛡️ ESCUDO ANTI-ESPERTINHO + TRAVA DE HISTÓRICO
    // =========================================================
    // A trava de coluna (não deixar editar o passado) só se aplica se NÃO estiver no Focus Mode.
    const isInvalidRow = (startRow < 9 || endRow > 104);
    const isInvalidCol = (!isFocusMode && (startCol < currentWeekStart || endCol > currentWeekEnd));

    if (isInvalidRow || isInvalidCol) {
      // Mostra o aviso bonito dentro da planilha
      SpreadsheetApp.getUi().alert(
        "🛑 Ação Bloqueada",
        "Você só pode resetar blocos dentro da SEMANA ATUAL!\n\nAs semanas passadas ficam travadas para preservar o seu histórico.\n\n(Dica: Use o 'Focus Mode' para liberar a edição de semanas anteriores).",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      // Retorna vazio para a barra lateral não exibir erro
      return; 
    }

    // =========================================================
    // 📏 AJUSTE DE MESCLAGEM E DNA DA COLUNA
    // =========================================================
    const mainColsOffsets = [2, 4, 6, 8, 10, 12, 14]; 
    const sideColsOffsets = [0, 1, 3, 5, 7, 9, 11, 13, 15, 16];
    
    let relStart = (startCol - 4) % 18;
    if (mainColsOffsets.includes(relStart)) {
      startCol = startCol - 1;
    }

    let minR = startRow;
    let maxR = endRow;
    for (let col = startCol; col <= endCol; col++) {
      let tempRange = sheet.getRange(startRow, col, endRow - startRow + 1, 1);
      let merges = tempRange.getMergedRanges();
      merges.forEach(m => {
        if (m.getRow() < minR) minR = m.getRow();
        if (m.getRow() + m.getNumRows() - 1 > maxR) maxR = m.getRow() + m.getNumRows() - 1;
      });
    }
    
    startRow = minR;
    endRow = maxR;
    let numRows = endRow - startRow + 1;

    // 1. Consulta o tema
    const theme = getActiveThemeColors();

    // 2. Loop de limpeza
    for (let col = startCol; col <= endCol; col++) {
      let relCol = (col - 4) % 18;
      if (relCol === 17) continue; // Pula a divisa

      let currentRange = sheet.getRange(startRow, col, numRows, 1);
      currentRange.clearContent();

      if (mainColsOffsets.includes(relCol)) {
        currentRange.setBackground(theme.fill)
                    .setFontColor(theme.text)
                    .setFontWeight('normal');
      } else if (sideColsOffsets.includes(relCol)) {
        currentRange.setBackground(theme.side);
      }
    }

    SpreadsheetApp.flush();
    return "SUCESSO";
  } catch (erro) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Erro interno: " + erro.message, "❌ Falha", 5);
    return;
  }
}

// ==========================================
// 🧹 VASSOURINHA: CLEAR LAST WEEK (VERSÃO RODAPÉ PROTEGIDO)
// ==========================================
function clearLastWeekContent() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    "🧹 Clear Week",
    "Do you want to KEEP the meal cells?\n\nYES = Clears tasks, but keeps meals.\nNO = Clears absolutely everything.",
    ui.ButtonSet.YES_NO_CANCEL
  );

  if (response === ui.Button.CANCEL || response === ui.Button.CLOSE) return; 
  const keepMeals = (response === ui.Button.YES);
  
  const blockWidth = 18;
  const dataStartCol = 4; 
  const currentMaxCols = sheet.getMaxColumns();
  let lastBlockCol = 0;
  
  for (let col = dataStartCol; col <= currentMaxCols; col += blockWidth) {
    if (col + blockWidth - 1 <= currentMaxCols) {
      const merges = sheet.getRange(1, col, 5, blockWidth).getMergedRanges();
      if (merges.length > 0) lastBlockCol = col;
    }
  }
  
  if (lastBlockCol === 0) {
    ui.alert("No blocks found.");
    return;
  }

  const cleanStartCol = lastBlockCol + 1; 
  
  // === AJUSTE CRÍTICO AQUI ===
  // Linha inicial: 9
  // Quantidade de linhas: 96 (Isso vai até a linha 104 exatamente)
  const numRows = 96; 
  const numCols = 14;
  const targetRange = sheet.getRange(9, cleanStartCol, numRows, numCols);
  
  if (keepMeals) {
    const values = targetRange.getValues();
    const mealKeywords = ["breakfast", "m. snack", "lunch", "a. snack", "l. snack", "dinner"];
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < numCols; c++) {
        const cellText = String(values[r][c]).toLowerCase();
        if (!mealKeywords.some(kw => cellText.includes(kw))) values[r][c] = ""; 
      }
    }
    targetRange.setValues(values);
  } else {
    targetRange.clearContent();
  }
  
  const theme = getActiveThemeColors();

  const rowColors = [];
  for (let c = 0; c < numCols; c++) {
    rowColors.push(c % 2 === 0 ? theme.side : theme.fill);
  }
  
  const backgroundPattern = [];
  for (let r = 0; r < numRows; r++) {
    backgroundPattern.push(rowColors);
  }
  targetRange.setBackgrounds(backgroundPattern);
  
  // Refaz as bordas apenas nas laterais, sem esmagar o rodapé
  targetRange.setBorder(null, true, null, true, false, null, theme.side, SpreadsheetApp.BorderStyle.SOLID);

  ui.alert("Cleanup Complete!", keepMeals ? "Tasks cleared. Your meals were kept!" : "The entire block was cleared from scratch.", ui.ButtonSet.OK);
}