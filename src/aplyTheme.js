// ==========================================
// 1. BANCO DE DADOS DE TEMAS (FÁBRICA)
// ==========================================
function getBaseThemes() {
  return {
    "DARK_CLASSIC": {
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff" },
      values: [
        ["#ff5f5f", "", "#2d0a0a", "", "#ffdbdb", "", "#d69ca5"], // C(2): Fundo | E(4): Texto
        ["#ff4b67", "", "#310d14", "", "#ffdae0", "", "#eeacb2"],
        ["#ff6d31", "", "#331500", "", "#ffeadb", "", "#f4bfb1"],
        ["#e07a3d", "", "#2d180c", "", "#f5e0d3", "", "#e5bba4"],
        ["#ffb347", "", "#3d2600", "", "#ffeed4", "", "#ffba66"],
        ["#ffcc00", "", "#3d3100", "", "#fff4cc", "", "#d39300"],
        ["#e1c03d", "", "#2d260c", "", "#f5eed3", "", "#dac157"],
        ["#966d5b", "", "#1f1511", "", "#e8ddd8", "", "#946b5c"],
        ["#a4e63e", "", "#212e0c", "", "#e9f7d4", "", "#a4cc6c"],
        ["#6dc131", "", "#152e0c", "", "#def0d4", "", "#85b44c"],
        ["#47d647", "", "#0c330c", "", "#d4f5d4", "", "#5ec75a"],
        ["#31a331", "", "#0c260c", "", "#d4e9d4", "", "#2d8e2d"],
        ["#3de1e1", "", "#0c2d2d", "", "#d3f5f5", "", "#32c8d1"],
        ["#3da3a3", "", "#0c2626", "", "#d3e9e9", "", "#218b8b"],
        ["#59b3f4", "", "#122633", "", "#dbedfc", "", "#4496a9"],
        ["#47d6ff", "", "#0c2d3d", "", "#d4f5ff", "", "#63d7f7"],
        ["#4b89ff", "", "#0d1a33", "", "#dae6ff", "", "#4178a3"],
        ["#a68aff", "", "#211a33", "", "#e9e1ff", "", "#a79cf1"],
        ["#c578ff", "", "#261833", "", "#efdcff", "", "#633e8f"],
        ["#ff66cc", "", "#331429", "", "#ffdaef", "", "#ef85c8"],
        ["#e65cb8", "", "#2e1225", "", "#f7daed", "", "#98246f"],
        ["#a8a19e", "", "#21201f", "", "#e9e7e6", "", "#afabaa"],
        ["#8a837e", "", "#1b1a19", "", "#e2e1e0", "", "#84817e"],
        ["#9eb3b8", "", "#1f2425", "", "#e7edef", "", "#a0aeb1"],
        ["#5c5c5c", "", "#1a1a1a", "", "#e0e0e0", "", "#686868"],
        ["#ffe599", "", "#fff2cc", "", "#fff2cc", "", "#fff2cc"]
      ]
    },
    "LIGHT_CLASSIC": {
      ui: { headerBg: "#f1f5f9", headerText: "#0f172a", timeBg: "#ffffff", timeText: "#334155" },
      values: [
        ["#ff4d4d", "", "#ffe6e6", "", "#800000", "", "#ff9999"],
        ["#d81b60", "", "#fce4ec", "", "#880e4f", "", "#f48fb1"],
        ["#e65100", "", "#fff3e0", "", "#bf360c", "", "#ffcc80"],
        ["#a1887f", "", "#efebe9", "", "#4e342e", "", "#d7ccc8"],
        ["#ffb74d", "", "#fff3e0", "", "#e65100", "", "#ffe0b2"],
        ["#ffb300", "", "#fff8e1", "", "#ff6f00", "", "#ffe082"],
        ["#fbc02d", "", "#fffde7", "", "#f57f17", "", "#fff59d"],
        ["#5d4037", "", "#d7ccc8", "", "#3e2723", "", "#8d6e63"],
        ["#7cb342", "", "#f1f8e9", "", "#33691e", "", "#c5e1a5"],
        ["#388e3c", "", "#e8f5e9", "", "#1b5e20", "", "#a5d6a7"],
        ["#81c784", "", "#e8f5e9", "", "#2e7d32", "", "#c8e6c9"],
        ["#2e7d32", "", "#c8e6c9", "", "#1b5e20", "", "#81c784"],
        ["#4dd0e1", "", "#e0f7fa", "", "#006064", "", "#b2ebf2"],
        ["#00897b", "", "#e0f2f1", "", "#004d40", "", "#80cbc4"],
        ["#546e7a", "", "#eceff1", "", "#263238", "", "#b0bec5"],
        ["#29b6f6", "", "#e1f5fe", "", "#01579b", "", "#81d4fa"],
        ["#1e88e5", "", "#e3f2fd", "", "#0d47a1", "", "#90caf9"],
        ["#9575cd", "", "#ede7f6", "", "#4527a0", "", "#d1c4e9"],
        ["#8e24aa", "", "#f3e5f5", "", "#4a148c", "", "#ce93d8"],
        ["#f06292", "", "#fce4ec", "", "#880e4f", "", "#f8bbd0"],
        ["#ab47bc", "", "#f3e5f5", "", "#4a148c", "", "#e1bee7"],
        ["#d7ccc8", "", "#f5f5f5", "", "#4e342e", "", "#efebe9"],
        ["#bcaaa4", "", "#efebe9", "", "#3e2723", "", "#d7ccc8"],
        ["#b0bec5", "", "#f5f5f5", "", "#37474f", "", "#cfd8dc"],
        ["#424242", "", "#eeeeee", "", "#212121", "", "#bdbdbd"],
        ["#e0e0e0", "", "#f9f9f9", "", "#212121", "", "#f9f9f9"]
      ]
    },
    "DARK_MODERN": {
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff" },
      values: [
        ["#ff5f5f", "", "#2d0a0a", "", "#ffdbdb", "", "#d69ca5"],
        ["#ff4b67", "", "#310d14", "", "#ffdae0", "", "#eeacb2"],
        ["#ff6d31", "", "#150a00", "", "#ffeadb", "", "#8c3d1f"],
        ["#e07a3d", "", "#1a0f08", "", "#f5e0d3", "", "#7d4525"],
        ["#ffd18a", "", "#1a140b", "", "#fff1db", "", "#8c6d3b"],
        ["#ffe066", "", "#1a1700", "", "#fff9e0", "", "#8c7b00"],
        ["#e1c03d", "", "#1a1705", "", "#f5eed3", "", "#8c7a26"],
        ["#966d5b", "", "#14100e", "", "#e8ddd8", "", "#5c4338"],
        ["#a4e63e", "", "#121706", "", "#e9f7d4", "", "#5a7a22"],
        ["#6dc131", "", "#0a1706", "", "#def0d4", "", "#3d641b"],
        ["#47d647", "", "#061a06", "", "#d4f5d4", "", "#2b7a2b"],
        ["#31a331", "", "#051405", "", "#d4e9d4", "", "#1d5c1d"],
        ["#3de1e1", "", "#061a1a", "", "#d3f5f5", "", "#267a7a"],
        ["#3da3a3", "", "#051414", "", "#d3e9e9", "", "#1b4d4d"],
        ["#59b3f4", "", "#0a131a", "", "#dbedfc", "", "#2d5a7a"],
        ["#47d6ff", "", "#061a21", "", "#d4f5ff", "", "#2b7a8c"],
        ["#4b89ff", "", "#070e1a", "", "#dae6ff", "", "#2b4d8c"],
        ["#a68aff", "", "#110e1a", "", "#e9e1ff", "", "#5b4d8c"],
        ["#c578ff", "", "#150d1a", "", "#efdcff", "", "#6d4d8c"],
        ["#ff66cc", "", "#1a0b14", "", "#ffdaef", "", "#8c3b6e"],
        ["#e65cb8", "", "#170912", "", "#f7daed", "", "#98246f"],
        ["#a8a19e", "", "#141414", "", "#e9e7e6", "", "#6b6766"],
        ["#8a837e", "", "#121212", "", "#e2e1e0", "", "#545250"],
        ["#9eb3b8", "", "#131718", "", "#e7edef", "", "#5c6b6e"],
        ["#5c5c5c", "", "#0d0d0d", "", "#e0e0e0", "", "#333333"],
        ["#333333", "", "#000000", "", "#666666", "", "#1a1a1a"]
      ]
    },
    "LIGHT_MODERN": {
      ui: { headerBg: "#f8fafc", headerText: "#0f172a", timeBg: "#ffffff", timeText: "#64748b" },
      values: [
        ["#ef5350", "", "#fef2f2", "", "#991b1b", "", "#fecaca"],
        ["#d81b60", "", "#fff1f2", "", "#831843", "", "#fbcfe8"],
        ["#f97316", "", "#fff7ed", "", "#7c2d12", "", "#ffedd5"],
        ["#a1887f", "", "#fafaf9", "", "#44403c", "", "#e7e5e4"],
        ["#fdba74", "", "#fffcf5", "", "#7c2d12", "", "#ffedd5"],
        ["#fbbf24", "", "#fffdf0", "", "#78350f", "", "#fef3c7"],
        ["#f59e0b", "", "#fffbeb", "", "#78350f", "", "#fef3c7"],
        ["#78350f", "", "#fdf8f6", "", "#451a03", "", "#f6e8e0"],
        ["#84cc16", "", "#f7fee7", "", "#365314", "", "#ecfccb"],
        ["#22c55e", "", "#f0fdf4", "", "#14532d", "", "#dcfce7"],
        ["#4ade80", "", "#f0fdf4", "", "#166534", "", "#bbf7d0"],
        ["#16a34a", "", "#f0fdf4", "", "#14532d", "", "#bbf7d0"],
        ["#06b6d4", "", "#ecfeff", "", "#164e63", "", "#cffafe"],
        ["#0891b2", "", "#ecfeff", "", "#164e63", "", "#cffafe"],
        ["#3b82f6", "", "#eff6ff", "", "#1e3a8a", "", "#dbeafe"],
        ["#0ea5e9", "", "#f0f9ff", "", "#0c4a6e", "", "#e0f2fe"],
        ["#2563eb", "", "#eff6ff", "", "#1e3a8a", "", "#dbeafe"],
        ["#a855f7", "", "#faf5ff", "", "#581c87", "", "#f3e8ff"],
        ["#9333ea", "", "#faf5ff", "", "#581c87", "", "#f3e8ff"],
        ["#ec4899", "", "#fdf2f8", "", "#831843", "", "#fce7f3"],
        ["#db2777", "", "#fdf2f8", "", "#831843", "", "#fce7f3"],
        ["#d6d3d1", "", "#fafaf9", "", "#44403c", "", "#e7e5e4"],
        ["#a8a29e", "", "#fafaf9", "", "#44403c", "", "#e7e5e4"],
        ["#94a3b8", "", "#f8fafc", "", "#1e293b", "", "#e2e8f0"],
        ["#475569", "", "#f1f5f9", "", "#0f172a", "", "#cbd5e1"],
        ["#e2e8f0", "", "#ffffff", "", "#1e293b", "", "#f1f5f9"]
      ]
    },
    "CUSTOM": {
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff" },
      values: [
        ["#ff5f5f", "", "#2d0a0a", "", "#ffdbdb", "", "#d69ca5"],
        ["#ff4b67", "", "#310d14", "", "#ffdae0", "", "#eeacb2"],
        ["#ff6d31", "", "#331500", "", "#ffeadb", "", "#f4bfb1"],
        ["#e07a3d", "", "#2d180c", "", "#f5e0d3", "", "#e5bba4"],
        ["#ffb347", "", "#3d2600", "", "#ffeed4", "", "#ffba66"],
        ["#ffcc00", "", "#3d3100", "", "#fff4cc", "", "#d39300"],
        ["#e1c03d", "", "#2d260c", "", "#f5eed3", "", "#dac157"],
        ["#966d5b", "", "#1f1511", "", "#e8ddd8", "", "#946b5c"],
        ["#a4e63e", "", "#212e0c", "", "#e9f7d4", "", "#a4cc6c"],
        ["#6dc131", "", "#152e0c", "", "#def0d4", "", "#85b44c"],
        ["#47d647", "", "#0c330c", "", "#d4f5d4", "", "#5ec75a"],
        ["#31a331", "", "#0c260c", "", "#d4e9d4", "", "#2d8e2d"],
        ["#3de1e1", "", "#0c2d2d", "", "#d3f5f5", "", "#32c8d1"],
        ["#3da3a3", "", "#0c2626", "", "#d3e9e9", "", "#218b8b"],
        ["#59b3f4", "", "#122633", "", "#dbedfc", "", "#4496a9"],
        ["#47d6ff", "", "#0c2d3d", "", "#d4f5ff", "", "#63d7f7"],
        ["#4b89ff", "", "#0d1a33", "", "#dae6ff", "", "#4178a3"],
        ["#a68aff", "", "#211a33", "", "#e9e1ff", "", "#a79cf1"],
        ["#c578ff", "", "#261833", "", "#efdcff", "", "#633e8f"],
        ["#ff66cc", "", "#331429", "", "#ffdaef", "", "#ef85c8"],
        ["#e65cb8", "", "#2e1225", "", "#f7daed", "", "#98246f"],
        ["#a8a19e", "", "#21201f", "", "#e9e7e6", "", "#afabaa"],
        ["#8a837e", "", "#1b1a19", "", "#e2e1e0", "", "#84817e"],
        ["#9eb3b8", "", "#1f2425", "", "#e7edef", "", "#a0aeb1"],
        ["#5c5c5c", "", "#1a1a1a", "", "#e0e0e0", "", "#686868"],
        ["#ffe599", "", "#fff2cc", "", "#fff2cc", "", "#fff2cc"]
      ]
    }
  };
}

/**
 * Retorna o tema atual para a Sidebar.
 * Assume que se o fundo da célula A1 for preto (#000000), o tema é DARK.
 */
function getSidebarTheme() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const bgColor = sheet.getRange("A1").getBackground();
  
  // Se o fundo da célula for preto, manda 'dark', senão 'light'
  return (bgColor === "#000000") ? "dark" : "light";
}

// ==========================================
// 2. RECUPERA O TEMA (AGORA ACEITA TEMAS NOVOS SEM CRASHAR)
// ==========================================
function getThemeData(themeName) {
  let base = getBaseThemes()[themeName];
  
  // Se não achou na fábrica, assume que é um tema novo criado por você.
  // Pega o esqueleto do "CUSTOM" original por precaução para não quebrar a planilha.
  if (!base) {
    base = JSON.parse(JSON.stringify(getBaseThemes()["CUSTOM"])); // Cópia segura
  }
  
  const customUIStr = PropertiesService.getDocumentProperties().getProperty(themeName + "_CUSTOM_UI");
  if (customUIStr) {
    const customUI = JSON.parse(customUIStr);
    base.ui = { ...base.ui, ...customUI }; 
  }
  return base;
}

// ==========================================
// FUNÇÃO NOVA: CRIAR E SALVAR TEMA NA LISTA
// ==========================================
function createNewCustomTheme(newThemeName, uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Criando seu novo tema...", "⏳ Aguarde", -1);
    const props = PropertiesService.getDocumentProperties();
    
    // 1. Cria uma chave única pro sistema não se perder (Ex: CUSTOM_MEU_TEMA_NEON)
    const themeKey = "CUSTOM_" + newThemeName.toUpperCase().replace(/\s+/g, "_");
    
    // 2. Salva as cores nesse slot novo
    props.setProperty(themeKey + "_CUSTOM_UI", JSON.stringify(uiConfig));
    
    // 3. Adiciona o nome na "Lista VIP" de temas salvos
    let savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    let savedThemes = savedThemesStr ? JSON.parse(savedThemesStr) : [];
    
    // Verifica se já existe pra não duplicar no menu
    let themeExists = savedThemes.find(t => t.key === themeKey);
    if (!themeExists) {
      savedThemes.push({ name: newThemeName, key: themeKey });
      props.setProperty("USER_SAVED_THEMES", JSON.stringify(savedThemes));
    }
    
    // 4. Aplica o tema na hora
    if (uiConfig.patternSide && uiConfig.patternFill) {
      updatePatternInPalette(uiConfig.patternSide, uiConfig.patternFill);
    }
    
    applyUIColorsOnly(uiConfig);
    updateThemeBorders(uiConfig); 
    props.setProperty("GLOBAL_THEME", themeKey); // Avisa que esse é o tema atual
    
    ss.toast("Tema '" + newThemeName + "' salvo com sucesso!", "✅ Sucesso", 5);
    return { status: "SUCESSO", themeKey: themeKey, themeName: newThemeName };
    
  } catch (e) {
    ss.toast("Erro ao criar tema: " + e.message, "❌ Erro", 5);
    return { status: "ERRO", message: e.message };
  }
}

// ==========================================
// FUNÇÃO NOVA: PUXAR A LISTA PARA O HTML
// ==========================================
function getUserSavedThemes() {
  const savedThemesStr = PropertiesService.getDocumentProperties().getProperty("USER_SAVED_THEMES");
  return savedThemesStr ? JSON.parse(savedThemesStr) : [];
}

// ==========================================
// 3. APLICA A COR NA INTERFACE
// ==========================================
function applyUIColorsOnly(uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orgSheet = ss.getSheetByName("Organisieren");
  const paletteSheet = ss.getSheetByName("Palette Entry"); 
  if (!orgSheet) return;

  const maxCol = orgSheet.getMaxColumns();
  const maxRow = orgSheet.getMaxRows();

  // HEADER BG: Linhas 1 a 8, começando da Coluna D (4)
  if (maxCol >= 4) {
    orgSheet.getRange(1, 4, 8, maxCol - 3).setBackground(uiConfig.headerBg).setFontColor(uiConfig.headerText);
  }

  // TIMELINE BG: Colunas A, B e C (1 a 3), até a linha 105
  const endRowTimeline = Math.min(maxRow, 105);
  if (maxCol >= 3 && endRowTimeline >= 1) {
    orgSheet.getRange(1, 1, endRowTimeline, 3).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
  }

  // FOOTER BG: Linha 105, começando da Coluna D (4)
  if (maxCol >= 4 && maxRow >= 105) {
    orgSheet.getRange(105, 4, 1, maxCol - 3).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
  }

  // DIVISAS BG: Coluna U (21) e seguintes (AM, BE...), pulando de 18 em 18
  if (maxRow >= 9 && maxCol >= 21) {
    const numRowsToPaint = Math.min(maxRow - 8, 96); // Cobre apenas de 9 a 104
    for (let c = 21; c <= maxCol; c += 18) {
      orgSheet.getRange(9, c, numRowsToPaint, 1).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
    }
  }

  // =======================================================
  // PINTAR FUNDO DAS COLUNAS SIDE (P) E PATTERN (Q)
  // =======================================================
  if (paletteSheet && maxRow >= 9) {
    const lastRow = paletteSheet.getLastRow();
    let sideColor = null;
    let patternColor = null;

    if (lastRow >= 1) {
      const pValues = paletteSheet.getRange(1, 1, lastRow, 17).getValues(); 
      for (let i = 0; i < pValues.length; i++) {
        // Puxa o nome da Coluna O, converte pra minúsculo e tira espaços para não ter erro
        const nameVal = pValues[i][14] ? pValues[i][14].toString().toLowerCase().trim() : ""; 
        
        // SÓ PEGA AS CORES SE ESTIVER NA LINHA DO "PATTERN"
        if (nameVal.includes("pattern")) {
          let valP = pValues[i][15] ? pValues[i][15].toString().trim() : ""; // Coluna P
          let valQ = pValues[i][16] ? pValues[i][16].toString().trim() : ""; // Coluna Q
          
          if (valP.startsWith("#")) sideColor = valP;
          if (valQ.startsWith("#")) patternColor = valQ;
          break; // Achou o pattern, não precisa ler mais nenhuma linha!
        }
      }
    }

    const numRowsPaint = Math.min(maxRow - 8, 96); // Da linha 9 a 104
    
    // Loop aplicando nas semanas (pula de 18 em 18 colunas a partir da D)
    for (let baseCol = 4; baseCol <= maxCol; baseCol += 18) {
      
      // APLICA COR DA SIDE (Coluna P) nas abas (E, G, I, K, M, O, Q, S)
      const sideOffsets = [1, 3, 5, 7, 9, 11, 13, 15]; 
      if (sideColor) {
        sideOffsets.forEach(offset => {
          const c = baseCol + offset;
          if (c <= maxCol) orgSheet.getRange(9, c, numRowsPaint, 1).setBackground(sideColor);
        });
      }

      // APLICA COR DO PATTERN (Coluna Q) nas vazias principais e extremidades
      // D(0), F(2), H(4), J(6), L(8), N(10), P(12), R(14), T(16)
      const patternOffsets = [0, 2, 4, 6, 8, 10, 12, 14, 16]; 
      if (patternColor) {
        patternOffsets.forEach(offset => {
          const c = baseCol + offset;
          if (c <= maxCol) orgSheet.getRange(9, c, numRowsPaint, 1).setBackground(patternColor);
        });
      }
    }
  }

  SpreadsheetApp.flush();
}

// ==========================================
// 4. SALVA CUSTOMIZAÇÃO DA SIDEBAR
// ==========================================
function saveCustomThemeUI(themeName, uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Saving custom theme and applying colors...", "⏳ Personalizando", -1);
    
    PropertiesService.getDocumentProperties().setProperty(themeName + "_CUSTOM_UI", JSON.stringify(uiConfig));
    
    // INJETA AS CORES DO PATTERN NA FOLHA ANTES DE PINTAR
    if (uiConfig.patternSide && uiConfig.patternFill) {
      updatePatternInPalette(uiConfig.patternSide, uiConfig.patternFill);
    }
    
    applyUIColorsOnly(uiConfig);
    updateThemeBorders(uiConfig); 
    
    ss.toast("Tema personalizado salvo e aplicado!", "✅ Sucesso", 3);
    return "SUCESSO";
  } catch (e) {
    ss.toast("Erro ao salvar o tema personalizado: " + e.message, "❌ Erro", 5);
    return "ERRO: " + e.message;
  }
}

// ==========================================
// 5. MOTOR PRINCIPAL: TROCAR DE TEMA
// ==========================================
function applyTheme(themeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Aguarde, aplicando novas cores e bordas...", "⏳ Alterando tema", -1);
    const sheet = ss.getSheetByName("Palette Entry");
    if (!sheet) return "ERRO: Aba Palette Entry não encontrada.";

    PropertiesService.getDocumentProperties().setProperty("GLOBAL_THEME", themeName);
    const themeData = getThemeData(themeName);
    if (!themeData) return "ERRO: Tema inexistente.";

    // Escreve a Matriz Base
    sheet.getRange(2, 1, themeData.values.length, 7).setValues(themeData.values);
    
    // MÁGICA AQUI: Força a atualização do Pattern. Se o tema não tiver, envia "" para limpar as células!
    const pSide = (themeData.ui && themeData.ui.patternSide) ? themeData.ui.patternSide : "";
    const pFill = (themeData.ui && themeData.ui.patternFill) ? themeData.ui.patternFill : "";
    updatePatternInPalette(pSide, pFill);

    applyUIColorsOnly(themeData.ui);
    updateThemeBorders(themeData.ui); 
    
    if (typeof applyOutlookColorsOptimized === "function") applyOutlookColorsOptimized();
    if (typeof applyFillFromOutlookColorsOptimized === "function") applyFillFromOutlookColorsOptimized();
    
    ss.toast("Tema aplicado com sucesso!", "✅ Sucesso", 3);
    return JSON.stringify({ status: "SUCESSO", ui: themeData.ui }); 
  } catch (erro) {
    ss.toast("Falha ao aplicar o tema: " + erro.message, "❌ Erro crítico", 5);
    return "ERRO CRÍTICO: " + erro.message;
  }
}

// ==========================================
// 6. PUXA A COR ATUAL 
// ==========================================
function getCurrentUIState() {
  const currentTheme = PropertiesService.getDocumentProperties().getProperty("GLOBAL_THEME") || "DARK_THEME";
  const themeData = getThemeData(currentTheme);
  return { theme: currentTheme, ui: themeData ? themeData.ui : null };
}

// ==========================================
// 7. BORDAS INTELIGENTES (VERSÃO CORRIGIDA PARA O RODAPÉ)
// ==========================================
function updateThemeBorders(uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orgSheet = ss.getSheetByName("Organisieren");
  const paletteSheet = ss.getSheetByName("Palette Entry");
  if (!orgSheet || !paletteSheet) return;

  // 1. BUSCA A COR DO STROKE E CORES DO TEMA ATUAL
  const theme = getActiveThemeColors(); // Puxa nosso "Cérebro de Cores"
  const lastRow = paletteSheet.getLastRow();
  let strokeColor = "#333333"; 
  
  if (lastRow >= 1) {
    const values = paletteSheet.getRange(1, 1, lastRow, 12).getValues(); 
    for (let i = 0; i < values.length; i++) {
      const nameVal = values[i][11];
      if (nameVal && nameVal.toString().toLowerCase().includes("pattern")) { 
        const strokeVal = values[i][6];
        if (strokeVal && strokeVal.toString().startsWith("#")) {
          strokeColor = strokeVal.toString().trim();
          break; 
        }
      }
    }
  }

  const maxCol = orgSheet.getMaxColumns();
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

  // =========================================================
  // PASSO 1: RESET E FUNDO DA LINHA 105 (Rodapé)
  // =========================================================
  if (maxCol >= 4) {
    const footerRange = orgSheet.getRange(105, 4, 1, maxCol - 3);
    footerRange.setBorder(false, false, false, false, false, false);
    // Aplica a cor do tema no fundo para não ficar branco/vazio
    footerRange.setBackground(theme.side); 
  }

  // =========================================================
  // PASSO 2: TIMELINE COL C
  // =========================================================
  if (maxCol >= 3) {
    orgSheet.getRange(9, 3, 95, 1).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);
    orgSheet.getRange(8, 3).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
    
    orgSheet.getRange(104, 3).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
    orgSheet.getRange(105, 3).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
  }

  // =========================================================
  // PASSO 3: LOOP POR SEMANAS
  // =========================================================
  for (let baseCol = 4; baseCol <= maxCol; baseCol += 18) {
    
    // Grade central do miolo (Aumentado para 96 para tocar a linha 104)
    orgSheet.getRange(9, baseCol, 96, 17).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);

    // Divisa Col U
    const divisa = baseCol + 17;
    if (divisa <= maxCol) {
      orgSheet.getRange(9, divisa, 96, 1).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);
      orgSheet.getRange(8, divisa).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      
      orgSheet.getRange(104, divisa).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      orgSheet.getRange(105, divisa).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
    }

    const patternColsOffsets = [0, 1, 3, 5, 7, 9, 11, 13, 15, 16]; 
    const mainColsOffsets = [2, 4, 6, 8, 10, 12, 14];             
    
    // TOPO (LINHA 8) E RODAPÉ (LINHAS 104/105)
    patternColsOffsets.forEach(offset => {
      const col = baseCol + offset;
      if (col <= maxCol) {
        orgSheet.getRange(8, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
        
        // REFORÇO NO RODAPÉ
        orgSheet.getRange(104, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
        orgSheet.getRange(105, col).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
      }
    });

    mainColsOffsets.forEach(offset => {
      const col = baseCol + offset;
      if (col <= maxCol) {
        orgSheet.getRange(8, col).setBorder(false, false, true, false, false, false, uiConfig.headerBg, borderStyle);
        
        // FECHAMENTO DA COLUNA LARGA NO RODAPÉ (Garante a linha horizontal de baixo)
        orgSheet.getRange(104, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      }
    });
  }

  SpreadsheetApp.flush();
}

// ==========================================
// 8. ATUALIZA O PATTERN NA MATRIZ RAIZ (PRESERVA FÓRMULAS E IGUALA FILL = STROKE)
// ==========================================
function updatePatternInPalette(sideColor, fillColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paletteSheet = ss.getSheetByName("Palette Entry");
  
  if (!paletteSheet) return;
  
  const lastRow = paletteSheet.getLastRow();
  if (lastRow < 1) return;
  
  // Lê APENAS a coluna L (índice 12), que é onde estão os nomes ("Pattern", etc)
  const labels = paletteSheet.getRange(1, 12, lastRow, 1).getValues(); 
  
  // 🧠 LÓGICA DE CORES DO TEXTO
  const currentTheme = PropertiesService.getDocumentProperties().getProperty("GLOBAL_THEME") || "";
  const isDark = currentTheme.includes("DARK");
  const textColor = isDark ? "#ffffff" : "#000000"; // Branco no tema escuro, Preto no claro

  for (let i = 0; i < labels.length; i++) {
    const nameVal = labels[i][0] ? labels[i][0].toString().toLowerCase().trim() : ""; 
    
    if (nameVal.includes("pattern")) {
      const row = i + 1; // Linha exata onde o "Pattern" está na matriz
      
      // INJETA NA MATRIZ BASE (A a G) - Suas fórmulas vão ler daqui!
      if (sideColor) {
        paletteSheet.getRange(row, 1).setValue(sideColor); // Coluna A (Side)
      }
      
      if (fillColor) {
        paletteSheet.getRange(row, 3).setValue(fillColor); // Coluna C (Fill)
        
        // ✨ A REGRA DE OURO: Stroke recebe exatamente o mesmo HEX do Fill!
        paletteSheet.getRange(row, 7).setValue(fillColor); // Coluna G (Stroke)
      }
      
      // Injeta o Texto
      paletteSheet.getRange(row, 5).setValue(textColor);   // Coluna E (Texto)
      
      break; // Achou e atualizou. Fim do loop!
    }
  }
  
  // Força o Google Sheets a calcular as fórmulas =SEERRO(...) antes do pincel passar
  SpreadsheetApp.flush();
}

// ==========================================
// FUNÇÃO NOVA: CRIAR TEMA AUTOMÁTICO (TEMA 1, TEMA 2...)
// ==========================================
function createNewAutoTheme() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const props = PropertiesService.getDocumentProperties();
    
    // Puxa a lista atual para saber qual é o número do próximo tema
    let savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    let savedThemes = savedThemesStr ? JSON.parse(savedThemesStr) : [];
    
    let nextNum = savedThemes.length + 1;
    let newThemeName = "TEMA " + nextNum;
    let themeKey = "CUSTOM_TEMA_" + nextNum;
    
    // Copia a paleta do DARK_MODERN como base para o novo tema não nascer "em branco"
    let base = JSON.parse(JSON.stringify(getBaseThemes()["DARK_MODERN"]));
    let newUi = base.ui;
    
    // Salva a base do tema novo
    props.setProperty(themeKey + "_CUSTOM_UI", JSON.stringify(newUi));
    
    // Põe na lista do Menu
    savedThemes.push({ name: newThemeName, key: themeKey });
    props.setProperty("USER_SAVED_THEMES", JSON.stringify(savedThemes));
    
    // Define como tema ativo e pinta a tela
    props.setProperty("GLOBAL_THEME", themeKey);
    if (newUi.patternSide && newUi.patternFill) {
      updatePatternInPalette(newUi.patternSide, newUi.patternFill);
    }
    applyUIColorsOnly(newUi);
    updateThemeBorders(newUi); 
    
    return { status: "SUCESSO", themeKey: themeKey, themeName: newThemeName, ui: newUi };
    
  } catch (e) {
    ss.toast("Erro ao criar tema: " + e.message, "❌ Erro", 5);
    return { status: "ERRO", message: e.message };
  }
}

// ==========================================
// FUNÇÃO NOVA: PREVIEW AO VIVO DAS CORES
// ==========================================
function previewThemeColors(uiConfig) {
  // Apenas pinta a tela, NÃO salva na memória permanente.
  if (uiConfig.patternSide && uiConfig.patternFill) {
    updatePatternInPalette(uiConfig.patternSide, uiConfig.patternFill);
  }
  applyUIColorsOnly(uiConfig);
  updateThemeBorders(uiConfig);
}

// ==========================================
// 🧠 FUNÇÃO NOVA: MATEMÁTICA DE COR PARA A BORDA (STROKE)
// ==========================================
function getSmartStrokeColor(hex, isDark) {
  // Garante que é um HEX válido de 6 caracteres
  if (!/^#[0-9A-F]{6}$/i.test(hex)) return hex;
  
  let r = parseInt(hex.slice(1, 3), 16);
  let g = parseInt(hex.slice(3, 5), 16);
  let b = parseInt(hex.slice(5, 7), 16);

  const amount = 35; // Intensidade da diferença da borda (0 a 255)
  
  if (isDark) {
    // Tema escuro: Clareia a borda para destacar
    r = Math.min(255, r + amount);
    g = Math.min(255, g + amount);
    b = Math.min(255, b + amount);
  } else {
    // Tema claro: Escurece a borda para destacar
    r = Math.max(0, r - amount);
    g = Math.max(0, g - amount);
    b = Math.max(0, b - amount);
  }

  const toHex = (n) => {
    let h = n.toString(16);
    return h.length === 1 ? "0" + h : h;
  };
  
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

// ==========================================
// 8. ATUALIZA O PATTERN NA MATRIZ RAIZ (COM STROKE INTELIGENTE)
// ==========================================
function updatePatternInPalette(sideColor, fillColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paletteSheet = ss.getSheetByName("Palette Entry");
  
  if (!paletteSheet) return;
  
  const lastRow = paletteSheet.getLastRow();
  if (lastRow < 1) return;
  
  const labels = paletteSheet.getRange(1, 12, lastRow, 1).getValues(); 
  
  const currentTheme = PropertiesService.getDocumentProperties().getProperty("GLOBAL_THEME") || "";
  const isDark = currentTheme.includes("DARK");
  const textColor = isDark ? "#ffffff" : "#000000"; 

  // ✨ GERA A COR DE BORDA AUTOMÁTICA
  const smartStroke = getSmartStrokeColor(fillColor, isDark);

  for (let i = 0; i < labels.length; i++) {
    const nameVal = labels[i][0] ? labels[i][0].toString().toLowerCase().trim() : ""; 
    
    if (nameVal.includes("pattern")) {
      const row = i + 1; 
      
      if (sideColor) paletteSheet.getRange(row, 1).setValue(sideColor); 
      if (fillColor) {
        paletteSheet.getRange(row, 3).setValue(fillColor); 
        paletteSheet.getRange(row, 7).setValue(smartStroke); // Aplica a borda gerada!
      }
      
      paletteSheet.getRange(row, 5).setValue(textColor);   
      break; 
    }
  }
  
  SpreadsheetApp.flush();
}

// ==========================================
// FUNÇÃO NOVA: DELETAR TEMA CUSTOMIZADO
// ==========================================
function deleteCustomTheme(themeKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ss.toast("Apagando tema...", "🗑️ Deletando", -1);
    const props = PropertiesService.getDocumentProperties();
    
    // 1. Deleta a configuração de cores
    props.deleteProperty(themeKey + "_CUSTOM_UI");
    
    // 2. Remove o tema da lista VIP
    let savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    let savedThemes = savedThemesStr ? JSON.parse(savedThemesStr) : [];
    savedThemes = savedThemes.filter(t => t.key !== themeKey);
    props.setProperty("USER_SAVED_THEMES", JSON.stringify(savedThemes));
    
    // 3. Se ele tava usando o tema que foi apagado, joga pro padrão
    let current = props.getProperty("GLOBAL_THEME");
    let fallback = "DARK_MODERN";
    if (current === themeKey) {
      props.setProperty("GLOBAL_THEME", fallback);
    } else {
      fallback = current; // Mantém onde tava
    }
    
    ss.toast("Tema deletado com sucesso!", "✅ Sucesso", 4);
    return { status: "SUCESSO", fallbackTheme: fallback };
    
  } catch (e) {
    ss.toast("Erro ao deletar tema: " + e.message, "❌ Erro", 5);
    return { status: "ERRO", message: e.message };
  }
}