// ==========================================
// 1. BANCO DE DADOS DE TEMAS (FÁBRICA)
// ==========================================
function getBaseThemes() {
  return {
    "DARK_CLASSIC": {
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff", blockBg: "#000000", borderColor: "#333333", patternSide: "#111111", footerBg: "#000000" },
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
    },
    "LIGHT_CLASSIC": {
      ui: { headerBg: "#f1f5f9", headerText: "#0f172a", timeBg: "#ffffff", timeText: "#334155", blockBg: "#ffffff", borderColor: "#cbd5e1", patternSide: "#f8fafc", footerBg: "#f1f5f9" },
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
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff", blockBg: "#000000", borderColor: "#262626", patternSide: "#111111", footerBg: "#000000" },
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
      ui: { headerBg: "#f8fafc", headerText: "#0f172a", timeBg: "#ffffff", timeText: "#64748b", blockBg: "#ffffff", borderColor: "#e2e8f0", patternSide: "#f1f5f9", footerBg: "#e2e8f0" },
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
      ui: { headerBg: "#000000", headerText: "#ffffff", timeBg: "#000000", timeText: "#ffffff", blockBg: "#000000", borderColor: "#333333", patternSide: "#111111", footerBg: "#000000" },
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

// ==========================================
// 🧠 MOTOR DE LUMINÂNCIA (Calcula Hierarquia Visual Robusta)
// ==========================================
function hexToRgb(hex) {
  hex = hex.replace(/^#/, '');
  if (hex.length === 3) hex = hex.split('').map(c => c + c).join('');
  const num = parseInt(hex, 16);
  return [num >> 16, (num >> 8) & 255, num & 255];
}

function rgbToHex(r, g, b) {
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
}

function mixColors(color1, color2, weight) {
  const rgb1 = hexToRgb(color1);
  const rgb2 = hexToRgb(color2);
  const w1 = weight;
  const w2 = 1 - weight;
  const r = Math.round(rgb1[0] * w1 + rgb2[0] * w2);
  const g = Math.round(rgb1[1] * w1 + rgb2[1] * w2);
  const b = Math.round(rgb1[2] * w1 + rgb2[2] * w2);
  return rgbToHex(r, g, b);
}

function getThemeAppearance(blockBgHex) {
  let bg = blockBgHex;
  if (!bg || !/^#[0-9A-F]{3,6}$/i.test(bg)) { bg = "#000000"; }
  
  const [r, g, b] = hexToRgb(bg);
  
  const a = [r, g, b].map(function (v) {
      v /= 255;
      return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  });
  const luminance = a[0] * 0.2126 + a[1] * 0.7152 + a[2] * 0.0722;
  
  const isDark = luminance < 0.35; 
  const targetMix = isDark ? "#FFFFFF" : "#000000";

  const success = isDark ? "#00E676" : "#008B46"; 
  const danger = isDark ? "#FF4A70" : "#D31B42";  
  const warning = isDark ? "#FFC107" : "#E65100"; 
  const accent = isDark ? "#4DA8DA" : "#0056B3";  

  return {
      background: bg,
      foreground: isDark ? "#FFFFFF" : "#000000",
      foregroundMuted: mixColors(isDark ? "#FFFFFF" : "#000000", bg, 0.65),
      bgCard: mixColors(targetMix, bg, 0.08),
      bgBtn: mixColors(targetMix, bg, 0.15),
      bgBtnHover: mixColors(targetMix, bg, 0.25),
      border: mixColors(targetMix, bg, 0.20),
      borderHover: mixColors(targetMix, bg, 0.40),
      accent: accent,
      success: success,
      warning: warning,
      danger: danger,
      isDark: isDark
  };
}

// ==========================================
// 2. RECUPERA O TEMA (BLINDADO)
// ==========================================
function getThemeData(themeName) {
  let baseRaw = getBaseThemes()[themeName];
  if (!baseRaw) baseRaw = getBaseThemes()["CUSTOM"]; 
  let base = JSON.parse(JSON.stringify(baseRaw)); 
  
  if (themeName && themeName.startsWith("CUSTOM_")) {
    const customUIStr = PropertiesService.getDocumentProperties().getProperty(themeName + "_CUSTOM_UI");
    if (customUIStr) {
      const customUI = JSON.parse(customUIStr);
      base.ui = { ...base.ui, ...customUI }; 
    }
  }
  
  if (!base.ui.blockBg) {
    base.ui.blockBg = "#000000"; 
  }
  return base;
}

// ==========================================
// SALVAR PERSISTÊNCIA GERAL DA UI
// ==========================================
function saveGlobalUIPreferences(theme, sidebarMode, dashboardMode) {
  try {
    const props = PropertiesService.getDocumentProperties();
    if (theme) props.setProperty("GLOBAL_THEME", theme);
    if (sidebarMode) props.setProperty("SIDEBAR_MODE", sidebarMode);
    if (dashboardMode) props.setProperty("DASHBOARD_MODE", dashboardMode);
    return "OK";
  } catch (e) {
    throw new Error(e.message);
  }
}

// ==========================================
// 6. PUXA O ESTADO ATUAL (CONTEXTO COMPLETO E SEGURO)
// ==========================================
function getCurrentUIState() {
  try {
    var props = PropertiesService.getDocumentProperties();
    var currentTheme = props.getProperty("GLOBAL_THEME");
    var sidebarMode = props.getProperty("SIDEBAR_MODE") || "THEME";
    var dashboardMode = props.getProperty("DASHBOARD_MODE") || "THEME";
    
    if (!currentTheme) {
      currentTheme = "DARK_MODERN";
      props.setProperty("GLOBAL_THEME", currentTheme);
    }
    
    var themeData = getThemeData(currentTheme);
    var appearance = getThemeAppearance(themeData.ui.blockBg);
    
    console.log("STATE LIDO - Theme: " + currentTheme + " | Side: " + sidebarMode + " | Dash: " + dashboardMode);

    return { 
        theme: currentTheme, 
        ui: themeData.ui,
        sidebarMode: sidebarMode,
        dashboardMode: dashboardMode,
        appearance: appearance
    };
  } catch(e) {
    console.log("ERRO FATAL AO LER TEMA ATUAL: " + e.message);
    var fallbackApp = getThemeAppearance("#000000");
    return { theme: "DARK_MODERN", ui: null, sidebarMode: "THEME", dashboardMode: "THEME", appearance: fallbackApp }; 
  }
}

// ==========================================
// FUNÇÃO: CRIAR TEMA AUTOMÁTICO (SOMENTE CRIA E RETORNA)
// ==========================================
function createNewAutoTheme() {
  const props = PropertiesService.getDocumentProperties();
  try {
    let savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    let savedThemes = savedThemesStr ? JSON.parse(savedThemesStr) : [];

    let nextNum = 1;
    while (savedThemes.some(t => t.key === "CUSTOM_TEMA_" + nextNum)) {
      nextNum++;
    }

    const themeName = "TEMA " + nextNum;
    const themeKey = "CUSTOM_TEMA_" + nextNum;

    const base = JSON.parse(JSON.stringify(getBaseThemes()["DARK_MODERN"]));
    const uiConfig = base.ui;

    props.setProperty(themeKey + "_CUSTOM_UI", JSON.stringify(uiConfig));

    savedThemes.push({ name: themeName, key: themeKey });
    props.setProperty("USER_SAVED_THEMES", JSON.stringify(savedThemes));

    props.setProperty("GLOBAL_THEME", themeKey);

    const appearance = getThemeAppearance(uiConfig.blockBg);

    return {
      status: "SUCESSO",
      themeKey: themeKey,
      themeName: themeName,
      ui: uiConfig,
      appearance: appearance
    };

  } catch (e) {
    console.error("ERRO AO CRIAR TEMA: " + e.message);
    return { status: "ERRO", message: e.message };
  }
}

// ==========================================
// FUNÇÃO: PUXAR A LISTA PARA O HTML
// ==========================================
function getUserSavedThemes() {
  try {
    var props = PropertiesService.getDocumentProperties();
    var savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    if (!savedThemesStr) return [];
    return JSON.parse(savedThemesStr);
  } catch(e) {
    console.log("ERRO AO LER LISTA DE TEMAS: " + e.message);
    return [];
  }
}

// ==========================================
// 3. APLICA A COR NA INTERFACE (PLANILHA)
// ==========================================
function applyUIColorsOnly(uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orgSheet = ss.getSheetByName("Organisieren");
  if (!orgSheet || !uiConfig) return;

  const maxCol = orgSheet.getMaxColumns();
  const maxRow = orgSheet.getMaxRows();

  if (maxCol >= 4) {
    orgSheet.getRange(1, 4, 8, maxCol - 3).setBackground(uiConfig.headerBg).setFontColor(uiConfig.headerText);
  }

  const endRowTimeline = Math.min(maxRow, 105);
  if (maxCol >= 3 && endRowTimeline >= 1) {
    orgSheet.getRange(1, 1, endRowTimeline, 3).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
  }

  if (maxCol >= 4 && maxRow >= 105) {
    orgSheet.getRange(105, 4, 1, maxCol - 3).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
  }

  if (maxRow >= 9 && maxCol >= 21) {
    const numRowsToPaint = Math.min(maxRow - 8, 96); 
    let divisasA1 = [];
    for (let c = 21; c <= maxCol; c += 18) {
      divisasA1.push(orgSheet.getRange(9, c, numRowsToPaint, 1).getA1Notation());
    }
    if (divisasA1.length > 0) {
      orgSheet.getRangeList(divisasA1).setBackground(uiConfig.timeBg).setFontColor(uiConfig.timeText);
    }
  }

  if (maxRow >= 9) {
    const numRowsPaint = Math.min(maxRow - 8, 96); 
    let sideA1List = [];
    let patternA1List = [];
    
    for (let baseCol = 4; baseCol <= maxCol; baseCol += 18) {
      const sideOffsets = [1, 3, 5, 7, 9, 11, 13, 15]; 
      sideOffsets.forEach(offset => {
        if (baseCol + offset <= maxCol) sideA1List.push(orgSheet.getRange(9, baseCol + offset, numRowsPaint, 1).getA1Notation());
      });

      const patternOffsets = [0, 2, 4, 6, 8, 10, 12, 14, 16]; 
      patternOffsets.forEach(offset => {
        if (baseCol + offset <= maxCol) patternA1List.push(orgSheet.getRange(9, baseCol + offset, numRowsPaint, 1).getA1Notation());
      });
    }

    if (uiConfig.patternSide && sideA1List.length > 0) orgSheet.getRangeList(sideA1List).setBackground(uiConfig.patternSide);
    if (uiConfig.blockBg && patternA1List.length > 0) orgSheet.getRangeList(patternA1List).setBackground(uiConfig.blockBg);
  }

  SpreadsheetApp.flush();
}

// ==========================================
// 4. SALVAR EDIÇÃO DE TEMA (APENAS PERSISTE!)
// ==========================================
function saveCustomThemeUI(themeName, uiConfig) {
  try {
    if (!themeName || !themeName.startsWith("CUSTOM_")) {
      return "ERRO: Tema inválido.";
    }
    PropertiesService.getDocumentProperties().setProperty(themeName + "_CUSTOM_UI", JSON.stringify(uiConfig));
    PropertiesService.getDocumentProperties().setProperty("GLOBAL_THEME", themeName);
    
    return "SUCESSO";
  } catch (e) {
    return "ERRO: " + e.message;
  }
}

// ==========================================
// 5. MOTOR PRINCIPAL: TROCAR DE TEMA (APLICAÇÃO TOTAL)
// ==========================================
function applyTheme(themeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    PropertiesService.getDocumentProperties().setProperty("GLOBAL_THEME", themeName);
    const themeData = getThemeData(themeName);
    const sheet = ss.getSheetByName("Palette Entry");
    if (sheet) sheet.getRange(2, 1, themeData.values.length, 7).setValues(themeData.values);
    
    if (themeData.ui) {
      updatePatternInPalette(themeData.ui);
      applyUIColorsOnly(themeData.ui);
      updateThemeBorders(themeData.ui); 
    }
    
    if (typeof applyOutlookColorsOptimized === "function") applyOutlookColorsOptimized();
    if (typeof applyFillFromOutlookColorsOptimized === "function") applyFillFromOutlookColorsOptimized();
    
    var appearance = getThemeAppearance(themeData.ui.blockBg);
    
    return JSON.stringify({ status: "SUCESSO", ui: themeData.ui, appearance: appearance }); 
  } catch (erro) {
    throw new Error(erro.message);
  }
}

// ==========================================
// FUNÇÃO: PREVIEW AO VIVO (NÃO PERSISTE, APENAS PLANILHA/SIDE)
// ==========================================
function previewThemeColors(uiConfig) {
  try {
    updatePatternInPalette(uiConfig);
    applyUIColorsOnly(uiConfig);
    updateThemeBorders(uiConfig);
    
    // Sem chamadas de outlook colors aqui para evitar travamentos cíclicos no preview
    
    return getThemeAppearance(uiConfig.blockBg);
  } catch(e) {
    throw new Error(e.message);
  }
}

// ==========================================
// 7. BORDAS INTELIGENTES DA PLANILHA
// ==========================================
function updateThemeBorders(uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orgSheet = ss.getSheetByName("Organisieren");
  if (!orgSheet || !uiConfig) return;

  let strokeColor = uiConfig.borderColor || uiConfig.patternFill || "#333333"; 
  let footerBg = uiConfig.footerBg || uiConfig.patternSide || "#e0e0e0";

  const maxCol = orgSheet.getMaxColumns();
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

  if (maxCol >= 4) {
    orgSheet.getRange(105, 4, 1, maxCol - 3).setBackground(footerBg).setBorder(false, false, false, false, false, false);
  }

  if (maxCol >= 3) {
    orgSheet.getRange(9, 3, 95, 1).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);
    orgSheet.getRange(8, 3).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
    orgSheet.getRange(104, 3).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
    orgSheet.getRange(105, 3).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
  }

  for (let baseCol = 4; baseCol <= maxCol; baseCol += 18) {
    orgSheet.getRange(9, baseCol, 96, 17).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);

    const divisa = baseCol + 17;
    if (divisa <= maxCol) {
      orgSheet.getRange(9, divisa, 96, 1).setBorder(false, false, true, false, false, true, strokeColor, borderStyle);
      orgSheet.getRange(8, divisa).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      orgSheet.getRange(104, divisa).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      orgSheet.getRange(105, divisa).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
    }

    const patternColsOffsets = [0, 1, 3, 5, 7, 9, 11, 13, 15, 16]; 
    const mainColsOffsets = [2, 4, 6, 8, 10, 12, 14];              
    
    patternColsOffsets.forEach(offset => {
      const col = baseCol + offset;
      if (col <= maxCol) {
        orgSheet.getRange(8, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
        orgSheet.getRange(104, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
        orgSheet.getRange(105, col).setBorder(true, false, false, false, false, false, strokeColor, borderStyle);
      }
    });

    mainColsOffsets.forEach(offset => {
      const col = baseCol + offset;
      if (col <= maxCol) {
        orgSheet.getRange(8, col).setBorder(false, false, true, false, false, false, uiConfig.headerBg, borderStyle);
        orgSheet.getRange(104, col).setBorder(false, false, true, false, false, false, strokeColor, borderStyle);
      }
    });
  }
  SpreadsheetApp.flush();
}

// ==========================================
// MATEMÁTICA DE COR PARA A BORDA DA PLANILHA
// ==========================================
function getSmartStrokeColor(hex, isDark) {
  if (!/^#[0-9A-F]{6}$/i.test(hex)) return hex;
  let r = parseInt(hex.slice(1, 3), 16); let g = parseInt(hex.slice(3, 5), 16); let b = parseInt(hex.slice(5, 7), 16);
  const amount = 35; 
  if (isDark) {
    r = Math.min(255, r + amount); g = Math.min(255, g + amount); b = Math.min(255, b + amount);
  } else {
    r = Math.max(0, r - amount); g = Math.max(0, g - amount); b = Math.max(0, b - amount);
  }
  const toHex = (n) => { let h = n.toString(16); return h.length === 1 ? "0" + h : h; };
  return "#" + toHex(r) + toHex(g) + toHex(b);
}

// ==========================================
// 8. ATUALIZA O PATTERN NA MATRIZ RAIZ
// ==========================================
function updatePatternInPalette(uiConfig) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const paletteSheet = ss.getSheetByName("Palette Entry");
  if (!paletteSheet || !uiConfig) return;

  const lastRow = paletteSheet.getLastRow();
  if (lastRow < 1) return;

  const labels = paletteSheet.getRange(1, 12, lastRow, 1).getValues();

  let blockBg = uiConfig.blockBg || "#000000";
  let app = getThemeAppearance(blockBg);
  let textColor = app.foreground;
  let isDark = app.isDark;

  for (let i = 0; i < labels.length; i++) {
    const nameVal = labels[i][0] ? labels[i][0].toString().toLowerCase().trim() : "";

    if (nameVal.includes("pattern")) {
      const row = i + 1;
      if (uiConfig.patternSide) paletteSheet.getRange(row, 1).setValue(uiConfig.patternSide);
      paletteSheet.getRange(row, 3).setValue(blockBg);
      let border = uiConfig.borderColor || getSmartStrokeColor(blockBg, isDark);
      paletteSheet.getRange(row, 7).setValue(border);
      paletteSheet.getRange(row, 5).setValue(textColor);
      break;
    }
  }
  SpreadsheetApp.flush();
}

// ==========================================
// FUNÇÃO: DELETAR TEMA CUSTOMIZADO
// ==========================================
function deleteCustomTheme(themeKey) {
  try {
    const props = PropertiesService.getDocumentProperties();
    props.deleteProperty(themeKey + "_CUSTOM_UI");
    
    let savedThemesStr = props.getProperty("USER_SAVED_THEMES");
    let savedThemes = savedThemesStr ? JSON.parse(savedThemesStr) : [];
    savedThemes = savedThemes.filter(t => t.key !== themeKey);
    props.setProperty("USER_SAVED_THEMES", JSON.stringify(savedThemes));
    
    let current = props.getProperty("GLOBAL_THEME");
    let fallback = "DARK_MODERN";
    
    if (current === themeKey) {
      props.setProperty("GLOBAL_THEME", fallback);
    } else {
      fallback = current; 
    }
    return { status: "SUCESSO", fallbackTheme: fallback };
  } catch (e) {
    throw new Error(e.message);
  }
}