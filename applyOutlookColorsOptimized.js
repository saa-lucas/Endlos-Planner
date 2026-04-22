function applyOutlookColorsOptimized(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e ? e.range.getSheet() : ss.getSheetByName("Palette Entry");
  
  // CORREÇÃO DE ALCANCE: Garante que passe da linha 27
  const lastRow = sheet.getMaxRows(); 
  if (lastRow < 1) return;

  const range = sheet.getRange(1, 1, lastRow, 18);
  let values = range.getValues();
  
  // Sincronia da Coluna O (Mantida original)
  const masterSelectionMap = {};
  for (let i = 0; i < values.length; i++) {
    const ctx = String(values[i][13] || "").trim().toLowerCase(); 
    const sub = String(values[i][17] || "").trim();               
    if (ctx !== "" && sub === "") {
      masterSelectionMap[ctx] = values[i][14];
    }
  }

  for (let i = 0; i < values.length; i++) {
    const ctx = String(values[i][13] || "").trim().toLowerCase();
    const sub = String(values[i][17] || "").trim();
    if (sub !== "" && masterSelectionMap[ctx] !== undefined) {
      if (values[i][14] !== masterSelectionMap[ctx]) {
        sheet.getRange(i + 1, 15).setValue(masterSelectionMap[ctx]);
      }
    }
  }

  SpreadsheetApp.flush();
  values = range.getValues(); 

  const bg = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const isColorMode = PropertiesService.getDocumentProperties().getProperty("COLOR_FONT_MODE") === "true";
  
  const globalTheme = PropertiesService.getDocumentProperties().getProperty("GLOBAL_THEME") || "DARK_MODERN";
  const isLightTheme = globalTheme.includes("LIGHT");
  const defaultTextColor = isLightTheme ? "#000000" : "#ffffff";
  
  const patternHex = "#fff2cc";

  // MAPEAMENTO: Índice 2 (C) é Fundo, Índice 4 (E) é Texto
  const fillToTextMap = {};
  for (let i = 0; i < values.length; i++) {
    const fillHex = values[i][2]; 
    const textHex = values[i][4]; 
    if (isValidHex(fillHex) && isValidHex(textHex)) {
      fillToTextMap[fillHex.toString().trim().toLowerCase()] = textHex.toString().trim();
    }
  }

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    // COLUNAS DE BLOCO PURO (A, C, E, G, P) - Mantido original
    [0, 2, 4, 6, 15].forEach(colIdx => {
      if (isValidHex(row[colIdx])) {
        const hex = row[colIdx].toString().trim();
        bg[i][colIdx] = hex;
        fontColors[i][colIdx] = hex; 
      }
    });

    // --- AJUSTE ESPECÍFICO COLUNA J (9) ---
    // Faz a J espelhar a cor da Side (A) para não ficar amarela na linha 27
    if (isValidHex(row[0])) {
      const sideHex = row[0].toString().trim();
      bg[i][9] = sideHex;
      fontColors[i][9] = sideHex;
    }

    // --- COLUNA K (10) MANTIDA COM SEU CONCEITO ---
    // Fundo de C + Texto mapeado de E + Borda de G
    if (isValidHex(row[2])) {
      const baseFillRaw = row[2].toString().trim();
      const baseFillLower = baseFillRaw.toLowerCase();
      bg[i][10] = baseFillRaw;

      if (baseFillLower === patternHex) {
        fontColors[i][10] = fillToTextMap[baseFillLower] || defaultTextColor;
      } else {
        fontColors[i][10] = isColorMode ? (fillToTextMap[baseFillLower] || defaultTextColor) : defaultTextColor;
      }
    }

    // --- COLUNA Q (16) MANTIDA ORIGINAL ---
    if (isValidHex(row[16])) {
      const qFillRaw = row[16].toString().trim();
      const qFillLower = qFillRaw.toLowerCase();
      bg[i][16] = qFillRaw; 
      
      if (qFillLower === patternHex) {
        fontColors[i][16] = fillToTextMap[qFillLower] || defaultTextColor;
      } else {
        fontColors[i][16] = isColorMode ? (fillToTextMap[qFillLower] || defaultTextColor) : defaultTextColor; 
      }
    }
  }

  range.setBackgrounds(bg);
  range.setFontColors(fontColors);

  // Bordas na coluna K vinda da G (Stroke) - Mantido funcionando
  for (let i = 0; i < values.length; i++) {
    if (isValidHex(values[i][2]) && isValidHex(values[i][6])) {
      sheet.getRange(i + 1, 11).setBorder(true, true, true, true, null, null, values[i][6], SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

// Única função de validação para evitar conflitos
function isValidHex(color) {
  return typeof color === 'string' && /^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/.test(color.toString().trim());
}