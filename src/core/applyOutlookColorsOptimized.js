function applyOutlookColorsOptimized(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e ? e.range.getSheet() : ss.getSheetByName("Palette Entry");
  
  const lastRow = sheet.getMaxRows(); 
  if (lastRow < 2) return;

  const range = sheet.getRange(1, 1, lastRow, 18);
  let values = range.getValues();
  
  // Sincronia da Coluna O
  const masterSelectionMap = {};
  for (let i = 1; i < values.length; i++) { 
    const ctx = String(values[i][13] || "").trim().toLowerCase(); 
    const sub = String(values[i][17] || "").trim();               
    if (ctx !== "" && sub === "") {
      masterSelectionMap[ctx] = values[i][14];
    }
  }

  let changedSync = false;
  for (let i = 1; i < values.length; i++) { 
    const ctx = String(values[i][13] || "").trim().toLowerCase();
    const sub = String(values[i][17] || "").trim();
    if (sub !== "" && masterSelectionMap[ctx] !== undefined) {
      if (values[i][14] !== masterSelectionMap[ctx]) {
        sheet.getRange(i + 1, 15).setValue(masterSelectionMap[ctx]);
        changedSync = true;
      }
    }
  }

  if (changedSync) {
    SpreadsheetApp.flush();
    values = range.getValues(); 
  }

  const bg = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const isColorMode = PropertiesService.getDocumentProperties().getProperty("COLOR_FONT_MODE") === "true";
  
  const globalTheme = PropertiesService.getDocumentProperties().getProperty("GLOBAL_THEME") || "DARK_MODERN";
  const isLightTheme = globalTheme.includes("LIGHT");
  const defaultTextColor = isLightTheme ? "#000000" : "#ffffff";
  
  // MAPEAMENTO: Fundo e Texto (Ignorando Cabeçalho)
  const fillToTextMap = {};
  for (let i = 1; i < values.length; i++) { 
    const fillHex = values[i][2]; 
    const textHex = values[i][4]; 
    if (isValidHex(fillHex) && isValidHex(textHex)) {
      fillToTextMap[fillHex.toString().trim().toLowerCase()] = textHex.toString().trim();
    }
  }

  // APLICAÇÃO DE CORES
  for (let i = 1; i < values.length; i++) { 
    const row = values[i];

    // COLUNAS DE BLOCO PURO (A, C, E, G, P) 
    [0, 2, 4, 6, 15].forEach(colIdx => {
      if (isValidHex(row[colIdx])) {
        const hex = row[colIdx].toString().trim();
        bg[i][colIdx] = hex;
        fontColors[i][colIdx] = hex; 
      }
    });

    // --- AJUSTE ESPECÍFICO COLUNA J (9) ---
    if (isValidHex(row[0])) {
      const sideHex = row[0].toString().trim();
      bg[i][9] = sideHex;
      fontColors[i][9] = sideHex;
    }

    // --- COLUNA K (10) MANTIDA COM SEU CONCEITO ---
    if (isValidHex(row[2])) {
      const baseFillRaw = row[2].toString().trim();
      bg[i][10] = baseFillRaw;
      fontColors[i][10] = isColorMode ? (fillToTextMap[baseFillRaw.toLowerCase()] || defaultTextColor) : defaultTextColor;
    }

    // --- COLUNA Q (16) MANTIDA ORIGINAL ---
    if (isValidHex(row[16])) {
      const qFillRaw = row[16].toString().trim();
      bg[i][16] = qFillRaw; 
      fontColors[i][16] = isColorMode ? (fillToTextMap[qFillRaw.toLowerCase()] || defaultTextColor) : defaultTextColor; 
    }
  }

  range.setBackgrounds(bg);
  range.setFontColors(fontColors);

  // Bordas na coluna K vinda da G (Stroke)
  for (let i = 1; i < values.length; i++) { 
    if (isValidHex(values[i][2]) && isValidHex(values[i][6])) {
      sheet.getRange(i + 1, 11).setBorder(true, true, true, true, null, null, values[i][6], SpreadsheetApp.BorderStyle.SOLID);
    }
  }
}

function isValidHex(color) {
  return typeof color === 'string' && /^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/.test(color.toString().trim());
}