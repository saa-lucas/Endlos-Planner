// =====================================================================
// 🛡️ MOTOR DE ESTABILIDADE E SERIALIZAÇÃO (MUTEX + DEBOUNCE)
// =====================================================================

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (sheetName !== "Palette Entry" && sheetName !== "Organisieren") return;

  const editData = {
    sheetName: sheetName,
    row: e.range.getRow(),
    col: e.range.getColumn(),
    numRows: e.range.getNumRows(),
    numCols: e.range.getNumColumns(),
    value: e.value,
    bgRaw: e.range.getBackground()
  };

  const cache = CacheService.getScriptCache();
  const editId = Date.now().toString() + "_" + Math.random().toString(36).substr(2, 5);
  
  cache.put('LAST_EDIT_ID', editId, 10); 
  Utilities.sleep(1500); 
  
  if (cache.get('LAST_EDIT_ID') !== editId) {
    return; 
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const safeRange = ss.getSheetByName(editData.sheetName).getRange(editData.row, editData.col, editData.numRows, editData.numCols);

    if (editData.sheetName === "Palette Entry") {
      if ([14, 15, 18, 19].includes(editData.col)) {
        triggerSafeGlobalRepaint();
      }
    } 
    else if (editData.sheetName === "Organisieren") {
      if (editData.numRows === 1 && editData.numCols === 1 && editData.value) {
        captureNewNameInContextSafe(safeRange, editData.value, editData.bgRaw);
      }
      applyFillFromOutlookColorsOptimized(safeRange);
    }
  } catch (err) {
    console.error("Erro Crítico na Execução Protegida: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// 🛡️ CONTROLE DE REPAINT GLOBAL
// =====================================================================

function triggerSafeGlobalRepaint() {
  const cache = CacheService.getScriptCache();
  const lastRepaint = cache.get('LAST_REPAINT_TIME');
  const now = Date.now();

  if (lastRepaint && (now - parseInt(lastRepaint) < 3000)) {
    return;
  }

  cache.put('LAST_REPAINT_TIME', now.toString(), 15);
  PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', now.toString());
  SpreadsheetApp.flush();
  
  try {
    applyFillFromOutlookColorsOptimized();
  } catch (err) {
    console.error("Erro no Global Repaint: " + err);
  }
}

function captureNewNameInContextSafe(safeRange, newValue, bgColor) {
  if (typeof newValue !== 'string' || newValue.trim() === "") return; 
  const defaultBgs = ["#ffffff", "#000000", "#111111", "#1e1e1e", "#fff2cc"]; 
  if (defaultBgs.includes(bgColor)) return; 

  const ss = safeRange.getSheet().getParent();
  const paletteSheet = ss.getSheetByName("Palette Entry");
  if (!paletteSheet) return;

  const lastRow = paletteSheet.getLastRow();
  if (lastRow < 1) return;

  const data = paletteSheet.getRange(1, 1, lastRow, 18).getValues();
  let rowToDuplicateIndex = -1;
  const typed = newValue.trim().toLowerCase();

  for (let i = 2; i < data.length; i++) {
    const contextText = String(data[i][13] || "").trim().toLowerCase(); 
    const subActionText = String(data[i][17] || "").trim().toLowerCase(); 
    if (contextText === typed || subActionText === typed) return; 

    if (rowToDuplicateIndex === -1) {
      const bgMain = String(data[i][16] || "").toLowerCase(); 
      if (bgMain === bgColor) rowToDuplicateIndex = i + 1; 
    }
  }

  if (rowToDuplicateIndex !== -1) {
    const newRowIndex = Math.max(lastRow + 1, 32); 
    const sourceRange = paletteSheet.getRange(rowToDuplicateIndex, 14, 1, 5);
    const targetRange = paletteSheet.getRange(newRowIndex, 14, 1, 5);
    sourceRange.copyTo(targetRange); 
    paletteSheet.getRange(newRowIndex, 18).setValue(newValue.trim());
    PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', Date.now().toString());
  }
}

function applyFillFromOutlookColorsOptimized(editedRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const testSheet = ss.getSheetByName("Organisieren");
  const outlookSheet = ss.getSheetByName("Palette Entry");
  if (!outlookSheet || !testSheet) return;

  const props = PropertiesService.getDocumentProperties();
  const isColorFont = props.getProperty("COLOR_FONT_MODE") === "true";
  const globalTheme = props.getProperty("GLOBAL_THEME") || "DARK_MODERN";
  const isLightTheme = globalTheme.includes("LIGHT");
  const defaultTextColor = isLightTheme ? "#000000" : "#ffffff";

  const lastRow = outlookSheet.getLastRow();
  if (lastRow < 1) return;

  const outlookValues = outlookSheet.getRange(1, 1, lastRow, 18).getValues();
  const fillToTextMap = {};
  const colorMap = {};
  const hexRegex = /^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/i;

  let globalPatternSide = isLightTheme ? "#e0e0e0" : "#262626";
  let globalPatternMain = isLightTheme ? "#f9f9f9" : "#111111";

  for (let i = 0; i < outlookValues.length; i++) {
    const rowStr = outlookValues[i].join("|").toLowerCase();
    if (rowStr.includes("pattern")) {
      const aColor = String(outlookValues[i][0]).trim(); 
      const cColor = String(outlookValues[i][2]).trim(); 
      if (hexRegex.test(aColor)) globalPatternSide = aColor;
      if (hexRegex.test(cColor)) globalPatternMain = cColor;
      break; 
    }
  }

  const timeBg = isLightTheme ? "#ffffff" : "#000000";

  for (let i = 2; i < outlookValues.length; i++) {
    const fillHex = String(outlookValues[i][2]).trim();
    const textHex = String(outlookValues[i][4]).trim();
    if (hexRegex.test(fillHex)) {
      fillToTextMap[fillHex.toLowerCase()] = textHex || defaultTextColor;
    }
  }

  for (let i = 2; i < outlookValues.length; i++) {
    const bgMainRaw = String(outlookValues[i][16]).trim().toLowerCase();
    if (!bgMainRaw || !hexRegex.test(bgMainRaw)) continue;
    
    const sideColor = String(outlookValues[i][15]).trim();
    const isValidSide = hexRegex.test(sideColor);
    const trueFontColor = fillToTextMap[bgMainRaw] || defaultTextColor;
    
    const config = { 
      main: bgMainRaw, 
      font: trueFontColor, 
      side: isValidSide ? sideColor : bgMainRaw 
    };
    
    const context = String(outlookValues[i][13]).trim().toLowerCase();
    const subAction = String(outlookValues[i][17]).trim().toLowerCase();
    
    if (context.includes("pattern") || subAction.includes("pattern")) continue;

    if (context) colorMap[context] = config;
    if (subAction) colorMap[subAction] = config;
  }

  let range;
  let ALLOW_LUXURY = false;

  if (editedRange) {
    let rRow = editedRange.getRow();
    let rCol = editedRange.getColumn();
    let rNumRows = editedRange.getNumRows();
    let rNumCols = editedRange.getNumColumns();
    if (rCol < 3) return; 
    if (rNumRows <= 6 && rNumCols <= 2) ALLOW_LUXURY = true;
    range = testSheet.getRange(rRow, rCol - 1, rNumRows, rNumCols + 1);
  } else {
    const maxCols = testSheet.getMaxColumns();
    const maxRows = Math.min(testSheet.getMaxRows(), 105);
    range = testSheet.getRange(1, 1, maxRows, maxCols);
  }

  let values = range.getValues();
  let bg = range.getBackgrounds();
  let fontColors = range.getFontColors();
  let fontWeights = range.getFontWeights();
  let vAligns = range.getVerticalAlignments(); 
  let hAligns = range.getHorizontalAlignments();

  let mergeMap = {};
  if (ALLOW_LUXURY) {
    let merges = range.getMergedRanges();
    for (let i = 0; i < merges.length; i++) {
      mergeMap[`${merges[i].getRow()}_${merges[i].getColumn()}`] = merges[i].getNumRows(); 
    }
  }

  const richTextUpdates = []; 
  const sideMergeUpdates = []; 
  const absRowOffset = range.getRow();
  const absColOffset = range.getColumn();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const absRow = absRowOffset + r;
      const absCol = absColOffset + c;

      if (absCol < 4 || absRow < 9) continue;

      const rawVal = values[r][c];
      const value = String(rawVal || "").trim();

      if (!value) {
        const colRel = absCol - 4; 
        const posInWeek = colRel % 18; 

        if (absRow <= 104) {
          if (posInWeek === 17) {
             bg[r][c] = timeBg; 
          } else {
             bg[r][c] = (posInWeek % 2 === 0) ? globalPatternMain : globalPatternSide; 
          }
          fontColors[r][c] = defaultTextColor;
        } 
        else if (absRow === 105) {
          // 🟢 FOOTER IGUAL À TIMELINE
          bg[r][c] = timeBg;
          fontColors[r][c] = defaultTextColor;
        }
        continue; 
      }

      const lowerValue = value.toLowerCase();
      const isOnlyDashes = value.length > 0 && value.replace(/[-\s]/g, '') === '';
      const mapData = colorMap[lowerValue];

      if (mapData || isOnlyDashes) {
        const currentRows = ALLOW_LUXURY ? (mergeMap[`${absRow}_${absCol}`] || 1) : 1;
        const currentSideRows = ALLOW_LUXURY ? (mergeMap[`${absRow}_${absCol - 1}`] || 1) : 1;
        if (ALLOW_LUXURY && currentSideRows !== currentRows) {
          sideMergeUpdates.push({ row: absRow, col: absCol - 1, numRows: currentRows });
        }
        let bgMain = mapData ? mapData.main : bg[r][c];
        let fgMain = mapData ? (isOnlyDashes ? bgMain : (isColorFont ? mapData.font : defaultTextColor)) : bg[r][c];
        let sideCol = mapData ? mapData.side : bgMain;
        for (let i = 0; i < currentRows; i++) {
          if (r + i < bg.length) {
            if (mapData) {
              bg[r + i][c] = bgMain;
              bg[r + i][c - 1] = sideCol;
            }
            fontWeights[r + i][c] = "bold";
            vAligns[r + i][c] = "middle";
            hAligns[r + i][c] = "center";
          }
        }
        fontColors[r][c] = fgMain; 
        if (ALLOW_LUXURY && mapData && currentRows === 1 && value.includes(" ") && value !== "Start" && value !== "End" && value.length > 15 && !isOnlyDashes) {
          let splitIndex = value.indexOf(" ");
          if (splitIndex !== -1 && splitIndex <= 3) {
            const nextSpace = value.indexOf(" ", splitIndex + 1);
            if (nextSpace !== -1) splitIndex = nextSpace;
          }
          const textStyleVis = SpreadsheetApp.newTextStyle().setForegroundColor(fgMain).setBold(true).build();
          const textStyleInvis = SpreadsheetApp.newTextStyle().setForegroundColor(bgMain).setBold(false).build();
          const rtv = SpreadsheetApp.newRichTextValue().setText(value).setTextStyle(0, splitIndex, textStyleVis).setTextStyle(splitIndex, value.length, textStyleInvis).build();
          richTextUpdates.push({ row: absRow, col: absCol, rtv: rtv });
        }
      }
    }
  }

  range.setBackgrounds(bg);
  range.setFontColors(fontColors);
  range.setFontWeights(fontWeights);
  range.setVerticalAlignments(vAligns);
  range.setHorizontalAlignments(hAligns);

  if (ALLOW_LUXURY) {
    for (let i = 0; i < richTextUpdates.length; i++) {
      testSheet.getRange(richTextUpdates[i].row, richTextUpdates[i].col).setRichTextValue(richTextUpdates[i].rtv);
    }
    for (let i = 0; i < sideMergeUpdates.length; i++) {
      let sideRange = testSheet.getRange(sideMergeUpdates[i].row, sideMergeUpdates[i].col, sideMergeUpdates[i].numRows, 1);
      sideRange.breakApart(); 
      if (sideMergeUpdates[i].numRows > 1) sideRange.merge();
    }
  }
}