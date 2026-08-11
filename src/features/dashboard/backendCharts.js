// =========================================================================
// FUNÇÕES DO DASHBOARD (Abertura, Extração de Dados e Temas)
// =========================================================================

// ABRIR O MODAL DO DASHBOARD (ATUALIZADO PARA O CÉREBRO HÍBRIDO)
function openDashboardModal() {
  var template = HtmlService.createTemplateFromFile('ui/dashboard');
  template.ambiente = 'PLANILHA';

  var htmlOutput = template.evaluate()
    .setWidth(1100)
    .setHeight(650)
    .setTitle('📈 Organisieren Analytics');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Analytics');
  }

// O CÉREBRO DO TEMA: Lê a cor real da planilha e define se é Light ou Dark
function getSpreadsheetThemeMode() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const theme = ss.getSpreadsheetTheme();
    
    if (!theme) return null; // Se não houver tema na planilha, o HTML decide

    const bgColor = theme.getThemeColorType(SpreadsheetApp.ThemeColorType.BACKGROUND).asRgbColor().asHexString();
    const hex = bgColor.replace('#', '');
    const r = parseInt(hex.substring(0, 2), 16);
    const g = parseInt(hex.substring(2, 4), 16);
    const b = parseInt(hex.substring(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b);
    
    return luminance > 128 ? 'light' : 'dark';
    
  } catch (e) {
    return null; 
  }
}

// =========================================================================
// FUNÇÕES DE CONFIGURAÇÃO (PREFERÊNCIAS)
// =========================================================================

function getSettingsState() {
  const props = PropertiesService.getDocumentProperties();
  return {
    isColorFont: props.getProperty("COLOR_FONT_MODE") === "true",
    isMealSplit: props.getProperty("POST_MEAL_1H_MODE") === "true"
  };
}

function toggleColorFont(isColor) {
  PropertiesService.getDocumentProperties().setProperty("COLOR_FONT_MODE", isColor.toString());
  if (typeof applyFillFromOutlookColorsOptimized === 'function') applyFillFromOutlookColorsOptimized(); 
}

function setMealSplitState(isMealSplit) {
  PropertiesService.getDocumentProperties().setProperty("POST_MEAL_1H_MODE", isMealSplit.toString());
  if (typeof updateHoursDashboard === "function") updateHoursDashboard(); 
}

// =====================================================================
// EVENTO DE ESTRUTURA (Deteta Inserção de Linhas/Colar Dados)
// =====================================================================

function onStructuralChange(e) {
  if (!e) return;
  // Qualquer mudança de estrutura avisa a Sidebar para buscar novas ações/órfãs
  PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', Date.now().toString());
}

// ⚠️ RODE ISTO UMA VEZ NO EDITOR PARA INSTALAR O GATILHO
function installArchitectureTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getUserTriggers(ss);
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onStructuralChange") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onStructuralChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
    
  console.log("Arquitetura de Sincronização instalada com sucesso.");
}

// =====================================================================
// API DO SYSTEM CONFIG BUILDER (ARQUITETURA DE FONTE ÚNICA)
// =====================================================================

function checkSystemVersion(clientVersion) {
  const currentVersion = PropertiesService.getDocumentProperties().getProperty('SYS_VERSION') || "0";
  return { needsUpdate: currentVersion !== clientVersion, newVersion: currentVersion };
}

function getSystemConfigFromPalette() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Palette Entry");
  if (!sheet) return JSON.stringify({ contexts: [], actions: [] });

  const data = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), 19).getValues();
  const state = { contexts: [], actions: [] };
  const ctxMap = {};
  const PROTECTED_CTX = ["start", "end", "entrada", "saída", "suporte", "desperdício", "recuperação", "refeição", "unwind"];

  for (let i = 1; i < data.length; i++) {
    const ctxName = String(data[i][13] || "").trim(); // N
    const colorName = String(data[i][14] || "").trim(); // O
    const colorHex = String(data[i][15] || "").trim(); // P -> AQUI ESTÁ O HEX REAL DA PLANILHA!
    const actName = String(data[i][17] || "").trim(); // R
    let modeName = String(data[i][18] || "").trim(); // S

    if (modeName === "-" || modeName.toLowerCase() === "sem categoria") modeName = "";
    
    const ctxLower = ctxName.toLowerCase();

    if (ctxName && !actName) {
      if (PROTECTED_CTX.includes(ctxLower)) continue; 
      let id = 'c_' + i;
      
      // AGORA ENVIAMOS O HEXADECIMAL PARA A SIDEBAR!
      state.contexts.push({ 
        id: id, 
        name: ctxName, 
        colorName: colorName || 'Pattern', 
        colorHex: colorHex || '#cccccc', 
        mode: modeName 
      });
      ctxMap[ctxLower] = id;
    } else if (ctxName && actName) {
      if (PROTECTED_CTX.includes(ctxLower)) continue; 
      let parentId = ctxMap[ctxLower] || ('c_unknown_' + i); 
      state.actions.push({ 
        id: 'a_' + i, 
        contextId: parentId, 
        name: actName, 
        modeOverride: modeName, 
        originalCtxName: ctxName 
      });
    }
  }
  return JSON.stringify(state);
}

function saveSystemConfigToPalette(jsonString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Palette Entry");
  if (!sheet) throw new Error("Aba 'Palette Entry' não encontrada.");

  const state = JSON.parse(jsonString);
  const PROTECTED_CTX = ["start", "end", "entrada", "saída", "suporte", "desperdício", "recuperação", "refeição", "unwind"];

  // 🛡️ ASSASSINO DE ESPAÇOS: Garante que "Pattern " vire "Pattern"
  const safeColor = (c) => c ? String(c).trim() : "Pattern";

  // 1. GUARDA OS FIXOS (Com travas contra Undefined)
  const currentData = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), 19).getValues();
  let protectedContexts = [], protectedActions = [];

  for (let i = 1; i < currentData.length; i++) {
    const ctxName = String(currentData[i][13] || "").trim(); // N
    if (ctxName && PROTECTED_CTX.includes(ctxName.toLowerCase())) {
      const colorName = safeColor(currentData[i][14]); // O
      const modeName = String(currentData[i][18] || "-").trim(); // S
      const actionName = String(currentData[i][17] || "").trim(); // R

      if (!actionName) {
        protectedContexts.push([ctxName, colorName, modeName]);
      } else {
        protectedActions.push([ctxName, colorName, actionName, modeName]);
      }
    }
  }

  // 2. LIMPEZA CIRÚRGICA (Não toca nas fórmulas da col P e Q, e remove o rastro fantasma)
  sheet.getRange("N2:O30").clearContent(); 
  sheet.getRange("O3:O30").clearDataValidations(); 
  
  sheet.getRange("S2:S30").clearContent();
  sheet.getRange("S3:S30").clearDataValidations(); 
  
  const lastRow = Math.max(sheet.getLastRow(), 33);
  if (lastRow >= 32) { 
    sheet.getRange("N32:O" + lastRow).clearContent(); 
    sheet.getRange("R32:S" + lastRow).clearContent(); 
    if (lastRow >= 33) {
      sheet.getRange("O33:O" + lastRow).clearDataValidations(); 
      sheet.getRange("S33:S" + lastRow).clearDataValidations();
    }
  }

  // 3. GRAVAÇÃO DE CONTEXTOS
  let ctxNO = [], ctxS = [];
  protectedContexts.forEach(pc => { 
    ctxNO.push([pc[0] || "", pc[1]]); 
    ctxS.push([pc[2] || ""]); 
  });
  
  state.contexts.forEach(c => { 
    ctxNO.push([c.name || "", safeColor(c.colorName)]); 
    ctxS.push([c.mode || ""]); 
  });
  
  const ctxRows = ctxNO.length;

  // 🛑 TRAVA DE SEGURANÇA (TETO DE CONTEXTOS)
  if (ctxRows > 26) {
    throw new Error("Teto atingido: O sistema suporta no máximo 26 contextos (limite na linha 27). Remova um antigo antes de adicionar outro.");
  }

  if (ctxRows > 0) { 
    sheet.getRange(2, 14, ctxRows, 2).setValues(ctxNO); 
    sheet.getRange(2, 19, ctxRows, 1).setValues(ctxS); 
    
    // 🎨 CORREÇÃO VISUAL: Força Alinhamento à Direita na Coluna N
    sheet.getRange(2, 14, ctxRows, 1).setHorizontalAlignment("right"); 
    
    if (ctxRows > 1) {
      // Clona a Validação de Dados (Dropdown) da Coluna O
      let sourceO = sheet.getRange("O2");
      let destO = sheet.getRange(3, 15, ctxRows - 1, 1);
      sourceO.copyTo(destO, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      sourceO.copyTo(destO, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      
      // Clona as Fórmulas de Cor das Colunas P e Q
      let sourcePQ = sheet.getRange("P2:Q2");
      let destPQ = sheet.getRange(3, 16, ctxRows - 1, 2);
      sourcePQ.copyTo(destPQ, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      sourcePQ.copyTo(destPQ, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }

  // 4. GRAVAÇÃO DE AÇÕES
  let actNO = [], actRS = [];
  protectedActions.forEach(pa => { 
    actNO.push([pa[0] || "", pa[1]]); 
    actRS.push([pa[2] || "", pa[3] || ""]); 
  });
  
  state.actions.forEach(a => {
    let parentCtx = state.contexts.find(c => c.id === a.contextId);
    let pName = parentCtx ? parentCtx.name : (a.originalCtxName || "Desconhecido");
    let pColor = parentCtx ? safeColor(parentCtx.colorName) : "Pattern";
    
    actNO.push([pName || "", pColor]); 
    actRS.push([a.name || "", a.modeOverride || ""]);
  });
  
  const actRows = actNO.length;
  if (actRows > 0) { 
    sheet.getRange(32, 14, actRows, 2).setValues(actNO); 
    sheet.getRange(32, 18, actRows, 2).setValues(actRS); 
    
    // 🎨 CORREÇÃO VISUAL: Força Alinhamento à Direita nas Colunas N e R
    sheet.getRange(32, 14, actRows, 1).setHorizontalAlignment("right"); 
    sheet.getRange(32, 18, actRows, 1).setHorizontalAlignment("right"); 
    
    if (actRows > 1) {
      // Clona a Validação de Dados (Dropdown) da Coluna O
      let sourceO_act = sheet.getRange("O32");
      let destO_act = sheet.getRange(33, 15, actRows - 1, 1);
      sourceO_act.copyTo(destO_act, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      sourceO_act.copyTo(destO_act, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      
      // Clona as Fórmulas de Cor das Colunas P e Q
      let sourcePQ_act = sheet.getRange("P32:Q32");
      let destPQ_act = sheet.getRange(33, 16, actRows - 1, 2);
      sourcePQ_act.copyTo(destPQ_act, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      sourcePQ_act.copyTo(destPQ_act, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
  }

  // 5. ATUALIZAÇÃO E ESTABILIDADE
  SpreadsheetApp.flush(); 
  PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', Date.now().toString());

  // Aguarda a planilha processar as fórmulas e manda pintar!
  Utilities.sleep(1000);
  if (typeof applyOutlookColorsOptimized === 'function') applyOutlookColorsOptimized(null);
  if (typeof triggerSafeGlobalRepaint === 'function') triggerSafeGlobalRepaint();

  return "SUCESSO";
}