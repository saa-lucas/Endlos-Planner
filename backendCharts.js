// =========================================================================
// FUNÇÕES DO DASHBOARD (Abertura, Extração de Dados e Temas)
// =========================================================================

// ABRIR O MODAL DO DASHBOARD (ATUALIZADO PARA O CÉREBRO HÍBRIDO)
function openDashboardModal() {
  var template = HtmlService.createTemplateFromFile('dashboard');
  
  // A MÁGICA AQUI: Avisamos o HTML que ele está dentro da Planilha
  template.ambiente = 'PLANILHA'; 
  
  var htmlOutput = template.evaluate()
      .setWidth(1100)
      .setHeight(650)
      .setTitle('📈 Organisieren Analytics');
      
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Analytics');
}

// PEGAR OS DADOS DA TIMELINE NA ABA DATEN E AS CORES DA PALETA
function getDashboardDataJson() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Daten");
    const palSheet = ss.getSheetByName("Palette Entry");
    
    if (!sheet) return { timeline: [], colors: {} };
    
    const colorMap = {};
    if (palSheet) {
      const palLastRow = palSheet.getLastRow();
      if (palLastRow > 0) {
        const palData = palSheet.getRange(1, 1, palLastRow, 18).getValues();
        for (let i = 0; i < palData.length; i++) {
          const ctx = String(palData[i][13] || "").trim(); // Coluna N
          const act = String(palData[i][17] || "").trim(); // Coluna R
          
          let hex = String(palData[i][15] || "").trim().toLowerCase(); 
          if (!hex) hex = String(palData[i][16] || "").trim().toLowerCase();
          
          if (/^#([0-9A-Fa-f]{6}|[0-9A-Fa-f]{3})$/i.test(hex)) {
            if (ctx && !act) colorMap[ctx] = hex;
            if (act) colorMap[act] = hex; 
          }
        }
      }
    }

    function hexToRgb(hex) {
      let c = hex.substring(1).split('');
      if(c.length === 3) c = [c[0], c[0], c[1], c[1], c[2], c[2]];
      c = '0x' + c.join('');
      return [(c>>16)&255, (c>>8)&255, c&255];
    }

    // --- REGRAS ESPECIAIS DE CORES (ATUALIZADAS PARA PT-BR) ---
    let startColor = colorMap["Start"] || colorMap["Início"];
    let endColor = colorMap["End"] || colorMap["Fim"];
    
    if (startColor && endColor) {
      let rgb1 = hexToRgb(startColor);
      let rgb2 = hexToRgb(endColor);
      
      let blendedHex = "#" + (
        (1 << 24) + 
        (Math.round((rgb1[0] + rgb2[0]) / 2) << 16) + 
        (Math.round((rgb1[1] + rgb2[1]) / 2) << 8) + 
        Math.round((rgb1[2] + rgb2[2]) / 2)
      ).toString(16).slice(1);
      
      colorMap["Sono"] = blendedHex;
      colorMap["Sono (Calculado)"] = blendedHex;
      colorMap["Sleep"] = blendedHex; 
    } else {
      colorMap["Sono"] = "#336699"; 
      colorMap["Sono (Calculado)"] = "#336699";
    }

    let leisureColor = colorMap["Lazer"] || colorMap["Leisure"] || colorMap["Unwind"] || "#00ffaa"; 
    if (leisureColor) {
      let rgbL = hexToRgb(leisureColor);
      let oppHex = "#" + ((1 << 24) + ((255 - rgbL[0]) << 16) + ((255 - rgbL[1]) << 8) + (255 - rgbL[2])).toString(16).slice(1);
      
      // BLINDAGEM AQUI: Cobre todas as variações para o gerador dinâmico não falhar
      colorMap["Distração"] = oppHex;
      colorMap["Distracao"] = oppHex;
      colorMap["Tempo Perdido"] = oppHex;
      colorMap["Wasted Time"] = oppHex;
    }

    // --- BUSCA OS DADOS DA TIMELINE ---
    const lastRow = sheet.getLastRow();
    let timeline = [];
    if (lastRow >= 2) {
      // Agora pegamos da Coluna A (1) até a Coluna N (14)
      const data = sheet.getRange(2, 1, lastRow - 1, 14).getValues();
      timeline = data.map(row => {
        
        let rawDay = row[7]; // Coluna H: Dia (Ex: 'Seg')
        let safeDayAbbr = "";
        
        // Blindado: Data Objeto vs Variações de Texto
        if (Object.prototype.toString.call(rawDay) === '[object Date]' && !isNaN(rawDay.getTime())) {
          const ptDays = ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"];
          safeDayAbbr = ptDays[rawDay.getDay()];
        } else {
          const map = {
            sunday: "Dom", sun: "Dom", dom: "Dom", domingo: "Dom",
            monday: "Seg", mon: "Seg", seg: "Seg", segunda: "Seg",
            tuesday: "Ter", tue: "Ter", ter: "Ter", terca: "Ter", terça: "Ter",
            wednesday: "Qua", wed: "Qua", qua: "Qua", quarta: "Qua",
            thursday: "Qui", thu: "Qui", qui: "Qui", quinta: "Qui",
            friday: "Sex", fri: "Sex", sex: "Sex", sexta: "Sex",
            saturday: "Sáb", sat: "Sáb", sab: "Sáb", sabado: "Sáb", sáb: "Sáb"
          };
          let clean = String(rawDay).replace(/'/g, "").trim().toLowerCase();
          safeDayAbbr = map[clean] || String(rawDay).replace(/'/g, "").trim();
        }

        // Tratamento crucial: converter vírgula europeia para ponto americano ("0,5" -> "0.5")
        let durationVal = String(row[12]).replace(',', '.');

        return {
          action: String(row[0] || "").trim(),
          context: String(row[1] || "").trim(),
          mode: String(row[2] || "").trim(),   // NOVO MAPEAMENTO: Mode
          year: String(row[3]).replace(/'/g, "").trim(),
          month: String(row[5]).replace(/'/g, "").trim(), // Lemos pelo Nome do Mês
          dayAbbr: safeDayAbbr, 
          week: String(row[8]).replace(/'/g, "").trim(),
          fullDate: String(row[9]).trim(),
          duration: parseFloat(durationVal) || 0,
          count: parseInt(row[13]) || 0
        };
      }).filter(item => item.duration > 0 && item.action !== ""); 
    }

    return { timeline: timeline, colors: colorMap };
    
  } catch (e) {
    throw new Error(e.toString());
  }
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
    const actName = String(data[i][17] || "").trim(); // R
    let modeName = String(data[i][18] || "").trim(); // S

    // 🛡️ SANITIZAÇÃO: Morte do "-"
    if (modeName === "-" || modeName.toLowerCase() === "sem categoria") modeName = "";
    
    const ctxLower = ctxName.toLowerCase();

    if (ctxName && !actName) {
      if (PROTECTED_CTX.includes(ctxLower)) continue; 
      let id = 'c_' + i;
      state.contexts.push({ id: id, name: ctxName, colorName: colorName || 'Pattern', mode: modeName });
      ctxMap[ctxLower] = id;
    } else if (ctxName && actName) {
      if (PROTECTED_CTX.includes(ctxLower)) continue; 
      let parentId = ctxMap[ctxLower] || ('c_unknown_' + i); 
      state.actions.push({ id: 'a_' + i, contextId: parentId, name: actName, modeOverride: modeName, originalCtxName: ctxName });
    }
  }
  return JSON.stringify(state);
}

function saveSingleActionMode(actionName, newMode) {
  if (!newMode || newMode === "-") throw new Error("Modo inválido.");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Palette Entry");
  const data = sheet.getRange(1, 18, Math.max(sheet.getLastRow(), 32), 2).getValues(); 

  for (let i = 31; i < data.length; i++) {
    if (String(data[i][0]).trim() === actionName) {
      sheet.getRange(i + 1, 19).setValue(newMode); 
      PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', Date.now().toString());
      return "SUCESSO";
    }
  }
  throw new Error("Ação não encontrada na planilha.");
}

function saveSystemConfigToPalette(jsonString) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Palette Entry");
  if (!sheet) throw new Error("Aba 'Palette Entry' não encontrada.");

  const state = JSON.parse(jsonString);
  const PROTECTED_CTX = ["start", "end", "entrada", "saída", "suporte", "desperdício", "recuperação", "refeição", "unwind"];

  // 1. GUARDA OS FIXOS (Com travas contra Undefined)
  const currentData = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 2), 19).getValues();
  let protectedContexts = [], protectedActions = [];

  for (let i = 1; i < currentData.length; i++) {
    const ctxName = String(currentData[i][13] || "").trim(); // N
    if (ctxName && PROTECTED_CTX.includes(ctxName.toLowerCase())) {
      const colorName = String(currentData[i][14] || "Pattern").trim(); // O
      const modeName = String(currentData[i][18] || "-").trim(); // S
      const actionName = String(currentData[i][17] || "").trim(); // R

      if (!actionName) {
        protectedContexts.push([ctxName, colorName, modeName]);
      } else {
        protectedActions.push([ctxName, colorName, actionName, modeName]);
      }
    }
  }

  // 2. LIMPEZA CIRÚRGICA (Não toca nas fórmulas da col P e Q)
  sheet.getRange("N2:O30").clearContent(); 
  sheet.getRange("S2:S30").clearContent();
  
  const lastRow = Math.max(sheet.getLastRow(), 32);
  if (lastRow >= 32) { 
    sheet.getRange("N32:O" + lastRow).clearContent(); 
    sheet.getRange("R32:S" + lastRow).clearContent(); 
  }

  // 3. GRAVAÇÃO DE CONTEXTOS (Com blindagem contra valores nulos)
  let ctxNO = [], ctxS = [];
  protectedContexts.forEach(pc => { 
    ctxNO.push([pc[0] || "", pc[1] || "Pattern"]); 
    ctxS.push([pc[2] || ""]); 
  });
  
  state.contexts.forEach(c => { 
    ctxNO.push([c.name || "", c.colorName || "Pattern"]); 
    ctxS.push([c.mode || ""]); 
  });
  
  if (ctxNO.length > 0) { 
    sheet.getRange(2, 14, ctxNO.length, 2).setValues(ctxNO); 
    sheet.getRange(2, 19, ctxS.length, 1).setValues(ctxS); 
  }

  // 4. GRAVAÇÃO DE AÇÕES (Com blindagem contra valores nulos)
  let actNO = [], actRS = [];
  protectedActions.forEach(pa => { 
    actNO.push([pa[0] || "", pa[1] || "Pattern"]); 
    actRS.push([pa[2] || "", pa[3] || ""]); 
  });
  
  state.actions.forEach(a => {
    let parentCtx = state.contexts.find(c => c.id === a.contextId);
    let pName = parentCtx ? parentCtx.name : (a.originalCtxName || "Desconhecido");
    let pColor = parentCtx ? parentCtx.colorName : "Pattern";
    
    actNO.push([pName || "", pColor || "Pattern"]); 
    actRS.push([a.name || "", a.modeOverride || ""]);
  });
  
  if (actNO.length > 0) { 
    sheet.getRange(32, 14, actNO.length, 2).setValues(actNO); 
    sheet.getRange(32, 18, actRS.length, 2).setValues(actRS); 
  }

  // 5. ATUALIZAÇÃO E ESTABILIDADE: Passa o controle para o Mutex
  // Isso garante que não teremos múltiplas repinturas simultâneas
  if (typeof triggerSafeGlobalRepaint === 'function') {
    triggerSafeGlobalRepaint();
  } else {
    SpreadsheetApp.flush(); 
    PropertiesService.getDocumentProperties().setProperty('SYS_VERSION', Date.now().toString());
  }

  return "SUCESSO";
}