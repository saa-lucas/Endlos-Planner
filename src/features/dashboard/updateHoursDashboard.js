// ==========================================
// O MOTOR DE DADOS PRINCIPAL (HIERARQUIA MODE)
// ==========================================

function updateHoursDashboard(mode = "LAST_WEEK") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const safeMode = String(mode).trim().toUpperCase();
  
  const sheetOrg = ss.getSheetByName("Organisieren");
  const sheetPal = ss.getSheetByName("Palette Entry");
  
  if (!sheetOrg || !sheetPal) return;

  let sheetDash = ss.getSheetByName("Daten");
  if (!sheetDash) sheetDash = ss.insertSheet("Daten");

  const props = PropertiesService.getDocumentProperties();
  const isPostMeal1hMode = props.getProperty("POST_MEAL_1H_MODE") !== "false";
  const isUnwind1hMode = props.getProperty("UNWIND_1H_MODE") !== "false";

  // ==========================================
  // 🛑 DICIONÁRIO DO SISTEMA (CONTROLE CENTRAL)
  // ==========================================
  const SYS = {
    start: "start",               
    end: "end",                   
    meal: "refeição",         
    unwind: "unwind",         
    leisure: "lazer",
    
    // Saídas automáticas do sistema e seus MODES rígidos
    wastedOut: "Distração", 
    wastedMode: "Desperdício",

    sleepOutAction: "Sono (Calculado)",
    sleepOutContext: "Sono",
    sleepMode: "Recuperação",
    
    uncategorized: "Sem Categoria",
    lazer: "leisure"
  };

  const { months, days } = getDateConstants();

  // ==========================================
  // 🧠 LEITURA DA PALETA (AGORA COM MODE)
  // ==========================================
  const palData = sheetPal.getDataRange().getValues();
  const validEntries = new Set();
  const hierarchy = {}; 

  for (let i = 0; i < palData.length; i++) {
    const ctxRaw = palData[i][13]; // N: Contexto
    const subRaw = palData[i][17]; // R: Ação
    const modeRaw = palData[i][18]; // S: Mode (A nova dimensão!)

    const ctx = ctxRaw ? String(ctxRaw).trim() : "";
    const sub = subRaw ? String(subRaw).trim() : "";
    const sysMode = modeRaw ? String(modeRaw).trim() : "";

    // O Contexto define o Mode base
    if (ctx) { 
      validEntries.add(ctx.toLowerCase()); 
      hierarchy[ctx.toLowerCase()] = { visualName: ctx, parent: ctx, mode: sysMode }; 
    }
    // A Ação pode sobrepor o Mode (ex: Reflexões -> Entrada). Se não, herda do Contexto.
    if (sub) { 
      validEntries.add(sub.toLowerCase()); 
      hierarchy[sub.toLowerCase()] = { 
        visualName: sub, 
        parent: ctx || sub, 
        mode: sysMode || (ctx ? hierarchy[ctx.toLowerCase()]?.mode : "")
      }; 
    }
  }

  const orgData = sheetOrg.getDataRange().getValues(); 
  const orgDisplay = sheetOrg.getDataRange().getDisplayValues(); 
  const numCols = orgData[0].length;

  const mergedRanges = sheetOrg.getDataRange().getMergedRanges();
  const mergeMap = {};
  const skipCells = new Set(); 
  for (let i = 0; i < mergedRanges.length; i++) {
    const merge = mergedRanges[i];
    const sR = merge.getRow() - 1; const sC = merge.getColumn() - 1; 
    mergeMap[`${sR}_${sC}`] = merge.getNumRows();
    for (let rr = 0; rr < merge.getNumRows(); rr++) {
      for (let cc = 0; cc < merge.getNumColumns(); cc++) {
        if (rr === 0 && cc === 0) continue; 
        skipCells.add(`${sR + rr}_${sC + cc}`);
      }
    }
  }

  // O CÉREBRO GEOMÉTRICO DAS DATAS
  const blockWidth = 18;
  const dataStartCol = 4; 
  let anchorDateObj = null;
  let anchorAbsDayIndex = 0; 

  for (let c = dataStartCol; c < numCols; c++) {
    if (orgData[6] && orgData[6][c] instanceof Date) {
      let safeD = new Date(orgData[6][c].getTime() + (12 * 60 * 60 * 1000));
      anchorDateObj = safeD;
      
      let wIdx = Math.floor((c - dataStartCol) / blockWidth);
      let dIdx = Math.floor(((c - dataStartCol) % blockWidth) / 2);
      if (dIdx > 6) dIdx = 6; 
      
      anchorAbsDayIndex = (wIdx * 7) + dIdx; 
      break;
    }
  }

  if (!anchorDateObj) {
    for (let c = dataStartCol; c < numCols; c++) {
      let txt = String(orgDisplay[6][c]).trim();
      let match = txt.match(/(\d{1,2})\/(\d{1,2})/); 
      if (match) {
        let m = parseInt(match[1]) - 1; 
        let d = parseInt(match[2]);
        let currYear = new Date().getFullYear();
        anchorDateObj = new Date(currYear, m, d, 12, 0, 0); 
        
        let wIdx = Math.floor((c - dataStartCol) / blockWidth);
        let dIdx = Math.floor(((c - dataStartCol) % blockWidth) / 2);
        if (dIdx > 6) dIdx = 6;
        anchorAbsDayIndex = (wIdx * 7) + dIdx;
        break;
      }
    }
  }

  if (!anchorDateObj) anchorDateObj = new Date();

  let absoluteStartDate = new Date(anchorDateObj.valueOf());
  absoluteStartDate.setDate(absoluteStartDate.getDate() - anchorAbsDayIndex);

  const colDates = [];
  for (let c = 0; c < numCols; c++) {
    if (c < dataStartCol) {
      colDates[c] = null;
    } else {
      let wIdx = Math.floor((c - dataStartCol) / blockWidth);
      let dIdx = Math.floor(((c - dataStartCol) % blockWidth) / 2);
      
      if (dIdx > 6) {
        colDates[c] = null; 
      } else {
        let absDay = (wIdx * 7) + dIdx;
        let colD = new Date(absoluteStartDate.valueOf());
        colD.setDate(colD.getDate() + absDay);
        colDates[c] = colD;
      }
    }
  }

  // A GUILHOTINA DINÂMICA
  let maxLoggedDate = null;
  for (let c = 2; c < numCols; c++) {
    let tempAtiva = colDates[c];
    if (tempAtiva) {
      for (let r = 8; r < orgDisplay.length; r++) {
        if (skipCells.has(`${r}_${c}`)) continue; 
        let val = String(orgDisplay[r][c]).trim().toLowerCase();
        let isDash = val.length > 0 && val.replace(/[-\s]/g, '') === '';
        
        if (val !== '' && !isDash) {
          let isSystemTemplate = false;
          if (validEntries.has(val)) {
            let pCtx = hierarchy[val].parent.toLowerCase();
            if (pCtx === SYS.start || pCtx === SYS.end || pCtx === SYS.meal || pCtx === SYS.unwind) isSystemTemplate = true;
          } else {
            if (val === SYS.start || val === SYS.end || val === SYS.meal || val === SYS.unwind) isSystemTemplate = true;
          }

          if (!isSystemTemplate) {
            if (!maxLoggedDate || tempAtiva > maxLoggedDate) {
              maxLoggedDate = new Date(tempAtiva.valueOf());
            }
            break; 
          }
        }
      }
    }
  }

  if (!maxLoggedDate) maxLoggedDate = new Date(); 

  let dayOfWeek = maxLoggedDate.getDay(); 
  let diffToSunday = (dayOfWeek === 0) ? 0 : (7 - dayOfWeek);
  let tetoDomingo = new Date(maxLoggedDate.valueOf());
  tetoDomingo.setDate(tetoDomingo.getDate() + diffToSunday);
  tetoDomingo.setHours(23, 59, 59, 999);

  const isColVisible = [];
  const limitePassado = new Date();
  limitePassado.setDate(limitePassado.getDate() - 20); 

  for (let c = 0; c < numCols; c++) {
    let dataAtiva = colDates[c];
    if (dataAtiva && dataAtiva > tetoDomingo) {
      isColVisible[c] = false;
      continue;
    }
    if (mode === "FULL_SYNC") {
      isColVisible[c] = true; 
    } else if (mode === "LAST_WEEK") {
      isColVisible[c] = dataAtiva ? (dataAtiva >= limitePassado) : false;
    } else {
      isColVisible[c] = !sheetOrg.isColumnHiddenByUser(c + 1); 
    }
  }

  function getExactTime(rIdx) {
    if (rIdx >= orgDisplay.length) return "--:--";
    let rUp = rIdx;
    while (rUp >= 8) {
      let val = orgDisplay[rUp][1]; 
      if (val) {
        let tm = String(val).match(/(\d{1,2})[:h](\d{2})/i);
        if (tm) return `${tm[1].padStart(2, '0')}:${tm[2]}`;
      }
      rUp--;
    }
    return "--:--";
  }

  const activityLog = []; 
  const sleepTracker = {}; 
  const visibleDates = new Set(); 
  const isAwakeForCol = {};

  for (let c = 2; c < numCols; c++) {
    if (isColVisible[c] && colDates[c]) {
      let safeD = colDates[c];
      let dN = String(safeD.getDate()).padStart(2, '0');
      let mN = String(safeD.getMonth() + 1).padStart(2, '0');
      visibleDates.add(`${dN}/${mN}/${safeD.getFullYear()}`);
    }
  }

  // ==========================================
  // O PARSER DE ATIVIDADES
  // ==========================================
  for (let r = 8; r < orgDisplay.length; r++) {
    for (let c = 2; c < orgDisplay[r].length; c++) {
      if (!isColVisible[c] || skipCells.has(`${r}_${c}`)) continue;

      const rawValue = orgDisplay[r][c];
      const clnVal = String(rawValue).trim().toLowerCase();
      const isDash = clnVal.length > 0 && clnVal.replace(/[-\s]/g, '') === '';

      let rName, pName, mName; // Ação, Contexto, Mode

      if (clnVal === "" || isDash) {
        if (isAwakeForCol[c]) {
          rName = SYS.wastedOut; pName = SYS.wastedOut; mName = SYS.wastedMode;
        } else continue; 
      } else if (validEntries.has(clnVal)) {
        const info = hierarchy[clnVal];
        rName = info.visualName;
        pName = hierarchy[info.parent.toLowerCase()] ? hierarchy[info.parent.toLowerCase()].visualName : info.parent;
        mName = info.mode || SYS.uncategorized; // Carimba o Mode!
      } else {
        rName = String(rawValue).trim(); 
        pName = SYS.uncategorized; 
        mName = SYS.uncategorized;
      }

      const safeR = rName.toLowerCase();
      const safeP = pName.toLowerCase();

      let isStartTrigger = (safeR === SYS.start || safeP === SYS.start);
      let isEndTrigger = (safeR === SYS.end || safeP === SYS.end);

      const rowSpan = mergeMap[`${r}_${c}`] || 1;
      let bHours = rowSpan * 0.25;

      if (isStartTrigger) { isAwakeForCol[c] = true; bHours = 0; }
      if (isEndTrigger) { isAwakeForCol[c] = false; bHours = 0; }

      if (isUnwind1hMode && (safeR === SYS.unwind || safeP === SYS.unwind) && rowSpan === 2) {
        bHours = 1.0;
      }

      // LÓGICA DA REFEIÇÃO
      let isPrevMeal = false; 
      let tempR = r - 1;
      while (tempR >= 8) {
        const tVal = String(orgDisplay[tempR][c]).trim().toLowerCase();
        const tDash = tVal.length > 0 && tVal.replace(/[-\s]/g, '') === '';
        const isSkip = skipCells.has(`${tempR}_${c}`);

        if (tVal !== "" && !tDash && !isSkip) {
          if (validEntries.has(tVal)) {
            const tInfo = hierarchy[tVal];
            const tParent = hierarchy[tInfo.parent.toLowerCase()] ? hierarchy[tInfo.parent.toLowerCase()].visualName : tInfo.parent;
            if (tParent.toLowerCase() === SYS.meal || tVal === SYS.meal) isPrevMeal = true;
          }
          break; 
        } 
        else if ((tVal === "" || tDash) && !isSkip) { break; }
        tempR--;
      }
      
      if (isPrevMeal && !SYS.meal.includes(safeP) && rowSpan === 2 && !isEndTrigger && !isStartTrigger) {
        if (safeR === SYS.wastedOut.toLowerCase()) {
          bHours = 1.0;
        } else if (isPostMeal1hMode) {
          bHours = 1.0;
        }
      }

      let rawDateObj = colDates[c];
      let year = "-", mNum = "-", mAbbr = "-", dNum = "-", dAbbr = "-", wNum = "-", fDate = "-";
      let sTime = getExactTime(r);
      let eTime = getExactTime(r + rowSpan);

      if (eTime === "--:--" && sTime !== "--:--") {
        let p = sTime.split(":");
        let totMins = parseInt(p[0]) * 60 + parseInt(p[1]) + (rowSpan * 15);
        eTime = `${String(Math.floor(totMins / 60)).padStart(2, '0')}:${String(totMins % 60).padStart(2, '0')}`;
      }
      
      if (rawDateObj) {
        const dateKey = rawDateObj.getFullYear() + "-" + (rawDateObj.getMonth() + 1).toString().padStart(2, '0') + "-" + rawDateObj.getDate().toString().padStart(2, '0');
        if (!sleepTracker[dateKey]) sleepTracker[dateKey] = { dateObj: rawDateObj, tracked: 0 };
        
        if (isStartTrigger) {
          sleepTracker[dateKey].start = sTime;
          if (bHours > 0) sleepTracker[dateKey].tracked += bHours; 
        } else if (isEndTrigger) {
          sleepTracker[dateKey].end = eTime;
          if (bHours > 0) sleepTracker[dateKey].tracked += bHours; 
        } else if (bHours > 0) {
          sleepTracker[dateKey].tracked += bHours;
        }

        year = "'" + rawDateObj.getFullYear();
        mNum = "'" + (rawDateObj.getMonth() + 1).toString().padStart(2, '0');
        mAbbr = monthsEng[rawDateObj.getMonth()];
        dNum = "'" + rawDateObj.getDate().toString().padStart(2, '0');
        dAbbr = daysEng[rawDateObj.getDay()];
        fDate = dNum.replace("'","") + "/" + mNum.replace("'","") + "/" + year.replace("'","");

        let tdt = new Date(rawDateObj.valueOf());
        let dayn = (tdt.getDay() + 6) % 7; 
        let monD = new Date(tdt.valueOf()); 
        monD.setDate(monD.getDate() - dayn); 
        let sunD = new Date(monD.valueOf()); 
        sunD.setDate(sunD.getDate() + 6); 
        let monStr = monthsEng[monD.getMonth()] + String(monD.getDate()).padStart(2, '0');
        let sunStr = monthsEng[sunD.getMonth()] + String(sunD.getDate()).padStart(2, '0');
        wNum = "'" + monStr + "/" + sunStr;
      }

      const countV = 1;
      // ADICIONAMOS O MODE (mName) AQUI COMO A COLUNA 3
      activityLog.push({ data: [rName, pName, mName, year, mNum, mAbbr, dNum, dAbbr, wNum, fDate, "'" + sTime, "'" + eTime, bHours, countV] });
    }
  }

  const lastRow = sheetDash.getLastRow();
  let dbTimeline = [];
  
  if (mode !== "FULL_SYNC" && lastRow > 1) {
    // Agora puxamos 14 colunas porque adicionamos o Mode!
    dbTimeline = sheetDash.getRange(2, 9, lastRow - 1, 14).getDisplayValues();
  }

  // ==========================================
  // O EXORCISMO DO BANCO (AJUSTADO PARA NOVOS ÍNDICES)
  // ==========================================
  const filteredDb = dbTimeline.filter(row => {
    const rAction = String(row[0]);
    if (rAction === SYS.sleepOutAction) return false; 
    
    // Data completa agora é o índice 9 (porque o Mode empurrou tudo)
    const rDate = String(row[9]).replace(/'/g, "");
    if (rDate === "-" || rDate === "") return false; 

    let p = rDate.split("/");
    if (p.length === 3) {
      let rowDateObj = new Date(p[2], parseInt(p[1])-1, p[0], 12, 0, 0);
      if (rowDateObj > tetoDomingo) return false; 
    }
    
    if (visibleDates.has(rDate)) return false; 
    return true; 
  });

  const combined = filteredDb.concat(activityLog.map(obj => obj.data));
  
  const sortable = combined.map(row => {
    let p = String(row[9]).replace(/'/g, "").split("/"); // Indice 9
    if (p.length < 3) p = [1, 1, 1970]; 
    
    let dObj = new Date(p[2], parseInt(p[1])-1, p[0], 12, 0, 0);
    
    let tm = String(row[10]).replace(/'/g, "").split(":"); // Hora início Indice 10
    let mins = tm.length === 2 ? parseInt(tm[0])*60 + parseInt(tm[1]) : 9999;
    
    let sortMins = mins;
    if (sortMins < 300) sortMins += 1440; 

    return { 
      data: row, 
      ts: dObj.getTime() + (sortMins * 60000), 
      dateObj: dObj, 
      action: String(row[0]), 
      context: String(row[1]), 
      sTime: String(row[10]).replace(/'/g, ""), 
      eTime: String(row[11]).replace(/'/g, "") 
    };
  }).sort((a, b) => a.ts - b.ts);

  const finalTimeline = [];
  let lastSaaEnd = null;

  for (let i = 0; i < sortable.length; i++) {
    let curr = sortable[i];
    
    let isLoopEnd = (curr.action.toLowerCase() === SYS.end || curr.context.toLowerCase() === SYS.end);
    let isLoopStart = (curr.action.toLowerCase() === SYS.start || curr.context.toLowerCase() === SYS.start);
    
    if (isLoopEnd && curr.eTime !== "--:--") {
      lastSaaEnd = curr;
    } else if (isLoopStart && curr.sTime !== "--:--" && lastSaaEnd) {
      let eM = parseInt(lastSaaEnd.eTime.split(":")[0])*60 + parseInt(lastSaaEnd.eTime.split(":")[1]);
      let sM = parseInt(curr.sTime.split(":")[0])*60 + parseInt(curr.sTime.split(":")[1]);
      let sH = (sM >= eM) ? (sM - eM) / 60 : ((1440 - eM) + sM) / 60;

      let d = curr.dateObj;
      let y = "'" + d.getFullYear(); let mN = "'" + String(d.getMonth() + 1).padStart(2, '0');
      let dN = "'" + String(d.getDate()).padStart(2, '0');
      let fD = dN.replace("'","") + "/" + mN.replace("'","") + "/" + y.replace("'","");
      
      let tdt = new Date(d.valueOf()); 
      let dayn = (tdt.getDay() + 6) % 7;
      let monD = new Date(tdt.valueOf()); 
      monD.setDate(monD.getDate() - dayn);
      let sunD = new Date(monD.valueOf()); 
      sunD.setDate(sunD.getDate() + 6);
      
      let monStr = monthsEng[monD.getMonth()] + String(monD.getDate()).padStart(2, '0');
      let sunStr = monthsEng[sunD.getMonth()] + String(sunD.getDate()).padStart(2, '0');
      let wN = "'" + monStr + "/" + sunStr;

      // SLEEP CARIMBADO COMO "RECUPERAÇÃO" NA COLUNA 3
      finalTimeline.push([SYS.sleepOutAction, SYS.sleepOutContext, SYS.sleepMode, y, mN, monthsEng[d.getMonth()], dN, daysEng[d.getDay()], wN, fD, "'" + lastSaaEnd.eTime, "'" + curr.sTime, sH, 1]);
      lastSaaEnd = null;
    }
    
    finalTimeline.push(curr.data);
  }

  const gC = {}, gO = {}, cC = {}, cO = {};
  const outputTimeline = [["Ação", "Contexto", "Mode", "Ano", "Mês #", "Mês", "Dia #", "Dia", "Semana #", "Data Completa", "Hora Início", "Hora Fim", "Duração (h)", "Contagem"]];

  finalTimeline.forEach(row => {
    outputTimeline.push(row);
    let act = row[0], ctx = row[1], dur = parseFloat(row[12]) || 0, cnt = parseInt(row[13]) || 0; // Índices 12 e 13 agora!
    if (dur > 0) {
      gC[act] = (gC[act] || 0) + dur; gO[act] = (gO[act] || 0) + cnt;
      cC[ctx] = (cC[ctx] || 0) + dur; cO[ctx] = (cO[ctx] || 0) + cnt;
    }
  });

  const outG = [["Ação Específica", "Total Horas", "Contagem"]];
  for (let k in gC) outG.push([k, gC[k], gO[k]]); outG.sort((a, b) => b[1] - a[1]);

  const outC = [["Contexto (Agrupado)", "Total Horas", "Contagem"]];
  for (let k in cC) outC.push([k, cC[k], cO[k]]); outC.sort((a, b) => b[1] - a[1]);

  sheetDash.clear();
  if (outG.length > 1) sheetDash.getRange(1, 1, outG.length, 3).setValues(outG);
  if (outC.length > 1) sheetDash.getRange(1, 5, outC.length, 3).setValues(outC);
  if (outputTimeline.length > 1) {
    sheetDash.getRange(1, 9, outputTimeline.length, 14).setValues(outputTimeline);
    sheetDash.getRange(2, 11, outputTimeline.length - 1, 12).setHorizontalAlignment("center");
  }
  sheetDash.getRange("A1:V1").setFontWeight("bold");
  sheetDash.autoResizeColumns(1, 22);
}

// Mantenha as outras funções `saveToggleSettings` e `getToggleStates` intactas aqui abaixo.