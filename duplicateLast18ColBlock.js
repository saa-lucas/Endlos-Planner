/**
 * Duplica o último bloco de 18 colunas.
 * Mantém as células de Segunda a Domingo, mas nomeia o botão de Domingo a Sábado.
 */
function duplicateLast18ColBlock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Duplicando o bloco e processando datas...", "Carregando ⏳", -1);
  const sheet = ss.getActiveSheet();
  
  const blockWidth = 18;
  const dataStartCol = 4; // Coluna D
  const fixedLastRow = 105; 
  const weekDayRow = 6;
  const dateRow = 7;
  
  const currentMaxCols = sheet.getMaxColumns();
  let lastBlockCol = 0;
  
  // 1. Localiza o último bloco
  for (let col = dataStartCol; col <= currentMaxCols; col += blockWidth) {
    if (col + blockWidth - 1 <= currentMaxCols) {
      const merges = sheet.getRange(1, col, 5, blockWidth).getMergedRanges();
      if (merges.length > 0) lastBlockCol = col;
    }
  }
  
  if (lastBlockCol === 0) {
    ss.toast("Nenhum bloco encontrado.", "Erro ❌", 5);
    return;
  }
  
  // 2. Criação das novas colunas
  const targetCol = lastBlockCol + blockWidth;
  if (targetCol + blockWidth - 1 > currentMaxCols) {
    sheet.insertColumnsAfter(currentMaxCols, (targetCol + blockWidth - 1) - currentMaxCols);
  }
  
  // 3. Copia a estrutura bruta
  const sourceRange = sheet.getRange(1, lastBlockCol, fixedLastRow, blockWidth);
  const targetRange = sheet.getRange(1, targetCol, fixedLastRow, blockWidth);
  sourceRange.copyTo(targetRange);
  SpreadsheetApp.flush(); 

  // ========================================================
  // 4. INJEÇÃO CIRÚRGICA NAS CÉLULAS (SEGUNDA A DOMINGO)
  // ========================================================
  const oldDates = sheet.getRange(dateRow, lastBlockCol, 1, blockWidth).getValues()[0];
  const oldHeaders = sheet.getRange(weekDayRow, lastBlockCol, 1, blockWidth).getValues()[0];
  const isValidDate = (d) => Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d.getTime());
  
  // Acha a Segunda-feira base do bloco anterior
  let baseMonday = oldDates.find(d => isValidDate(d) && d.getDay() === 1);
  if (!baseMonday) baseMonday = oldDates.find(d => isValidDate(d)); // Fallback

  let newHeaders = new Array(blockWidth).fill("");
  let newDates = new Array(blockWidth).fill("");

  // Variáveis para o nome do Botão (Named Range)
  let buttonStartDate = null;
  let buttonEndDate = null;

  if (baseMonday) {
    // Nova Segunda = +7 dias (Isto vai para a primeira célula)
    let newMonday = new Date(baseMonday.getFullYear(), baseMonday.getMonth(), baseMonday.getDate() + 7, 12, 0, 0);

    // ========================================================
    // 🧠 MÁGICA AQUI: Datas para o Botão (Domingo a Sábado)
    // ========================================================
    buttonStartDate = new Date(newMonday.getTime());
    buttonStartDate.setDate(newMonday.getDate() - 1); // Puxa 1 dia para trás (Domingo)

    buttonEndDate = new Date(newMonday.getTime());
    buttonEndDate.setDate(newMonday.getDate() + 5); // Puxa para o Sábado

    // Índices onde a mesclagem guarda o texto! (Segunda a Domingo)
    const dayIndices = [1, 3, 5, 7, 9, 11, 13];
    
    dayIndices.forEach((idx, offset) => {
      let d = new Date(newMonday.getTime());
      d.setDate(newMonday.getDate() + offset);
      
      newDates[idx] = d; // Injeta a data na célula (Vai do dia 6 ao 12)
      
      let oldHead = oldHeaders[idx];
      if (isValidDate(oldHead)) {
         newHeaders[idx] = d;
      } else if (oldHead !== "" && oldHead !== undefined) {
         newHeaders[idx] = oldHead; 
      } else {
         const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
         newHeaders[idx] = days[d.getDay()]; 
      }
    });
  }

  sheet.getRange(weekDayRow, targetCol, 1, blockWidth).setValues([newHeaders]);
  sheet.getRange(dateRow, targetCol, 1, blockWidth).setValues([newDates]);

  // ========================================================
  // 5. NAMED RANGE PARA O BOTÃO (DOMINGO A SÁBADO)
  // ========================================================
  if (buttonStartDate && buttonEndDate) {
    const tz = Session.getScriptTimeZone();
    const startYear = Utilities.formatDate(buttonStartDate, tz, "yyyy"); 
    const startDateStr = Utilities.formatDate(buttonStartDate, tz, "MMMdd");
    const endDateStr = Utilities.formatDate(buttonEndDate, tz, "MMMdd");
    
    // O nome gerado será ex: Week_2026_Apr05_Apr11
    let rangeName = "Week_" + startYear + "_" + startDateStr + "_" + endDateStr;
    
    try {
      const existingRanges = ss.getNamedRanges();
      for (let r = 0; r < existingRanges.length; r++) {
        if (existingRanges[r].getName() === rangeName) existingRanges[r].remove();
      }
      ss.setNamedRange(rangeName, targetRange);
    } catch(e) {
      console.error("Erro ao definir o intervalo nomeado: " + e);
    }
  }
  
  // 6. Ajusta larguras
  for (let w = 0; w < blockWidth; w++) {
    sheet.setColumnWidth(targetCol + w, sheet.getColumnWidth(lastBlockCol + w));
  }

  ss.toast("Novo bloco gerado com sucesso!", "Concluído ✅", 5);
}