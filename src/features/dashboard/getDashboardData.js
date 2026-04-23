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
      // 👇 AQUI ESTAVA O BUG: Mudamos de ler da Coluna 1 (A) para a Coluna 9 (I)
      const data = sheet.getRange(2, 9, lastRow - 1, 14).getValues();
      timeline = data.map(row => {
        
        let rawDay = row[7]; // Agora row[7] mapeia perfeitamente para a Coluna P
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