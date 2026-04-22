function toggleSmartMerge() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== "Organisieren") {
      return "ERRO: Use esta ferramenta apenas na aba Organisieren.";
    }

    const activeRange = sheet.getActiveRange();
    let startRow = activeRange.getRow();
    let numRows = activeRange.getNumRows();
    let colStart = activeRange.getColumn();

    // 1. A MATEMÁTICA CORRIGIDA: No seu layout, Main é Par e Side é Ímpar.
    // Se clicar na Ímpar (Side), a Main é a próxima. Se clicar na Par (Main), a Side é a anterior.
    let mainCol = (colStart % 2 === 0) ? colStart : colStart + 1;
    let sideCol = mainCol - 1;

    if (numRows === 1) numRows = 2; // Auto-expansão para o mínimo de 30m

    // 2. Expande a área para englobar os blocos da Main perfeitamente
    let mainTarget = sheet.getRange(startRow, mainCol, numRows, 1);
    let mainMerges = mainTarget.getMergedRanges();
    if (mainMerges.length > 0) {
      let minR = startRow;
      let maxR = startRow + numRows - 1;
      mainMerges.forEach(m => {
        if (m.getRow() < minR) minR = m.getRow();
        if (m.getRow() + m.getNumRows() - 1 > maxR) maxR = m.getRow() + m.getNumRows() - 1;
      });
      startRow = minR;
      numRows = maxR - minR + 1;
    }

    // Refazemos as variáveis com a área real e 100% corrigida
    let finalMain = sheet.getRange(startRow, mainCol, numRows, 1);
    let finalSide = sheet.getRange(startRow, sideCol, numRows, 1);
    
    // Essa é a ZONA TOTAL que engloba a Side e a Main juntas
    let fullZone = sheet.getRange(startRow, sideCol, numRows, 2);

    // 3. Tira a "foto" da estética original para clonar depois
    let mainValue = finalMain.getCell(1, 1).getValue();
    let sideBg = finalSide.getCell(1, 1).getBackground();
    let mainBg = finalMain.getCell(1, 1).getBackground();
    let mainFont = finalMain.getCell(1, 1).getFontColor();

    // 4. Descobre se o bloco já está inteiro (para podermos Dividir na metade)
    let isFullyMerged = false;
    let checkMerges = finalMain.getMergedRanges();
    if (checkMerges.length === 1 && checkMerges[0].getNumRows() === numRows) {
      isFullyMerged = true;
    }

    // >>> O TRATOR: Estilhaça ABSOLUTAMENTE TUDO (Side e Main) na zona alvo <<<
    // Destrói qualquer mesclagem torta antes de reconstruir perfeitamente
    fullZone.breakApart();

    // 5. Constrói a Side e a Main como Gêmeas Idênticas
    if (isFullyMerged) {
      // DESCER O DEGRAU (Dividir ao meio)
      let chunk = Math.floor(numRows / 2);
      for (let i = 0; i < numRows; i += chunk) {
        let actualChunk = Math.min(chunk, numRows - i);
        
        let sR = sheet.getRange(startRow + i, sideCol, actualChunk, 1);
        let mR = sheet.getRange(startRow + i, mainCol, actualChunk, 1);
        
        // Aplica as mesclagens nos novos pedaços
        if (actualChunk > 1) {
          sR.mergeVertically();
          mR.mergeVertically();
        }
        
        // Pinta instantaneamente a Side e a Main clonando o visual
        sR.setBackground(sideBg);
        mR.setBackground(mainBg).setFontColor(mainFont).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
        if (mainValue) mR.setValue(mainValue);
      }
    } else {
      // SUBIR O DEGRAU (Fundir tudo num bloco só)
      finalSide.mergeVertically().setBackground(sideBg);
      finalMain.mergeVertically().setBackground(mainBg).setFontColor(mainFont).setHorizontalAlignment("center").setVerticalAlignment("middle").setFontWeight("bold");
      if (mainValue) finalMain.setValue(mainValue);
    }

    SpreadsheetApp.flush();
    return "SUCESSO";
  } catch (erro) {
    return "ERRO CRÍTICO: " + erro.message;
  }
}