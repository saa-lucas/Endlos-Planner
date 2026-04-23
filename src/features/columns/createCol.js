function adicionarColunasLimpias() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Organisieren");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Aba 'Organisieren' não encontrada.");
    return;
  }

  const numCols = 18; // Tamanho do seu bloco (1 semana)
  const maxCols = sheet.getMaxColumns(); // Pega a última coluna que existe na planilha

  // 1. Insere 18 colunas novas no final
  sheet.insertColumnsAfter(maxCols, numCols);

  // 2. Garante que estão completamente limpas (sem herdar formatação ou mesclagem)
  const novaArea = sheet.getRange(1, maxCols + 1, sheet.getMaxRows(), numCols);
  novaArea.clear(); // Limpa dados, cores e formatações
  novaArea.breakApart(); // Desfaz qualquer célula mesclada que possa ter vindo de brinde

  // 3. Copia as larguras exatas do bloco anterior
  const inicioBlocoAnterior = maxCols - numCols + 1;
  const inicioNovoBloco = maxCols + 1;

  for (let i = 0; i < numCols; i++) {
    const larguraOriginal = sheet.getColumnWidth(inicioBlocoAnterior + i);
    sheet.setColumnWidth(inicioNovoBloco + i, larguraOriginal);
  }

  // Aviso discreto no final
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Bloco de 18 colunas criado com larguras ajustadas!", "Sucesso", 5);
}