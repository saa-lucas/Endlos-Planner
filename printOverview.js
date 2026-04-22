function printOverview(focusName, includeHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet(); 
  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  // ==========================================
  // LÓGICA DE DIMENSIONAMENTO E EXPORTAÇÃO
  // ==========================================
  const blockWidth = 18;
  const dataStartCol = 4; // Coluna D (início dos dados)
  const fixedLastRow = 105; 
  
  const currentMaxCols = sheet.getMaxColumns();
  let lastVisibleCol = dataStartCol + blockWidth - 1;

  // Localiza o último bloco visível para definir a largura do PDF
  for (let col = dataStartCol; col <= currentMaxCols; col += blockWidth) {
    if (col + blockWidth - 1 <= currentMaxCols) {
      if (!sheet.isColumnHiddenByUser(col)) {
        const merges = sheet.getRange(1, col, 5, blockWidth).getMergedRanges();
        if (merges.length > 0) {
          lastVisibleCol = col + blockWidth - 1; 
        }
      }
    }
  }

  const printRange = sheet.getRange(1, 1, fixedLastRow, lastVisibleCol).getA1Notation();

  let scaleValue, fzcValue;
  
  if (focusName && focusName !== "ALL") {
    scaleValue = 4; // Foco em uma semana
    fzcValue = 'false';
  } else {
    scaleValue = 3; // Visão geral
    fzcValue = 'true';
  }

  let headerParams = "";
  let topMargin = 0.2;
  let bottomMargin = 0.2;

  if (includeHeaders !== false) { 
    headerParams = "&printtitle=true&sheetnames=true&pagenum=RIGHT";
  } else { 
    headerParams = "&printtitle=false&sheetnames=false"; 
    topMargin = 0.6;  
    bottomMargin = 0.6; 
  }

  // Montagem da URL de exportação para PDF
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&gid=${sheetId}&range=${printRange}&size=A4&portrait=true&scale=${scaleValue}&top_margin=${topMargin}&bottom_margin=${bottomMargin}&left_margin=0.2&right_margin=0.2${headerParams}&fzc=${fzcValue}&fzr=false&printnotes=false`;

  try {
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + token }
    });
    
    // Converte para Base64 para exibir no Painel Lateral
    const blob = response.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    
    return base64;

  } catch (error) {
    console.error("Erro ao gerar o PDF do ENDLOS Planner: ", error);
    throw error;
  }
}