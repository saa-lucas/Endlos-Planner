// ==========================================
// MOTOR DO WEB APP (O GERADOR DE LINK)
// ==========================================

function doGet(e) {
  // Aqui ele puxa o visual que nós criamos. 
  // IMPORTANTE: Se o seu arquivo HTML se chamar "clientdashboard.html", mude o nome abaixo para 'clientdashboard'
  var htmlOutput = HtmlService.createTemplateFromFile('dashboard')
      .evaluate()
      .setTitle('📊 Vida em Padrões - Diagnóstico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) 
      .addMetaTag('viewport', 'width=device-width, initial-scale=1'); // Deixa bonito no celular
      
  return htmlOutput;
}

// Como o WebApp roda fora da planilha, a função de tema precisa de um "fallback" 
// para o gráfico não quebrar e manter o seu Dark Mode lindão.
function getCurrentUIState() {
  return { theme: 'DARK' }; 
}