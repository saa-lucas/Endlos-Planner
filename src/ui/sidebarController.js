function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ui/sidebar')
      .setTitle('Schedule Selector')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function saveToggleSettings(mealStatus, unwindStatus, headerStatus, fontStatus, skipUpdateStatus) {
  try {
    var props = PropertiesService.getDocumentProperties();
    
    // Converte os booleanos para strings ('true' ou 'false') e salva
    var m = mealStatus === true ? 'true' : 'false';
    var u = unwindStatus === true ? 'true' : 'false';
    var h = headerStatus === true ? 'true' : 'false';
    var f = fontStatus === true ? 'true' : 'false';
    var s = skipUpdateStatus === true ? 'true' : 'false';

    props.setProperty('isMealSplit', m);
    props.setProperty('isUnwind', u);
    props.setProperty('isHeaderEnabled', h);
    props.setProperty('isColorFont', f);
    props.setProperty('isSkipUpdate', s);
    
    // Log para a gente ver se ele salvou certo
    console.log("SALVOU: Meal=" + m + " Unwind=" + u + " Header=" + h + " Font=" + f + " Skip=" + s);
    
    return "OK"; 
  } catch (e) {
    throw new Error(e.message);
  }
}

function getSettingsState() {
  try {
    var props = PropertiesService.getDocumentProperties();
    
    // Lê tudo do banco de dados (tudo vem como texto!)
    var mStr = props.getProperty('isMealSplit');
    var uStr = props.getProperty('isUnwind');
    var hStr = props.getProperty('isHeaderEnabled');
    var fStr = props.getProperty('isColorFont');
    var sStr = props.getProperty('isSkipUpdate');
    
    // Log para a gente ver o que ele está lendo
    console.log("LEU DO BANCO: Meal=" + mStr + " Unwind=" + uStr + " Header=" + hStr + " Font=" + fStr + " Skip=" + sStr);

    return {
      // Se não existir (null) ou não for explicitamente 'false', assume o padrão de fábrica
      isMealSplit: mStr === 'false' ? false : true, 
      isUnwind: uStr === 'false' ? false : true,       
      isSkipUpdate: sStr === 'false' ? false : true, 
      
      // Cabeçalho e Fonte vêm OFF por padrão
      isHeaderEnabled: hStr === 'true' ? true : false, 
      isColorFont: fStr === 'true' ? true : false  
    };
  } catch(e) {
    console.log("ERRO FATAL NA LEITURA: " + e.message);
    return { isMealSplit: true, isUnwind: true, isSkipUpdate: true, isHeaderEnabled: false, isColorFont: false };
  }
}