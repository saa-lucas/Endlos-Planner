function setupDashboardCharts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard");
  const datenSheet = ss.getSheetByName("Daten");
  if (!datenSheet) return;

  const charts = dashSheet.getCharts();
  charts.forEach(c => dashSheet.removeChart(c));
  dashSheet.clear().clearFormats();

  // Deteção real de dados
  const lastA = datenSheet.getRange("A:A").getValues().filter(r => r[0] !== "").length;
  const lastD = datenSheet.getRange("D:D").getValues().filter(r => r[0] !== "").length;

  if (lastA <= 1) return;

  dashSheet.getRange("B2").setValue("PERSONAL PERFORMANCE ANALYTICS")
    .setFontSize(20).setFontWeight("bold").setFontColor("#122633");

  // --- GRÁFICO 1: CONTEXTOS (ROSCA) ---
  const contextRange = datenSheet.getRange(1, 4, lastD, 2);
  const chart1 = dashSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(contextRange)
    .setPosition(4, 2, 0, 0)
    .setOption('title', 'Esforço por Contexto (%)')
    .setOption('pieHole', 0.5)
    .setOption('colors', ['#122633', '#59b3f4', '#3da3a3', '#4b89ff', '#a68aff'])
    .setOption('chartArea', {width: '90%', height: '80%'})
    .setNumHeaders(1) // CRUCIAL: Diz que a linha 1 é cabeçalho
    .build();
  dashSheet.insertChart(chart1);

  // --- GRÁFICO 2: TOP 10 ACTIONS (BARRAS COM VALORES) ---
  const actionRange = datenSheet.getRange(1, 1, Math.min(lastA, 11), 2);
  const chart2 = dashSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(actionRange)
    .setPosition(4, 9, 0, 0)
    .setOption('title', 'Top 10 Ações (Horas)')
    .setOption('legend', {position: 'none'})
    .setOption('annotations', {alwaysOnTop: true, textStyle: {bold: true}})
    .setOption('hAxis', {gridlines: {color: 'none'}, textStyle: {color: 'none'}})
    .setOption('colors', ['#59b3f4'])
    .setOption('chartArea', {left: 140, width: '80%', height: '80%'})
    .setNumHeaders(1) // CRUCIAL
    .build();
  dashSheet.insertChart(chart2);

  dashSheet.activate();
}