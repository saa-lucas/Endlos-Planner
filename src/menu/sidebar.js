function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Schedule Selector')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}