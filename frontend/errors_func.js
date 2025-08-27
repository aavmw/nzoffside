function showError(error_message = 'Error') {
  const template = HtmlService.createTemplateFromFile('error');
  template.error_message = error_message;
  const html = template.evaluate()
    .setWidth(400)
    .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Error');
}

function openInstruction() {
    
  var html = HtmlService.createHtmlOutput(`
    <script>
      window.open('https://aavuas.yonote.ru/share/11e01b3a-0cb2-4f73-aa0b-135223f83194', '_blank');
      google.script.host.close();
    </script>
  `).setWidth(10).setHeight(10);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Redirect...');

}