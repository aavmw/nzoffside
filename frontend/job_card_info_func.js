function fetchAndShowData() {

  if (Ctx.getSheet().getName() !== 'master') {
    showError('Change list to master');
    return
  }

  const drive_id = getCurrentOperationJobCardData(false).jobCardCode;
  res = get_job_card_database_data(drive_id);
  showJobCardInfo(res.data);
  
}

function showJobCardInfo(data) {

  const baseHeight = 300; // Base height for static content
  const rowHeight = 60;   // Height per operation row
  const opCount = Object.keys(data.operations || {}).length;
  const extraHeight = opCount * rowHeight;
  const totalHeight = Math.min(baseHeight + extraHeight, 800); // max 800px for safety

  const template = HtmlService.createTemplateFromFile('job_card_info');
  template.data = data;

  const html = template.evaluate()
    .setWidth(900)
    .setHeight(totalHeight);

  SpreadsheetApp.getUi().showModalDialog(html, 'Job Card Info');

}