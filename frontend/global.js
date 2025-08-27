const CFG = {
  get apiBase() {
    return PropertiesService.getScriptProperties().getProperty('API_BASE_URL');
  },
  get apiKey() {
    return PropertiesService.getScriptProperties().getProperty('API_KEY');
  },
  get tz() {
    // prefer a property; fall back to the spreadsheet's timezone; then to UTC
    return PropertiesService.getScriptProperties().getProperty('SHEETS_TZ')
        || SpreadsheetApp.getActive().getSpreadsheetTimeZone()
        || 'Etc/UTC';
  },
};


const Ctx = {
  getSpreadsheet() { return SpreadsheetApp.getActive(); },
  getSheet() { return SpreadsheetApp.getActiveSheet(); },
  getRangeSafe() {
    try { return SpreadsheetApp.getActiveRange(); } catch (e) { return null; }
  },
  getUserEmail() {
    // ActiveUser often empty on consumer accounts; EffectiveUser is more reliable
    const u = Session.getActiveUser().getEmail();
    return u || Session.getEffectiveUser().getEmail();
  },
  getRowCol() {
    const r = this.getRangeSafe();
    return r ? { row: r.getRow(), col: r.getColumn() } : { row: null, col: null };
  },
  getCellColor(row, col) {
    return this.getSheet().getRange(row, col).getBackground();
  },
};


function onOpen() {

    const ui = SpreadsheetApp.getUi();

    // Define menus in a data structure
    const menus = [
        {
        name: 'Data',
        items: [
            ['Update data', 'call_dataflow_update'],
            ['Job card info', 'fetchAndShowData'],
            ['Instruction', 'openInstruction'],
        ]
        },
        {
        name: 'Operation',
        items: [
            ['Get in work', 'getOperationInWork'],
            ['Close operation', 'closeOperation'],
            ['Edit comment', 'openCommentDialog'],
            ['Edit operation', 'openEditJobCardForm'],
            ['Clear operation', 'dropOperationTimes'],
        ]
        }
    ];

    // Create menus
    menus.forEach(menuDef => {
        const menu = ui.createMenu(menuDef.name);
        menuDef.items.forEach(([label, func]) => menu.addItem(label, func));
        // Check if adminItems exists before iterating
        (menuDef.adminItems || []).forEach(([label, func]) => menu.addItem(label, func));
        menu.addToUi();
    });
}

function nowDateString() {

  const now = new Date();
  const formattedDate = now.toLocaleString('en-CA', {
    timeZone: 'Europe/Istanbul',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).replace(', ', 'T');

  return formattedDate

}

function getUserEmail() {
  return Ctx.getUserEmail()
}

function getMasterSheetCellColor(row, col) {
  return Ctx.getCellColor(row, col)
}