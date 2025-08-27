function isUserAdmin() {
  const values = Sheets.Spreadsheets.Values.get(Ctx.getSpreadsheet().getId(), 'users!A:B').values || [];
  const user = Ctx.getUserEmail();
  return values.some(row => row[0] === user && row[1] === 'ADMIN');
}

function check_input() {
  if (Ctx.getSheet().getName() !== 'master') {
    return false;
  }

  const current_operation = Ctx.getRangeSafe()?.getValue();
  if (!is_operation(current_operation)) {
    return false;
  }

  const values = Sheets.Spreadsheets.Values.get(Ctx.getSpreadsheet().getId(), 'users!A:B').values || [];
  const current_user = Ctx.getUserEmail();

  const user_accesses = values
    .filter(row => row[0] === current_user)
    .map(row => row[1]);

  if (user_accesses.length === 0) {
    return false;
  }

  return user_accesses.some(op => current_operation.includes(op) || op === 'ADMIN');
}

function checkPrevOperationClosed() {
  const { row, col } = Ctx.getRowCol();
  if (!row || !col) return false;

  const sheet = Ctx.getSheet();
  const last_col = sheet.getLastColumn();
  const columns = sheet.getRange(1, 1, 1, last_col).getValues();
  const opIdx = columns[0]?.indexOf('operations');

  if (col - 1 === opIdx) {
    return true;
  }

  const data = getCurrentOperationJobCardData(true, col - 1);
  return data.endDateTime !== null;
}

function is_operation(value) {
  return /^(?:PPK|\d{3}_[A-Z]+(?:_[A-Z]+)*)$/.test((value ?? '').toString().trim());
}

function checkHealthz() {

  const base  = CFG.apiBase;
  const key   = CFG.apiKey;

  if (!base || !key) {
    throw new Error("Missing API_BASE or API_KEY");
  }

  const url = base + '/healthz';

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'X-API-Key': key,
      'ngrok-skip-browser-warning': 'true'
    },
    muteHttpExceptions: true
  });

  return {
    ok: resp.getResponseCode() >= 200 && resp.getResponseCode() < 300,
    status: resp.getResponseCode(),
    body: resp.getContentText()
  };
}