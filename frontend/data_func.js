function call_dataflow_update() {

  update_db_from_drive();
  update_master_from_db();
  update_projects_sheet();

}

function getCurrentOperationJobCardData(with_operation_dttm, col_number = null) {
  const sheet = Ctx.getSheet();
  const { row, col } = Ctx.getRowCol();

  // if caller didn't pass a column, use the active column
  const targetCol = col_number ?? col;
  if (!row || !targetCol) {
    return {
      email: Ctx.getUserEmail(),
      project: null,
      jobCardName: null,
      jobCardCode: null,
      operation: null,
      startDateTime: null,
      endDateTime: null,
      comment: null,
    };
  }

  // header indexes
  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  const projectIdx = header.indexOf('project');   // 0-based
  const nameIdx    = header.indexOf('name');      // 0-based

  // row data
  const rowVals = sheet.getRange(row, 1, 1, lastCol).getValues()[0] || [];

  // spreadsheetId from HYPERLINK in the name cell (if any)
  let spreadsheetId = null;
  if (nameIdx >= 0) {
    const nameCell = sheet.getRange(row, nameIdx + 1); // Sheets is 1-based
    const formula = nameCell.getFormula();
    if (formula && formula.startsWith('=HYPERLINK(')) {
      const m = formula.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (m) spreadsheetId = m[1];
    }
  }

  // current operation value from the target column
  const operationValue = rowVals[targetCol - 1] ?? null;

  // defaults
  let start_dttm = null;
  let end_dttm   = null;
  let comment    = null;

  // optionally enrich with operation start/end/comment via your existing function
  if (with_operation_dttm && spreadsheetId && operationValue) {
    try {
      // your existing helper; assumed to return JSON string or object like:
      // { status: 'success', data: { start_dttm, end_dttm, comment } }
      const raw = get_single_operation_data(spreadsheetId, operationValue);
      const op = typeof raw === 'string' ? JSON.parse(raw) : raw;

      if (op && op.status === 'success' && op.data) {
        const d = op.data;
        start_dttm = d.start_dttm ?? null;
        end_dttm   = d.end_dttm ?? null;
        comment    = (typeof d.comment === 'string' && d.comment.trim()) ? d.comment : null;
      }
    } catch (e) {
      // swallow parsing/network errors; keep defaults
      // Logger.log('get_single_operation_data error: ' + e);
    }
  }

  return {
    email: Ctx.getUserEmail(),
    project: projectIdx >= 0 ? rowVals[projectIdx] : null,
    jobCardName: nameIdx >= 0 ? rowVals[nameIdx] : null,
    jobCardCode: spreadsheetId,
    operation: operationValue,
    startDateTime: start_dttm,
    endDateTime: end_dttm,
    comment: comment,
  };
}


function save_google_log(data) {

  const log_sheet = Ctx.getSpreadsheet().getSheetByName('log');
  log_sheet.appendRow([new Date(), data]);

}