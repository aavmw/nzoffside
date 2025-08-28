function getOperationInWork() {

  if (!check_input()) return showError('Wrong list or you cant perform such operation');
  if (!checkPrevOperationClosed()) return showError('Previous operations is not closed');

  let data = getCurrentOperationJobCardData(true);
  if (data.endDateTime) return showError('Operation already closed');

  data.startDateTime = nowDateString();

  const payload = {
    email: data.email,
    jobCardCode: data.jobCardCode,
    operation: data.operation,
    startDateTime: data.startDateTime
  }

  return sendToDbAndColorAfterJobEdit(payload)
}

function closeOperation() {
  if (!check_input()) return showError('Wrong list or you cant perform such operation');
  if (!checkPrevOperationClosed()) return showError('Previous operations is not closed');

  return openDurationDialogForClose();
}

function openDurationDialogForClose() {
  const tpl = HtmlService.createTemplateFromFile('used_manhours'); // file below
  const html = tpl.evaluate().setWidth(320).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter Duration (HH:MM)');
}

function durationRequestAndStore(used_manhours) {
  console.log(used_manhours);
  const data = getCurrentOperationJobCardData(true);
  const now = nowDateString();

  const mins = Math.max(0, parseInt(used_manhours, 10) || 0);

  // Base payload for closing
  const payload = {
    email: data.email,
    jobCardCode: data.jobCardCode,
    operation: data.operation,
    used_mnhrs: mins,
    endDateTime: now,
  };

  if (!data.startDateTime) {
    payload.startDateTime = now;
  }

  return sendToDbAndColorAfterJobEdit(payload);

}


function processForm(data) {

  let payload = {
    email: data.email,
    jobCardCode: data.jobCardCode,
    operation: data.operation
  };

  // Add only non-null / non-empty startDateTime, endDateTime
  ['startDateTime', 'endDateTime'].forEach(key => {
    if (data[key] != null && data[key] !== '') {
      payload[key] = data[key];
    }
  });

  return sendToDbAndColorAfterJobEdit(payload);

}

function dropOperationTimes() {

  if (!isUserAdmin()) {
    return showError('Only ADMIN can perform this operation')
  }

  if (is_operation(Ctx.getRangeSafe().getValue())) {

    data = getCurrentOperationJobCardData(false);

    const payload = {
          email: null,
          jobCardCode: data.jobCardCode,
          operation: data.operation,
          startDateTime : data.startDateTime,
          endDateTime : data.endDateTime
    };

    sendToDbAndColorAfterJobEdit(payload);

  } else {
    showError('Cell has no operation in it!');
  }

}

function sendToDbAndColorAfterJobEdit(data) {

  const { row, col } = Ctx.getRowCol();

  // Decide the status once
  const status =
    data.endDateTime ? 'completed' :
    data.startDateTime ? 'in_work' :
    null;

  // Do the DB call and color only after it succeeds
  const p = Promise.resolve(send_to_db(data))
    .then(() => {
      // Color current cell
      if (status) {
        color_current_cell(row, col, status);
      } else {
        color_current_cell(row, col); // clear/default
      }

      if (status === 'completed') {
        if (data.operation === 'PPK') {
          // Special case: close JC color
          close_jc_color(row);
        } else if (Ctx.getCellColor(row, col + 1) === '#ffffff') {
          // Otherwise, prepare the next operation
          color_current_cell(row, col + 1, 'pending');
        }
      }
    })
    .catch(err => {
      showError(err?.message || 'Failed to update operation');
      throw err;
    });

  return p; // safe to ignore if you don’t need chaining
}

function openEditJobCardForm() {

  if (isUserAdmin() && check_input()) {
    const html = HtmlService.createHtmlOutputFromFile('operation_form_filled')
      .setWidth(600)
      .setHeight(500)
    SpreadsheetApp.getUi().showModalDialog(html, 'Operation Tracker');
  } else {
    showError('Only ADMIN can perform this operation');
  }

}

function getFormData() {
  return getCurrentOperationJobCardData(true);
}