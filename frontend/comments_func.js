/** Open the dialog */
function openCommentDialog() {

  if (!check_input()) return showError('Wrong list or you cant perform such operation');

  const html = HtmlService.createTemplateFromFile('comment_form').evaluate()
      .setWidth(420)
      .setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Operation Comment');

}

/** Provide context for the dialog (which op / job card, etc.) */
function getCommentContext() {
  const ctx = getCurrentOperationJobCardData(true);
  return ctx; 
}

/** Save handler called from the dialog */
function saveOperationComment(commentText) {
  const ctx = getCommentContext();

  const toStr = v => (v == null ? '' : String(v));
  const cleaned = toStr(commentText).trim(); // removes spaces/newlines

  if (!cleaned) {
    // nothing to save
    return { status: 'noop', reason: 'empty_comment' };
  }

  const userPart = String(ctx.email || '').split('@')[0];
  const comment_text = `${nowDateString()} ${userPart}: ${cleaned}`

  const payload = {
    jobCardCode: ctx.jobCardCode,
    operation: ctx.operation,
    comment: comment_text
  };

  const { row, col } = Ctx.getRowCol();
  // backend expects 0-based indices:
  place_note_in_cell(row - 1, col - 1, comment_text);

  send_to_db(payload)
}

function deleteOperationComment() {
  const ctx = getCommentContext();
  const payload = {
    jobCardCode: ctx.jobCardCode,
    operation: ctx.operation,
    comment: null
  };

  send_to_db(payload);

  const { row, col } = Ctx.getRowCol();
  place_note_in_cell(row - 1, col - 1);

}