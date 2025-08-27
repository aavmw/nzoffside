/* Setup function: ensure required sheets exist  */
function setupNzoffside() {
  const required = ['master', 'users', 'projects', 'log'];
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  required.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    } else {
      // clear existing content if sheet already exists
      sheet.clear();
    }
    // optional: set a header row placeholder
    sheet.getRange(1, 1).setValue(`# ${name.toUpperCase()} sheet`);
  });

  SpreadsheetApp.getUi().alert('Setup complete: master, users, projects, log created/cleared.');
}

/* Resolve library symbol (prefer nzoffside_front, fallback to nzoffside) */
const LIB = (typeof nzoffside_front !== 'undefined' && nzoffside_front)
         || (typeof nzoffside !== 'undefined' && nzoffside)
         || null;

function BRIDGE_requireLib_() {
  if (!LIB) {
    throw new Error(
      'Library not found. Ensure you added it with identifier "nzoffside_front" ' +
      'or "nzoffside" in Extensions → Apps Script → Libraries.'
    );
  }
  return LIB;
}

/* ============== Optional: menu hook ============== */

function onOpen() { return BRIDGE_requireLib_().onOpen(); }

/* ============== Simple one-to-one forwards ============== */
function openInstruction() { return BRIDGE_requireLib_().openInstruction(); }
function openCommentDialog() { return BRIDGE_requireLib_().openCommentDialog(); }
function getCommentContext() { return BRIDGE_requireLib_().getCommentContext(); }
function saveOperationComment(message) { return BRIDGE_requireLib_().saveOperationComment(message); }
function deleteOperationComment() { return BRIDGE_requireLib_().deleteOperationComment(); }
function openJobCardInfoDialog() { return BRIDGE_requireLib_().openJobCardInfoDialog(); }
function getFormData() { return BRIDGE_requireLib_().getFormData(); }
function getUserEmail() { return BRIDGE_requireLib_().getUserEmail(); }
function processForm(data) { return BRIDGE_requireLib_().processForm(data); }
function call_dataflow_update() { return BRIDGE_requireLib_().call_dataflow_update(); }
function fetchAndShowData() { return BRIDGE_requireLib_().fetchAndShowData(); }
function getOperationInWork() { return BRIDGE_requireLib_().getOperationInWork(); }
function closeOperation() {return BRIDGE_requireLib_().closeOperation(); }
function dropOperationTimes() { return BRIDGE_requireLib_().dropOperationTimes(); }
function openEditJobCardForm() { return BRIDGE_requireLib_().openEditJobCardForm(); }