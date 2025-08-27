// server_request.js — refactored to reuse apiFetch_ + CFG

// Base segment for this service
const WSOP_BASE = '/wsop';
const GO_PATH   = `${WSOP_BASE}/go`;

/**
 * POST an action to /wsop/go
 * @param {string} action
 * @param {object} data
 * @returns {object} parsed JSON
 */
function wsopGo_(action, data = {}) {
  return apiFetch_(GO_PATH, {
    method: 'post',
    body: { action, data }
  });
}

/**
 * Generic GET under /wsop
 * @param {string[]} parts - path parts after /wsop
 * @returns {object} parsed JSON
 */
function wsopGet_(...parts) {
  // Ensure we join with exactly one slash between segments
  const path = [WSOP_BASE, ...parts.map(String)].join('/').replace(/\/{2,}/g, '/');
  return apiFetch_(path, { method: 'get' });
}

/* === Public convenience wrappers (stable names kept) === */

function color_current_cell(row_number, col_number, status) {
  return wsopGo_('color', { row: row_number, col: col_number, status });
}

function update_db_from_drive() {
  return wsopGo_('db_upd');
}

function update_projects_sheet() {
  return wsopGo_('prj_upd');
}

function update_master_from_db() {
  return wsopGo_('mstr_upd');
}

function close_jc_color(row_number) {
  return wsopGo_('jc_clr', { row: row_number});
}

function place_note_in_cell(row_number, col_number, note = null) {
  return wsopGo_('cell_note', { row: row_number, col: col_number, note });
}

function send_to_db(data) {
  // Preserving the original semantics: POST to /wsop with the payload
  try { save_google_log(data); } catch { /* pass */ }
  return apiFetch_(WSOP_BASE, { method: 'post', body: data });
}

function get_single_operation_data(drive_id, operation) {
  return wsopGet_(drive_id, operation);
}

function get_job_card_database_data(drive_id) {
  return wsopGet_(drive_id);
}
