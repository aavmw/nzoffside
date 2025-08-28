# app.py
from flask import Flask, request, jsonify
from ..services import workshop as ws
from . import api_bp

app = Flask(__name__)

@api_bp.get("/ping")
def ping():
    return jsonify(ok=True, msg="pong")

# @api_bp.post("/acdc")
# def execute_code():
#     try:
#         acdc.getFull()
#         return jsonify({"status": "success"}), 200
#     except Exception as e:
#         app.logger.error("An error occurred", exc_info=True)
#         return jsonify({"status": "error", "message": str(e)}), 500 

# @app.route("/download_toptenalts")
# def download_top_ten_alts():
#     file_path = tta.top_ten_alts()
#     return send_file(
#         file_path,
#         as_attachment=True,
#         download_name="toptenalts.xlsx",
#         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     )

@api_bp.post("/wsop")
def workshop_operations():
    try:
        json_data = request.get_json()
        if not json_data:
            return jsonify({"error": "No JSON data received"}), 400

        ws.OperationManager().update_operation(json_data)

        return jsonify({"status": "success", "data_received": json_data}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@api_bp.get("/wsop/<drive_id>")
@api_bp.get("/wsop/<drive_id>/<operation>")
def get_jobcard_data(drive_id, operation=None):
    try:

        op = ws.OperationManager()

        if operation:
            op_data = op.get_single_operation_data(drive_id, operation)
        else:
            op_data = op.get_job_card_data(drive_id)

        if op_data is not None:
            return jsonify({"status": "success", "data": op_data}), 200

        else:
            return jsonify({"error": "Data not found"}), 404

    except Exception as e:
        return jsonify({"error": str(e)}), 500

@api_bp.post("/wsop/go")
def google_interactions():
    function_map = {
        "db_upd": (ws.update_db_job_cards_info, False),
        "mstr_upd": (ws.update_google_master_table, False),
        "color": (lambda data: ws.color_master_table_cell(**data), True),
        "prj_upd": (ws.update_projects_google_sheet, False),
        "cell_note": (lambda data: ws.place_note_master_table(**data), True),
        "jc_clr": (lambda data: ws.close_job_card_color(**data), True),
    }

    try:
        json_data = request.get_json()
        if not json_data:
            return jsonify({"error": "No JSON data received"}), 400

        action = json_data.get("action")
        if not action or action not in function_map:
            return jsonify({"error": "Invalid or missing action"}), 400

        func, needs_args = function_map[action]

        if needs_args:
            func(json_data.get("data", {}))
        else:
            func()

        return jsonify({"status": "success", "action": action}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500