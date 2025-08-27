import os
from dotenv import load_dotenv
from flask import Flask, jsonify, request
from sqlalchemy import text
from .db import engine, SessionLocal
from .api import api_bp

# Find project root (backend/)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ENV_PATH = os.path.join(BASE_DIR, ".env.dev")

# Load env vars
load_dotenv(ENV_PATH, override=True)

DATABASE_URL = os.getenv("DATABASE_URL")
API_KEY_DEV = os.getenv("API_KEY_DEV", "")


def create_app():
    app = Flask(__name__)

    # ---- auth (kept here so it applies to all blueprints) ----
    @app.before_request
    def _auth():
        # Simple dev-only header check. We'll match this from Sheets later.
        from flask import request, abort

        if request.path.startswith("/healthz"):
            return
        if request.headers.get("X-Api-Key") != API_KEY_DEV:
            abort(401)

    # ---- health ----
    @app.get("/healthz")
    def healthz():
        try:
            with engine.connect() as conn:
                conn.execute(text("SELECT 1"))
            return jsonify(ok=True, db="up")
        except Exception as e:
            return jsonify(ok=False, db="down", error=str(e)), 500

    # ---- errors ----
    def json_error(status: int, code: str, message: str, details: dict | None = None):
        payload = {"ok": False, "error": {"code": code, "message": message}}
        if details:
            payload["error"]["details"] = details
        return jsonify(payload), status

    @app.errorhandler(400)
    def _bad_request(e):
        return json_error(400, "BAD_REQUEST", "Bad request")

    @app.errorhandler(401)
    def _unauth(e):
        return json_error(401, "UNAUTHORIZED", "Invalid or missing API key")

    @app.errorhandler(404)
    def _not_found(e):
        return json_error(404, "NOT_FOUND", "Endpoint not found")

    @app.errorhandler(500)
    def _server(e):
        return json_error(500, "INTERNAL", "Internal error")

    # Root helper so hitting / is never confusing
    @app.get("/")
    def root():
        return jsonify(ok=True, message="Use /healthz or /v1/*"), 200

    app.register_blueprint(api_bp, url_prefix="/v1")

    return app

app = create_app()

if __name__ == "__main__":
    # Handy for: python backend/app/wsgi.py
    app.run(host="0.0.0.0", port=5000, debug=True)
