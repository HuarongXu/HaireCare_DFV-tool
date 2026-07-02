"""
DFV Web App - internal-only Flask server for editing Owner / Action plan.

Trust model: trusted internal LAN, no auth (see spec
specs/editable-owner-actionplan/spec.md). All request input is still validated
at this boundary and all DB writes are parameterized. Security response headers
are set to reduce clickjacking / MIME sniffing.
"""
import json
import os
import sys

from flask import Flask, request, jsonify, Response

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import history
from dashboard import _build_html
import email_report

OWNER_MAX = 200
ACTION_PLAN_MAX = 2000
# Drop control chars except tab/newline/carriage-return (prevents log/HTML
# injection via stored values).
_CTRL = {c: None for c in range(32) if c not in (9, 10, 13)}


def _clean_text(v):
    """Coerce to str and strip disallowed control characters."""
    return str(v).translate(_CTRL)


def create_app():
    app = Flask(__name__)

    @app.after_request
    def _security_headers(resp):
        resp.headers["X-Frame-Options"] = "DENY"
        resp.headers["X-Content-Type-Options"] = "nosniff"
        resp.headers["Referrer-Policy"] = "no-referrer"
        return resp

    @app.get("/")
    def index():
        runs = history.get_all_data()
        data_json = json.dumps(runs, ensure_ascii=False, default=str)
        return Response(_build_html(data_json, runs), mimetype="text/html")

    @app.get("/api/errors")
    def get_errors():
        run_id = request.args.get("run_id", type=int)
        if run_id is None:
            runs = history.get_all_runs()
            if not runs:
                return jsonify([])
            run_id = runs[0]["id"]
        return jsonify(history.get_errors_for_run(run_id))

    @app.post("/api/errors/<int:error_id>")
    def update_error(error_id):
        data = request.get_json(silent=True) or {}
        fields = {}
        if "owner" in data:
            owner = _clean_text(data["owner"])
            if len(owner) > OWNER_MAX:
                return jsonify({"error": "owner too long"}), 400
            fields["owner"] = owner
        if "action_plan" in data:
            ap = _clean_text(data["action_plan"])
            if len(ap) > ACTION_PLAN_MAX:
                return jsonify({"error": "action_plan too long"}), 400
            fields["action_plan"] = ap
        if not fields:
            return jsonify({"error": "no editable field provided"}), 400

        # No reverse proxy -> trust remote_addr, not X-Forwarded-For.
        ip = request.remote_addr
        n = history.update_error_fields(error_id, ip=ip, **fields)
        if n == 0:
            return jsonify({"error": "not found"}), 404
        return jsonify({"id": error_id, **fields})

    def _prev_run(runs, run):
        """runs is DESC by run_date (as get_all_data returns). Previous week is
        the next-older run in that list, or None if this is the earliest."""
        for i, r in enumerate(runs):
            if r["id"] == run["id"]:
                return runs[i + 1] if i + 1 < len(runs) else None
        return None

    @app.post("/api/email/weekly")
    def email_weekly():
        data = request.get_json(silent=True) or {}
        run_id = data.get("run_id")
        if not isinstance(run_id, int) or isinstance(run_id, bool):
            return jsonify({"ok": False, "error": "run_id (int) required"}), 400
        runs = history.get_all_data()
        run = next((r for r in runs if r["id"] == run_id), None)
        if run is None:
            return jsonify({"ok": False, "error": "run not found"}), 404
        subject, html = email_report.build_weekly_email(run, _prev_run(runs, run))
        recips = email_report.load_recipients()
        try:
            email_report.open_outlook_draft(subject, recips["to"], recips["cc"], html)
        except Exception as e:  # COM/system failure -> 500 (R13)
            app.logger.exception("weekly email draft failed run_id=%s", run_id)
            return jsonify({"ok": False, "error": str(e)}), 500
        # Audit: no recipient PII in the log line (R14).
        app.logger.info("weekly email draft opened run_id=%s ip=%s", run_id, request.remote_addr)
        return jsonify({"ok": True, "subject": subject})

    return app


if __name__ == "__main__":
    from waitress import serve
    serve(create_app(), host="0.0.0.0", port=8000)
