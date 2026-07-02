"""
Tests for the weekly-report-email feature.
Run:  .venv\\Scripts\\python.exe dfv_tool\\test_weekly_email.py

Pure builder tests need no DB. Endpoint tests monkeypatch history.get_all_data
and email_report.open_outlook_draft so nothing touches Outlook or the real DB.
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import email_report
import history
import app as app_module
import dashboard

# Valid 1x1 PNG (base64, no data-URL prefix) for screenshot tests.
_PNG_1x1 = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk"
            "+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==")


def _run(**kw):
    base = {
        "id": 1,
        "run_date": "2026-06-30 11:53",
        "week_label": "W27/2026",
        "diff_pct": 0.57,
        "errors": [],
    }
    base.update(kw)
    return base


def _err(**kw):
    base = {
        "apo_product": "83929605", "description": "H&S SHM", "brand": "HD&SHLDRS",
        "apo_location": "A673", "error_message": "Missing Mat/Loc", "idp_forecast": 11438.0,
        "reason": "Missing master data", "action": "Apply T-lane", "action_plan": "",
        "owner": "Becky", "first_time": "W24/2026", "duration": 3, "priority": "Mid",
    }
    base.update(kw)
    return base


def test_build_email_subject_and_intro():
    run = _run(errors=[_err(duration=5), _err(duration=2), _err(duration=6)])
    subject, html = email_report.build_weekly_email(run)
    assert subject == "GC HC DFV Weekly Result-20260630", subject
    assert "0.57%" in html, "result % missing"
    assert ">3<" in html or "<b>3</b>" in html, "item count missing"
    # 2 items have duration >= 4 (5 and 6)
    assert "<b>2</b>" in html, "aging>=4 count missing"
    print("PASS test_build_email_subject_and_intro")


def test_build_email_table_columns_order():
    run = _run(errors=[_err()])
    _, html = email_report.build_weekly_email(run)
    for h in ["Product", "Description", "Brand", "Location", "Error", "Forecast",
              "Reason", "Action", "First Time", "Duration", "Priority", "Owner", "Action Plan"]:
        assert ">" + h + "<" in html, "missing column " + h
    # Owner must come before Action Plan; Action Plan is the last data column.
    assert html.index(">Owner<") < html.index(">Action Plan<"), "Owner must precede Action Plan"
    assert "83929605" in html, "product value missing"
    print("PASS test_build_email_table_columns_order")


def test_build_email_escapes_html():
    run = _run(errors=[_err(owner='<img src=x onerror=alert(1)>')])
    _, html = email_report.build_weekly_email(run)
    assert "<img src=x onerror=alert(1)>" not in html, "raw payload embedded"
    assert "&lt;img" in html, "payload not escaped"
    print("PASS test_build_email_escapes_html")


def test_build_email_prev_week_comparison():
    run = _run(diff_pct=0.57, errors=[_err(), _err(), _err()])          # 3 items
    prev = _run(id=0, run_date="2026-06-23 10:00", diff_pct=0.69,
                errors=[_err(), _err(), _err(), _err(), _err(), _err(), _err(), _err()])  # 8 items
    _, html = email_report.build_weekly_email(run, prev)
    assert "vs last week" in html, "prev-week line missing"
    assert "0.69% &rarr; 0.57%" in html, "result delta text missing"
    assert "-5" in html, "item delta missing (8 -> 3)"
    print("PASS test_build_email_prev_week_comparison")


def test_build_email_prev_week_omitted_when_none():
    _, html = email_report.build_weekly_email(_run(errors=[_err()]))
    assert "vs last week" not in html, "prev-week line must be omitted"
    print("PASS test_build_email_prev_week_omitted_when_none")


def test_build_email_owner_summary_bold():
    run = _run(errors=[_err(owner="Becky"), _err(owner="Becky"), _err(owner="Lucy"),
                       _err(owner="")])
    _, html = email_report.build_weekly_email(run)
    assert "<b>Becky</b>: 2" in html, "Becky count/bold wrong"
    assert "<b>Lucy</b>: 1" in html, "Lucy count/bold wrong"
    assert "(未分配)</b>: 1" in html, "blank owner not grouped"
    print("PASS test_build_email_owner_summary_bold")


def test_build_email_inserts_screenshot_when_cid():
    _, html = email_report.build_weekly_email(_run(errors=[_err()]), image_cid="dashboard")
    assert 'src="cid:dashboard"' in html, "screenshot img missing"
    assert html.index("Executive Summary") < html.index("cid:dashboard"), "img must be after summary"
    assert html.index("cid:dashboard") < html.index("Details as below"), "img must be before details"
    print("PASS test_build_email_inserts_screenshot_when_cid")


def test_build_email_no_screenshot_by_default():
    _, html = email_report.build_weekly_email(_run(errors=[_err()]))
    assert "cid:" not in html, "no cid image expected by default"
    print("PASS test_build_email_no_screenshot_by_default")


def test_load_recipients_reads_file():
    fd, path = tempfile.mkstemp(suffix=".json")
    os.close(fd)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"to": ["A <a@pg.com>", "B <b@pg.com>"], "cc": ["C <c@pg.com>"]}, f)
    try:
        r = email_report.load_recipients(path)
        assert r["to"] == ["A <a@pg.com>", "B <b@pg.com>"], r
        assert r["cc"] == ["C <c@pg.com>"], r
    finally:
        os.remove(path)
    print("PASS test_load_recipients_reads_file")


def test_load_recipients_missing_returns_blank():
    r = email_report.load_recipients(os.path.join(tempfile.gettempdir(), "no_such_email_cfg.json"))
    assert r == {"to": [], "cc": []}, r
    print("PASS test_load_recipients_missing_returns_blank")


def test_open_draft_never_sends():
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_report.py"),
               encoding="utf-8").read()
    assert ".Send(" not in src, "email_report must never call .Send()"
    assert ".Display(" in src, "open_outlook_draft must use .Display()"
    assert hasattr(email_report, "open_outlook_draft"), "open_outlook_draft missing"
    print("PASS test_open_draft_never_sends")


def _client(monkeypatch_runs, draft_stub):
    """Build a Flask test client with history.get_all_data and the Outlook COM
    call replaced. Returns a client; draft_stub records/handles draft invocations."""
    history.get_all_data = lambda: monkeypatch_runs           # simple stub
    email_report.open_outlook_draft = draft_stub
    return app_module.create_app().test_client()


def test_endpoint_generates_draft_200():
    calls = []
    runs = [_run(id=2, run_date="2026-06-30 11:53", errors=[_err(duration=5)]),
            _run(id=1, run_date="2026-06-23 10:00", errors=[_err(), _err()])]
    client = _client(runs, lambda *a, **k: calls.append(a))
    r = client.post("/api/email/weekly", json={"run_id": 2})
    assert r.status_code == 200, r.status_code
    assert r.get_json()["ok"] is True, r.get_json()
    assert r.get_json()["subject"] == "GC HC DFV Weekly Result-20260630", r.get_json()
    assert len(calls) == 1, "draft not opened exactly once"
    print("PASS test_endpoint_generates_draft_200")


def test_endpoint_bad_run_id_404():
    client = _client([_run(id=2)], lambda *a, **k: None)
    r = client.post("/api/email/weekly", json={"run_id": 999})
    assert r.status_code == 404, r.status_code
    print("PASS test_endpoint_bad_run_id_404")


def test_endpoint_missing_run_id_400():
    client = _client([_run(id=2)], lambda *a, **k: None)
    r = client.post("/api/email/weekly", json={})
    assert r.status_code == 400, r.status_code
    print("PASS test_endpoint_missing_run_id_400")


def test_endpoint_bool_run_id_400():
    # bool is an int subclass; it must not be accepted as a run_id.
    client = _client([_run(id=1, errors=[_err()])], lambda *a, **k: None)
    r = client.post("/api/email/weekly", json={"run_id": True})
    assert r.status_code == 400, r.status_code
    print("PASS test_endpoint_bool_run_id_400")


def test_endpoint_com_failure_500():
    def boom(*a, **k):
        raise RuntimeError("Outlook not installed")
    client = _client([_run(id=2, errors=[_err()])], boom)
    r = client.post("/api/email/weekly", json={"run_id": 2})
    assert r.status_code == 500, r.status_code
    assert r.get_json()["ok"] is False, r.get_json()
    print("PASS test_endpoint_com_failure_500")


def test_endpoint_with_screenshot_200():
    captured = {}
    def stub(subject, to, cc, html, inline_images=None):
        captured["inline"] = inline_images
        captured["html"] = html
    history.get_all_data = lambda: [_run(id=2, errors=[_err()])]
    email_report.open_outlook_draft = stub
    client = app_module.create_app().test_client()
    r = client.post("/api/email/weekly",
                    json={"run_id": 2, "screenshot": "data:image/png;base64," + _PNG_1x1})
    assert r.status_code == 200, r.status_code
    assert captured["inline"] and captured["inline"][0][0] == "dashboard", captured.get("inline")
    import base64 as _b64
    assert captured["inline"][0][1] == _b64.b64decode(_PNG_1x1), "png bytes mismatch"
    assert "cid:dashboard" in captured["html"], "html missing cid ref"
    print("PASS test_endpoint_with_screenshot_200")


def test_endpoint_bad_screenshot_type_400():
    client = _client([_run(id=2, errors=[_err()])], lambda *a, **k: None)
    r = client.post("/api/email/weekly",
                    json={"run_id": 2, "screenshot": "data:text/html;base64,AAAA"})
    assert r.status_code == 400, r.status_code
    print("PASS test_endpoint_bad_screenshot_type_400")


def test_endpoint_oversize_screenshot_400():
    old = app_module.SCREENSHOT_MAX_B64
    app_module.SCREENSHOT_MAX_B64 = 4
    try:
        client = _client([_run(id=2, errors=[_err()])], lambda *a, **k: None)
        r = client.post("/api/email/weekly",
                        json={"run_id": 2, "screenshot": "data:image/png;base64," + _PNG_1x1})
        assert r.status_code == 400, r.status_code
    finally:
        app_module.SCREENSHOT_MAX_B64 = old
    print("PASS test_endpoint_oversize_screenshot_400")


def test_open_draft_supports_inline_images():
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_report.py"),
               encoding="utf-8").read()
    assert "inline_images" in src, "open_outlook_draft must accept inline_images"
    assert "3712001F" in src, "must set PR_ATTACH_CONTENT_ID for inline image"
    assert "Attachments" in src, "must attach the image"
    print("PASS test_open_draft_supports_inline_images")


def test_dashboard_email_button_wired():
    t = dashboard._get_template()
    assert "生成周报邮件" in t, "email button label missing"
    assert "function generateEmail" in t, "generateEmail() missing"
    assert "/api/email/weekly" in t, "email endpoint not wired"
    assert "run_id" in t, "run_id not sent"
    print("PASS test_dashboard_email_button_wired")


def test_dashboard_inline_js_syntax_ok_email():
    import re, shutil, subprocess
    node = shutil.which("node")
    if not node:
        print("SKIP test_dashboard_inline_js_syntax_ok_email (node not found)")
        return
    html = dashboard._get_template()
    blocks = re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S)
    js = "\n;\n".join(blocks).replace("__DATA_JSON__", "[]")
    fd, path = tempfile.mkstemp(suffix=".js")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(js)
        rc = subprocess.run([node, "--check", path], capture_output=True, text=True)
        assert rc.returncode == 0, "inline JS syntax error:\n" + rc.stderr
    finally:
        os.remove(path)
    print("PASS test_dashboard_inline_js_syntax_ok_email")


if __name__ == "__main__":
    test_build_email_subject_and_intro()
    test_build_email_table_columns_order()
    test_build_email_escapes_html()
    test_build_email_prev_week_comparison()
    test_build_email_prev_week_omitted_when_none()
    test_build_email_owner_summary_bold()
    test_build_email_inserts_screenshot_when_cid()
    test_build_email_no_screenshot_by_default()
    test_load_recipients_reads_file()
    test_load_recipients_missing_returns_blank()
    test_open_draft_never_sends()
    test_open_draft_supports_inline_images()
    test_endpoint_generates_draft_200()
    test_endpoint_bad_run_id_404()
    test_endpoint_missing_run_id_400()
    test_endpoint_bool_run_id_400()
    test_endpoint_com_failure_500()
    test_endpoint_with_screenshot_200()
    test_endpoint_bad_screenshot_type_400()
    test_endpoint_oversize_screenshot_400()
    test_dashboard_email_button_wired()
    test_dashboard_inline_js_syntax_ok_email()
    print("\nALL WEEKLY EMAIL TESTS PASSED")
