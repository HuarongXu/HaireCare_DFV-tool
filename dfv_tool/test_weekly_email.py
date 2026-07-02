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


if __name__ == "__main__":
    test_build_email_subject_and_intro()
    test_build_email_table_columns_order()
    test_build_email_escapes_html()
    print("\nALL WEEKLY EMAIL TESTS PASSED")
