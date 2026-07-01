"""
Tests for editable owner/action_plan web feature.
Run:  .venv\\Scripts\\python.exe dfv_tool\\test_editable_actionplan.py

DB isolation: each DB-touching test points history.DB_PATH at a fresh temp DB.
Pure functions are exercised by monkeypatching the looked-up helper.
"""
import os
import sys
import tempfile

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import history
import pipeline
import dashboard


def _fresh_db():
    """Point history at a brand-new temp DB; return the path."""
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    os.remove(path)  # let sqlite create it fresh
    history.DB_PATH = path
    history.init_db()
    return path


def test_schema_has_action_plan_and_audit():
    _fresh_db()
    conn = history._get_conn()
    cols = [r[1] for r in conn.execute("PRAGMA table_info(errors)").fetchall()]
    assert "action_plan" in cols, cols
    tables = [r[0] for r in conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
    assert "audit_log" in tables, tables
    conn.close()
    print("PASS test_schema_has_action_plan_and_audit")


def test_latest_action_plan_map():
    _fresh_db()
    # Two runs; same code/plant appears in both. Latest non-empty wins;
    # if latest run's value is empty, fall back to the most recent non-empty.
    conn = history._get_conn()
    conn.execute("INSERT INTO runs (id, run_date) VALUES (1, '2026-06-01 09:00')")
    conn.execute("INSERT INTO runs (id, run_date) VALUES (2, '2026-06-08 09:00')")
    # code=P1/plant=L1: run1 has plan 'old', run2 empty -> expect 'old'
    conn.execute("INSERT INTO errors (run_id, apo_product, apo_location, action_plan) "
                 "VALUES (1,'P1','L1','old plan')")
    conn.execute("INSERT INTO errors (run_id, apo_product, apo_location, action_plan) "
                 "VALUES (2,'P1','L1','')")
    # code=P2/plant=L2: run1 'a', run2 'b' -> latest 'b'
    conn.execute("INSERT INTO errors (run_id, apo_product, apo_location, action_plan) "
                 "VALUES (1,'P2','L2','a')")
    conn.execute("INSERT INTO errors (run_id, apo_product, apo_location, action_plan) "
                 "VALUES (2,'P2','L2','b')")
    conn.commit()
    conn.close()

    m = history.get_latest_action_plan_map()
    assert m[("P1", "L1")] == "old plan", m.get(("P1", "L1"))
    assert m[("P2", "L2")] == "b", m.get(("P2", "L2"))
    print("PASS test_latest_action_plan_map")


def test_save_run_persists_action_plan():
    _fresh_db()
    summary = {"total_rows": 1, "total_idp": 100.0, "total_apo": 100.0,
               "diff_pct": 0.0, "vol_status": "OK", "total_errors": 1,
               "actionable_errors": 1, "hktw_errors": 0, "sku_status": "OK",
               "impact_volume_3m": 0.0}
    errors_df = pd.DataFrame([{
        "File_ID": "F", "APO_Product": "P1", "Description": "d", "Category": "c",
        "Brand": "b", "APO_Location": "L1", "SNP_Planner": "s",
        "Error_Message": "Missing Mat/Loc", "IDP_Forecast": 100.0,
        "APO_Forecast": 100.0, "Reason": "r", "Action": "auto",
        "Owner": "GC DRP", "Is_HKTW": False, "First_Time": "W01/2026",
        "Duration": 0, "Priority": "Low", "Action_Plan": "call planner",
    }])
    rid = history.save_run(summary, errors_df)
    conn = history._get_conn()
    val = conn.execute("SELECT action_plan FROM errors WHERE run_id=?", (rid,)).fetchone()[0]
    conn.close()
    assert val == "call planner", val
    print("PASS test_save_run_persists_action_plan")


def test_update_error_fields_and_audit():
    _fresh_db()
    conn = history._get_conn()
    conn.execute("INSERT INTO runs (id, run_date) VALUES (1,'2026-06-08 09:00')")
    conn.execute("INSERT INTO errors (id, run_id, apo_product, apo_location, "
                 "owner, action_plan) VALUES (10,1,'P1','L1','GC DRP','')")
    conn.commit()
    conn.close()

    n = history.update_error_fields(10, owner="Alice", action_plan="do X", ip="10.0.0.5")
    assert n == 1, n
    conn = history._get_conn()
    row = conn.execute("SELECT owner, action_plan FROM errors WHERE id=10").fetchone()
    assert row["owner"] == "Alice" and row["action_plan"] == "do X", dict(row)
    audits = conn.execute("SELECT field, old_value, new_value, ip FROM audit_log "
                          "WHERE error_id=10 ORDER BY field").fetchall()
    conn.close()
    fields = {a["field"]: (a["old_value"], a["new_value"], a["ip"]) for a in audits}
    assert fields["owner"] == ("GC DRP", "Alice", "10.0.0.5"), fields
    assert fields["action_plan"] == ("", "do X", "10.0.0.5"), fields
    # Non-existent id -> 0 rows, no crash
    assert history.update_error_fields(9999, owner="Z") == 0
    print("PASS test_update_error_fields_and_audit")


def test_enrich_action_plan_materializes():
    original = history.get_latest_action_plan_map
    history.get_latest_action_plan_map = lambda: {("P1", "L1"): "reuse me"}
    try:
        df = pd.DataFrame([
            {"APO_Product": "P1", "APO_Location": "L1"},   # has history
            {"APO_Product": "NEW", "APO_Location": "Z9"},   # no history -> ""
            {"APO_Product": 999, "APO_Location": "X1"},     # int key, no history -> ""
        ])
        out = pipeline.enrich_action_plan(df.copy())
        assert out.iloc[0]["Action_Plan"] == "reuse me", out.iloc[0]["Action_Plan"]
        assert out.iloc[1]["Action_Plan"] == "", out.iloc[1]["Action_Plan"]
        assert out.iloc[2]["Action_Plan"] == "", out.iloc[2]["Action_Plan"]
        assert pipeline.enrich_action_plan(pd.DataFrame()).empty
    finally:
        history.get_latest_action_plan_map = original
    print("PASS test_enrich_action_plan_materializes")


def test_api_get_and_post_validation():
    _fresh_db()
    conn = history._get_conn()
    conn.execute("INSERT INTO runs (id, run_date) VALUES (1,'2026-06-08 09:00')")
    conn.execute("INSERT INTO errors (id, run_id, apo_product, apo_location, "
                 "error_message, owner, action, action_plan, is_hktw) "
                 "VALUES (20,1,'P1','L1','Missing Mat/Loc','GC DRP','auto','',0)")
    conn.commit()
    conn.close()

    import app as appmod
    client = appmod.create_app().test_client()

    # GET returns the row
    r = client.get("/api/errors")
    assert r.status_code == 200, r.status_code
    rows = r.get_json()
    assert any(x["id"] == 20 for x in rows), rows

    # POST valid update
    r = client.post("/api/errors/20", json={"owner": "Alice", "action_plan": "do X"})
    assert r.status_code == 200, (r.status_code, r.get_data(as_text=True))

    # Non-existent id -> 404
    assert client.post("/api/errors/9999", json={"owner": "Z"}).status_code == 404
    # Over-length owner -> 400
    assert client.post("/api/errors/20", json={"owner": "x" * 201}).status_code == 400
    # Over-length action_plan -> 400
    assert client.post("/api/errors/20", json={"action_plan": "y" * 2001}).status_code == 400
    # No editable fields -> 400
    assert client.post("/api/errors/20", json={"foo": "bar"}).status_code == 400
    print("PASS test_api_get_and_post_validation")


def test_dashboard_editable_wiring():
    t = dashboard._get_template()
    assert "function escapeHtml" in t, "escapeHtml missing"
    assert "<th>Action Plan</th>" in t, "Action Plan header missing"
    # Editable cells use contenteditable + blur save hook
    assert "contenteditable" in t, "editable cell missing"
    assert "function saveCell" in t, "saveCell missing"
    assert "/api/errors/" in t, "save endpoint not wired"
    # User-controlled fields escaped
    assert "escapeHtml(e.action_plan" in t, "action_plan not escaped"
    assert "escapeHtml(e.owner" in t, "owner not escaped"
    print("PASS test_dashboard_editable_wiring")


def test_dashboard_owner_lock_and_column_order():
    t = dashboard._get_template()
    # Owner + Action Plan columns moved to AFTER Priority; Action Plan is last.
    order = t.find("<th>Priority</th>")
    ap = t.find("<th>Action Plan</th>")
    ow = t.find("<th>Owner</th>")
    assert order != -1 and ap != -1 and ow != -1, "headers missing"
    assert ow > order and ap > order, "Owner/Action Plan not after Priority"
    assert ap > ow, "Action Plan must be the last column (after Owner)"
    # Owner + Action Plan editing is gated behind a single toggle (locked by default).
    assert "editEnabled" in t, "edit lock state missing"
    assert "function toggleOwnerEdit" in t, "edit toggle missing"
    assert 'id="ownerEditToggle"' in t, "toggle button missing"
    assert "var editEnabled = false" in t, "editing must default locked"
    # Both editable cells are gated by the toggle.
    assert t.count("editEnabled ?") >= 2, "action_plan/owner cells not both gated by toggle"
    # New owner names propagate into the filter after save.
    assert "function applyOwnerEdit" in t, "owner->filter refresh missing"
    print("PASS test_dashboard_owner_lock_and_column_order")


def test_dashboard_inline_js_syntax_ok():
    # A broken edit to the inline <script> silently blanks the whole dashboard
    # (all charts + table). Guard the generated JS with a real parser when node
    # is available; skip cleanly if it is not.
    import re, shutil, subprocess, tempfile, os
    node = shutil.which("node")
    if not node:
        print("SKIP test_dashboard_inline_js_syntax_ok (node not found)")
        return
    html = dashboard._get_template()
    blocks = re.findall(r"<script(?![^>]*\bsrc=)[^>]*>(.*?)</script>", html, re.S)
    js = "\n;\n".join(blocks).replace("__DATA_JSON__", "[]")
    fd, path = tempfile.mkstemp(suffix=".js")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(js)
        r = subprocess.run([node, "--check", path], capture_output=True, text=True)
        assert r.returncode == 0, "inline JS syntax error:\n" + r.stderr
    finally:
        os.remove(path)
    print("PASS test_dashboard_inline_js_syntax_ok")


def test_dashboard_json_embedding_escaped():
    # User-editable action_plan flows into the embedded <script>DATA. json.dumps
    # does NOT escape </script>, so a raw payload could break out of the script
    # tag (stored XSS). _build_html must \u-escape < > & before embedding.
    import json
    payload = [{"week_label": "W01/2026", "run_date": "2026-01-01 09:00",
                "errors": [{"id": 1, "action_plan": "</script><script>x"}]}]
    html = dashboard._build_html(json.dumps(payload, ensure_ascii=False), payload)
    assert "</script><script>x" not in html, "raw script breakout embedded"
    assert "\\u003c/script\\u003e" in html, "payload not unicode-escaped"
    print("PASS test_dashboard_json_embedding_escaped")


def test_dashboard_owner_filter_escaped():
    # Owner filter buttons are built client-side from user-editable owner values.
    # The label and the click argument must not concatenate raw owner text into
    # innerHTML/onclick (stored XSS). Enforce escapeHtml + data-owner delegation.
    t = dashboard._get_template()
    assert "escapeHtml(owners[j])" in t, "owner filter label/attr not escaped"
    assert "data-owner=" in t, "owner filter not using data-owner delegation"
    assert "setFilter(&quot;' + esc" not in t, "raw owner still inlined into onclick"
    assert "'\">' + owners[j] +" not in t, "raw owner still used as button label"
    print("PASS test_dashboard_owner_filter_escaped")


if __name__ == "__main__":
    test_schema_has_action_plan_and_audit()
    test_latest_action_plan_map()
    test_save_run_persists_action_plan()
    test_update_error_fields_and_audit()
    test_enrich_action_plan_materializes()
    test_api_get_and_post_validation()
    test_dashboard_editable_wiring()
    test_dashboard_owner_lock_and_column_order()
    test_dashboard_inline_js_syntax_ok()
    test_dashboard_json_embedding_escaped()
    test_dashboard_owner_filter_escaped()
    print("\nALL EDITABLE TESTS PASSED")
