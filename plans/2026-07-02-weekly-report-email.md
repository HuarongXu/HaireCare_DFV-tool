# 一键生成周报邮件（Outlook 草稿）实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在现有 DFV Flask Web 应用上加一个按钮，一键根据当前选中周的数据生成一封 Outlook 草稿邮件（收件人/主题/开头总结/Executive Summary/明细表），由用户在 Outlook 里核对后手动发送。

**Architecture:** 纯函数 `build_weekly_email(run, prev_run)` 负责拼 HTML（可独立单测，不碰 COM）；`open_outlook_draft()` 单独封装 Outlook COM 调用，只 `.Display()` 绝不 `.Send()`；`load_recipients()` 从 gitignored 配置读 To/CC。新增 `POST /api/email/weekly` 端点串起来。前端加按钮调用该端点。

**Tech Stack:** Python 3.13 / Flask / pywin32(win32com) / SQLite（现有）。测试沿用仓库既有的"手写 assert 函数 + main runner"风格，非 pytest。

**关联规范：** specs/weekly-report-email/spec.md（R1–R14）

---

## 文件结构

- **Create** `dfv_tool/email_report.py` — 邮件构建纯函数 + 收件人加载 + Outlook COM 封装
- **Create** `dfv_tool/email_config.example.json` — 收件人配置模板（占位邮箱，无真实 PII）
- **Create** `dfv_tool/test_weekly_email.py` — 本功能全部测试
- **Modify** `dfv_tool/app.py` — 新增 `POST /api/email/weekly` 端点
- **Modify** `dfv_tool/dashboard.py` — 顶部加"生成周报邮件"按钮 + `generateEmail()` JS
- **Modify** `.gitignore` — 忽略 `email_config.json`、`*.msg`、`.tmp_msgtools/`

**测试运行命令（全程统一）：**
```
.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py
```

---

## Task 1: 邮件构建纯函数 build_weekly_email

**Files:**
- Create: `dfv_tool/email_report.py`
- Test: `dfv_tool/test_weekly_email.py`

覆盖 R3（主题）、R5（开头总结）、R7（Aging）、R8（Owner 汇总）、R9（表格列）、R12（HTML 转义）。

- [ ] **Step 1: 写失败测试**

创建 `dfv_tool/test_weekly_email.py`：

```python
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
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: FAIL — `ModuleNotFoundError: No module named 'email_report'`

- [ ] **Step 3: 写最小实现**

创建 `dfv_tool/email_report.py`：

```python
"""
Weekly DFV report email.

Split into a pure HTML builder (`build_weekly_email`) that is fully unit-testable
without Outlook, and a thin COM wrapper (`open_outlook_draft`) that opens the mail
as a DRAFT only. This module NEVER calls .Send() — the user reviews and sends in
Outlook themselves (see specs/weekly-report-email/spec.md, R10).
"""
import html as _html
import json
import logging
import os

log = logging.getLogger(__name__)

_CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_config.json")

# (data key, column header) in the exact order used by the email table (R9).
_COLS = [
    ("apo_product", "Product"), ("description", "Description"), ("brand", "Brand"),
    ("apo_location", "Location"), ("error_message", "Error"), ("idp_forecast", "Forecast"),
    ("reason", "Reason"), ("action", "Action"), ("first_time", "First Time"),
    ("duration", "Duration"), ("priority", "Priority"), ("owner", "Owner"),
    ("action_plan", "Action Plan"),
]


def _esc(v):
    """HTML-escape any value (None -> '')."""
    return "" if v is None else _html.escape(str(v))


def _yyyymmdd(run_date):
    return str(run_date)[:10].replace("-", "")


def _dur(e):
    d = e.get("duration")
    return d if isinstance(d, (int, float)) else 0


def _owner_counts(errors):
    counts = {}
    for e in errors:
        o = (str(e.get("owner") or "")).strip() or "(未分配)"
        counts[o] = counts.get(o, 0) + 1
    return sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))


def build_weekly_email(run, prev_run=None):
    """Return (subject, html_body) for one weekly run. Pure; no I/O, no COM."""
    errors = run.get("errors") or []
    n = len(errors)
    diff = run.get("diff_pct")
    diff_s = ("%.2f" % diff) if isinstance(diff, (int, float)) else "?"
    a4 = sum(1 for e in errors if _dur(e) >= 4)
    a8 = sum(1 for e in errors if _dur(e) >= 8)
    subject = "GC HC DFV Weekly Result-" + _yyyymmdd(run.get("run_date"))

    intro = (
        "<p>Dear all,</p>"
        "<p>This week's result is <b>%s%%</b>, and we still have <b>%d</b> items to close, "
        "of which <b>%d</b> have been outstanding for over 4 weeks. "
        "Could all owners please resolve them as soon as possible. Thank you.</p>"
        % (diff_s, n, a4)
    )

    summ = ["<p><b>Executive Summary</b></p><ul>"]
    if prev_run is not None:
        pdiff = prev_run.get("diff_pct")
        pn = len(prev_run.get("errors") or [])
        if isinstance(diff, (int, float)) and isinstance(pdiff, (int, float)):
            d = diff - pdiff
            arrow = "&#9650;" if d > 0 else ("&#9660;" if d < 0 else "=")
            summ.append("<li>vs last week: result %s %.2fpp (%.2f%% &rarr; %.2f%%), "
                        "items %+d (%d &rarr; %d)</li>" % (arrow, abs(d), pdiff, diff, n - pn, pn, n))
        else:
            summ.append("<li>vs last week: items %+d (%d &rarr; %d)</li>" % (n - pn, pn, n))
    summ.append("<li>Aging: %d items &ge; 4 weeks, %d items &ge; 8 weeks</li>" % (a4, a8))
    owner_bits = ", ".join("<b>%s</b>: %d" % (_esc(o), c) for o, c in _owner_counts(errors))
    summ.append("<li>By owner: %s</li>" % (owner_bits or "&mdash;"))
    summ.append("</ul>")

    thead = "".join(
        '<th style="background:#1a1a2e;color:#fff;padding:6px 10px;text-align:left;'
        'border:1px solid #ccc;">%s</th>' % _esc(h) for _, h in _COLS)
    trows = []
    for e in errors:
        tds = []
        for key, _h in _COLS:
            val = e.get(key)
            align = "left"
            if key == "idp_forecast":
                try:
                    val = "{:,}".format(int(round(float(val))))
                except (TypeError, ValueError):
                    val = ""
                align = "right"
            tds.append('<td style="padding:5px 10px;border:1px solid #e0e0e0;text-align:%s;">%s</td>'
                       % (align, _esc(val)))
        trows.append("<tr>%s</tr>" % "".join(tds))
    table = (
        "<p>Details as below:</p>"
        '<table style="border-collapse:collapse;font-family:Segoe UI,sans-serif;font-size:12px;">'
        "<thead><tr>%s</tr></thead><tbody>%s</tbody></table>" % (thead, "".join(trows))
    )

    body = ('<div style="font-family:Segoe UI,sans-serif;font-size:13px;color:#222;">'
            + intro + "".join(summ) + table
            + "<p>Any questions please let me know, thanks.</p></div>")
    return subject, body
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS（3 个 build 测试）

- [ ] **Step 5: 安全核对**（frontend-xss / injection-data-access）：所有 error 字段经 `_esc`（`html.escape`）转义后才入 HTML（R12）；无 SQL、无外部命令。✓

- [ ] **Step 6: 提交**

```bash
git add dfv_tool/email_report.py dfv_tool/test_weekly_email.py
git commit -m "feat(email): build_weekly_email pure HTML builder with escaping"
```

---

## Task 2: 环比上周与 Owner 汇总（补测已实现逻辑）

**Files:**
- Modify: `dfv_tool/test_weekly_email.py`（追加测试；`build_weekly_email` 已在 Task 1 覆盖 R6/R8 逻辑）

覆盖 R6（环比）、R8（Owner 汇总加粗）。

- [ ] **Step 1: 追加失败测试**

在 `test_weekly_email.py` 的 `if __name__` 之前插入：

```python
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
```

并在 `if __name__` 块追加调用：

```python
    test_build_email_prev_week_comparison()
    test_build_email_prev_week_omitted_when_none()
    test_build_email_owner_summary_bold()
```

- [ ] **Step 2: 跑测试**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS（现有实现应已满足；若失败按断言修 `build_weekly_email`）

- [ ] **Step 3: 提交**

```bash
git add dfv_tool/test_weekly_email.py
git commit -m "test(email): cover prev-week comparison and owner summary"
```

---

## Task 3: 收件人配置加载 load_recipients

**Files:**
- Modify: `dfv_tool/email_report.py`（新增 `load_recipients`）
- Create: `dfv_tool/email_config.example.json`
- Modify: `dfv_tool/test_weekly_email.py`

覆盖 R4（收件人；缺失时留空不报错）。

- [ ] **Step 1: 写失败测试**

在 `test_weekly_email.py` 追加：

```python
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
```

并在 `if __name__` 追加：

```python
    test_load_recipients_reads_file()
    test_load_recipients_missing_returns_blank()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: FAIL — `AttributeError: module 'email_report' has no attribute 'load_recipients'`

- [ ] **Step 3: 写最小实现**

在 `email_report.py` 末尾（`build_weekly_email` 之后）新增：

```python
def load_recipients(path=_CONFIG_PATH):
    """Read {"to": [...], "cc": [...]} from a local JSON config.

    Recipients are internal PII and must NOT be committed to git — the real file
    (email_config.json) is gitignored (R4). Missing/unreadable config is not fatal:
    return empty lists so the draft still opens (user can fill recipients in Outlook).
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        return {
            "to": [str(x) for x in (cfg.get("to") or [])],
            "cc": [str(x) for x in (cfg.get("cc") or [])],
        }
    except FileNotFoundError:
        log.warning("email_config.json not found at %s; recipients left blank", path)
        return {"to": [], "cc": []}
    except (ValueError, OSError) as e:
        log.warning("email_config.json unreadable (%s); recipients left blank", e)
        return {"to": [], "cc": []}
```

- [ ] **Step 4: 创建配置模板**

创建 `dfv_tool/email_config.example.json`（占位邮箱，无真实 PII）：

```json
{
  "to": [
    "Recipient One <person1@example.com>",
    "Recipient Two <person2@example.com>",
    "Some Distribution List <dl-name@example.com>"
  ],
  "cc": [
    "CC Person <cc1@example.com>"
  ]
}
```

- [ ] **Step 5: 跑测试确认通过**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS（含两个 recipients 测试）

- [ ] **Step 6: 安全核对**（secrets-sensitive-data）：真实收件人只存 gitignored `email_config.json`（Task 6 加忽略规则）；模板文件仅占位邮箱；加载失败不抛敏感信息。✓

- [ ] **Step 7: 提交**

```bash
git add dfv_tool/email_report.py dfv_tool/email_config.example.json dfv_tool/test_weekly_email.py
git commit -m "feat(email): load_recipients from gitignored config + example template"
```

---

## Task 4: Outlook COM 封装 open_outlook_draft（只 Display 不 Send）

**Files:**
- Modify: `dfv_tool/email_report.py`（新增 `open_outlook_draft`）
- Modify: `dfv_tool/test_weekly_email.py`（源码静态断言，不真调 COM）

覆盖 R10（草稿而非发送）、R11（COM 与构建分离）。

- [ ] **Step 1: 写失败测试**

在 `test_weekly_email.py` 追加（用源码扫描保证绝不 `.Send()`，不真的弹 Outlook）：

```python
def test_open_draft_never_sends():
    src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_report.py"),
               encoding="utf-8").read()
    assert ".Send(" not in src, "email_report must never call .Send()"
    assert ".Display(" in src, "open_outlook_draft must use .Display()"
    assert hasattr(email_report, "open_outlook_draft"), "open_outlook_draft missing"
    print("PASS test_open_draft_never_sends")
```

并在 `if __name__` 追加：

```python
    test_open_draft_never_sends()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: FAIL — `AssertionError: open_outlook_draft missing`（`.Display(` 也不在源码里）

- [ ] **Step 3: 写最小实现**

在 `email_report.py` 末尾新增：

```python
def open_outlook_draft(subject, to, cc, html):
    """Open an Outlook draft (never send). win32com is imported lazily so unit
    tests and non-Windows tooling don't require Outlook to import this module.
    Raises on COM failure; the caller maps that to an HTTP 500 (R13)."""
    import win32com.client  # lazy import (R11/R13)
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = subject
    mail.To = "; ".join(to or [])
    mail.CC = "; ".join(cc or [])
    mail.HTMLBody = html
    mail.Display(False)  # open as editable draft — NEVER .Send()
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS

- [ ] **Step 5: 安全核对**（ai-agent-tool / business-logic）：只 `.Display()`，源码级测试禁止 `.Send()`（R10）；COM 异常向上抛给端点处理（R13）。✓

- [ ] **Step 6: 提交**

```bash
git add dfv_tool/email_report.py dfv_tool/test_weekly_email.py
git commit -m "feat(email): open_outlook_draft COM wrapper (Display only, never Send)"
```

---

## Task 5: Flask 端点 POST /api/email/weekly

**Files:**
- Modify: `dfv_tool/app.py`（新增端点 + `import email_report` + `_prev_run` 辅助）
- Modify: `dfv_tool/test_weekly_email.py`

覆盖 R1（被前端调用的端点）、R2（200/400/404）、R13（COM 失败→500）、R14（审计日志）。

- [ ] **Step 1: 写失败测试**

在 `test_weekly_email.py` 顶部 `import email_report` 下方加 `import history` 与 `import app as app_module`：

```python
import history
import app as app_module
```

追加测试：

```python
def _client(monkeypatch_runs, draft_stub):
    """Build a Flask test client with history.get_all_data and the Outlook COM
    call replaced. Returns (client, calls) where calls records draft invocations."""
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


def test_endpoint_com_failure_500():
    def boom(*a, **k):
        raise RuntimeError("Outlook not installed")
    client = _client([_run(id=2, errors=[_err()])], boom)
    r = client.post("/api/email/weekly", json={"run_id": 2})
    assert r.status_code == 500, r.status_code
    assert r.get_json()["ok"] is False, r.get_json()
    print("PASS test_endpoint_com_failure_500")
```

并在 `if __name__` 追加：

```python
    test_endpoint_generates_draft_200()
    test_endpoint_bad_run_id_404()
    test_endpoint_missing_run_id_400()
    test_endpoint_com_failure_500()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: FAIL — 端点不存在，返回 404（对 200 用例失败）

- [ ] **Step 3: 写最小实现**

在 `app.py` 顶部 import 区，`from dashboard import _build_html` 之后加：

```python
import email_report
```

在 `update_error` 端点之后、`return app` 之前，新增：

```python
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
        if not isinstance(run_id, int):
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
```

- [ ] **Step 4: 跑测试确认通过**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS（全部 build + recipients + endpoint 测试）

- [ ] **Step 5: 安全核对**（api-endpoint / logging-monitoring / request-browser）：`run_id` 强类型校验（非 int→400）；查库用现有 `get_all_data`（参数化）；日志只记 run_id + remote_addr，无收件人 PII（R14）；COM 异常统一 500、`str(e)` 不含机密。✓

- [ ] **Step 6: 提交**

```bash
git add dfv_tool/app.py dfv_tool/test_weekly_email.py
git commit -m "feat(app): POST /api/email/weekly generates Outlook draft"
```

---

## Task 6: 前端按钮 + generateEmail()

**Files:**
- Modify: `dfv_tool/dashboard.py`（顶部工具条加按钮 + `generateEmail()` JS）
- Modify: `dfv_tool/test_weekly_email.py`（模板接线断言 + 内联 JS 语法检查）

覆盖 R1（按钮触发端点）。

- [ ] **Step 1: 写失败测试**

在 `test_weekly_email.py` 顶部加 `import dashboard`，并追加：

```python
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
```

并在 `if __name__` 追加：

```python
    test_dashboard_email_button_wired()
    test_dashboard_inline_js_syntax_ok_email()
```

- [ ] **Step 2: 跑测试确认失败**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: FAIL — `email button label missing`

- [ ] **Step 3: 写最小实现（HTML 按钮）**

在 `dashboard.py` 的 Action Items 标题栏，`<button class="copy-btn" onclick="copyTable()">Copy Table</button>` 之后加一行：

```html
    <button class="copy-btn" onclick="generateEmail(this)">📧 生成周报邮件</button>
```

- [ ] **Step 4: 写最小实现（JS 函数）**

在 `dashboard.py` 的 `copyTable()` 函数定义之前（或 `saveCell` 附近的函数区）加入 `generateEmail()`。它取当前选中周的 `run_id`（`DATA[idx].id`）POST 到端点：

```javascript
function generateEmail(btn) {
  var idx = document.getElementById("weekPicker").value;
  var run = DATA[idx];
  if (!run) { alert("No week selected"); return; }
  if (btn) { btn.textContent = "生成中…"; btn.disabled = true; }
  fetch("/api/email/weekly", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({run_id: run.id})
  }).then(function(r) {
    return r.json().then(function(j) { return {ok: r.ok, j: j}; });
  }).then(function(res) {
    if (btn) { btn.textContent = "📧 生成周报邮件"; btn.disabled = false; }
    if (res.ok && res.j.ok) { alert("已在 Outlook 打开草稿，请核对后发送。"); }
    else { alert("生成失败：" + ((res.j && res.j.error) || "unknown")); }
  }).catch(function() {
    if (btn) { btn.textContent = "📧 生成周报邮件"; btn.disabled = false; }
    alert("生成失败（网络或服务器错误）");
  });
}
```

- [ ] **Step 5: 跑测试确认通过**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: PASS（含按钮接线 + 内联 JS 语法检查）

- [ ] **Step 6: 回归 — 跑既有测试确保没弄坏 Dashboard**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_editable_actionplan.py`
Expected: `ALL EDITABLE TESTS PASSED`

- [ ] **Step 7: 提交**

```bash
git add dfv_tool/dashboard.py dfv_tool/test_weekly_email.py
git commit -m "feat(dashboard): add Generate Weekly Email button + generateEmail()"
```

---

## Task 7: .gitignore、真实配置落地、清理、onboarding

**Files:**
- Modify: `.gitignore`
- Create（本地，不提交）: `dfv_tool/email_config.json`
- Modify: `.github/project-onboarding.md`

覆盖 R4（PII 不入库）与收尾。

- [ ] **Step 1: 更新 .gitignore**

在 `.gitignore` 的 `# Secrets` 段追加：

```
# Email recipients (internal PII) + reference .msg + temp tooling
dfv_tool/email_config.json
*.msg
.tmp_msgtools/
```

- [ ] **Step 2: 落地真实收件人配置（本地，不提交）**

创建 `dfv_tool/email_config.json`（用参考邮件里的真实分发列表；此文件已被 .gitignore 忽略）：

```json
{
  "to": [
    "Xie, XueYing <xie.xy@pg.com>",
    "Xu, Lucy <xu.x.17@pg.com>",
    "HC SIP team <HCSIPteam@pgone.onmicrosoft.com>",
    "GC HC CSP Team <GCHCCSPTeam@pgone.onmicrosoft.com>",
    "GC-HC-IOL <GC-HC-IOL@pgone.onmicrosoft.com>",
    "Zeng, Rebecca <zeng.re@pg.com>",
    "Wenhao, Wu <wenhao.wu@pg.com>",
    "Chen, Guimin <chen.g.9@pg.com>",
    "Li, Yanrong <li.y.47@pg.com>",
    "Sun, Yuchen <sun.y.25@pg.com>",
    "GC DRP <GCDRP@pgone.onmicrosoft.com>",
    "Huang, Molly <huang.h.9@pg.com>",
    "Zhong, Doris <zhong.d@pg.com>",
    "Liu, Becky <liu.b.2@pg.com>",
    "Wang, Lia <wang.r.19@pg.com>"
  ],
  "cc": [
    "Xu, Xia <xu.xi.3@pg.com>",
    "Chen, Mingjia <chen.m.29@pg.com>"
  ]
}
```

- [ ] **Step 3: 确认配置未被 git 跟踪**

Run: `git status --porcelain dfv_tool/email_config.json`
Expected: 无输出（被忽略）。若出现，检查 .gitignore 路径。

- [ ] **Step 4: 清理临时文件**

Run:
```
Remove-Item -Recurse -Force .tmp_msgtools -ErrorAction SilentlyContinue
```
（`.msg` 参考文件保留在本地但已被忽略，不删。）

- [ ] **Step 5: 回写 onboarding**

在 `.github/project-onboarding.md` 的 Web 应用/功能小节补一句：新增"生成周报邮件"功能（`dfv_tool/email_report.py` + `POST /api/email/weekly`），收件人存本地 gitignored `email_config.json`（模板见 `email_config.example.json`），依赖 Outlook + pywin32，只 `.Display()` 不 `.Send()`。

- [ ] **Step 6: 提交**

```bash
git add .gitignore .github/project-onboarding.md
git commit -m "chore(email): gitignore recipients/.msg, onboarding note"
```

---

## Task 8: 端到端冒烟 + 完成前验证

**Files:** 无（验证任务）

- [ ] **Step 1: 全量跑本功能测试**

Run: `.\.venv\Scripts\python.exe dfv_tool\test_weekly_email.py`
Expected: `ALL WEEKLY EMAIL TESTS PASSED`

- [ ] **Step 2: 回归既有测试**

Run:
```
.\.venv\Scripts\python.exe dfv_tool\test_editable_actionplan.py
.\.venv\Scripts\python.exe dfv_tool\test_first_seen_sort.py
```
Expected: 两个都全绿。

- [ ] **Step 3: 真实启动冒烟**

停掉旧实例后启动：`.\.venv\Scripts\python.exe dfv_tool\app.py`，浏览器打开 `http://localhost:8000/`，点"📧 生成周报邮件"。
Expected: Outlook 弹出草稿，主题/收件人/正文正确；**不自动发送**。核对后可手动发送。

- [ ] **Step 4: 勾掉 spec 未决项确认 & 汇报**

对照 spec R1–R14 逐条确认已实现；向用户汇报完成情况，并提示：本地已提交，**推送需用户同意**（铁律）。

---

## 备注

- **推送**：所有 commit 在本地进行；`git push` 须先征得用户同意。
- **每个 commit** 末尾附：`Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`。
- **改规范**走 `changes/` 提案流程，不直接改 `specs/`。
