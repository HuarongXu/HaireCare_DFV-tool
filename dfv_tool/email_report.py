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
