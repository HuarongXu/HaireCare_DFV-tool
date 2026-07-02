"""
DFV Dashboard Generator - Creates a single HTML file with:
  1. Current week KPI status cards
  2. Action items table (grouped by Owner, filterable)
  3. Historical trend charts (volume diff%, error count)
  4. Week selector to view past weeks' actions
"""
import json
import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from history import get_all_data

OUTPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output")


def generate_dashboard():
    """Generate DFV_Dashboard.html with all history data."""
    runs = get_all_data()
    if not runs:
        print("No history data. Run the pipeline first.")
        return

    # Prepare JSON data for embedding
    # runs[0] is the latest (DESC order)
    data_json = json.dumps(runs, ensure_ascii=False, default=str)

    html = _build_html(data_json, runs)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, "DFV_Dashboard.html")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Dashboard: {out_path}")
    return out_path


def _build_html(data_json, runs):
    # Use string.replace instead of f-string to avoid brace escaping issues
    template = _get_template()
    # Harden JSON embedded in <script>: json.dumps does not escape </script>,
    # so user-editable fields (owner/action_plan) could break out of the script
    # tag (stored XSS). \u-escapes are equivalent inside JSON string literals.
    safe = (data_json.replace('&', '\\u0026')
                     .replace('<', '\\u003c')
                     .replace('>', '\\u003e')
                     .replace('\u2028', '\\u2028')
                     .replace('\u2029', '\\u2029'))
    return template.replace('__DATA_JSON__', safe)


def _get_template():
    return r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DFV Dashboard</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root {
  --bg:#eef2f7; --card:#ffffff; --text:#0f172a; --text2:#64748b; --text3:#94a3b8;
  --border:#e2e8f0; --border2:#eef2f6;
  --primary:#2563eb; --primary-d:#1d4ed8; --primary-l:#eff4ff;
  --green:#16a34a; --red:#dc2626; --amber:#d97706;
  --shadow-sm:0 1px 2px rgba(15,23,42,.04), 0 1px 3px rgba(15,23,42,.06);
  --shadow-md:0 6px 18px rgba(15,23,42,.08);
  --radius:12px;
}
* { margin:0; padding:0; box-sizing:border-box; }
body { font-family:'Inter','Segoe UI',system-ui,sans-serif; background:var(--bg); color:var(--text);
       font-variant-numeric:tabular-nums; -webkit-font-smoothing:antialiased; }

.topbar { background:linear-gradient(120deg,#0f172a 0%,#15233f 55%,#1e3a8a 150%); color:#fff;
          padding:15px 32px; display:flex; align-items:center; justify-content:space-between;
          box-shadow:var(--shadow-sm); position:sticky; top:0; z-index:100; }
.topbar .brand { display:flex; align-items:center; gap:13px; }
.topbar .logo { width:36px; height:36px; border-radius:10px; background:linear-gradient(135deg,#3b82f6,#2563eb);
                display:flex; align-items:center; justify-content:center; font-weight:800; font-size:13px;
                letter-spacing:.5px; color:#fff; box-shadow:0 2px 10px rgba(37,99,235,.55); }
.topbar h1 { font-size:1.12em; font-weight:700; letter-spacing:.2px; line-height:1.2; }
.topbar h1 small { display:block; font-size:.58em; font-weight:400; color:#94a3b8; letter-spacing:.4px; margin-top:3px; }
.topbar .week-sel { display:flex; align-items:center; gap:12px; }
.topbar .sync-btn { padding:8px 16px; border-radius:8px; border:1px solid rgba(255,255,255,.18);
                    background:rgba(255,255,255,.07); color:#e2e8f0; font-size:.8em; font-weight:500;
                    cursor:pointer; transition:all .15s; }
.topbar .sync-btn:hover { background:var(--primary); color:#fff; border-color:var(--primary); }
.topbar .sync-btn.done { background:var(--green); color:#fff; border-color:var(--green); }
.topbar .sel-label { font-size:.78em; color:#94a3b8; }
.topbar select { padding:8px 12px; border-radius:8px; border:1px solid rgba(255,255,255,.18);
                 background:rgba(255,255,255,.07); color:#fff; font-size:.84em; cursor:pointer; }
.topbar select option { color:#0f172a; }

.kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:18px; padding:24px 32px 6px;
           max-width:1600px; margin:0 auto; }
.kpi { background:var(--card); border-radius:var(--radius); padding:18px 20px; box-shadow:var(--shadow-sm);
       border:1px solid var(--border); position:relative; overflow:hidden; }
.kpi::before { content:""; position:absolute; left:0; top:0; bottom:0; width:4px; background:var(--primary); }
.kpi.on-track::before { background:var(--green); }
.kpi.off-track::before { background:var(--red); }
.kpi-label { font-size:.7em; color:var(--text2); text-transform:uppercase; letter-spacing:.7px; font-weight:600; }
.kpi-value { font-size:1.9em; font-weight:700; margin:6px 0 2px; letter-spacing:-.5px; }
.kpi-sub { font-size:.78em; color:var(--text3); }
.kpi.on-track .kpi-value { color:var(--green); }
.kpi.off-track .kpi-value { color:var(--red); }

.main { padding:14px 32px 48px; max-width:1600px; margin:0 auto; }
.chart-row { display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-top:20px; }
.chart-row-3 { grid-template-columns:repeat(3,1fr); }
.chart-card { background:var(--card); border-radius:var(--radius); padding:18px 20px;
              box-shadow:var(--shadow-sm); border:1px solid var(--border); display:flex; flex-direction:column; }
.chart-card h3 { font-size:.72em; color:var(--text2); margin-bottom:14px; text-transform:uppercase;
                 letter-spacing:.6px; font-weight:700; }
.chart-card h3 span { text-transform:none; letter-spacing:0; color:var(--text3); font-weight:400; }
.chart-card canvas { max-height:260px; }

.section-title { font-size:1.05em; font-weight:700; color:var(--text); margin:30px 0 14px;
                 display:flex; align-items:center; gap:12px; }
.section-title .badge { background:var(--red); color:#fff; padding:3px 11px; border-radius:20px;
                        font-size:.7em; font-weight:600; }
.toolbar { display:flex; flex-wrap:wrap; gap:20px; align-items:center; margin-bottom:14px; }
.filter-group { display:flex; align-items:center; gap:9px; }
.filter-label { font-size:.68em; text-transform:uppercase; letter-spacing:.6px; color:var(--text3); font-weight:700; }
.filters { display:flex; gap:7px; flex-wrap:wrap; }
.filters button { padding:6px 14px; border:1px solid var(--border); border-radius:18px;
                  background:#fff; color:var(--text2); cursor:pointer; font-size:.78em; font-weight:500;
                  transition:all .15s; }
.filters button:hover { border-color:var(--primary); color:var(--primary); background:var(--primary-l); }
.filters button.active { background:var(--primary); color:#fff; border-color:var(--primary);
                         box-shadow:0 2px 6px rgba(37,99,235,.3); }

.table-wrap { background:#fff; border-radius:var(--radius); border:1px solid var(--border);
              box-shadow:var(--shadow-sm); overflow:hidden; }
.action-table { width:100%; border-collapse:collapse; font-size:.82em; }
.action-table thead th { background:#0f172a; color:#e2e8f0; padding:11px 12px; text-align:left;
                         font-weight:600; font-size:.92em; letter-spacing:.2px; position:sticky; top:0; white-space:nowrap; }
.action-table tbody td { padding:10px 12px; border-bottom:1px solid var(--border2); color:#334155; }
.action-table tbody tr:nth-child(even) { background:#fafbfc; }
.action-table tbody tr:hover { background:var(--primary-l); }
.action-table tbody tr:last-child td { border-bottom:none; }
.action-table .owner-cell { white-space:nowrap; }
.tag { display:inline-block; padding:3px 9px; border-radius:20px; font-size:.74em; font-weight:600; white-space:nowrap; }
.tag-dp { background:#fef3c7; color:#92400e; }
.tag-drp { background:#dbeafe; color:#1e40af; }
.tag-iol { background:#dcfce7; color:#166534; }
.tag-high { background:#fee2e2; color:#991b1b; }
.tag-mid { background:#fef3c7; color:#92400e; }
.tag-low { background:#dcfce7; color:#166534; }
.no-actions { text-align:center; padding:48px; color:var(--text3); font-size:1.05em; }
.action-table td.editable { background:#fffef5; cursor:text; outline:none; min-width:120px; }
.action-table td.editable:focus { box-shadow:inset 0 0 0 2px var(--primary); }
.action-table td.saving { background:#fef9c3; }
.action-table td.save-ok { background:#dcfce7; }
.action-table td.save-err { background:#fee2e2; }
.edit-toggle { padding:6px 12px; background:#fff; color:var(--primary); border:1px solid var(--primary);
               border-radius:8px; cursor:pointer; font-size:.78em; font-weight:600; transition:all .15s; }
.edit-toggle:hover { background:#f1f5f9; }
.edit-toggle.on { background:var(--primary); color:#fff; }
.copy-btn { padding:8px 16px; background:var(--primary); color:#fff; border:none; border-radius:8px;
            cursor:pointer; font-size:.8em; font-weight:600; margin-left:auto; transition:all .15s;
            box-shadow:0 2px 6px rgba(37,99,235,.25); }
.copy-btn:hover { background:var(--primary-d); }
.copy-btn.copied { background:var(--green); box-shadow:none; }
@media(max-width:1100px) { .chart-row-3 { grid-template-columns:1fr; } }
@media(max-width:900px) {
  .kpi-row { grid-template-columns:repeat(2,1fr); }
  .chart-row { grid-template-columns:1fr; }
}
@media(max-width:600px) {
  .kpi-row { grid-template-columns:1fr; }
  .topbar { flex-direction:column; gap:12px; }
  .main { padding:14px 16px 40px; }
}
</style>
</head>
<body>

<div class="topbar">
  <div class="brand">
    <div class="logo">DFV</div>
    <h1>Demand Flow Validation<small>Forecast Flow Health Check</small></h1>
  </div>
  <div class="week-sel">
    <button class="sync-btn" onclick="syncOwners()" title="Copy sync command to clipboard">Sync Owners</button>
    <span class="sel-label">Week</span>
    <select id="weekPicker" onchange="switchWeek(this.value)"></select>
  </div>
</div>

<div class="kpi-row" id="kpiRow"></div>

<div class="main">
  <div class="chart-row">
    <div class="chart-card">
      <h3>Volume Difference % (18 Months) &mdash; Target &lt; 2%</h3>
      <canvas id="chartDiff"></canvas>
    </div>
    <div class="chart-card">
      <h3>Error SKU Count (13 Weeks)</h3>
      <canvas id="chartSKU"></canvas>
    </div>
  </div>

  <div class="chart-row chart-row-3">
    <div class="chart-card">
      <h3>Action Priority <span>(current view)</span></h3>
      <canvas id="chartPriority"></canvas>
    </div>
    <div class="chart-card">
      <h3>Aging Distribution <span>(weeks since first seen)</span></h3>
      <canvas id="chartAging"></canvas>
    </div>
    <div class="chart-card">
      <h3>By Owner <span>(actionable items)</span></h3>
      <canvas id="chartOwner"></canvas>
    </div>
  </div>

  <div class="section-title">
    Action Items <span class="badge" id="actionCount">0</span>
    <button class="copy-btn" onclick="copyTable()">Copy Table</button>
    <button class="copy-btn" onclick="generateEmail(this)">📧 生成周报邮件</button>
  </div>
  <div class="toolbar">
    <div class="filter-group"><span class="filter-label">Priority</span><div class="filters" id="priorityFilters"></div></div>
    <div class="filter-group"><span class="filter-label">Owner</span><div class="filters" id="ownerFilters"></div></div>
    <div class="filter-group"><button id="ownerEditToggle" class="edit-toggle" onclick="toggleOwnerEdit()">&#128274; 允许编辑</button></div>
  </div>
  <div class="table-wrap" id="actionTableWrap"></div>
</div>

<script>
var DATA = __DATA_JSON__;

function escapeHtml(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function saveCell(cell) {
  var id = cell.getAttribute("data-id");
  var field = cell.getAttribute("data-field");
  var val = cell.innerText.trim();
  cell.classList.remove("save-ok", "save-err");
  cell.classList.add("saving");
  var body = {}; body[field] = val;
  fetch("/api/errors/" + id, {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify(body)
  }).then(function(r) {
    cell.classList.remove("saving");
    if (r.ok) {
      cell.classList.add("save-ok");
      if (field === "owner") { applyOwnerEdit(id, val); }
    }
    else { cell.classList.add("save-err"); alert("Save failed (" + r.status + ")"); }
  }).catch(function() {
    cell.classList.remove("saving"); cell.classList.add("save-err");
    alert("Save failed (network)");
  });
}

// Owner + Action Plan editing is locked by default to avoid accidental edits;
// a toolbar toggle unlocks both.
var editEnabled = false;
function toggleOwnerEdit() {
  editEnabled = !editEnabled;
  var btn = document.getElementById("ownerEditToggle");
  if (btn) {
    btn.textContent = editEnabled ? "🔓 可编辑（点此锁定）" : "🔒 允许编辑";
    btn.classList.toggle("on", editEnabled);
  }
  var idx = document.getElementById("weekPicker").value;
  renderActions((DATA[idx] && DATA[idx].errors) || []);
}

// After an owner is saved, reflect the (possibly new) name in the in-memory data
// so the Owner filter list includes it and no rows are hidden/missed.
function applyOwnerEdit(id, val) {
  var idx = document.getElementById("weekPicker").value;
  var errs = (DATA[idx] && DATA[idx].errors) || [];
  for (var i = 0; i < errs.length; i++) {
    if (String(errs[i].id) === String(id)) { errs[i].owner = val; break; }
  }
  renderActions(errs);
}

var picker = document.getElementById("weekPicker");
for (var i = 0; i < DATA.length; i++) {
  var opt = document.createElement("option");
  opt.value = i;
  opt.textContent = DATA[i].week_label + " (" + DATA[i].run_date.split(" ")[0] + ")";
  picker.appendChild(opt);
}

var currentFilter = "all";
var currentPriority = "all";
var priorityChart = null, agingChart = null, ownerChart = null;

// Inline plugins to draw value labels directly on the charts (no extra CDN).
var doughnutLabels = {
  id: "doughnutLabels",
  afterDatasetsDraw: function(chart) {
    var ctx = chart.ctx;
    chart.getDatasetMeta(0).data.forEach(function(arc, i) {
      var val = chart.data.datasets[0].data[i];
      if (!val) return;
      var p = arc.tooltipPosition();
      ctx.save();
      ctx.fillStyle = "#fff";
      ctx.font = "700 13px Inter, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(val, p.x, p.y);
      ctx.restore();
    });
  }
};
var barTopLabels = {
  id: "barTopLabels",
  afterDatasetsDraw: function(chart) {
    var ctx = chart.ctx;
    chart.getDatasetMeta(0).data.forEach(function(bar, i) {
      var val = chart.data.datasets[0].data[i];
      if (val == null) return;
      ctx.save();
      ctx.fillStyle = "#334155";
      ctx.font = "700 12px Inter, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "bottom";
      ctx.fillText(val, bar.x, bar.y - 5);
      ctx.restore();
    });
  }
};
var hbarLabels = {
  id: "hbarLabels",
  afterDatasetsDraw: function(chart) {
    var ctx = chart.ctx;
    chart.getDatasetMeta(0).data.forEach(function(bar, i) {
      var val = chart.data.datasets[0].data[i];
      if (val == null) return;
      ctx.save();
      ctx.fillStyle = "#334155";
      ctx.font = "700 12px Inter, sans-serif";
      ctx.textAlign = "left";
      ctx.textBaseline = "middle";
      ctx.fillText(val, bar.x + 6, bar.y);
      ctx.restore();
    });
  }
};

function switchWeek(idx) {
  var run = DATA[idx];
  renderKPI(run);
  renderActions(run.errors || []);
}

function fmt(n) {
  if (n == null) return "?";
  if (n >= 1e6) return (n / 1e6).toFixed(1) + "M";
  if (n >= 1e3) return (n / 1e3).toFixed(0) + "K";
  return n.toFixed(0);
}

function renderKPI(run) {
  var diff = run.diff_pct != null ? run.diff_pct.toFixed(2) : "?";
  var volCls = run.vol_status === "on-track" ? "on-track" : "off-track";
  var skuCls = run.sku_status === "on-track" ? "on-track" : "off-track";
  document.getElementById("kpiRow").innerHTML =
    '<div class="kpi ' + volCls + '">' +
      '<div class="kpi-label">Volume Diff %</div>' +
      '<div class="kpi-value">' + diff + '%</div>' +
      '<div class="kpi-sub">Target &lt; 2% &mdash; ' + run.vol_status.toUpperCase() + '</div>' +
    '</div>' +
    '<div class="kpi ' + skuCls + '">' +
      '<div class="kpi-label">Error SKUs</div>' +
      '<div class="kpi-value">' + run.total_errors + '</div>' +
      '<div class="kpi-sub">' + run.actionable_errors + ' actionable, ' + run.hktw_errors + ' HKTW</div>' +
    '</div>' +
    '<div class="kpi">' +
      '<div class="kpi-label">Demand Hub</div>' +
      '<div class="kpi-value">' + fmt(run.total_idp) + '</div>' +
      '<div class="kpi-sub">Total IDP Forecast</div>' +
    '</div>' +
    '<div class="kpi">' +
      '<div class="kpi-label">APO-DRP</div>' +
      '<div class="kpi-value">' + fmt(run.total_apo) + '</div>' +
      '<div class="kpi-sub">Impact: ' + fmt(run.impact_volume) + ' MSU</div>' +
    '</div>';
}

function renderActions(errors) {
  // Items matching the active Priority facet. Owner buttons are derived from THIS set,
  // so every owner shown has >=1 item under the current priority (never empty on click).
  var priorityScoped = currentPriority === "all" ? errors :
    errors.filter(function(e) { return (e.priority || "") === currentPriority; });

  // Owner filter options: distinct, non-blank owners present in the priority-scoped set.
  var owners = [];
  var seen = {};
  for (var i = 0; i < priorityScoped.length; i++) {
    var ow = priorityScoped[i].owner;
    if (ow && !seen[ow]) { seen[ow] = true; owners.push(ow); }
  }
  owners.sort();
  // If the selected owner no longer has items under this priority, fall back to All.
  if (currentFilter !== "all" && owners.indexOf(currentFilter) < 0) currentFilter = "all";

  var fhtml = '<button class="' + (currentFilter === "all" ? "active" : "") +
    '" data-owner="all">All</button>';
  for (var j = 0; j < owners.length; j++) {
    fhtml += '<button class="' + (currentFilter === owners[j] ? "active" : "") +
      '" data-owner="' + escapeHtml(owners[j]) + '">' + escapeHtml(owners[j]) + '</button>';
  }
  var ownerFilters = document.getElementById("ownerFilters");
  ownerFilters.innerHTML = fhtml;
  var ownerBtns = ownerFilters.getElementsByTagName("button");
  for (var k = 0; k < ownerBtns.length; k++) {
    ownerBtns[k].onclick = function() { setFilter(this.getAttribute("data-owner")); };
  }

  // Priority quick-filter buttons + summary charts reflect the current owner selection.
  var ownerScoped = currentFilter === "all" ? errors :
    errors.filter(function(e) { return e.owner === currentFilter; });
  renderPriorityFilters(ownerScoped);
  renderPriorityChart(ownerScoped);
  renderAgingChart(ownerScoped);
  // By-Owner chart is a cross-owner overview -> always the full week, not faceted.
  renderOwnerChart(errors);

  // Table = both facets applied (priority AND owner).
  var filtered = priorityScoped.filter(function(e) {
    return currentFilter === "all" || e.owner === currentFilter;
  });
  // Sort: Owner ascending, then Duration descending (longest-standing first).
  filtered = filtered.slice().sort(function(a, b) {
    var oa = a.owner || "", ob = b.owner || "";
    if (oa < ob) return -1;
    if (oa > ob) return 1;
    var da = a.duration != null ? a.duration : -1;
    var db = b.duration != null ? b.duration : -1;
    return db - da;
  });
  document.getElementById("actionCount").textContent = filtered.length;

  if (filtered.length === 0) {
    document.getElementById("actionTableWrap").innerHTML =
      '<div class="no-actions">No action items for this week</div>';
    return;
  }

  var html = '<table class="action-table" id="actionTable"><thead><tr>' +
    '<th>Product</th><th>Description</th><th>Brand</th><th>Location</th><th>Error</th>' +
    '<th>Forecast</th><th>Reason</th><th>Action</th>' +
    '<th>First Time</th><th>Duration</th><th>Priority</th>' +
    '<th>Owner</th><th>Action Plan</th>' +
    '</tr></thead><tbody>';

  for (var k = 0; k < filtered.length; k++) {
    var e = filtered[k];
    var fcst = Math.round(e.idp_forecast).toLocaleString();
    html += '<tr>' +
      '<td><strong>' + escapeHtml(e.apo_product) + '</strong></td>' +
      '<td>' + escapeHtml(e.description) + '</td>' +
      '<td>' + escapeHtml(e.brand) + '</td>' +
      '<td>' + escapeHtml(e.apo_location) + '</td>' +
      '<td>' + escapeHtml(e.error_message) + '</td>' +
      '<td style="text-align:right">' + fcst + '</td>' +
      '<td>' + escapeHtml(e.reason) + '</td>' +
      '<td><strong>' + escapeHtml(e.action) + '</strong></td>' +
      '<td class="owner-cell">' + escapeHtml(e.first_time) + '</td>' +
      '<td style="text-align:center">' + (e.duration != null ? e.duration : "") + '</td>' +
      '<td style="text-align:center">' + (e.priority ? '<span class="tag ' + (e.priority === "High" ? "tag-high" : e.priority === "Mid" ? "tag-mid" : "tag-low") + '">' + escapeHtml(e.priority) + '</span>' : "") + '</td>' +
      '<td class="owner-cell' + (editEnabled ? ' editable" contenteditable="true" onblur="saveCell(this)' : '') +
        '" data-id="' + e.id + '" data-field="owner">' +
        escapeHtml(e.owner) + '</td>' +
      '<td class="' + (editEnabled ? 'editable" contenteditable="true" onblur="saveCell(this)' : '') +
        '" data-id="' + e.id + '" data-field="action_plan">' +
        escapeHtml(e.action_plan) + '</td>' +
      '</tr>';
  }
  html += '</tbody></table>';
  document.getElementById("actionTableWrap").innerHTML = html;
}

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

function copyTable() {
  var table = document.getElementById("actionTable");
  if (!table) return;
  // Build HTML table for rich paste (email)
  var hdr = '#1a1a2e';
  var html = '<table style="border-collapse:collapse;font-family:Segoe UI,sans-serif;font-size:12px;">';
  var rows = table.querySelectorAll("tr");
  for (var i = 0; i < rows.length; i++) {
    html += '<tr>';
    var cells = rows[i].querySelectorAll("th, td");
    for (var j = 0; j < cells.length; j++) {
      var tag = cells[j].tagName === 'TH' ? 'th' : 'td';
      var style = tag === 'th'
        ? 'background:' + hdr + ';color:#fff;padding:6px 10px;text-align:left;border:1px solid #ccc;'
        : 'padding:5px 10px;border:1px solid #e0e0e0;' + (j === 5 ? 'text-align:right;' : '');
      html += '<' + tag + ' style="' + style + '">' + cells[j].innerText + '</' + tag + '>';
    }
    html += '</tr>';
  }
  html += '</table>';
  // Also build plain text version
  var tsv = [];
  for (var i2 = 0; i2 < rows.length; i2++) {
    var cells2 = rows[i2].querySelectorAll("th, td");
    var row = [];
    for (var j2 = 0; j2 < cells2.length; j2++) row.push(cells2[j2].innerText.replace(/\t/g, ' '));
    tsv.push(row.join('\t'));
  }
  var plain = tsv.join('\n');
  // Write both HTML and plain text to clipboard
  var blob = new Blob([html], {type: 'text/html'});
  var blobText = new Blob([plain], {type: 'text/plain'});
  var item = new ClipboardItem({'text/html': blob, 'text/plain': blobText});
  navigator.clipboard.write([item]).then(function() {
    var btn = document.querySelector('.copy-btn');
    btn.textContent = 'Copied!';
    btn.classList.add('copied');
    setTimeout(function() { btn.textContent = 'Copy Table'; btn.classList.remove('copied'); }, 2000);
  });
}

function syncOwners() {
  var cmd = "sync_owners.bat";
  navigator.clipboard.writeText(cmd).then(function() {
    var btn = document.querySelector(".sync-btn");
    btn.textContent = "Run sync_owners.bat";
    btn.classList.add("done");
    setTimeout(function() { btn.textContent = "Sync Owners"; btn.classList.remove("done"); }, 3000);
  });
}

function setFilter(f) {
  currentFilter = f;
  var idx = document.getElementById("weekPicker").value;
  renderActions(DATA[idx].errors || []);
}

function setPriority(p) {
  currentPriority = p;
  var idx = document.getElementById("weekPicker").value;
  renderActions(DATA[idx].errors || []);
}

function priorityCounts(items) {
  var c = { High: 0, Mid: 0, Low: 0 };
  for (var i = 0; i < items.length; i++) {
    if (c[items[i].priority] != null) c[items[i].priority]++;
  }
  return c;
}

function renderPriorityFilters(items) {
  var c = priorityCounts(items);
  var defs = [["all", "All", items.length], ["High", "High", c.High],
              ["Mid", "Mid", c.Mid], ["Low", "Low", c.Low]];
  var html = "";
  for (var i = 0; i < defs.length; i++) {
    html += '<button class="' + (currentPriority === defs[i][0] ? "active" : "") +
      '" onclick="setPriority(&quot;' + defs[i][0] + '&quot;)">' +
      defs[i][1] + ' (' + defs[i][2] + ')</button>';
  }
  document.getElementById("priorityFilters").innerHTML = html;
}

function renderPriorityChart(items) {
  var ctx = document.getElementById("chartPriority");
  if (!ctx || typeof Chart === "undefined") return;
  var c = priorityCounts(items);
  if (priorityChart) priorityChart.destroy();
  priorityChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: ["High", "Mid", "Low"],
      datasets: [{ data: [c.High, c.Mid, c.Low],
        backgroundColor: ["#dc2626", "#d97706", "#16a34a"], borderWidth: 2, borderColor: "#fff" }]
    },
    options: { responsive: true, cutout: "60%",
      plugins: { legend: { position: "bottom",
        labels: { usePointStyle: true, pointStyle: "circle", padding: 14, font: { size: 11 } } } } },
    plugins: [doughnutLabels]
  });
}

function renderAgingChart(items) {
  var ctx = document.getElementById("chartAging");
  if (!ctx || typeof Chart === "undefined") return;
  var b = [0, 0, 0]; // 1-2, 3-4, 5+ (the "0/new" bucket is intentionally excluded)
  for (var i = 0; i < items.length; i++) {
    var d = items[i].duration;
    if (d == null || d === 0) continue;
    if (d <= 2) b[0]++;
    else if (d <= 4) b[1]++;
    else b[2]++;
  }
  if (agingChart) agingChart.destroy();
  agingChart = new Chart(ctx, {
    type: "bar",
    data: {
      labels: ["1-2 wk", "3-4 wk", "5+ wk"],
      datasets: [{ label: "Items", data: b,
        backgroundColor: ["#16a34a", "#d97706", "#dc2626"], borderRadius: 5, borderSkipped: false }]
    },
    options: { responsive: true, layout: { padding: { top: 16 } },
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#eef2f6" },
        title: { display: true, text: "Items" } },
        x: { grid: { display: false } } } },
    plugins: [barTopLabels]
  });
}

function renderOwnerChart(items) {
  var ctx = document.getElementById("chartOwner");
  if (!ctx || typeof Chart === "undefined") return;
  var counts = {};
  for (var i = 0; i < items.length; i++) {
    var o = items[i].owner;
    if (!o) continue; // skip unassigned/blank owners (not filterable)
    counts[o] = (counts[o] || 0) + 1;
  }
  var entries = Object.keys(counts).map(function(k) { return [k, counts[k]]; });
  entries.sort(function(a, b) { return b[1] - a[1]; });
  var labels = entries.map(function(e) { return e[0]; });
  var data = entries.map(function(e) { return e[1]; });
  if (ownerChart) ownerChart.destroy();
  ownerChart = new Chart(ctx, {
    type: "bar",
    data: { labels: labels, datasets: [{ label: "Items", data: data,
      backgroundColor: "#2563eb", hoverBackgroundColor: "#1d4ed8", borderRadius: 5, barThickness: 16 }] },
    options: { indexAxis: "y", responsive: true,
      layout: { padding: { right: 24 } },
      plugins: { legend: { display: false } },
      scales: { x: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: "#eef2f6" } },
        y: { grid: { display: false }, ticks: { font: { size: 11 } } } },
      onClick: function(e, els) { if (els.length) { setFilter(labels[els[0].index]); } } },
    plugins: [hbarLabels]
  });
}

function renderCharts() {
  var sorted = DATA.slice().reverse();
  var labels = sorted.map(function(r) { return r.week_label; });
  var diffs = sorted.map(function(r) { return r.diff_pct; });
  var skus = sorted.map(function(r) { return r.total_errors; });
  var actionable = sorted.map(function(r) { return r.actionable_errors; });

  new Chart(document.getElementById("chartDiff"), {
    type: "line",
    data: {
      labels: labels,
      datasets: [
        { label: "Volume Diff %", data: diffs, borderColor: "#0078d4",
          backgroundColor: "rgba(0,120,212,0.08)", fill: true, tension: 0.3, pointRadius: 4 },
        { label: "Target (2%)", data: labels.map(function() { return 2; }),
          borderColor: "#dc3545", borderDash: [6,4], pointRadius: 0, borderWidth: 2 }
      ]
    },
    options: { responsive: true, plugins: { legend: { position: "bottom" } },
               scales: { y: { min: 0.2, ticks: { stepSize: 0.1 }, title: { display: true, text: "%" } } } }
  });

  new Chart(document.getElementById("chartSKU"), {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        { label: "Total Errors", data: skus, backgroundColor: "rgba(220,53,69,0.15)",
          borderColor: "#dc3545", borderWidth: 1 },
        { label: "Actionable", data: actionable, backgroundColor: "rgba(0,120,212,0.6)",
          borderColor: "#0078d4", borderWidth: 1 }
      ]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { position: "bottom" },
        datalabels: {
          anchor: "end", align: "top", font: { weight: "bold", size: 11 },
          formatter: function(v) { return v; }
        }
      },
      scales: { y: { beginAtZero: true, title: { display: true, text: "SKU Count" } } }
    },
    plugins: [{
      id: "barLabels",
      afterDatasetsDraw: function(chart) {
        var ctx = chart.ctx;
        chart.data.datasets.forEach(function(ds, di) {
          var meta = chart.getDatasetMeta(di);
          meta.data.forEach(function(bar, idx) {
            var val = ds.data[idx];
            if (val == null) return;
            ctx.save();
            ctx.fillStyle = ds.borderColor || "#333";
            ctx.font = "bold 11px Segoe UI,sans-serif";
            ctx.textAlign = "center";
            ctx.fillText(val, bar.x, bar.y - 6);
            ctx.restore();
          });
        });
      }
    }]
  });
}

switchWeek(0);
if (typeof Chart !== "undefined") {
  renderCharts();
} else {
  var cards = document.querySelectorAll(".chart-card");
  for (var c = 0; c < cards.length; c++) {
    cards[c].innerHTML = '<p style="color:#999;padding:20px">Chart.js CDN not reachable.</p>';
  }
}
</script>
</body>
</html>'''


if __name__ == "__main__":
    generate_dashboard()
