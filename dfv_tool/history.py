"""
DFV History - SQLite storage for weekly DFV results.
Stores KPI summary + all error details for each run.
"""
import sqlite3
import os
import pandas as pd
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "dfv_history.db")


def _get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Create tables if they don't exist."""
    conn = _get_conn()
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_date TEXT NOT NULL,
            week_label TEXT,
            total_rows INTEGER,
            total_idp REAL,
            total_apo REAL,
            diff_pct REAL,
            vol_status TEXT,
            total_errors INTEGER,
            actionable_errors INTEGER,
            hktw_errors INTEGER,
            sku_status TEXT,
            impact_volume REAL
        );

        CREATE TABLE IF NOT EXISTS errors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id INTEGER NOT NULL,
            file_id TEXT,
            apo_product TEXT,
            description TEXT,
            category TEXT,
            brand TEXT,
            apo_location TEXT,
            snp_planner TEXT,
            error_message TEXT,
            idp_forecast REAL,
            apo_forecast REAL,
            reason TEXT,
            action TEXT,
            owner TEXT,
            is_hktw INTEGER,
            first_time TEXT,
            duration INTEGER,
            priority TEXT,
            action_plan TEXT,
            FOREIGN KEY (run_id) REFERENCES runs(id)
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT NOT NULL,
            ip TEXT,
            error_id INTEGER,
            field TEXT,
            old_value TEXT,
            new_value TEXT
        );
    """)
    # Migrations for existing DBs (CREATE TABLE IF NOT EXISTS won't add columns).
    for col, decl in (("first_time", "TEXT"), ("duration", "INTEGER"),
                      ("priority", "TEXT"), ("action_plan", "TEXT")):
        try:
            conn.execute(f"ALTER TABLE errors ADD COLUMN {col} {decl}")
        except sqlite3.OperationalError:
            pass  # column already exists
    conn.commit()
    conn.close()


def save_run(summary, errors_df):
    """Save a DFV run to history. Returns run_id."""
    init_db()
    conn = _get_conn()

    # Week label: e.g. "W20/2026"
    today = datetime.now()
    iso = today.isocalendar()
    week_label = f"W{iso[1]:02d}/{iso[0]}"

    cur = conn.execute("""
        INSERT INTO runs (run_date, week_label, total_rows, total_idp, total_apo,
                          diff_pct, vol_status, total_errors, actionable_errors,
                          hktw_errors, sku_status, impact_volume)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        today.strftime("%Y-%m-%d %H:%M"),
        week_label,
        summary["total_rows"],
        summary["total_idp"],
        summary["total_apo"],
        summary["diff_pct"],
        summary["vol_status"],
        summary["total_errors"],
        summary["actionable_errors"],
        summary["hktw_errors"],
        summary["sku_status"],
        summary["impact_volume_3m"],
    ))
    run_id = cur.lastrowid

    # Save error details
    if not errors_df.empty:
        rows = []
        for _, r in errors_df.iterrows():
            rows.append((
                run_id,
                r.get("File_ID", ""),
                str(r.get("APO_Product", "")),
                r.get("Description", ""),
                r.get("Category", ""),
                r.get("Brand", ""),
                r.get("APO_Location", ""),
                r.get("SNP_Planner", ""),
                r.get("Error_Message", ""),
                float(r.get("IDP_Forecast", 0)),
                float(r.get("APO_Forecast", 0)),
                r.get("Reason", ""),
                r.get("Action", ""),
                r.get("Owner", ""),
                1 if r.get("Is_HKTW", False) else 0,
                r.get("First_Time", ""),
                int(r.get("Duration", 0)) if str(r.get("Duration", "")) != "" else None,
                r.get("Priority", ""),
                r.get("Action_Plan", ""),
            ))
        conn.executemany("""
            INSERT INTO errors (run_id, file_id, apo_product, description, category,
                                brand, apo_location, snp_planner, error_message,
                                idp_forecast, apo_forecast, reason, action, owner, is_hktw,
                                first_time, duration, priority, action_plan)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)

    conn.commit()
    conn.close()
    return run_id


def get_all_runs():
    """Get all run summaries, ordered by date."""
    init_db()
    conn = _get_conn()
    rows = conn.execute("SELECT * FROM runs ORDER BY run_date DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_errors_for_run(run_id):
    """Get all errors for a specific run."""
    conn = _get_conn()
    rows = conn.execute("SELECT * FROM errors WHERE run_id = ? ORDER BY owner, error_message",
                        (run_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_first_seen_map():
    """Return the earliest run_date each code-by-plant (product+location) was
    ever recorded as an error.

    Returns {(apo_product, apo_location): first_run_date_str}. Keys are strings
    so callers can match regardless of numeric/text types. Reflects history
    BEFORE the current run is saved, so a brand-new item is absent (caller then
    treats the current run as its first appearance).
    """
    init_db()
    conn = _get_conn()
    rows = conn.execute("""
        SELECT e.apo_product AS p, e.apo_location AS l, MIN(r.run_date) AS first_date
        FROM errors e JOIN runs r ON e.run_id = r.id
        GROUP BY e.apo_product, e.apo_location
    """).fetchall()
    conn.close()
    return {(str(row["p"]), str(row["l"])): row["first_date"] for row in rows}


def get_latest_action_plan_map():
    """Return {(apo_product, apo_location): most-recent non-empty action_plan}.

    Keyed by string product/location so callers match regardless of numeric
    vs text types (mirrors get_first_seen_map). "Most recent" = the non-empty
    action_plan from the run with the largest run_date.
    """
    init_db()
    conn = _get_conn()
    rows = conn.execute("""
        SELECT e.apo_product AS p, e.apo_location AS l,
               e.action_plan AS ap, r.run_date AS rd
        FROM errors e JOIN runs r ON e.run_id = r.id
        WHERE e.action_plan IS NOT NULL AND e.action_plan <> ''
        ORDER BY r.run_date ASC
    """).fetchall()
    conn.close()
    # ASC order means later rows overwrite earlier -> last (most recent) wins.
    result = {}
    for row in rows:
        result[(str(row["p"]), str(row["l"]))] = row["ap"]
    return result


def delete_runs(run_ids):
    """Delete one or more runs (and all their error rows) by run id.

    Args:
        run_ids: a single run id or an iterable of run ids.

    Returns:
        (deleted_runs, deleted_errors) counts.
    """
    if isinstance(run_ids, (int, str)):
        run_ids = [run_ids]
    ids = [int(r) for r in run_ids]
    if not ids:
        return (0, 0)

    conn = _get_conn()
    placeholders = ",".join("?" * len(ids))
    cur_e = conn.execute(
        f"DELETE FROM errors WHERE run_id IN ({placeholders})", ids)
    deleted_errors = cur_e.rowcount
    cur_r = conn.execute(
        f"DELETE FROM runs WHERE id IN ({placeholders})", ids)
    deleted_runs = cur_r.rowcount
    conn.commit()
    conn.close()
    return (deleted_runs, deleted_errors)


def delete_run(run_id):
    """Delete a single run (and its errors) by run id."""
    return delete_runs([run_id])


def get_all_data():
    """Get all runs + their errors. For dashboard generation."""
    runs = get_all_runs()
    conn = _get_conn()
    for run in runs:
        rows = conn.execute(
            "SELECT * FROM errors WHERE run_id = ? AND is_hktw = 0 ORDER BY owner, error_message",
            (run["id"],)
        ).fetchall()
        run["errors"] = [dict(r) for r in rows]
    conn.close()
    return runs


def update_owners(mapping, run_id=None):
    """
    Batch update owners in the DB.

    Args:
        mapping: dict of {error_message: new_owner_name}
                 e.g. {"Missing Mat/Loc": "张三 / IOL", "Matloc Deleted": "李四"}
        run_id:  If None, update ALL runs. If specified, only that run.

    Example usage:
        from history import update_owners
        update_owners({"Missing Mat/Loc": "张三 / IOL", "No SNP Assigned": "王五"})
    """
    conn = _get_conn()
    for error_msg, new_owner in mapping.items():
        if run_id:
            conn.execute(
                "UPDATE errors SET owner = ? WHERE error_message = ? AND run_id = ?",
                (new_owner, error_msg, run_id)
            )
        else:
            conn.execute(
                "UPDATE errors SET owner = ? WHERE error_message = ?",
                (new_owner, error_msg)
            )
    conn.commit()
    affected = conn.execute("SELECT changes()").fetchone()[0]
    conn.close()
    print(f"Updated {affected} rows")
    return affected


def sync_owners_from_excel(excel_path, run_id=None):
    """
    Read editable columns (Owner, Action, Reason) from an edited DFV_actions
    workbook and sync them back to the DB. Matches rows by
    APO_Product + APO_Location + Error_Message.

    Reads BOTH the date-named detail sheet AND every per-owner sheet, so edits
    made in either place are picked up. Per-owner sheets are processed AFTER
    the date sheet, so if the same row was edited in a per-owner tab that value
    wins (per-owner tabs are the natural per-planner working surface).

    Column headers differ between the two sheet styles and both are handled:
        date sheet:  "APO - Product", "APO - Location", "Error Message"
        owner sheet: "APO_Product",   "APO_Location",   "Error_Message"

    Args:
        excel_path: Path to edited DFV_actions_YYYYMMDD.xlsx
        run_id: If None, update the latest run. If specified, that run.

    Usage:
        from history import sync_owners_from_excel
        sync_owners_from_excel("output/DFV_actions_20260518.xlsx")
    """
    # Aliases: every accepted header -> canonical key
    PRODUCT_HDRS = ("APO - Product", "APO_Product")
    LOCATION_HDRS = ("APO - Location", "APO_Location")
    ERROR_HDRS = ("Error Message", "Error_Message")

    def _find(df, headers):
        for h in headers:
            if h in df.columns:
                return h
        return None

    def _clean(v):
        return "" if pd.isna(v) else str(v)

    xl = pd.ExcelFile(excel_path)

    # Order sheets: date-named detail first, then per-owner sheets (override),
    # always skipping the Summary sheet.
    date_sheets = [n for n in xl.sheet_names if n.isdigit() and len(n) == 8]
    owner_sheets = [n for n in xl.sheet_names
                    if n not in date_sheets and n.lower() != "summary"]
    ordered_sheets = date_sheets + owner_sheets
    if not ordered_sheets:
        raise ValueError("No editable sheets found in workbook")

    # Determine run_id (latest run if not given)
    if run_id is None:
        conn = _get_conn()
        row = conn.execute("SELECT id FROM runs ORDER BY run_date DESC LIMIT 1").fetchone()
        if not row:
            raise ValueError("No runs in DB")
        run_id = row["id"]
        conn.close()

    conn = _get_conn()
    updated = 0
    used_action = False
    used_reason = False

    for sheet in ordered_sheets:
        df = pd.read_excel(excel_path, sheet_name=sheet)

        p_col = _find(df, PRODUCT_HDRS)
        l_col = _find(df, LOCATION_HDRS)
        e_col = _find(df, ERROR_HDRS)
        if not (p_col and l_col and e_col and "Owner" in df.columns):
            # Not a recognizable detail/owner sheet; skip it.
            continue

        has_action = "Action" in df.columns
        has_reason = "Reason" in df.columns
        used_action = used_action or has_action
        used_reason = used_reason or has_reason

        set_cols = ["owner = ?"]
        if has_action:
            set_cols.append("action = ?")
        if has_reason:
            set_cols.append("reason = ?")
        set_clause = ", ".join(set_cols)

        for _, row in df.iterrows():
            product = str(row[p_col])
            location = str(row[l_col])
            error_msg = str(row[e_col])

            params = [_clean(row["Owner"])]
            if has_action:
                params.append(_clean(row["Action"]))
            if has_reason:
                params.append(_clean(row["Reason"]))
            params.extend([run_id, product, location, error_msg])

            cur = conn.execute(f"""
                UPDATE errors SET {set_clause}
                WHERE run_id = ? AND apo_product = ? AND apo_location = ? AND error_message = ?
            """, params)
            updated += cur.rowcount

    conn.commit()
    conn.close()
    synced_cols = ["Owner"] + (["Action"] if used_action else []) + (["Reason"] if used_reason else [])
    print(f"Synced {updated} row-updates from {len(ordered_sheets)} sheet(s) "
          f"(run_id={run_id}) | columns: {', '.join(synced_cols)}")
    return updated
