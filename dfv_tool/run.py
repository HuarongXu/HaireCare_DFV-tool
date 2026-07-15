"""
DFV End-to-End Automation
=========================
1. Open Excel workbook (with AO Add-In)
2. Wait for AO ribbon to appear (via pywinauto)
3. Click Refresh All button in Analysis ribbon -> triggers Prompts dialog
4. Fill Prompts via coordinate+clipboard
5. Wait for BW data to load
6. Fast export to CSV (Range().Value bulk read)
7. Run pipeline (filter + classify + owner-assign + report)

NOTE: We do NOT use xl.Run('SAPGetProperty/SAPLogon/SAPExecuteCommand') because
      XLL function calls from COM break the AO ribbon (it disappears).
      All SAP interactions go through pywinauto ribbon button clicks.

Usage:  python run.py
Output: output/DFV_actions_YYYYMMDD.xlsx
        output/DFV_errors_YYYYMMDD.csv
"""
import win32com.client
import win32gui
import win32clipboard
import pythoncom
import ctypes
import csv
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import WORKBOOK_PATH, OUTPUT_DIR, REGION, CATEGORY

# Variant name saved in SAP AO Prompts dialog (create manually once)
# To create: open Prompts -> fill Region=XA + Category=HAIRCARE -> Save Variant -> name it
VARIANT_NAME = "HAIRCARE_XA"


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


def find_excel_by_workbook(path):
    """Find the Excel.Application that actually owns `path` via the Running
    Object Table (ROT).

    GetObject(Class="Excel.Application") binds to whichever Excel instance
    registered first in the ROT. When the user has multiple Excel instances
    open (or a stale/empty background instance), that may NOT be the instance
    holding our workbook, so Workbooks.Count stays 0 forever and the tool
    appears to hang. Matching by workbook path is reliable across instances.

    Returns (app, workbook) or (None, None) if the workbook is not open.
    """
    target = os.path.basename(path).lower()
    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
    except Exception:
        return None, None
    for mk in rot:
        try:
            name = mk.GetDisplayName(ctx, None)
        except Exception:
            continue
        if not name or not name.lower().endswith(target):
            continue
        try:
            obj = rot.GetObject(mk)
            disp = obj.QueryInterface(pythoncom.IID_IDispatch)
            wb = win32com.client.Dispatch(disp)
            return wb.Application, wb
        except Exception:
            continue
    return None, None


def data_signature(ws):
    """Lightweight content fingerprint of the Main data block.

    Used to tell whether a BW refresh has actually replaced the data, rather
    than just checking row count (the previous run's stale data is still in
    the sheet, so a row-count check passes immediately and exports OLD data).

    Returns (last_row, sum_IDP, sum_APO) or None if the sheet is busy.
    """
    try:
        last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
        if last_row < 22:
            return (last_row, 0.0, 0.0)
        wf = ws.Application.WorksheetFunction
        col_idp = ws.Range(ws.Cells(22, 9), ws.Cells(last_row, 9))
        col_apo = ws.Range(ws.Cells(22, 10), ws.Cells(last_row, 10))
        s_idp = round(float(wf.Sum(col_idp)), 3)
        s_apo = round(float(wf.Sum(col_apo)), 3)
        return (last_row, s_idp, s_apo)
    except Exception:
        return None


def real_click(x, y, pause=0.8):
    """Physical mouse click at screen coordinates."""
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.15)
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # left down
    time.sleep(0.05)
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # left up
    time.sleep(pause)


def clipboard_set(text):
    """Set system clipboard to text."""
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    win32clipboard.CloseClipboard()


def find_prompts_dialog():
    """Find the SAP Prompts dialog window handle."""
    results = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t == 'Prompts' or t.startswith('Prompts for'):
                results.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return results[0] if results else None


def get_ok_button(hwnd):
    """Get OK button state and wrapper from Prompts dialog."""
    from pywinauto.uia_element_info import UIAElementInfo
    from pywinauto.controls.uiawrapper import UIAWrapper
    dlg = UIAWrapper(UIAElementInfo(hwnd))
    for c in dlg.descendants():
        if c.element_info.automation_id == 'OkButton':
            return c.is_enabled(), c
    return False, None


def fill_variable(dlg, value, input_x, input_y):
    """
    Fill a variable in the Prompts dialog right panel.

    The right panel inputs are stacked vertically at fixed positions.
    We click at the pre-computed (input_x, input_y), Ctrl+A to select all,
    Ctrl+V to paste. No Enter key (scrolls panel), no left-label click
    (shifts panel). The value is confirmed by the next click or OK.
    """
    real_click(input_x, input_y, pause=0.5)

    clipboard_set(value)
    dlg.type_keys('^a^v', pause=0.05)
    time.sleep(1)
    time.sleep(2)


# ============================================================
# STEP 1: Open Excel
# ============================================================
def step1_open_workbook():
    log("Step 1: Opening workbook...")

    # Already open? Match by workbook path (robust against multiple instances).
    app, wb = find_excel_by_workbook(WORKBOOK_PATH)
    if app is not None:
        log(f"  Already open: {wb.Name}")
        return app

    log(f"  Launching: {os.path.basename(WORKBOOK_PATH)}")
    os.startfile(WORKBOOK_PATH)

    for i in range(40):
        time.sleep(3)
        app, wb = find_excel_by_workbook(WORKBOOK_PATH)
        if app is not None:
            log(f"  Ready ({i*3}s): {wb.Name}")
            return app
    raise RuntimeError("Excel did not open within 120s")


def _get_excel_app():
    """Return a clean, late-bound Excel.Application from the ROT, or None.

    Used as a fallback for COM property access (e.g. COMAddIns) that fails on
    the dynamic Application proxy obtained via find_excel_by_workbook.
    """
    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None


def ensure_ao_addin(xl):
    """Make sure the SAP Analysis for Office COM add-in is loaded.

    When Excel was force-closed during a prior hung run (or the add-in soft-
    failed on startup), the Analysis ribbon does not appear even though the
    add-in's registry LoadBehavior is 3. We reconnect it through the standard
    COMAddIns API (progID 'SapExcelAddIn'), which loads the ribbon WITHOUT any
    XLL/SAPExecuteCommand call, so it does not break the ribbon.

    Returns True if the add-in is connected (or was reconnected), else False.
    """
    # The Application object handed in comes from the Running Object Table as a
    # late-bound dynamic dispatch; reading .COMAddIns on it can raise
    # "<unknown>.COMAddIns". Re-fetch a clean Excel.Application and retry a few
    # times in case Excel is momentarily busy right after opening the workbook.
    addins = None
    for attempt in range(5):
        for candidate in (xl, _get_excel_app()):
            if candidate is None:
                continue
            try:
                addins = candidate.COMAddIns
                break
            except Exception:
                addins = None
        if addins is not None:
            break
        time.sleep(2)

    if addins is None:
        log("  [addin] Cannot read COMAddIns (Excel busy?); skipping auto-load")
        return False

    sap = None
    for a in addins:
        try:
            pid = (a.progID or "")
        except Exception:
            continue
        if "Sap" in pid or "Analysis" in pid:
            sap = a
            break

    if sap is None:
        log("  [addin] SAP Analysis add-in not registered in this Excel")
        return False

    try:
        if sap.Connect:
            return True
        log("  [addin] Analysis add-in not loaded -> connecting...")
        sap.Connect = True
        # Give Excel a moment to build the ribbon after connecting.
        for _ in range(10):
            time.sleep(1)
            if sap.Connect:
                log("  [addin] Analysis add-in connected")
                return True
        log("  [addin] Connect requested but add-in still not active")
        return False
    except Exception as e:
        log(f"  [addin] Failed to connect Analysis add-in: {e}")
        return False


# ============================================================
# STEP 2: Wait for AO ribbon
# ============================================================
def step2_wait_ao(xl):
    log("Step 2: Waiting for AO Analysis ribbon...")
    from pywinauto import Application

    # Make sure the SAP add-in is actually loaded before waiting for its ribbon.
    ensure_ao_addin(xl)

    # Get Excel window handle via win32gui (avoid COM for hwnd lookup)
    excel_hwnd = None
    def _find_excel(h, _):
        nonlocal excel_hwnd
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if 'Excel' in title:
                excel_hwnd = h
        return True
    win32gui.EnumWindows(_find_excel, None)

    if not excel_hwnd:
        raise RuntimeError("Cannot find Excel window handle")

    app = Application(backend='uia').connect(handle=excel_hwnd)
    main_win = app.window(handle=excel_hwnd)

    # Wait for Analysis tab to appear in ribbon
    for attempt in range(60):
        for d in main_win.descendants():
            if d.element_info.control_type == 'TabItem':
                name = d.element_info.name or ''
                if 'Analysis' == name:
                    log(f"  Analysis ribbon found ({attempt*3}s)")
                    return main_win
        if attempt % 10 == 0 and attempt > 0:
            log(f"  Still waiting for Analysis ribbon... ({attempt*3}s)")
            # Ribbon still missing after a while -> try reconnecting the add-in.
            ensure_ao_addin(xl)
        time.sleep(3)

    raise RuntimeError("Analysis ribbon did not appear within 180s")


# ============================================================
# STEP 3: Click Refresh All -> triggers Prompts dialog
# ============================================================
def step3_refresh(xl, main_win):
    log("Step 3: Clicking Refresh All (triggers Prompts)...")

    # Click the Analysis tab first
    analysis_tab = main_win.child_window(title='Analysis', control_type='TabItem')
    analysis_tab.click_input()
    time.sleep(1)

    # Find and click "Refresh All" button
    # The Refresh All SplitButton has a Button child — click the button part
    refresh_btn = None
    for d in main_win.descendants():
        if d.element_info.control_type == 'Button':
            name = d.element_info.name or ''
            if name == 'Refresh All':
                refresh_btn = d
                break

    if not refresh_btn:
        raise RuntimeError("Refresh All button not found in Analysis ribbon")

    log("  Clicking Refresh All...")
    refresh_btn.click_input()

    # Wait for Prompts dialog to appear
    log("  Waiting for Prompts dialog...")
    for w in range(60):
        time.sleep(2)
        hwnd = find_prompts_dialog()
        if hwnd:
            log(f"  Prompts dialog appeared: hwnd={hwnd}")
            return hwnd

    raise RuntimeError("Prompts dialog did not appear within 120s")


# ============================================================
# STEP 4: Fill Prompts dialog
# ============================================================
# Fill mandatory variables with * in the left panel:
#   - * Region - MR  (mandatory, fill with REGION)
# Always fill Category to reduce data transfer from SAP.
# Calendar Week is retained by SAP automatically — no need to fill.
# ============================================================
def step4_fill_prompts(hwnd):
    log("Step 4: Filling Prompts dialog...")
    from pywinauto import Application

    # Press Alt to allow SetForegroundWindow from background process
    ctypes.windll.user32.keybd_event(0x12, 0, 0, 0)       # Alt down
    ctypes.windll.user32.keybd_event(0x12, 0, 0x0002, 0)  # Alt up
    time.sleep(0.2)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # Fallback: ShowWindow + BringWindowToTop
        win32gui.ShowWindow(hwnd, 5)  # SW_SHOW
        win32gui.BringWindowToTop(hwnd)
    time.sleep(1)

    app = Application(backend='uia').connect(handle=hwnd)
    dlg = app.window(handle=hwnd)

    # Right panel positioning:
    # Right panel Group DP_4 does NOT scroll when clicking left panel labels.
    # Inputs are stacked vertically at: Y = rr.top + 62 + index * 40
    # X = center of right pane.
    # IMPORTANT: Do NOT press Enter after pasting — Enter scrolls the panel.
    FIRST_OFFSET = 62
    SPACING = 40

    right_pane = dlg.child_window(auto_id='VariableItemsScroll', control_type='Pane')
    rr = right_pane.rectangle()
    input_x = (rr.left + rr.right) // 2
    log(f"  Right panel: ({rr.left},{rr.top})-({rr.right},{rr.bottom})")

    # Get variable order from left panel labels (sorted by Y position)
    var_entries = []
    for d in dlg.descendants():
        if d.element_info.control_type == 'Text':
            name = d.element_info.name or ''
            rect = d.element_info.rectangle
            if ': ' in name and rect.top > rr.top - 100:
                var_entries.append((name.strip(), rect.top))
    var_entries.sort(key=lambda x: x[1])
    log(f"  Variables ({len(var_entries)}): {[n for n, _ in var_entries]}")

    def get_idx(keyword):
        return next((i for i, (n, _) in enumerate(var_entries) if keyword in n), None)

    def get_y(idx):
        return rr.top + FIRST_OFFSET + idx * SPACING

    # Fill Region if mandatory
    region_idx = get_idx('Region')
    if region_idx is not None and var_entries[region_idx][0].startswith('*'):
        region_y = get_y(region_idx)
        log(f"  Filling Region = {REGION} at ({input_x},{region_y}) [idx={region_idx}]")
        fill_variable(dlg, REGION, input_x, region_y)
        time.sleep(1)

    # Always fill Category
    category_idx = get_idx('Category')
    if category_idx is not None:
        category_y = get_y(category_idx)
        log(f"  Filling Category = {CATEGORY} at ({input_x},{category_y}) [idx={category_idx}]")
        fill_variable(dlg, CATEGORY, input_x, category_y)
        time.sleep(1)

    # Calendar Week: skip — SAP retains the previous value automatically.
    # Only Region and Category need to be filled each run.

    # Final OK
    ok, ok_btn = get_ok_button(hwnd)
    log(f"  All variables set, OK enabled: {ok}")
    if not ok:
        # If OK is still disabled, the right panel input might not have received
        # the value. Log variable summary for debugging.
        log("  Checking variable states...")
        for d in dlg.descendants():
            if d.element_info.control_type == 'Text':
                name = d.element_info.name or ''
                if name and ('Region' in name or 'Category' in name or 'Calendar' in name):
                    log(f"    {name}")
        raise RuntimeError("Variables not accepted (OK still disabled)")

    log("  Clicking OK...")
    ok_rect = ok_btn.element_info.rectangle
    ok_cx = (ok_rect.left + ok_rect.right) // 2
    ok_cy = (ok_rect.top + ok_rect.bottom) // 2
    real_click(ok_cx, ok_cy, pause=1.0)
    time.sleep(3)

    # Retry OK if dialog still open (may need a second click)
    for retry in range(3):
        if not find_prompts_dialog():
            break
        log(f"  Dialog still open, retrying OK ({retry+1})...")
        ok, ok_btn = get_ok_button(hwnd)
        if ok and ok_btn:
            ok_rect = ok_btn.element_info.rectangle
            ok_cx = (ok_rect.left + ok_rect.right) // 2
            ok_cy = (ok_rect.top + ok_rect.bottom) // 2
            real_click(ok_cx, ok_cy, pause=1.0)
            time.sleep(3)

    if find_prompts_dialog():
        raise RuntimeError("Prompts dialog did not close after clicking OK")

    log("  Prompts submitted successfully!")


# ============================================================
# STEP 5: Wait for BW data
# ============================================================
def step5_wait_data(xl, baseline=None):
    log("Step 5: Waiting for BW data refresh...")
    log(f"  Pre-refresh baseline: {baseline}")

    # Re-acquire COM object — the background Refresh thread may have
    # disrupted the original COM proxy in this apartment.
    ws = None
    for retry in range(5):
        try:
            app, wb = find_excel_by_workbook(WORKBOOK_PATH)
            if app is not None:
                xl = app
                ws = wb.Sheets("Main")
                break
        except Exception:
            pass
        log(f"  COM reconnect attempt {retry+1}...")
        time.sleep(3)
    if ws is None:
        raise RuntimeError("Cannot re-acquire Excel COM object")

    changed = False
    last_sig = None
    stable_count = 0
    for w in range(120):
        time.sleep(3)
        sig = data_signature(ws)
        if sig is None:
            # COM object may be busy/stale during refresh; reconnect.
            try:
                app, wb = find_excel_by_workbook(WORKBOOK_PATH)
                if app is not None:
                    xl = app
                    ws = wb.Sheets("Main")
            except Exception:
                pass
            continue

        lr = sig[0]
        # Has the data actually changed from the stale pre-refresh baseline?
        if baseline is None or sig != baseline:
            changed = True

        if changed and lr > 100:
            # Wait until the signature stops changing (refresh finished writing).
            if sig == last_sig:
                stable_count += 1
            else:
                stable_count = 0
            if stable_count >= 2:  # stable across ~2 consecutive polls (~6s)
                log(f"  Data refreshed & stable: {lr} rows, sig={sig} ({w*3}s)")
                return lr

        last_sig = sig
        if w % 10 == 0 and w > 0:
            state = "changed" if changed else "UNCHANGED (still stale)"
            log(f"  Still refreshing... {state}, sig={sig} ({w*3}s)")

    if not changed:
        raise RuntimeError(
            "Data never changed from the pre-refresh baseline within 360s — "
            "BW refresh likely failed; refusing to export STALE data.")
    raise RuntimeError("Data did not stabilize within 360s")


# ============================================================
# STEP 6: Fast export to CSV
# ============================================================
def step6_export(xl, last_row):
    log("Step 6: Exporting to CSV...")
    ws = xl.ActiveWorkbook.Sheets("Main")

    # Read headers (row 21)
    headers = [str(v) if v else f"Col{i+1}" for i, v in enumerate(
        ws.Range(ws.Cells(21, 1), ws.Cells(21, 10)).Value[0])]

    # Read data in chunks (fast bulk read via Range().Value)
    all_data = []
    chunk = 2000
    r = 22
    t0 = time.time()
    while r <= last_row:
        end_r = min(r + chunk - 1, last_row)
        data = ws.Range(ws.Cells(r, 1), ws.Cells(end_r, 10)).Value
        if isinstance(data[0], tuple):
            all_data.extend(data)
        else:
            all_data.append(data)
        r = end_r + 1
    log(f"  Read {len(all_data)} rows in {time.time()-t0:.1f}s")

    # Write CSV
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = time.strftime("%Y%m%d_%H%M%S")
    csv_path = os.path.join(OUTPUT_DIR, f"DFV_{ts}.csv")

    with open(csv_path, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        w.writerow(headers)
        for row in all_data:
            w.writerow([v if v is not None else '' for v in row])

    log(f"  Exported: {os.path.basename(csv_path)}")
    return csv_path


# ============================================================
# STEP 7: Pipeline (filter + classify + report)
# ============================================================
def step7_pipeline(csv_path):
    log("Step 7: Running pipeline...")
    from pipeline import run_pipeline

    df, errors, summary, reports = run_pipeline(csv_path)
    return df, errors, summary


# ============================================================
# MAIN
# ============================================================
def main():
    t0 = time.time()

    xl = step1_open_workbook()
    main_win = step2_wait_ao(xl)

    # Snapshot the (stale) data currently in the sheet BEFORE refreshing, so
    # step5 can detect when the BW refresh has actually replaced it.
    _app, _wb = find_excel_by_workbook(WORKBOOK_PATH)
    baseline = data_signature(_wb.Sheets("Main")) if _wb is not None else None
    log(f"Baseline data signature: {baseline}")

    hwnd = step3_refresh(xl, main_win)
    step4_fill_prompts(hwnd)

    last_row = step5_wait_data(xl, baseline)
    # Re-acquire COM after data load (match by workbook path)
    app, _wb = find_excel_by_workbook(WORKBOOK_PATH)
    if app is not None:
        xl = app
    csv_path = step6_export(xl, last_row)
    df, errors, summary = step7_pipeline(csv_path)

    elapsed = time.time() - t0
    log(f"\n{'='*50}")
    log(f"DONE in {elapsed:.0f}s")
    log(f"  Data: {len(df)} rows")
    log(f"  Errors: {len(errors)} ({summary['actionable_errors']} actionable)")
    log(f"  Vol: {summary['vol_status'].upper()} ({summary['diff_pct']:.2f}%)  |  SKU: {summary['sku_status'].upper()} ({summary['total_errors']} errors)")
    log(f"{'='*50}")


if __name__ == "__main__":
    main()
