#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_orders.py  -  KPI Dashboard Version 11.17  "Rolling Orders"   (Tableau report: Number of Order with unique "Cus+Adress+Timeslot")

Runs EVERY HOUR, non-stop (00:00 ... 23:00). Each run reads the orders created TODAY (T+0), counted per Expected Delivery Date x Timeslot.
  1. Logs in to Tableau with TABLEAU_USER / TABLEAU_PASS (read from the .env file in the log-kpi-tracker folder, see below), opens the view
         https://inhouse-analytics.hktv.com.hk/#/views/NumberofOrderwithuniqueCusAdressTimeslot/data?:iid=2
     (v11.17: the sheet "data" is ALREADY selected in the Crosstab dialog, so the script never clicks a sheet thumbnail - a click would
     de-select it):  Download -> Crosstab -> CSV -> Download. The file is saved to "Folder for Rolling Orders" under the original name and
     overwritten every run:  AI Fetching for Order Date T-1 (MMM-DD of T-1).csv
     (UTF-16 tab separated). It is deleted again once the numbers are in the json (--keep-files keeps it).
  2. Report layout (v11.17: one long table, a header row + one row per delivery date x timeslot x district x order date):
         A delivery_date | B delivery_timeslot (AM / EV / PM) | C Delivery Address Type | D district (group) | E order creation date (2026年11月10日)
         | F # of Consignee Address | G # of Order | H # of Parent Order | I Same pk-address-date-timeslot Combination
     - Used: A, B, D, E and H (the count = # of Parent Order). The header names of A, B, D, E, H are checked; a different layout stops the run.
     - Column D district (group) = the 10 district codes; "NT-YT & WTH" is the district WTH. Every other code (pick-up points, Macau,
       overseas ...) is skipped.
     - Column E Order Date: only the rows whose Order Date = T+0 (today) are read. e.g. on Nov-10 -> every row with 2026年11月10日 in column E.
     - Per delivery date: Total = AM + EV + PM. EV (and EV2 if it ever appears) are stored as "EV&EV2".
  3. Update rule
       - every hourly run updates the Order Date T+0 (today);
       - the 00:00 run updates TWO order dates in the same session: yesterday (its final count, marked complete) and today
         (e.g. 00:00 on Nov-11 -> Nov-10 and Nov-11). If the 00:00 run was missed, the next run finishes yesterday as well.
       - delivery dates that the report no longer shows (e.g. yesterday's same-day column after midnight) keep their last hourly value.
  4. Writes rolling_orders.json next to index.html AND embeds it in index.html's window.__EMBEDDED_DATA__ block, then pushes
     rolling_orders.json + index.html to the dashboard repo (git pull --rebase -> commit -> push, with the automatic fix for the
     EMBEDDED_DATA conflict, see push_dashboard). Disable with --skip-push or ROLLING_DEPLOY=none.

JSON (days[orderDate]):  {"v":2, "slot":"HH:00", "asOf":..., "complete":bool, "cells":{deliveryDate:{district|"OVERALL":[total, AM, PM, EV&EV2]}}}
Days written before version 11.15 (no "v") keep their old [orders, waybills] cells; the dashboard reads both.

Usage (WSL):
    python3 rolling_orders.py                              # the hourly cron call
    python3 rolling_orders.py --date 2026-11-10            # run as if today were Nov-10 (T+0)
    python3 rolling_orders.py --both                       # force the two-day (yesterday + today) update
    python3 rolling_orders.py --order-date 2026-11-09 --tableau-file "10_Districts.csv" --process-only   # read one named order date from a file
    python3 rolling_orders.py --skip-push
    python3 rolling_orders.py --tableau-only               # only download the csv
    python3 rolling_orders.py --push-only                  # embed + push the json that is already there

Cron (WSL, machine timezone = HKT) - ONE line, every hour, all day. Replace the old 06:00 line:
    0 * * * *  cd "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker" && python3 rolling_orders.py >> logs/rolling_hourly.log 2>&1

Credentials (never hard-code them): TABLEAU_USER / TABLEAU_PASS in  C:\\Users\\chipanl\\Downloads\\Whatsapp Session\\log-kpi-tracker\\.env
(= KPI_DIR/.env, KEY=VALUE lines). Real environment variables win over the file. Needs: pip install playwright && playwright install chromium.
"""
import argparse
import contextlib
import csv
import fcntl
import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

# WSL paths (C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\...)
KPI_DIR = os.environ.get("KPI_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker")


# ----------------------------------------------------------------------------- config
def _load_env_file():
    """KPI_DIR/.env first (the file named in the task), then next to the script / one folder up. Never overrides a real environment variable."""
    seen = set()
    cands = [os.path.join(KPI_DIR, ".env")]
    for folder in (HERE, os.path.dirname(HERE)):
        for name in (".env", "rolling_orders.env"):
            cands.append(os.path.join(folder, name))
    for p in cands:
        rp = os.path.realpath(p)
        if rp in seen or not os.path.exists(p):
            continue
        seen.add(rp)
        for line in open(p, encoding="utf-8-sig"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env_file()

# --- Tableau: report "Number of Order with unique "Cus+Adress+Timeslot"" - login = TABLEAU_USER / TABLEAU_PASS of the .env file ---
TABLEAU_SIGNIN_URL = os.environ.get("ROLLING_TABLEAU_SIGNIN_URL", "https://inhouse-analytics.hktv.com.hk/#/signin")
# v11.17 - view changed to the "data" sheet (.../NumberofOrderwithuniqueCusAdressTimeslot/data?:iid=2); env ROLLING_TABLEAU_URL still overrides it
TABLEAU_VIEW_URL = os.environ.get("ROLLING_TABLEAU_URL",
    "https://inhouse-analytics.hktv.com.hk/#/views/NumberofOrderwithuniqueCusAdressTimeslot/data?:iid=2")
TABLEAU_VIEW_TITLE = os.environ.get("ROLLING_TABLEAU_TITLE", "Number of Order with unique")      # only used for a log line / warning
TABLEAU_USER = os.environ.get("TABLEAU_USER", "")
TABLEAU_PASS = os.environ.get("TABLEAU_PASS", "")
TABLEAU_SETTLE_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_SETTLE_WAIT", "15"))        # after the view appeared, let the crosstab finish rendering
TABLEAU_ATTEMPTS = max(1, int(os.environ.get("TABLEAU_DOWNLOAD_ATTEMPTS", "3")))        # fresh browser sessions, as kpi_pipeline.py
TABLEAU_RECOVERY_WAIT_S = int(os.environ.get("TABLEAU_RECOVERY_WAIT_S", "20"))
TABLEAU_LOAD_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_LOAD_WAIT", "180"))           # max wait for the view to appear
TABLEAU_HEADLESS = os.environ.get("ROLLING_TABLEAU_HEADLESS", "1").strip().lower() not in ("0", "false", "no")
KEEP_FILES = os.environ.get("ROLLING_KEEP_FILES", "0").strip().lower() in ("1", "true", "yes", "on")   # 1 = keep the downloaded csv file

# The raw csv lives OUTSIDE log-kpi-tracker, one level up: C:\Users\chipanl\Downloads\Whatsapp Session\Folder for Rolling Orders
ROLLING_DIR = os.environ.get("ROLLING_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/Folder for Rolling Orders")
PUBLIC_DIR = os.environ.get("KPI_PUBLIC_DIR", os.path.join(KPI_DIR, "public"))   # same folder kpi_pipeline.py serves from
INDEX_HTML = os.environ.get("KPI_INDEX_HTML", os.path.join(PUBLIC_DIR, "index.html"))
ROLLING_JSON = os.environ.get("ROLLING_JSON", os.path.join(PUBLIC_DIR, "rolling_orders.json"))   # next to index.html
LOG_DIR = os.environ.get("ROLLING_LOG_DIR", os.path.join(KPI_DIR, "logs"))
TABLEAU_FILE_PREFIX = "AI Fetching for Order Date T-1"      # original file name kept on purpose: "<prefix> (MMM-DD of T-1).csv"

KEEP_ORDER_DATES = 400                       # order dates kept in rolling_orders.json (the dashboard shows the latest 3 months day by day)
KEEP_SNAPSHOT_DATES = 14                     # legacy per-slot snapshots (no longer written; old ones are only trimmed)
MAX_LEAD_DAYS = 45                           # a delivery date further ahead than this after the order date is not stored

# auto-push (git). Run from the repo root = KPI_DIR (the folder that holds public/ and history.json).
PUSH_ENABLED = os.environ.get("ROLLING_DEPLOY", "1").strip().lower() not in ("0", "none", "off", "false", "no")
DEPLOY_CWD = os.environ.get("ROLLING_DEPLOY_CWD") or os.environ.get("DEPLOY_CWD") or KPI_DIR
GIT_TIMEOUT = int(os.environ.get("ROLLING_GIT_TIMEOUT", "180"))
PUSH_ATTEMPTS = 3

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]     # the dashboard's 10 districts
DISTRICT_ALIAS = {"NT-YT&WTH": "WTH", "NT-YT": "WTH"}     # column A "NT-YT & WTH" (spaces ignored) = WTH
SLOT_ORDER = ["TOTAL", "AM", "PM", "EV"]                  # position in a stored cell: [total, AM, PM, EV&EV2]
SOURCE_NAME = 'Tableau: Number of Order with unique "Cus+Adress+Timeslot"'


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# ----------------------------------------------------------------------------- dates / names
def mmm_dd(d):
    return d.strftime("%b-%d")                        # Nov-09


def tableau_report_path(order_date):
    """Original name + folder kept: 'AI Fetching for Order Date T-1 (MMM-DD of T-1).csv', same name every run => overwritten."""
    return os.path.join(ROLLING_DIR, f"{TABLEAU_FILE_PREFIX} ({mmm_dd(order_date)}).csv")


def order_dates_for_run(now, t0_arg=None, both=False, catchup=False):
    """-> [(order date 'YYYY-MM-DD', final)] in the order they are written.
    T+0 is always updated. The 00:00 run (hour 0) also finishes yesterday: [yesterday final, today]. `both` forces it, `catchup` = the
    00:00 run was missed (yesterday is still open in the json)."""
    t0 = t0_arg or datetime(now.year, now.month, now.day)
    out = []
    if both or catchup or (t0_arg is None and now.hour == 0):
        out.append(((t0 - timedelta(days=1)).strftime("%Y-%m-%d"), True))
    out.append((t0.strftime("%Y-%m-%d"), False))
    return out


# ----------------------------------------------------------------------------- Tableau download
def _view_ready(page):
    """True when the report is on screen: the viz iframe, the toolbar Download button or the viz canvas exists in the page or in any frame
    (Tableau versions differ in how the view is embedded)."""
    sel = ('iframe[title="Data Visualisation"], iframe[src*="views"], #download, [aria-label="Download"], '
           '[data-tb-test-id*="download"], [data-tb-test-id="viz-viewer-toolbar-button-download"], '
           '.tab-tableArea, .tabCanvas, [id^="tabZoneId"], tableau-viz')
    for fr in [page.main_frame] + list(page.frames):
        try:
            if fr.locator(sel).count():
                return True
        except Exception:
            continue
    return False


def _wait_for_view(page, final_path):
    """Wait (up to TABLEAU_LOAD_WAIT_S) until the report is on screen. Re-issues the view URL every 90s in case the SPA bounced back to the
    explore/home page after sign-in. Every 30s it logs what the page shows; on failure a screenshot + the page text are saved next to the csv."""
    end = time.time() + TABLEAU_LOAD_WAIT_S
    last_goto = last_diag = time.time()
    while time.time() < end:
        if _view_ready(page):
            log("Tableau: view loaded")
            return
        if time.time() - last_diag > 30:
            last_diag = time.time()
            try:
                txt = re.sub(r"\s+", " ", page.inner_text("body", timeout=3000))[:160]
            except Exception:
                txt = "?"
            log(f"Tableau: waiting for the view - {len(page.frames)} frame(s), page text: '{txt}'")
        if time.time() - last_goto > 90:
            log(f"Tableau: view not up yet (at {page.url}) - opening it again")
            try:
                page.goto(TABLEAU_VIEW_URL)
            except Exception:
                pass
            last_goto = time.time()
        time.sleep(3)
    base = os.path.splitext(final_path)[0] + "_view_timeout"
    try:
        page.screenshot(path=base + ".png", full_page=True)
        log(f"Tableau: screenshot saved to {base}.png (page was at {page.url})")
    except Exception:
        pass
    try:
        open(base + ".txt", "w", encoding="utf-8").write(page.inner_text("body", timeout=5000)[:5000])
        log(f"Tableau: page text saved to {base}.txt")
    except Exception:
        pass
    raise RuntimeError(f"Tableau view did not load within {TABLEAU_LOAD_WAIT_S}s (page at {page.url})")


def _check_view_title(page):
    """Log which view was opened; warn when its title does not contain TABLEAU_VIEW_TITLE. Never stops the run."""
    norm = lambda x: re.sub(r"\s+", "", x or "").lower()
    seen = []
    try:
        seen.append(page.title() or "")
    except Exception:
        pass
    try:
        seen.append(page.inner_text("body", timeout=5000)[:20000])
    except Exception:
        pass
    if any(norm(TABLEAU_VIEW_TITLE) in norm(t) for t in seen):
        log(f"Tableau: view title OK ('{TABLEAU_VIEW_TITLE}...' found)")
    else:
        log(f"WARNING: the opened Tableau view does not show '{TABLEAU_VIEW_TITLE}...' (page {page.url}) - check ROLLING_TABLEAU_URL")


def _click_crosstab_item(page):
    """Click the 'Crosstab' entry of the Download menu: the known selectors first, then any visible element whose text contains 'Crosstab'
    (page or any frame, any case)."""
    ct_sel = ["#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
              "#viz-viewer-toolbar-download-menu div:nth-of-type(3) span",
              "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]",
              '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
              "span:has-text('Crosstab')", "text='Crosstab'"]
    if _fast_click(page, ct_sel, 3000):
        return True
    for fr in [page.main_frame] + list(page.frames):               # by role: the menu is Image / Data / Crosstab / PDF / PowerPoint
        try:
            loc = fr.get_by_role("menuitem", name=re.compile(r"crosstab", re.I))
            if loc.count() and loc.first.is_visible():
                loc.first.click(force=True, timeout=3000)
                return True
        except Exception:
            continue
    for fr in [page.main_frame] + list(page.frames):
        try:
            loc = fr.get_by_text(re.compile(r"crosstab", re.I))
            for i in range(min(loc.count(), 6)):
                it = loc.nth(i)
                if it.is_visible():
                    it.click(force=True, timeout=3000)
                    return True
        except Exception:
            continue
    return False


def _menu_text(page):
    """What the Download menu / dialog shows (for the log when Crosstab cannot be found)."""
    seen = []
    for fr in [page.main_frame] + list(page.frames):
        try:
            for t in fr.locator('[role="menuitem"], [role="menu"] li, [role="dialog"] button, #viz-viewer-toolbar-download-menu *').all_inner_texts()[:30]:
                t = re.sub(r"\s+", " ", t).strip()
                if t and t not in seen:
                    seen.append(t[:40])
        except Exception:
            continue
    return seen[:15] or "nothing visible"


def _open_crosstab(page, final_path):
    """Download button -> Crosstab. Up to 3 tries (Escape + click again); screenshots *_crosstab_N.png are saved when Crosstab is not found."""
    base = os.path.splitext(final_path)[0] + "_crosstab"
    dl_sel = ["#download", '[aria-label="Download"]', "button:has-text('Download')"]
    for attempt in range(1, 4):
        if not (_fast_click(page, dl_sel, 4000) or _smart_click(page, dl_sel, 15)):
            raise RuntimeError("Download button not found")
        time.sleep(2)
        if _click_crosstab_item(page):
            return
        log(f"Tableau: Crosstab not found after the Download click (try {attempt}/3) - menu shows: {_menu_text(page)}")
        try:
            page.screenshot(path=f"{base}_{attempt}.png", full_page=True)
        except Exception:
            pass
        if attempt == 3:                                           # last resort: keyboard. The menu opens on "Image" -> Down, Down = "Crosstab"
            try:
                _fast_click(page, dl_sel, 4000)
                time.sleep(1.5)
                page.keyboard.press("ArrowDown"); page.keyboard.press("ArrowDown"); page.keyboard.press("Enter")
                log("Tableau: Crosstab selected with the keyboard (Down, Down, Enter)")
                return
            except Exception:
                pass
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        time.sleep(2)
    raise RuntimeError("Crosstab option not found")


SHEET_NAME = os.environ.get("ROLLING_TABLEAU_SHEET", "data")        # v11.17: the view opens on the sheet "data", which is ALREADY selected in the Crosstab dialog


def _norm_title(s):
    return " ".join(str(s or "").split()).lower()


def _sheet_selected_state(page):
    """True / False = the thumbnail of SHEET_NAME in the Crosstab dialog is / is not selected (aria-selected, as kpi_pipeline.is_sheet_already_selected);
    None = the thumbnail could not be found at all. Exact title first, then a case / whitespace tolerant match."""
    viz = page.frame_locator('iframe[title="Data Visualisation"]')
    targets = [viz, page.main_frame] + list(page.frames)
    sels = [f'[role="option"][title="{SHEET_NAME}"]',
            f'[data-tb-test-id^="sheet-thumbnail"][title="{SHEET_NAME}"]',
            f'div[title="{SHEET_NAME}"][aria-selected]']
    for t in targets:
        for sel in sels:
            try:
                loc = t.locator(sel).first
                if loc.count() == 0:
                    continue
                st = loc.get_attribute("aria-selected")
                if st is not None:
                    return st == "true"
            except Exception:
                continue
    want = _norm_title(SHEET_NAME)
    for t in targets:
        try:
            opts = t.locator('[role="option"]')
            for i in range(min(opts.count(), 40)):
                o = opts.nth(i)
                if _norm_title(o.get_attribute("title")) == want:
                    st = o.get_attribute("aria-selected")
                    if st is not None:
                        return st == "true"
        except Exception:
            continue
    return None


def _log_thumbnails(page):
    """Debug line: which sheet thumbnails the Crosstab dialog lists and which one is selected."""
    seen = []
    for t in [page.frame_locator('iframe[title="Data Visualisation"]'), page.main_frame] + list(page.frames):
        try:
            opts = t.locator('[role="option"]')
            for i in range(min(opts.count(), 40)):
                o = opts.nth(i)
                title = (o.get_attribute("title") or "").strip()
                if title:
                    item = f"{title}{' [selected]' if o.get_attribute('aria-selected') == 'true' else ''}"
                    if item not in seen:
                        seen.append(item)
        except Exception:
            continue
    log(f"Tableau: sheets in the Crosstab dialog: {seen if seen else 'none detected'}")


def _select_sheet(page, final_path):
    """v11.17: the sheet "data" is already selected (pre-clicked) when the Crosstab dialog opens. The thumbnail list is a TOGGLE - a click on
    a selected sheet DE-selects it - so this step never clicks: it only logs the thumbnails and the aria-selected state as a diagnostic.
    (ROLLING_TABLEAU_CLICK_SHEET=1 brings back the old behaviour: click the thumbnail only when it is not selected.)"""
    _log_thumbnails(page)
    state = _sheet_selected_state(page)
    if state is True:
        log(f"Tableau: sheet '{SHEET_NAME}' is already selected - not clicked (a click would de-select it)")
        return
    if os.environ.get("ROLLING_TABLEAU_CLICK_SHEET", "0").strip().lower() not in ("1", "true", "yes", "on"):
        log(f"Tableau: sheet '{SHEET_NAME}' is expected to be pre-selected - not clicked (selected state read as {state}); continuing")
        return
    base = os.path.splitext(final_path)[0] + "_sheet"
    sel = [f'[role="option"][title="{SHEET_NAME}"]',
           f'[data-tb-test-id^="sheet-thumbnail"][title="{SHEET_NAME}"]',
           f'div[title="{SHEET_NAME}"][aria-selected]',
           f"[role='option']:has-text('{SHEET_NAME}')"]
    if not (_fast_click(page, sel, 3000) or _smart_click_with_scroll(page, sel, 15)):
        try:
            page.screenshot(path=base + ".png", full_page=True)
        except Exception:
            pass
        raise RuntimeError(f"sheet thumbnail '{SHEET_NAME}' not found in the Crosstab dialog (screenshot {base}.png)")
    time.sleep(1.5)
    log(f"Tableau: clicked sheet '{SHEET_NAME}' -> selected state = {_sheet_selected_state(page)}")


def _open_view(page, final_path):
    """Same sequence as kpi_pipeline._attempt_download: open the view, then a FORCED RELOAD so the toolbar (Download button) is initialised
    cleanly (a hash-only navigation leaves the toolbar without event handlers), then guard against being thrown back to the /#/user/ page."""
    def load():
        page.goto(TABLEAU_VIEW_URL)
        try:
            page.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            pass                                   # Tableau polls in the background: networkidle often never happens
        time.sleep(10)
        page.wait_for_timeout(5000)
        log("Tableau: forced reload so the toolbar initialises correctly")
        try:
            page.reload(wait_until="load", timeout=45000)
        except Exception as ex:
            log(f"Tableau: reload() wait ended ({ex}) - continuing")
        time.sleep(8)
        page.wait_for_timeout(3000)
    load()
    retry = 0
    while "/#/user/" in page.url and retry < 2:
        retry += 1
        log(f"Tableau: sent to the settings page ({page.url}) - opening the view again ({retry}/2)")
        load()
    if "/#/user/" in page.url:
        raise RuntimeError(f"still on the settings page after {retry} retries ({page.url})")
    _wait_for_view(page, final_path)


def _format_states(page):
    """[(label text, checked)] of the file-format options (CSV / Excel ...) in the Crosstab dialog."""
    out = []
    for root in [page.main_frame] + list(page.frames):
        try:
            labs = root.locator('#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label, [role="dialog"] label')
            for i in range(min(labs.count(), 10)):
                l = labs.nth(i)
                inp = l.locator("input")
                txt = re.sub(r"\s+", " ", l.inner_text(timeout=1000)).strip()
                if txt and inp.count():
                    out.append((txt, inp.first.is_checked()))
        except Exception:
            continue
    return out


def _select_csv(page):
    """Choose the CSV file format and VERIFY it (the dialog may default to Excel; an unnoticed miss gives an .xlsx that is then read as text)."""
    for attempt in range(1, 4):
        st = _format_states(page)
        log(f"Tableau: file formats in the dialog (try {attempt}): {st if st else 'none detected'}")
        if any(t.upper().startswith("CSV") and c for t, c in st):
            return True
        done = False
        for root in [page.main_frame] + list(page.frames):
            try:
                lab = root.locator('[role="dialog"] label, #export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label').filter(has_text=re.compile(r"^\s*CSV\s*$", re.I)).first
                if lab.count():
                    lab.click(force=True, timeout=3000)
                    done = True
                    break
            except Exception:
                continue
        if not done:
            csv_sel = ["#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
                       "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2)",
                       "label:has-text('CSV')", "text='CSV'"]
            if not _fast_click(page, csv_sel, 2000):
                _smart_click(page, csv_sel, 5)
        time.sleep(1)
    st = _format_states(page)
    log(f"Tableau: file formats after clicking CSV: {st}")
    if not st:                                   # markup without readable radio inputs: cannot verify, go on (the file type is checked after the download)
        return True
    return any(t.upper().startswith("CSV") and c for t, c in st)


def _dialog_roots(page):
    """Where the Crosstab dialog may live: the main page or the viz iframe (and any other frame)."""
    return [page.main_frame, page.frame_locator('iframe[title="Data Visualisation"]')] + list(page.frames)


def _log_dialog_buttons(page):
    """Log the visible buttons of the Crosstab dialog (name | aria-label | test id | enabled) so a wrong click can be diagnosed from the log."""
    rows = []
    for root in [page.main_frame] + list(page.frames):
        try:
            btns = root.locator('[role="dialog"] button, #export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id button')
            for i in range(min(btns.count(), 12)):
                b = btns.nth(i)
                if not b.is_visible():
                    continue
                txt = re.sub(r"\s+", " ", b.inner_text(timeout=1000)).strip()[:25]
                rows.append("[%s|%s|%s|%s]" % (txt, b.get_attribute("aria-label"), b.get_attribute("data-tb-test-id"), "on" if b.is_enabled() else "OFF"))
        except Exception:
            continue
    log("Tableau: buttons in the Crosstab dialog: " + (" ".join(rows) if rows else "none found"))


def _press_download(page, final_path):
    """Press the dialog's own Download button and return the Download object. The old code clicked the FIRST match of
    '#export-crosstab-...-Body-Id button' - that can be another button of the dialog (e.g. a sheet-list scroll / info button), which counts as a click
    but starts no download, so the run then waited 180s for nothing. Now the specific Download button is tried first, the generic ones last
    (taking the LAST button = the footer action), and every candidate gets only 25s to start a download before the next one is tried."""
    dlg = '#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id'
    cands = [
        ("test-id export button", lambda r: r.locator('button[data-tb-test-id="export-crosstab-export-Button"]').first),
        ("aria-label 'Download Crosstab'", lambda r: r.locator("button[aria-label='Download Crosstab']").first),
        ("dialog button named Download", lambda r: r.locator('[role="dialog"]').get_by_role("button", name=re.compile(r"^\s*download", re.I)).last),
        ("aria-label 'Download' in dialog", lambda r: r.locator("[role='dialog'] button[aria-label='Download']").last),
        ("last button of the dialog", lambda r: r.locator('[role="dialog"] button').last),
        ("last button of the dialog body", lambda r: r.locator(dlg + ' button').last),
        ("button with text Download", lambda r: r.locator("button:has-text('Download')").last),
    ]
    base = os.path.splitext(final_path)[0] + "_dlbutton"
    for label, get in cands:
        for root in _dialog_roots(page):
            try:
                btn = get(root)
                if btn.count() == 0 or not btn.is_visible():
                    continue
                log(f"Tableau: pressing the dialog Download button via '{label}'")
                with page.expect_download(timeout=25000) as info:
                    try:
                        btn.click(timeout=5000)
                    except Exception:
                        btn.click(force=True, timeout=5000)
                return info.value
            except Exception as ex:
                log(f"Tableau: '{label}' did not start a download ({str(ex).splitlines()[0][:110]}) - trying the next one")
                try:
                    page.screenshot(path=base + ".png", full_page=True)
                except Exception:
                    pass
    raise RuntimeError(f"no Download button of the Crosstab dialog started a download (screenshot {base}.png)")


def _tableau_session(final_path):
    from playwright.sync_api import sync_playwright
    tmp = final_path + ".part"
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=TABLEAU_HEADLESS, args=["--disable-popup-blocking"])
        try:
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            page.set_viewport_size({"width": 1920, "height": 1080})     # wide: a narrow window hides the Download icon
            log("Tableau: login (TABLEAU_USER)")
            page.goto(TABLEAU_SIGNIN_URL)
            page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
            page.locator("input[type='text'], input[name='username']").first.fill(TABLEAU_USER)
            page.locator("input[type='password'], input[name='password']").first.fill(TABLEAU_PASS)
            page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()
            try:                                                        # let Tableau finish its own post-login redirect BEFORE we navigate
                page.wait_for_load_state("networkidle", timeout=60000)
            except Exception:
                pass
            try:
                page.wait_for_url(re.compile(r".*/#/(user|explore)\b.*"), timeout=30000)
                log(f"Tableau: signed in, landed on {page.url}")
            except Exception:
                log(f"Tableau: no /#/user|explore redirect detected (at {page.url}) - continuing")
            time.sleep(3)
            log("Tableau: opening 'Number of Order with unique Cus+Adress+Timeslot'")
            _open_view(page, final_path)
            time.sleep(TABLEAU_SETTLE_WAIT_S)
            _check_view_title(page)
            _open_crosstab(page, final_path)                            # Download -> Crosstab
            time.sleep(2)
            _select_sheet(page, final_path)                             # sheet "data" is pre-selected: only checked, never clicked
            time.sleep(1)
            if not _select_csv(page):
                raise RuntimeError("could not select the CSV file format in the Crosstab dialog")
            time.sleep(1)
            _log_dialog_buttons(page)
            dl = _press_download(page, final_path)                      # the final "Download" button of the Crosstab dialog
            os.makedirs(os.path.dirname(final_path), exist_ok=True)
            dl.save_as(tmp)
            log(f"Tableau: downloaded '{dl.suggested_filename}' ({os.path.getsize(tmp)} bytes)")
        except Exception:
            try:                                                        # picture of the failure for later checking
                page.screenshot(path=os.path.splitext(final_path)[0] + "_download_failed.png", full_page=True)
            except Exception:
                pass
            raise
        finally:
            browser.close()
    return tmp


def download_tableau_report(final_path):
    """Fresh browser session per attempt (up to TABLEAU_ATTEMPTS), as the GMV download does. Saved as .part, checked, then renamed over the old copy."""
    if not TABLEAU_USER or not TABLEAU_PASS:
        raise RuntimeError(f"TABLEAU_USER / TABLEAU_PASS not set (put them in {os.path.join(KPI_DIR, '.env')})")
    last = None
    for n in range(1, TABLEAU_ATTEMPTS + 1):
        if n > 1:
            log(f"Tableau: retry {n}/{TABLEAU_ATTEMPTS} (fresh browser + login)")
            time.sleep(TABLEAU_RECOVERY_WAIT_S)
        try:
            tmp = _tableau_session(final_path)
            if os.path.getsize(tmp) < 200:
                raise RuntimeError(f"downloaded file is only {os.path.getsize(tmp)} bytes")
            try:
                parse_report(tmp)                              # layout check before it replaces the old copy
            except Exception:
                bad = final_path + ".bad"                      # keep what Tableau really exported, for checking which sheet / layout it is
                try:
                    os.replace(tmp, bad)
                    log(f"Tableau: the downloaded file does not have the expected layout - kept as {bad} for checking")
                except OSError:
                    pass
                raise
            os.replace(tmp, final_path)                        # same name each run -> the previous copy is covered
            log(f"Tableau: saved {final_path} ({os.path.getsize(final_path) / 1048576:.2f} MB)")
            return final_path
        except Exception as ex:
            last = ex
            log(f"Tableau: attempt {n}/{TABLEAU_ATTEMPTS} failed: {type(ex).__name__}: {ex}")
    raise RuntimeError(f"Tableau 'Number of Order with unique Cus+Adress+Timeslot' download failed after {TABLEAU_ATTEMPTS} sessions: {last}")



# ----------------------------------------------------------------------------- shared helpers

def _frames(page):
    return [page.frame_locator('iframe[title="Data Visualisation"]'), page.main_frame] + list(page.frames)


def _fast_click(page, selectors, timeout_ms=5000):
    per = max(300, timeout_ms // max(1, len(selectors)))
    end = time.time() + timeout_ms / 1000.0
    while time.time() < end:
        for target in _frames(page):
            for sel in selectors:
                try:
                    target.locator(sel).first.click(timeout=per)
                    return True
                except Exception:
                    continue
    return False


def _smart_click(page, selectors, timeout_sec=15):
    end = time.time() + timeout_sec
    while time.time() < end:
        for target in _frames(page):
            for sel in selectors:
                try:
                    loc = target.locator(sel).first
                    if loc.is_visible():
                        loc.click(force=True)
                        return True
                except Exception:
                    continue
        time.sleep(1)
    return False


def _smart_click_with_scroll(page, selectors, timeout_sec=15):
    """As kpi_pipeline.smart_click_with_scroll: scroll the target into view first (the sheet thumbnail list scrolls; a force click does not scroll)."""
    end = time.time() + timeout_sec
    while time.time() < end:
        for target in _frames(page):
            for sel in selectors:
                try:
                    loc = target.locator(sel).first
                    loc.scroll_into_view_if_needed(timeout=3000)
                    if loc.is_visible():
                        loc.click(force=True)
                        return True
                except Exception:
                    continue
        time.sleep(1)
    return False


def cleanup_files(paths, why="written to the json"):
    """v11.9 - delete the downloaded csv files once their numbers are safely in rolling_orders.json. Only files inside
    "Folder for Rolling Orders" are touched (+ a leftover .part / .new next to them); never raises."""
    base = os.path.realpath(ROLLING_DIR) + os.sep
    for p in paths:
        if not p:
            continue
        for q in (p, p + ".part", p + ".new"):
            try:
                if os.path.isfile(q) and os.path.realpath(q).startswith(base):
                    sz = os.path.getsize(q) / 1048576
                    os.remove(q)
                    log(f"cleanup: deleted {os.path.basename(q)} ({sz:.1f} MB) - {why}")
            except Exception as ex:
                log(f"cleanup: could not delete {q}: {ex}")


# ----------------------------------------------------------------------------- report parsing
def _read_text(path):
    """Tableau crosstab text. UTF-16 (BOM, the normal 'CSV' crosstab download) or UTF-8; tab or comma separated."""
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return raw.decode("utf-16", errors="replace")
    return raw.decode("utf-8-sig", errors="replace")


def _date_format(values):
    """'iso' | 'mdy' | 'dmy' from the date strings of the file (Tableau US export = M/D/YYYY; a first part > 12 proves D/M/YYYY)."""
    vals = [v for v in values if v]
    if vals and all(re.match(r"^\d{4}-\d{2}-\d{2}", v) for v in vals):
        return "iso"
    for v in vals:
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", v)
        if m:
            if int(m.group(1)) > 12:
                return "dmy"
            if int(m.group(2)) > 12:
                return "mdy"
    return "mdy"


_EXTRA_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%b %d %Y", "%B %d %Y",
                       "%A, %B %d, %Y", "%a, %b %d, %Y", "%d-%b-%Y", "%d-%b-%y", "%b-%d-%Y", "%m/%d/%y", "%d/%m/%y")


def _to_iso(s, fmt):
    """Date text -> 'YYYY-MM-DD' or None. The file's own format (iso / mdy / dmy) first, then other common Tableau spellings
    (2026年11月10日, Nov 10, 2026, 10 Nov 2026, Tuesday, November 10, 2026, ...). A time part after the date is ignored."""
    s = (s or "").strip()
    if not s:
        return None
    try:
        if fmt == "iso":
            return datetime.strptime(s[:10], "%Y-%m-%d").strftime("%Y-%m-%d")
        return datetime.strptime(s, "%d/%m/%Y" if fmt == "dmy" else "%m/%d/%Y").strftime("%Y-%m-%d")
    except ValueError:
        pass
    m = re.match(r"^(\d{4})\s*[年]\s*(\d{1,2})\s*[月]\s*(\d{1,2})", s)            # 2026年11月10日
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).strftime("%Y-%m-%d")
        except ValueError:
            return None
    s2 = re.sub(r"\s+\d{1,2}:\d{2}(:\d{2})?\s*([AP]M)?$", "", s, flags=re.I)         # drop a trailing time
    for f in (("%m/%d/%Y", "%d/%m/%Y") if fmt != "dmy" else ("%d/%m/%Y", "%m/%d/%Y")) + _EXTRA_DATE_FORMATS:
        try:
            return datetime.strptime(s2, f).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def _num(s):
    s = (s or "").strip().replace(",", "")
    if not s:
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def _slot_key(label):
    t = re.sub(r"\s+", "", (label or "")).upper()
    if t == "TOTAL":
        return "TOTAL"
    if t == "AM":
        return "AM"
    if t == "PM":
        return "PM"
    if t.startswith("EV"):                     # EV, EV2, EV&EV2 -> one bucket
        return "EV"
    return None


def _district(name):
    t = re.sub(r"\s+", "", (name or "")).upper()
    if not t or t == "GRANDTOTAL":
        return None
    t = DISTRICT_ALIAS.get(t, t)
    return t if t in DISTRICTS else "?" + (name or "").strip()


# v11.17 - the report is now the sheet "data": one LONG table, one row per delivery date x timeslot x district x order date.
#   A delivery_date | B delivery_timeslot | C Delivery Address Type | D district (group) | E order creation date (2026年11月10日) | F # of Consignee Address |
#   G # of Order | H # of Parent Order | I Same pk-address-date-timeslot Combination
# Used: A delivery date, B timeslot, D district, E order date, H parent order count. (Pick-up / Macau / overseas rows carry other district codes -> skipped.)
COL_DELIVERY, COL_SLOT, COL_DISTRICT, COL_ORDER_DATE, COL_COUNT = 0, 1, 3, 4, 7


def parse_report(path):
    """-> {'rows': [(district, order date ISO, delivery date ISO, 'AM'|'PM'|'EV'|None, parent order count)], 'unknownSlots': set, 'unknownDistricts': set,
           'badDates': int, 'format': 'data'}. Raises when the layout is not the expected one."""
    with open(path, "rb") as _f:
        _head = _f.read(8)
    if _head[:2] == b"PK":
        raise RuntimeError("the downloaded file is an Excel/zip file, not a CSV - the CSV format was not selected in the Crosstab dialog")
    if _head[:5].lower() in (b"<!doc", b"<html"):
        raise RuntimeError("the downloaded file is an HTML page, not a CSV (Tableau error / sign-in page?)")
    text = _read_text(path)
    first = text.split("\n", 1)[0]
    rows = list(csv.reader(io.StringIO(text), delimiter="\t" if "\t" in first else ","))
    if not rows or len(rows[0]) <= COL_COUNT:
        raise RuntimeError(f"unexpected Tableau csv layout: the header has {len(rows[0]) if rows else 0} columns, expected at least {COL_COUNT + 1} (first row: {rows[0][:9] if rows else []})")
    hd = [(c or "").strip().lower() for c in rows[0]]
    checks = [(COL_DELIVERY, "delivery", "A delivery_date"), (COL_SLOT, "timeslot", "B delivery_timeslot"), (COL_DISTRICT, "district", "D district (group)"),
              (COL_ORDER_DATE, "order", "E order date"), (COL_COUNT, "parent", "H # of Parent Order")]
    bad = [name for ix, key, name in checks if key not in hd[ix]]
    if bad:
        raise RuntimeError(f"unexpected Tableau csv layout: header column(s) {bad} not as expected. Header row: {[c.strip() for c in rows[0][:9]]}")
    out, unknown_slots, unknown_d, bad_dates = [], set(), set(), 0
    for r in rows[1:]:
        if len(r) <= COL_COUNT:
            continue
        d = _district(r[COL_DISTRICT])
        if d is None:
            continue
        if d.startswith("?"):
            unknown_d.add(d[1:])                                  # pick-up points, Macau, overseas ... (not one of the 10 districts)
            continue
        dd, od = _to_iso(r[COL_DELIVERY], "iso"), _to_iso(r[COL_ORDER_DATE], "iso")
        if dd is None or od is None:
            bad_dates += 1
            continue
        slot = _slot_key(r[COL_SLOT])
        if slot is None or slot == "TOTAL":
            unknown_slots.add((r[COL_SLOT] or "").strip())
            slot = None
        out.append((d, od, dd, slot, _num(r[COL_COUNT])))
    if not out:
        raise RuntimeError(f"unexpected Tableau csv layout: no usable data rows ({len(rows) - 1} lines read, {bad_dates} with an unreadable date)")
    return {"rows": out, "unknownSlots": unknown_slots, "unknownDistricts": unknown_d, "badDates": bad_dates, "format": "data"}


def build_cells(rep, order_date):
    """Rows with Order Date (column E) = order_date -> {deliveryDate: {district|'OVERALL': [total, AM, PM, EV&EV2]}}, stats.
    Count = # of Parent Order (column H). NT-YT & WTH rows are WTH. Total = the sum of the timeslots (a timeslot other than AM / PM / EV counts in Total only)."""
    od_dt = datetime.strptime(order_date, "%Y-%m-%d")
    acc, matched = {}, 0
    for dist, od, dd, slot, n in rep["rows"]:
        if od != order_date:
            continue
        matched += 1
        a = acc.setdefault(dd, {}).setdefault(dist, {})
        a["TOTAL"] = a.get("TOTAL", 0) + n
        if slot:
            a[slot] = a.get(slot, 0) + n
    cells, mismatch, skipped = {}, 0, []
    for dd in sorted(acc):
        lead = (datetime.strptime(dd, "%Y-%m-%d") - od_dt).days
        if lead < 0 or lead > MAX_LEAD_DAYS:
            skipped.append(dd)
            continue
        blk, overall = {}, [0, 0, 0, 0]
        for dist in DISTRICTS:
            a = acc[dd].get(dist, {})
            am, pm, ev = a.get("AM", 0), a.get("PM", 0), a.get("EV", 0)
            tot = a.get("TOTAL", 0)
            if tot != am + pm + ev:
                mismatch += 1
            blk[dist] = [tot, am, pm, ev]
            overall = [x + y for x, y in zip(overall, blk[dist])]
        cells[dd] = {"OVERALL": overall, **blk}
    return cells, {"rows": matched, "deliveryDates": len(cells), "skippedDeliveryDates": skipped, "totalNotEqualSlots": mismatch}


# ----------------------------------------------------------------------------- JSON + embedding
def load_json():
    if os.path.exists(ROLLING_JSON):
        try:
            return json.load(open(ROLLING_JSON, encoding="utf-8"))
        except Exception as ex:
            log(f"WARNING: {ROLLING_JSON} could not be read ({ex}) - starting a new file")
    return {}


def yesterday_open(data, t0):
    """True when yesterday is a version-2 day that the 00:00 run never finished (missed run) -> this run finishes it."""
    d = (data.get("days") or {}).get((t0 - timedelta(days=1)).strftime("%Y-%m-%d"))
    return bool(d and d.get("v") == 2 and not d.get("complete"))


def _merge_cells(old, new):
    """Delivery dates that the report shows replace the stored ones (a count can also go DOWN); a delivery date the report no longer shows
    keeps its last value; a brand-new delivery date that is all zero is not created."""
    out = dict(old or {})
    for dd, blk in new.items():
        if dd not in out and not any(v for vec in blk.values() for v in vec):
            continue
        out[dd] = blk
    return dict(sorted(out.items()))


def update_json(updates, run_slot, stats):
    """updates = [(order date, cells, final)]"""
    data = load_json()
    data.setdefault("days", {})
    data.setdefault("snapshots", {})
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    written = []
    for od, cells, final in updates:
        old = data["days"].get(od)
        if old and old.get("v") != 2:
            log(f"WARNING: {od} already holds a day written by the old (parent-order) logic - left untouched")
            continue
        merged = _merge_cells((old or {}).get("cells"), cells)
        data["days"][od] = {"v": 2, "slot": run_slot, "asOf": now, "complete": bool(final or (old or {}).get("complete")), "cells": merged}
        written.append(od)
    if not written:
        raise SystemExit("nothing was written to the json")
    keep = sorted(data["days"])[-KEEP_ORDER_DATES:]
    data["days"] = {k: data["days"][k] for k in keep}
    keep_snap = set(sorted(data["days"])[-KEEP_SNAPSHOT_DATES:])
    data["snapshots"] = {k: v for k, v in data["snapshots"].items() if k in keep_snap}
    data["districts"] = DISTRICTS
    data["generatedAt"] = now
    data["lastRun"] = {"orderDate": max(written), "orderDates": written, "slot": run_slot, "source": SOURCE_NAME, "stats": stats}
    tmp = ROLLING_JSON + ".tmp"
    os.makedirs(os.path.dirname(ROLLING_JSON), exist_ok=True)
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))     # compact: it is pushed + embedded
    os.replace(tmp, ROLLING_JSON)
    log(f"wrote {ROLLING_JSON}  (order dates {', '.join(written)})")
    return data, written


# ----------------------------------------------------------------------------- git push to the dashboard
def _script_safe(js):
    """Same escaping kpi_pipeline.py uses: a raw "</" would end the <script> tag, U+2028/9 break JS string literals."""
    return js.replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")


def embed_in_index(data, path=None):
    """Splice rolling_orders.json into window.__EMBEDDED_DATA__ (between the EMBEDDED_DATA markers) of `path`
    (default INDEX_HTML). Everything else in the page - and every other embedded file - is left as it is."""
    path = path or INDEX_HTML
    if not os.path.exists(path):
        log("index.html not found - embedding skipped")
        return False
    s = open(path, encoding="utf-8").read()
    a = s.find("<!-- EMBEDDED_DATA_START")
    b = s.find("EMBEDDED_DATA_END")
    if a < 0 or b < 0:
        log("EMBEDDED_DATA markers not found - embedding skipped")
        return False
    e = s.index("-->", b) + 3
    blk = s[a:e]
    m = re.search(r"window\.__EMBEDDED_DATA__ = ", blk)
    js = blk[m.end():blk.rindex("</script>")].rstrip().rstrip(";")
    emb = json.loads(js)
    emb["rolling_orders.json"] = data
    head = blk[:blk.index("<script>")]
    tail = blk[blk.index("</script>") + len("</script>"):]
    new_blk = head + "<script>window.__EMBEDDED_DATA__ = " + _script_safe(json.dumps(emb, ensure_ascii=False)) + ";</script>" + tail
    s2 = s[:a] + new_blk + s[e:]
    if os.path.realpath(path) == os.path.realpath(INDEX_HTML):
        shutil.copy2(path, path + ".bak_rolling")
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(s2)
    os.replace(tmp, path)
    log(f"embedded rolling_orders.json into {os.path.basename(path)}")
    return True


# ----------------------------------------------------------------------------- git push to the dashboard
def _git(args, cwd):
    env = dict(os.environ, GIT_EDITOR="true", GIT_TERMINAL_PROMPT="0")
    r = subprocess.run(["git"] + args, cwd=cwd, capture_output=True, text=True, timeout=GIT_TIMEOUT, env=env)
    return r.returncode, ((r.stdout or "") + (r.stderr or "")).strip()


def _rebase_in_progress(top):
    for name in ("rebase-merge", "rebase-apply"):
        rc, out = _git(["rev-parse", "--git-path", name], top)
        if rc == 0 and os.path.isdir(os.path.join(top, out) if not os.path.isabs(out) else out):
            return True
    return False


_CONFLICT_RE = re.compile(r"<<<<<<< [^\n]*\n(.*?)(?:\|\|\|\|\|\|\| [^\n]*\n.*?)?=======\n(.*?)>>>>>>> [^\n]*(?:\n|$)", re.S)


def _resolve_index_conflict(path, data):
    """index.html conflicts because kpi_pipeline.py and this script both rewrite the one-line EMBEDDED_DATA block.
    During a rebase the first side of a conflict is the UPSTREAM one. If every conflict hunk is inside the data block,
    keep upstream's block (its other embedded files are the newest) and splice rolling_orders.json back into it -
    code edits on either side were merged by git outside the hunks and are untouched. Anything else is not ours to guess."""
    s = open(path, encoding="utf-8").read()
    hunks = list(_CONFLICT_RE.finditer(s))
    if not hunks or not all("__EMBEDDED_DATA__" in m.group(1) or "__EMBEDDED_DATA__" in m.group(2) for m in hunks):
        return False
    s = _CONFLICT_RE.sub(lambda m: m.group(1), s)
    if "<<<<<<<" in s or ">>>>>>>" in s or "<!-- EMBEDDED_DATA_START" not in s:
        return False
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(s)
    return embed_in_index(data, path)


def _resolve_rebase_conflicts(top, rel_json, rel_index, data):
    """True when the rebase finished (all conflicts were the data block / our own json); False when it was aborted."""
    for _ in range(10):                         # one round per local commit being replayed
        rc, out = _git(["diff", "--name-only", "--diff-filter=U"], top)
        conflicted = [x.strip() for x in out.splitlines() if x.strip()]
        if not conflicted:
            break
        for f in conflicted:
            if f == rel_json:
                ok = _git(["checkout", "--theirs", "--", f], top)[0] == 0      # rebase: "theirs" = our commit = newest rolling data
            elif f == rel_index:
                ok = _resolve_index_conflict(os.path.join(top, f), data)
            else:
                ok = False
            if not ok:
                log(f"push: cannot auto-resolve the conflict in {f} - aborting the rebase, nothing was lost locally")
                _git(["rebase", "--abort"], top)
                return False
            _git(["add", "--", f], top)
        rc, out = _git(["rebase", "--continue"], top)
        if rc == 0:
            break
    if _rebase_in_progress(top):
        log("push: rebase still not finished after auto-resolving - aborting it")
        _git(["rebase", "--abort"], top)
        return False
    return True


def push_dashboard(order_date, slot, data):
    """pull --rebase -> commit our two files -> push, retried. NEVER raises: the data files are already written locally."""
    if not PUSH_ENABLED:
        log("push: disabled (ROLLING_DEPLOY=none / --skip-push) - dashboard files updated locally only")
        return False
    try:
        return _push_dashboard(order_date, slot, data)
    except Exception as ex:
        log(f"push: FAILED ({ex}) - the dashboard was NOT updated online")
        return False


def _push_dashboard(order_date, slot, data):
    rc, top = _git(["rev-parse", "--show-toplevel"], DEPLOY_CWD)
    if rc != 0:
        log(f"push: {DEPLOY_CWD} is not inside a git repository ({top}) - set ROLLING_DEPLOY_CWD")
        return False
    real_top = os.path.realpath(top)
    rels = {}
    for name, f in (("json", ROLLING_JSON), ("index", INDEX_HTML)):
        if os.path.exists(f):
            r = os.path.relpath(os.path.realpath(f), real_top)
            if r.startswith(".."):
                log(f"push: {f} is outside the repository - skipped")
            else:
                rels[name] = r.replace(os.sep, "/")
    if not rels:
        log("push: nothing to push")
        return False
    files = list(rels.values())
    rc, out = _git(["add", "--"] + files, top)
    if rc != 0:
        log(f"push: git add failed (is public/rolling_orders.json in .gitignore?): {out[-400:]}")
        return False
    if _git(["diff", "--cached", "--quiet", "--"] + files, top)[0] != 0:          # something to commit (only OUR files)
        rc, out = _git(["commit", "-m", f"Rolling orders auto update - order date {order_date:%Y-%m-%d}, run {datetime.now():%Y-%m-%d} {slot}", "--"] + files, top)
        if rc != 0:
            log(f"push: commit failed: {out[-400:]}")
            return False
    else:
        log("push: files unchanged since the last commit")
    for attempt in range(1, PUSH_ATTEMPTS + 1):
        rc, out = _git(["pull", "--rebase", "--autostash"], top)
        if rc != 0:
            if _rebase_in_progress(top):
                if not _resolve_rebase_conflicts(top, rels.get("json"), rels.get("index"), data):
                    log("push: " + out[-400:])
                    return False
            else:
                # rebase itself completed; typically the autostash of someone else's uncommitted file did not re-apply cleanly
                log(f"push: pull reported a problem but no rebase is pending - continuing ({out[-300:]})")
        rc, out = _git(["push"], top)
        if rc == 0:
            log(f"push: dashboard pushed (attempt {attempt})" + (" - " + out.splitlines()[-1] if out else ""))
            return True
        log(f"push: attempt {attempt}/{PUSH_ATTEMPTS} rejected ({out[-250:]}) - pulling again")
        time.sleep(3)
    log("push: gave up - the commit is local; the next session will push it together with its own update")
    return False



# ----------------------------------------------------------------------------- main
@contextlib.contextmanager
def run_lock():
    os.makedirs(LOG_DIR, exist_ok=True)
    fh = open("/tmp/rolling_orders.lock", "w")
    try:
        fcntl.flock(fh, fcntl.LOCK_EX)         # runs queue up instead of overlapping
        yield
    finally:
        fcntl.flock(fh, fcntl.LOCK_UN)
        fh.close()



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="T+0 = the run date YYYY-MM-DD (default today). Only this order date is updated (add --both for yesterday too)")
    ap.add_argument("--both", action="store_true", help="force the two-day update: yesterday (final) + today, as the 00:00 run does")
    ap.add_argument("--order-date", action="append", metavar="YYYY-MM-DD", help="read this order date instead of T+0 (repeatable / comma separated); "
                    "stored as not complete unless --final")
    ap.add_argument("--final", action="store_true", help="with --order-date: mark those days complete")
    ap.add_argument("--skip-download", "--skip-tableau-download", dest="skip_download", action="store_true", help="use the csv already in Folder for Rolling Orders")
    ap.add_argument("--tableau-file", help="process this csv instead of the downloaded one (never deleted)")
    ap.add_argument("--process-only", action="store_true", help="no download, no git push; uses --tableau-file or the file in the folder")
    ap.add_argument("--headed", action="store_true", help="show the Chrome window during the Tableau download (same as ROLLING_TABLEAU_HEADLESS=0)")
    ap.add_argument("--tableau-only", "--download-only", dest="tableau_only", action="store_true", help="only download the csv, then stop")
    ap.add_argument("--push-only", action="store_true", help="no download / processing: embed the rolling_orders.json that is already in the public folder into index.html and git-push it")
    ap.add_argument("--keep-files", action="store_true", help="do not delete the downloaded csv after it was written to the json (same as ROLLING_KEEP_FILES=1)")
    ap.add_argument("--skip-embed", action="store_true")
    ap.add_argument("--skip-push", action="store_true", help="do not git-push the dashboard files (same as ROLLING_DEPLOY=none)")
    ap.add_argument("--slot", help="ignored - kept so an old cron line does not break")
    ap.add_argument("--skip-whatsapp", action="store_true", help="ignored - kept so old command lines still work")
    a = ap.parse_args()
    global TABLEAU_HEADLESS
    if a.headed:
        TABLEAU_HEADLESS = False

    now = datetime.now()
    run_slot = f"{now:%H}:00"
    if a.push_only:
        with run_lock():
            data = json.load(open(ROLLING_JSON, encoding="utf-8"))
            last = max(data.get("days", {}) or [now.strftime("%Y-%m-%d")])
            log(f"push-only: {ROLLING_JSON} holds {len(data.get('days', {}))} order dates, latest {last}, generatedAt {data.get('generatedAt')}")
            if not a.skip_embed:
                embed_in_index(data)
            if not a.skip_push:
                log("push-only: result = " + ("PUSHED" if push_dashboard(datetime.strptime(last, "%Y-%m-%d"), run_slot, data) else "NOT pushed (see the push: lines above)"))
        return

    t0 = datetime.strptime(a.date, "%Y-%m-%d") if a.date else datetime(now.year, now.month, now.day)
    with run_lock():
        if a.order_date:
            plan = [(d.strip(), a.final) for x in a.order_date for d in x.split(",") if d.strip()]
        else:
            plan = order_dates_for_run(now, t0 if a.date else None, a.both, catchup=yesterday_open(load_json(), t0) if not a.date else False)
        log(f"run {now:%Y-%m-%d %H:%M}  T+0={t0:%Y-%m-%d}  ->  order dates to update: " + ", ".join(f"{d}{' (final)' if f else ''}" for d, f in plan))
        path = a.tableau_file or tableau_report_path(t0 - timedelta(days=1))
        if a.tableau_only or not (a.skip_download or a.process_only or a.tableau_file):
            download_tableau_report(path)
        if a.tableau_only:
            log("tableau-only: done")
            return
        log(f"tableau file: {path}  ->  {'FOUND' if os.path.exists(path) else 'NOT FOUND'}")
        if not os.path.exists(path):
            raise SystemExit("no Tableau file - run without --skip-download / --process-only, or pass --tableau-file")
        rep = parse_report(path)
        dd_all = sorted({r[2] for r in rep["rows"]})
        ods = sorted({r[1] for r in rep["rows"]})
        log(f"report: {len(rep['rows'])} data rows, delivery dates {dd_all[0]} -> {dd_all[-1]} ({len(dd_all)}), order dates {ods[0] if ods else '-'} -> {ods[-1] if ods else '-'}, date format {rep['format']}")
        if rep["unknownSlots"]:
            log(f"WARNING: timeslot values not recognised (not counted, only in Total): {sorted(rep['unknownSlots'])}")
        if rep["unknownDistricts"]:
            log(f"skipped {len(rep['unknownDistricts'])} district (group) values that are not one of the 10 districts (pick-up points, Macau, overseas ...)")
        updates, stats = [], {"reportRows": len(rep["rows"]), "deliveryFrom": dd_all[0], "deliveryTo": dd_all[-1], "dates": {}}
        for od, final in plan:
            cells, st = build_cells(rep, od)
            stats["dates"][od] = st
            if not st["rows"] or not cells:
                log(f"WARNING: no rows with Order Date = {od} in the report - {od} not updated")
                continue
            if st["totalNotEqualSlots"]:
                log(f"WARNING: {od}: Total <> AM+PM+EV in {st['totalNotEqualSlots']} district-day blocks (a timeslot other than AM / PM / EV?)")
            ov = {dd: c["OVERALL"] for dd, c in cells.items()}
            log(f"{od}{' (final)' if final else ''}: {st['rows']} district rows, {len(cells)} delivery dates, T+0 total {ov.get(od, [0])[0]}, all dates total {sum(v[0] for v in ov.values())}")
            updates.append((od, cells, final))
        if not updates:
            raise SystemExit(f"no usable rows for the order date(s) {', '.join(d for d, _ in plan)} - nothing written")
        data, written = update_json(updates, run_slot, stats)
        if not (a.keep_files or KEEP_FILES or a.process_only):
            cleanup_files([None if a.tableau_file else path])
        if not a.skip_embed:
            embed_in_index(data)
        if not (a.skip_push or a.process_only):
            push_dashboard(datetime.strptime(max(written), "%Y-%m-%d"), run_slot, data)       # publish the update to the Rolling Orders tab


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
