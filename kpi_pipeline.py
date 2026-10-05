#!/usr/bin/env python3
"""
LOG · KPI Tracker — Data Pipeline  (v9.5)
==================================
v10.4 — a backfilled OIX day now ALSO refreshes everything that is calculated from that day's manpower, not just the
  Manpower Distribution tab: dailyProductivityLog (HKTV Staff manpower with leader-exclusion + ODS/VAN manpower, and each
  row's productivity = orders / manpower; Cost-Report productivity is kept), the 7-day rolling series, every month-to-date
  productivityMtdLog snapshot from the corrected day onward, data.json's HKTV Staff / ODS Ratio actual + forecast,
  manpower_staging.json (when the corrected day is the staged T-1), productivity_history.json and the embedded snapshot.
  See apply_manpower_corrections(). The 14:00 job also runs the backfill before it reads the staging file.
v10.3 — OIX manpower backfill: a day's HKTV Manpower Distribution is recomputed whenever its OIX_Record file is new or has
  been UPDATED since it was last processed (SHA-1 fingerprint in history.json "oixManpowerProcessed"). Runs inside the
  on demand via `--section oixbackfill` [--oix-days N] [--force-oix], by default for OIX_Record files whose file
  modified time is today (T+0); older days are covered by the Daily Cost Report. The 03:00 `productivity` job
  fingerprints its T-1 file and also handles any other file modified today. Cost-Report days keep their real Courier/Driver; only ODS/VAN is refreshed.
v9.5 — `--section newestate` (new_estate_tracker.py): tracks newly launched private / public housing that is ready for
  move-in -> New_Estate_Tracker.xlsx + new_estates.json (address -> lat/long). Weekly job; see the module docstring.
v9.0 — Fulfillment Cost % (monthly): when the Daily Cost Report is processed, Total Cost (Overview tab,
  'Total Cost' row, MTD column - Overall and per district) / GMV summed from the 1st of the month through the
  report's Last Update date -> other_aspects_history.json 'fulfillmentCost' (Other Aspects Tracking tab).
  The month still open (report not yet at month end) is shown as 'YYYY-MM (MTD)'.
v8.2 — the WhatsApp cost-report download (`--section costreport`) no longer shows a browser window.
  Windows/macOS: Chromium runs in Chrome's "new headless" mode (a full browser engine that WhatsApp Web
  accepts, unlike the old headless shell) with a normal Chrome user-agent. Linux: Xvfb virtual display.
  WA_HIDE_MODE = auto | headless | offscreen | xvfb | off  (default auto; `off` = visible window, use it to
  re-scan the WhatsApp QR code).
v8.1 — the "Actual Delivery - 10 Districts" Tableau sheet is no longer treated as pre-selected;
  its thumbnail is clicked like the other Delivery Dashboard sheets.

v8.0 — Adjustment in Data Fetching:
  1. ODS order / waybill counts now come from the Tableau "Delivery Dashboard"
     sheet "Actual Delivery - 10 Districts" (rows headed "包派"), not OIX.
     HKTV order count = total order count - ODS order count.
  2. HKTV manpower distribution (FT/PT Driver & Courier), actual Cost per Order
     (Controllable) and Staff Productivity now come from the "Daily Cost Report"
     Excel posted in the WhatsApp group; run with `--section costreport`
     (scheduled 09:00 every Tuesday and Friday). The 03:00 OIX manpower job is
     unchanged; the fetched cost-report data then covers the older figures.
整合 Playwright 自動化下載 Tableau Crosstab 與 Pandas 數據處理。
1. 使用 Playwright 登入 Tableau，模擬點擊 Download -> Crosstab 下載 CSV。
2. 清洗 CSV 數據並應用 KPI 商業邏輯。
3. 輸出 data.json 供 Dashboard 讀取。
"""

import os
import re
import sys
import json
import glob
import hashlib
import time
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd
from playwright.sync_api import sync_playwright

# Windows Task Scheduler 執行時，stdout/stderr 會被導向到檔案 (pipeline_log.txt)
# 而非真正的主控台，此時 Python 常會退回系統的 ANSI codepage (cp1252)，
# 導致印出 emoji 或中文字元時丟出 UnicodeEncodeError 而讓整支程式當掉。
# 這裡強制 stdout/stderr 用 UTF-8，不管是手動執行還是排程執行都不會再炸掉。
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass  # 極舊版本 Python 或非標準串流時，安靜略過，不影響主流程

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# =============================================================================
# 1. 環境變數與路徑設定 (CONFIGURATION & PATHS)
# =============================================================================

# Tableau 登入資訊 (建議放 .env)
TABLEAU_URL = os.environ.get("TABLEAU_URL", "https://inhouse-analytics.hktv.com.hk/#/signin")
TABLEAU_DASHBOARD_URL = os.environ.get("TABLEAU_DASHBOARD_URL", "")
TABLEAU_USER = os.environ.get("TABLEAU_USER", "")
TABLEAU_PASS = os.environ.get("TABLEAU_PASS", "")

if not TABLEAU_USER or not TABLEAU_PASS:
    raise SystemExit(
        "TABLEAU_USER / TABLEAU_PASS are not set. Create a .env file next to this "
        "script (see .env.example) with your real credentials — the script no "
        "longer falls back to a hardcoded password."
    )

# --- v3.0 §3: GMV lives on a SEPARATE Tableau account from TABLEAU_USER/PASS
# above ("LOG GMV for AI Fetching" is only visible under that other login),
# so it gets its own sign-in URL + credentials. fetch_gmv_report() opens its
# own browser context and logs in with these, independently of
# fetch_tableau_reports(). Not required unless you actually run the GMV
# fetch, so no SystemExit here — run_section_tableau() just skips GMV with a
# warning if the file never showed up (see there).
TABLEAU_GMV_URL = os.environ.get("TABLEAU_GMV_URL", "https://inhouse-analytics.hktv.com.hk/#/signin")
TABLEAU_GMV_DASHBOARD_URL = os.environ.get(
    "TABLEAU_GMV_DASHBOARD_URL",
    "https://inhouse-analytics.hktv.com.hk/#/views/LogGMVforAIFetching/Sheet1?:iid=1",
)
TABLEAU_GMV_USER = os.environ.get("TABLEAU_GMV_USER", "")
TABLEAU_GMV_PASS = os.environ.get("TABLEAU_GMV_PASS", "")
# v3.0 §3 fix: numeric workbook ID for "LOG GMV for AI Fetching", used to log
# in via a redirect straight to that workbook's views listing — same pattern
# as oix_returning_waybill.py's TABLEAU_URL ("...?redirect=%2Fworkbooks%2F4896%2Fviews").
# Find it by opening the GMV workbook itself (not the direct view link) while
# logged in as the GMV account, and copying the number from the resulting
# "#/workbooks/<ID>/views" URL. Left blank, fetch_gmv_report() falls back to
# the old direct-view goto()+reload() approach, which is what was failing.
TABLEAU_GMV_WORKBOOK_ID = os.environ.get("TABLEAU_GMV_WORKBOOK_ID", "")

# 目錄設定
OIX_FOLDER = os.environ.get("OIX_FOLDER", r"C:\Users\chipanl\Downloads\Digimobi Report")
REPORT_FOLDER = os.environ.get("REPORT_FOLDER", r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\Folder for KPI Dashboard")
# v10.0 — Staff List switched from the local "Logistics_Staff_List_YYYYMMDD.xlsx"
# export (v3.0 §4, STAFF_LIST_FOLDER below) to the "Master LOG Staff List"
# Google Sheet — it's the actively-maintained source and doesn't depend on
# someone remembering to re-export Excel. Sheet URL, tab name, credential
# path, and column layout are copied verbatim from ot_time_alert_workflow.py's
# STAFF_MASTER_SHEET_URL / STAFF_MASTER_TAB_NAME / GOOGLE_CREDENTIAL_JSON /
# STAFF_MASTER_*_COL — same sheet, same service-account credential, so both
# scripts stay in sync if the sheet ever moves. See load_staff_master_df().
STAFF_MASTER_SHEET_URL = "https://docs.google.com/spreadsheets/d/1HT7KstK1iLUxWeprTZkGPzzjLzQ3PeXSeE9vdZ6cVp0/edit?gid=0#gid=0"
STAFF_MASTER_TAB_NAME = "LOG Master"
GOOGLE_CREDENTIAL_JSON = os.environ.get(
    "GOOGLE_CREDENTIAL_JSON",
    # v10.1 fix — ot_time_alert_workflow.py's own comment on its WORK_DIR
    # says its /mnt/c/... paths are a WSL-only convention and need
    # adjusting to a plain Windows path when running directly under
    # Windows Python instead of WSL. I copied the credential path from
    # there verbatim and missed that adjustment; kpi_pipeline.py runs via
    # run_productivity_0300.bat -> `py kpi_pipeline.py` under Task
    # Scheduler, i.e. plain Windows Python (same reason OIX_FOLDER /
    # REPORT_FOLDER / STAFF_LIST_FOLDER above are all raw C:\ paths, not
    # /mnt/c/...) — confirmed by the trial-run pipeline_log.txt traceback:
    # "credential not found at '/mnt/c/Users/...'". Fixed to the native
    # Windows path here. Also made configurable via env var, matching the
    # convention every other path in this section already follows.
    r"C:\Users\chipanl\Downloads\Whatsapp Session\digimobi-temperature-review-d88603b35531.json"
)
STAFF_MASTER_STAFFID_COL = "A"    # Staff ID
STAFF_MASTER_DEPT_CODE_COL = "E"  # Dept
STAFF_MASTER_POSITION_COL = "G"   # Position
# The sheet has no separate "Employment Status" column the way the old Excel
# export did — Last Working Date (I) blank is treated as still-employed.
# See load_staff_master_df()'s docstring for the notice-period caveat.
STAFF_MASTER_LAST_WORKING_DATE_COL = "I"
STAFF_MASTER_HEADER_ROW = 1  # row 1 is the header row

# v3.0 §4 (DEPRECATED as of v10.0, replaced by the Google Sheet above — kept
# here, unused, only so old pipeline_log.txt entries referencing this path
# still make sense when read back later).
STAFF_LIST_FOLDER = os.environ.get("STAFF_LIST_FOLDER", r"C:\Users\chipanl\Downloads\Staff List")
DATA_JSON_PATH = os.environ.get("DATA_JSON_PATH", "./public/data.json")
HISTORY_PATH = os.environ.get("HISTORY_PATH", "./history.json")
# The dashboard's Productivity Detail / Daily Records tabs fetch this file
# directly (not data.json) — see loadDataSource() in the HTML. It lives next
# to data.json in ./public so both get served by the same static host.
PRODUCTIVITY_HISTORY_PATH = os.environ.get("PRODUCTIVITY_HISTORY_PATH", "./public/productivity_history.json")
# v3.0 §3: GMV / Basket Size tab's data file — daily rows for the current
# (still-open) month, plus one accumulated row per CLOSED month. See
# build_gmv_monthly().
GMV_HISTORY_PATH = os.environ.get("GMV_HISTORY_PATH", "./public/gmv_history.json")
# v3.0 §4: HKTV Manpower Distribution tab's data file — same daily-log /
# trimmed-window pattern as PRODUCTIVITY_HISTORY_PATH.
MANPOWER_HISTORY_PATH = os.environ.get("MANPOWER_HISTORY_PATH", "./public/manpower_distribution.json")
# v10.3 — OIX manpower backfill. An OIX_Record export is often picked up while it only holds PART of the
# day's records (e.g. 2026-10-02..10-04 were first logged with ODS/VAN = 22 / 43 / 47 and the completed
# files give 77 / 189 / 145). Whenever the file for a day later turns out to be different (updated), the
# day's manpower distribution is recomputed from it. Default: only OIX_Record files whose FILE MODIFIED TIME
# is today (T+0) are looked at, whatever date is in their name; older, untouched files are left alone (the
# Daily Cost Report covers those days). Set OIX_BACKFILL_DAYS (or use --oix-days N) to instead re-check every
# file of the last N days regardless of modified time; the check itself is cheap (one SHA-1 per file).
OIX_BACKFILL_DAYS = int(os.environ["OIX_BACKFILL_DAYS"]) if os.environ.get("OIX_BACKFILL_DAYS", "").strip() else None
OIX_FILE_RE = re.compile(r"^OIX_Record_(\d{8})\.(xlsx|csv)$", re.IGNORECASE)
# v4.0 §1: Productivity now needs BOTH manpower (from OIX, available at
# 03:00) and parent order counts (from the Tableau "Delivery Dashboard"
# report, only downloaded in the 14:00 Tableau job — see run_tableau_1400.*).
# The 03:00 job can no longer finish the Productivity matrices on its own, so
# it stashes yesterday's manpower headcounts here; the 14:00 job picks this
# up once the Tableau order counts are in and finishes the calculation. Not
# served to the dashboard — internal handoff file only.
MANPOWER_STAGING_PATH = os.environ.get("MANPOWER_STAGING_PATH", "./manpower_staging.json")
# v4.0 §2: new "Delay %" tab's data source — daily (T-1) + MTD, by timeslot
# and district, same daily/monthly-after-close pattern as gmv_history.json.
DELAY_HISTORY_PATH = os.environ.get("DELAY_HISTORY_PATH", "./public/delay_history.json")
# v4.0 §4: new "Other Aspects Tracking" tab — poor rating %, missing & lost
# amount, RFID missing tote, logged monthly in the same layout as GMV/Basket
# Size (see build_gmv_monthly()).
OTHER_ASPECTS_HISTORY_PATH = os.environ.get("OTHER_ASPECTS_HISTORY_PATH", "./public/other_aspects_history.json")
# v28.0 — single deployable dashboard file. Every JSON file's parsed content
# gets burned into a window.__EMBEDDED_DATA__ script tag inside this same
# file at the end of every "tableau" run, so it works both live off a real
# server (fetchJSON() tries fetch() first) and opened from disk with zero
# network calls (falls back to the embedded snapshot) — see
# update_embedded_data(). Only the marked block between EMBEDDED_DATA_START/
# END gets rewritten each run; hand-editing the rest of the file in between
# runs is safe.
INDEX_HTML_PATH = os.environ.get("INDEX_HTML_PATH", "./public/index.html")
# How many days of raw order-count/manpower history productivity_history.json
# carries. history.json (not this) is the durable full log, so raising this
# later doesn't lose anything already run — it just widens the served window.
DAILY_PRODUCTIVITY_KEEP_DAYS = int(os.environ.get("DAILY_PRODUCTIVITY_KEEP_DAYS", "60"))
# v3.0 §3 fix: run_section_tableau()'s file check used to only check
# existence, not freshness. Since REPORT_FOLDER's CSVs are never deleted
# between runs, a report whose *download* silently failed this run (Crosstab
# menu not found, download button not found, etc. — all seen intermittently
# in pipeline_log.txt) would leave yesterday's leftover CSV sitting there,
# existence-check would pass, and the pipeline would silently recompute
# today's dashboard numbers from stale data with no warning. A file older
# than this many hours is now treated the same as a missing file. 2h is
# generous slack over how long fetch_tableau_reports()+fetch_gmv_report()
# actually take to run immediately before run_section_tableau() reads them.
STALE_REPORT_HOURS = float(os.environ.get("STALE_REPORT_HOURS", "2"))

# v8.0 — WhatsApp / Daily Cost Report (see run_cost_report_section()).
# The WhatsApp Web session is a persistent Chromium profile that has already
# been logged in once by scanning the QR code (re-scan if the session expires).
WA_SESSION_DIR = os.environ.get("WA_SESSION_DIR", "/home/chipanl/whatsapp_session_2")
WA_TARGET_GROUP = os.environ.get("WA_TARGET_GROUP", "LOG 區頭 x Head office")
WA_HEADLESS = os.environ.get("WA_HEADLESS", "0") == "1"   # WhatsApp Web often refuses headless; default = visible window
# v8.2 — hide the WhatsApp Web browser window (WA_HIDE_MODE):
#   auto      -> "headless" on Windows/macOS, "xvfb" on Linux (default)
#   headless  -> Chromium's NEW headless mode (--headless=new): a full browser, not the old headless shell that
#                WhatsApp Web refuses, with a normal Chrome user-agent (the new mode still says "HeadlessChrome")
#   offscreen -> normal window parked far off-screen (fallback; may still flash in the taskbar)
#   xvfb      -> Linux virtual display (needs `apt install xvfb`)
#   off       -> visible window (use this to re-scan the QR code after the session expires)
# The old WA_VIRTUAL_DISPLAY=0 switch still works and means "off".
WA_HIDE_MODE = os.environ.get("WA_HIDE_MODE", "auto").strip().lower()
if os.environ.get("WA_VIRTUAL_DISPLAY") == "0":
    WA_HIDE_MODE = "off"
if WA_HIDE_MODE == "auto":
    WA_HIDE_MODE = "xvfb" if sys.platform.startswith("linux") else "headless"
WA_VIRTUAL_SCREEN = os.environ.get("WA_VIRTUAL_SCREEN", "1920x1080x24")
WA_USER_AGENT = os.environ.get(
    "WA_USER_AGENT",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")
COST_REPORT_FOLDER = os.environ.get("COST_REPORT_FOLDER", REPORT_FOLDER)
# "Daily Cost Report_YYYYMM_Last Update_MMM DD.xlsx" (e.g. Daily Cost Report_202609_Last Update_Sep 27.xlsx).
# Spaces may appear as underscores once a file has been saved/renamed by a browser or chat app, so both are accepted.
COST_REPORT_NAME_RE = re.compile(
    r"^Daily[ _]Cost[ _]Report_(\d{6})_Last[ _]Update_([A-Za-z]{3})[ _](\d{1,2})(?:\b.*)?\.xlsx$", re.IGNORECASE)

os.makedirs(REPORT_FOLDER, exist_ok=True)
os.makedirs(COST_REPORT_FOLDER, exist_ok=True)

# 報表精確名稱對應
REPORT_FILES = {
    "report_a": "Summary By RP Group (MTD).csv",
    "report_b": "MTD Summary By RP.csv",
    "report_c": "MTD Summary By RP Group.csv",
    # v6.0 §3 — replaced by the "Delivery Rating Report" Tableau report's
    # "MTD Courier Rating by District" sheet (see TABLEAU_TARGETS below);
    # filename mirrors the sheet name, same convention as every other entry
    # here. The old "LogisticsKPIReport"/"Delivery Rating" source is gone —
    # this file_key now points at the new report exclusively.
    "poor_rating": "MTD Courier Rating by District.csv",
    "rfid": "RP Breakdown (7days).csv",  # 保持無空格，依據您之前提供的檔名
    "gmv": "Sheet 1.csv",  # v3.0 §3 — fixed download filename, per spec

    # v4.0 §1/§2/§4 — "Delivery Dashboard" Tableau report replaces both the
    # old OIX-based order count AND the old "Rank_On Time" delay-rate source.
    # Two views feed these 5 files:
    #   DeliverySummary        (T-1 daily)   -> actual_delivery_timeslot, delay_early
    #   DeliverySummary-MTD    (month-to-date) -> actual_delivery_timeslot_mtd,
    #                                             delay_zone_type_mtd, mtd_delay_early_ontime
    "actual_delivery_timeslot": "Actual Delivery by Timeslot.csv",
    "delay_early": "Actual Delivery - Delay & Early %.csv",
    "actual_delivery_timeslot_mtd": "Actual Delivery by Timeslot - MTD - 10 Districts.csv",
    "delay_zone_type_mtd": "delay rate by zone type.csv",
    "mtd_delay_early_ontime": "MTD Actual Delivery - Delay, Early & On Time %.csv",

    # v8.0 §1 — Delivery Dashboard (Delivery Summary tab), sheet "Actual Delivery -
    # 10 Districts": source of the ODS order count (Column C "Parent Order #") and
    # ODS waybill count (Column E "No. of Waybill") on the rows headed "包派".
    "actual_delivery_10d": "Actual Delivery - 10 Districts.csv",
}

# v4.0 §2 — "Expected Timeslot w/ same day" raw values -> the short codes the
# dashboard's "Delay %" tab dropdown uses. "Total" (the daily/T-1 file's
# per-district summary row) maps to "Overall".
DELAY_TIMESLOT_MAP = {
    "1000-1400": "AM",
    "1400-1800": "PM",
    "1800-2200": "EV",
    "same day EV": "EV2",
    "Total": "Overall",
}

# v3.0 §4: the 6 position codes that get classified into Courier / Driver
# for the HKTV Manpower Distribution tab. Matched case-insensitively /
# whitespace-trimmed against the Staff List's "Position" column.
COURIER_POSITIONS = {"COURIER", "SENIOR COURIER"}
DRIVER_POSITIONS = {"DRIVER", "DRIVER II", "DRIVER AT", "DRIVER C"}
MANPOWER_GROUP_POSITIONS = {"courier": COURIER_POSITIONS, "driver": DRIVER_POSITIONS}

# v9.0 — Staff Allocation Cap (Suggest Headcount, Full Perspective Forecast
# tab): the same 4 buckets the Per-District Staff Plan table itself has
# (FT/PT Driver, FT/PT Courier), each mapped from the Staff List's "Position"
# column. FT_DRIVER/FT_COURIER reuse the existing DRIVER_POSITIONS/
# COURIER_POSITIONS sets above (same positions, same meaning) rather than
# duplicating them. Any Position not covered here — team leaders,
# supervisors, managers, transit-truck drivers, fleet/ops/HQ roles, etc. —
# is deliberately excluded from every cap: those staff aren't part of the
# Driver/Courier variable-cost pool Suggest Headcount allocates against, so
# counting them toward the cap would let a district with lots of
# supervisory headcount look like it has more allocatable Driver/Courier
# capacity than it actually does.
PT_DRIVER_POSITIONS = {"PART-TIME DRIVER"}
PT_COURIER_POSITIONS = {"PART-TIME COURIER"}
STAFF_CAP_BUCKET_POSITIONS = {
    "ftDriver": DRIVER_POSITIONS,
    "ptDriver": PT_DRIVER_POSITIONS,
    "ftCourier": COURIER_POSITIONS,
    "ptCourier": PT_COURIER_POSITIONS,
}
# Reverse lookup (Position -> bucket key), built once from the table above
# rather than kept as a second hand-written mapping that could drift out of
# sync with it.
STAFF_CAP_POSITION_TO_BUCKET = {
    pos: bucket for bucket, positions in STAFF_CAP_BUCKET_POSITIONS.items() for pos in positions
}

# v10.2 — Staff List "Department Code" (Column E) -> District, for the
# Staff Allocation Cap. Switched from an exact-code dict (each specific
# code like "LOGETH01"/"LOGETHP01" hand-enumerated) to the SAME
# prefix/startswith() matching ot_time_alert_workflow.py's
# dept_code_to_district() uses against its own DEPT_CODE_TO_DISTRICT list
# — same sheet, same Department Code column, so both scripts now classify
# a given code identically instead of running two independently-maintained
# mappings that could silently drift apart. This also closes the exact gap
# the old dict's own docstring flagged: a new sub-code the district office
# starts using (e.g. a third "LOGETH03") is picked up automatically here,
# where the old exact-dict approach would have silently dropped it into
# "unmapped" until someone noticed and added it by hand.
# Codes with no matching prefix (LOGCFM, LOGOPR, LOGPD*, LOGTD02, LOGMGT,
# LOGEXP, LOGPED, LOGMEN06, LOGMTM04, LOGFD02, LOGMET01, LOGMWT02, ...) are
# HQ / fleet-management / other non-district departments and are excluded
# from the cap entirely, the same way district_from_truck_no() excludes an
# unmatched truck number.
DEPT_CODE_TO_DISTRICT = [
    ("LOGETH", "ETH"),
    ("LOGETK", "ETK"),
    ("LOGETX", "ETX"),
    ("LOGST", "NT-ST"),
    ("LOGTM", "NT-TM"),
    ("LOGETN", "NT-TSM"),
    ("LOGWTW", "NT-TW"),
    ("LOGWTH", "WTH"),
    ("LOGWTK", "WTK"),
    ("LOGWTX", "WTX"),
]


def dept_code_to_district(dept_code):
    """v10.2 — identical logic to ot_time_alert_workflow.py's function of
    the same name: first prefix in DEPT_CODE_TO_DISTRICT that dept_code
    starts with wins (order matters for that reason, even though every
    prefix here happens to be distinct enough that it doesn't currently
    matter in practice). Returns "" — not None — for an unmapped code, to
    match the reference script's own return value exactly; callers here
    treat "" the same way they'd treat None (falsy)."""
    for prefix, district in DEPT_CODE_TO_DISTRICT:
        if dept_code.startswith(prefix):
            return district
    return ""

# v3.0 §4.1: staff in these positions are leads/supervisors/managers, not
# individual couriers/drivers — excluded from HKTV Staff Productivity's
# manpower denominator (they still appear in the raw OIX data, just not
# counted as "manpower" for that calculation).
LEADER_EXCLUDE_POSITIONS = {
    "TEAM LEADER ASSISTANT II",
    "ASSISTANT LOGISTIC OPERATIONS SUPERVISOR",
    "LOGISTIC OPERATIONS SUPERVISOR",
    "TEAM LEADER ASSISTANT I",
    "LOGISTIC OPERATIONS MANAGER",
    "ASSISTANT LOGISTIC OPERATIONS MANAGER",
    "SENIOR LOGISTIC OPERATIONS SUPERVISOR",
}

# 報表在 Tableau 彈出選單中的確切 Sheet 名稱，用於 Playwright 點擊
TABLEAU_TARGETS = [
    {
        "file_key": "report_a",
        "sheet_name": "Summary By RP Group (MTD)",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/MonthlyRP-3RReportA/MonthlyRP-3RSummary"
    },
    {
        "file_key": "report_b",
        "sheet_name": "MTD Summary By RP",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/MonthlyRP-CSCancelMerchantPaymentReportB/MonthlyRP-CSCancelMerchantPaymentSummary"
    },
    {
        "file_key": "report_c",
        "sheet_name": "MTD Summary By RP Group",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/MonthlyRP-CSCancelReportC/MonthlyRP-CSCancelSummary"
    },
    # v6.0 §3 — "Delivery Rating Report" Tableau report, "MTD Courier Rating
    # By District" sheet. Per spec this sheet is ALREADY pre-selected in the
    # Crosstab dialog when it opens, and must NOT be clicked again.
    #
    # v6.0 fix — the generic is_sheet_already_selected() check was not enough
    # here (pipeline_log.txt: both attempts ended in a 60s timeout waiting
    # for the download event, and no "already selected" line was logged).
    # The thumbnail list is a TOGGLE, so clicking the pre-selected sheet
    # DESELECTS it and the final Download then exports nothing. The old
    # sheet_name was spelled "...by District" while the sheet is titled
    # "...By District": the exact-title selectors are case-sensitive, so the
    # already-selected check found nothing and returned False, and the
    # case-INsensitive has-text fallback then matched the thumbnail and
    # clicked it. `preselected` makes the loop skip the thumbnail step
    # outright, whatever the check says (and the title is now spelled to
    # match, with the check itself made case-insensitive as well).
    {
        "file_key": "poor_rating",
        "sheet_name": "MTD Courier Rating By District",
        "preselected": True,
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/Deliveryratingreport-emaildata-Up/MTDCourierRatingByDistrict?:iid=1"
    },
    {
        "file_key": "rfid",
        "sheet_name": "RP Breakdown (7days)",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/RFIDReport_V3/RFIDReport-lastactiondate"
    },

    # v4.0 §1/§2 — Tableau Report "Delivery Dashboard" (Delivery Summary tab,
    # T-1 data). Both sheets live on the same view/URL; "Actual Delivery -
    # Delay & Early %" is the pre-selected sheet on that view (per spec, no
    # thumbnail click needed) — is_sheet_already_selected() inside the
    # download loop already handles that gracefully either way.
    {
        "file_key": "actual_delivery_timeslot",
        "sheet_name": "Actual Delivery by Timeslot",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary?:iid=1"
    },
    {
        "file_key": "delay_early",
        "sheet_name": "Actual Delivery - Delay & Early %",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary?:iid=1"
    },

    # v4.0 §3/§4 — same Tableau Report, "Delivery Summary - MTD" tab.
    {
        "file_key": "delay_zone_type_mtd",
        "sheet_name": "delay rate by zone type",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary-MTD?:iid=1"
    },
    {
        "file_key": "actual_delivery_timeslot_mtd",
        "sheet_name": "Actual Delivery by Timeslot - MTD - 10 Districts",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary-MTD?:iid=1"
    },
    {
        "file_key": "mtd_delay_early_ontime",
        "sheet_name": "MTD Actual Delivery - Delay, Early & On Time %",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary-MTD?:iid=1"
    },

    # v8.0 §1 — Delivery Dashboard > Delivery Summary tab, sheet "Actual Delivery -
    # 10 Districts" (ODS order / waybill counts).
    #
    # v8.1 fix — this sheet is NOT pre-selected in the Crosstab dialog. The v8.0
    # "preselected": True flag made the loop skip the thumbnail step, so Tableau
    # exported whatever sheet was selected by default ('Actual Delivery by TimeSlot
    # - Others' in pipeline_log.txt) under this filename, and the parser then
    # found no '包派' row and fell back to the OIX figures. The flag is removed, so
    # the normal sheet-select logic runs again: skip the click only if the sheet
    # is genuinely already aria-selected (is_sheet_already_selected), otherwise
    # click its thumbnail (scrolled into view) before choosing CSV.
    {
        "file_key": "actual_delivery_10d",
        "sheet_name": "Actual Delivery - 10 Districts",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/DeliveryDashboard/DeliverySummary?:iid=2"
    },
]

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]
DISTRICT_MATCH_ORDER = sorted(DISTRICTS + ["NT-YT"], key=len, reverse=True)
MONTH_RE = re.compile(r"^\d{4}年\d{1,2}月$")

# =============================================================================
# Business targets, per Target__Other_Aspects_.xlsx (received 2026-09-15).
# Each matrix's "target" block is written into data.json alongside "actual"/
# "forecast" so the dashboard's getTargetOverall()/getTargetForDistrict() pick
# it up automatically (see index.html) instead of falling back to the
# placeholder flat number in the page's MATRICES config.
#
# NOTE: the source workbook only gave per-district ceilings for Missing &
# Lost Amount and RFID Missing Tote; Customer Rating (Poor Rating %) and
# (MTD) Delay Rate were confirmed as flat, network-wide targets (no
# per-district breakdown) rather than per-district ceilings.
# =============================================================================
DELAY_RATE_TARGET = {"overall": 4.0}     # MTD Delay Rate target, confirmed flat (no per-district split)
POOR_RATING_TARGET = {"overall": 0.07}   # Customer Rating (Poor Rating %) target, confirmed flat (no per-district split)
MISSING_LOST_TARGETS = {  # "<=" row, Missing & Lost sheet
    "overall": 400000,
    "districts": {
        "ETH": 41543.52, "ETK": 39507.93, "ETX": 36261.54, "NT-ST": 35860.28,
        "NT-TM": 35457.72, "NT-TSM": 39123.05, "NT-TW": 36312.45, "WTH": 48711.29,
        "WTK": 31378.04, "WTX": 33844.18,
    },
}
RFID_TOTE_TARGETS = {  # "<=" row, RFID Tote sheet
    "overall": 146,
    "districts": {
        "ETH": 16, "ETK": 15, "ETX": 14, "NT-ST": 14, "NT-TM": 14,
        "NT-TSM": 15, "NT-TW": 14, "WTH": 19, "WTK": 12, "WTX": 13,
    },
}


# =============================================================================
# 2. 共用 Helper 函數
# =============================================================================
def today_hkt():
    return dt.date.today()

def normalize_district_code(code):
    # v3.1 (Sept 2026): district grouping changed at the source — NT-YT is
    # now folded into WTH (it used to be folded into NT-TW). Fires when the
    # raw code is the standalone string "NT-YT" — e.g. GMV's per-column
    # headers, or a party name starting with "NT-YT" in report_a/b/c.
    return "WTH" if code == "NT-YT" else code

def district_from_party_name(name):
    if not isinstance(name, str): return None
    for code in DISTRICT_MATCH_ORDER:
        if name.startswith(code):
            return normalize_district_code(code)
    return None

# v3.1 (Sept 2026): unlike the party-name/column-header case above, the
# Delivery Rating (and possibly Rank_On Time) crosstab doesn't hand us a bare
# "NT-YT" row to fold — Tableau's own district dimension now emits the
# already-merged group as a single combined label, "NT-YT & WTH". Confirmed
# from the Sept CSV export: row that used to read "WTH" now reads
# "NT-YT & WTH" outright. normalize_district_code() won't catch this (the
# code here isn't "NT-YT", it's the whole combined string), so map the
# display label itself.
DISTRICT_LABEL_ALIASES = {
    "NT-YT & WTH": "WTH",
}

def normalize_district_label(label):
    label = str(label).strip()
    return DISTRICT_LABEL_ALIASES.get(label, label)

def col(letter):
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1

# Replicates the Excel IFS() formula on 送貨車號 (truck number) EXACTLY in
# order — IFS returns the first TRUE condition, so order matters. Excel's
# SEARCH() is case-insensitive, matched here via .upper().
TRUCK_NO_RULES = [
    ("ETK", "ETK"), ("將軍澳", "ETK"),
    ("CH", "WTH"), ("CHA", "WTH"),
    ("CTK", "WTK"), ("WK", "WTK"),
    ("WX", "WTX"),
    ("CTW", "NT-TW"),
    ("ENH", "ETH"),
    ("KTX", "ETX"), ("KX", "ETX"),
    ("NST", "NT-ST"),
    ("NTM", "NT-TM"),
    ("ZTS", "NT-TSM"),
    ("WTW", "NT-TW"),
]


def district_from_truck_no(truck_no):
    if not isinstance(truck_no, str):
        return None
    upper = truck_no.upper()
    for needle, code in TRUCK_NO_RULES:
        if needle.upper() in upper:
            return code
    return None

# =============================================================================
# 3. PLAYWRIGHT TABLEAU 下載邏輯
# =============================================================================
def fast_click(page, selectors, timeout_ms=5000):
    """用於「視窗期很短」的元素（例如 Download 選單、Crosstab 選項）。

    smart_click 每輪只重新檢查一次、間隔 time.sleep(1)——如果選單只
    存在大約 1 秒，這個間隔本身就跟目標視窗一樣長，變成純粹賭時機。
    這裡改用 Playwright locator.click(timeout=...) 內建的 auto-wait，
    它是每 ~100ms 就重新嘗試一次，抓到瞬間出現/消失的元素的機率高得多。
    """
    viz_frame = page.frame_locator('iframe[title="Data Visualisation"]')
    candidates = [viz_frame] + [page.main_frame] + list(page.frames)
    per_target_timeout = max(300, timeout_ms // max(1, len(selectors)))
    end_time = time.time() + (timeout_ms / 1000.0)
    while time.time() < end_time:
        for target in candidates:
            for sel in selectors:
                try:
                    loc = target.locator(sel).first
                    loc.click(timeout=per_target_timeout)
                    return True
                except Exception:
                    continue
    return False


def smart_click(page, selectors, timeout_sec=15):
    """依序嘗試點擊 selectors 裡的元素。

    Tableau 的實際視覺化內容 (包含整個工具列、Download 按鈕、Crosstab 選單、
    Sheet 縮圖、CSV 對話框) 是包在一個 <iframe id="viz" tb-test-id="viz">
    裡面的，不是直接在主頁面上。優先用 page.frame_locator() 直接鎖定這個
    iframe —— 這是 Playwright 官方建議處理 iframe 的方式，比手動遍歷
    page.frames 更穩定（會自動等待 iframe 附加、內容就緒）。
    如果找不到這個 iframe 或元素不在裡面，才退回原本「掃描全部 frame」
    的作法當備援，涵蓋 iframe id 未來改變或多層巢狀的情況。
    """
    start_time = time.time()
    # 注意：id="viz" / tb-test-id="viz" 其實是包住 iframe 的外層 <div> 上的屬性，
    # 不是 iframe 標籤本身！iframe 標籤本身沒有 id，只有 title="Data Visualisation"
    # 是穩定存在於 iframe 標籤上的屬性，所以改用這個來鎖定正確的 frame。
    viz_frame = page.frame_locator('iframe[title="Data Visualisation"]')
    while time.time() - start_time < timeout_sec:
        # 優先：直接鎖定已知的 viz iframe
        for sel in selectors:
            try:
                loc = viz_frame.locator(sel).first
                if loc.is_visible():
                    loc.click(force=True)
                    return True
            except Exception:
                continue
        # 備援：掃描主頁面 + 所有 frame（含巢狀）
        all_frames = [page.main_frame] + page.frames
        for frame in all_frames:
            for sel in selectors:
                try:
                    loc = frame.locator(sel).first
                    if loc.is_visible():
                        loc.click(force=True)
                        return True
                except Exception:
                    continue
        time.sleep(1)
    return False


def is_sheet_already_selected(page, sheet_name):
    """檢查目標 sheet 縮圖現在是否已經是「已選取」狀態 (aria-selected="true")。

    根本原因（感謝實測回報確認）：這個縮圖清單的選取是「切換式」
    (toggle) 而非「單選式」。如果 Crosstab 對話框打開時，目標 sheet
    剛好已經是預設選取狀態（例如該報表本來就只有/預設停留在這個
    sheet），我們的腳本又照原本邏輯點它一次，會把它「取消選取」，
    導致後面按下 Download 時沒有任何有效 sheet 被選定 —— 按鈕點得到、
    click 事件也正常觸發，但 Tableau 端不會真的產生檔案，所以症狀是
    expect_download 累積 60 秒逾時，而不是找不到按鈕。
    這正是 Delivery Rating 會失敗、但 Summary By RP Group 不會失敗的差異：
    後者預設選取的縮圖跟目標 sheet 不同名，點擊是「選取」而非「取消」。

    回傳 True 時，呼叫端應該跳過點擊，直接視為已選定。
    """
    viz_frame = page.frame_locator('iframe[title="Data Visualisation"]')
    candidates = [viz_frame, page.main_frame] + list(page.frames)
    selectors = [
        f'[role="option"][title="{sheet_name}"]',
        f'[data-tb-test-id^="sheet-thumbnail"][title="{sheet_name}"]',
        f'div[title="{sheet_name}"][aria-selected]',
    ]
    for target in candidates:
        for sel in selectors:
            try:
                loc = target.locator(sel).first
                if loc.count() == 0:
                    continue
                state = loc.get_attribute("aria-selected")
                if state is not None:
                    return state == "true"
            except Exception:
                continue
    # v6.0 fix — 上面的 selector 是「title 完全相等」(大小寫敏感)。如果設定檔裡的
    # sheet_name 跟 Tableau 縮圖的 title 只差大小寫/多餘空白（例如 "by" vs "By"），
    # 上面全部會找不到、誤判成「未選取」，接著後面 has-text 備援（大小寫不敏感）
    # 卻點得到 → 把已選取的縮圖切換成取消選取。這裡補一次大小寫不敏感的比對。
    want = " ".join(str(sheet_name).split()).lower()
    for target in candidates:
        try:
            opts = target.locator('[role="option"]')
            for i in range(min(opts.count(), 40)):
                opt = opts.nth(i)
                title = " ".join((opt.get_attribute("title") or "").split()).lower()
                if title == want:
                    state = opt.get_attribute("aria-selected")
                    if state is not None:
                        return state == "true"
        except Exception:
            continue
    return False  # 找不到就當作未選取，走原本點擊流程（不影響原本能成功的報表）


def log_selected_sheet_thumbnails(page):
    """v6.0 — 純除錯用：印出 Crosstab 對話框裡目前「已選取」的工作表縮圖 title。
    用在「預先選取、不點擊」的報表，萬一之後又下載失敗，log 裡就看得到當時
    Tableau 實際預選的是哪一個 sheet（不會點擊任何東西）。"""
    viz_frame = page.frame_locator('iframe[title="Data Visualisation"]')
    seen = []
    for target in [viz_frame, page.main_frame] + list(page.frames):
        try:
            opts = target.locator('[role="option"]')
            for i in range(min(opts.count(), 40)):
                opt = opts.nth(i)
                if opt.get_attribute("aria-selected") == "true":
                    title = (opt.get_attribute("title") or "").strip()
                    if title and title not in seen:
                        seen.append(title)
        except Exception:
            continue
    print(f"  🔎 目前已選取的工作表縮圖: {seen if seen else '（偵測不到縮圖清單，可能此報表本來就只有單一工作表）'}")


def smart_click_with_scroll(page, selectors, timeout_sec=15):
    """跟 smart_click 幾乎一樣，但用在「元素可能在可捲動清單中、不在目前可視
    範圍內」的情況 —— 例如 Sheet 縮圖清單 (role="listbox" ... scroll)。
    force=True 點擊會跳過 Playwright 內建的自動捲動，所以這裡在點擊前先
    明確呼叫 scroll_into_view_if_needed()，確保目標真的被捲動到畫面上再點。
    """
    start_time = time.time()
    viz_frame = page.frame_locator('iframe[title="Data Visualisation"]')
    while time.time() - start_time < timeout_sec:
        for sel in selectors:
            try:
                loc = viz_frame.locator(sel).first
                loc.scroll_into_view_if_needed(timeout=3000)
                if loc.is_visible():
                    loc.click(force=True)
                    return True
            except Exception:
                continue
        all_frames = [page.main_frame] + page.frames
        for frame in all_frames:
            for sel in selectors:
                try:
                    loc = frame.locator(sel).first
                    loc.scroll_into_view_if_needed(timeout=3000)
                    if loc.is_visible():
                        loc.click(force=True)
                        return True
                except Exception:
                    continue
        time.sleep(1)
    return False

def fetch_tableau_reports():
    """使用 Playwright 自動登入 Tableau 並下載所有目標 Crosstab CSV"""
    print("🚀 啟動 Tableau 自動化下載程序...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--disable-popup-blocking"])
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        # 加寬視窗 — 1280 太窄，Tableau 工具列在窄螢幕下會自動收起部分圖示
        # (包含 Download)，即使按鈕還在 DOM 裡也不會顯示/可點擊。
        page.set_viewport_size({"width": 1920, "height": 1080})

        # 1. 登入 Tableau
        print("  🌐 導航至 Tableau 登入頁面...")
        page.goto(TABLEAU_URL)
        page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
        page.locator("input[type='text'], input[name='username']").first.fill(TABLEAU_USER)
        page.locator("input[type='password'], input[name='password']").first.fill(TABLEAU_PASS)
        page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()

        # 登入後 Tableau 會自己非同步跳轉到預設頁面
        # (您提到會先跳到 /#/user/local/Warehouse/settings)。
        # 如果我們在這個跳轉還沒完成前就急著呼叫 page.goto(report_url)，
        # 會有 race condition：我們的導航先發生，接著 Tableau 自己的跳轉
        # 才完成，結果把畫面蓋回 settings 頁，後續怎麼點都點不到東西。
        # 所以先明確等待、讓它先跳轉完、畫面穩定下來，才開始逐一導航到報表。
        print("  ⏳ 等待登入後的跳轉完成...")
        page.wait_for_load_state("networkidle", timeout=60000)
        try:
            page.wait_for_url("**/#/user/**", timeout=30000)
            print(f"  ✅ 已到達登入後預設頁面: {page.url}")
        except Exception:
            print(f"  ⚠ 未偵測到預期的 /#/user/ 跳轉，目前網址: {page.url}（仍會繼續嘗試導航）")
        time.sleep(3)

        # 2+3. 每個報表：導航 → 等待載入 → 下載，合併成單一迴圈
        #      (原本分成兩個迴圈：第一個迴圈把全部 6 個網址都導航過一遍，
        #       第二個迴圈才嘗試點擊 Download —— 但第二個迴圈完全沒有再次
        #       呼叫 page.goto()，所以實際上是在「上一輪導航留下的最後一頁」
        #       上操作，而不是對應到當下這個報表。合併成一個迴圈，確保每次
        #       點擊 Download 之前，頁面一定是剛導航到的那個正確報表。)
        for target in TABLEAU_TARGETS:
            sheet_name = target["sheet_name"]
            report_url = target["url"]
            target_filename = REPORT_FILES[target["file_key"]]
            target_filepath = os.path.join(REPORT_FOLDER, target_filename)

            # v3.0 §3 fix: the whole download sequence (Download -> Crosstab ->
            # sheet select -> CSV -> confirm) is a chain of Tableau UI clicks
            # that, per the actual pipeline_log.txt from several trial runs,
            # fails intermittently at basically any step and for any report —
            # not consistently the same report twice, and not consistently the
            # same step (Download button not found, Crosstab not found, sheet
            # toggle silently deselecting, confirm button not found, or the
            # download event itself never firing). Since none of that points
            # to one specific selector being wrong, wrap one full retry around
            # the whole sequence (fresh navigate + reload) instead of giving
            # up on the first miss — this is the same shape of retry the
            # "stuck on /#/user/ settings page" case below already uses.
            def _attempt_download():
                print(f"\n🌐 Opening Tableau Report:")
                print(f"   {sheet_name}")
                print(f"   {report_url}")

                page.goto(report_url)
                try:
                    page.wait_for_load_state("networkidle", timeout=30000)
                except Exception:
                    pass  # 同上，Tableau 背景流量常讓 networkidle 逾時，不視為錯誤
                time.sleep(10)
                page.wait_for_timeout(5000)

                # 強制整頁重新載入 (reload)，而非只依賴 page.goto() 的 hash 導航。
                # 原因：這個網址只有 # 後面的部分不同 (同一個 origin/path)，瀏覽器會
                # 把它當成「同文件」的輕量導航，不會真的重新載入整個頁面 —
                # 畫面內容看起來雖然正確 (Tableau 用 JS 更新畫面)，但工具列
                # (包含 Download 按鈕) 的事件綁定經常沒有隨之重新初始化，導致
                # 按鈕看得到卻點不動/找不到。強制 reload() 讓 Tableau 針對這個
                # 特定 view 做一次「乾淨」的完整啟動，工具列才會確實可用。
                #
                # 注意：用 "load" 而非 "networkidle" —— Tableau 的 view 會持續有
                # 背景輪詢/websocket 流量，網路幾乎不會真正「idle」，用 networkidle
                # 當作 reload() 的等待條件很容易 30 秒逾時。改用 "load"（頁面的
                # load 事件，通常幾秒內就會觸發），實際「畫面真的準備好了沒」
                # 交給後面自己的 sleep + wait_for_selector 判斷即可。
                print(f"  🔄 強制重新載入頁面，確保工具列正確初始化...")
                try:
                    page.reload(wait_until="load", timeout=45000)
                except Exception as e:
                    print(f"  ⚠ reload() 等待逾時或發生問題（{e}），仍繼續嘗試後續步驟...")
                time.sleep(8)
                page.wait_for_timeout(3000)

                # 保護：如果被 Tableau 自己的跳轉蓋回 /#/user/ 設定頁
                # (目前只在第一個報表看過，但保留重試邏輯以防其他報表也偶發發生)，
                # 就再導航一次。最多重試 2 次，避免無限迴圈。
                retry_count = 0
                while "/#/user/" in page.url and retry_count < 2:
                    print(f"  ⚠ 目前網址被導向設定頁 ({page.url})，重新導航一次...")
                    retry_count += 1
                    page.goto(report_url)
                    try:
                        page.wait_for_load_state("networkidle", timeout=30000)
                    except Exception:
                        pass
                    time.sleep(8)
                    page.wait_for_timeout(3000)
                    try:
                        page.reload(wait_until="load", timeout=45000)
                    except Exception as e:
                        print(f"  ⚠ reload() 等待逾時或發生問題（{e}），仍繼續嘗試後續步驟...")
                    time.sleep(8)
                    page.wait_for_timeout(3000)
                if "/#/user/" in page.url:
                    shot_path = os.path.join(REPORT_FOLDER, f"_debug_{target['file_key']}_stuck_on_settings.png")
                    try:
                        page.screenshot(path=shot_path, full_page=True)
                    except Exception:
                        pass
                    print(f"  ❌ 重試 {retry_count} 次後仍停留在設定頁，跳過 {sheet_name}（截圖: {shot_path}）")
                    return False

                print(f"  ⬇️ 正在下載報表: {sheet_name} ...")

                # Download 按鈕本身通常撐得住，用一般 smart_click 找就好。
                # 真正容易「一閃即逝」的是點下去之後彈出的選單/選項，所以那些
                # 一律改用 fast_click（Playwright 內建 ~100ms 頻率的 auto-wait），
                # 而不是 smart_click 自訂的 1 秒間隔 polling。
                dl_selectors = ["#download", '[aria-label="Download"]', "button:has-text('Download')"]
                if not fast_click(page, dl_selectors, 4000):
                    if not smart_click(page, dl_selectors, 15):
                        print(f"  ❌ smart_click 也無法點擊 Download 按鈕，跳過 {sheet_name}")
                        return False

                # 點擊 Crosstab —— 這是「選單彈出後一閃即逝」的關鍵一步，
                # 緊接著上一個點擊立刻嘗試，中間不要 sleep，把握選單開啟的短暫視窗。
                crosstab_selectors = [
                    "#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
                    "#viz-viewer-toolbar-download-menu div:nth-of-type(3) span",
                    "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]",
                    "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]/div/div/span[2]",
                    '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
                    "span:has-text('Crosstab')",
                    "text='Crosstab'"
                ]
                if not fast_click(page, crosstab_selectors, 4000):
                    # fast_click 沒抓到 → 選單可能還沒完全跳出來，補一次完整的
                    # smart_click 當備援（涵蓋選單延遲較久才出現的情況）。
                    if not smart_click(page, crosstab_selectors, 10):
                        print(f"  ❌ 找不到 Crosstab 選項，跳過 {sheet_name}")
                        return False
                time.sleep(2)

                # 選擇工作表 (Sheet) — Tableau 這個版本用「縮圖卡片」(role="option")
                # 而非傳統下拉選單，所以直接用 title 屬性比對卡片，不需要先點開下拉選單。
                # 對應您提供的 HTML：<div role="option" title="Summary By RP Group (MTD)"
                #   data-tb-test-id="sheet-thumbnail-2" aria-selected="true">
                #
                # 注意：這個縮圖清單本身是可捲動的 (role="listbox" ... scroll)，如果目標
                # 縮圖排在很後面 (例如 RFID 的 "RP Breakdown (7days)" 排在第 11 個)，
                # 用 smart_click 的 force=True 點擊會跳過 Playwright 內建的
                # 「自動捲動到可視範圍」機制，導致即使選到正確元素也點不中。
                # 所以這裡改用專門的 scroll_into_view_if_needed() 再點擊。
                sheet_option_selectors = [
                    f'[role="option"][title="{sheet_name}"]',
                    f'[data-tb-test-id^="sheet-thumbnail"][title="{sheet_name}"]',
                    f'div[title="{sheet_name}"][aria-selected]',
                    # 備援：文字比對（含大小寫/多餘空白容錯）
                    f"[role='option']:has-text('{sheet_name}')",
                ]
                if target.get("preselected"):
                    # v6.0 fix — 依規格這個 sheet 在開啟 Crosstab 時已經預先選取，完全不碰縮圖。
                    print(f"  ℹ️ '{sheet_name}' 依規格已預先選取，不再點擊工作表縮圖（縮圖是切換式，再點一次會變成取消選取）")
                    log_selected_sheet_thumbnails(page)
                elif is_sheet_already_selected(page, sheet_name):
                    print(f"  ℹ️ 工作表縮圖 '{sheet_name}' 已經是選取狀態，跳過點擊（避免切換式選取被點成取消選取）")
                elif not fast_click(page, sheet_option_selectors, 3000):
                    if not smart_click_with_scroll(page, sheet_option_selectors, 15):
                        print(f"⚠ 找不到工作表縮圖 '{sheet_name}'")
                        print(f"   請確認 sheet_name 拼字是否與縮圖 title 完全一致（含空格/大小寫）。")
                time.sleep(1)

                # 選擇 CSV 格式 —— 套用與 Download/Crosstab 相同的「成功組合」：
                # fast_click（Playwright 內建高頻 auto-wait）優先，找不到才退回
                # smart_click 的完整跨 frame 掃描備援。
                csv_selectors = [
                    "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
                    "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2)",
                    "xpath=//*[@id='export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id']/div/div[2]/div[2]/label[2]",
                    "label:has-text('CSV')",
                    "text='CSV'"
                ]
                if not fast_click(page, csv_selectors, 2000):
                    smart_click(page, csv_selectors, 5)
                time.sleep(1)

                # 點擊 Download 確認並攔截檔案
                confirm_selectors = [
                    "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id button",
                    "button[aria-label='Download Crosstab']",
                    "button[aria-label='Download']",
                    'button[data-tb-test-id="export-crosstab-export-Button"]'
                ]

                try:
                    with page.expect_download(timeout=60000) as download_info:

                        if not fast_click(page, confirm_selectors, 3000):
                            if not smart_click(page, confirm_selectors, 30):
                                raise TimeoutError(
                                    f"Unable to click download button for {sheet_name}"
                                )

                    download = download_info.value
                    download.save_as(target_filepath)
                    print(f"  ✅ 成功儲存: {target_filename}")
                    time.sleep(2) # 緩衝時間
                    return True
                except Exception as e:
                    print(f"  ❌ 下載 {sheet_name} 失敗: {e}")
                    # v6.0 — 存一張截圖，之後能直接看到失敗當下 Crosstab 對話框的狀態
                    # (哪個 sheet 有被選取、CSV 有沒有選到)，不用再靠推測。
                    try:
                        shot_path = os.path.join(REPORT_FOLDER, f"_debug_{target['file_key']}_download_failed.png")
                        page.screenshot(path=shot_path, full_page=True)
                        print(f"  📸 失敗當下截圖: {shot_path}")
                    except Exception:
                        pass
                    return False

            succeeded = False
            for attempt in (1, 2):
                if attempt == 2:
                    print(f"  🔁 重試 {sheet_name}（第 2 次嘗試）...")
                if _attempt_download():
                    succeeded = True
                    break
            if not succeeded:
                print(f"  ❌ {sheet_name} 兩次嘗試皆失敗，放棄此報表。")

        context.close()
        browser.close()
        print("🎉 Tableau 報表下載完畢！\n")


def fetch_gmv_report():
    """v3.0 §3 — downloads 'Sheet 1.csv' (GMV) from the SEPARATE 'LOG GMV
    for AI Fetching' Tableau account. Its own browser/login, independent of
    fetch_tableau_reports() above, since it's a different account entirely.

    Per spec this sheet is "pre-click" — i.e. already sitting on the right
    view/sheet by default — so this skips the sheet-thumbnail selection step
    that fetch_tableau_reports() needs for its 6 reports, and just does
    Download -> Crosstab -> CSV -> confirm.
    """
    if not TABLEAU_GMV_USER or not TABLEAU_GMV_PASS:
        print("  ⚠️ TABLEAU_GMV_USER / TABLEAU_GMV_PASS not set — skipping GMV download. "
              "Set these in .env (separate account from TABLEAU_USER/PASS) to enable it.")
        return

    print("🚀 啟動 GMV Tableau 自動化下載程序 (獨立帳號)...")
    target_filepath = os.path.join(REPORT_FOLDER, REPORT_FILES["gmv"])

    # v3.0 §3 fix: same "log in via a redirect straight to the workbook's
    # views listing" approach as oix_returning_waybill.py's download_tableau_raw(),
    # instead of logging in then goto()-ing the deep view URL directly. The
    # GMV account doesn't land on /#/user/ after login the way the main
    # account does (it lands on /#/explore), so jumping straight to a
    # deep-linked view hash right after login was racing the SPA's own
    # routing/toolbar init — Download would click something, but the
    # Crosstab menu item never actually rendered.
    login_url = TABLEAU_GMV_URL
    if TABLEAU_GMV_WORKBOOK_ID:
        login_url = (f"https://inhouse-analytics.hktv.com.hk/#/signin"
                     f"?redirect=%2Fworkbooks%2F{TABLEAU_GMV_WORKBOOK_ID}%2Fviews")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--disable-popup-blocking"])
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        print("  🌐 導航至 GMV Tableau 登入頁面...")
        page.goto(login_url)
        page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
        page.locator("input[type='text'], input[name='username']").first.fill(TABLEAU_GMV_USER)
        page.locator("input[type='password'], input[name='password']").first.fill(TABLEAU_GMV_PASS)
        page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()

        if TABLEAU_GMV_WORKBOOK_ID:
            print("  📄 等待並點擊 Sheet1...")
            loc1 = page.locator('[aria-label="Sheet1"][role="link"]')
            loc2 = page.get_by_text("Sheet1", exact=True)
            loc3 = page.get_by_text("Sheet 1", exact=True)
            sheet_link = loc1.or_(loc2).or_(loc3)
            sheet_link.first.wait_for(state="visible", timeout=60000)
            sheet_link.first.click()
            time.sleep(3)
        else:
            # Fallback: old direct-view approach. Kept only for the case
            # TABLEAU_GMV_WORKBOOK_ID hasn't been set yet — this is the path
            # that was producing "找不到 Crosstab 選項".
            print("  ⚠ TABLEAU_GMV_WORKBOOK_ID not set — using the older direct-view "
                  "navigation, which is the flow that was failing. Set "
                  "TABLEAU_GMV_WORKBOOK_ID in .env to use the more reliable path.")
            print("  ⏳ 等待登入後的跳轉完成...")
            page.wait_for_load_state("networkidle", timeout=60000)
            try:
                page.wait_for_url("**/#/user/**", timeout=30000)
            except Exception:
                print(f"  ⚠ 未偵測到預期的 /#/user/ 跳轉，目前網址: {page.url}（仍會繼續嘗試導航）")
            time.sleep(3)

            print(f"\n🌐 Opening GMV Tableau Report: {TABLEAU_GMV_DASHBOARD_URL}")
            page.goto(TABLEAU_GMV_DASHBOARD_URL)
            try:
                page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
            time.sleep(10)
            page.wait_for_timeout(5000)
            try:
                page.reload(wait_until="load", timeout=45000)
            except Exception as e:
                print(f"  ⚠ reload() 等待逾時或發生問題（{e}），仍繼續嘗試後續步驟...")
            time.sleep(8)
            page.wait_for_timeout(3000)

        print("  ⬇️ 正在下載 GMV 報表...")
        dl_selectors = ["#download", '[aria-label="Download"]', "button:has-text('Download')"]
        if not fast_click(page, dl_selectors, 4000):
            if not smart_click(page, dl_selectors, 15):
                print("  ❌ 找不到 Download 按鈕，GMV 下載中止")
                context.close(); browser.close()
                return

        crosstab_selectors = [
            "#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
            "#viz-viewer-toolbar-download-menu div:nth-of-type(3) span",
            "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]",
            "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]/div/div/span[2]",
            '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
            "span:has-text('Crosstab')",
            "text='Crosstab'"
        ]
        if not fast_click(page, crosstab_selectors, 4000):
            if not smart_click(page, crosstab_selectors, 10):
                print("  ❌ 找不到 Crosstab 選項，GMV 下載中止")
                context.close(); browser.close()
                return
        time.sleep(2)

        # Sheet is pre-selected per spec — no thumbnail click needed. If a
        # future export ever isn't pre-selected, is_sheet_already_selected()
        # returning False just means we fall through to the CSV step anyway
        # (same graceful-skip behavior as the 6-report loop).
        if not is_sheet_already_selected(page, "Sheet1"):
            sheet_option_selectors = [
                '[role="option"][title="Sheet1"]',
                '[data-tb-test-id^="sheet-thumbnail"][title="Sheet1"]',
                "[role='option']:has-text('Sheet1')",
            ]
            fast_click(page, sheet_option_selectors, 2000)
        time.sleep(1)

        csv_selectors = [
            "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
            "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2)",
            "xpath=//*[@id='export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id']/div/div[2]/div[2]/label[2]",
            "label:has-text('CSV')",
            "text='CSV'"
        ]
        if not fast_click(page, csv_selectors, 2000):
            smart_click(page, csv_selectors, 5)
        time.sleep(1)

        confirm_selectors = [
            "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id button",
            "button[aria-label='Download Crosstab']",
            "button[aria-label='Download']",
            'button[data-tb-test-id="export-crosstab-export-Button"]'
        ]
        try:
            with page.expect_download(timeout=60000) as download_info:
                if not fast_click(page, confirm_selectors, 3000):
                    if not smart_click(page, confirm_selectors, 30):
                        raise TimeoutError("Unable to click download button for GMV Sheet1")
            download = download_info.value
            download.save_as(target_filepath)
            print(f"  ✅ 成功儲存: {REPORT_FILES['gmv']}")
        except Exception as e:
            print(f"  ❌ 下載 GMV 報表失敗: {e}")

        context.close()
        browser.close()
        print("🎉 GMV 報表下載完畢！\n")

# =============================================================================
# 4. 數據處理邏輯 (Tableau CSV 解析)
# =============================================================================
def parse_money(val):
    if pd.isna(val): return 0.0
    s = str(val).replace("$", "").replace(",", "").strip()
    return float(s) if s and s.lower() != "nan" else 0.0

def parse_number(val):
    if pd.isna(val): return 0.0
    s = str(val).replace(",", "").strip()
    return float(s) if s and s.lower() != "nan" else 0.0

def parse_percent(val, decimals=2):
    # v4.0.1 §1 correction — round to 2dp at the source (Tableau exports these
    # as e.g. "7.3%"), so every Delay/Early/On Time % figure written to
    # data.json / delay_history.json is consistently 2 decimal places
    # instead of inheriting whatever precision the source happened to have.
    # v6.0 §3 — `decimals` lets a caller ask for a different precision
    # (Poor Rating % now needs 3dp, per spec) without duplicating this
    # function; every existing caller keeps the default 2dp behaviour.
    if pd.isna(val): return None
    s = str(val).replace("%", "").strip()
    return round(float(s), decimals) if s and s.lower() != "nan" else None

def load_crosstab(filename, marker_col0_values):
    path = os.path.join(REPORT_FOLDER, filename)
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    for i in range(min(10, len(raw))):
        if str(raw.iloc[i, 0]).strip() in marker_col0_values:
            return raw.iloc[i + 1:].reset_index(drop=True)
    raise ValueError(f"無法在 {path} 中找到標頭 {marker_col0_values}。")

def split_district_and_other(df, rp_group_idx, final_rp_idx, total_idx, allowed_rp_groups, month_idx=None):
    per_district = {d: 0.0 for d in DISTRICTS}
    other_total = 0.0
    for _, row in df.iterrows():
        rp_group = str(row.iloc[rp_group_idx]).strip()
        if rp_group not in allowed_rp_groups: continue
        
        final_rp = str(row.iloc[final_rp_idx]).strip()
        if final_rp in ("Total", "nan", "NaN", "None"): continue
        
        # Report A 專用過濾：僅計算 YYYY年MM月 的 Row
        if month_idx is not None:
            if not MONTH_RE.match(str(row.iloc[month_idx]).strip()):
                continue
                
        val = parse_money(row.iloc[total_idx])
        d = district_from_party_name(final_rp)
        if d:
            per_district[d] += val
        else:
            other_total += val
    return per_district, round(other_total, 2)

def parse_report_a():
    df = load_crosstab(REPORT_FILES["report_a"], {"RP Group"})
    per_d, other = split_district_and_other(df, 0, 1, 6, {"Bert (Log)"}, month_idx=2)
    return {"overall": round(sum(per_d.values()) + other, 2), "districts": {d: round(v, 2) for d, v in per_d.items()}}

def parse_report_b():
    df = load_crosstab(REPORT_FILES["report_b"], {"RP Group"})
    per_d, other = split_district_and_other(df, 0, 1, 5, {"Bert (Log)", "Bert (Log) - DAMUP"})
    return {"overall": round(sum(per_d.values()) + other, 2), "districts": {d: round(v, 2) for d, v in per_d.items()}}

def parse_report_c():
    df = load_crosstab(REPORT_FILES["report_c"], {"RP Group"})
    per_d, other = split_district_and_other(df, 0, 1, 5, {"Bert (Log)"})
    return {"overall": round(sum(per_d.values()) + other, 2), "districts": {d: round(v, 2) for d, v in per_d.items()}}

def parse_actual_delivery_timeslot(file_key):
    """v4.0 §1/§4 — parses 'Actual Delivery by Timeslot(.csv)' (T-1 daily) or
    'Actual Delivery by Timeslot - MTD - 10 Districts.csv' (month-to-date) —
    both share the exact same 2-row-header crosstab shape:
      row 0: 'Parent Order #' / 'No. of Waybill' (repeated) / blank
      row 1: timeslot label ('1000-1400' etc.) or 'Total'
      row 2+: one row per district (col 0), 'Grand Total' last.
    We only need the two "Total" columns (overall parent-order count and
    overall waybill count per district) — per spec: "Row 1 Column Header =
    'Parent Order' and Row 2 Column Header = 'Total' is the total parent
    order count in district basis" (and same for 'No. of Waybill').
    Returns {"overall": {"order":..,"waybill":..}, "districts": {d: {...}}}.
    """
    path = os.path.join(REPORT_FOLDER, REPORT_FILES[file_key])
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    row0 = raw.iloc[0].tolist()
    row1 = raw.iloc[1].tolist()
    parent_total_idx = waybill_total_idx = None
    for idx, (a, b) in enumerate(zip(row0, row1)):
        a = "" if pd.isna(a) else str(a).strip()
        b = "" if pd.isna(b) else str(b).strip()
        if a == "Parent Order #" and b == "Total":
            parent_total_idx = idx
        if a == "No. of Waybill" and b == "Total":
            waybill_total_idx = idx
    if parent_total_idx is None or waybill_total_idx is None:
        raise ValueError(f"Could not find 'Parent Order # / Total' or 'No. of "
                          f"Waybill / Total' columns in {path!r}.")

    per_d_order = {d: 0.0 for d in DISTRICTS}
    per_d_waybill = {d: 0.0 for d in DISTRICTS}
    overall_order = overall_waybill = None
    for _, row in raw.iloc[2:].iterrows():
        label = normalize_district_label(row.iloc[0])
        order_val = parse_number(row.iloc[parent_total_idx])
        waybill_val = parse_number(row.iloc[waybill_total_idx])
        if label in DISTRICTS:
            per_d_order[label] += order_val
            per_d_waybill[label] += waybill_val
        elif label.lower() == "grand total":
            overall_order, overall_waybill = order_val, waybill_val

    return {
        "overall": {"order": overall_order, "waybill": overall_waybill},
        "districts": {d: {"order": per_d_order[d], "waybill": per_d_waybill[d]} for d in DISTRICTS},
    }


def _read_tableau_crosstab_raw(path):
    """Tableau crosstab CSVs are UTF-16 / tab-separated; fall back to UTF-8 in case an export differs."""
    try:
        return pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    except (UnicodeError, pd.errors.ParserError):
        return pd.read_csv(path, encoding="utf-8-sig", sep=None, engine="python", header=None, dtype=str)


ODS_ROW_HEADER = "包派"


def parse_ods_counts_from_delivery_dashboard():
    """v8.0 §1 — ODS order and waybill count per district from the Tableau
    "Actual Delivery - 10 Districts" crosstab (replaces the OIX-derived figures):
      - Column A = district, Column B = row header, Column C = Parent Order #,
        Column E = No. of Waybill.
      - Only rows whose Column B is "包派" are ODS. A district with no such row
        has ODS order count = waybill count = 0.
    Column A is forward-filled, so it works whether or not the export repeats
    the district label on every row. Raises if the file is stale/missing or has
    no "包派" row at all (a real day is never all-zero) so the caller can fall
    back to the staged OIX figures instead of silently logging zeros.
    Returns (ods_order, ods_waybill) in the {"overall", "districts"} shape of
    order_count_for_group() / waybill_count_for_group()."""
    path = os.path.join(REPORT_FOLDER, REPORT_FILES["actual_delivery_10d"])
    if not os.path.exists(path):
        raise FileNotFoundError(f"{path!r} not found")
    if os.path.getmtime(path) < time.time() - STALE_REPORT_HOURS * 3600:
        raise FileNotFoundError(f"{path!r} is older than {STALE_REPORT_HOURS}h — this run's download failed")
    raw = _read_tableau_crosstab_raw(path)

    orders = {d: 0.0 for d in DISTRICTS}
    waybills = {d: 0.0 for d in DISTRICTS}
    current, matched = None, 0
    for _, row in raw.iterrows():
        a = row.iloc[0]
        if not pd.isna(a) and str(a).strip():
            current = normalize_district_label(a)
        b = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
        if b != ODS_ROW_HEADER or current not in DISTRICTS:
            continue
        orders[current] += parse_number(row.iloc[2])      # Column C — Parent Order #
        waybills[current] += parse_number(row.iloc[4])    # Column E — No. of Waybill
        matched += 1
    if matched == 0:
        raise ValueError(f"no '{ODS_ROW_HEADER}' row found in {path!r} — layout may have changed")
    ods_order = {"overall": int(round(sum(orders.values()))), "districts": {d: int(round(v)) for d, v in orders.items()}}
    ods_waybill = {"overall": int(round(sum(waybills.values()))), "districts": {d: int(round(v)) for d, v in waybills.items()}}
    return ods_order, ods_waybill


def parse_delay_early_pct(file_key):
    """v4.0 §2 — parses 'Actual Delivery - Delay & Early %.csv' (T-1) or
    'MTD Actual Delivery - Delay, Early & On Time %.csv' (month-to-date).
    Both share the shape: District1 (group) | Expected Timeslot w/ same day |
    Delay % | Early % | On Time %, with a 'Grand Total'/'Total' row for the
    whole-network figure and one row per district per timeslot.

    NOTE: unlike the T-1 file, the MTD export does NOT include a per-district
    'Total' (i.e. per-district "Overall") row — only the network-wide Grand
    Total, plus each district broken out by timeslot. So districts[d] will
    have "AM"/"PM"/"EV"/"EV2" keys from the MTD file, but no "Overall" key;
    the per-district MTD "Overall" delay% instead comes from
    parse_delay_rate_by_zone_type()'s "overall" (Grand Total row, all zones).

    Returns {"overall": {slot: {"delay","early","onTime"}},
             "districts": {d: {slot: {...}}}}.
    """
    path = os.path.join(REPORT_FOLDER, REPORT_FILES[file_key])
    df = pd.read_csv(path, encoding="utf-16", sep="\t", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    overall = {}
    districts = {d: {} for d in DISTRICTS}
    for _, row in df.iterrows():
        label = normalize_district_label(row.iloc[0])
        slot_raw = str(row.iloc[1]).strip()
        slot = DELAY_TIMESLOT_MAP.get(slot_raw, slot_raw)
        rec = {
            "delay": parse_percent(row["Delay %"]),
            "early": parse_percent(row["Early %"]),
            "onTime": parse_percent(row["On Time %"]),
        }
        if label.lower() == "grand total":
            overall[slot] = rec
        elif label in DISTRICTS:
            districts[label][slot] = rec
    return {"overall": overall, "districts": districts}


def parse_delay_rate_by_zone_type():
    """v4.0 §3/§4 — parses 'delay rate by zone type.csv': commercial_zone
    (0 = Residential, 1 = Commercial) x date, with a 'Total'/'Total' row per
    zone giving that zone's MTD delay% per district, and a final
    'Grand Total' row giving the combined (both zones) MTD delay% per
    district — this is the per-district "Overall" MTD delay rate used
    elsewhere as matrices.delayRate.actual.districts.
    Returns {"residential": {d:..}, "commercial": {d:..}, "overall": {d:..}}.
    """
    path = os.path.join(REPORT_FOLDER, REPORT_FILES["delay_zone_type_mtd"])
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header = raw.iloc[0].tolist()
    col_to_dist = {}
    for idx, label in enumerate(header):
        if idx < 3:
            continue  # commercial_zone / date / MAX(DATE(...)) columns
        norm = normalize_district_label(label)
        if norm in DISTRICTS:
            col_to_dist.setdefault(norm, []).append(idx)

    residential, commercial, overall = {}, {}, {}
    for _, row in raw.iloc[1:].iterrows():
        zone = str(row.iloc[0]).strip()
        col1 = str(row.iloc[1]).strip()
        col2 = str(row.iloc[2]).strip()
        if not (col1 == "Total" and col2 == "Total"):
            continue  # only the per-zone/grand MTD "Total" rows, skip daily rows
        target = {"0": residential, "1": commercial, "Grand Total": overall}.get(zone)
        if target is None:
            continue
        for dist, idxs in col_to_dist.items():
            for i in idxs:
                v = parse_percent(row.iloc[i])
                if v is not None:
                    target[dist] = v
    return {"residential": residential, "commercial": commercial, "overall": overall}


def backfill_delay_rate_history(history):
    """v4.0 §3 — 'delay rate by zone type.csv' also carries one row per date
    per commercial_zone (0/1), so a missing day in history["delayRate"] (used
    for the 30-day rolling forecast) can be recovered from it.

    Caveat: this file only gives Residential (zone 0) and Commercial (zone 1)
    rates separately per day — not a true order-volume-weighted "Overall".
    As a backfill-only approximation (never used for today's normal T-1
    write, which always comes from parse_delay_early_pct("delay_early")
    instead), each missing day's per-district "Overall" is taken as the
    simple average of that day's zone-0 and zone-1 rates. Good enough to
    keep the 30-day forecast window from having a hole; not a substitute for
    the real daily figure.
    """
    path = os.path.join(REPORT_FOLDER, REPORT_FILES["delay_zone_type_mtd"])
    if not os.path.exists(path):
        return
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header = raw.iloc[0].tolist()
    col_to_dist = {}
    for idx, label in enumerate(header):
        if idx < 3:
            continue
        norm = normalize_district_label(label)
        if norm in DISTRICTS:
            col_to_dist.setdefault(norm, []).append(idx)

    daily_zone_vals = {}  # date_str -> {"0": {d: val}, "1": {d: val}}
    for _, row in raw.iloc[1:].iterrows():
        zone = str(row.iloc[0]).strip()
        if zone not in ("0", "1"):
            continue
        date_label = str(row.iloc[2]).strip()  # "1/9/2026" — D/M/YYYY
        m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", date_label)
        if not m:
            continue  # skips that zone's own "Total" row, which isn't a real date
        date_str = f"{int(m.group(3)):04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
        bucket = daily_zone_vals.setdefault(date_str, {"0": {}, "1": {}})
        for dist, idxs in col_to_dist.items():
            for i in idxs:
                v = parse_percent(row.iloc[i])
                if v is not None:
                    bucket[zone][dist] = v

    delay_log = history.setdefault("delayRate", {})
    filled = []
    for date_str, zones in daily_zone_vals.items():
        if date_str in delay_log:
            continue  # never overwrite a real T-1 figure already on record
        per_d = {}
        for d in DISTRICTS:
            vals = [zones["0"].get(d), zones["1"].get(d)]
            vals = [v for v in vals if v is not None]
            per_d[d] = round(sum(vals) / len(vals), 2) if vals else None
        present = [v for v in per_d.values() if v is not None]
        overall = round(sum(present) / len(present), 2) if present else None
        delay_log[date_str] = {"overall": overall, "districts": per_d}
        filled.append(date_str)
    if filled:
        print(f"  🩹 Backfilled {len(filled)} missing Delay Rate day(s) (zone-average "
              f"approximation) from this run's download: {filled}")


def parse_mtd_overall_delay():
    """v4.0 §4 — the single Overall MTD Delay % headline figure for the
    Overview tab's main cell, from 'MTD Actual Delivery - Delay, Early & On
    Time %.csv' — its 'Grand Total' / 'Total' row's 'Delay %' column."""
    df = pd.read_csv(os.path.join(REPORT_FOLDER, REPORT_FILES["mtd_delay_early_ontime"]),
                      encoding="utf-16", sep="\t", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    for _, row in df.iterrows():
        if str(row.iloc[0]).strip().lower() == "grand total" and str(row.iloc[1]).strip() == "Total":
            return parse_percent(row["Delay %"])
    raise ValueError("Could not find the Grand Total/Total row in "
                      f"{REPORT_FILES['mtd_delay_early_ontime']!r}.")

def parse_poor_rating():
    """v6.0 §3 — parses REPORT_FILES["poor_rating"], now downloaded from the
    "Delivery Rating Report" Tableau report's "MTD Courier Rating by
    District" sheet (replaces the old "Delivery Rating"/LogisticsKPIReport
    source entirely — see TABLEAU_TARGETS above).

    Shape (confirmed from the sample export): row 0 repeats the MTD month
    label across every column (ignored); row 1 is the real header —
    "District (adjusted)" | "Courier Rating <=2 #" | "# of Order with
    rating" | "Total # of Order" | "1-2* Ratio"; row 2+ is one row per
    district plus a final "Grand Total" row.

    The value column is located by its header text ("1-2* Ratio") rather
    than a fixed column index — the old source's poor-rating figure sat at
    a hardcoded index 5, which no longer matches this report's column
    layout. Per spec, rounded to 3 decimal places (the old source used 2)."""
    path = os.path.join(REPORT_FOLDER, REPORT_FILES["poor_rating"])
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header_row_idx = ratio_col_idx = None
    for i in range(min(10, len(raw))):
        for j, cell in enumerate(raw.iloc[i]):
            if str(cell).strip() == "1-2* Ratio":
                header_row_idx, ratio_col_idx = i, j
                break
        if header_row_idx is not None:
            break
    if header_row_idx is None:
        raise ValueError(f"無法在 {path} 中找到 '1-2* Ratio' 欄位標頭。")

    per_district, overall = {}, None
    for _, row in raw.iloc[header_row_idx + 1:].iterrows():
        label = normalize_district_label(row.iloc[0])
        val = parse_percent(row.iloc[ratio_col_idx], decimals=3)
        if val is None: continue
        if label in DISTRICTS: per_district[label] = val
        elif label.lower() == "grand total": overall = val
    return {"overall": overall, "districts": {d: per_district.get(d) for d in DISTRICTS}}

def parse_rfid(target_date):
    df = pd.read_csv(os.path.join(REPORT_FOLDER, REPORT_FILES["rfid"]), encoding="utf-16", sep="\t")
    df.iloc[:, 0] = df.iloc[:, 0].ffill()
    log_rows = df[df.iloc[:, 0].astype(str).str.strip() == "LOG"]

    date_col_label = f"{target_date.strftime('%B')} {target_date.day}"
    if date_col_label not in df.columns:
        print(f"⚠️ 找不到日期欄位: {date_col_label}，可能因為 7 天滾動已被推掉。")
        return {"overall": 0.0, "districts": {d: 0.0 for d in DISTRICTS}}

    per_district = {d: 0.0 for d in DISTRICTS}
    for _, row in log_rows.iterrows():
        # v3.1: defensively normalized too, in case RFID's district column
        # ever starts emitting the same combined "NT-YT & WTH" label as
        # Delivery Rating — unconfirmed for this report as of this fix, but
        # cheap insurance since the alias table is a no-op for any other value.
        code = normalize_district_label(row.iloc[1])
        if code in DISTRICTS:
            per_district[code] += parse_number(row[date_col_label])
    return {"overall": round(sum(per_district.values()), 2), "districts": {d: round(v, 2) for d, v in per_district.items()}}


def parse_gmv(target_date):
    """v3.0 §3 — parses REPORT_FILES['gmv'] ("Sheet 1.csv"), downloaded by
    fetch_gmv_report() from the separate GMV Tableau account.

    Layout (same UTF-16 tab-separated crosstab shape as the other reports):
      row 0: 'delivery_district' marker row (ignored)
      row 1: real header — col 0 is the date-pivot label, remaining columns
             are district codes as exported by Tableau. If a "NT-YT" column
             is present, its values are merged into "WTH" (v3.1, Sept 2026 —
             was "NT-TW" under v3.0 §3, changed when the district grouping
             itself changed at the source) via normalize_district_code(),
             same helper §1/§2 use.
      row 2+: one row per date, col 0 = "YYYY年M月D日" (no zero-padding —
             confirmed against the sample export), remaining columns = GMV
             amounts (may contain "$"/"," — parsed with parse_money()).

    Returns the row for target_date (T-1) as {"overall", "districts"}.
    """
    path = os.path.join(REPORT_FOLDER, REPORT_FILES["gmv"])
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header = raw.iloc[1].tolist()
    data_rows = raw.iloc[2:].reset_index(drop=True)

    col_to_district = {}
    for idx, label in enumerate(header):
        if idx == 0:
            continue
        norm = normalize_district_code(str(label).strip())
        if norm in DISTRICTS:
            col_to_district[idx] = norm

    missing = [d for d in DISTRICTS if d not in col_to_district.values()]
    if missing:
        print(f"  ⚠️ GMV export is missing column(s) for district(s): {missing} "
              f"— those will be treated as $0 for every date.")

    target_label = f"{target_date.year}年{target_date.month}月{target_date.day}日"
    for _, row in data_rows.iterrows():
        if str(row.iloc[0]).strip() == target_label:
            per_district = {d: 0.0 for d in DISTRICTS}
            for idx, dist in col_to_district.items():
                per_district[dist] += parse_money(row.iloc[idx])
            return {"overall": round(sum(per_district.values()), 2),
                    "districts": {d: round(v, 2) for d, v in per_district.items()}}

    raise ValueError(f"Could not find GMV row for {target_label!r} in {path!r}.")


def backfill_gmv_history(history, cutoff_date=None):
    """v4.0 §3 — 'Sheet 1.csv' (the GMV crosstab) contains one row per date,
    not just T-1 — so if a day's GMV never made it into history["gmv"] (a
    missed run, a past failure, etc.), we can recover it straight from
    whatever date rows happen to still be present in today's download,
    instead of leaving a permanent gap. Only fills gaps; never overwrites a
    date that's already recorded (today's own T-1 write from parse_gmv()
    still happens separately/normally after this).

    v4.0.1 §3 correction — never backfills a day AFTER T-1 (cutoff_date,
    default = today - 1). Tableau's GMV export can include a row for today
    (or, in theory, later) that's still accumulating and not yet a complete
    day's GMV — e.g. if today is Sept 10, only Sept 9 and earlier are
    eligible; a same-day or future row is skipped even if present in the
    download, so a partial number never gets treated as that day's final
    GMV.
    """
    if cutoff_date is None:
        cutoff_date = dt.date.today() - dt.timedelta(days=1)  # T-1

    path = os.path.join(REPORT_FOLDER, REPORT_FILES["gmv"])
    if not os.path.exists(path):
        return
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header = raw.iloc[1].tolist()
    col_to_district = {}
    for idx, label in enumerate(header):
        if idx == 0:
            continue
        norm = normalize_district_code(str(label).strip())
        if norm in DISTRICTS:
            col_to_district[idx] = norm

    gmv_log = history.setdefault("gmv", {})
    filled, skipped_future = [], []
    for _, row in raw.iloc[2:].iterrows():
        label = str(row.iloc[0]).strip()
        m = re.match(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$", label)
        if not m:
            continue
        row_date = dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if row_date > cutoff_date:
            skipped_future.append(row_date.isoformat())
            continue  # never backfill past T-1, even if the export has the row
        date_str = row_date.isoformat()
        if date_str in gmv_log:
            continue  # already have this day — don't overwrite
        per_district = {d: 0.0 for d in DISTRICTS}
        for idx, dist in col_to_district.items():
            per_district[dist] += parse_money(row.iloc[idx])
        gmv_log[date_str] = {"overall": round(sum(per_district.values()), 2),
                              "districts": {d: round(v, 2) for d, v in per_district.items()}}
        filled.append(date_str)
    if filled:
        print(f"  🩹 Backfilled {len(filled)} missing GMV day(s) from this run's download: {filled}")
    if skipped_future:
        print(f"  ⏭️  Skipped {len(skipped_future)} GMV row(s) newer than T-1 ({cutoff_date.isoformat()}) "
              f"— not backfilled: {skipped_future}")

# =============================================================================
# 5. OIX Productivity 處理 (03:00 job — Excel/CSV-based, no Tableau involved)
# =============================================================================

def find_oix_file(target_date):
    """OIX_Record_YYYYMMDD.* for the given date (T-1), inside OIX_FOLDER."""
    for ext in ("xlsx", "csv"):
        fname = f"OIX_Record_{target_date.strftime('%Y%m%d')}.{ext}"
        path = os.path.join(OIX_FOLDER, fname)
        if os.path.exists(path):
            return path
    matches = glob.glob(os.path.join(OIX_FOLDER, f"OIX_Record_{target_date.strftime('%Y%m%d')}*"))
    if matches:
        return matches[0]
    raise FileNotFoundError(f"Could not find an OIX_Record file for {target_date} in {OIX_FOLDER}")


def load_oix(path):
    """Row 1 is the report title ('Waybill Status History Report') — skip it,
    row 2 is the real header. Handles both .xlsx and .csv exports."""
    if path.lower().endswith(".csv"):
        try:
            return pd.read_csv(path, header=1, dtype=str, encoding="utf-8-sig")
        except UnicodeDecodeError:
            return pd.read_csv(path, header=1, dtype=str, encoding="cp950")
    return pd.read_excel(path, header=1, dtype=str)


def load_staff_master_df():
    """v10.0 — Pulls the "Master LOG Staff List" Google Sheet (LOG Master
    tab) in place of the old local 'Logistics_Staff_List_YYYYMMDD.xlsx'
    export. gspread auth pattern, credential path, sheet URL/tab, and
    column layout are borrowed verbatim from ot_time_alert_workflow.py's
    load_staff_master() (same sheet, same service-account credential):
      A Staff ID, B User ID, C Fullname (EN), D Fullname (ZH), E Dept,
      F Supervisor, G Position, H Join Date, I Last Working Date,
      J Remarks, K Driving License Expiry Date.

    Returns a DataFrame with columns: Staff ID, Department Code (upper-
    cased/stripped), Position (upper-cased/stripped — comparable directly
    against LEADER_EXCLUDE_POSITIONS / COURIER_POSITIONS / DRIVER_POSITIONS
    / STAFF_CAP_POSITION_TO_BUCKET), and Active (bool — True when Last
    Working Date, column I, is blank). The sheet has no separate
    "Employment Status" column the way the old Excel export did, so a
    blank Last Working Date is used as the still-employed signal here.
    TODO: confirm this matches how the LOG team actually uses that column
    — e.g. whether someone serving notice gets a future-dated Last Working
    Date (would still read as Active here) or an immediate one.

    Deliberately pads any row shorter than column K to a full-length row
    of "" before indexing, rather than skipping it the way
    ot_time_alert_workflow.py's load_staff_master() does — gspread's
    get_all_values() trims trailing blank cells per row, so any row whose
    Last Working Date (I) / Remarks (J) / Driving License Expiry (K) are
    all blank — i.e. most currently-active staff — would otherwise come
    back shorter than expected and get silently dropped entirely. That's
    fine for an OT-alert script (a few missed staff just means a few
    missed alerts) but not here, where this same map drives Position
    classification and the Staff Allocation Cap for every district.

    Raises FileNotFoundError (same exception type the old Excel-based
    loader raised, so every existing caller's `except FileNotFoundError`
    soft-fail below keeps working unchanged) if gspread/google-auth aren't
    installed, the credential file is missing, or the sheet/tab can't be
    opened.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as e:
        raise FileNotFoundError(
            "gspread / google-auth not installed — run: pip install gspread "
            "google-auth --break-system-packages"
        ) from e

    if not os.path.exists(GOOGLE_CREDENTIAL_JSON):
        raise FileNotFoundError(
            f"Google service-account credential not found at {GOOGLE_CREDENTIAL_JSON!r} "
            f"— can't read the '{STAFF_MASTER_TAB_NAME}' staff master sheet."
        )

    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIAL_JSON, scopes=scopes)
        gc = gspread.authorize(creds)
        sheet_id = STAFF_MASTER_SHEET_URL.split("/d/")[1].split("/")[0]
        sh = gc.open_by_key(sheet_id)
        ws = sh.worksheet(STAFF_MASTER_TAB_NAME)
        values = ws.get_all_values()
    except Exception as e:
        # Covers gspread's SpreadsheetNotFound/WorksheetNotFound/APIError and
        # any transient network error alike — all treated as "staff master
        # not available this run", the same soft-fail contract a missing
        # Excel file used to have.
        raise FileNotFoundError(
            f"Could not read the '{STAFF_MASTER_TAB_NAME}' Google Sheet "
            f"({STAFF_MASTER_SHEET_URL}): {e}"
        ) from e

    header_idx = STAFF_MASTER_HEADER_ROW - 1
    rows = values[header_idx + 1:]

    staffid_i = col(STAFF_MASTER_STAFFID_COL)
    dept_i = col(STAFF_MASTER_DEPT_CODE_COL)
    pos_i = col(STAFF_MASTER_POSITION_COL)
    lwd_i = col(STAFF_MASTER_LAST_WORKING_DATE_COL)
    width = max(staffid_i, dept_i, pos_i, lwd_i) + 1

    records = []
    for row in rows:
        if len(row) < width:
            row = row + [""] * (width - len(row))  # see docstring — pad, don't skip
        staff_id = row[staffid_i].strip()
        if not staff_id:
            continue
        records.append({
            "Staff ID": staff_id,
            "Department Code": row[dept_i].strip().upper(),
            "Position": row[pos_i].strip().upper(),
            "Active": row[lwd_i].strip() == "",
        })

    df = pd.DataFrame(records)
    print(f"  📋 Loaded staff master from Google Sheet '{STAFF_MASTER_TAB_NAME}' ({len(df)} employees)")
    return df


def load_staff_position_map():
    """v10.0 — Staff ID -> Position, from the Google Sheet staff master
    (see load_staff_master_df()). Previously read the latest local
    'Logistics_Staff_List_YYYYMMDD.xlsx' export (v3.0 §4) — replaced
    because the Google Sheet is the actively-maintained source and doesn't
    depend on someone remembering to re-export Excel. Position values are
    already upper-cased/stripped by load_staff_master_df()."""
    df = load_staff_master_df()
    return dict(zip(df["Staff ID"], df["Position"]))


def compute_staff_allocation_cap():
    """v9.0, updated v10.0 — Reads the Google Sheet staff master (via
    load_staff_master_df(), same source now used for the Courier/Driver
    Position classification above) and reduces it to a maximum-cap table
    for Suggest Headcount: current ACTIVE headcount per district, split
    into the same 4 buckets the Per-District Staff Plan table uses (FT
    Driver, PT Driver, FT Courier, PT Courier).

    v10.0 — switched from the local 'Logistics_Staff_List_YYYYMMDD.xlsx'
    export to the Google Sheet (more stable, no dependency on someone
    remembering to re-export). "Active" is now an empty Last Working Date
    (column I) rather than an 'Employment Status' column — the sheet
    doesn't have one; see load_staff_master_df()'s docstring for the
    notice-period caveat.

    - Position -> bucket via STAFF_CAP_POSITION_TO_BUCKET, unchanged. A
      Position not in that map (team leader, supervisor, manager, transit-
      truck driver, fleet/ops/HQ role, ...) is excluded from every cap on
      purpose — see the comment on STAFF_CAP_BUCKET_POSITIONS above.
    - v10.2 — Department Code -> District via dept_code_to_district(),
      the SAME prefix/startswith() matching ot_time_alert_workflow.py's
      function of the same name uses against its own DEPT_CODE_TO_DISTRICT
      list, reading the same sheet column. Previously an exact-code dict
      (DEPARTMENT_CODE_DISTRICT_MAP) that hand-enumerated every specific
      code ("LOGETH01"/"LOGETHP01"/...); switched so a new sub-code the
      district office starts using is picked up automatically instead of
      silently landing in "unmapped" until someone notices and edits the
      dict — and so this script and ot_time_alert_workflow.py can't drift
      apart on what a given Department Code means.

    Returns {district: {"ftDriver":n, "ftCourier":n, "ptDriver":n,
    "ptCourier":n}} for every one of the 10 DISTRICTS (0 where the staff
    list currently has nobody in that bucket), plus "_asOf" (today's date —
    the Google Sheet has no filename date stamp the old Excel export did)
    and "_source" (the sheet name, for the same reason).

    Raises FileNotFoundError if the sheet can't be read — same as
    load_staff_position_map(), left for the caller to soft-fail on so a
    temporarily-unreachable sheet doesn't take down the rest of the 14:00
    run.
    """
    df = load_staff_master_df()

    cap = {d: {"ftDriver": 0, "ftCourier": 0, "ptDriver": 0, "ptCourier": 0} for d in DISTRICTS}

    active = df[df["Active"]]

    unmapped_depts, unmapped_positions = set(), set()
    for dept_code, pos in zip(active["Department Code"], active["Position"]):
        district = dept_code_to_district(dept_code)  # v10.2 — prefix match, "" = unmapped
        if not district:
            if dept_code:
                unmapped_depts.add(dept_code)
            continue
        bucket = STAFF_CAP_POSITION_TO_BUCKET.get(pos)
        if bucket is None:
            if pos:
                unmapped_positions.add(pos)
            continue
        cap[district][bucket] += 1

    if unmapped_positions:
        print(f"  ℹ️ Staff Allocation Cap: {len(unmapped_positions)} Position title(s) not in "
              f"STAFF_CAP_BUCKET_POSITIONS, excluded from every district's cap (supervisory/other "
              f"roles): {sorted(unmapped_positions)}")
    if unmapped_depts:
        print(f"  ℹ️ Staff Allocation Cap: {len(unmapped_depts)} Department Code(s) matched no "
              f"prefix in DEPT_CODE_TO_DISTRICT, excluded (non-district departments — HQ, fleet "
              f"management, etc.): {sorted(unmapped_depts)}")

    cap["_asOf"] = dt.date.today().isoformat()
    cap["_source"] = f"Google Sheet: {STAFF_MASTER_TAB_NAME}"
    return cap


def process_oix(df, position_map=None, dedupe=True):
    """Cleaning + district tagging steps: drop O2O rows from non-LF/LP/ODS/VAN
    users, fill blank Parent Order from Order Number, dedupe, tag District.

    v3.0 §4: if position_map (Employee No. -> Position, from
    load_staff_position_map()) is given, also tags a "Position" column
    (matching the spec's "Column X [NEW]") by looking up each row's User
    (Column E) — used downstream for Courier/Driver classification
    (manpower_distribution_for_group) and the §4.1 leader-exclusion in
    productivity_for_group(). If position_map is None (e.g. the staff list
    couldn't be found this run), "Position" is left all-blank and those two
    features simply have nothing to classify — everything else is
    unaffected.

    v19.0 §2 — `dedupe` (default True, unchanged behavior for every existing
    caller) controls the drop_duplicates(subset=[Parent Order, User]) step
    below. That collapse is correct for manpower/order-count purposes (one
    row per staff/order pair) but WRONG for a true waybill count: a single
    Parent Order + User pair can legitimately carry several distinct Waybill
    Numbers (multiple totes/parcels on the same order), and the (Parent
    Order, User) dedupe silently keeps only one of them. Callers that need
    an accurate unique-Waybill-Number count (waybill_count_for_group) should
    pass dedupe=False to get a still-cleaned, still-District-tagged frame
    with every waybill row intact — Waybill Number itself is still safe to
    call .nunique() on directly, since a given waybill number's only
    repetition in this extract is its own status-history rows, not a
    different waybill."""
    c_user, c_addr = col("E"), col("R")
    c_order_no, c_parent = col("K"), col("L")
    c_truck = col("P")

    user = df.iloc[:, c_user].fillna("")
    addr = df.iloc[:, c_addr].fillna("")

    valid_prefixes = ("LF", "LP", "ODS", "VAN")
    remove_mask = addr.str.contains("O2O", na=False) & ~user.str.startswith(valid_prefixes)
    df = df.loc[~remove_mask].copy()

    parent = df.iloc[:, c_parent]
    order_no = df.iloc[:, c_order_no]

    def normalize_parent_order(l_val, k_val):
        """Two Parent Order formats now exist:
          - "H..." — the original/existing format, unchanged.
          - "EM..." — newer format; the leading "E" gets stripped (so "EM..."
            becomes "M...") wherever it's found, whether that's already
            sitting in Column L or has to be derived from Column K.
        Priority: Column L's OWN existing value wins whenever it's non-blank
        — its own prefix decides the transformation, and Column K's prefix
        is irrelevant in that case (e.g. L starts with "H" and K starts with
        "EM" -> still counted as an "H" order, using L's value as-is).
        Only when L is blank do we derive a value from K at all."""
        l_str = "" if pd.isna(l_val) else str(l_val).strip()
        if l_str:
            if l_str.startswith("EM"):
                return l_str[1:]              # "EM..." -> "M..." — strip the leading E only
            return l_str                       # "H..." (or anything else) — unchanged

        k_str = "" if pd.isna(k_val) else str(k_val).strip()
        if k_str.startswith("EM"):
            return k_str[1:][:13]              # strip "E" first, THEN take 13 chars of the
                                                # remainder — keeps the same 13-char, single-
                                                # leading-letter shape as the "H" format below
        return k_str[:13]                      # existing "H" (or unrecognized-prefix fallback)
                                                # behavior, unchanged from before

    df.iloc[:, c_parent] = [normalize_parent_order(l, k) for l, k in zip(parent, order_no)]

    if dedupe:
        df = df.drop_duplicates(subset=[df.columns[c_parent], df.columns[c_user]])

    df["District"] = df.iloc[:, c_truck].apply(district_from_truck_no)
    unmatched = df["District"].isna().sum()
    if unmatched:
        print(f"  ⚠️ {unmatched} row(s) had a truck number matching none of the 14 "
              f"district patterns — excluded from every total. Check df.iloc[:, {c_truck}].")

    if position_map:
        user_stripped = df.iloc[:, c_user].fillna("").astype(str).str.strip()
        df["Position"] = user_stripped.map(lambda u: position_map.get(u, "") or None)
    else:
        df["Position"] = None

    return df


def manpower_for_group(df, prefixes, exclude_positions=None):
    """v4.0 §1 — replaces the manpower half of the old productivity_for_group()
    (order-count is no longer computed from OIX at all; see
    parse_actual_delivery_timeslot() / finish_productivity_with_orders()).
    exclude_positions (v3.0 §4.1): staff whose "Position" column falls in
    this set are dropped from the manpower headcount. Positions come from
    process_oix()'s "Position" column, so this only has an effect when that
    run had a position_map available.
    Returns {"overall": int, "districts": {d: int}} — unique HKTV-user
    headcount per district for this group (LF/LP, or ODS/VAN)."""
    c_user = col("E")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(prefixes)]
    if exclude_positions:
        sub = sub[~sub["Position"].fillna("").isin(exclude_positions)]
    per_district = {}
    for d in DISTRICTS:
        rows = sub[sub["District"] == d]
        per_district[d] = int(rows.iloc[:, c_user].nunique())
    return {"overall": sum(per_district.values()), "districts": per_district}


def order_count_for_group(df, prefixes):
    """v4.0.1 — restores the ODS/VAN half of the OLD (pre-v4.0)
    productivity_for_group()'s order-count logic: unique Parent Orders
    (Column L) per district, counted straight from the OIX extract, for
    rows whose User (Column E) starts with `prefixes`.

    Per the v4.0.1 correction: ODS order count goes back to being computed
    this way from OIX (not from the Tableau "Delivery Dashboard" report at
    all) — HKTV order count is then derived as Tableau's network-wide total
    MINUS this OIX-derived ODS figure (see finish_productivity_with_orders()),
    rather than both groups sharing the same raw Tableau total as before.
    Returns {"overall": int, "districts": {d: int}}."""
    c_user, c_parent = col("E"), col("L")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(prefixes)]
    per_district = {}
    for d in DISTRICTS:
        rows = sub[sub["District"] == d]
        per_district[d] = int(rows.iloc[:, c_parent].nunique())
    return {"overall": sum(per_district.values()), "districts": per_district}


def waybill_count_for_group(df, prefixes):
    """v4.0.1 §5 — the ODS/VAN waybill-count counterpart to
    order_count_for_group() above: IDENTICAL OIX-based logic (same User
    prefix filter, same per-district grouping), the only difference being
    the column deduplicated on — unique Waybill Number (Column G) instead
    of unique Parent Order (Column L). Per spec: ODS waybill count is no
    longer half of the Tableau network-wide waybill total; it's counted
    straight from OIX the same way ODS order count is, just deduped on
    Column G.
    Returns {"overall": int, "districts": {d: int}}."""
    c_user, c_waybill = col("E"), col("G")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(prefixes)]
    per_district = {}
    for d in DISTRICTS:
        rows = sub[sub["District"] == d]
        per_district[d] = int(rows.iloc[:, c_waybill].nunique())
    return {"overall": sum(per_district.values()), "districts": per_district}


def split_hktv_ods_totals(tableau_totals, ods_order, ods_waybill):
    """v4.0.1 §2/§5 — combines the OIX-derived ODS/VAN order count
    (order_count_for_group) and waybill count (waybill_count_for_group)
    with the Tableau network-wide order/waybill totals
    (parse_actual_delivery_timeslot() output) into two group-specific
    totals, in the same {"overall":{"order","waybill"},
    "districts":{d:{"order","waybill"}}} shape as tableau_totals itself:
      - ODS/VAN: the OIX-derived figures, used as-is.
      - HKTV: Tableau's network-wide total MINUS the OIX-derived ODS
        figure, per district and overall — restores the pre-v4.0 approach
        (each group has its own genuine order/waybill count that sums back
        to the network total) instead of both groups sharing the same raw
        Tableau total.
    """
    ods_districts = {
        d: {"order": ods_order["districts"][d], "waybill": ods_waybill["districts"][d]}
        for d in DISTRICTS
    }
    hktv_districts = {
        d: {
            "order": tableau_totals["districts"][d]["order"] - ods_order["districts"][d],
            "waybill": tableau_totals["districts"][d]["waybill"] - ods_waybill["districts"][d],
        }
        for d in DISTRICTS
    }
    ods_totals = {
        "overall": {"order": ods_order["overall"], "waybill": ods_waybill["overall"]},
        "districts": ods_districts,
    }
    hktv_totals = {
        "overall": {
            "order": tableau_totals["overall"]["order"] - ods_order["overall"],
            "waybill": tableau_totals["overall"]["waybill"] - ods_waybill["overall"],
        },
        "districts": hktv_districts,
    }
    return hktv_totals, ods_totals


def manpower_distribution_for_group(df, group_key):
    """v3.0 §4 — distinct HKTV (LF/LP) staff headcount per district, for
    group_key 'courier' or 'driver' (see MANPOWER_GROUP_POSITIONS). Feeds
    the HKTV Manpower Distribution tab, which is pure headcount (not an
    orders/productivity ratio like the other tabs)."""
    positions = MANPOWER_GROUP_POSITIONS[group_key]
    c_user = col("E")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(("LF", "LP"))]
    sub = sub[sub["Position"].fillna("").isin(positions)]
    per_district = {}
    for d in DISTRICTS:
        rows = sub[sub["District"] == d]
        per_district[d] = int(rows.iloc[:, c_user].nunique())
    total = sum(per_district.values())
    return {"districts": per_district, "total": total}


def odsvan_manpower_group(ods_ratio_manpower):
    """v5.0 §1 — reshapes manpower_for_group(df, ("ODS","VAN"))'s output
    ({"overall","districts"}) into the same {"districts","total"} shape
    manpower_distribution_for_group() returns for courier/driver, so the
    HKTV Manpower Distribution tab's ODS/VAN group can be logged and read
    identically to the Courier/Driver groups. Per the patch spec, this is
    the SAME unique-Staff-code-starting-with-ODS/VAN count, by the SAME
    Districts classification, already computed for the ODS / VAN
    Productivity tab — just relabelled for this tab's log shape, not a
    second computation."""
    return {"districts": ods_ratio_manpower["districts"], "total": ods_ratio_manpower["overall"]}


def save_manpower_staging(target_date, hktv_staff_manpower, ods_ratio_manpower, courier_group, driver_group,
                           ods_order_count, ods_waybill_count):
    """v4.0 §1 — 03:00 hand-off to the 14:00 job (see MANPOWER_STAGING_PATH):
    everything the 03:00 OIX run can produce (manpower headcounts, plus —
    v4.0.1 §2/§5 — the OIX-derived ODS/VAN order and waybill counts) before
    the Tableau order counts even exist yet."""
    payload = {
        "date": target_date.isoformat(),
        "hktvStaffManpower": hktv_staff_manpower,
        "odsRatioManpower": ods_ratio_manpower,
        "courierGroup": courier_group,
        "driverGroup": driver_group,
        "odsOrderCount": ods_order_count,
        "odsWaybillCount": ods_waybill_count,
    }
    with open(MANPOWER_STAGING_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {MANPOWER_STAGING_PATH} (staged for the 14:00 Tableau job)")


def load_manpower_staging():
    if not os.path.exists(MANPOWER_STAGING_PATH):
        return None
    with open(MANPOWER_STAGING_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def run_productivity_section():
    """v4.0 §1 — 03:00 job. OIX now only supplies MANPOWER (unique HKTV-user
    headcounts) — parent order counts moved to the Tableau-based 14:00 job
    (see finish_productivity_with_orders()). This job stages its manpower
    output to MANPOWER_STAGING_PATH for the 14:00 job to pick up, and still
    writes the HKTV Manpower Distribution tab's own file directly (that tab
    is pure headcount and doesn't depend on order counts at all)."""
    print("🚀 開始處理 OIX Manpower 數據...")
    target_date = dt.date.today() - dt.timedelta(days=1)  # T-1
    path = find_oix_file(target_date)
    df = load_oix(path)
    raw_rows = len(df)      # v10.3 — recorded with the file fingerprint, see mark_oix_processed()

    # v3.0 §4 / §4.1: without a position_map, process_oix() still runs fine
    # (Position column just stays blank) — HKTV Staff Productivity falls
    # back to its pre-v3.0 behavior (no leader exclusion) and the Manpower
    # Distribution tab simply has nothing new for today, rather than the
    # whole 03:00 job failing over a missing/late staff list file.
    try:
        position_map = load_staff_position_map()
    except FileNotFoundError as e:
        print(f"  ⚠️ {e} — skipping Position/Courier/Driver classification for "
              f"today's run (§4.1 leader-exclusion and HKTV Manpower Distribution "
              f"won't reflect today until a staff list file is present).")
        position_map = None

    df = process_oix(df, position_map)

    hktv_staff_manpower = manpower_for_group(df, ("LF", "LP"), exclude_positions=LEADER_EXCLUDE_POSITIONS)
    ods_ratio_manpower = manpower_for_group(df, ("ODS", "VAN"))
    courier_group = manpower_distribution_for_group(df, "courier")
    driver_group = manpower_distribution_for_group(df, "driver")
    ods_van_group = odsvan_manpower_group(ods_ratio_manpower)  # v5.0 §1 — Manpower Distribution tab's 3rd group

    # v4.0.1 §2/§5 — ODS/VAN order count AND waybill count go back to being
    # computed straight from OIX (old-version logic), not derived from
    # Tableau at all; HKTV's side is backed out of the Tableau total against
    # this figure later, in finish_productivity_with_orders() (see
    # split_hktv_ods_totals()).
    ods_order_count = order_count_for_group(df, ("ODS", "VAN"))
    ods_waybill_count = waybill_count_for_group(df, ("ODS", "VAN"))

    save_manpower_staging(target_date, hktv_staff_manpower, ods_ratio_manpower, courier_group, driver_group,
                           ods_order_count, ods_waybill_count)

    # v3.0 §4: HKTV Manpower Distribution — independent of order counts,
    # still written straight from the 03:00 run as before.
    history = load_history()
    # v10.3 — written through write_oix_manpower_entry() instead of append_manpower_log(): it keeps the
    # figures already on record when the staff list was unreadable (append_manpower_log wrote all-zero
    # Courier/Driver headcounts in that case), refreshes ODS/VAN on Daily-Cost-Report days, and the
    # file's fingerprint is remembered so a later UPDATE of this OIX file is detected.
    written, wrote_cd = write_oix_manpower_entry(history, target_date.isoformat(), df, have_positions=bool(position_map))
    if written:
        mark_oix_processed(history, target_date.isoformat(), path, raw_rows,
                           courier_driver=(wrote_cd or history["manpowerDistributionLog"][target_date.isoformat()].get("_source") == "costReport"))
        history["oixManpowerProcessed"][target_date.isoformat()]["hktvLeaderExcluded"] = bool(position_map)   # v10.4
    else:
        print(f"  ⚠️ {target_date.isoformat()}: staff list unreadable and no earlier Courier/Driver on record — "
              f"Manpower Distribution not written; the OIX backfill will fill it once the staff list is readable.")
    # v10.3 — T-1 itself was just written above and fingerprinted. Any OTHER OIX_Record file that was modified
    # today (or, if OIX_BACKFILL_DAYS is set, any file of the last N days) is re-checked here too.
    res = backfill_manpower_from_oix(history, position_map=position_map, skip_dates={target_date.isoformat()})
    save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
    persist_manpower_corrections(history, res)       # v10.4 — productivity / MTD / data.json follow the corrected manpower
    save_history(history)
    print("✅ Manpower 數據處理完成，已寫入 staging 檔案（Productivity 將於 14:00 Tableau job 完成）")


def _group_to_log_shape(manpower_group, order_totals):
    """v4.0 §1 — combines a group's OIX-derived manpower with the SHARED
    Tableau order+waybill totals (order_totals = parse_actual_delivery_timeslot()
    output; per the v4.0 decision, HKTV Staff and ODS Ratio both use the same
    Tableau parent-order figure as numerator) into the dashboard's
    dailyProductivity schema:
    {districts:{d:{orderCount, waybillCount, manpower, productivity}}, total:{...}}."""
    districts = {}
    for d in DISTRICTS:
        manpower = manpower_group["districts"].get(d, 0)
        order_count = order_totals["districts"][d]["order"]
        waybill_count = order_totals["districts"][d]["waybill"]
        districts[d] = {
            "orderCount": order_count,
            "waybillCount": waybill_count,
            "manpower": manpower,
            "productivity": round(order_count / manpower, 2) if manpower and order_count is not None else None,
        }
    total_orders = order_totals["overall"]["order"]
    total_waybill = order_totals["overall"]["waybill"]
    total_manpower = sum(v["manpower"] for v in districts.values())
    return {
        "districts": districts,
        "total": {
            "orderCount": total_orders,
            "waybillCount": total_waybill,
            "manpower": total_manpower,
            "productivity": round(total_orders / total_manpower, 2) if total_manpower and total_orders is not None else None,
        },
    }


def append_daily_productivity_log(history, date_str, hktv_manpower, ods_manpower,
                                   hktv_order_totals, ods_order_totals):
    """v4.0.1 §2/§5: each group now gets its OWN order/waybill totals —
    ods_order_totals is the OIX-derived ODS/VAN figure, hktv_order_totals is
    the Tableau total minus that figure (see split_hktv_ods_totals() /
    finish_productivity_with_orders()) — instead of both groups sharing the
    same raw Tableau total as under v4.0."""
    log = history.setdefault("dailyProductivityLog", {})
    log[date_str] = {
        "hktvStaff": _group_to_log_shape(hktv_manpower, hktv_order_totals),
        "odsRatio": _group_to_log_shape(ods_manpower, ods_order_totals),
    }


def trimmed_daily_productivity(history, keep_days):
    """Recomputed fresh from history.json (the durable store) every run, same
    pattern as rfidMonthly below — data.json only ever carries a recent
    window so it doesn't grow unbounded, while history.json keeps everything."""
    log = history.get("dailyProductivityLog", {})
    recent_dates = sorted(log.keys())[-keep_days:]
    return {d: log[d] for d in recent_dates}


def append_manpower_log(history, date_str, courier_group, driver_group, ods_van_group=None):
    """v3.0 §4 — durable full log of daily HKTV Courier/Driver headcount,
    same shape/role as append_daily_productivity_log() above.
    v5.0 §1 — optionally also logs that day's ODS/VAN headcount group
    (odsvan_manpower_group()) alongside Courier/Driver, so the Manpower
    Distribution tab can offer ODS/VAN as a third Group option. Optional
    (defaults to None / omitted) so old callers and old log entries without
    an ODS/VAN figure keep working unchanged."""
    log = history.setdefault("manpowerDistributionLog", {})
    # v8.0 §2 — a day already filled from the Daily Cost Report is the real figure;
    # the OIX-based 03:00 calculation must never cover it again.
    if log.get(date_str, {}).get("_source") == "costReport":
        print(f"  ℹ️ {date_str} manpower distribution already comes from the Daily Cost Report — OIX figure not written.")
        return
    entry = {"courier": courier_group, "driver": driver_group}
    if ods_van_group is not None:
        entry["odsVan"] = ods_van_group
    log[date_str] = entry


def trimmed_manpower_distribution(history, keep_days):
    """Recent-window slice of manpowerDistributionLog, same pattern as
    trimmed_daily_productivity() above — history.json keeps everything,
    manpower_distribution.json only ever serves the last `keep_days`."""
    log = history.get("manpowerDistributionLog", {})
    recent_dates = sorted(log.keys())[-keep_days:]
    return {d: log[d] for d in recent_dates}


# -----------------------------------------------------------------------------
# v10.3 — OIX manpower backfill (re-computes the HKTV Manpower Distribution for any
# day whose OIX_Record file has been updated since it was last processed).
# -----------------------------------------------------------------------------
_UNSET = object()


def _oix_fingerprint(path):
    """Content fingerprint of an OIX_Record file: size + SHA-1. Deliberately NOT the modified time —
    copying/re-downloading an identical file changes the mtime without changing a single record,
    and an updated file can in principle keep an old mtime."""
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return {"file": os.path.basename(path), "size": os.path.getsize(path), "sha1": h.hexdigest()}


def list_oix_files(folder=None):
    """{date: path} for every OIX_Record_YYYYMMDD.(xlsx|csv) in `folder`. If a day has both an .xlsx and a
    .csv, the more recently modified one wins (the updated export is the one to trust)."""
    folder = folder or OIX_FOLDER
    found = {}
    for p in glob.glob(os.path.join(folder, "OIX_Record_*")):
        m = OIX_FILE_RE.match(os.path.basename(p))
        if not m:
            continue
        try:
            d = dt.datetime.strptime(m.group(1), "%Y%m%d").date()
        except ValueError:
            continue
        if d not in found or os.path.getmtime(p) > os.path.getmtime(found[d]):
            found[d] = p
    return found


def write_oix_manpower_entry(history, date_str, df, have_positions):
    """Writes one day's OIX-derived figures into history["manpowerDistributionLog"][date_str] and returns
    (written, courier_driver_written). `df` must come from process_oix().

      - odsVan comes straight from the unique ODS/VAN users and needs no staff list, so it is always refreshed.
      - courier / driver need the Position tags (staff list). Without them (have_positions False) the figures
        already on record are KEPT — the old code wrote all-zero Courier/Driver headcounts in that case.
      - A day whose _source is 'costReport' keeps the Daily Cost Report's courier / driver / courierPT /
        driverPT (the real figures); only its odsVan, which the cost report doesn't carry, is refreshed.
      - A day with no courier/driver to show (no staff list and nothing on record) is not created at all
        rather than half-created; it is picked up on a later run once the staff list is readable."""
    log = history.setdefault("manpowerDistributionLog", {})
    entry = dict(log.get(date_str, {}))
    is_cost = entry.get("_source") == "costReport"
    wrote_cd = False
    if not is_cost and have_positions:
        entry["courier"] = manpower_distribution_for_group(df, "courier")
        entry["driver"] = manpower_distribution_for_group(df, "driver")
        wrote_cd = True
    if "courier" not in entry or "driver" not in entry:
        return False, False
    entry["odsVan"] = odsvan_manpower_group(manpower_for_group(df, ("ODS", "VAN")))
    log[date_str] = entry
    return True, wrote_cd


def mark_oix_processed(history, date_str, path, rows, courier_driver, fp=None):
    """Remembers which version of the OIX file a day's manpower figures were computed from."""
    rec = dict(fp) if fp else _oix_fingerprint(path)
    rec.update({"rows": int(rows), "courierDriver": bool(courier_driver),
                "processedAt": dt.datetime.now().isoformat(timespec="seconds")})
    history.setdefault("oixManpowerProcessed", {})[date_str] = rec


def _manpower_totals(entry):
    return {g: (entry.get(g) or {}).get("total") for g in ("courier", "driver", "odsVan")} if entry else {}


def backfill_manpower_from_oix(history, position_map=_UNSET, days=None, force=False, skip_dates=None, as_of=None):
    """v10.3 — re-computes the HKTV Manpower Distribution for the days whose OIX_Record file is new or has been
    UPDATED since it was last processed.

    Which files are looked at: by default (days=None) only files whose file MODIFIED TIME is today (T+0),
    whatever date is in the file name — a file named for an earlier day but touched today is a file that was
    updated today; untouched older files are left to the Daily Cost Report. A file named for today or a
    later day is never used (that day is still running). With days=N, every file of the last N days up to T-1
    is checked instead, regardless of modified time.

    A day is (re)processed when its file exists in OIX_FOLDER and any of these holds:
      * it was never processed by this logic (no history["oixManpowerProcessed"] record — covers every
        day logged before v10.3, whose figures may have come from a partial export);
      * the file's SHA-1 differs from the processed one (the file was updated / replaced);
      * its Courier/Driver could not be classified last time (no staff list) and the staff list is
        readable now;
      * the day has no log entry yet or its entry has no odsVan;
      * force=True.
    Guards:
      * costReport days keep their Cost-Report courier/driver (see write_oix_manpower_entry);
      * a changed file with FEWER rows than the one already processed is skipped with a warning (an old /
        partial export copied over a complete one must not shrink the figures) unless force=True;
      * a file that fails to parse is skipped without touching the day.
    `position_map`: pass the already-loaded Staff ID -> Position map (None = it could not be loaded);
    left unset, it is loaded lazily from the staff master sheet, and only if a day needs it.
    `skip_dates`: ISO dates handled by the caller (the 03:00 job's own T-1 write).
    Mutates `history` only (the caller saves). Returns {"updated": [...], "unchanged": [...],
    "skipped": [...]}; every refreshed day also prints old -> new totals."""
    days = OIX_BACKFILL_DAYS if days is None else days     # None = "modified today" mode
    today = as_of or dt.date.today()
    skip_dates = set(skip_dates or ())
    result = {"updated": [], "unchanged": [], "skipped": [], "manpower": {}}   # v10.4: "manpower" = corrections for apply_manpower_corrections()
    files = list_oix_files()
    if days is None:     # default: files whose modified time is today (T+0)
        candidates = sorted(d for d, pth in files.items()
                            if d < today and d.isoformat() not in skip_dates
                            and dt.date.fromtimestamp(os.path.getmtime(pth)) == today)
        scope = "modified today"
    else:                # explicit window: last N days up to T-1, regardless of modified time
        window = [today - dt.timedelta(days=n) for n in range(days, 0, -1)]
        candidates = [d for d in window if d in files and d.isoformat() not in skip_dates]
        scope = f"of the last {days} day(s)"
    if not candidates:
        print(f"  ℹ️ OIX manpower backfill: no OIX_Record file {scope} in {OIX_FOLDER!r} to check.")
        return result

    log = history.setdefault("manpowerDistributionLog", {})
    processed = history.setdefault("oixManpowerProcessed", {})
    pm = {"map": position_map}

    def get_position_map():
        if pm["map"] is _UNSET:
            try:
                pm["map"] = load_staff_position_map()
            except FileNotFoundError as e:
                print(f"  ⚠️ OIX manpower backfill: {e} — Courier/Driver of refreshed days stay as they are; "
                      f"only ODS/VAN is refreshed until the staff list can be read.")
                pm["map"] = None
        return pm["map"]

    for d in candidates:
        date_str, path = d.isoformat(), files[d]
        entry, rec = log.get(date_str), processed.get(date_str)
        is_cost = bool(entry) and entry.get("_source") == "costReport"
        try:
            fp = _oix_fingerprint(path)
        except OSError as e:
            print(f"  ⚠️ OIX manpower backfill: cannot read {os.path.basename(path)!r} ({e}) — {date_str} skipped.")
            result["skipped"].append(date_str)
            continue

        reasons = []
        if force:
            reasons.append("forced")
        if rec is None:
            reasons.append("not processed before (may have been logged from a partial export)")
        elif rec.get("sha1") != fp["sha1"]:
            reasons.append("OIX file updated")
        if rec is not None and not rec.get("courierDriver") and not is_cost and get_position_map():
            reasons.append("Courier/Driver were not classified last time")
        if rec is not None and rec.get("hktvLeaderExcluded") is False and get_position_map():   # v10.4
            reasons.append("HKTV manpower had no leader-exclusion last time")
        if entry is None or "odsVan" not in entry:
            reasons.append("log entry missing or without ODS/VAN")
        if not reasons:
            result["unchanged"].append(date_str)
            continue

        try:
            df_raw = load_oix(path)
            rows = len(df_raw)
        except Exception as e:
            print(f"  ⚠️ OIX manpower backfill: {os.path.basename(path)!r} failed to parse ({e}) — {date_str} skipped.")
            result["skipped"].append(date_str)
            continue
        if (rec is not None and not force and rec.get("sha1") != fp["sha1"]
                and rows < (rec.get("rows") or 0)):
            print(f"  ⚠️ OIX manpower backfill: {os.path.basename(path)!r} has {rows:,} rows but the version already "
                  f"processed had {rec['rows']:,} — looks like an older/partial export, {date_str} NOT updated "
                  f"(use --force-oix to override).")
            result["skipped"].append(date_str)
            continue

        positions = get_position_map()     # v10.4: also needed on cost-report days — HKTV leader-exclusion uses Position
        have_positions = bool(positions)
        before = _manpower_totals(entry)
        df = process_oix(df_raw, positions)
        written, wrote_cd = write_oix_manpower_entry(history, date_str, df, have_positions)
        if not written:
            print(f"  ⚠️ OIX manpower backfill: {date_str} has no Courier/Driver on record and the staff list is "
                  f"unavailable — left for a later run.")
            result["skipped"].append(date_str)
            continue
        mark_oix_processed(history, date_str, path, rows, courier_driver=(wrote_cd or is_cost), fp=fp)
        history["oixManpowerProcessed"][date_str]["hktvLeaderExcluded"] = have_positions
        result["manpower"][date_str] = oix_manpower_correction(df, df_raw, have_positions)
        after = _manpower_totals(log[date_str])
        diff = ", ".join(f"{g} {before.get(g)}→{after[g]}" for g in after if before.get(g) != after[g]) or "figures unchanged"
        print(f"  🩹 OIX manpower backfill {date_str} [{'; '.join(reasons)}]: {diff}")
        result["updated"].append(date_str)
    return result


# -----------------------------------------------------------------------------
# v10.4 — everything calculated from a day's manpower follows an OIX backfill.
# -----------------------------------------------------------------------------
def oix_manpower_correction(df, df_raw, have_positions):
    """The corrected manpower of ONE OIX day (df = process_oix() output), in the shapes the stores below use.
    HKTV Staff manpower needs the leader-exclusion, i.e. the staff list: without it (have_positions False)
    hktvStaff / courier / driver are None and the HKTV figures already on record are left alone — writing the
    un-excluded headcount would inflate the denominator and understate productivity (the original bug)."""
    hktv = manpower_for_group(df, ("LF", "LP"), exclude_positions=LEADER_EXCLUDE_POSITIONS) if have_positions else None
    return {
        "hktvStaff": hktv,
        "odsRatio": manpower_for_group(df, ("ODS", "VAN")),
        "courier": manpower_distribution_for_group(df, "courier") if have_positions else None,
        "driver": manpower_distribution_for_group(df, "driver") if have_positions else None,
        # fallback-only order / waybill counts (Tableau is the primary source) — only used for manpower_staging.json
        "odsOrderCount": order_count_for_group(df, ("ODS", "VAN")),
        "odsWaybillCount": waybill_count_for_group(df, ("ODS", "VAN")),
    }


def _plog_manpower(entry, group):
    g = (entry or {}).get(group) or {}
    return {"overall": (g.get("total") or {}).get("manpower"),
            "districts": {d: (g.get("districts", {}).get(d) or {}).get("manpower") for d in DISTRICTS}}


def _mtd_manpower(log, group, month_key, up_to, override=None):
    """Sum of a group's logged daily manpower from the 1st of month_key through up_to (inclusive).
    `override` {date: manpower dict} swaps in a day's figure (used to price the 'before' picture)."""
    overall, per_d, any_data = 0, {d: 0 for d in DISTRICTS}, False
    for ds, entry in log.items():
        if not ds.startswith(month_key) or ds > up_to:
            continue
        mp = (override or {}).get(ds) or _plog_manpower(entry, group)
        if mp["overall"] is None and all(v is None for v in mp["districts"].values()):
            continue
        overall += mp["overall"] or 0
        for d in DISTRICTS:
            per_d[d] += mp["districts"].get(d) or 0
        any_data = True
    return (overall, per_d) if any_data else (None, {d: None for d in DISTRICTS})


def _rewrite_group_manpower(group_entry, mp):
    """Puts a corrected manpower into one dailyProductivityLog group entry and re-derives productivity =
    orderCount / manpower for it. A hktvStaff day whose productivity comes from the Daily Cost Report
    (productivitySource == 'costReport') keeps that real figure. Returns True when anything changed."""
    changed = False
    def apply(rec, manpower):
        nonlocal changed
        if rec is None:
            return
        if rec.get("manpower") != manpower:
            rec["manpower"] = manpower
            changed = True
        if rec.get("productivitySource") == "costReport":
            return
        orders = rec.get("orderCount")
        new_prod = round(orders / manpower, 2) if manpower and orders is not None else None
        if rec.get("productivity") != new_prod:
            rec["productivity"] = new_prod
            changed = True
    for d in DISTRICTS:
        apply(group_entry.get("districts", {}).get(d), mp["districts"].get(d, 0))
    apply(group_entry.get("total"), sum(mp["districts"].get(d, 0) for d in DISTRICTS))
    return changed


def apply_manpower_corrections(history, corrections, matrices=None, staging=None):
    """v10.4 — propagates OIX-backfilled manpower (corrections = {date: oix_manpower_correction()}) through every
    figure that is calculated from manpower. Mutates `history`, `matrices` (data.json's, optional) and `staging`
    (manpower_staging.json's dict, optional); the caller saves. Returns a summary dict.

      1. dailyProductivityLog[date]  hktvStaff / odsRatio  manpower + productivity (orders / manpower)
      2. the 7-day rolling series history['hktvStaff'|'odsRatio'][date] (Cost-Report days: reapply_cost_report_fields)
      3. productivityMtdLog — every month-to-date snapshot from the earliest corrected day of its month onward.
         HKTV: MTD productivity = Tableau MTD orders / MTD manpower. Manpower is the only thing that changed, so the
         MTD ORDER numerator is held fixed — read from history['productivityMtdBasis'] (stored since v10.4), or, for
         snapshots older than that, backed out as  stored productivity x the manpower sum as it was BEFORE this
         correction — and divided by the corrected MTD manpower.
      4. matrices (data.json): HKTV Staff actual (= the as-of day's snapshot) + forecast, ODS Ratio actual (MTD log
         orders / MTD log manpower) + forecast
      5. staging: if the staged T-1 is a corrected day, its manpower (and fallback ODS counts) are replaced
    A day with no dailyProductivityLog entry is skipped here: backfill_productivity_log() builds it from the fresh
    OIX file at the next 14:00 run."""
    summary = {"days": [], "mtdSnapshots": [], "staging": False, "noLogEntry": []}
    corrections = {ds: c for ds, c in (corrections or {}).items() if c}
    plog = history.setdefault("dailyProductivityLog", {})

    # --- the 'before' picture of every group's per-day manpower, taken BEFORE anything is overwritten
    old_mp = {g: {ds: _plog_manpower(e, g) for ds, e in plog.items()} for g in ("hktvStaff", "odsRatio")}

    # --- 1. per-day rows
    for ds in sorted(corrections):
        corr, entry, did = corrections[ds], plog.get(ds), False
        if not entry:
            summary["noLogEntry"].append(ds)
            continue
        for g in ("hktvStaff", "odsRatio"):
            mp = corr.get(g)
            if mp is not None and entry.get(g):
                did = _rewrite_group_manpower(entry[g], mp) or did
        if did:
            summary["days"].append(ds)
            print(f"  🩹 Productivity log {ds}: HKTV manpower {old_mp['hktvStaff'][ds]['overall']}→{_plog_manpower(entry, 'hktvStaff')['overall']}, "
                  f"ODS/VAN manpower {old_mp['odsRatio'][ds]['overall']}→{_plog_manpower(entry, 'odsRatio')['overall']} "
                  f"(productivity re-derived)")

    # --- 2. 7-day rolling series
    for ds in summary["days"]:
        for g in ("hktvStaff", "odsRatio"):
            ge = plog[ds].get(g) or {}
            if g == "hktvStaff" and ds in history.get("costReportDaily", {}):
                continue                       # real Cost-Report productivity — put back below
            tot = ge.get("total") or {}
            if tot.get("productivity") is None and tot.get("orderCount") is None:
                continue                       # nothing to derive from (HKTV order count unrecoverable) — leave the series as is
            history.setdefault(g, {})[ds] = {
                "overall": tot.get("productivity"),
                "districts": {d: (ge.get("districts", {}).get(d) or {}).get("productivity") for d in DISTRICTS},
            }
    if history.get("costReportDaily"):
        reapply_cost_report_fields(history)

    # --- 3. month-to-date snapshots (HKTV)
    mtd_log = history.setdefault("productivityMtdLog", {})
    basis = history.setdefault("productivityMtdBasis", {})
    first_changed = {}
    for ds in summary["days"]:
        first_changed.setdefault(ds[:7], ds)
    for month_key, first_ds in first_changed.items():
        for snap in sorted(s for s in mtd_log if s.startswith(month_key) and s >= first_ds):
            old_overall, old_d = _mtd_manpower(plog, "hktvStaff", month_key, snap, override=old_mp["hktvStaff"])
            new_overall, new_d = _mtd_manpower(plog, "hktvStaff", month_key, snap)
            stored = mtd_log[snap]
            b = basis.get(snap)
            if b is not None:               # exact numerator stored by the 14:00 job (v10.4+)
                o_overall, o_d = b.get("overall"), b.get("districts") or {}
            else:                           # older snapshot: back the numerator out of the stored ratio
                def back_out(prod, m_old):
                    return prod * m_old if prod is not None and m_old else None
                o_overall = back_out(stored.get("overall"), old_overall)
                o_d = {d: back_out(stored.get("districts", {}).get(d), old_d.get(d)) for d in DISTRICTS}
                basis[snap] = {"overall": o_overall, "districts": o_d}      # keep it so the next refresh is exact
            before_overall = stored.get("overall")
            stored["overall"] = round(o_overall / new_overall, 2) if o_overall is not None and new_overall else None
            stored["districts"] = {d: (round(o_d[d] / new_d[d], 2) if o_d.get(d) is not None and new_d.get(d) else None)
                                   for d in DISTRICTS}
            summary["mtdSnapshots"].append(snap)
            print(f"  🩹 MTD productivity snapshot {snap}: overall {before_overall}→{stored['overall']}")

    # --- 4. data.json matrices
    if matrices is not None and summary["days"]:
        refresh_productivity_matrices(history, matrices)

    # --- 5. staging
    if staging is not None and staging.get("date") in corrections:
        c = corrections[staging["date"]]
        if c.get("hktvStaff") is not None:
            staging["hktvStaffManpower"] = c["hktvStaff"]
            staging["courierGroup"], staging["driverGroup"] = c["courier"], c["driver"]
        staging["odsRatioManpower"] = c["odsRatio"]
        staging["odsOrderCount"], staging["odsWaybillCount"] = c["odsOrderCount"], c["odsWaybillCount"]
        summary["staging"] = True
        print(f"  🩹 {MANPOWER_STAGING_PATH} ({staging['date']}) refreshed from the updated OIX file")
    return summary


def refresh_productivity_matrices(history, matrices):
    """Re-derives data.json's HKTV Staff / ODS Ratio matrices from the (corrected) history: the as-of day's
    month-to-date actual + the 7-day rolling forecast (same as finish_productivity_with_orders())."""
    plog = history.get("dailyProductivityLog", {})
    for key in ("hktvStaff", "odsRatio"):
        m = matrices.get(key)
        if not m or not m.get("asOf"):
            continue
        as_of, month_key = m["asOf"], m["asOf"][:7]
        if key == "hktvStaff":
            snap = history.get("productivityMtdLog", {}).get(as_of)
            if snap:
                m["actual"] = {"overall": snap["overall"], "districts": dict(snap["districts"])}
        else:       # ODS/VAN: MTD orders from the log / MTD manpower from the log — exactly how the 14:00 job builds it
            mp_overall, mp_d = _mtd_manpower(plog, "odsRatio", month_key, as_of)
            o_overall, o_d, any_o = 0, {d: 0 for d in DISTRICTS}, False
            for ds, e in plog.items():
                if ds.startswith(month_key) and ds <= as_of:
                    g = e.get("odsRatio") or {}
                    o_overall += (g.get("total") or {}).get("orderCount") or 0
                    for d in DISTRICTS:
                        o_d[d] += (g.get("districts", {}).get(d) or {}).get("orderCount") or 0
                    any_o = True
            if any_o:
                m["actual"] = {
                    "overall": round(o_overall / mp_overall, 2) if mp_overall else None,
                    "districts": {d: (round(o_d[d] / mp_d[d], 2) if mp_d.get(d) else None) for d in DISTRICTS},
                }
        fc_overall, fc_districts = rolling_average(history, key, 7, dt.date.today())
        m["forecast"] = {"overall": fc_overall, "districts": fc_districts}


def persist_manpower_corrections(history, res):
    """Runs apply_manpower_corrections() for a backfill result and writes every file it touched (the callers
    save history.json themselves): productivity_history.json, data.json, manpower_staging.json."""
    corrections = (res or {}).get("manpower") or {}
    if not corrections:
        return None
    payload = load_data_json()
    staging = load_manpower_staging()
    summary = apply_manpower_corrections(history, corrections, matrices=payload.get("matrices"), staging=staging)
    if summary["days"]:
        save_productivity_history(trimmed_daily_productivity(history, DAILY_PRODUCTIVITY_KEEP_DAYS),
                                  trimmed_productivity_mtd(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
        if payload.get("matrices"):
            save_data_json(payload)
    if summary["staging"]:
        with open(MANPOWER_STAGING_PATH, "w", encoding="utf-8") as f:
            json.dump(staging, f, ensure_ascii=False, indent=2)
        print(f"Wrote {MANPOWER_STAGING_PATH}")
    if summary["noLogEntry"]:
        print(f"  ℹ️ No productivity-log entry yet for {summary['noLogEntry']} — the 14:00 job builds them from the updated OIX file.")
    return summary



def run_oix_backfill_section(days=None, force=False):
    """v10.3 — `--section oixbackfill`: run the OIX manpower backfill on its own — run it (or schedule it) after
    an OIX_Record file has been updated on T+0. By default only files whose modified time is today are checked;
    it refreshes the served JSON files and the dashboard snapshot."""
    print("🚀 檢查 OIX Record 是否有更新，回補 HKTV Manpower Distribution...")
    history = load_history()
    res = backfill_manpower_from_oix(history, days=days, force=force)
    if not res["updated"]:
        print(f"✅ Nothing to backfill ({len(res['unchanged'])} day(s) up to date, {len(res['skipped'])} skipped).")
        save_history(history)       # a first-time fingerprint record may still have been added
        return
    save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
    persist_manpower_corrections(history, res)       # v10.4 — productivity / MTD / data.json follow the corrected manpower
    save_history(history)
    update_embedded_data()
    run_deploy_hook()
    print(f"✅ Manpower Distribution + dependent productivity figures backfilled for {len(res['updated'])} day(s): {res['updated']}")


# =============================================================================
# 6. data.json 讀寫 + history（滾動平均） + forecast 共用邏輯
# =============================================================================

def backfill_productivity_log(history, target_date, position_map=None):
    """v4.0.2 (trial-run adjustment #3) — fills gaps in
    history["dailyProductivityLog"] for the current month, so a day the
    03:00 OIX job missed (staff list not ready, box crashed, OIX_Record not
    downloaded yet, etc. — exactly the 2026-08-31 / 09-01 / 09-02 gap seen in
    productivity_history.json) doesn't silently under-count this month's
    MTD accumulation forever.

    Why this backfills ODS/VAN's figures (order count, waybill count,
    manpower), HKTV's manpower, AND (v19.0 §2) HKTV's waybill count — but
    still never HKTV's order count:
    ODS/VAN's order and waybill counts are entirely OIX-derived
    (order_count_for_group() / waybill_count_for_group()), so any day whose
    OIX_Record file is still sitting in OIX_FOLDER can be fully
    reconstructed after the fact, no Tableau data required. HKTV's order
    count, by contrast, is Tableau's network-wide total for that day MINUS
    the OIX-derived ODS figure (split_hktv_ods_totals()) — and Tableau's
    Delivery Dashboard reports only ever give us T-1 (one day) or MTD (one
    pre-summed total), never a re-queryable per-day total for an arbitrary
    past date. So there is nothing to back that half out of once the day
    has rolled past T-1, and orderCount/productivity stay None for the
    hktvStaff side — honest about what can't be recovered.

    HKTV's waybill count is different: it doesn't need Tableau's network
    total at all. Like ODS/VAN's, it can be counted straight from OIX with
    waybill_count_for_group(df, ("LF", "LP")) — unique Waybill Number
    (Column G) per district, for HKTV's own user prefixes. v19.0 §2 adds
    this recovery, using process_oix(..., dedupe=False) (see its docstring)
    rather than the (Parent Order, User)-deduped `df` used for manpower/
    order-count above, since that dedupe silently drops every waybill past
    the first whenever one order/user pair carries several — HKTV orders
    routinely do, so the corrected count is substantially higher than a
    naive count on the deduped frame would be.

    This is exactly what the Overview tab's HKTV Staff Productivity actually
    needs, though: its MTD "actual" is (Tableau's single MTD order total) −
    (ODS/VAN's MTD order count, SUMMED FROM THIS LOG) ÷ (MTD manpower,
    ALSO SUMMED FROM THIS LOG) — see finish_productivity_with_orders() below.
    It was never built by summing daily HKTV order counts, so recovering
    ODS's per-day figures (which this function does) plus both groups'
    per-day manpower is exactly enough to fix the MTD accumulation; the
    unrecoverable HKTV per-day order count/productivity only affects that
    one day's own row in the Productivity Detail / Daily Records table (and
    the 7-day rolling forecast, which already tolerates missing days).

    `position_map` is passed through from the caller so a single staff list
    load can be reused across every date backfilled in one run, instead of
    re-reading the Excel file per missing day.
    """
    log = history.setdefault("dailyProductivityLog", {})
    manpower_log = history.setdefault("manpowerDistributionLog", {})
    month_key = target_date.strftime("%Y-%m")
    first_of_month = target_date.replace(day=1)

    missing_dates = []
    d = first_of_month
    while d < target_date:  # target_date itself is filled by the normal T-1 flow right after this
        date_str = d.isoformat()
        if date_str.startswith(month_key) and date_str not in log:
            missing_dates.append(d)
        d += dt.timedelta(days=1)

    if not missing_dates:
        return

    if position_map is None:
        try:
            position_map = load_staff_position_map()
        except FileNotFoundError as e:
            print(f"  ⚠️ Productivity backfill: {e} — leader-exclusion / Courier-Driver "
                  f"classification will be skipped for any day recovered below.")
            position_map = None

    filled, unavailable = [], []
    for missing_date in missing_dates:
        try:
            path = find_oix_file(missing_date)
        except FileNotFoundError:
            unavailable.append(missing_date.isoformat())
            continue

        try:
            df_loaded = load_oix(path)
            df = process_oix(df_loaded, position_map)
            hktv_manpower = manpower_for_group(df, ("LF", "LP"), exclude_positions=LEADER_EXCLUDE_POSITIONS)
            ods_manpower = manpower_for_group(df, ("ODS", "VAN"))
            ods_order_count = order_count_for_group(df, ("ODS", "VAN"))
            ods_waybill_count = waybill_count_for_group(df, ("ODS", "VAN"))
            courier_group = manpower_distribution_for_group(df, "courier")
            driver_group = manpower_distribution_for_group(df, "driver")
            # v19.0 §2 — HKTV waybill count, recovered from the SAME OIX file
            # but on a separately-processed, un-deduped frame (see
            # process_oix()'s dedupe=False docstring) so multi-waybill orders
            # aren't collapsed to one. Order count/productivity stay
            # unrecoverable (see docstring above) — this only fills waybillCount.
            df_for_waybill = process_oix(df_loaded, position_map, dedupe=False)
            hktv_waybill_count = waybill_count_for_group(df_for_waybill, ("LF", "LP"))
        except Exception as e:
            print(f"  ⚠️ Productivity backfill: found {os.path.basename(path)!r} for "
                  f"{missing_date.isoformat()} but failed to parse it ({e}) — skipping this date.")
            unavailable.append(missing_date.isoformat())
            continue

        # ODS/VAN: fully real, OIX-derived — orderCount/waybillCount/manpower/productivity all populate.
        ods_entry = _group_to_log_shape(ods_manpower, {
            "overall": {"order": ods_order_count["overall"], "waybill": ods_waybill_count["overall"]},
            "districts": {dist: {"order": ods_order_count["districts"][dist],
                                  "waybill": ods_waybill_count["districts"][dist]} for dist in DISTRICTS},
        })
        # HKTV Staff: manpower is real; waybillCount is now recovered too
        # (v19.0 §2, see docstring above); order count/productivity stay
        # None — Tableau's per-day network total can't be recovered after T-1.
        hktv_entry = {
            "districts": {
                dist: {"orderCount": None,
                       "waybillCount": hktv_waybill_count["districts"].get(dist),
                       "manpower": hktv_manpower["districts"].get(dist, 0), "productivity": None}
                for dist in DISTRICTS
            },
            "total": {"orderCount": None, "waybillCount": hktv_waybill_count["overall"],
                      "manpower": hktv_manpower["overall"], "productivity": None},
        }
        log[missing_date.isoformat()] = {"hktvStaff": hktv_entry, "odsRatio": ods_entry}

        # Also backfill the ODS/VAN 7-day rolling-forecast series — this half
        # is fully computable from OIX alone, same as the entry above.
        ods_daily_district = {
            dist: (round(ods_order_count["districts"][dist] / ods_manpower["districts"][dist], 2)
                   if ods_manpower["districts"].get(dist) else None)
            for dist in DISTRICTS
        }
        ods_daily_overall = (round(ods_order_count["overall"] / ods_manpower["overall"], 2)
                              if ods_manpower["overall"] else None)
        append_history(history, "odsRatio", missing_date.isoformat(), ods_daily_overall, ods_daily_district)

        # HKTV Manpower Distribution tab — same gap, same fix, straight from OIX.
        # v5.0 §1 — backfilled days also get the ODS/VAN group, same as the
        # normal 03:00 run, so a recovered day isn't missing that group later.
        if missing_date.isoformat() not in manpower_log:
            manpower_log[missing_date.isoformat()] = {
                "courier": courier_group, "driver": driver_group,
                "odsVan": odsvan_manpower_group(ods_manpower),
            }

        filled.append(missing_date.isoformat())

    if filled:
        print(f"  🩹 Productivity backfill: recovered {len(filled)} missing day(s) from OIX_Record "
              f"files still in {OIX_FOLDER!r} (ODS/VAN order+waybill+manpower, HKTV manpower+waybill "
              f"— HKTV order count/productivity still unrecoverable, see "
              f"backfill_productivity_log() docstring): {filled}")
    if unavailable:
        print(f"  ⚠️ Productivity backfill: {len(unavailable)} day(s) this month still have no "
              f"dailyProductivityLog entry AND no OIX_Record file left in {OIX_FOLDER!r} to "
              f"recover them from — MTD accumulation for this month is missing these days "
              f"permanently unless that file resurfaces: {unavailable}")


def finish_productivity_with_orders(history, matrices, staging, order_totals_t1, order_totals_mtd, target_date):
    """v4.0 §1/§4, corrected by v4.0.1 §2/§5 — runs inside the 14:00 Tableau
    job, once the new order-count source is available. Picks up staging
    (manpower AND the OIX-derived ODS/VAN order+waybill counts from the
    03:00 OIX run, for the SAME target_date — see save_manpower_staging())
    and combines it with the Tableau-sourced network-wide order/waybill
    totals:
      - ODS/VAN's order/waybill counts are the OIX-derived figures, as-is.
      - HKTV's order/waybill counts are the Tableau total MINUS the
        OIX-derived ODS figure, per district and overall (split_hktv_ods_totals()) —
        restores the pre-v4.0 approach where each group carries its own
        genuine count instead of both sharing the same raw Tableau total.
      - appends today's DAILY order/manpower productivity into history[key]
        (still used for the unchanged 7-day rolling forecast)
      - computes the Overview tab's "actual" as an MTD figure: month-to-date
        order count (HKTV: Tableau MTD total minus ODS's OIX-derived MTD
        sum; ODS: that OIX-derived MTD sum itself, summed from each day's
        entry already logged this month in dailyProductivityLog) ÷
        cumulative month-to-date manpower
      - logs the Productivity Detail / Daily Records entry, now including
        waybillCount alongside orderCount/manpower/productivity, each
        group's own figures (§1/§5's updated 4-row cell layout: Order /
        Waybill / Manpower / Productivity)
    Returns the trimmed productivity_history.json-ready dict.
    """
    # v4.0.2 — recover any earlier-this-month gap in dailyProductivityLog
    # BEFORE summing MTD below, so a day the 03:00 job missed doesn't
    # silently under-count this run's MTD actual (see docstring above).
    backfill_productivity_log(history, target_date)

    hktv_manpower = staging["hktvStaffManpower"]
    ods_manpower = staging["odsRatioManpower"]
    ods_order_count = staging["odsOrderCount"]
    ods_waybill_count = staging["odsWaybillCount"]

    # v4.0.1 §2/§5 — split today's Tableau network totals into HKTV's and
    # ODS/VAN's own order+waybill figures.
    hktv_order_totals, ods_order_totals = split_hktv_ods_totals(order_totals_t1, ods_order_count, ods_waybill_count)
    group_order_totals = {"hktvStaff": hktv_order_totals, "odsRatio": ods_order_totals}

    # --- Daily append, for the unchanged 7-day rolling forecast ---
    for key, manpower_group in (("hktvStaff", hktv_manpower), ("odsRatio", ods_manpower)):
        order_totals_for_group = group_order_totals[key]
        daily_district = {}
        for d in DISTRICTS:
            manpower = manpower_group["districts"].get(d)
            order_count = order_totals_for_group["districts"][d]["order"]
            daily_district[d] = round(order_count / manpower, 2) if manpower and order_count is not None else None
        overall_manpower = manpower_group["overall"]
        overall_order = order_totals_for_group["overall"]["order"]
        daily_overall = round(overall_order / overall_manpower, 2) if overall_manpower and overall_order is not None else None
        append_history(history, key, target_date.isoformat(), daily_overall, daily_district)

    # --- Productivity Detail / Daily Records (Order / Waybill / Manpower /
    # Productivity) — logged BEFORE the MTD sum below so today's own entry is
    # included in month-to-date manpower AND month-to-date ODS order/waybill
    # on the very first run of the month (and on every run thereafter). ---
    append_daily_productivity_log(history, target_date.isoformat(), hktv_manpower, ods_manpower,
                                   hktv_order_totals, ods_order_totals)
    # v8.0 §2 — the line above rebuilt T-1's whole entry from OIX/Tableau; if the Daily
    # Cost Report already covers that day, put the real Cost/Order & Productivity back.
    reapply_cost_report_fields(history)

    # --- MTD actual for the Overview tab ---
    month_key = target_date.strftime("%Y-%m")

    def manpower_mtd(group_key):
        log = history.get("dailyProductivityLog", {})
        total_overall, total_d, any_data = 0, {d: 0 for d in DISTRICTS}, False
        for date_str, entry in log.items():
            if not date_str.startswith(month_key):
                continue
            g = entry.get(group_key, {})
            total_overall += g.get("total", {}).get("manpower") or 0
            for d in DISTRICTS:
                total_d[d] += g.get("districts", {}).get(d, {}).get("manpower") or 0
            any_data = True
        return (total_overall, total_d) if any_data else (None, {d: None for d in DISTRICTS})

    def orders_mtd_from_log(group_key):
        """v4.0.1 §2/§5 — sums a group's own per-day orderCount entries from
        dailyProductivityLog for the current month (same walk pattern as
        manpower_mtd() above). Used to build ODS/VAN's OIX-derived MTD order
        count — HKTV's MTD figure is then the Tableau MTD total minus this.
        Note: only days logged AFTER this fix is deployed carry each group's
        own genuine order count; older days in the log (logged under the old
        shared-Tableau-total behavior) will still be off until they roll out
        of the window naturally."""
        log = history.get("dailyProductivityLog", {})
        total_overall, total_d = 0, {d: 0 for d in DISTRICTS}
        for date_str, entry in log.items():
            if not date_str.startswith(month_key):
                continue
            g = entry.get(group_key, {})
            total_overall += g.get("total", {}).get("orderCount") or 0
            for d in DISTRICTS:
                total_d[d] += g.get("districts", {}).get(d, {}).get("orderCount") or 0
        return total_overall, total_d

    tableau_orders_mtd_overall = order_totals_mtd["overall"]["order"]
    tableau_orders_mtd_district = {d: order_totals_mtd["districts"][d]["order"] for d in DISTRICTS}
    ods_orders_mtd_overall, ods_orders_mtd_district = orders_mtd_from_log("odsRatio")

    group_orders_mtd = {
        "odsRatio": (ods_orders_mtd_overall, ods_orders_mtd_district),
        "hktvStaff": (
            (tableau_orders_mtd_overall - ods_orders_mtd_overall) if tableau_orders_mtd_overall is not None else None,
            {d: (tableau_orders_mtd_district[d] - ods_orders_mtd_district[d])
                if tableau_orders_mtd_district[d] is not None else None
             for d in DISTRICTS},
        ),
    }

    for key in ("hktvStaff", "odsRatio"):
        mp_overall, mp_district = manpower_mtd(key)
        orders_overall, orders_district = group_orders_mtd[key]
        actual_overall = (round(orders_overall / mp_overall, 2)
                           if mp_overall and orders_overall is not None else None)
        actual_district = {
            d: (round(orders_district[d] / mp_district[d], 2)
                if mp_district.get(d) and orders_district.get(d) is not None else None)
            for d in DISTRICTS
        }
        if key == "hktvStaff":
            # Per-day snapshot of the Tableau-based MTD productivity, so the dashboard's
            # date picker can read the SAME figure for a past date instead of re-deriving
            # it from summed daily order rows (which drift from Tableau's MTD total).
            history.setdefault("productivityMtdLog", {})[target_date.isoformat()] = {
                "overall": actual_overall,
                "districts": dict(actual_district),
            }
            # v10.4 — the MTD ORDER numerator behind that figure, so a later OIX manpower backfill can re-divide it
            # by the corrected manpower exactly (see apply_manpower_corrections()).
            history.setdefault("productivityMtdBasis", {})[target_date.isoformat()] = {
                "overall": orders_overall,
                "districts": dict(orders_district),
            }
        fc_overall, fc_districts = rolling_average(history, key, 7, dt.date.today())
        matrices[key] = {
            "actual": {"overall": actual_overall, "districts": actual_district},
            "forecast": {"overall": fc_overall, "districts": fc_districts},
            "asOf": target_date.isoformat(),
        }

    return trimmed_daily_productivity(history, DAILY_PRODUCTIVITY_KEEP_DAYS)


def load_data_json():
    if os.path.exists(DATA_JSON_PATH):
        with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"generatedAt": None, "matrices": {}}


def save_data_json(payload):
    payload["generatedAt"] = dt.datetime.now().isoformat(timespec="minutes")
    os.makedirs(os.path.dirname(DATA_JSON_PATH) or ".", exist_ok=True)
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {DATA_JSON_PATH}")


def update_embedded_data():
    """v28.0 — splices every JSON file this run just wrote into a
    `window.__EMBEDDED_DATA__ = {...}` <script> tag, in place, inside
    INDEX_HTML_PATH ("index.html") itself, between the
    `<!-- EMBEDDED_DATA_START -->` / `_END` marker comments that file's
    template carries for this purpose (see the comment beside those markers,
    and the fetchJSON() shim in loadDataSource() that tries a real fetch()
    first and falls back to window.__EMBEDDED_DATA__ only if that fails).

    v26.0–v27.x wrote this into a separate dashboard.html so a live,
    server-fetched template (index.html) and a self-contained offline copy
    (dashboard.html) could coexist. v28.0 merged them: index.html now does
    both jobs itself (fetch first, embedded snapshot as fallback), so there's
    one deployable file instead of two. This function only ever rewrites the
    text between the two marker comments — everything else in index.html
    (markup, CSS, the rest of the <script>) round-trips untouched, so
    hand-editing the file between pipeline runs is still safe; the next run
    just re-splices a fresh data block into whatever is there at the time.

    Soft-fails (prints a warning, doesn't raise) if index.html is missing or
    doesn't have the marker comments yet, or if any of the six source JSON
    files isn't there to embed — those are the same files this run's other
    save_*() calls just wrote, so a missing one usually just means an earlier
    step in this run failed and already printed its own warning above."""
    if not os.path.exists(INDEX_HTML_PATH):
        print(f"  ⚠️ {INDEX_HTML_PATH!r} not found — skipping embedded-data update (nothing to embed data into).")
        return
    with open(INDEX_HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    start_marker = "<!-- EMBEDDED_DATA_START"
    end_marker = "<!-- EMBEDDED_DATA_END -->"
    start_idx = html.find(start_marker)
    end_idx = html.find(end_marker)
    if start_idx == -1 or end_idx == -1:
        # v9.3 - self-heal: index.html often carries the bare `<script>window.__EMBEDDED_DATA__ = {...};</script>`
        # block without the marker comments (they get lost when the file is re-saved/merged by hand), which made
        # this step skip silently on EVERY run. Wrap the existing block in markers and carry on.
        m = re.search(r"<script>\s*window\.__EMBEDDED_DATA__\s*=.*?</script>", html, re.S)
        if not m:
            print(f"  ⚠️ {INDEX_HTML_PATH!r} has neither EMBEDDED_DATA_START/END markers nor a "
                  f"window.__EMBEDDED_DATA__ script block — skipping embedded-data update.")
            return
        html = (html[:m.start()] + "<!-- EMBEDDED_DATA_START (auto-managed by kpi_pipeline.py) -->"
                + m.group(0) + "<!-- EMBEDDED_DATA_END -->" + html[m.end():])
        print(f"  🔧 {INDEX_HTML_PATH!r}: EMBEDDED_DATA markers were missing - re-added around the existing data block.")
        start_idx = html.find(start_marker)
        end_idx = html.find(end_marker)
    end_idx += len(end_marker)

    sources = {
        "data.json": DATA_JSON_PATH,
        "productivity_history.json": PRODUCTIVITY_HISTORY_PATH,
        "gmv_history.json": GMV_HISTORY_PATH,
        "manpower_distribution.json": MANPOWER_HISTORY_PATH,
        "delay_history.json": DELAY_HISTORY_PATH,
        "other_aspects_history.json": OTHER_ASPECTS_HISTORY_PATH,
    }
    embedded = {}
    missing = []
    for filename, path in sources.items():
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                embedded[filename] = json.load(f)
        else:
            missing.append(filename)
    if missing:
        print(f"  ⚠️ embedded data: {', '.join(missing)} not found — embedding whatever "
              f"the rest of this run did produce; index.html's own fetchJSON() will "
              f"report those as not-found (both its live fetch and its embedded "
              f"fallback) until a future run supplies them.")

    # ensure_ascii=False keeps the embedded JSON human-diffable; the '</' escape
    # is load-bearing — the HTML tokenizer ends a <script> element on the raw
    # byte sequence "</script" wherever it appears (comment or string alike),
    # so any "</" inside embedded text (e.g. a URL) would otherwise truncate
    # this script tag early. \u2028/\u2029 are valid in JSON strings but are
    # illegal raw line terminators inside a JS string literal.
    payload = json.dumps(embedded, ensure_ascii=False)
    payload = (payload.replace("</", "<\\/")
                       .replace("\u2028", "\\u2028")
                       .replace("\u2029", "\\u2029"))
    # v9.3 - the replaced span [start_idx, end_idx) INCLUDES both marker comments, so the new block must carry them
    # again. Before, they were dropped, so only the first run after adding markers worked and every later run (03:00,
    # 09:00, 14:00) printed "no EMBEDDED_DATA_START/END markers" and left the dashboard's snapshot stale.
    block = ("<!-- EMBEDDED_DATA_START (auto-managed by kpi_pipeline.py; do not edit between the markers) -->"
             f"<script>window.__EMBEDDED_DATA__ = {payload};</script>"
             "<!-- EMBEDDED_DATA_END -->")

    updated_html = html[:start_idx] + block + html[end_idx:]
    with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(updated_html)
    print(f"Updated {INDEX_HTML_PATH} ({len(updated_html):,} bytes, embedded data refreshed)")


def load_history():
    if os.path.exists(HISTORY_PATH):
        with open(HISTORY_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_history(history):
    with open(HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


def trimmed_productivity_mtd(history, keep_days):
    """Recent-window slice of productivityMtdLog (Tableau-based MTD productivity per
    as-of date), same pattern as trimmed_daily_productivity()."""
    log = history.get("productivityMtdLog", {})
    return {d: log[d] for d in sorted(log.keys())[-keep_days:]}


def save_productivity_history(daily_dict, mtd_daily=None):
    """Writes ./public/productivity_history.json as {"daily": {...}} — the
    exact shape the dashboard's loadDataSource() fetches. Recomputed fresh
    from history.json's full log every run (see trimmed_daily_productivity),
    so this file is always a derived, disposable window — same pattern as
    data.json itself, never hand-edited or incrementally patched."""
    os.makedirs(os.path.dirname(PRODUCTIVITY_HISTORY_PATH) or ".", exist_ok=True)
    with open(PRODUCTIVITY_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump({"daily": daily_dict, "mtdDaily": mtd_daily or {}}, f, ensure_ascii=False, indent=2)
    print(f"Wrote {PRODUCTIVITY_HISTORY_PATH}")


def save_manpower_history(daily_dict):
    """Writes ./public/manpower_distribution.json as {"daily": {...}} — the
    HKTV Manpower Distribution tab's data source. Same disposable/derived
    pattern as save_productivity_history() above."""
    os.makedirs(os.path.dirname(MANPOWER_HISTORY_PATH) or ".", exist_ok=True)
    with open(MANPOWER_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump({"daily": daily_dict}, f, ensure_ascii=False, indent=2)
    print(f"Wrote {MANPOWER_HISTORY_PATH}")


def append_history(history, metric_key, date_str, overall_value, district_values):
    history.setdefault(metric_key, {})
    history[metric_key][date_str] = {"overall": overall_value, "districts": district_values}


def rolling_average(history, metric_key, n_days, as_of_date):
    series = history.get(metric_key, {})
    dates = sorted(d for d in series if d <= as_of_date.isoformat())[-n_days:]
    if not dates:
        return None, {d: None for d in DISTRICTS}
    overall_vals = [series[d]["overall"] for d in dates if series[d]["overall"] is not None]
    overall_avg = round(sum(overall_vals) / len(overall_vals), 2) if overall_vals else None
    district_avgs = {}
    for d in DISTRICTS:
        vals = [series[dt_]["districts"].get(d) for dt_ in dates if series[dt_]["districts"].get(d) is not None]
        district_avgs[d] = round(sum(vals) / len(vals), 2) if vals else None
    return overall_avg, district_avgs


def prorate_forecast(actual, data_date, total_days):
    if actual is None:
        return None
    return round(actual * (total_days / data_date), 2)


def combine_missing_lost(a, b, c):
    per_district = {d: round(a["districts"][d] + b["districts"][d] + c["districts"][d], 2) for d in DISTRICTS}
    overall = round(a["overall"] + b["overall"] + c["overall"], 2)
    return {"overall": overall, "districts": per_district}


# =============================================================================
# 7. GMV / Basket Size (v3.0 §3)
# =============================================================================

def append_gmv_history(history, date_str, gmv_group):
    """Durable full daily GMV log, in history.json — same role as
    append_history() for the other metrics, kept as its own top-level key
    ("gmv") since GMV also needs month-closing rollups (build_gmv_monthly)
    that the plain rolling_average()-style metrics don't."""
    history.setdefault("gmv", {})[date_str] = gmv_group


def total_parent_orders_for(history, date_str):
    """Basket Size's denominator ("GMV / Total Parent Order") for date_str,
    read from the SAME dailyProductivityLog entry the Productivity Detail tab
    uses (see append_daily_productivity_log() / §1).

    v4.0.1 §2/§5 correction: hktvStaff and odsRatio each carry their own
    genuine order count again (ODS straight from OIX, HKTV = Tableau's
    network-wide total minus that OIX figure — see split_hktv_ods_totals()),
    so the true network total is the SUM of both groups, not either one
    alone. (This reverts the v4.0-era behavior of reading only "hktvStaff",
    which was only correct while both groups shared one duplicated Tableau
    figure — that's no longer the case.)
    Returns (overall, {district: count}) — None wherever that day's
    productivity processing never ran.
    """
    log = history.get("dailyProductivityLog", {}).get(date_str)
    if not log:
        return None, {d: None for d in DISTRICTS}

    def orders(group_key):
        g = log.get(group_key, {})
        return (g.get("total", {}).get("orderCount"),
                {d: g.get("districts", {}).get(d, {}).get("orderCount") for d in DISTRICTS})

    hk_overall, hk_d = orders("hktvStaff")
    od_overall, od_d = orders("odsRatio")
    overall = (hk_overall or 0) + (od_overall or 0)
    districts = {d: (hk_d.get(d) or 0) + (od_d.get(d) or 0) for d in DISTRICTS}
    return overall, districts


def basket_size(gmv_group, orders_overall, orders_districts):
    """Basket Size = GMV / Total Parent Order (v3.0 §3), per district and
    overall. None wherever either side is missing/zero, rather than a
    misleading 0 or a divide-by-zero."""
    def bs(gmv_val, orders_val):
        return round(gmv_val / orders_val, 2) if gmv_val is not None and orders_val else None
    overall = bs(gmv_group["overall"], orders_overall)
    districts = {d: bs(gmv_group["districts"][d], orders_districts.get(d)) for d in DISTRICTS}
    return {"overall": overall, "districts": districts}


def build_gmv_monthly(history):
    """v3.0 §3 — recomputed fresh from history.json's full "gmv" log every
    run (same disposable-derived-file pattern as productivity_history.json /
    rfidMonthly):
      - "daily": one row per date in the CURRENT (still-open) calendar month
        — the "show it like the Daily Records Tab" requirement.
      - "monthly": one accumulated row per CLOSED month, keyed "YYYY-MM" —
        "after each month, the GMV information can be stacked in 1 row ...
        this is accumulated, please do not remove". Every closed month that
        has ever been logged stays here permanently (there are only ~12/yr,
        so this never needs trimming the way daily logs do).
    Both GMV and Basket Size ($ = GMV, # = GMV ÷ Total Parent Order) are
    included at every granularity.
    """
    gmv_log = history.get("gmv", {})
    if not gmv_log:
        return {"daily": {}, "monthly": {}}

    # v4.0.2 fix (trial-run adjustment #2) — "GMV information for dates after
    # T-1" turned out to be contamination sitting in history["gmv"] itself,
    # not a bug in *today's* write path: backfill_gmv_history() already
    # refuses to backfill anything past its cutoff_date (T-1), but it also
    # only ever FILLS gaps and never removes/overwrites an entry that's
    # already on record — so any future-dated (or otherwise bogus, e.g. a
    # garbled crosstab date label parsed into a bad year) row that made it
    # into history["gmv"] before that cutoff existed, or from any other
    # source, stays there forever and gets faithfully re-served by this
    # function on every single run. Guard against that here, at the point
    # this file is actually built, so a clean gmv_history.json doesn't
    # depend on history["gmv"] having never been contaminated.
    #
    # v4.0.3 correction: the first cut of this guard only dropped dates
    # LATER than today, which still let TODAY ITSELF (T-0) through — and
    # that's exactly what happened: 2026-09-10 (today, at the time of that
    # run) was sitting in history["gmv"] with basketSize entirely null (a
    # dead giveaway of a still-accumulating, not-yet-complete day), because
    # it had been legitimately backfilled on an earlier run when "today"
    # was a later date and 09-10 was that run's genuine T-1. GMV is only
    # ever a FINAL, complete figure as of T-1 — today's own number is still
    # accumulating intraday and is never valid to show, no matter how it
    # got into the log. So the cutoff here now matches every other T-1
    # cutoff in this file (cutoff_date = today - 1 in backfill_gmv_history,
    # gmv_date = today - 1 in run_section_tableau's normal write): any date
    # >= today, or implausibly far outside a +/-2 year window, is dropped
    # from the rollup and removed from history["gmv"] itself, so it stops
    # being carried forward and re-checked on every future run too.
    today = dt.date.today()
    cutoff_date = today - dt.timedelta(days=1)  # T-1 — the last date GMV can ever be "final" for
    valid_year_range = (today.year - 2, today.year + 2)
    bad_dates = []
    for date_str in list(gmv_log.keys()):
        try:
            row_date = dt.date.fromisoformat(date_str)
        except ValueError:
            bad_dates.append(date_str)
            continue
        if row_date > cutoff_date or not (valid_year_range[0] <= row_date.year <= valid_year_range[1]):
            bad_dates.append(date_str)
    if bad_dates:
        for date_str in bad_dates:
            del gmv_log[date_str]
        print(f"  🧹 Dropped {len(bad_dates)} contaminated/future-or-today-dated GMV day(s) "
              f"found in history[\"gmv\"] (later than T-1 {cutoff_date.isoformat()!r} — "
              f"GMV is never final for today itself — or an implausible year): {sorted(bad_dates)}")

    current_month = today.strftime("%Y-%m")
    daily = {}
    sums = {}  # month -> running totals, used to build the closed-month rollup

    for date_str, gmv_group in sorted(gmv_log.items()):
        month = date_str[:7]
        orders_overall, orders_districts = total_parent_orders_for(history, date_str)

        if month == current_month:
            daily[date_str] = {
                "gmv": gmv_group,
                "basketSize": basket_size(gmv_group, orders_overall, orders_districts),
            }

        bucket = sums.setdefault(month, {
            "gmv_overall": 0.0, "gmv_districts": {d: 0.0 for d in DISTRICTS},
            "orders_overall": 0, "orders_districts": {d: 0 for d in DISTRICTS},
        })
        bucket["gmv_overall"] += gmv_group["overall"] or 0
        bucket["orders_overall"] += orders_overall or 0
        for d in DISTRICTS:
            bucket["gmv_districts"][d] += gmv_group["districts"].get(d) or 0
            bucket["orders_districts"][d] += orders_districts.get(d) or 0

    monthly = {}
    for month, b in sums.items():
        if month == current_month:
            continue  # current month stays as daily rows only, per spec
        gmv_group = {
            "overall": round(b["gmv_overall"], 2),
            "districts": {d: round(v, 2) for d, v in b["gmv_districts"].items()},
        }
        monthly[month] = {
            "gmv": gmv_group,
            "basketSize": basket_size(gmv_group, b["orders_overall"], b["orders_districts"]),
        }

    return {"daily": daily, "monthly": monthly}


def append_delay_history(history, date_str, delay_early_t1):
    """v4.0 §2 — durable full daily log for the 'Delay %' tab: one row per
    day, by timeslot (AM/PM/EV/EV2/Overall) and district (+ overall).

    v6.0 fix — `date_str` is the DATA date of the record (T-1), NOT the
    date the pipeline happened to run. The 'Actual Delivery - Delay & Early
    %' report is the DeliverySummary T-1 view, so a run on Monday holds
    Sunday's figures and must file them under Sunday's date. It used to be
    filed under the run date (T+0), which showed as a same-day row on the
    Delay % tab and shifted every later day's label by one."""
    history.setdefault("delayPercentDaily", {})[date_str] = delay_early_t1


def migrate_delay_t1_keys(history):
    """v6.0 fix — ONE-TIME re-keying of the delay history written before
    append_delay_history() was given the T-1 data date.

    Every record in history["delayPercentDaily"] (and its twin in
    history["delayRate"], the series behind the 30-day forecast) was filed
    under the day the pipeline RAN, but holds the previous day's T-1
    figures. So each key is shifted back by one day: a record filed as
    2026-09-21 becomes 2026-09-20.

    history["delayPercentDaily"] is only ever written by
    append_delay_history(), so all of it is run-date keyed and is shifted
    wholesale. history["delayRate"] is mixed — besides those T-1 records it
    holds day-level ESTIMATES that backfill_delay_rate_history() recovered
    from the zone-type file under their TRUE dates — so it is only touched
    where its value matches the delayPercentDaily record for the same key
    (i.e. is provably the T-1 write). Each such entry is moved to the
    corrected date, replacing any zone-average estimate sitting there (the
    real T-1 figure wins), and the old run-date slot is cleared so no
    T+0 row survives. A slot that ends up empty is refilled from the zone
    file by the next run's backfill, exactly like any other missing day.

    Idempotent: guarded by history["_migrations"]["delay_t1_key_shift"], so
    it runs once and is a no-op on every later run (and on a history.json
    that has already been migrated by hand). Returns True if it changed
    anything."""
    marker = history.setdefault("_migrations", {})
    if marker.get("delay_t1_key_shift"):
        return False
    one_day = dt.timedelta(days=1)
    daily = history.get("delayPercentDaily", {})
    rate = history.setdefault("delayRate", {})
    moved_rate = {}
    for old_key, rec in daily.items():
        new_key = (dt.date.fromisoformat(old_key) - one_day).isoformat()
        t1_overall = rec.get("overall", {}).get("Overall", {}).get("delay")
        t1_districts = {d: rec.get("districts", {}).get(d, {}).get("Overall", {}).get("delay") for d in DISTRICTS}
        existing = rate.get(old_key)
        if existing is not None and existing.get("overall") == t1_overall and existing.get("districts") == t1_districts:
            moved_rate[new_key] = existing
            del rate[old_key]
    rate.update(moved_rate)
    shifted = {(dt.date.fromisoformat(k) - one_day).isoformat(): v for k, v in daily.items()}
    if daily:
        history["delayPercentDaily"] = dict(sorted(shifted.items()))
    history["delayRate"] = dict(sorted(rate.items()))
    marker["delay_t1_key_shift"] = dt.date.today().isoformat()
    if daily:
        print(f"  🔧 One-time fix: re-keyed {len(daily)} Delay % day(s) from run date to T-1 data date "
              f"({min(daily)}…{max(daily)} → {min(shifted)}…{max(shifted)}).")
    return bool(daily)


def build_delay_monthly(history, delay_early_mtd, zone_type, mtd_overall_delay):
    """v4.0 §2 — 'Delay %' tab data:
      - "daily": delayPercentDaily entries for the CURRENT (still-open) month
        only — the full log stays in history.json (same disposable/derived
        pattern as productivity_history.json), but the tab itself only needs
        this month plus the closed-month rollup below.
      - "monthly": one accumulated row per CLOSED month (spec: "having the
        monthly record after the end of the month"), keyed "YYYY-MM" — for
        each of the 5 slots (AM/PM/EV/EV2/Overall), overall + per-district
        Delay % is the plain average of that slot's daily values across the
        month. Never removed once a month closes, same as gmv_history.json /
        other_aspects_history.json.
      - "mtd": today's month-to-date snapshot for the dropdown's 5 options
        (AM/PM/EV/EV2 come straight from the MTD file; "Overall" per district
        comes from delay_rate_by_zone_type's Grand-Total row, since the MTD
        delay/early file itself has no per-district Total row — see
        parse_delay_early_pct()).
      - "mtdDaily" (v27.0; full history since v28.0) — history["delayRateMtdDaily"]
        entries for EVERY month recorded, not just the current one: the REAL
        MTD-to-date delay % as the source system reported it on each day it
        was captured, keyed by that date. Lets the dashboard's "Data Date"
        picker look up an earlier day's true MTD figure directly — including
        a day in a month that has since closed — instead of approximating it
        from a plain average of delayPercentDaily's raw per-day values. Only
        exists from the day this log started (v27.0) onward — an earlier
        date has no entry here and the dashboard falls back to the old
        averaging approach for those. history["delayRateMtdDaily"] itself is
        never trimmed, so this was always available; v28.0 just stopped
        filtering it down to the current month before handing it to the
        dashboard.
    """
    current_month = dt.date.today().strftime("%Y-%m")
    full_log = history.get("delayPercentDaily", {})
    daily_log = {k: v for k, v in full_log.items() if k[:7] == current_month}
    mtd_daily_log = dict(history.get("delayRateMtdDaily", {}))

    slots = ["AM", "PM", "EV", "EV2", "Overall"]
    sums = {}  # month -> slot -> {"overall": [...], "districts": {d: [...]}}
    for date_str, rec in full_log.items():
        month = date_str[:7]
        if month == current_month:
            continue
        bucket = sums.setdefault(month, {s: {"overall": [], "districts": {d: [] for d in DISTRICTS}} for s in slots})
        for s in slots:
            v = rec.get("overall", {}).get(s, {}).get("delay")
            if v is not None:
                bucket[s]["overall"].append(v)
            for d in DISTRICTS:
                dv = rec.get("districts", {}).get(d, {}).get(s, {}).get("delay")
                if dv is not None:
                    bucket[s]["districts"][d].append(dv)
    monthly = {}
    for month, by_slot in sums.items():
        monthly[month] = {}
        for s in slots:
            vals = by_slot[s]
            overall = round(sum(vals["overall"]) / len(vals["overall"]), 2) if vals["overall"] else None
            districts = {d: (round(sum(vs) / len(vs), 2) if (vs := vals["districts"][d]) else None) for d in DISTRICTS}
            monthly[month][s] = {"overall": overall, "districts": districts}

    mtd_districts = {}
    for d in DISTRICTS:
        entry = dict(delay_early_mtd["districts"].get(d, {}))
        entry["Overall"] = {"delay": zone_type["overall"].get(d)}
        mtd_districts[d] = entry
    mtd = {
        "overall": {**delay_early_mtd["overall"], "Overall": {"delay": mtd_overall_delay}},
        "districts": mtd_districts,
    }
    return {"daily": daily_log, "monthly": monthly, "mtd": mtd, "mtdDaily": mtd_daily_log}


def save_delay_history(payload):
    os.makedirs(os.path.dirname(DELAY_HISTORY_PATH) or ".", exist_ok=True)
    with open(DELAY_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {DELAY_HISTORY_PATH}")


def build_other_aspects_monthly(history, matrices, payload):
    """'Other Aspects Tracking' tab: Poor Rating %, Missing & Lost Amount,
    RFID Missing Tote.

    v4.0.3 correction (trial-run adjustment): all three of these figures are
    MTD-cumulative already at the source — Delivery Rating.csv is itself an
    MTD report, so is each of the three "Summary By RP" reports that feed
    Missing & Lost Amount, and RFID's own bucket is a running sum of its
    per-day ledger. That means history["poorRating"][date] /
    history["missingLostAmount"][date] were never independent single-day
    figures — each one is "the month's running total as best known on that
    date", and the LAST one recorded in a month IS that month's true final
    total. The previous version of this function didn't know that: it
    averaged poorRating's daily snapshots and summed missingLostAmount's
    across a closed month, which diluted the true final percentage in one
    case and double/triple/quadruple-counted every earlier day's running
    total in the other. Every closed month now uses its LAST recorded
    snapshot instead — the same approach rfidMissingTote's rfidMonthlyClosed
    archive already used correctly, so all three metrics are now consistent
    with each other.

    Also per the trial-run feedback, this file no longer serves day-by-day
    rows at all — only ever the MTD figure, and it's read straight from
    `matrices` (this run's already-computed Overview-tab values) rather than
    re-derived from history, so the two tabs can never disagree:
      - "current": the single OPEN (current) month's MTD-to-date snapshot,
        overwritten in place every run, never accumulated as a per-day log.
      - "monthly": one permanent row per CLOSED month, keyed "YYYY-MM".
    The dashboard itself adds "(MTD)" to "current"'s label.

    v27.0 §2 — "daily" (full history since v28.0): each metric's OWN true
    MTD-to-date value as recorded on every day this pipeline ran, for EVERY
    month on record — not just the current one — so the dashboard's "Data
    Date" picker can look up any earlier day's real figure directly,
    including a day in a month that has since closed, instead of falling
    back to today's snapshot once the month rolls over. poorRating/
    missingLostAmount already had this per-day MTD record sitting in
    history.json the whole time, for every date ever recorded (see the
    correction above — each entry there already IS that day's running MTD
    total, and history[metric_key] is never trimmed); v28.0 just stopped
    filtering it down to the current month before handing it to the
    dashboard. rfidMissingTote has no such per-day MTD record, only a
    per-day ledger of INCREMENTS — that ledger used to live only in
    payload["rfidMonthly"], which is trimmed to the 2 most recent months, so
    a closed month's day-by-day path was lost forever once it aged out.
    v28.0 adds a second, durable copy of every increment
    (history["rfidDailyLedger"], never trimmed) written alongside the
    existing one; "daily" below is built by grouping that durable ledger by
    month and summing each month's entries cumulatively in date order,
    resetting the running total at each month's first entry (the same
    MTD-cumulative convention the other two metrics already have natively).
    A month that closed before v28.0 shipped only has its final total
    (already preserved in "monthly" via rfidMonthlyClosed) — its day-by-day
    path can't be recovered retroactively since the old ledger had already
    been trimmed away by the time this durable copy started being written.
    """
    current_month = dt.date.today().strftime("%Y-%m")

    def closed_months_from_last_snapshot(metric_key):
        """Each closed month's value = its LAST recorded daily snapshot
        (already a complete MTD total for that month by definition — see
        docstring above), not a sum or average across the month's entries."""
        series = history.get(metric_key, {})
        latest_date_for_month = {}
        for date_str in series:
            month = date_str[:7]
            if month == current_month:
                continue
            if month not in latest_date_for_month or date_str > latest_date_for_month[month]:
                latest_date_for_month[month] = date_str
        return {
            month: {"overall": series[date_str].get("overall"),
                    "districts": series[date_str].get("districts", {})}
            for month, date_str in latest_date_for_month.items()
        }

    def current_from_matrix(metric_key):
        mm = matrices.get(metric_key, {})
        m = mm.get("actual", {})
        # v9.4 - RFID is a T-4 metric: its open month is the month of the T-4 date (matrix "monthKey"), which
        # lags the calendar month for the first 4 days. The other metrics have no monthKey -> calendar month.
        return {"monthKey": mm.get("monthKey") or current_month, "overall": m.get("overall"),
                "districts": m.get("districts", {})}

    def full_history_daily_mtd(metric_key):
        """poorRating/missingLostAmount, EVERY month on record: history[metric_key]
        is already a per-day MTD-cumulative log (append_history() records the
        day's full running total, not an increment) and is never trimmed, so
        returning it in full — rather than filtering to the current month —
        costs nothing and is all that's needed."""
        series = history.get(metric_key, {})
        return {
            date_str: {"overall": rec.get("overall"), "districts": rec.get("districts", {})}
            for date_str, rec in series.items()
        }

    def rfid_all_daily_mtd():
        """rfidMissingTote, EVERY month on record: history["rfidDailyLedger"]
        (v28.0, durable, never trimmed) holds every day's own increment,
        keyed by the T-4 date it belongs to. Grouped by month and summed
        cumulatively in date order, resetting to 0 at each month's first
        entry, so each date's entry is "everything recorded up to and
        including that day, within its own month" — matching the
        MTD-cumulative convention the other two metrics have natively.
        Rebuilt fresh every run straight from the ledger, so a late-arriving
        or corrected T-4 entry is reflected at every date on or after it
        within that same month, not just the day it was recorded."""
        ledger = history.get("rfidDailyLedger", {})
        by_month = {}
        for date_str in sorted(ledger.keys()):
            by_month.setdefault(date_str[:7], []).append(date_str)
        out = {}
        for month, date_strs in by_month.items():
            running_overall = 0.0
            running_districts = {d: 0.0 for d in DISTRICTS}
            for date_str in date_strs:
                day = ledger[date_str]
                running_overall += day.get("overall") or 0
                for d in DISTRICTS:
                    running_districts[d] += (day.get("districts") or {}).get(d) or 0
                out[date_str] = {
                    "overall": round(running_overall, 2),
                    "districts": {d: round(v, 2) for d, v in running_districts.items()},
                }
        return out

    return {
        "poorRating": {
            "current": current_from_matrix("poorRating"),
            "monthly": closed_months_from_last_snapshot("poorRating"),
            "daily": full_history_daily_mtd("poorRating"),
        },
        "missingLostAmount": {
            "current": current_from_matrix("missingLostAmount"),
            "monthly": closed_months_from_last_snapshot("missingLostAmount"),
            "daily": full_history_daily_mtd("missingLostAmount"),
        },
        "rfidMissingTote": {
            "current": current_from_matrix("rfidMissingTote"),
            "monthly": {
                k: {"overall": v.get("overall"), "districts": v.get("districts", {})}
                for k, v in history.get("rfidMonthlyClosed", {}).items()
            },
            "daily": rfid_all_daily_mtd(),
        },
        # v9.0 - Fulfillment Cost % (Total Cost / GMV), written by the Daily Cost Report job; kept here so this
        # full rebuild of the file never drops it.
        "fulfillmentCost": build_fulfillment_cost_block(history),
    }


def save_other_aspects_history(payload):
    os.makedirs(os.path.dirname(OTHER_ASPECTS_HISTORY_PATH) or ".", exist_ok=True)
    with open(OTHER_ASPECTS_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {OTHER_ASPECTS_HISTORY_PATH}")


def save_gmv_history(payload):
    os.makedirs(os.path.dirname(GMV_HISTORY_PATH) or ".", exist_ok=True)
    with open(GMV_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {GMV_HISTORY_PATH}")


def update_rfid_missing_tote(payload, history, matrices, today):
    """RFID Missing Tote - a T-4 record. The Tableau extract pulled on day D reports the records of D-4 (the Oct 1 run
    reports Sep 27), so every record belongs to the month of its T-4 date, NOT to the month of the run date.

    v9.4 - before, the bucket key was the run date's month: the Oct 1-4 runs put Sep 27-30 records into "October"
    (the card showed Oct = 14 while September was already archived without them), and the first days of every
    month leaked into the month before. Now:
      * bucket / month_key = month of the T-4 date. September stays the open month until the T-4 date reaches Oct 1
        (the Oct 5 run); only then does October start counting.
      * stored as a per-day ledger keyed by the T-4 date, totals recomputed from it each run (re-runs overwrite
        that day's entry instead of adding to it).
      * one-time self-repair: ledger days filed under the wrong month bucket are moved to the right one, a month
        that was archived too early is re-opened, and an archived month is refreshed from its ledger.
      * forecast = actual / (days of the month the data actually covers = T-4 day-of-month) * days in the month.
    Months that were already closed before v9.4 keep their archived total (their day-level data is gone)."""
    t4_date = today - dt.timedelta(days=4)
    rfid_increment = parse_rfid(t4_date)

    month_key = t4_date.strftime("%Y-%m")
    rfid_total_days = (dt.date(t4_date.year + (t4_date.month == 12), (t4_date.month % 12) + 1, 1)
                       - dt.timedelta(days=1)).day
    rfid_state = payload.setdefault("rfidMonthly", {})
    closed_archive = history.setdefault("rfidMonthlyClosed", {})

    # the month still being counted is, by definition, not closed - undo a premature archive entry
    if closed_archive.pop(month_key, None) is not None:
        print(f"  🔧 RFID: {month_key} was archived as closed too early (T-4 date {t4_date} is still in it) - re-opened.")

    # move ledger days that sit in the wrong month bucket
    moved = {}
    for bk, b in rfid_state.items():
        if isinstance(b.get("days"), dict):
            for ds in [x for x in b["days"] if x[:7] != bk]:
                moved.setdefault(ds[:7], {})[ds] = b["days"].pop(ds)
    for mk, days in moved.items():
        if mk in closed_archive and mk != month_key:
            continue      # that month's total is already archived; the durable rfidDailyLedger still keeps the days
        tgt = rfid_state.setdefault(mk, {}).setdefault("days", {})
        for ds, v in days.items():
            tgt.setdefault(ds, v)
        print(f"  🔧 RFID: moved {len(days)} misfiled day(s) into {mk}: {sorted(days)}")

    bucket = rfid_state.setdefault(month_key, {})
    # One-time migration: older data.json files stored a running total with no per-day ledger; those totals can't be
    # trusted (duplicate same-day additions), so drop them and start a fresh ledger.
    if "overall" in bucket and "days" not in bucket:
        print(f"  ⚠️ {month_key} had an old-style running total with no per-day ledger — resetting it. "
              f"Starting a fresh ledger from today.")
        bucket = {"days": {}}
        rfid_state[month_key] = bucket
    days_ledger = bucket.setdefault("days", {})
    days_ledger[t4_date.isoformat()] = rfid_increment  # overwrite, not add

    # v28.0 - second, DURABLE copy (never trimmed) read by build_other_aspects_monthly()'s rfid_all_daily_mtd().
    history.setdefault("rfidDailyLedger", {})[t4_date.isoformat()] = rfid_increment

    # recompute every bucket that has a ledger; drop empty ones that are not the open month
    for bk in list(rfid_state):
        b = rfid_state[bk]
        if "days" not in b:
            continue
        if not b["days"] and bk != month_key:
            del rfid_state[bk]
            continue
        b["overall"] = round(sum(d["overall"] for d in b["days"].values()), 2)
        b["districts"] = {dist: round(sum(day["districts"][dist] for day in b["days"].values()), 2)
                          for dist in DISTRICTS}
    bucket = rfid_state[month_key]

    matrices["rfidMissingTote"] = {
        "actual": {"overall": bucket["overall"], "districts": bucket["districts"]},
        "forecast": {
            "overall": prorate_forecast(bucket["overall"], t4_date.day, rfid_total_days),
            "districts": {d: prorate_forecast(bucket["districts"][d], t4_date.day, rfid_total_days) for d in DISTRICTS},
        },
        "target": RFID_TOTE_TARGETS,
        "asOf": today.isoformat(),
        "dataThrough": t4_date.isoformat(),      # v9.4 - last day the figures cover (T-4)
        "monthKey": month_key,
    }

    # archive every month that is no longer the open one (refresh from its ledger - authoritative)
    for k, v in rfid_state.items():
        if k != month_key and "days" in v:
            closed_archive[k] = {"overall": v.get("overall"), "districts": v.get("districts", {})}

    # keep the previous month on the card during the first days of a month (when its bucket was already trimmed)
    prev_key = (f"{t4_date.year - 1:04d}-12" if t4_date.month == 1
                else f"{t4_date.year:04d}-{t4_date.month - 1:02d}")
    if prev_key not in rfid_state and prev_key in closed_archive:
        rfid_state[prev_key] = {"overall": closed_archive[prev_key].get("overall"),
                                "districts": closed_archive[prev_key].get("districts", {})}

    keep_keys = sorted(rfid_state.keys())[-2:]
    payload["rfidMonthly"] = {k: rfid_state[k] for k in keep_keys}


def run_section_tableau():
    """No network — reads the 6 CSVs fetch_tableau_reports() already
    downloaded into REPORT_FOLDER, parses them, computes forecasts, writes
    data.json."""
    print("🚀 開始處理 Tableau 數據...")
    # v3.0 §3 fix: GMV ("gmv" -> Sheet 1.csv) is handled as its own soft-fail
    # block further down (see the `if os.path.exists(gmv_path)` block below) —
    # a GMV-account hiccup should never block the other reports that
    # already downloaded fine, so it's excluded from this hard check.
    # Also now checks *freshness*, not just existence — see STALE_REPORT_HOURS.
    # v8.0 — "actual_delivery_10d" (ODS counts) is also soft: if its download failed the
    # ODS figures fall back to the OIX-derived ones staged at 03:00 (see below) instead of
    # taking the whole 15:00 job down.
    required_files = {k: v for k, v in REPORT_FILES.items() if k not in ("gmv", "actual_delivery_10d")}
    # (v4.0: "delay_rate"/Rank_On Time.csv is retired — replaced by the 5
    # Delivery Dashboard files above, already included in REPORT_FILES.)
    cutoff_time = time.time() - STALE_REPORT_HOURS * 3600
    missing, stale = [], []
    for f in required_files.values():
        fpath = os.path.join(REPORT_FOLDER, f)
        if not os.path.exists(fpath):
            missing.append(f)
        elif os.path.getmtime(fpath) < cutoff_time:
            stale.append(f)
    if missing or stale:
        raise FileNotFoundError(
            f"Report file(s) not freshly downloaded this run in {REPORT_FOLDER!r} — "
            f"missing: {missing or 'none'}; stale/leftover from an earlier run whose "
            f"download failed (older than {STALE_REPORT_HOURS}h): {stale or 'none'}. "
            f"fetch_tableau_reports() should have downloaded these just now — check "
            f"pipeline_log.txt for which report(s) failed to download this run."
        )

    today = dt.date.today()
    total_days = (dt.date(today.year + (today.month == 12), (today.month % 12) + 1, 1) - dt.timedelta(days=1)).day

    history = load_history()
    migrate_delay_t1_keys(history)  # v6.0 — one-time, idempotent; must run BEFORE today's T-1 delay record is appended below
    payload = load_data_json()
    matrices = payload.setdefault("matrices", {})

    # v9.0 — Staff Allocation Cap for Suggest Headcount (Full Perspective
    # Forecast tab): recomputed fresh every run straight from whatever is
    # currently the latest staff list file, independently of the 03:00 job's
    # own position_map load below — so the cap always reflects the most
    # up-to-date staff list regardless of when it lands during the day.
    # Soft-fails the same way load_staff_position_map() does elsewhere: a
    # missing/late staff list just means the dashboard keeps yesterday's cap
    # (or none, before the first successful run) rather than blocking the
    # rest of this job.
    try:
        payload["staffAllocationCap"] = compute_staff_allocation_cap()
    except FileNotFoundError as e:
        print(f"  ⚠️ {e} — skipping Staff Allocation Cap update this run; "
              f"Suggest Headcount on the dashboard keeps its previous cap (if any).")

    # --- v4.0 §2/§4: Delay Rate ---
    # "actual" shown on the dashboard is now the MTD figure (spec §4), but the
    # 30-day forecast still rolls up genuine T-1 DAILY values (spec: "No
    # effect on the forecast value calculation") — so we still append today's
    # single-day delay% into the history series used by rolling_average(),
    # completely separately from the MTD "actual" we display.
    # v6.0 fix — the "Actual Delivery - Delay & Early %" file is the T-1 view,
    # so its record belongs to the DATA date (today - 1), not the run date.
    # Filing it under `today` showed a same-day (T+0) row on the Delay % tab
    # and shifted the 30-day forecast window by a day. Same T-1 convention as
    # productivity_target_date and the GMV write below.
    delay_data_date = today - dt.timedelta(days=1)
    delay_early_t1 = parse_delay_early_pct("delay_early")
    t1_overall_delay = delay_early_t1["overall"].get("Overall", {}).get("delay")
    t1_district_delay = {d: delay_early_t1["districts"][d].get("Overall", {}).get("delay") for d in DISTRICTS}
    append_history(history, "delayRate", delay_data_date.isoformat(), t1_overall_delay, t1_district_delay)
    delay_fc_overall, delay_fc_districts = rolling_average(history, "delayRate", 30, today)

    zone_type = parse_delay_rate_by_zone_type()          # per-district Residential/Commercial/Overall, MTD
    mtd_overall_delay = parse_mtd_overall_delay()         # single network-wide MTD headline %

    # v27.0 — durable per-day record of the REAL MTD delay % this run actually
    # saw (network-wide + per-district), keyed by today's run date. Until now
    # this MTD figure was computed fresh every run and only ever shown for
    # "today" — nothing kept yesterday's or last week's true MTD-to-date value
    # around. The dashboard's "Data Date" picker (v7.0 §1) filled that gap by
    # RE-DERIVING an approximate MTD-through-date as a plain average of the
    # raw per-day Overall delay% (delayPercentDaily) up to the chosen date —
    # a reasonable stand-in, but not the actual number the source system
    # reported as of that day, since a plain average of daily percentages
    # isn't necessarily identical to a true volume-weighted MTD rollup. Now
    # that the real figure is captured daily, build_delay_monthly() exposes
    # it as "mtdDaily" and the dashboard prefers a direct lookup there,
    # falling back to the old averaging approach only for dates before this
    # log started.
    history.setdefault("delayRateMtdDaily", {})[today.isoformat()] = {
        "overall": mtd_overall_delay,
        "districts": {d: zone_type["overall"].get(d) for d in DISTRICTS},
    }

    matrices["delayRate"] = {
        "actual": {"overall": mtd_overall_delay, "districts": {d: zone_type["overall"].get(d) for d in DISTRICTS}},
        "forecast": {"overall": delay_fc_overall, "districts": delay_fc_districts},
        "target": DELAY_RATE_TARGET,
        "asOf": today.isoformat(),
    }
    # v4.0 §4 — feeds the Overview tab's district-cell Residential:/Commercial:/
    # Overall: breakdown display.
    matrices["delayRateByZone"] = {
        "residential": {d: zone_type["residential"].get(d) for d in DISTRICTS},
        "commercial": {d: zone_type["commercial"].get(d) for d in DISTRICTS},
        "overall": {d: zone_type["overall"].get(d) for d in DISTRICTS},
        "asOf": today.isoformat(),
    }

    # v4.0 §2 — "Delay %" tab: T-1 daily record (by timeslot + overall, in
    # district basis and overall) plus a monthly rollup after month-end, same
    # daily/monthly pattern as gmv_history.json. MTD-to-date timeslot view
    # (AM/PM/EV/EV2) comes from the MTD file directly; MTD's per-district
    # "Overall" selection reuses zone_type["overall"] computed above, since
    # the MTD delay/early file itself has no per-district Total row (see
    # parse_delay_early_pct docstring).
    delay_early_mtd = parse_delay_early_pct("mtd_delay_early_ontime")
    append_delay_history(history, delay_data_date.isoformat(), delay_early_t1)  # v6.0 — T-1 data date, see delay_data_date above
    save_delay_history(build_delay_monthly(history, delay_early_mtd, zone_type, mtd_overall_delay))

    # v4.0 §3 — backfill any missing GMV / Delay Rate days from whatever
    # extra dates happen to still be sitting in this run's own downloads.
    backfill_delay_rate_history(history)

    # --- v4.0 §1: Productivity (HKTV Staff / ODS Ratio) — now finished here,
    # in the 14:00 job, using the manpower staged by the 03:00 OIX job plus
    # the new Tableau-sourced order/waybill counts (T-1 daily + MTD). ---
    order_totals_t1 = parse_actual_delivery_timeslot("actual_delivery_timeslot")
    order_totals_mtd = parse_actual_delivery_timeslot("actual_delivery_timeslot_mtd")
    productivity_target_date = today - dt.timedelta(days=1)  # T-1, matches the OIX staging date
    # v10.4 — an OIX_Record updated since the 03:00 run is picked up here, BEFORE staging is read, so T-1's staged
    # manpower and every earlier day's manpower-based figures are current when today's productivity is finished.
    try:
        _res = backfill_manpower_from_oix(history)
        if _res.get("updated"):
            apply_manpower_corrections(history, _res["manpower"], matrices=matrices, staging=None)
            _stg = load_manpower_staging()
            if _stg is not None and _stg.get("date") in _res["manpower"]:
                apply_manpower_corrections({}, {_stg["date"]: _res["manpower"][_stg["date"]]}, staging=_stg)
                with open(MANPOWER_STAGING_PATH, "w", encoding="utf-8") as _f:
                    json.dump(_stg, _f, ensure_ascii=False, indent=2)
            save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
    except Exception as e:
        print(f"  ⚠️ OIX manpower backfill before productivity skipped: {e}")
    staging = load_manpower_staging()
    if staging is None:
        print(f"  ⚠️ {MANPOWER_STAGING_PATH!r} not found — the 03:00 Manpower job hasn't run yet "
              f"(or hasn't run since this file was last cleared). Skipping Productivity update "
              f"this run; HKTV Staff / ODS Ratio matrices keep their previous values.")
    elif staging.get("date") != productivity_target_date.isoformat():
        print(f"  ⚠️ Manpower staging is for {staging.get('date')!r} but today's Tableau order "
              f"counts are for {productivity_target_date.isoformat()!r} — dates don't match "
              f"(the 03:00 job may not have run today). Skipping Productivity update this run.")
    else:
        # v8.0 §1 — ODS order / waybill counts now come from Tableau's "Actual Delivery -
        # 10 Districts" sheet; HKTV = Tableau total - this ODS figure (split_hktv_ods_totals()).
        # The OIX-derived counts staged at 03:00 stay only as the fallback.
        try:
            ods_order_t, ods_waybill_t = parse_ods_counts_from_delivery_dashboard()
            staging["odsOrderCount"], staging["odsWaybillCount"] = ods_order_t, ods_waybill_t
            print(f"  ✅ ODS counts from Tableau 'Actual Delivery - 10 Districts': "
                  f"{ods_order_t['overall']} orders / {ods_waybill_t['overall']} waybills")
        except Exception as e:
            print(f"  ⚠️ Could not use the Tableau ODS report ({e}) — falling back to the OIX-derived "
                  f"ODS order/waybill counts staged at 03:00.")
        trimmed = finish_productivity_with_orders(history, matrices, staging, order_totals_t1,
                                                    order_totals_mtd, productivity_target_date)
        save_productivity_history(trimmed, trimmed_productivity_mtd(history, DAILY_PRODUCTIVITY_KEEP_DAYS))

    # --- Poor Rating (30-day rolling forecast, unchanged) ---
    poor = parse_poor_rating()
    append_history(history, "poorRating", today.isoformat(), poor["overall"], poor["districts"])
    poor_fc_overall, poor_fc_districts = rolling_average(history, "poorRating", 30, today)
    matrices["poorRating"] = {
        "actual": poor,
        "forecast": {"overall": poor_fc_overall, "districts": poor_fc_districts},
        "target": POOR_RATING_TARGET,
        "asOf": today.isoformat(),
    }

    # --- Missing & Lost Amount (prorate forecast) ---
    a = parse_report_a()
    b = parse_report_b()
    c = parse_report_c()
    missing_lost = combine_missing_lost(a, b, c)
    matrices["missingLostAmount"] = {
        "actual": missing_lost,
        "forecast": {
            "overall": prorate_forecast(missing_lost["overall"], today.day, total_days),
            "districts": {d: prorate_forecast(missing_lost["districts"][d], today.day, total_days) for d in DISTRICTS},
        },
        "target": MISSING_LOST_TARGETS,
        "asOf": today.isoformat(),
    }
    # v4.0 §4 — daily log feeding "Other Aspects Tracking"'s monthly rollup
    # (build_other_aspects_monthly()); purely additive, doesn't touch the
    # prorated forecast above.
    append_history(history, "missingLostAmount", today.isoformat(), missing_lost["overall"], missing_lost["districts"])

    # --- RFID Missing Tote (accumulates within month; T-4 record; prorate forecast) ---
    # Stored as a per-day ledger keyed by the T-4 date, NOT a running sum —
    # this makes reruns safe: rerunning on the same day overwrites that day's
    # entry instead of adding on top of it again. The bucket totals are always
    # recomputed fresh from the ledger, never incremented directly.
    update_rfid_missing_tote(payload, history, matrices, today)   # v9.4 - T-4 month attribution, see function

    # --- v3.0 §3: GMV / Basket Size — separate account/file, so handled as
    # its own soft-fail block rather than being added to the `missing` check
    # above: a GMV-account hiccup shouldn't block the other 6 reports that
    # already downloaded fine. Same T-1 cadence as OIX productivity, since
    # basket size needs that same day's Total Parent Order to divide by.
    gmv_path = os.path.join(REPORT_FOLDER, REPORT_FILES["gmv"])
    if os.path.exists(gmv_path):
        gmv_date = today - dt.timedelta(days=1)  # T-1
        try:
            gmv_group = parse_gmv(gmv_date)
            append_gmv_history(history, gmv_date.isoformat(), gmv_group)
            backfill_gmv_history(history)  # v4.0 §3 — fill any other missing days this download still has
            orders_overall, orders_districts = total_parent_orders_for(history, gmv_date.isoformat())
            matrices["gmv"] = {
                "actual": gmv_group,
                "basketSize": basket_size(gmv_group, orders_overall, orders_districts),
                "asOf": gmv_date.isoformat(),
            }
            save_gmv_history(build_gmv_monthly(history))
        except Exception as e:
            print(f"  ⚠️ GMV parse failed, skipping this run's GMV update: {e}")
    else:
        print(f"  ⚠️ {REPORT_FILES['gmv']!r} not found in {REPORT_FOLDER!r} — skipping GMV/Basket "
              f"Size update. fetch_gmv_report() should have downloaded it (needs "
              f"TABLEAU_GMV_USER/PASS set in .env).")

    # v4.0 §4 — "Other Aspects Tracking" tab (poor rating %, missing & lost
    # amount, RFID missing tote), monthly.
    save_other_aspects_history(build_other_aspects_monthly(history, matrices, payload))

    save_history(history)
    save_data_json(payload)
    update_embedded_data()  # v28.0 — refresh index.html's embedded fallback snapshot with the files just written above
    print("✅ 報表解析完成，已寫入 data.json")


# =============================================================================
# 7. v8.0 — Daily Cost Report (WhatsApp) : manpower distribution, Cost/Order, Staff Productivity
# =============================================================================
_MONTH_ABBR = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
CR_FT_COST_TYPE = "Manpower Cost (FT & Leader)"
CR_PT_COST_TYPE = "Manpower Cost (PT)"
# Title grouping (spec §2). Cost Type + Function decide the group; the "借人" (borrowed
# staff) row and every Leader / Sick Leave row match none of these and are ignored.
CR_FT_DRIVER_FUNCTIONS = {"driver", "driver at", "driver c", "b shift driver c"}
CR_GROUPS = ("ftDriver", "ftCourier", "ptDriver", "ptCourier")
# Labels used on the report's tabs -> dashboard district code.
CR_OVERVIEW_LABELS = {"ETH": "ETH", "ETK": "ETK", "ETX": "ETX", "TSM": "NT-TSM", "WTK": "WTK",
                      "WTW": "NT-TW", "WTH": "WTH", "WTX": "WTX", "NT-TM": "NT-TM", "NT-ST": "NT-ST"}
CR_BREAKDOWN_LABELS = {"ETH": "ETH", "ETK": "ETK", "ETX": "ETX", "NT-ST": "NT-ST", "NT-TM": "NT-TM",
                       "TSM": "NT-TSM", "NT-TSM": "NT-TSM", "NT-TW": "NT-TW", "WTH": "WTH", "WTK": "WTK", "WTX": "WTX"}


def cost_report_sort_key(filename):
    """(year, month, day) from 'Daily Cost Report_YYYYMM_Last Update_MMM DD.xlsx'; None if the name doesn't match."""
    m = COST_REPORT_NAME_RE.match(os.path.basename(filename).strip())
    if not m:
        return None
    ym, mon, day = m.groups()
    if mon.lower() not in _MONTH_ABBR:
        return None
    return (int(ym[:4]), _MONTH_ABBR.index(mon.lower()) + 1, int(day))


def find_latest_local_cost_report():
    """Newest matching Daily Cost Report already sitting in COST_REPORT_FOLDER, or None."""
    best = None
    for p in glob.glob(os.path.join(COST_REPORT_FOLDER, "*.xlsx")):
        k = cost_report_sort_key(p)
        if k and (best is None or k > best[0]):
            best = (k, p)
    return best[1] if best else None


# v8.1 — how many screens the chat may be scrolled up while hunting for the newest report, as a LAST
# resort only (the normal routes — the already-loaded bottom of the chat, then WhatsApp's in-chat search —
# never scroll through the history).
WA_MAX_SCROLLS = int(os.environ.get("WA_MAX_SCROLLS", "8"))

# Finds elements that carry a Daily Cost Report file name, either in their `title` attribute or in their
# text. Matching on TEXT (not just title="…") matters: WhatsApp Web's document bubbles often have no title
# attribute at all, which made the old title-only lookup find nothing and scroll forever.
# mode 'main'    -> only inside the open chat (#main)
# mode 'outside' -> only outside the chat (the in-chat search results panel)
# Each hit is tagged data-cr-idx=N so Python can address it; returns [{i, name}] in document order.
_CR_SCAN_JS = r"""
(mode) => {
  const NAME = /Daily[ _]Cost[ _]Report_\d{6}_Last[ _]Update_[A-Za-z]{3}[ _]\d{1,2}(?:\.xlsx)?/i;
  document.querySelectorAll('[data-cr-idx]').forEach(e => e.removeAttribute('data-cr-idx'));
  const main = document.querySelector('#main');
  const out = [];
  let n = 0;
  for (const el of document.querySelectorAll('span, div, a, p')) {
    const inMain = !!(main && main.contains(el));
    if ((mode === 'main') !== inMain) continue;
    if (el.closest('#pane-side')) continue;
    let m = NAME.exec(el.getAttribute('title') || '');
    if (!m) {
      const t = (el.textContent || '').trim();
      if (t.length > 300) continue;
      m = NAME.exec(t);
      if (m) {
        // keep only the innermost element that holds the name
        const child = Array.from(el.children).some(c => NAME.test(c.textContent || ''));
        if (child) m = null;
      }
    }
    if (!m) continue;
    el.setAttribute('data-cr-idx', String(n));
    out.push({ i: n, name: m[0] });
    n++;
  }
  return out;
}
"""


from contextlib import contextmanager


@contextmanager
def virtual_display():
    """v8.2 — starts a private Xvfb (virtual X server) and yields the DISPLAY string Chromium should use
    (e.g. ':99'), then shuts Xvfb down again. Yields None — meaning "use whatever display already exists" —
    when WA_HIDE_MODE is not "xvfb", on non-Linux systems, or when the
    `Xvfb` binary is missing / fails to start; in that case the browser simply opens as before."""
    import shutil
    import subprocess
    if WA_HIDE_MODE != "xvfb" or not sys.platform.startswith("linux"):
        yield None
        return
    xvfb = shutil.which("Xvfb")
    if not xvfb:
        print("  ⚠️ Xvfb not installed (sudo apt install xvfb) — WhatsApp Chromium will use the normal display instead.")
        yield None
        return
    proc, display = None, None
    for num in range(99, 120):
        # skip display numbers already taken by a running X server
        if os.path.exists(f"/tmp/.X{num}-lock") or os.path.exists(f"/tmp/.X11-unix/X{num}"):
            continue
        cand = f":{num}"
        proc = subprocess.Popen([xvfb, cand, "-screen", "0", WA_VIRTUAL_SCREEN, "-nolisten", "tcp"],
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        for _ in range(50):                       # wait up to ~5s for the X socket to appear
            if proc.poll() is not None or os.path.exists(f"/tmp/.X11-unix/X{num}"):
                break
            time.sleep(0.1)
        if proc.poll() is None and os.path.exists(f"/tmp/.X11-unix/X{num}"):
            display = cand
            break
        if proc.poll() is None:
            proc.terminate()
        proc = None
    if not display:
        print("  ⚠️ Could not start Xvfb — WhatsApp Chromium will use the normal display instead.")
        yield None
        return
    print(f"  🖥️ Virtual display {display} started ({WA_VIRTUAL_SCREEN}) — Chromium will not show a window.")
    try:
        yield display
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()


def dismiss_wa_popups(page):
    """v9.1 - WhatsApp Web rolls out one-time promo tooltips (e.g. "View recent calls, or start a new one with up to 32
    people." / "Voice and video calling is now available") that float over the search box. Playwright then refuses to
    click the box ("<div> intercepts pointer events") and the run dies. Close any such tooltip/dialog first."""
    for _ in range(4):
        page.keyboard.press("Escape")
        time.sleep(0.3)
        clicked = False
        for sel in ['[role="tooltip"] [aria-label*="Close" i]', '[role="tooltip"] [aria-label*="Dismiss" i]',
                    '[role="dialog"] [aria-label*="Close" i]', '[role="dialog"] button:has-text("OK")',
                    '[role="dialog"] button:has-text("Got it")',
                    'span[data-icon="x-alt"]', 'span[data-icon="x"]:visible']:
            try:
                loc = page.locator(sel).first
                if loc.is_visible():
                    loc.click(timeout=1500)
                    clicked = True
                    time.sleep(0.5)
                    break
            except Exception:
                continue
        if not clicked:
            break


def wa_scroll_chat_to_bottom(page, rounds=8):
    """v9.2 - a chat with unread messages does NOT open at its newest message: WhatsApp Web jumps to the first
    unread one, and only the messages around the viewport exist in the DOM. Route 1 then "found" an OLD report and
    never saw today's one further down. Jump to the real bottom (and let lazy-loaded messages render) first."""
    for sel in ['#main [aria-label*="Scroll to bottom" i]', '#main [data-icon="chevron-down-circle"]',
                '#main [data-icon="down"]', '#main button[aria-label*="bottom" i]']:
        try:
            loc = page.locator(sel).first
            if loc.is_visible():
                loc.click(timeout=1500)
                time.sleep(1.5)
                break
        except Exception:
            continue
    last = None
    for _ in range(rounds):
        try:
            h = page.evaluate("""() => {
                const row = document.querySelector('#main div[data-id]');
                let el = row;
                while (el && el !== document.body) {
                    const cs = getComputedStyle(el);
                    if ((cs.overflowY === 'auto' || cs.overflowY === 'scroll') && el.scrollHeight > el.clientHeight) break;
                    el = el.parentElement;
                }
                if (!el || el === document.body) return -1;
                el.scrollTop = el.scrollHeight;
                return el.scrollHeight;
            }""")
        except Exception:
            h = -1
        try:
            page.mouse.move(900, 500)
            page.mouse.wheel(0, 4000)
        except Exception:
            pass
        time.sleep(1.2)
        if h == last:
            break
        last = h


def fetch_cost_report_from_whatsapp():
    """v8.0 §2 / v8.1 — downloads the NEWEST 'Daily Cost Report_YYYYMM_Last Update_MMM DD.xlsx' from the
    WhatsApp group WA_TARGET_GROUP via WhatsApp Web, using the already-logged-in persistent profile in
    WA_SESSION_DIR. Returns the saved file's path (inside COST_REPORT_FOLDER).

    v8.1 — never walks the whole chat history. The newest report is looked for, in order:
      1. in the messages already loaded when the chat opens (it opens at the newest message);
      2. through WhatsApp's own in-chat search ("Daily Cost Report"), jumping straight to a result;
      3. only if both fail, scrolling up at most WA_MAX_SCROLLS screens, stopping at the first report seen.
    Matching is by file-name TEXT as well as title attribute, and several download triggers are tried.
    A failure leaves a screenshot (_debug_whatsapp_*.png) in COST_REPORT_FOLDER."""
    def shot(page, tag):
        try:
            p = os.path.join(COST_REPORT_FOLDER, f"_debug_whatsapp_{tag}.png")
            page.screenshot(path=p, full_page=True)
            print(f"  📸 screenshot: {p}")
        except Exception:
            pass

    print(f"🚀 Opening WhatsApp Web (session {WA_SESSION_DIR!r}) for group {WA_TARGET_GROUP!r}...")
    with virtual_display() as wa_display, sync_playwright() as p:
        # v8.2 — headed Chromium, but drawn on the Xvfb virtual display (when available) so no window appears
        launch_env = {**os.environ, "DISPLAY": wa_display} if wa_display else None
        wa_args = ["--disable-popup-blocking"]
        wa_kwargs = {}
        if WA_HIDE_MODE == "headless":
            # New headless is requested through a flag on a HEADED launch (headless=False): Playwright's own
            # headless=True would start the old headless shell, which WhatsApp Web refuses.
            wa_args.append("--headless=new")
            wa_kwargs["user_agent"] = WA_USER_AGENT
            print("  🕶️ WhatsApp Chromium: new-headless mode (no window).")
        elif WA_HIDE_MODE == "offscreen":
            wa_args += ["--window-position=-32000,-32000", "--window-size=1600,1000"]
            print("  🕶️ WhatsApp Chromium: window parked off-screen.")
        ctx = p.chromium.launch_persistent_context(
            WA_SESSION_DIR, headless=WA_HEADLESS, accept_downloads=True, env=launch_env,
            viewport={"width": 1600, "height": 1000}, args=wa_args, **wa_kwargs)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto("https://web.whatsapp.com", wait_until="domcontentloaded")
            try:
                page.wait_for_selector("#pane-side", timeout=120000)
            except Exception:
                shot(page, "login")
                raise RuntimeError("WhatsApp Web never showed the chat list — the session in "
                                   f"{WA_SESSION_DIR!r} has probably expired (QR code); re-link it once by hand.")
            time.sleep(3)
            dismiss_wa_popups(page)

            # --- open the group ---
            title_sel = "#pane-side span[title=" + json.dumps(WA_TARGET_GROUP, ensure_ascii=False) + "]"
            opened_group = False
            try:                                    # v9.1 - the group is normally already in the chat list: no search needed
                page.locator(title_sel).first.click(timeout=6000)
                opened_group = True
            except Exception:
                dismiss_wa_popups(page)
            if not opened_group:
                box = None
                for sel in ['div[contenteditable="true"][data-tab="3"]', '[aria-label="Search input textbox"]',
                            '[aria-label="Search or start a new chat"]', '#side div[role="textbox"]']:
                    loc = page.locator(sel).first
                    try:
                        loc.wait_for(state="visible", timeout=4000)
                        box = loc
                        break
                    except Exception:
                        continue
                if box is None:
                    shot(page, "search")
                    raise RuntimeError("Could not find WhatsApp's search box.")
                try:
                    box.click(timeout=5000)
                except Exception:
                    dismiss_wa_popups(page)
                    box.click(force=True, timeout=5000)     # v9.1 - last resort: ignore whatever floats above it
                box.fill(WA_TARGET_GROUP)
                time.sleep(2)
                try:
                    page.locator(title_sel).first.click(timeout=15000)
                except Exception:
                    try:
                        page.get_by_text(WA_TARGET_GROUP, exact=True).first.click(timeout=8000)
                    except Exception:
                        shot(page, "group")
                        raise RuntimeError(f"Could not open the WhatsApp group {WA_TARGET_GROUP!r}.")
            page.wait_for_selector("#main", timeout=30000)
            try:
                page.wait_for_selector("#main div[data-id]", timeout=20000)
            except Exception:
                pass
            time.sleep(3)

            # --- locate the newest report ---
            def scan(mode):
                out = []
                for it in page.evaluate(_CR_SCAN_JS, mode):
                    name = it["name"].strip()
                    if not name.lower().endswith(".xlsx"):
                        name += ".xlsx"
                    k = cost_report_sort_key(name)
                    if k:
                        out.append((k, it["i"], name))
                return out

            wa_scroll_chat_to_bottom(page)          # v9.2 - make sure we are looking at the NEWEST messages
            found = scan("main")
            if found:
                print(f"  🔎 route 1: {len(found)} report message(s) already loaded at the bottom of the chat")

            if not found:
                # route 2 — WhatsApp's in-chat search; jump to a hit, no scrolling through history
                print("  🔎 route 2: in-chat search for 'Daily Cost Report'")
                try:
                    opened = False
                    for sel in ['#main header [aria-label="Search"]', '#main header [data-icon="search-refreshed"]',
                                '#main header [data-icon="search"]', '#main header button[title="Search…"]']:
                        try:
                            page.locator(sel).first.click(timeout=3000)
                            opened = True
                            break
                        except Exception:
                            continue
                    if not opened:
                        page.keyboard.press("Control+Shift+F")
                    time.sleep(1)
                    page.keyboard.type("Daily Cost Report", delay=40)
                    time.sleep(4)
                    hits = scan("outside")
                    if hits:
                        best_key = max(k for k, _, _ in hits)
                        _, i0, n0 = [h for h in hits if h[0] == best_key][0]     # results list newest first
                        print(f"  🔎 search hit: {n0!r}")
                        page.locator(f'[data-cr-idx="{i0}"]').first.click(timeout=5000)
                        time.sleep(4)
                    else:
                        print("  ℹ️ the in-chat search showed no report hit")
                    page.keyboard.press("Escape")
                    time.sleep(1)
                except Exception as e:
                    print(f"  ⚠️ in-chat search route failed: {e}")
                found = scan("main")

            if not found:
                # route 3 — bounded scroll up; the first report reached going up IS the newest one
                print(f"  🔎 route 3: scrolling up (at most {WA_MAX_SCROLLS} screens)")
                for n in range(WA_MAX_SCROLLS):
                    page.mouse.move(900, 500)
                    page.mouse.wheel(0, -1500)
                    time.sleep(1.5)
                    found = scan("main")
                    if found:
                        break
            if not found:
                shot(page, "nofile")
                raise RuntimeError("No 'Daily Cost Report_YYYYMM_Last Update_MMM DD.xlsx' message found in the group "
                                   f"(tried the loaded messages, in-chat search and {WA_MAX_SCROLLS} scrolls).")

            # --- pick the newest (latest of any duplicates) and download it ---
            print("  📋 reports visible in the chat: " + ", ".join(sorted({n for _, _, n in found})))
            best_key = max(k for k, _, _ in found)
            _, idx, name = [f for f in found if f[0] == best_key][-1]
            print(f"  📎 newest cost report in the chat: {name!r}")
            target = page.locator(f'[data-cr-idx="{idx}"]').first
            target.scroll_into_view_if_needed()
            msg = target.locator("xpath=ancestor::div[@data-id][1]")
            if msg.count() == 0:
                msg = target.locator("xpath=ancestor::div[@role='row'][1]")
            ICONS = ('span[data-icon*="download"], span[data-icon="down"], '
                     'button[aria-label*="Download" i], [title*="Download" i]')

            def a_icon():
                msg.first.hover()
                msg.locator(ICONS).first.click(timeout=3000)

            def a_name():
                target.click(timeout=3000)

            def a_menu():
                msg.first.hover()
                msg.locator('span[data-icon="down-context"], span[data-icon="ic-chevron-down-menu"], '
                            '[aria-label="Menu"], [aria-label="Message options"]').first.click(timeout=3000)
                page.get_by_text(re.compile(r"^\s*Download\s*$", re.I)).first.click(timeout=3000)

            saved = None
            for label, act in (("download icon", a_icon), ("file name click", a_name), ("message menu", a_menu)):
                try:
                    with page.expect_download(timeout=25000) as dl_info:
                        act()
                    dl = dl_info.value
                    fname = dl.suggested_filename or name
                    if not cost_report_sort_key(fname):
                        fname = name
                    saved = os.path.join(COST_REPORT_FOLDER, fname)
                    dl.save_as(saved)
                    print(f"  ✅ saved via {label}: {saved}")
                    break
                except Exception as e:
                    print(f"  ⚠️ {label} did not start a download ({str(e).splitlines()[0][:120]})")
            if not saved:
                shot(page, "download")
                raise RuntimeError(f"Found {name!r} but none of the download triggers worked — see the screenshot.")
            return saved
        except Exception:
            try:
                shot(ctx.pages[0], "error")
            except Exception:
                pass
            raise
        finally:
            ctx.close()


def _cr_num(v):
    """Numeric cell -> float; '#DIV/0!', text, blanks -> None."""
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return None
    return float(v)


def _cr_overview_district(label):
    s = str(label or "").strip()
    if s.lower() == "overall":
        return "Overall"
    for k, v in CR_OVERVIEW_LABELS.items():    # 'ETH - Chun', 'TSM (Yun)', 'WTW (Shing)', ...
        if s == k or s.startswith(k + " ") or s.startswith(k + "("):
            return v
    return None


def parse_cost_report(path):
    """v8.0 §2 — reads the Daily Cost Report workbook.
    Overview tab: each section (District / Overall mark in the top-left cell of its header row) has a
    'Cost Per Order (Controllable)' and a 'Staff Productivity' row; the date columns are matched to
    the header row's dates (shown as DD-MMM, e.g. 29-Sep). Blank / #DIV/0! / 0 = not reported yet.
    Breakdown tab, FIRST table only (not the yellow one): per district block and per day column (DD),
    the FT/PT Driver & Courier man-days, grouped by Cost Type + Function; row '借人' ignored.
    Only days up to the report's 'Last Update' date (Breakdown!B3) are used.
    Returns {"file","asOf","manpower":{date:{group:{"districts","total"}}},
             "daily":{date:{"costPerOrder":{overall,districts},"productivity":{...}}},
             "mtd":{"costPerOrder":{...},"productivity":{...}},
             "totalCost":{overall,districts}}   # v9.0 — Overview 'Total Cost' (MTD, col C) per section."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    names = {n.strip().lower(): n for n in wb.sheetnames}
    for need in ("overview", "breakdown"):
        if need not in names:
            raise ValueError(f"{os.path.basename(path)!r} has no '{need.title()}' sheet (found {wb.sheetnames})")
    ws_o, ws_b = wb[names["overview"]], wb[names["breakdown"]]

    upd = ws_b["B3"].value
    if isinstance(upd, dt.datetime):
        upd = upd.date()
    if not isinstance(upd, dt.date):
        raise ValueError("Breakdown!B3 (the report's 'Date' / last-update day) is not a date")

    # ---- Overview: Cost Per Order (Controllable) + Staff Productivity ----
    row_metric = {"Cost Per Order (Controllable)": "costPerOrder", "Staff Productivity": "productivity"}
    blank = lambda: {"overall": None, "districts": {}}
    daily, mtd = {}, {"costPerOrder": blank(), "productivity": blank()}
    total_cost = blank()      # v9.0 — Overview 'Total Cost' row, column C (MTD), Overall + per district
    r, max_r = 1, ws_o.max_row
    while r <= max_r:
        if str(ws_o.cell(r, 3).value or "").strip() != "MTD":
            r += 1
            continue
        key = _cr_overview_district(ws_o.cell(r, 1).value)
        date_cols = {}
        for c in range(4, ws_o.max_column + 1):
            v = ws_o.cell(r, c).value
            if isinstance(v, dt.datetime):
                v = v.date()
            if isinstance(v, dt.date) and v <= upd:
                date_cols[c] = v.isoformat()
        rr = r + 1
        while rr <= max_r and str(ws_o.cell(rr, 3).value or "").strip() != "MTD":
            if key and str(ws_o.cell(rr, 1).value or "").strip() == "Total Cost":     # exact: not '...(Controllable Cost)'
                tc = _cr_num(ws_o.cell(rr, 3).value)
                if tc is not None and tc > 0:
                    if key == "Overall":
                        total_cost["overall"] = tc
                    else:
                        total_cost["districts"][key] = tc
            metric = row_metric.get(str(ws_o.cell(rr, 1).value or "").strip())
            if metric and key:
                m = _cr_num(ws_o.cell(rr, 3).value)
                if key == "Overall":
                    mtd[metric]["overall"] = m
                else:
                    mtd[metric]["districts"][key] = m
                for c, d in date_cols.items():
                    v = _cr_num(ws_o.cell(rr, c).value)
                    if v is not None and v <= 0:
                        v = None                       # 0 = day not filled in yet
                    slot = daily.setdefault(d, {"costPerOrder": blank(), "productivity": blank()})[metric]
                    if key == "Overall":
                        slot["overall"] = v
                    else:
                        slot["districts"][key] = v
            rr += 1
        r = rr
    daily = {d: v for d, v in daily.items()
             if any(v[m]["overall"] is not None or any(x is not None for x in v[m]["districts"].values())
                    for m in ("costPerOrder", "productivity"))}

    # ---- Breakdown: first table, manpower man-days per group / district / day ----
    hdr = next((r for r in range(1, ws_b.max_row + 1) if str(ws_b.cell(r, 1).value or "").strip() == "Cost Type"), None)
    if hdr is None:
        raise ValueError("Breakdown tab: could not find the 'Cost Type' header row of the first table")
    end = next((r for r in range(hdr + 1, ws_b.max_row + 1)
                if str(ws_b.cell(r, 5).value or "").strip().startswith("Total")), ws_b.max_row + 1)
    starts = [(c, CR_BREAKDOWN_LABELS[str(ws_b.cell(hdr - 1, c).value).strip()])
              for c in range(1, ws_b.max_column + 1)
              if str(ws_b.cell(hdr - 1, c).value or "").strip() in CR_BREAKDOWN_LABELS]
    blocks = {}
    for i, (c0, dist) in enumerate(starts):
        c1 = starts[i + 1][0] if i + 1 < len(starts) else ws_b.max_column + 1
        blocks[dist] = {int(ws_b.cell(hdr, c).value): c for c in range(c0, c1)
                        if isinstance(ws_b.cell(hdr, c).value, (int, float)) and not isinstance(ws_b.cell(hdr, c).value, bool)
                        and 1 <= int(ws_b.cell(hdr, c).value) <= 31}
    missing = [d for d in DISTRICTS if d not in blocks]
    if missing:
        raise ValueError(f"Breakdown tab: district block(s) not found: {missing}")

    cost_type, row_group = None, {}
    for r in range(hdr + 1, end):
        a = str(ws_b.cell(r, 1).value or "").strip()
        if a:
            cost_type = a                                   # Cost Type is only written on a block's first row
        func = str(ws_b.cell(r, 2).value or "").strip().lower()
        if cost_type == CR_FT_COST_TYPE:
            if func in CR_FT_DRIVER_FUNCTIONS:
                row_group[r] = "ftDriver"
            elif func == "courier":
                row_group[r] = "ftCourier"
        elif cost_type == CR_PT_COST_TYPE:
            if func == "driver":
                row_group[r] = "ptDriver"
            elif func == "courier":
                row_group[r] = "ptCourier"

    manpower = {}
    for day in range(1, upd.day + 1):
        per = {g: {"districts": {d: 0.0 for d in DISTRICTS}, "total": 0} for g in CR_GROUPS}
        seen = False
        for dist in DISTRICTS:
            c = blocks[dist].get(day)
            if c is None:
                continue
            for r, g in row_group.items():
                v = _cr_num(ws_b.cell(r, c).value)
                if v is not None:
                    seen = True
                    per[g]["districts"][dist] += v
        if not seen:
            continue
        for g in CR_GROUPS:
            per[g]["districts"] = {d: int(round(v)) for d, v in per[g]["districts"].items()}
            per[g]["total"] = sum(per[g]["districts"].values())
        if sum(per[g]["total"] for g in CR_GROUPS) == 0:
            continue                                        # day exists in the sheet but nothing filled in yet
        manpower[dt.date(upd.year, upd.month, day).isoformat()] = per

    return {"file": os.path.basename(path), "asOf": upd.isoformat(),
            "manpower": manpower, "daily": daily, "mtd": mtd, "totalCost": total_cost}


def reapply_cost_report_fields(history):
    """v8.0 §2 — (re)writes the report's real Cost/Order and Staff Productivity into
    dailyProductivityLog's hktvStaff entries (and the 7-day rolling-forecast series). Everything
    comes from history['costReportDaily'] (the durable copy), so it can be called again after any
    other job rebuilds a day's entry from OIX/Tableau. ODS/VAN is untouched (spec: no effect)."""
    stored = history.get("costReportDaily", {})
    plog = history.get("dailyProductivityLog", {})
    series = history.setdefault("hktvStaff", {})
    for date_str, d in stored.items():
        prod, cost = d["productivity"], d["costPerOrder"]
        entry = plog.get(date_str, {}).get("hktvStaff")
        if entry:
            for dist in DISTRICTS:
                rec = entry.get("districts", {}).get(dist)
                if rec is None:
                    continue
                if prod["districts"].get(dist) is not None:
                    rec["productivity"] = round(prod["districts"][dist], 2)
                    rec["productivitySource"] = "costReport"
                if cost["districts"].get(dist) is not None:
                    rec["costPerOrder"] = round(cost["districts"][dist], 2)
            tot = entry.get("total")
            if tot is not None:
                if prod["overall"] is not None:
                    tot["productivity"] = round(prod["overall"], 2)
                    tot["productivitySource"] = "costReport"
                if cost["overall"] is not None:
                    tot["costPerOrder"] = round(cost["overall"], 2)
        cur = series.get(date_str, {"overall": None, "districts": {}})
        series[date_str] = {
            "overall": round(prod["overall"], 2) if prod["overall"] is not None else cur.get("overall"),
            "districts": {dist: (round(prod["districts"][dist], 2) if prod["districts"].get(dist) is not None
                                 else cur.get("districts", {}).get(dist)) for dist in DISTRICTS},
        }


def apply_cost_report(history, report):
    """v8.0 §2 — covers the old data with the fetched data:
      - manpowerDistributionLog[date]: courier / driver (= Full-Time, same meaning as before) are
        overwritten and courierPT / driverPT added; the ODS/VAN group is kept; entry tagged
        _source='costReport' so the 03:00 OIX job won't cover it again.
      - Cost/Order + Staff Productivity: stored durably in history['costReportDaily'] and re-applied."""
    mlog = history.setdefault("manpowerDistributionLog", {})
    for date_str, g in report["manpower"].items():
        entry = dict(mlog.get(date_str, {}))
        entry["courier"], entry["driver"] = g["ftCourier"], g["ftDriver"]
        entry["courierPT"], entry["driverPT"] = g["ptCourier"], g["ptDriver"]
        entry["_source"] = "costReport"
        mlog[date_str] = entry
    stored = history.setdefault("costReportDaily", {})
    for date_str, d in report["daily"].items():
        stored[date_str] = d
    reapply_cost_report_fields(history)
    plog = history.get("dailyProductivityLog", {})
    no_log = sorted(d for d in report["daily"] if "hktvStaff" not in plog.get(d, {}))
    return {"manpowerDays": len(report["manpower"]), "costDays": len(report["daily"]), "noProductivityLogDays": no_log}


def compute_fulfillment_cost(history, report):
    """v9.0 — Fulfillment Cost % for the report's month, MTD through the report's 'Last Update' day:
         Fulfillment Cost % = Total Cost / GMV x 100
       - Total Cost: Daily Cost Report > Overview tab > 'Total Cost' row, MTD column (col C), for Overall
         and for each district (parse_cost_report()['totalCost']).
       - GMV: the SUM of history['gmv'] daily figures from the 1st of that month through the report's
         'Last Update' date (inclusive) - the same period the cost covers - overall and per district.
       Stored durably in history['fulfillmentCostMonthly'][YYYY-MM]; a newer report for the same month
       replaces the older one (an older file never overwrites a newer one). Returns the stored record,
       or None when there is nothing to compute."""
    tc = report.get("totalCost") or {}
    if tc.get("overall") is None and not tc.get("districts"):
        print("  ⚠️ Fulfillment Cost %: no 'Total Cost' row found in the Overview tab - skipped.")
        return None
    as_of = dt.date.fromisoformat(report["asOf"])
    month = as_of.strftime("%Y-%m")
    first = as_of.replace(day=1)
    gmv_log = history.get("gmv", {})
    days = sorted(d for d in gmv_log if first.isoformat() <= d <= as_of.isoformat())
    gmv_overall = sum((gmv_log[d].get("overall") or 0) for d in days)
    gmv_districts = {x: sum(((gmv_log[d].get("districts") or {}).get(x) or 0) for d in days) for x in DISTRICTS}
    if len(days) < as_of.day:
        print(f"  ⚠️ Fulfillment Cost %: GMV history covers {len(days)} of {as_of.day} day(s) "
              f"({first.isoformat()}..{as_of.isoformat()}) - the ratio uses only those days' GMV.")
    ratio = lambda cost, gmv: round(cost / gmv * 100, 2) if cost is not None and gmv else None
    rec = {
        "asOf": as_of.isoformat(), "source": report["file"],
        "gmvDays": len(days), "expectedDays": as_of.day,
        "totalCost": {"overall": tc.get("overall"), "districts": {x: tc.get("districts", {}).get(x) for x in DISTRICTS}},
        "gmv": {"overall": gmv_overall, "districts": gmv_districts},
        "overall": ratio(tc.get("overall"), gmv_overall),
        "districts": {x: ratio(tc.get("districts", {}).get(x), gmv_districts[x]) for x in DISTRICTS},
    }
    dist_sum = sum(v for v in tc.get("districts", {}).values() if v)
    if tc.get("overall") and dist_sum and abs(dist_sum - tc["overall"]) / tc["overall"] > 0.005:
        print(f"  ⚠️ Fulfillment Cost %: district Total Costs sum to {dist_sum:,.0f} but Overall says {tc['overall']:,.0f}.")
    store = history.setdefault("fulfillmentCostMonthly", {})
    old = store.get(month)
    if old and old.get("asOf", "") > rec["asOf"]:
        print(f"  ℹ️ Fulfillment Cost %: {month} already holds a newer report (as of {old['asOf']}) - kept.")
        return old
    store[month] = rec
    return rec


def build_fulfillment_cost_block(history):
    """v9.0 - the 'fulfillmentCost' block of other_aspects_history.json, built from
    history['fulfillmentCostMonthly']:
      - 'current': the latest month while its report does not yet reach the month's last day - the dashboard
        shows it as 'YYYY-MM (MTD)';
      - 'monthly': every other month (its report reaches month end, or a later month has a report).
    Each row: {asOf, overall, districts} with the % in percent units (6.21 = 6.21%)."""
    store = history.get("fulfillmentCostMonthly", {})
    block = {"current": None, "monthly": {}}
    if not store:
        return block
    latest = max(store)
    for month in sorted(store):
        rec = store[month]
        row = {"asOf": rec.get("asOf"), "overall": rec.get("overall"), "districts": rec.get("districts", {})}
        y, m = int(month[:4]), int(month[5:7])
        month_end = (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).isoformat()
        if month == latest and (rec.get("asOf") or "") < month_end:
            block["current"] = {"monthKey": month, **row}
        else:
            block["monthly"][month] = row
    return block


def write_fulfillment_cost_to_other_aspects(history):
    """v9.0 - merges the 'fulfillmentCost' block into other_aspects_history.json WITHOUT touching the other
    metrics (that file is rebuilt in full by the 14:00 Tableau job, which now includes this block too)."""
    payload = {}
    if os.path.exists(OTHER_ASPECTS_HISTORY_PATH):
        try:
            with open(OTHER_ASPECTS_HISTORY_PATH, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (OSError, ValueError):
            payload = {}
    payload["fulfillmentCost"] = build_fulfillment_cost_block(history)
    save_other_aspects_history(payload)


def run_deploy_hook():
    """v9.3 - push the refreshed dashboard (index.html + the JSON files) to wherever it is served from, right after
    the 09:00 cost-report run. The command is yours, via DEPLOY_CMD (shell string, run from the project folder), e.g.
        DEPLOY_CMD=git add -A && git commit -m "cost report" && git push
        DEPLOY_CMD=firebase deploy --only hosting
        DEPLOY_CMD=robocopy public \\\\server\\share\\dashboard /E
    Never raises: a failed deploy is printed loudly but the data files are already written."""
    import subprocess
    cmd = os.environ.get("DEPLOY_CMD", "").strip()
    if not cmd:
        print("  ℹ️ DEPLOY_CMD is not set - dashboard files were refreshed locally but NOT pushed anywhere.")
        return False
    print(f"  🚚 deploying dashboard: {cmd}")
    try:
        r = subprocess.run(cmd, shell=True, cwd=os.environ.get("DEPLOY_CWD") or None,
                           capture_output=True, text=True, timeout=int(os.environ.get("DEPLOY_TIMEOUT", "300")))
        out = ((r.stdout or "") + (r.stderr or "")).strip()
        if out:
            print("     " + out.replace("\n", "\n     ")[-1500:])
        if r.returncode == 0:
            print("  ✅ dashboard pushed.")
            return True
        print(f"  ❌ deploy command exited with code {r.returncode} - dashboard NOT updated online.")
    except Exception as e:
        print(f"  ❌ deploy command failed: {e}")
    return False


def run_cost_report_section(local_file=None):
    """v8.0 §2 — the 09:00 Tuesday/Friday job: download (or read `local_file`), parse, cover the old
    manpower / Cost-per-Order / Productivity data, rewrite the served JSON files and index.html's snapshot."""
    print("🚀 開始處理 Daily Cost Report...")
    path = local_file
    if not path:
        try:
            path = fetch_cost_report_from_whatsapp()
        except Exception as e:
            path = find_latest_local_cost_report()
            if not path:
                raise
            print(f"  ⚠️ WhatsApp download failed ({e}) — re-using the newest report already on disk: {os.path.basename(path)!r}")
    report = parse_cost_report(path)
    print(f"  📄 {report['file']}: data through {report['asOf']} — {len(report['manpower'])} day(s) of manpower, "
          f"{len(report['daily'])} day(s) of Cost/Order & Productivity")

    history = load_history()
    summary = apply_cost_report(history, report)
    fc = compute_fulfillment_cost(history, report)        # v9.0 - Fulfillment Cost % (Total Cost / GMV)
    if fc:
        print(f"  💲 Fulfillment Cost % {report['asOf'][:7]} (through {fc['asOf']}): overall {fc['overall']}% "
              f"= {fc['totalCost']['overall']:,.0f} / {fc['gmv']['overall']:,.0f}")
    if summary["noProductivityLogDays"]:
        print(f"  ℹ️ no HKTV productivity-log entry yet for {summary['noProductivityLogDays']} — their Cost/Order "
              f"& Productivity are stored and will be applied once the daily job logs those days.")

    payload = load_data_json()
    last_day = max(report["daily"]) if report["daily"] else report["asOf"]
    payload["costReport"] = {
        "source": report["file"], "asOf": report["asOf"], "latestDay": last_day,
        "latest": report["daily"].get(last_day),
        "mtd": report["mtd"],     # the report's own MTD column, kept for reference (Overview 'actual' is unchanged)
        "totalCost": report.get("totalCost"),      # v9.0 - Overview 'Total Cost' (MTD), overall + per district
    }
    matrices = payload.setdefault("matrices", {})
    if "hktvStaff" in matrices:
        fc_overall, fc_districts = rolling_average(history, "hktvStaff", 7, dt.date.today())
        matrices["hktvStaff"]["forecast"] = {"overall": fc_overall, "districts": fc_districts}

    save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
    save_productivity_history(trimmed_daily_productivity(history, DAILY_PRODUCTIVITY_KEEP_DAYS),
                              trimmed_productivity_mtd(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
    save_history(history)
    save_data_json(payload)
    if fc:
        write_fulfillment_cost_to_other_aspects(history)  # v9.0 - before update_embedded_data() so the snapshot has it
    update_embedded_data()
    run_deploy_hook()                                 # v9.3 - push the refreshed dashboard in the same 09:00 run
    print("✅ Daily Cost Report 處理完成")


# =============================================================================
# 7. Delivery Map data (Delivery Map v2) — Excel -> window.__MAP_DATA__ in index.html
# =============================================================================
# Run:  python kpi_pipeline.py --section map [--map-excel "<estate list>.xlsx"]
# (also runs, soft-failing, as the last step of --section all)
#
#   * CARTO key  — read from CARTO_KEY_FILE ("Basemaps API key: <key>"), embedded as
#                  window.__MAP_CFG__.cartoKey; the page builds
#                  https://basemaps.cartocdn.com/rastertiles/<style>/{z}/{x}/{y}{r}.png?key=<key>
#                  and falls back to OpenStreetMap when no key is found.
#   * Zone view  — Delivery Zone code = Column C of the Excel (header name is used only
#                  when Column C does not look like zone codes).
#   * Plotting   — a row is plotted exactly at its Latitude/Longitude whenever that point
#                  lies inside Hong Kong (HK_POLYGON). Only rows with no usable coordinates,
#                  or coordinates outside Hong Kong (e.g. Shenzhen), are left without
#                  lat/lng; the page places those approximately inside their area and
#                  lists them under "Data Check". There is NO distance-from-area rejection.
CARTO_KEY_FILE = os.environ.get(
    "CARTO_KEY_FILE",
    r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\Carto Map API Key.txt")
MAP_EXCEL_PATH = os.environ.get("MAP_EXCEL_PATH", "")      # exact file; blank = auto-detect in MAP_EXCEL_FOLDER
MAP_EXCEL_FOLDER = os.environ.get(
    "MAP_EXCEL_FOLDER", r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker")
MAP_ZONE_COLUMN_INDEX = 2    # Column C (0-based)

# Coarse Hong Kong outline (lat, lng) incl. surrounding waters; the northern edge follows the
# Shenzhen River / Sha Tau Kok border. Used only to separate Hong Kong from Shenzhen.
HK_POLYGON = [(22.10, 113.78), (22.42, 113.78), (22.50, 113.84), (22.498, 113.93), (22.4985, 113.9435),
              (22.5025, 114.00), (22.515, 114.068), (22.533, 114.113), (22.55, 114.125), (22.566, 114.15),
              (22.572, 114.19), (22.562, 114.23), (22.56, 114.28), (22.57, 114.35), (22.57, 114.50),
              (22.10, 114.50)]

_MAP_HEADERS = {   # field -> accepted header names (lower-case, spaces/underscores ignored)
    "code": ["estatecode", "estateid", "code"],
    "en":   ["estatenameen", "estatenameenglish", "englishname", "nameen", "estatename", "estate", "name"],
    "zh":   ["estatenamezh", "estatenamechinese", "chinesename", "namezh", "estatenamecn", "中文名稱", "中文名", "中文"],
    "dd":   ["deliverydistrict", "dd", "district", "deliverydist"],
    "type": ["estatetype", "deliverytype", "type", "typecode"],
    "zone": ["deliveryzone", "deliveryzonecode", "zonecode", "zone"],
    "area": ["area", "estatearea", "areaname", "region", "areaen"],
    "lat":  ["latitude", "lat"],
    "lng":  ["longitude", "lng", "lon", "long"],
}
_ZONE_RE = re.compile(r"^[A-Za-z]{1,4}\d+(-\d+)?$")


def load_carto_key():
    """Basemaps API key from CARTO_KEY_FILE ('Basemaps API key: <key>'); '' if missing."""
    try:
        txt = Path(CARTO_KEY_FILE).read_text(encoding="utf-8-sig", errors="ignore")
    except OSError as e:
        print(f"  ⚠️ CARTO key file not readable ({CARTO_KEY_FILE!r}: {e}) — map will use OpenStreetMap tiles.")
        return ""
    m = re.search(r"api\s*key\s*:\s*([^\s]+)", txt, re.IGNORECASE)
    if not m:   # tolerate a file that holds just the bare key on its own line
        toks = [t for t in re.findall(r"[A-Za-z0-9_\-\.]{16,}", txt)]
        key = toks[-1] if toks else ""
    else:
        key = m.group(1).strip().strip("\"'")
    if not key:
        print("  ⚠️ No key found after 'Basemaps API key:' in the CARTO key file — map will use OpenStreetMap tiles.")
    return key


def point_in_hk(lat, lng):
    inside, n = False, len(HK_POLYGON)
    j = n - 1
    for i in range(n):
        yi, xi = HK_POLYGON[i]
        yj, xj = HK_POLYGON[j]
        if (yi > lat) != (yj > lat) and lng < (xj - xi) * (lat - yi) / (yj - yi) + xi:
            inside = not inside
        j = i
    return inside


def _clean_coord(lat, lng):
    """-> (lat, lng) floats, or None. Accepts '22.3,' / '22.3°N' strings; fixes swapped lat/lng."""
    def f(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        m = re.search(r"-?\d+(?:\.\d+)?", str(v).replace(",", ""))
        return float(m.group()) if m else None
    la, ln = f(lat), f(lng)
    if la is None or ln is None:
        return None
    if 113 <= la <= 115 and 22 <= ln <= 23:       # swapped columns
        la, ln = ln, la
    if la == 0 or ln == 0:
        return None
    return la, ln


def _norm_header(h):
    return re.sub(r"[\s_\-\(\)（）\.]+", "", str(h)).lower()


def _find_map_excel(explicit=None):
    if explicit:
        return explicit
    if MAP_EXCEL_PATH:
        return MAP_EXCEL_PATH
    cands = []
    for ext in ("*.xlsx", "*.xlsm", "*.csv"):
        cands += glob.glob(os.path.join(MAP_EXCEL_FOLDER, ext))
    cands = [p for p in cands if not os.path.basename(p).lower().startswith(("~$", "daily cost report", "oix_record", "new_estate_tracker"))]
    for p in sorted(cands, key=os.path.getmtime, reverse=True):   # newest first, first one with lat+lng headers
        try:
            head = pd.read_csv(p, nrows=8, header=None, dtype=str, encoding="utf-8-sig") if p.lower().endswith(".csv") \
                else pd.read_excel(p, nrows=8, header=None, dtype=str)
        except Exception:
            continue
        for _, row in head.iterrows():
            names = {_norm_header(x) for x in row.dropna()}
            if names & set(_MAP_HEADERS["lat"]) and names & set(_MAP_HEADERS["lng"]):
                return p
    raise FileNotFoundError(
        f"No estate/address list with Latitude + Longitude headers found in {MAP_EXCEL_FOLDER!r}. "
        f"Pass --map-excel <file> or set MAP_EXCEL_PATH.")


def load_map_excel(path):
    """-> DataFrame with canonical columns code,en,zh,dd,type,zone,area,lat,lng (header row auto-detected)."""
    raw = pd.read_csv(path, header=None, dtype=str, encoding="utf-8-sig") if path.lower().endswith(".csv") \
        else pd.read_excel(path, header=None, dtype=str)
    hdr_row = None
    for i in range(min(10, len(raw))):
        names = {_norm_header(x) for x in raw.iloc[i].dropna()}
        if names & set(_MAP_HEADERS["lat"]) and names & set(_MAP_HEADERS["lng"]):
            hdr_row = i
            break
    if hdr_row is None:
        raise ValueError(f"{path}: could not find a header row with Latitude/Longitude in the first 10 rows.")
    header = [_norm_header(x) if pd.notna(x) else "" for x in raw.iloc[hdr_row]]
    body = raw.iloc[hdr_row + 1:].reset_index(drop=True)
    pick = {}
    for field, names in _MAP_HEADERS.items():
        for nm in names:            # first alias in priority order that exists
            if nm in header:
                pick[field] = header.index(nm)
                break
    # Delivery Zone = Column C when it looks like zone codes (per spec), else the header match
    if body.shape[1] > MAP_ZONE_COLUMN_INDEX:
        colc = body.iloc[:, MAP_ZONE_COLUMN_INDEX].dropna().astype(str).str.strip()
        if len(colc) and colc.map(lambda v: bool(_ZONE_RE.match(v))).mean() >= 0.5:
            pick["zone"] = MAP_ZONE_COLUMN_INDEX
    missing = [k for k in ("code", "zone", "lat", "lng") if k not in pick]
    if missing:
        raise ValueError(f"{path}: cannot locate column(s) {missing}. Headers seen: {header}")
    print("  Map Excel columns -> " + ", ".join(f"{k}=col {chr(65 + v) if v < 26 else v}" for k, v in pick.items()))
    out = pd.DataFrame({k: (body.iloc[:, v].fillna("").astype(str).str.strip() if k not in ("lat", "lng")
                            else body.iloc[:, v]) for k, v in pick.items()})
    for k in ("en", "zh", "dd", "type", "area"):
        if k not in out:
            out[k] = ""
    out = out[out["code"] != ""].drop_duplicates(subset="code", keep="last").reset_index(drop=True)
    return out


MAP_LOOKUP_PATH = os.environ.get("MAP_LOOKUP_PATH", "./map_lookup.json")


def _existing_map_lookups():
    """Lookups used to fill columns the Excel lacks (district / type / area / names):
    1) MAP_LOOKUP_PATH (map_lookup.json, kept next to this script and refreshed after every good build),
    2) otherwise the map block already embedded in index.html.
    -> (by_code {code: {en,zh,dd,type,area}}, zone_dd {zone: district}, centres [(area, lat, lng)])"""
    from collections import Counter
    if os.path.exists(MAP_LOOKUP_PATH):
        try:
            L = json.loads(Path(MAP_LOOKUP_PATH).read_text(encoding="utf-8"))
            by_code = {c: dict(zip(("en", "zh", "dd", "type", "area"), v)) for c, v in L["codes"].items()}
            return by_code, L["zoneDd"], [tuple(c) for c in L["centres"]]
        except Exception as e:
            print(f"  ⚠️ {MAP_LOOKUP_PATH} unreadable ({e}) — trying the map block in index.html.")
    try:
        html = Path(INDEX_HTML_PATH).read_text(encoding="utf-8")
        m = re.search(r"window\.__MAP_DATA__=(\{.*?\});?/\*MAP_DATA_END\*/", html, re.DOTALL)
        old = json.loads(m.group(1))
    except Exception:
        return {}, {}, []
    by_code, zc, centres = {}, {}, []
    for ar in old.get("areas", []):
        pts = [(r[6], r[7]) for r in ar["e"] if len(r) >= 8]
        if pts:
            centres.append((ar["k"], sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts)))
        for r in ar["e"]:
            dd, ty, zn = old["dd"][r[3]], old["types"][r[4]], old["zones"][r[5]]
            by_code[r[0]] = {"en": r[1], "zh": r[2], "dd": dd, "type": ty, "area": ar["k"]}
            zc.setdefault(zn, Counter())[dd] += 1
    return by_code, {z: c.most_common(1)[0][0] for z, c in zc.items()}, centres


def save_map_lookup(data):
    """Refreshes MAP_LOOKUP_PATH from a freshly built map dict — only when it carries real districts."""
    if len([d for d in data["dd"] if d != "(none)"]) < 2:
        return
    from collections import Counter
    codes, zc, cen = {}, {}, []
    for ar in data["areas"]:
        pts = [(r[6], r[7]) for r in ar["e"] if len(r) >= 8]
        if pts:
            cen.append([ar["k"], sum(p[0] for p in pts) / len(pts), sum(p[1] for p in pts) / len(pts)])
        for r in ar["e"]:
            codes[r[0]] = [r[1], r[2], data["dd"][r[3]], data["types"][r[4]], ar["k"]]
            zc.setdefault(data["zones"][r[5]], Counter())[data["dd"][r[3]]] += 1
    Path(MAP_LOOKUP_PATH).write_text(json.dumps(
        {"_note": "refreshed by kpi_pipeline.py --section map", "codes": codes,
         "zoneDd": {z: c.most_common(1)[0][0] for z, c in zc.items()}, "centres": cen},
        ensure_ascii=False, separators=(",", ":")), encoding="utf-8")


def enrich_map_frame(df):
    """Fills blank district / type / area / names: 1) from the previous map data by estate code,
    2) district from the delivery zone, 3) area from the nearest known area centre."""
    by_code, zone_dd, centres = _existing_map_lookups()
    if not by_code:
        return df
    before = {k: int((df[k] == "").sum()) for k in ("dd", "type", "area")}
    def fill(row):
        o = by_code.get(row["code"], {})
        for k in ("en", "zh", "dd", "type", "area"):
            if not row[k] and o.get(k):
                row[k] = o[k]
        if not row["dd"] and row["zone"] in zone_dd:
            row["dd"] = zone_dd[row["zone"]]
        if not row["area"] and centres:
            c = _clean_coord(row["lat"], row["lng"])
            if c and point_in_hk(*c):
                row["area"] = min(centres, key=lambda a: (a[1] - c[0]) ** 2 + (a[2] - c[1]) ** 2)[0]
        return row
    df = df.apply(fill, axis=1)
    after = {k: int((df[k] == "").sum()) for k in ("dd", "type", "area")}
    print("  Filled from previous map data / zone: " + ", ".join(f"{k} {before[k] - after[k]:,} (still blank {after[k]:,})" for k in before))
    return df


def build_map_data(df):
    """Excel frame -> the dict the Delivery Map tab embeds (same shape as before + hkPoly)."""
    import statistics
    dd_list = sorted({v for v in df["dd"] if v}) or ["(none)"]
    type_list = sorted({v for v in df["type"] if v}) or ["-"]
    zone_list = sorted({v for v in df["zone"] if v} | ({"(none)"} if (df["zone"] == "").any() else set()))
    area_rows = {}
    n_ok = n_missing = n_outside = 0
    problems = []
    for r in df.itertuples(index=False):
        c = _clean_coord(r.lat, r.lng)
        rec = [r.code, r.en or r.code, r.zh or r.en or r.code,
               dd_list.index(r.dd) if r.dd in dd_list else 0,
               type_list.index(r.type) if r.type in type_list else 0,
               zone_list.index(r.zone or "(none)")]
        if c is None:
            n_missing += 1
            problems.append((r.code, "missing/invalid lat-long", r.lat, r.lng))
        elif not point_in_hk(*c):
            n_outside += 1
            problems.append((r.code, "outside Hong Kong", c[0], c[1]))
        else:
            n_ok += 1
            rec += [round(c[0], 6), round(c[1], 6)]
        area_rows.setdefault(r.area or "(no area)", []).append(rec)
    areas = []
    for key in sorted(area_rows):
        rows = area_rows[key]
        pts = [(x[6], x[7]) for x in rows if len(x) >= 8]
        if pts:
            clat, clng = statistics.median(p[0] for p in pts), statistics.median(p[1] for p in pts)
            d = sorted(((p[0] - clat) * 111000) ** 2 + ((p[1] - clng) * 111000 * 0.92) ** 2 for p in pts)
            rad = int(min(3000, max(300, 1.2 * d[int(0.9 * (len(d) - 1))] ** 0.5)))
        else:
            clat, clng, rad = 22.355, 114.15, 1500
        areas.append({"k": key, "lat": round(clat, 4), "lng": round(clng, 4), "r": rad, "e": rows})
    today = today_hkt()
    data = {"asOf": f"{today:%b} {today.day}", "dd": dd_list, "types": type_list, "zones": zone_list,
            "areas": areas, "hkPoly": [list(p) for p in HK_POLYGON]}
    return data, {"total": len(df), "plotted": n_ok, "missing": n_missing, "outsideHK": n_outside,
                  "zones": len(zone_list), "problems": problems}


def inject_map_data(data, carto_key):
    """Rewrites only the /*MAP_DATA_START*/ … /*MAP_DATA_END*/ block of index.html."""
    if not os.path.exists(INDEX_HTML_PATH):
        print(f"  ⚠️ {INDEX_HTML_PATH!r} not found — map data not embedded.")
        return False
    html = Path(INDEX_HTML_PATH).read_text(encoding="utf-8")
    cfg = json.dumps({"cartoKey": carto_key})
    blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    block = f"/*MAP_DATA_START*/window.__MAP_CFG__={cfg};window.__MAP_DATA__={blob};/*MAP_DATA_END*/"
    new, n = re.subn(r"/\*MAP_DATA_START\*/.*?/\*MAP_DATA_END\*/", lambda m: block, html, count=1, flags=re.DOTALL)
    if n != 1:
        print("  ⚠️ MAP_DATA_START/END markers not found in index.html — map data not embedded.")
        return False
    Path(INDEX_HTML_PATH).write_text(new, encoding="utf-8")
    return True


def _inject_map_key_only(carto_key):
    """Updates only window.__MAP_CFG__ in index.html (map data untouched)."""
    if not os.path.exists(INDEX_HTML_PATH):
        print(f"  ⚠️ {INDEX_HTML_PATH!r} not found.")
        return False
    html = Path(INDEX_HTML_PATH).read_text(encoding="utf-8")
    cfg = "window.__MAP_CFG__=" + json.dumps({"cartoKey": carto_key}) + ";"
    new, n = re.subn(r"window\.__MAP_CFG__=\{[^}]*\};", lambda m: cfg, html, count=1)
    if n != 1:
        print("  ⚠️ window.__MAP_CFG__ not found in index.html — use the Delivery Map v2 index.html first.")
        return False
    Path(INDEX_HTML_PATH).write_text(new, encoding="utf-8")
    return True


def run_map_section(excel_path=None):
    print("🗺️ 開始更新 Delivery Map 數據...")
    key = load_carto_key()
    try:
        path = _find_map_excel(excel_path)
    except FileNotFoundError as e:
        # No Excel this run: still push the CARTO key into the existing map block so the
        # Carto base maps come back; the estate/zone data already in index.html is kept.
        print(f"  ⚠️ {e}")
        if _inject_map_key_only(key):
            print(f"  ✅ index.html: CARTO key {'embedded' if key else 'NOT found — OSM fallback'} (map data unchanged)")
            run_deploy_hook()
        return
    print(f"  Source: {path}")
    data, st = build_map_data(enrich_map_frame(load_map_excel(path)))
    print(f"  Estates {st['total']:,} · plotted {st['plotted']:,} · missing coords {st['missing']} · "
          f"outside Hong Kong {st['outsideHK']} · delivery zones {st['zones']:,}")
    for code, why, la, ln in st["problems"][:25]:
        print(f"    - {code}: {why} ({la}, {ln})")
    if len(st["problems"]) > 25:
        print(f"    … +{len(st['problems']) - 25} more (listed in the dashboard's Data Check)")
    save_map_lookup(data)
    if inject_map_data(data, key):
        print(f"  ✅ index.html updated (CARTO key {'embedded' if key else 'NOT found — OSM fallback'})")
        run_deploy_hook()



def main():
    # v4.0 §1: "productivity" (03:00, OIX) now only stages MANPOWER —
    # HKTV Manpower Distribution's own file is still written directly, but
    # the Productivity matrices (HKTV Staff / ODS Ratio) can't be finished
    # until "tableau" (14:00) supplies the new Tableau-sourced order counts.
    # See MANPOWER_STAGING_PATH / finish_productivity_with_orders().
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", choices=["productivity", "tableau", "costreport", "oixbackfill", "map", "newestate", "all"], required=True)
    parser.add_argument("--cost-report-file", default=None,
                        help="v8.0 — with --section costreport: parse this local .xlsx instead of "
                             "downloading from WhatsApp (manual re-run / testing).")
    parser.add_argument("--oix-days", type=int, default=None,
                        help="v10.3 — with --section oixbackfill: check every OIX_Record file of the last N days (from T-1) "
                             "regardless of modified time (default: only files whose modified time is today).")
    parser.add_argument("--force-oix", action="store_true",
                        help="v10.3 — with --section oixbackfill: recompute every day in the window even if its OIX file "
                             "is unchanged (also overrides the 'fewer rows than before' guard).")
    parser.add_argument("--map-excel", default=None,
                        help="Delivery Map v2 — with --section map: estate/address Excel (Delivery Zone in Column C, Latitude, Longitude); "
                             "default = newest matching file in MAP_EXCEL_FOLDER.")
    parser.add_argument("--dry-run", action="store_true",
                        help="v9.5 — with --section newestate: read sources + match Raw only; no geocoding, no files written.")
    parser.add_argument("--limit", type=int, default=None, help="v9.5 — with --section newestate: geocode at most N estates this run.")
    parser.add_argument("--watchlist", default=None,
                        help="v9.5 — with --section newestate: Zone Grouping workbook whose 'New Address…' sheet is the hand-kept watchlist.")
    args = parser.parse_args()

    # v9.5 — new-estate tracker (private + public housing ready for move-in); own schedule (e.g. weekly), NOT part of \"all\".
    if args.section == "newestate":
        from new_estate_tracker import run as run_new_estates
        run_new_estates(raw_loader=lambda: load_map_excel(_find_map_excel(args.map_excel)), in_hk=point_in_hk,
                        dry_run=args.dry_run, limit=args.limit, watchlist=args.watchlist)
        return

    # v10.3 — OIX manpower backfill on its own (run it after an updated OIX_Record has been dropped in OIX_FOLDER).
    if args.section == "oixbackfill":
        run_oix_backfill_section(days=args.oix_days, force=args.force_oix)
        return

    # v8.0 §2 — separate Task Scheduler job (09:00 every Tuesday and Friday); deliberately NOT part of "all".
    if args.section == "costreport":
        run_cost_report_section(args.cost_report_file)
        return

    if args.section == "map":      # Delivery Map v2 — Delivery Map data + CARTO key
        run_map_section(args.map_excel)
        return

    if args.section in ("productivity", "all"):
        run_productivity_section()
    if args.section in ("tableau", "all"):
        fetch_tableau_reports()  # 1. 執行 Playwright 下載 (Delivery Dashboard + Logistics KPI 報表)
        fetch_gmv_report()       # 1b. 下載 GMV 報表 (獨立帳號，v3.0 §3)
        run_section_tableau()    # 2. 執行 CSV 解析、完成 Productivity 計算並寫入
    if args.section == "all":
        try:
            run_map_section()    # Delivery Map v2 — soft-fail: a missing map Excel must not break the KPI run
        except Exception as e:
            print(f"  ⚠️ Delivery Map update skipped: {e}")

if __name__ == "__main__":
    main()