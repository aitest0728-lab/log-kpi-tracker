#!/usr/bin/env python3
"""
LOG · KPI Tracker — Data Pipeline
==================================
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

os.makedirs(REPORT_FOLDER, exist_ok=True)

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
    append_manpower_log(history, target_date.isoformat(), courier_group, driver_group, ods_van_group)
    save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))
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
        print(f"  ⚠️ {INDEX_HTML_PATH!r} has no EMBEDDED_DATA_START/END markers — "
              f"skipping embedded-data update. (Markers were added in v26.0; an "
              f"older index.html won't have them yet.)")
        return
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
    block = f"<script>window.__EMBEDDED_DATA__ = {payload};</script>"

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


def save_productivity_history(daily_dict):
    """Writes ./public/productivity_history.json as {"daily": {...}} — the
    exact shape the dashboard's loadDataSource() fetches. Recomputed fresh
    from history.json's full log every run (see trimmed_daily_productivity),
    so this file is always a derived, disposable window — same pattern as
    data.json itself, never hand-edited or incrementally patched."""
    os.makedirs(os.path.dirname(PRODUCTIVITY_HISTORY_PATH) or ".", exist_ok=True)
    with open(PRODUCTIVITY_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump({"daily": daily_dict}, f, ensure_ascii=False, indent=2)
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
        m = matrices.get(metric_key, {}).get("actual", {})
        return {"monthKey": current_month, "overall": m.get("overall"), "districts": m.get("districts", {})}

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
    required_files = {k: v for k, v in REPORT_FILES.items() if k != "gmv"}
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
        trimmed = finish_productivity_with_orders(history, matrices, staging, order_totals_t1,
                                                    order_totals_mtd, productivity_target_date)
        save_productivity_history(trimmed)

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
    t4_date = today - dt.timedelta(days=4)
    rfid_increment = parse_rfid(t4_date)

    month_key = today.strftime("%Y-%m")
    rfid_state = payload.setdefault("rfidMonthly", {})
    bucket = rfid_state.setdefault(month_key, {})
    days_ledger = bucket.setdefault("days", {})

    # One-time migration: older data.json files (before this fix) stored a
    # running "overall"/"districts" total directly with no day-by-day ledger,
    # so their accumulated numbers can't be trusted (they may include
    # duplicate same-day additions from repeated runs) — drop them and start
    # the ledger fresh from today.
    if "overall" in bucket and "days" not in bucket:
        print(f"  ⚠️ {month_key} had an old-style running total with no per-day ledger — "
              f"resetting it, since past duplicate-run inflation can't be un-mixed from it. "
              f"Starting a fresh ledger from today.")
        bucket = {"days": {}}
        rfid_state[month_key] = bucket
        days_ledger = bucket["days"]

    days_ledger[t4_date.isoformat()] = rfid_increment  # overwrite, not add

    # v28.0 — second, DURABLE copy of the same increment, never trimmed
    # (unlike payload["rfidMonthly"] below, which keeps only the 2 most
    # recent months). This is what build_other_aspects_monthly()'s
    # rfid_all_daily_mtd() reads to give RFID Missing Tote the same
    # every-month "daily" MTD history the other two Other Aspects metrics
    # already have — without it, a closed month's day-by-day path would be
    # lost forever once it aged out of payload["rfidMonthly"].
    history.setdefault("rfidDailyLedger", {})[t4_date.isoformat()] = rfid_increment

    bucket["overall"] = round(sum(d["overall"] for d in days_ledger.values()), 2)
    bucket["districts"] = {
        dist: round(sum(day["districts"][dist] for day in days_ledger.values()), 2)
        for dist in DISTRICTS
    }

    matrices["rfidMissingTote"] = {
        "actual": {"overall": bucket["overall"], "districts": bucket["districts"]},
        "forecast": {
            "overall": prorate_forecast(bucket["overall"], today.day, total_days),
            "districts": {d: prorate_forecast(bucket["districts"][d], today.day, total_days) for d in DISTRICTS},
        },
        "target": RFID_TOTE_TARGETS,
        "asOf": today.isoformat(),
        "monthKey": month_key,
    }

    # v4.0 §4 — archive any month that just closed into history.json (the
    # durable store) before data.json's rfidMonthly window gets trimmed to
    # the last 2 months below — "Other Aspects Tracking" needs every closed
    # month kept permanently, the same way build_gmv_monthly() does for GMV.
    closed_archive = history.setdefault("rfidMonthlyClosed", {})
    for k, v in rfid_state.items():
        if k != month_key and k not in closed_archive:
            closed_archive[k] = {"overall": v.get("overall"), "districts": v.get("districts", {})}

    keep_keys = sorted(rfid_state.keys())[-2:]
    payload["rfidMonthly"] = {k: rfid_state[k] for k in keep_keys}

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


def main():
    # v4.0 §1: "productivity" (03:00, OIX) now only stages MANPOWER —
    # HKTV Manpower Distribution's own file is still written directly, but
    # the Productivity matrices (HKTV Staff / ODS Ratio) can't be finished
    # until "tableau" (14:00) supplies the new Tableau-sourced order counts.
    # See MANPOWER_STAGING_PATH / finish_productivity_with_orders().
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", choices=["productivity", "tableau", "all"], required=True)
    args = parser.parse_args()

    if args.section in ("productivity", "all"):
        run_productivity_section()
    if args.section in ("tableau", "all"):
        fetch_tableau_reports()  # 1. 執行 Playwright 下載 (Delivery Dashboard + Logistics KPI 報表)
        fetch_gmv_report()       # 1b. 下載 GMV 報表 (獨立帳號，v3.0 §3)
        run_section_tableau()    # 2. 執行 CSV 解析、完成 Productivity 計算並寫入

if __name__ == "__main__":
    main()