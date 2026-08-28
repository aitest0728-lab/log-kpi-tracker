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

# 目錄設定
OIX_FOLDER = os.environ.get("OIX_FOLDER", r"C:\Users\chipanl\Downloads\Digimobi Report")
REPORT_FOLDER = os.environ.get("REPORT_FOLDER", r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\Folder for KPI Dashboard")
# v3.0 §4: where the latest "Logistics_Staff_List_YYYYMMDD.xlsx" lives.
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
# How many days of raw order-count/manpower history productivity_history.json
# carries. history.json (not this) is the durable full log, so raising this
# later doesn't lose anything already run — it just widens the served window.
DAILY_PRODUCTIVITY_KEEP_DAYS = int(os.environ.get("DAILY_PRODUCTIVITY_KEEP_DAYS", "60"))

os.makedirs(REPORT_FOLDER, exist_ok=True)

# 報表精確名稱對應
REPORT_FILES = {
    "report_a": "Summary By RP Group (MTD).csv",
    "report_b": "MTD Summary By RP.csv",
    "report_c": "MTD Summary By RP Group.csv",
    "poor_rating": "Delivery Rating.csv",
    "delay_rate": "Rank_On Time.csv",
    "rfid": "RP Breakdown (7days).csv",  # 保持無空格，依據您之前提供的檔名
    "gmv": "Sheet 1.csv",  # v3.0 §3 — fixed download filename, per spec
}

# v3.0 §4: the 6 position codes that get classified into Courier / Driver
# for the HKTV Manpower Distribution tab. Matched case-insensitively /
# whitespace-trimmed against the Staff List's "Position" column.
COURIER_POSITIONS = {"COURIER", "SENIOR COURIER"}
DRIVER_POSITIONS = {"DRIVER", "DRIVER II", "DRIVER AT", "DRIVER C"}
MANPOWER_GROUP_POSITIONS = {"courier": COURIER_POSITIONS, "driver": DRIVER_POSITIONS}

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
    {
        "file_key": "poor_rating",
        "sheet_name": "Delivery Rating",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/LogisticsKPIReport/LogisticsKPI"
    },
    {
        "file_key": "delay_rate",
        "sheet_name": "Rank_On Time",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/LogisticsKPIReport/LogisticsKPI"
    },
    {
        "file_key": "rfid",
        "sheet_name": "RP Breakdown (7days)",
        "url": "https://inhouse-analytics.hktv.com.hk/#/views/RFIDReport_V3/RFIDReport-lastactiondate"
    }
]

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]
DISTRICT_MATCH_ORDER = sorted(DISTRICTS + ["NT-YT"], key=len, reverse=True)
MONTH_RE = re.compile(r"^\d{4}年\d{1,2}月$")


# =============================================================================
# 2. 共用 Helper 函數
# =============================================================================
def today_hkt():
    return dt.date.today()

def normalize_district_code(code):
    return "NT-TW" if code == "NT-YT" else code

def district_from_party_name(name):
    if not isinstance(name, str): return None
    for code in DISTRICT_MATCH_ORDER:
        if name.startswith(code):
            return normalize_district_code(code)
    return None

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
    return False  # 找不到就當作未選取，走原本點擊流程（不影響原本能成功的報表）


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
                continue

            print(f"  ⬇️ 正在下載報表: {sheet_name} ...")

            # Download 按鈕本身通常撐得住，用一般 smart_click 找就好。
            # 真正容易「一閃即逝」的是點下去之後彈出的選單/選項，所以那些
            # 一律改用 fast_click（Playwright 內建 ~100ms 頻率的 auto-wait），
            # 而不是 smart_click 自訂的 1 秒間隔 polling。
            dl_selectors = ["#download", '[aria-label="Download"]', "button:has-text('Download')"]
            if not fast_click(page, dl_selectors, 4000):
                if not smart_click(page, dl_selectors, 15):
                    print(f"  ❌ smart_click 也無法點擊 Download 按鈕，跳過 {sheet_name}")
                    continue

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
                    continue
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
            if is_sheet_already_selected(page, sheet_name):
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
            except Exception as e:
                print(f"  ❌ 下載 {sheet_name} 失敗: {e}")

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

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--disable-popup-blocking"])
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        print("  🌐 導航至 GMV Tableau 登入頁面...")
        page.goto(TABLEAU_GMV_URL)
        page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
        page.locator("input[type='text'], input[name='username']").first.fill(TABLEAU_GMV_USER)
        page.locator("input[type='password'], input[name='password']").first.fill(TABLEAU_GMV_PASS)
        page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()

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

def parse_percent(val):
    if pd.isna(val): return None
    s = str(val).replace("%", "").strip()
    return float(s) if s and s.lower() != "nan" else None

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

def parse_delay_rate():
    df = load_crosstab(REPORT_FILES["delay_rate"], {"District1"})
    per_district, overall = {}, None
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        val = parse_percent(row.iloc[6])
        if label in DISTRICTS: per_district[label] = val
        elif label.lower() == "grand total": overall = val
    return {"overall": overall, "districts": {d: per_district.get(d) for d in DISTRICTS}}

def parse_poor_rating():
    df = load_crosstab(REPORT_FILES["poor_rating"], {"district"})
    per_district, overall = {}, None
    for _, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        val = parse_percent(row.iloc[5])
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
        code = str(row.iloc[1]).strip()
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
             is present, its values are merged into "NT-TW" (v3.0 §3: "if
             the column header has NT-YT, please group ... with NT-TW
             first") via normalize_district_code(), same helper §1/§2 use.
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


def find_latest_staff_list():
    """v3.0 §4 — finds the most-recently-dated
    'Logistics_Staff_List_YYYYMMDD.xlsx' inside STAFF_LIST_FOLDER (YYYYMMDD
    = the file's own "final update date", per spec — so we sort on that,
    not on filesystem mtime)."""
    pattern = re.compile(r"^Logistics_Staff_List_(\d{8})\.xlsx$", re.IGNORECASE)
    candidates = []
    if os.path.isdir(STAFF_LIST_FOLDER):
        for fname in os.listdir(STAFF_LIST_FOLDER):
            m = pattern.match(fname)
            if m:
                candidates.append((m.group(1), fname))
    if not candidates:
        raise FileNotFoundError(
            f"No 'Logistics_Staff_List_YYYYMMDD.xlsx' file found in {STAFF_LIST_FOLDER!r}."
        )
    candidates.sort()  # YYYYMMDD sorts correctly as a string
    _, latest_fname = candidates[-1]
    return os.path.join(STAFF_LIST_FOLDER, latest_fname)


def load_staff_position_map():
    """v3.0 §4 — Employee No. (Column A) -> Position (Column G), from the
    latest staff list. Position values are upper-cased/stripped so they
    compare cleanly against LEADER_EXCLUDE_POSITIONS / COURIER_POSITIONS /
    DRIVER_POSITIONS (which are already all-caps)."""
    path = find_latest_staff_list()
    df = pd.read_excel(path, dtype=str)
    emp_col = df.columns[col("A")]
    pos_col = df.columns[col("G")]
    out = {}
    for emp, pos in zip(df[emp_col], df[pos_col]):
        if pd.isna(emp):
            continue
        out[str(emp).strip()] = "" if pd.isna(pos) else str(pos).strip().upper()
    print(f"  📋 Loaded staff position map from {os.path.basename(path)} ({len(out)} employees)")
    return out


def process_oix(df, position_map=None):
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
    unaffected."""
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


def productivity_for_group(df, prefixes, exclude_positions=None):
    """exclude_positions (v3.0 §4.1): staff whose "Position" column falls in
    this set are dropped BEFORE computing manpower (unique-user counts) —
    but their orders still count if the same order also has other couriers
    attached (excluding a row just means that row's user isn't tallied as
    manpower). Positions come from process_oix()'s "Position" column, so
    this only has an effect when that run had a position_map available."""
    c_user, c_parent = col("E"), col("L")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(prefixes)]
    if exclude_positions:
        sub = sub[~sub["Position"].fillna("").isin(exclude_positions)]
    per_district = {}
    for d in DISTRICTS:
        rows = sub[sub["District"] == d]
        unique_users = rows.iloc[:, c_user].nunique()
        order_count = rows.iloc[:, c_parent].nunique()
        per_district[d] = {
            "orderCount": int(order_count),
            "userCount": int(unique_users),
            "productivity": round(order_count / unique_users, 2) if unique_users else None,
        }
    total_orders = sum(v["orderCount"] for v in per_district.values())
    total_users = sum(v["userCount"] for v in per_district.values())
    overall = round(total_orders / total_users, 2) if total_users else None
    return {"overall": overall, "districts": per_district}


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


def run_productivity_section():
    print("🚀 開始處理 OIX Productivity 數據...")
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

    hktv_staff = productivity_for_group(df, ("LF", "LP"), exclude_positions=LEADER_EXCLUDE_POSITIONS)
    ods_ratio = productivity_for_group(df, ("ODS", "VAN"))
    courier_group = manpower_distribution_for_group(df, "courier")
    driver_group = manpower_distribution_for_group(df, "driver")

    history = load_history()
    payload = load_data_json()
    matrices = payload.setdefault("matrices", {})

    for key, group in (("hktvStaff", hktv_staff), ("odsRatio", ods_ratio)):
        append_history(history, key, target_date.isoformat(), group["overall"],
                        {d: group["districts"][d]["productivity"] for d in DISTRICTS})
        fc_overall, fc_districts = rolling_average(history, key, 7, dt.date.today())
        matrices[key] = {
            "actual": {"overall": group["overall"],
                       "districts": {d: group["districts"][d]["productivity"] for d in DISTRICTS}},
            "forecast": {"overall": fc_overall, "districts": fc_districts},
            "asOf": target_date.isoformat(),
        }

    # Dashboard's Productivity Detail / Daily Records tabs (week-to-week,
    # month-to-month, and raw order-count-per-manpower history). This is
    # purely additive on top of the block above — hktv_staff/ods_ratio are
    # the exact same already-computed groups (LF/LP and ODS/VAN unique-user
    # counts via productivity_for_group(), untouched), just also persisted
    # in their raw orderCount/manpower form instead of only as the reduced
    # productivity ratio the matrices/forecast block above keeps. Written to
    # its own file (productivity_history.json), not data.json — see the
    # PRODUCTIVITY_HISTORY_PATH comment above.
    append_daily_productivity_log(history, target_date.isoformat(), hktv_staff, ods_ratio)
    save_productivity_history(trimmed_daily_productivity(history, DAILY_PRODUCTIVITY_KEEP_DAYS))

    # v3.0 §4: HKTV Manpower Distribution — same daily-log / trimmed-window
    # pattern as productivity_history.json above, own file.
    append_manpower_log(history, target_date.isoformat(), courier_group, driver_group)
    save_manpower_history(trimmed_manpower_distribution(history, DAILY_PRODUCTIVITY_KEEP_DAYS))

    save_history(history)
    save_data_json(payload)
    print("✅ Productivity 數據處理完成，已寫入 data.json")


def _group_to_log_shape(group):
    """productivity_for_group() gives {overall, districts:{d:{orderCount,userCount,productivity}}}.
    The dashboard's dailyProductivity schema wants {districts:{d:{orderCount,manpower}}, total:{...}}
    (manpower = the same unique-user count, just under the name the dashboard uses)."""
    districts = {
        d: {"orderCount": v["orderCount"], "manpower": v["userCount"]}
        for d, v in group["districts"].items()
    }
    total_orders = sum(v["orderCount"] for v in districts.values())
    total_manpower = sum(v["manpower"] for v in districts.values())
    return {"districts": districts, "total": {"orderCount": total_orders, "manpower": total_manpower}}


def append_daily_productivity_log(history, date_str, hktv_group, ods_group):
    log = history.setdefault("dailyProductivityLog", {})
    log[date_str] = {
        "hktvStaff": _group_to_log_shape(hktv_group),
        "odsRatio": _group_to_log_shape(ods_group),
    }


def trimmed_daily_productivity(history, keep_days):
    """Recomputed fresh from history.json (the durable store) every run, same
    pattern as rfidMonthly below — data.json only ever carries a recent
    window so it doesn't grow unbounded, while history.json keeps everything."""
    log = history.get("dailyProductivityLog", {})
    recent_dates = sorted(log.keys())[-keep_days:]
    return {d: log[d] for d in recent_dates}


def append_manpower_log(history, date_str, courier_group, driver_group):
    """v3.0 §4 — durable full log of daily HKTV Courier/Driver headcount,
    same shape/role as append_daily_productivity_log() above."""
    log = history.setdefault("manpowerDistributionLog", {})
    log[date_str] = {"courier": courier_group, "driver": driver_group}


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
    """Basket Size's denominator (v3.0 §3: "GMV / Total Parent Order") —
    HKTV Staff + ODS/VAN order counts combined for date_str, read from the
    SAME dailyProductivityLog entry the Productivity Detail tab uses (see
    append_daily_productivity_log() / §1). Returns (overall, {district:count})
    — any value is None where that day's OIX processing never ran."""
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

    current_month = dt.date.today().strftime("%Y-%m")
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
    missing = [f for f in REPORT_FILES.values() if not os.path.exists(os.path.join(REPORT_FOLDER, f))]
    if missing:
        raise FileNotFoundError(f"Missing report file(s) in {REPORT_FOLDER!r}: {missing}. "
                                 f"fetch_tableau_reports() should have downloaded these first.")

    today = dt.date.today()
    total_days = (dt.date(today.year + (today.month == 12), (today.month % 12) + 1, 1) - dt.timedelta(days=1)).day

    history = load_history()
    payload = load_data_json()
    matrices = payload.setdefault("matrices", {})

    # --- Delay Rate & Poor Rating (30-day rolling forecast) ---
    delay = parse_delay_rate()
    poor = parse_poor_rating()
    for key, val in (("delayRate", delay), ("poorRating", poor)):
        append_history(history, key, today.isoformat(), val["overall"], val["districts"])
        fc_overall, fc_districts = rolling_average(history, key, 30, today)
        matrices[key] = {
            "actual": val,
            "forecast": {"overall": fc_overall, "districts": fc_districts},
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
        "asOf": today.isoformat(),
    }

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
        "asOf": today.isoformat(),
        "monthKey": month_key,
    }

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

    save_history(history)
    save_data_json(payload)
    print("✅ 報表解析完成，已寫入 data.json")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", choices=["productivity", "tableau", "all"], required=True)
    args = parser.parse_args()

    if args.section in ("productivity", "all"):
        run_productivity_section()
    if args.section in ("tableau", "all"):
        fetch_tableau_reports()  # 1. 執行 Playwright 下載 (6 份報表，主帳號)
        fetch_gmv_report()       # 1b. 下載 GMV 報表 (獨立帳號，v3.0 §3)
        run_section_tableau()    # 2. 執行 CSV 解析與寫入

if __name__ == "__main__":
    main()