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

# 目錄設定
OIX_FOLDER = os.environ.get("OIX_FOLDER", r"C:\Users\chipanl\Downloads\Digimobi Report")
REPORT_FOLDER = os.environ.get("REPORT_FOLDER", r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\Folder for KPI Dashboard")
DATA_JSON_PATH = os.environ.get("DATA_JSON_PATH", "./public/data.json")
HISTORY_PATH = os.environ.get("HISTORY_PATH", "./history.json")

os.makedirs(REPORT_FOLDER, exist_ok=True)

# 報表精確名稱對應
REPORT_FILES = {
    "report_a": "Summary By RP Group (MTD).csv",
    "report_b": "MTD Summary By RP.csv",
    "report_c": "MTD Summary By RP Group.csv",
    "poor_rating": "Delivery Rating.csv",
    "delay_rate": "Rank_On Time.csv",
    "rfid": "RP Breakdown (7days).csv",  # 保持無空格，依據您之前提供的檔名
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
    ("WH", "WTH"),
    ("CTK", "WTK"), ("WK", "WTK"),
    ("WX", "WTX"),
    ("CTW", "NT-TW"),
    ("ENH", "ETH"),
    ("CTX", "ETX"), ("CX", "ETX"),
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


def process_oix(df):
    """Cleaning + district tagging steps: drop O2O rows from non-LF/LP/ODS/VAN
    users, fill blank Parent Order from Order Number, dedupe, tag District."""
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
    blank_parent = parent.isna() | (parent.astype(str).str.strip() == "")
    filled_parent = order_no.astype(str).str.slice(0, 13)
    df.iloc[:, c_parent] = parent.where(~blank_parent, filled_parent)

    df = df.drop_duplicates(subset=[df.columns[c_parent], df.columns[c_user]])

    df["District"] = df.iloc[:, c_truck].apply(district_from_truck_no)
    unmatched = df["District"].isna().sum()
    if unmatched:
        print(f"  ⚠️ {unmatched} row(s) had a truck number matching none of the 14 "
              f"district patterns — excluded from every total. Check df.iloc[:, {c_truck}].")
    return df


def productivity_for_group(df, prefixes):
    c_user, c_parent = col("E"), col("L")
    sub = df[df.iloc[:, c_user].fillna("").str.startswith(prefixes)]
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


def run_productivity_section():
    print("🚀 開始處理 OIX Productivity 數據...")
    target_date = dt.date.today() - dt.timedelta(days=1)  # T-1
    path = find_oix_file(target_date)
    df = load_oix(path)
    df = process_oix(df)

    hktv_staff = productivity_for_group(df, ("LF", "LP"))
    ods_ratio = productivity_for_group(df, ("ODS", "VAN"))

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

    save_history(history)
    save_data_json(payload)
    print("✅ Productivity 數據處理完成，已寫入 data.json")


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
        fetch_tableau_reports()  # 1. 執行 Playwright 下載
        run_section_tableau()    # 2. 執行 CSV 解析與寫入

if __name__ == "__main__":
    main()