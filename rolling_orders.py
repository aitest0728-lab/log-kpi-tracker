#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_orders.py  -  KPI Dashboard Version 11.14  "Rolling Orders"   (MFNAS report + Tableau "AI Fetching for Order Date = T-1" mapping)

One run per day, at 06:00 (T-0). It reports the orders created YESTERDAY (T-1):
  1. Logs in to TKO-MFNAS (Synology File Station API, CRAB_NAS_USER / CRAB_NAS_PASS - same as project_crab.py),
     opens  ECOM-MFNAS / ECOM-MFNAS_hktvAccumulatedDeliveryLeadTimeReport / <T-0 as YYYYMMDD>/  and downloads its one csv.
     The csv holds the full list of orders with Order Date = T-1, so it is saved to "Folder for Rolling Orders" as
         Rolling Order Count (MMM-DD of T-1).csv        (same name every run => overwritten)
  1b. (v11.3) Right after the MFNAS download it logs in to Tableau (same TABLEAU_GMV_USER / TABLEAU_GMV_PASS account as the
     GMV report in kpi_pipeline.py), opens the view "AI Fetching for Order Date = T-1", types T-1 as DD/MM/YYYY into the
     delivery_date START-date box (top right), waits 5 minutes for the page to refresh, then Download -> Crosstab -> CSV
     (the sheet is already pre-selected, no thumbnail click). Saved to "Folder for Rolling Orders" as
         AI Fetching for Order Date T-1 (MMM-DD of T-1).csv        (overwritten every run; UTF-8 or UTF-16, with or without the waybill column - all read)
  1c. (v11.9) Once the day is written to rolling_orders.json the two downloaded csv files (MFNAS + Tableau) are DELETED from the folder.
     Keep them with --keep-files or ROLLING_KEEP_FILES=1; --process-only and files passed with --report-file / --tableau-file are never deleted.
  2. Processes it (column letters as in the file):
       - Order Date (C) and Delivery Date (P) are converted to MM/DD/YYYY
       - rows are removed when  Delivery District (Z) is not one of ETH ETK ETX NT-ST NT-TM NT-TSM NT-TW NT-YT WTH WTX WTK,
         or  Consignment Status (J) contains CANCEL,  or  Delivery Zone (U) = KN015
       - delivery date = T-1 (same as the order date): only the 18:00 - 22:00 Delivery Timeslot (Q) is counted
       - (v11.10) PARENT ORDER mapping + tracking across days. After the district / status / zone / same-day-timeslot filters above:
           a) The Tableau report now holds the Order Date (col D) AND the delivery date (col E). Every MFNAS Order Number (col A) is looked
              up in the Tableau rows of the SAME delivery date (for delivery date = T-1: the Tableau rows with delivery date = T-1) ->
              order_number (col B) -> Parent Order (col A).
           b) An order that is not in the Tableau report uses its own order number as its parent order.
           c) rolling_parent_registry.json.gz (in "Folder for Rolling Orders") remembers, per delivery date, which parent orders the
              earlier order-date records already counted. A parent that an earlier record counted for the SAME delivery date is not counted
              again (e.g. an add-on order that shares the parent of an order placed on a previous day). This replaces the v11.3 test
              "earliest order date of the parent = T-1".
         ORDER count = distinct PARENT ORDER that no earlier record counted for that delivery date, per Delivery Date x district and overall
         WAYBILL count (v11.11) = distinct Waybill Number, with the same rule: a waybill an earlier record already counted for the same
         delivery date is not counted again (the registry keeps the waybills too)
       - counted per (Delivery Date x district); NT-YT is grouped into WTH
  3. Writes rolling_orders.json next to index.html AND embeds it in index.html's window.__EMBEDDED_DATA__ block.
  4. Pushes rolling_orders.json + index.html to the dashboard repo (git pull --rebase -> commit -> push), with an
     automatic fix for the one conflict this always risks: the EMBEDDED_DATA block of index.html, which
     kpi_pipeline.py rewrites too (see push_dashboard). Disable with --skip-push or ROLLING_DEPLOY=none.
  (No WhatsApp message / screenshot is sent any more - the dashboard is the only output.)

Usage (WSL):
    python3 rolling_orders.py                          # the 06:00 cron call (T-0 = today, order date = yesterday)
    python3 rolling_orders.py --date 2026-10-14        # re-run for another T-0 (reports the orders of 2026-10-13)
    python3 rolling_orders.py --skip-push
    python3 rolling_orders.py --skip-download --skip-tableau-download      # re-process the two files already in the folder
    python3 rolling_orders.py --tableau-only --date 2026-10-22              # only (re)download the Tableau csv
    python3 rolling_orders.py --process-only --report-file "Rolling Order Count (Oct-13).csv" --date 2026-10-14

Cron (WSL, machine timezone = HKT) - ONE line; delete the old 13:00 / 18:00 / 00:00 lines (no xvfb-run needed any more - no browser is used):
    0 6 * * *  cd "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker" && python3 rolling_orders.py >> logs/rolling_0600.log 2>&1

Credentials (never hard-code them): environment variables CRAB_NAS_USER / CRAB_NAS_PASS (MFNAS) and TABLEAU_GMV_USER /
TABLEAU_GMV_PASS (Tableau, same as kpi_pipeline.py), or a KEY=VALUE file ("rolling_orders.env", "crab.env" or ".env",
next to this script or one folder up). Needs: pip install playwright && playwright install chromium (already used by kpi_pipeline.py).
"""
import argparse
import contextlib
import fcntl
import gzip
import json
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime, timedelta

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)


# ----------------------------------------------------------------------------- config
def _load_env_file():
    for folder in (HERE, os.path.dirname(HERE)):
        for name in ("rolling_orders.env", "crab.env", ".env"):
            p = os.path.join(folder, name)
            if os.path.exists(p):
                for line in open(p, encoding="utf-8-sig"):
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env_file()

# --- TKO-MFNAS (Synology File Station) - same server + credentials as project_crab.py ---
NAS_BASE_URL = "https://tko-mfnas.hktv.com.hk:5001"
NAS_USERNAME = os.environ.get("CRAB_NAS_USER", "")
NAS_PASSWORD = os.environ.get("CRAB_NAS_PASS", "")
NAS_REPORT_FOLDER = "ECOM-MFNAS_hktvAccumulatedDeliveryLeadTimeReport"
# Where the report folder sits. project_crab.py uses "<share>/<share>/<report folder>/<YYYYMMDD>" (/OIX-MFNAS/OIX-MFNAS/...),
# so the ECOM equivalent is tried first. Set ROLLING_NAS_ROOT to pin it; otherwise the candidates below are tried and,
# if none lists, the shares are searched for the report folder. The path that worked is written to the log.
NAS_ROOT_CANDIDATES = [r for r in (os.environ.get("ROLLING_NAS_ROOT", "").strip().rstrip("/"),
                                   "/ECOM-MFNAS/ECOM-MFNAS", "/ECOM-MFNAS",
                                   "/OIX-MFNAS/OIX-MFNAS",            # project_crab.py's NAS_REMOTE_ROOT
                                   "/OIX-MFNAS/OIX-MFNAS/ECOM-MFNAS", "/OIX-MFNAS/ECOM-MFNAS", "/OIX-MFNAS") if r]
# --- Tableau (v11.3) - same account / host as the GMV report in kpi_pipeline.py ---
TABLEAU_SIGNIN_URL = os.environ.get("TABLEAU_GMV_URL", "https://inhouse-analytics.hktv.com.hk/#/signin")
TABLEAU_AI_VIEW_URL = os.environ.get("TABLEAU_AI_T1_URL",
    "https://inhouse-analytics.hktv.com.hk/#/views/AIFetchingforOrderDateT-1/AIFetchingforOrderDateT-1_1")
# The view must be the one TITLED "AI Fetching for Order Date = T-1" (with the "="). A Tableau URL never shows the "=" (spaces and special
# characters are stripped from the view name, and a "_1" is added when two views end up with the same URL name), so the title found on
# the opened page is logged and compared with this text. TABLEAU_AI_T1_STRICT=1 stops the run when it is not found.
TABLEAU_AI_VIEW_TITLE = os.environ.get("TABLEAU_AI_T1_TITLE", "AI Fetching for Order Date = T-1")
TABLEAU_AI_VIEW_STRICT = os.environ.get("TABLEAU_AI_T1_STRICT", "") == "1"
TABLEAU_USER = os.environ.get("TABLEAU_GMV_USER", "")
TABLEAU_PASS = os.environ.get("TABLEAU_GMV_PASS", "")
TABLEAU_REFRESH_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_REFRESH_WAIT", "300"))   # 5 min after the start date is typed
TABLEAU_ATTEMPTS = max(1, int(os.environ.get("TABLEAU_DOWNLOAD_ATTEMPTS", "3")))      # fresh browser sessions, as kpi_pipeline.py
TABLEAU_RECOVERY_WAIT_S = int(os.environ.get("TABLEAU_RECOVERY_WAIT_S", "20"))
KEEP_FILES = os.environ.get("ROLLING_KEEP_FILES", "0").strip().lower() in ("1", "true", "yes", "on")   # 1 = keep the downloaded csv files
TABLEAU_LOAD_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_LOAD_WAIT", "180"))          # max wait for the view to appear
TABLEAU_HEADLESS = os.environ.get("ROLLING_TABLEAU_HEADLESS", "1").strip().lower() not in ("0", "false", "no")

NAS_CONNECT_TIMEOUT = 10
NAS_READ_TIMEOUT = 120
NAS_POLL_ATTEMPTS = int(os.environ.get("ROLLING_NAS_POLL_ATTEMPTS", "6"))      # the file is generated ~01:30; wait if it is late
NAS_POLL_INTERVAL_SEC = int(os.environ.get("ROLLING_NAS_POLL_INTERVAL", "300"))

# WSL paths (C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\...)
KPI_DIR = os.environ.get("KPI_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker")
# The raw csv is too large to push to the dashboard, so it lives OUTSIDE log-kpi-tracker, one level up:
# C:\Users\chipanl\Downloads\Whatsapp Session\Folder for Rolling Orders
ROLLING_DIR = os.environ.get("ROLLING_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/Folder for Rolling Orders")
PUBLIC_DIR = os.environ.get("KPI_PUBLIC_DIR", os.path.join(KPI_DIR, "public"))   # same folder kpi_pipeline.py serves from
INDEX_HTML = os.environ.get("KPI_INDEX_HTML", os.path.join(PUBLIC_DIR, "index.html"))
ROLLING_JSON = os.environ.get("ROLLING_JSON", os.path.join(PUBLIC_DIR, "rolling_orders.json"))   # next to index.html
LOG_DIR = os.environ.get("ROLLING_LOG_DIR", os.path.join(KPI_DIR, "logs"))
# v11.10 - parent orders already counted by earlier order dates (per delivery date). Outside the dashboard repo, next to the raw csv files.
REGISTRY_PATH = os.environ.get("ROLLING_PARENT_REGISTRY", os.path.join(ROLLING_DIR, "rolling_parent_registry.json.gz"))
MAX_UNMAPPED_SHARE = float(os.environ.get("ROLLING_MAX_UNMAPPED", "0.2"))          # more unmapped orders than this share = wrong / stale Tableau file -> stop
PARTIAL_DAY_SHARE = 0.25            # newest Tableau order date with fewer rows than this share of the day before = exported too early, not complete
REGISTRY_KEEP_DAYS = int(os.environ.get("ROLLING_REGISTRY_KEEP_DAYS", "30"))      # delivery dates older than this (vs the newest order date) are dropped

SLOTS = ["06:00"]                            # Version 11.1: one run a day
KEEP_ORDER_DATES = 400                       # order dates kept in rolling_orders.json (the dashboard shows the latest 3 months
                                             # day by day and collapses older months - they have to stay in the file to be opened)
KEEP_SNAPSHOT_DATES = 14                     # per-slot snapshots are only a short audit trail (the dashboard does not read them)

# auto-push (git). Run from the repo root = KPI_DIR (the folder that holds public/ and history.json).
PUSH_ENABLED = os.environ.get("ROLLING_DEPLOY", "1").strip().lower() not in ("0", "none", "off", "false", "no")
DEPLOY_CWD = os.environ.get("ROLLING_DEPLOY_CWD") or os.environ.get("DEPLOY_CWD") or KPI_DIR
GIT_TIMEOUT = int(os.environ.get("ROLLING_GIT_TIMEOUT", "180"))
PUSH_ATTEMPTS = 3

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]     # the dashboard's 10 districts
VALID_DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "NT-YT", "WTH", "WTX", "WTK"]   # the 11 codes of the spec
DISTRICT_ALIAS = {"NT-YT": "WTH"}            # the Delivery Map groups NT-YT together with WTH
EXCLUDE_ZONE = "KN015"                       # Delivery Zone (U) that is not counted
SAME_DAY_TIMESLOT = "18:00-22:00"            # delivery date = order date: only this Delivery Timeslot is counted (spaces ignored)

# report columns: canonical key -> (header in the file, position = the spec's column letter, 0-based)
REPORT_COLS = {
    "order_no":   ("Order Number", 0),                       # A
    "order_date": ("Order Date", 2),                         # C
    "status":     ("Consignment Status", 9),                 # J
    "waybill":    ("Waybill Number", 11),                    # L
    "optimum":    ("Optimum Delivery Timeslot", 13),         # N
    "first":      ("First Available Delivery Timeslot", 14), # O
    "delivery":   ("Delivery Date", 15),                     # P
    "slot":       ("Delivery Timeslot", 16),                 # Q
    "zone":       ("Delivery Zone", 20),                     # U
    "district":   ("Delivery District", 25),                 # Z
}


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# ----------------------------------------------------------------------------- dates / names
def run_dates(now=None):
    """(T-0, order date). T-0 = the day the 06:00 job runs; the report holds the orders created on T-1."""
    now = now or datetime.now()
    t0 = datetime(now.year, now.month, now.day)
    return t0, t0 - timedelta(days=1)


def mmm_dd(d):
    return d.strftime("%b-%d")                        # Oct-13


def report_path(order_date):
    return os.path.join(ROLLING_DIR, f"Rolling Order Count ({mmm_dd(order_date)}).csv")


def tableau_report_path(order_date):
    return os.path.join(ROLLING_DIR, f"AI Fetching for Order Date T-1 ({mmm_dd(order_date)}).csv")


# ----------------------------------------------------------------------------- MFNAS download
class NASClient:
    """Synology File Station Web API - same calls as project_crab.py (login / list / download)."""

    def __init__(self):
        import requests
        import urllib3
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.s = requests.Session()
        self.sid = None

    def _get(self, path, params, **kw):
        return self.s.get(f"{NAS_BASE_URL}/webapi/{path}", params=params, verify=False,
                          timeout=(NAS_CONNECT_TIMEOUT, NAS_READ_TIMEOUT), **kw)

    def login(self):
        if not NAS_USERNAME or not NAS_PASSWORD:
            raise RuntimeError("CRAB_NAS_USER / CRAB_NAS_PASS not set (env var, rolling_orders.env or crab.env)")
        log("MFNAS: login")
        data = self._get("auth.cgi", {"api": "SYNO.API.Auth", "version": 3, "method": "login",
                                      "account": NAS_USERNAME, "passwd": NAS_PASSWORD,
                                      "session": "FileStation", "format": "sid"}).json()
        if not data.get("success"):
            raise RuntimeError(f"[NAS] login failed: {data}")
        self.sid = data["data"]["sid"]
        log("MFNAS: login successful")

    def logout(self):
        if not self.sid:
            return
        try:
            self._get("auth.cgi", {"api": "SYNO.API.Auth", "version": 1, "method": "logout",
                                   "session": "FileStation", "_sid": self.sid})
        except Exception:
            pass
        self.sid = None

    def list_folder(self, folder_path, quiet=False):
        """[{name, path, isdir}] ; None when the folder cannot be listed (does not exist / no permission)."""
        data = self._get("entry.cgi", {"api": "SYNO.FileStation.List", "version": 2, "method": "list",
                                       "folder_path": folder_path, "_sid": self.sid}).json()
        if not data.get("success"):
            if not quiet:
                log(f"[NAS] cannot list '{folder_path}': {data}")
            return None
        return [{"name": f.get("name", ""), "path": f.get("path"), "isdir": bool(f.get("isdir"))}
                for f in data["data"].get("files", [])]

    def list_shares(self):
        data = self._get("entry.cgi", {"api": "SYNO.FileStation.List", "version": 2,
                                       "method": "list_share", "_sid": self.sid}).json()
        return [s.get("path") for s in data.get("data", {}).get("shares", [])] if data.get("success") else []

    def download(self, remote_path, local_path):
        log(f"[NAS] downloading {os.path.basename(remote_path)} ...")
        r = self._get("entry.cgi", {"api": "SYNO.FileStation.Download", "version": 2, "method": "download",
                                    "path": remote_path, "mode": "open", "_sid": self.sid}, stream=True)
        r.raise_for_status()
        if "json" in (r.headers.get("Content-Type") or "").lower():          # File Station answers errors as json
            raise RuntimeError(f"[NAS] download refused: {r.text[:300]}")
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        total, last = 0, time.time()
        with open(local_path, "wb") as f:
            for chunk in r.iter_content(65536):
                if chunk:
                    f.write(chunk)
                    total += len(chunk)
                    if time.time() - last > 10:
                        log(f"[NAS] ... {total / 1048576:.0f} MB")
                        last = time.time()
        log(f"[NAS] saved {local_path} ({total / 1048576:.1f} MB)")
        return total


def find_date_folder(nas, t0s):
    """Remote path of <report folder>/<YYYYMMDD>; None if the report folder cannot be found at all."""
    for root in NAS_ROOT_CANDIDATES:
        p = f"{root}/{NAS_REPORT_FOLDER}/{t0s}"
        if nas.list_folder(p, quiet=True) is not None:
            log(f"MFNAS: folder found: {p}")
            return p
        # the report folder may exist while today's date folder has not been created yet
        if nas.list_folder(f"{root}/{NAS_REPORT_FOLDER}", quiet=True) is not None:
            log(f"MFNAS: report folder found at {root}/{NAS_REPORT_FOLDER}, but no {t0s} folder yet")
            return p
    # not at any expected place: search the shares (3 levels deep) for the report folder
    log(f"MFNAS: '{NAS_REPORT_FOLDER}' not at the expected paths {NAS_ROOT_CANDIDATES} - searching the shares ...")
    queue = [(s, 1) for s in nas.list_shares()]
    seen = 0
    while queue and seen < 300:
        path, depth = queue.pop(0)
        seen += 1
        for k in nas.list_folder(path, quiet=True) or []:
            if k["name"] == NAS_REPORT_FOLDER:
                log(f"MFNAS: report folder found by search: {k['path']}  (set ROLLING_NAS_ROOT={path} to skip the search)")
                return f"{k['path']}/{t0s}"
            if k["isdir"] and depth < 3:
                queue.append((k["path"], depth + 1))
    return None


def download_report(t0, final_path):
    """Download the csv of the T-0 folder and save it as `final_path` (= Rolling Order Count (T-1 date).csv)."""
    t0s = t0.strftime("%Y%m%d")
    os.makedirs(os.path.dirname(final_path), exist_ok=True)
    nas = NASClient()
    nas.login()
    try:
        folder, pick = None, None
        for attempt in range(1, NAS_POLL_ATTEMPTS + 1):
            folder = folder or find_date_folder(nas, t0s)
            files = [f for f in (nas.list_folder(folder, quiet=True) if folder else None) or []
                     if not f["isdir"] and f["name"].lower().endswith(".csv")]
            if files:
                files.sort(key=lambda f: f["name"])
                pick = files[-1]                  # the folder holds ONE csv; if there are several, the newest timestamp wins
                if len(files) > 1:
                    log(f"MFNAS: WARNING {len(files)} csv files in the folder, using the newest: {pick['name']}")
                break
            log(f"MFNAS: no csv in {folder or NAS_REPORT_FOLDER + '/' + t0s} yet (attempt {attempt}/{NAS_POLL_ATTEMPTS})")
            if attempt < NAS_POLL_ATTEMPTS:
                time.sleep(NAS_POLL_INTERVAL_SEC)
        if not pick:
            raise RuntimeError(f"MFNAS: no csv found for {t0s}. Visible shares: {nas.list_shares()}. "
                               f"Set ROLLING_NAS_ROOT if the report folder sits somewhere else.")
        log(f"MFNAS: {pick['name']}  ->  {os.path.basename(final_path)}")
        tmp = final_path + ".part"
        size = nas.download(pick["path"], tmp)
        if size < 1024:
            raise RuntimeError(f"MFNAS: downloaded file is only {size} bytes")
        os.replace(tmp, final_path)               # same name each run -> the previous copy is covered
        return final_path
    finally:
        nas.logout()


# ----------------------------------------------------------------------------- Tableau download (v11.3)
# Playwright helpers - the same logic as fast_click / smart_click in kpi_pipeline.py (copied, not imported: importing the
# 6,500-line pipeline would run its module-level setup). The Tableau viz lives inside <iframe title="Data Visualisation">.
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


def _read_date_box(page):
    """Text shown in the delivery_date START-date box (div.readoutText ... domLowerText), '' if not readable."""
    for target in _frames(page):
        try:
            loc = target.locator('[dojoattachpoint="domLowerText"]').first
            if loc.count():
                return (loc.inner_text(timeout=2000) or "").strip()
        except Exception:
            continue
    return ""


def _set_start_date(page, order_date):
    """Type T-1 as DD/MM/YYYY into the delivery_date start-date box (top right). Click the readout text -> the hidden
    <input dojoattachpoint="domLowerInput"> appears -> select all, type, Enter. Verified by re-reading the readout text."""
    want = order_date.strftime("%d/%m/%Y")
    for attempt in range(1, 4):
        if not _fast_click(page, ['[dojoattachpoint="domLowerText"]', "div.readoutText.tab-ctrl-formatted-fixedsize"], 15000):
            log(f"Tableau: start-date box not found (attempt {attempt}/3)")
            time.sleep(5)
            continue
        time.sleep(0.5)
        done = False
        for target in _frames(page):
            try:
                inp = target.locator('input[dojoattachpoint="domLowerInput"]').first
                if not inp.count():
                    continue
                inp.click(timeout=3000, force=True)
                inp.press("Control+A")
                inp.fill(want)
                inp.press("Enter")
                done = True
                break
            except Exception:
                continue
        time.sleep(3)
        shown = _read_date_box(page)
        log(f"Tableau: start date typed {want} (attempt {attempt}/3) -> box now shows '{shown}'")
        if done and shown.lstrip("0").replace("/0", "/") == want.lstrip("0").replace("/0", "/"):
            return True
    return False


class _VizActivity:
    """Tracks Tableau's vizql requests (every refresh of the view goes through /vizql/). Used to leave the 5-minute wait early:
    once the date change has produced vizql traffic, nothing is in flight any more and it has been quiet for QUIET_S seconds,
    the view has finished refreshing."""
    QUIET_S = 20

    def __init__(self, page):
        self.inflight, self.count, self.last = set(), 0, time.time()
        page.on("request", self._start)
        page.on("requestfinished", self._end)
        page.on("requestfailed", self._end)

    def _start(self, req):
        if "/vizql/" in req.url:
            self.inflight.add(req); self.count += 1; self.last = time.time()

    def _end(self, req):
        if req in self.inflight:
            self.inflight.discard(req); self.last = time.time()

    def mark(self):
        self.count = 0
        self.last = time.time()

    def refreshed(self):
        return self.count > 0 and not self.inflight and time.time() - self.last >= self.QUIET_S


def _wait_for_view(page, final_path):
    """Wait (up to TABLEAU_LOAD_WAIT_S) until the viz iframe or the date box exists. Re-issues the view URL every 60s in case the
    SPA bounced back to the explore/home page after sign-in. On failure a screenshot is saved next to the csv for diagnosis."""
    end = time.time() + TABLEAU_LOAD_WAIT_S
    last_goto = time.time()
    while time.time() < end:
        try:
            if page.locator('iframe[title="Data Visualisation"]').count() or any(
                    f.locator('[dojoattachpoint="domLowerText"]').count() for f in page.frames):
                log("Tableau: view loaded")
                return
        except Exception:
            pass
        if time.time() - last_goto > 60:
            log(f"Tableau: view not up yet (at {page.url}) - opening it again")
            try:
                page.goto(TABLEAU_AI_VIEW_URL)
            except Exception:
                pass
            last_goto = time.time()
        time.sleep(3)
    shot = os.path.splitext(final_path)[0] + "_view_timeout.png"
    try:
        page.screenshot(path=shot, full_page=True)
        log(f"Tableau: screenshot saved to {shot} (page was at {page.url})")
    except Exception:
        pass
    raise RuntimeError(f"Tableau view did not load within {TABLEAU_LOAD_WAIT_S}s (page at {page.url})")


def _check_view_title(page):
    """v11.14 - log which Tableau view was opened and warn (or stop with TABLEAU_AI_T1_STRICT=1) when its title is not
    TABLEAU_AI_VIEW_TITLE ("AI Fetching for Order Date = T-1", with the "=")."""
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
    ok = any(norm(TABLEAU_AI_VIEW_TITLE) in norm(t) for t in seen)
    found = next((t.strip().splitlines()[0][:120] for t in seen[:1] if t and t.strip()), "?")
    if ok:
        log(f"Tableau: view title OK - '{TABLEAU_AI_VIEW_TITLE}' found (browser title: '{found}')")
        return
    msg = (f"the opened Tableau view does not show the title '{TABLEAU_AI_VIEW_TITLE}' (browser title: '{found}', url {page.url}). "
           f"A view named WITHOUT the '=' ('AI Fetching for Order Date T-1') gets almost the same URL - check TABLEAU_AI_T1_URL")
    if TABLEAU_AI_VIEW_STRICT:
        raise RuntimeError(msg)
    log("WARNING: " + msg)


def _tableau_session(order_date, final_path):
    from playwright.sync_api import sync_playwright
    tmp = final_path + ".part"
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=TABLEAU_HEADLESS, args=["--disable-popup-blocking"])
        try:
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            page.set_viewport_size({"width": 1920, "height": 1080})
            log("Tableau: login")
            page.goto(TABLEAU_SIGNIN_URL)
            page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
            page.locator("input[type='text'], input[name='username']").first.fill(TABLEAU_USER)
            page.locator("input[type='password'], input[name='password']").first.fill(TABLEAU_PASS)
            page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()
            try:                                                        # leave the sign-in page before jumping to the view
                page.wait_for_function("() => !location.hash.includes('signin')", timeout=60000)
            except Exception:
                log(f"Tableau: still on {page.url} after sign-in - continuing")
            time.sleep(5)
            log("Tableau: opening 'AI Fetching for Order Date = T-1'")
            page.goto(TABLEAU_AI_VIEW_URL)
            activity = _VizActivity(page)
            _wait_for_view(page, final_path)
            time.sleep(15)
            _check_view_title(page)
            activity.mark()
            if not _set_start_date(page, order_date):
                raise RuntimeError(f"could not set the delivery_date start date to {order_date:%d/%m/%Y}")
            log(f"Tableau: waiting for the page to refresh (up to {TABLEAU_REFRESH_WAIT_S}s, or as soon as the view has reloaded) ...")
            t_start = time.time()
            while time.time() - t_start < TABLEAU_REFRESH_WAIT_S:
                page.wait_for_timeout(2000)
                if activity.refreshed():
                    log(f"Tableau: view refreshed after {time.time() - t_start:.0f}s ({activity.count} vizql requests) - continuing")
                    break
            else:
                log(f"Tableau: no refresh completion detected - waited the full {TABLEAU_REFRESH_WAIT_S}s")
            shown = _read_date_box(page)
            if shown and shown.lstrip("0").replace("/0", "/") != order_date.strftime("%d/%m/%Y").lstrip("0").replace("/0", "/"):
                raise RuntimeError(f"start date reverted to '{shown}' during the refresh wait")
            dl_sel = ["#download", '[aria-label="Download"]', "button:has-text('Download')"]
            if not (_fast_click(page, dl_sel, 4000) or _smart_click(page, dl_sel, 15)):
                raise RuntimeError("Download button not found")
            ct_sel = ["#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
                      "#viz-viewer-toolbar-download-menu div:nth-of-type(3) span",
                      "xpath=//*[@id='viz-viewer-toolbar-download-menu']/div[3]",
                      '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
                      "span:has-text('Crosstab')", "text='Crosstab'"]
            if not (_fast_click(page, ct_sel, 4000) or _smart_click(page, ct_sel, 10)):
                raise RuntimeError("Crosstab option not found")
            time.sleep(2)          # the sheet is already pre-selected in the dialog: NO thumbnail click (it is a toggle)
            csv_sel = ["#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
                       "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2)",
                       "label:has-text('CSV')", "text='CSV'"]
            if not _fast_click(page, csv_sel, 2000):
                _smart_click(page, csv_sel, 5)
            time.sleep(1)
            ok_sel = ["#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id button",
                      "button[aria-label='Download Crosstab']", "button[aria-label='Download']",
                      'button[data-tb-test-id="export-crosstab-export-Button"]']
            with page.expect_download(timeout=180000) as info:      # ~730k rows: give the export more than the usual 60s
                if not (_fast_click(page, ok_sel, 3000) or _smart_click(page, ok_sel, 30)):
                    raise RuntimeError("Download Crosstab button not found")
            os.makedirs(os.path.dirname(final_path), exist_ok=True)
            info.value.save_as(tmp)
        finally:
            browser.close()
    return tmp


def read_tableau_csv(path, nrows=None):
    """Tableau crosstab as strings. Every layout is accepted: UTF-16 tab-separated or UTF-8 (comma / tab), with or without the waybill_no
    column - (Parent Order, order_number, [waybill_no,] Order Date, delivery_date); see tableau_columns."""
    with open(path, "rb") as f:
        head = f.read(4096)
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        enc, sep = "utf-16", "\t"
    else:
        enc = "utf-8-sig"
        first = head.decode("utf-8-sig", errors="ignore").split("\n", 1)[0]
        sep = "\t" if "\t" in first else ","
    return pd.read_csv(path, encoding=enc, sep=sep, dtype=str, keep_default_na=False, index_col=False, nrows=nrows)


def _check_tableau_file(path):
    """Cheap sanity check of the downloaded crosstab (UTF-16 tab csv, order_number in column B) before it replaces the old copy."""
    head = read_tableau_csv(path, nrows=5)
    if head.shape[1] < 3 or "order_number" not in [c.strip() for c in head.columns]:
        raise RuntimeError(f"unexpected Tableau csv layout, columns = {list(head.columns)[:6]}")


def download_tableau_report(order_date, final_path):
    """Fresh browser session per attempt (up to TABLEAU_ATTEMPTS), as the GMV download does."""
    if not TABLEAU_USER or not TABLEAU_PASS:
        raise RuntimeError("TABLEAU_GMV_USER / TABLEAU_GMV_PASS not set (env var, rolling_orders.env or .env)")
    last = None
    for n in range(1, TABLEAU_ATTEMPTS + 1):
        if n > 1:
            log(f"Tableau: retry {n}/{TABLEAU_ATTEMPTS} (fresh browser + login)")
            time.sleep(TABLEAU_RECOVERY_WAIT_S)
        try:
            tmp = _tableau_session(order_date, final_path)
            if os.path.getsize(tmp) < 1024:
                raise RuntimeError(f"downloaded file is only {os.path.getsize(tmp)} bytes")
            _check_tableau_file(tmp)
            os.replace(tmp, final_path)                # same name each run -> the previous copy is covered
            log(f"Tableau: saved {final_path} ({os.path.getsize(final_path) / 1048576:.1f} MB)")
            return final_path
        except Exception as ex:
            last = ex
            log(f"Tableau: attempt {n}/{TABLEAU_ATTEMPTS} failed: {type(ex).__name__}: {ex}")
    raise RuntimeError(f"Tableau 'AI Fetching for Order Date = T-1' download failed after {TABLEAU_ATTEMPTS} sessions: {last}")


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


# ----------------------------------------------------------------------------- Tableau parent-order mapping (v11.3 / v11.10)
def _parse_ymd(series):
    """'2026年10月21日' (the Chinese-locale export), or 2026-10-21 / 21/10/2026: three numbers; a 4-digit one tells the order."""
    n = series.str.extract(r"(\d+)\D+(\d+)\D+(\d+)").astype(float)
    ymd = n[0] > 31
    y = n[0].where(ymd, n[2]); mo = n[1]; d = n[2].where(ymd, n[0])
    return pd.to_datetime(pd.DataFrame({"year": y, "month": mo, "day": d}), errors="coerce")


class ParentMap:
    """The Tableau "AI Fetching for Order Date = T-1" crosstab as a lookup  (order number, delivery date) -> Parent Order.
    v11.10: the report now holds the Order Date (col D) AND the delivery date (col E). An MFNAS order is mapped through the Tableau
    rows of the SAME delivery date - for delivery date = T-1 that is exactly "take the Tableau rows with delivery date = T-1,
    then map". A report without the delivery column (old 3 / 4-column layout) maps by order number alone."""

    def __init__(self, lookup, has_delivery, orders, max_order_date, min_delivery, rows, complete_through=None, last_day_rows=0):
        self.lookup, self.has_delivery, self.orders = lookup, has_delivery, orders
        self.max_order_date, self.min_delivery, self.rows = max_order_date, min_delivery, rows
        # newest order date whose orders are all in the report. The newest order date of an export taken early in the day holds only
        # the first hours of that day, so it is NOT complete (v11.12).
        self.complete_through = complete_through if complete_through is not None else max_order_date
        self.last_day_rows = last_day_rows

    def parents(self, order_no, delivery):
        """Parent Order of every (order number, delivery date) pair; NaN where the pair is not in the report."""
        k = order_no.astype("string")
        if self.has_delivery:
            k = k + "|" + delivery.dt.strftime("%Y%m%d")
        return k.map(self.lookup)


def tableau_columns(t):
    """Column positions of the Tableau crosstab -> (parent, order_no, order_date, delivery_date or None, waybill or None).
    Headers decide (so the waybill_no column may be present or removed); only when a header is missing / renamed the layout is
    taken from the number of columns:
        5  Parent Order, order_number, waybill_no, Order Date, delivery_date
        4  Parent Order, order_number, Order Date, delivery_date          (waybill column removed - the daily download)
           or Parent Order, order_number, waybill_no, Order Date          (old layout, no delivery date)
        3  Parent Order, order_number, Order Date
    A trailing empty column (Tableau ends every line with a tab) is ignored."""
    cols = [str(c).strip() for c in t.columns]
    real = [i for i, c in enumerate(cols) if c and not c.lower().startswith("unnamed")]
    n = (max(real) + 1) if real else len(cols)
    norm = [c.lower().replace("_", " ") for c in cols]
    def find(*keys):
        return next((i for i, c in enumerate(norm[:n]) if any(k in c for k in keys)), None)
    oi = find("order number")
    di = find("order date")
    ei = find("delivery date")
    wi = find("waybill")
    pi = find("parent")
    if pi is None:
        pi = 0
    if oi is None:
        oi = 1
    if di is None or (n >= 4 and ei is None and wi is None):
        if n >= 5:
            wi, di, ei = (wi if wi is not None else 2), 3, 4
        elif n == 4:
            third = _parse_ymd(t.iloc[:200, 2].astype(str).str.strip())
            if third.notna().mean() > 0.9:          # col C is a date -> (Parent, order, Order Date, delivery date)
                wi, di, ei = None, 2, 3
            else:                                   # col C is a waybill -> old layout
                wi, di, ei = 2, 3, None
        else:
            wi, di, ei = None, 2, None
    return pi, oi, di, ei, wi


def load_parent_map(path):
    """Tableau crosstab -> ParentMap. Columns: A = Order (Unlimited & Add-on Group to Parent) = Parent Order, order_number (B),
    the "Day, Month, Year of Order Date" column and, since v11.10, the "Day, Month, Year of delivery_date" column. The waybill_no
    column (C) is optional: without it the order date is column C and the delivery date column D (see tableau_columns)."""
    t = read_tableau_csv(path)
    if t.shape[1] < 3:
        raise ValueError(f"{os.path.basename(path)}: expected at least 3 columns (Parent Order, order_number, order date), got {t.shape[1]}")
    pi, oi, di, ei, wi = tableau_columns(t)
    names = f"parent = {t.columns[pi]!r}, order = {t.columns[oi]!r}, order date = {t.columns[di]!r}, " + \
            (f"delivery date = {t.columns[ei]!r}" if ei is not None else "no delivery date")
    log(f"Tableau columns ({os.path.basename(path)}): {names}")
    t = t.iloc[:, [pi, oi, di] + ([ei] if ei is not None else [])]
    t.columns = ["parent", "order_no", "odate"] + (["ddate"] if ei is not None else [])
    t["parent"] = t["parent"].str.strip()
    t["order_no"] = t["order_no"].str.strip()
    t["odate"] = _parse_ymd(t["odate"])
    t = t[(t["parent"] != "") & (t["order_no"] != "")]
    bad = int(t["odate"].isna().sum())
    if bad:
        log(f"WARNING: {bad} Tableau rows have an unreadable order date (ignored)")
    if len(t) and bad == len(t):
        raise ValueError(f"{os.path.basename(path)}: no Tableau order date could be read")
    has_delivery = ei is not None
    min_delivery = None
    if has_delivery:
        t["ddate"] = _parse_ymd(t["ddate"])
        bad_d = int(t["ddate"].isna().sum())
        if bad_d:
            log(f"WARNING: {bad_d} Tableau rows have an unreadable delivery date - they cannot be used for the mapping")
        if len(t) and bad_d == len(t):
            raise ValueError(f"{os.path.basename(path)}: no Tableau delivery date could be read")
        t = t[t["ddate"].notna()]
        t["key"] = t["order_no"] + "|" + t["ddate"].dt.strftime("%Y%m%d")
        min_delivery = t["ddate"].min()
    else:
        log(f"WARNING: {os.path.basename(path)} has no delivery date column (old layout) - orders are mapped by order number only")
        t["key"] = t["order_no"]
    multi = int((t.groupby("key")["parent"].nunique() > 1).sum())
    if multi:
        log(f"WARNING: {multi} order numbers (per delivery date) map to more than one Parent Order in the Tableau report - the first one is used")
    lookup = t.drop_duplicates("key").set_index("key")["parent"]
    lookup.index = lookup.index.astype("string")
    per_day = t.groupby("odate").size().sort_index()
    last_rows = int(per_day.iloc[-1])
    complete = per_day.index[-1]
    if len(per_day) > 1 and last_rows < PARTIAL_DAY_SHARE * per_day.iloc[-2]:      # a few hours of the newest day only
        complete = per_day.index[-2]
    pm = ParentMap(lookup, has_delivery, frozenset(t["order_no"]), t["odate"].max(), min_delivery, len(t), complete, last_rows)
    log(f"Tableau map: {len(t)} rows, {lookup.size} order numbers{' x delivery dates' if has_delivery else ''}, "
        f"{t['parent'].nunique()} parent orders, order dates up to {pm.max_order_date:%Y-%m-%d}"
        + (f", delivery dates from {min_delivery:%Y-%m-%d}" if has_delivery else ""))
    if complete != pm.max_order_date:
        log(f"WARNING: the newest Order Date in {os.path.basename(path)} ({pm.max_order_date:%Y-%m-%d}) has only {last_rows} rows "
            f"({100.0 * last_rows / per_day.iloc[-2]:.0f}% of the day before) - the report was exported before that day was complete; "
            f"orders are complete only through {complete:%Y-%m-%d}")
    return pm


# ----------------------------------------------------------------------------- report processing
def _read_csv_tolerant(path, cols):
    """Rows with more / fewer fields than the header (an unquoted comma in a text field, a trailing comma) are cut / padded to
    the header width instead of aborting the whole run - project_crab.py needed the same for its MFNAS reports."""
    import csv
    rows, bad = [], 0
    with open(path, encoding="utf-8-sig", newline="") as f:
        rd = csv.reader(f)
        next(rd)
        n = len(cols)
        for r in rd:
            if not r:
                continue
            if len(r) != n:
                bad += 1
                r = (r + [""] * (n - len(r)))[:n]
            rows.append(r)
    log(f"tolerant read: {len(rows)} rows, {bad} with a different number of fields than the header")
    return pd.DataFrame(rows, columns=cols, dtype=str)


def read_report(path):
    """The 10 columns the spec uses, renamed to canonical keys. Looked up by header name; if a header was renamed the
    spec's column letter (position) is used instead."""
    head = pd.read_csv(path, nrows=0, dtype=str, encoding="utf-8-sig")
    cols = list(head.columns)
    use, rename = [], {}
    for key, (name, pos) in REPORT_COLS.items():
        if name in cols:
            use.append(name)
            rename[name] = key
        elif pos < len(cols):
            log(f"WARNING: header '{name}' not found - using column {pos + 1} ('{cols[pos]}') instead")
            use.append(cols[pos])
            rename[cols[pos]] = key
        else:
            raise ValueError(f"{os.path.basename(path)}: column '{name}' (position {pos + 1}) is missing")
    try:
        # index_col=False: the MFNAS csv has a trailing comma on every data row (27 fields vs 26 header names). Without it pandas
        # silently uses the first column as the index and shifts EVERY header one column to the right (Order Date got Product ID).
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False, usecols=use, index_col=False)
    except pd.errors.ParserError as ex:
        log(f"WARNING: malformed rows in the csv ({str(ex)[:110]}) - re-reading tolerantly")
        df = _read_csv_tolerant(path, cols)[use]
    df = df.rename(columns=rename)[list(REPORT_COLS)]
    # sanity check: the first Order Number must look like an order number, not a date / timeslot (catches a header/data shift early)
    first = df["order_no"].replace("", pd.NA).dropna().head(1).tolist()
    if first and re.fullmatch(r"[\d:\- .]+", first[0]):
        raise ValueError(f"{os.path.basename(path)}: header and data rows are misaligned (Order Number = {first[0]!r})")
    return df


_DATE_FORMATS = ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%m-%d-%Y", "%Y%m%d")


def parse_dates(series, expected=None, prefer=None):
    """Date part of an OIX/MFNAS date(-time) string -> (Timestamps, format). The export format is not always m/d/Y, so every
    known format is tried; the winner is the one that gives the most rows equal to `expected` (the date we know the column
    must hold), else the one that parses the most rows (`prefer` wins ties)."""
    head = series.astype(str).str.strip().str.split(r"[ T]", n=1, regex=True).str[0]
    best, best_key, best_fmt = None, None, None
    for fmt in _DATE_FORMATS:
        parsed = pd.to_datetime(head, format=fmt, errors="coerce")
        hit = int((parsed == pd.Timestamp(expected)).sum()) if expected is not None else 0
        key = (hit, int(parsed.notna().sum()), fmt == prefer)
        if best_key is None or key > best_key:
            best, best_key, best_fmt = parsed, key, fmt
    if best_key[1] == 0:
        return pd.Series(pd.NaT, index=series.index), None
    return best, best_fmt


class ParentRegistry:
    """v11.10 - the Parent Orders that earlier records already counted, kept per DELIVERY date and per ORDER-date record:
        {delivery date: {order date: [parent orders first counted by that record]}}
    v11.11 - the Waybill Numbers are tracked the same way (self.wb, same layout).
    A record stores only its NEW parents / waybills, so the union of the records of the order dates BEFORE X is everything already
    counted for that delivery date. The file lives next to the raw csv files (outside the dashboard repo): those csv files are deleted
    once processed, so this is the only memory of what the previous days counted."""

    def __init__(self, path=None):
        self.path = path or REGISTRY_PATH
        self.days = {}
        self.wb = {}                            # v11.11: same layout as self.days, for waybill numbers
        self.quiet = False                      # the backfill re-computes many days in a row and warns once at its end

    def load(self):
        if not os.path.exists(self.path):
            log(f"registry: {self.path} does not exist yet - starting with no history")
            return self
        try:
            with gzip.open(self.path, "rt", encoding="utf-8") as fh:
                raw = json.load(fh)
            self.days, self.wb = raw.get("days", {}), raw.get("wb", {})
            if self.days and not self.wb:
                log("WARNING: the registry has no waybill history (written before v11.11) - run rolling_orders_backfill.py "
                    "--from <first order date> so the waybills of the earlier days are known")
        except Exception as ex:
            raise RuntimeError(f"parent registry {self.path} is unreadable ({type(ex).__name__}: {ex}). Restore it, or delete it and run "
                               f"rolling_orders_backfill.py --from <first order date> to rebuild it - a run without it would count "
                               f"parents again that earlier days already counted.")
        n = sum(len(v) for v in self.days.values())
        log(f"registry: {len(self.days)} delivery dates, {n} (delivery date x order date) records loaded")
        return self

    def order_dates(self):
        return sorted({o for src in (self.days, self.wb) for rec in src.values() for o in rec})

    def seen_before(self, order_date, ddate, kind="parent"):
        """Parents (kind="parent") or waybills (kind="waybill") already counted for delivery date `ddate` by the records of
        order dates BEFORE `order_date`."""
        cut = f"{order_date:%Y-%m-%d}"
        seen = set()
        for od, parents in (self.wb if kind == "waybill" else self.days).get(ddate, {}).items():
            if od < cut:
                seen.update(parents)
        return seen

    def later_dates(self, order_date):
        cut = f"{order_date:%Y-%m-%d}"
        return [o for o in self.order_dates() if o > cut]

    def record(self, order_date, new_by_ddate, new_wb_by_ddate=None):
        """Replace the record of `order_date` (a re-run overwrites it) with its new parents / waybills per delivery date."""
        od = f"{order_date:%Y-%m-%d}"
        for store, new in ((self.days, new_by_ddate), (self.wb, new_wb_by_ddate or {})):
            for rec in store.values():
                rec.pop(od, None)
            for dte, items in new.items():
                store.setdefault(dte, {})[od] = sorted(items)
        self.days = {d: r for d, r in self.days.items() if r}
        self.wb = {d: r for d, r in self.wb.items() if r}

    def save(self, newest_order_date):
        keep_from = f"{newest_order_date - timedelta(days=REGISTRY_KEEP_DAYS):%Y-%m-%d}"
        self.days = {d: r for d, r in self.days.items() if d >= keep_from}      # delivery dates long past are finished
        self.wb = {d: r for d, r in self.wb.items() if d >= keep_from}
        os.makedirs(os.path.dirname(os.path.abspath(self.path)), exist_ok=True)
        tmp = self.path + ".tmp"
        with gzip.open(tmp, "wt", encoding="utf-8") as fh:
            json.dump({"version": 2, "savedAt": datetime.now().strftime("%Y-%m-%dT%H:%M"), "days": self.days, "wb": self.wb}, fh,
                      ensure_ascii=False, separators=(",", ":"))
        os.replace(tmp, self.path)
        log(f"registry: saved {self.path} ({os.path.getsize(self.path) / 1048576:.1f} MB, {len(self.days)} delivery dates)")


def process_report(path, order_date, tableau_path, parent_map=None, registry=None):
    """Returns (rows, stats). rows = DataFrame (one row per report line left after the filters) with
    district, delivery (date), order_no, parent, is_new, waybill, wb_new.
      parent (v11.10) = the Tableau Parent Order of (order number, delivery date); an order that is not in the Tableau report
                        uses its own order number as its parent.
      is_new          = the parent was NOT counted by an earlier order-date record for the same delivery date (registry).
      wb_new (v11.11) = the same test for the Waybill Number (a waybill an earlier record counted for that delivery date is skipped)."""
    df = read_report(path)
    st = {"rawRows": len(df)}
    # --- 1. dates -> MM/DD/YYYY
    od, od_fmt = parse_dates(df["order_date"], expected=order_date)
    dd, dd_fmt = parse_dates(df["delivery"], prefer=od_fmt)
    if od_fmt is None:
        log(f"WARNING: Order Date column cannot be read as dates (e.g. {df['order_date'].replace('', pd.NA).dropna().head(3).tolist()}) "
            f"- the file is taken as the orders of {order_date:%Y-%m-%d}")
    else:
        off = int((od.notna() & (od != pd.Timestamp(order_date))).sum())
        st["orderDates"] = {k: int(v) for k, v in od.dt.strftime("%Y-%m-%d").value_counts().head(4).items()}
        if off:
            log(f"WARNING: {off} rows have an Order Date other than {order_date:%Y-%m-%d}: {st['orderDates']} (kept, as the file is the T-1 list)")
    if dd_fmt is None:
        raise ValueError(f"Delivery Date (column P) cannot be read as dates - sample values "
                         f"{df['delivery'].replace('', pd.NA).dropna().head(5).tolist()}. "
                         f"If the csv was opened and re-saved in Excel its dates are gone; use the file exactly as downloaded from MFNAS.")
    st["dateFormat"] = {"orderDate": od_fmt, "delivery": dd_fmt}
    df["Order Date"] = od.dt.strftime("%m/%d/%Y")            # the converted columns (spec: MM/DD/YYYY)
    df["Delivery Date"] = dd.dt.strftime("%m/%d/%Y")
    df["delivery"] = dd
    bad = int(dd.isna().sum())
    df = df[dd.notna()]
    st["badDeliveryDate"] = bad
    # --- 2. remove: districts outside the 11 codes / CANCEL / zone KN015
    dist = df["district"].str.strip().str.upper()
    df = df[dist.isin(VALID_DISTRICTS)]
    st["afterDistrict"] = len(df)
    df = df[~df["status"].str.upper().str.contains("CANCEL", na=False)]
    st["afterCancel"] = len(df)
    df = df[df["zone"].str.strip().str.upper() != EXCLUDE_ZONE]
    st["afterZone"] = len(df)
    # --- 3. delivery date = order date (T-1): only the 18:00 - 22:00 timeslot is counted
    same = df["delivery"] == pd.Timestamp(order_date)
    slot = df["slot"].str.replace(r"\s+", "", regex=True).str.replace(r"[\u2013\u2014~]", "-", regex=True)   # "18:00 - 22:00" / en dash
    drop_same = same & (slot != SAME_DAY_TIMESLOT)
    st["sameDayRows"] = int(same.sum())
    st["sameDayDropped"] = int(drop_same.sum())
    st["sameDaySlots"] = {k: int(v) for k, v in slot[same].value_counts().head(6).items()}      # what the timeslot values look like
    if st["sameDayRows"] and st["sameDayRows"] == st["sameDayDropped"]:
        log(f"WARNING: none of the {st['sameDayRows']} same-day rows has Delivery Timeslot {SAME_DAY_TIMESLOT}; values seen: {st['sameDaySlots']}")
    df = df[~drop_same].copy().reset_index(drop=True)
    st["rowsKept"] = len(df)
    # --- 4. district (NT-YT -> WTH) and the Parent Order of every order (v11.10)
    d = df["district"].str.strip().str.upper()
    df["district"] = d.replace(DISTRICT_ALIAS)
    df["order_no"] = df["order_no"].str.strip().replace("", pd.NA)
    df["waybill"] = df["waybill"].str.strip().replace("", pd.NA)
    pm = parent_map if parent_map is not None else load_parent_map(tableau_path)
    mapped = pm.parents(df["order_no"], df["delivery"])                    # Tableau rows of the SAME delivery date
    df["parent"] = mapped.where(mapped.notna(), df["order_no"])           # not in the Tableau report -> the order number IS the parent
    pairs = df[df["order_no"].notna()].drop_duplicates(["order_no", "delivery"])      # one line per (order number, delivery date)
    unm = pairs[mapped.loc[pairs.index].isna()]
    if pm.has_delivery and pm.min_delivery is not None:
        before = unm["delivery"] < pm.min_delivery                            # earlier than the report's start date: cannot be in it
        covered_pairs = int((pairs["delivery"] >= pm.min_delivery).sum())
        unm_covered = int((~before).sum())
        if int(before.sum()):
            log(f"WARNING: {int(before.sum())} orders have a delivery date before the first delivery date of the Tableau report "
                f"({pm.min_delivery:%Y-%m-%d}) - not mappable, their own order number is the parent")
    else:
        covered_pairs, unm_covered = len(pairs), len(unm)
    st["ordersInFilteredReport"] = len(pairs)
    st["ordersNotInTableau"] = len(unm)                                  # own order number used as the parent
    st["ordersOtherDeliveryDate"] = int(unm["order_no"].isin(pm.orders).sum()) if pm.has_delivery else 0    # in Tableau, but under another delivery date
    if st["ordersNotInTableau"]:
        pct = 100.0 * st["ordersNotInTableau"] / max(1, st["ordersInFilteredReport"])
        log(f"WARNING: {st['ordersNotInTableau']} of {st['ordersInFilteredReport']} orders ({pct:.2f}%) are not in the Tableau report for their "
            f"delivery date ({st['ordersOtherDeliveryDate']} of them are in it under another delivery date) - their order number is used as the parent")
    # data guard removed: a low Tableau match rate is only logged (the run continues; unmapped orders use their own order number as parent)
    if covered_pairs and unm_covered > MAX_UNMAPPED_SHARE * covered_pairs:
        log(f"WARNING: only {100 - 100.0 * unm_covered / covered_pairs:.0f}% of the orders were found in {os.path.basename(tableau_path)} "
            f"(Tableau newest Order Date {pm.max_order_date:%Y-%m-%d}, {pm.last_day_rows} rows) - continuing anyway (guard removed).")
    # --- 5. (v11.10) a parent already counted by an earlier record for the same delivery date is not counted again
    if registry is not None and not registry.quiet:
        later = registry.later_dates(order_date)
        if later:
            log(f"WARNING: the registry already holds records of later order dates ({later[0]} .. {later[-1]}); they were deduplicated against "
                f"the old history of {order_date:%Y-%m-%d}. Re-run those days in order (rolling_orders_backfill.py --from {order_date:%Y-%m-%d}).")
    df["is_new"] = False
    df["wb_new"] = False
    new_by, n_parents, n_seen = {}, 0, 0
    new_wb, n_wb, n_wb_seen = {}, 0, 0
    ddates = df["delivery"].dt.strftime("%Y-%m-%d")
    for dte, g in df[df["parent"].notna()].groupby(ddates):
        ps = pd.Series(g["parent"].unique())
        seen = registry.seen_before(order_date, dte) if registry is not None else set()
        new = ps[~ps.isin(seen)]
        df.loc[g.index, "is_new"] = g["parent"].isin(set(new)).values
        new_by[dte] = new.tolist()
        n_parents += len(ps)
        n_seen += len(ps) - len(new)
    for dte, g in df[df["waybill"].notna()].groupby(ddates):               # v11.11 - the waybills, same rule
        ws = pd.Series(g["waybill"].unique())
        seen = registry.seen_before(order_date, dte, "waybill") if registry is not None else set()
        new = ws[~ws.isin(seen)]
        df.loc[g.index, "wb_new"] = g["waybill"].isin(set(new)).values
        new_wb[dte] = new.tolist()
        n_wb += len(ws)
        n_wb_seen += len(ws) - len(new)
    st["parentOrders"] = n_parents - n_seen                             # new parents = what this record adds to the dashboard
    st["parentsAlreadyCounted"] = n_seen                                # skipped: counted by an earlier order date, same delivery date
    st["parentRows"] = int(df["is_new"].sum())
    st["waybills"] = n_wb - n_wb_seen                                   # new waybills = what this record adds to the dashboard
    st["waybillsAlreadyCounted"] = n_wb_seen
    if registry is not None:
        registry.record(order_date, new_by, new_wb)
    log(f"rows {st['rawRows']} -> districts {st['afterDistrict']} -> no CANCEL {st['afterCancel']} -> no {EXCLUDE_ZONE} {st['afterZone']} "
        f"-> same-day slot {st['rowsKept']} | parents {n_parents} (new {st['parentOrders']}, already counted {n_seen}), "
        f"waybills {n_wb} (new {st['waybills']}, already counted {n_wb_seen})")
    return df, st


def count_cells(rows):
    """{deliveryDate: {district|'OVERALL': [orders, waybills]}}
    orders = distinct PARENT ORDER among the rows whose parent no earlier order-date record counted for that delivery date (v11.10);
    waybills = distinct Waybill Number that no earlier order-date record counted for that delivery date (v11.11)."""
    def cell(g):
        return [int(g.loc[g["is_new"], "parent"].nunique()), int(g.loc[g["wb_new"], "waybill"].nunique())]
    cells = {}
    for dte, g in rows.groupby(rows["delivery"].dt.strftime("%Y-%m-%d")):
        c = {"OVERALL": cell(g)}
        for dist, gd in g.groupby("district"):
            c[dist] = cell(gd)
        cells[dte] = c
    return cells


def build_snapshot(order_date, report_file, tableau_file, registry=None):
    rows, st = process_report(report_file, order_date, tableau_file, registry=registry)
    return count_cells(rows), {"report": st}


# ----------------------------------------------------------------------------- JSON + embedding
def update_json(order_date, slot, cells, stats):
    data = {}
    if os.path.exists(ROLLING_JSON):
        try:
            data = json.load(open(ROLLING_JSON, encoding="utf-8"))
        except Exception:
            data = {}
    data.setdefault("days", {})
    data.setdefault("snapshots", {})
    ods = order_date.strftime("%Y-%m-%d")
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    data["days"][ods] = {"slot": slot, "asOf": now, "cells": cells}
    data["snapshots"].setdefault(ods, {})[slot] = {"asOf": now, "cells": cells}
    keep = sorted(data["days"])[-KEEP_ORDER_DATES:]
    data["days"] = {k: data["days"][k] for k in keep}
    keep_snap = set(sorted(data["days"])[-KEEP_SNAPSHOT_DATES:])
    data["snapshots"] = {k: v for k, v in data["snapshots"].items() if k in keep_snap}
    data["districts"] = DISTRICTS
    data["generatedAt"] = now
    data["lastRun"] = {"orderDate": ods, "slot": slot, "source": "ECOM-MFNAS", "stats": stats}
    tmp = ROLLING_JSON + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))     # compact: it is pushed + embedded
    os.replace(tmp, ROLLING_JSON)
    log(f"wrote {ROLLING_JSON}")
    return data


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
    ap.add_argument("--slot", help="kept so an old cron line does not break; only 06:00 runs (Version 11.1)")
    ap.add_argument("--date", help="T-0 = the run date YYYY-MM-DD (default today). The report holds the orders of the day BEFORE it")
    ap.add_argument("--skip-download", action="store_true", help="use the file already in Folder for Rolling Orders")
    ap.add_argument("--process-only", action="store_true", help="skip download + git push; uses --report-file or the file in the folder")
    ap.add_argument("--report-file", help="process this csv instead of Rolling Order Count (T-1).csv")
    ap.add_argument("--skip-tableau-download", action="store_true", help="use the Tableau csv already in Folder for Rolling Orders")
    ap.add_argument("--tableau-file", help="process with this Tableau csv instead of AI Fetching for Order Date T-1 (T-1).csv")
    ap.add_argument("--headed", action="store_true", help="show the Chrome window during the Tableau download (same as ROLLING_TABLEAU_HEADLESS=0)")
    ap.add_argument("--tableau-only", action="store_true", help="only (re)download the Tableau csv, then stop")
    ap.add_argument("--download-only", action="store_true", help="only download the MFNAS csv + Tableau csv into Folder for Rolling Orders, then stop")
    ap.add_argument("--skip-whatsapp", action="store_true", help="ignored - WhatsApp sending was removed; kept so old command lines still work")
    ap.add_argument("--push-only", action="store_true", help="no download / processing: embed the rolling_orders.json that is already in the public folder into index.html and git-push it")
    ap.add_argument("--keep-files", action="store_true", help="do not delete the downloaded csv files after they were written to the json (same as ROLLING_KEEP_FILES=1)")
    ap.add_argument("--skip-embed", action="store_true")
    ap.add_argument("--skip-push", action="store_true", help="do not git-push the dashboard files (same as ROLLING_DEPLOY=none)")
    a = ap.parse_args()
    global TABLEAU_HEADLESS
    if a.headed:
        TABLEAU_HEADLESS = False

    slot = SLOTS[0]
    if a.push_only:
        with run_lock():
            data = json.load(open(ROLLING_JSON, encoding="utf-8"))
            last = max(data.get("days", {}) or [datetime.now().strftime("%Y-%m-%d")])
            log(f"push-only: {ROLLING_JSON} holds {len(data.get('days', {}))} order dates, latest {last}, generatedAt {data.get('generatedAt')}")
            if not a.skip_embed:
                embed_in_index(data)
            if not a.skip_push:
                log("push-only: result = " + ("PUSHED" if push_dashboard(datetime.strptime(last, "%Y-%m-%d"), slot, data) else "NOT pushed (see the push: lines above)"))
        return
    if a.slot and a.slot not in SLOTS:
        log(f"slot {a.slot}: nothing to do - since Version 11.1 Rolling Orders runs once a day at {slot} only "
            f"(remove this cron / Task Scheduler entry)")
        return
    t0 = datetime.strptime(a.date, "%Y-%m-%d") if a.date else run_dates()[0]
    order_date = t0 - timedelta(days=1)
    log(f"T-0={t0:%Y-%m-%d}  ->  Order Date (T-1)={order_date:%Y-%m-%d}")
    with run_lock():
        report = a.report_file or report_path(order_date)
        tableau = a.tableau_file or tableau_report_path(order_date)
        if a.tableau_only:
            download_tableau_report(order_date, tableau)
            log("tableau-only: done")
            return
        if not (a.skip_download or a.process_only or a.report_file):
            download_report(t0, report)
        # v11.3: right after the MFNAS csv, fetch the Tableau parent-order report (start date = T-1, 5 min refresh wait)
        if not (a.skip_tableau_download or a.process_only or a.tableau_file):
            download_tableau_report(order_date, tableau)
        if a.download_only:
            log("download-only: done")
            return
        log(f"report file: {report}  ->  {'FOUND' if os.path.exists(report) else 'NOT FOUND'}")
        if not os.path.exists(report):
            raise SystemExit("no report file - run without --skip-download / --process-only, or pass --report-file")
        log(f"tableau file: {tableau}  ->  {'FOUND' if os.path.exists(tableau) else 'NOT FOUND'}")
        if not os.path.exists(tableau):
            raise SystemExit("no Tableau file - run without --skip-tableau-download, or pass --tableau-file")
        registry = ParentRegistry().load()                  # v11.10: parents the earlier order dates already counted
        if not registry.order_dates() and os.path.exists(ROLLING_JSON):
            log("WARNING: the parent registry is empty but rolling_orders.json already holds days - run rolling_orders_backfill.py "
                "--from <first order date> once so the earlier days are known; otherwise parents they counted are counted again")
        cells, stats = build_snapshot(order_date, report, tableau, registry)
        log("stats: " + json.dumps(stats, ensure_ascii=False))
        if not cells:
            raise SystemExit(f"report found but 0 usable rows for order date {order_date:%Y-%m-%d} - nothing written")
        data = update_json(order_date, slot, cells, stats)
        registry.save(order_date)                          # only after the json was written, so the two never disagree
        # v11.9 - the numbers are in the json now: delete the two downloaded csv files (not with --keep-files / --process-only,
        # and never a file given with --report-file / --tableau-file)
        if not (a.keep_files or KEEP_FILES or a.process_only):
            cleanup_files([None if a.report_file else report, None if a.tableau_file else tableau])
        if not a.skip_embed:
            embed_in_index(data)
        if not (a.skip_push or a.process_only):
            push_dashboard(order_date, slot, data)       # publish the update to the Rolling Orders tab


if __name__ == "__main__":
    sys.path.insert(0, HERE)
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
