#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_orders.py  -  KPI Dashboard Version 11.2  "Rolling Orders"   (raw data now comes from MFNAS, not OIX)

One run per day, at 06:00 (T-0). It reports the orders created YESTERDAY (T-1):
  1. Logs in to TKO-MFNAS (Synology File Station API, CRAB_NAS_USER / CRAB_NAS_PASS - same as project_crab.py),
     opens  ECOM-MFNAS / ECOM-MFNAS_hktvAccumulatedDeliveryLeadTimeReport / <T-0 as YYYYMMDD>/  and downloads its one csv.
     The csv holds the full list of orders with Order Date = T-1, so it is saved to "Folder for Rolling Orders" as
         Rolling Order Count (MMM-DD of T-1).csv        (same name every run => overwritten)
  2. Processes it (column letters as in the file):
       - Order Date (C) and Delivery Date (P) are converted to MM/DD/YYYY
       - rows are removed when  Delivery District (Z) is not one of ETH ETK ETX NT-ST NT-TM NT-TSM NT-TW NT-YT WTH WTX WTK,
         or  Consignment Status (J) contains CANCEL,  or  Delivery Zone (U) = KN015
       - delivery date = T-1 (same as the order date): only the 18:00 - 22:00 Delivery Timeslot (Q) is counted
       - ORDER count = distinct Order Number among the rows where BOTH Optimum Delivery Timeslot (N) and
         First Available Delivery Timeslot (O) have a value (= parent orders, de-duplicated)
         WAYBILL count = distinct Waybill Number over all remaining rows (rows without both N and O only count here)
       - counted per (Delivery Date x district); NT-YT is grouped into WTH
  3. Writes rolling_orders.json next to index.html AND embeds it in index.html's window.__EMBEDDED_DATA__ block.
  4. Pushes rolling_orders.json + index.html to the dashboard repo (git pull --rebase -> commit -> push), with an
     automatic fix for the one conflict this always risks: the EMBEDDED_DATA block of index.html, which
     kpi_pipeline.py rewrites too (see push_dashboard). Disable with --skip-push or ROLLING_DEPLOY=none.
  5. Takes a screenshot of the "Rolling Orders" tab table and sends it + the message to the
     WhatsApp group via wa_common (shared WhatsApp Web session).

Usage (WSL):
    python3 rolling_orders.py                          # the 06:00 cron call (T-0 = today, order date = yesterday)
    python3 rolling_orders.py --date 2026-10-14        # re-run for another T-0 (reports the orders of 2026-10-13)
    python3 rolling_orders.py --skip-whatsapp --skip-push
    python3 rolling_orders.py --process-only --report-file "Rolling Order Count (Oct-13).csv" --date 2026-10-14

Cron (WSL, machine timezone = HKT) - ONE line; delete the old 13:00 / 18:00 / 00:00 lines:
    0 6 * * *  cd "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker" && python3 rolling_orders.py >> logs/rolling_0600.log 2>&1

Credentials (never hard-code them): environment variables CRAB_NAS_USER / CRAB_NAS_PASS, or a KEY=VALUE file
("rolling_orders.env" or "crab.env", next to this script or one folder up).
"""
import argparse
import contextlib
import fcntl
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
sys.path.insert(0, os.path.dirname(HERE))   # wa_common.py may sit one level up (…/Whatsapp Session/)
sys.path.insert(0, os.environ.get("WA_COMMON_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session"))   # same default as project_crab.py


# ----------------------------------------------------------------------------- config
def _load_env_file():
    for folder in (HERE, os.path.dirname(HERE)):
        for name in ("rolling_orders.env", "crab.env"):
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
SHOT_DIR = os.environ.get("ROLLING_SHOT_DIR", ROLLING_DIR)

WA_TARGET_GROUP = os.environ.get("WA_TARGET_GROUP", "LOG 區頭 x Head office")
# WA_SESSION_DIR is read by wa_common itself (default /home/chipanl/whatsapp_session_2)

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


def process_report(path, order_date):
    """Returns (rows, stats). rows = DataFrame (one row per report line left after the filters) with
    district, delivery (date), is_parent, order_no, waybill."""
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
    df = df[~drop_same].copy()
    st["rowsKept"] = len(df)
    # --- 4. parent order = both N and O have a value; district (NT-YT -> WTH)
    df["is_parent"] = (df["optimum"].str.strip() != "") & (df["first"].str.strip() != "")
    d = df["district"].str.strip().str.upper()
    df["district"] = d.replace(DISTRICT_ALIAS)
    df["order_no"] = df["order_no"].str.strip().replace("", pd.NA)
    df["waybill"] = df["waybill"].str.strip().replace("", pd.NA)
    st["parentRows"] = int(df["is_parent"].sum())
    st["parentOrders"] = int(df.loc[df["is_parent"], "order_no"].nunique())
    st["waybills"] = int(df["waybill"].nunique())
    log(f"rows {st['rawRows']} -> districts {st['afterDistrict']} -> no CANCEL {st['afterCancel']} -> no {EXCLUDE_ZONE} {st['afterZone']} "
        f"-> same-day slot {st['rowsKept']} | parent orders {st['parentOrders']}, waybills {st['waybills']}")
    return df, st


def count_cells(rows):
    """{deliveryDate: {district|'OVERALL': [orders, waybills]}}
    orders = distinct Order Number among the parent rows (both N and O filled); waybills = distinct Waybill Number of all rows."""
    def cell(g):
        return [int(g.loc[g["is_parent"], "order_no"].nunique()), int(g["waybill"].nunique())]
    cells = {}
    for dte, g in rows.groupby(rows["delivery"].dt.strftime("%Y-%m-%d")):
        c = {"OVERALL": cell(g)}
        for dist, gd in g.groupby("district"):
            c[dist] = cell(gd)
        cells[dte] = c
    return cells


def build_snapshot(order_date, report_file):
    rows, st = process_report(report_file, order_date)
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


# ----------------------------------------------------------------------------- screenshot + WhatsApp
def screenshot_table(order_date, slot):
    from playwright.sync_api import sync_playwright
    os.makedirs(SHOT_DIR, exist_ok=True)
    out = os.path.join(SHOT_DIR, f"Rolling Order Trend ({mmm_dd(order_date)})_{slot.replace(':', '')}.png")
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        try:
            pg = b.new_page(viewport={"width": 1700, "height": 1100}, device_scale_factor=1.5)
            pg.add_init_script("window.__ROLLING_SHOT__ = true;")     # dashboard JS: latest 14 order dates, nothing collapsed, no scroll boxes
            pg.goto("file://" + INDEX_HTML, wait_until="domcontentloaded", timeout=120000)
            pg.wait_for_selector(".main-tab[data-target='view-rolling']", timeout=120000)
            pg.click(".main-tab[data-target='view-rolling']")
            pg.wait_for_selector("#rollingTableWrap table", timeout=60000)
            pg.add_style_tag(content='#rollingShot .tbl-dl-bar{display:none !important}')
            time.sleep(1)
            w = pg.evaluate("() => Math.ceil(document.querySelector('#rollingShot').getBoundingClientRect().width)")
            if w + 140 > 1700:      # wide table: widen the page so the Total column is never clipped
                pg.set_viewport_size({"width": w + 140, "height": 1100})
                time.sleep(1)
            pg.locator("#rollingShot").screenshot(path=out)
        finally:
            b.close()
    log(f"screenshot {out}")
    return out


def whatsapp_text(order_date, slot):
    return (f"Rolling Order Trend ({mmm_dd(order_date)})\n\n"
            f"Attached please find the rolling orders trend as of {mmm_dd(order_date)} Overall above.")


def send_whatsapp(image_path, text):
    """Send the trend screenshot as a PHOTO (with the message as caption).
    Attach logic ported from project_crab.py: the plain `input[accept*="image"]` (image/* only) is WhatsApp's
    STICKER picker, which is what made the screenshot arrive as a sticker. The real Photos & videos
    input also accepts video, so it is chosen via the menu item / file chooser, or by accept=image+video."""
    import wa_common
    with wa_common.wa_session() as (context, page):
        search_sel = 'div[contenteditable="true"][data-tab="3"], input[data-tab="3"]'
        page.wait_for_selector(search_sel, timeout=60000)
        time.sleep(3)
        sb = page.locator(search_sel).first
        sb.fill("")
        sb.type(WA_TARGET_GROUP, delay=80)
        time.sleep(2)
        gsel = f'span[title="{WA_TARGET_GROUP}"]'
        page.wait_for_selector(gsel, timeout=10000)
        page.click(gsel)
        time.sleep(3)
        msg_sel = 'div[contenteditable="true"][data-tab="10"]'
        page.wait_for_selector(msg_sel, timeout=10000)

        attach_selectors = [
            'div[aria-label="Attach"]', 'button[aria-label="Attach"]',
            'button[title="Attach"]', 'div[title="Attach"]',
            '[aria-label="附加"]', '[aria-label="附件"]',
            'span[data-icon="plus"]', 'span[data-icon="plus-rounded"]',
            'span[data-icon="attach-menu-plus"]', 'span[data-icon="attach-menu-item"]',
            'span[data-icon="clip"]',
        ]
        media_selectors = [
            'span[data-icon="attach-image"]', 'button[aria-label="Photos & videos"]',
            'button[aria-label="相片和影片"]', 'button[aria-label="照片和影片"]',
            'button[aria-label="相片及影片"]', 'li[data-testid="mi-attach-media"]',
            'span:has-text("Photos & videos")', 'span:has-text("相片和影片")',
            'span:has-text("照片和影片")', 'span:has-text("相片及影片")',
        ]
        caption_selectors = [
            'div[aria-label*="aption"]', 'div[aria-label*="說明"]', 'div[aria-label*="標題"]',
        ]
        send_selectors = [
            'span[data-icon="send"]', 'span[data-icon="wds-ic-send-filled"]',
            'div[aria-label="Send"]', 'div[aria-label="發送"]', 'div[aria-label="傳送"]',
            'button[aria-label="Send"]', 'button[aria-label="發送"]', 'button[aria-label="傳送"]',
        ]

        def open_attach_menu():
            for sel in attach_selectors:
                try:
                    btn = page.locator(sel).first
                    if btn.is_visible():
                        btn.click(timeout=3000)
                        time.sleep(1)
                        log(f"WhatsApp: attach menu opened via {sel}")
                        return True
                except Exception:
                    continue
            log("WhatsApp: attach button not found by selector; relying on hidden file input")
            return False

        def choose_via_menu(path, selectors):
            """Click the 'Photos & videos' entry and hand the file to the native file chooser."""
            open_attach_menu()
            for sel in selectors:
                try:
                    el = page.locator(sel).first
                    if el.is_visible():
                        with page.expect_file_chooser(timeout=5000) as fc:
                            el.click()
                        fc.value.set_files(path)
                        log(f"WhatsApp: photo chosen via menu item {sel}")
                        page.wait_for_timeout(3000)   # preview dialog
                        return True
                except Exception:
                    continue
            return False

        def send_photo(path):
            if choose_via_menu(path, media_selectors):
                return
            accepts = page.evaluate(
                "() => Array.from(document.querySelectorAll('input[type=file]'))"
                ".map(i => i.getAttribute('accept') || '')")
            log(f"WhatsApp: Photos menu not found; hidden inputs (accept): {accepts}")
            for i, acc in enumerate(accepts):
                # the real Photos & videos input also accepts video; image/* alone = the STICKER picker
                if "image" in acc and "video" in acc:
                    page.locator('input[type="file"]').nth(i).set_input_files(path)
                    page.wait_for_timeout(3000)
                    return
            # last resort: put the PNG on the clipboard and paste it into the composer (opens a photo preview)
            log("WhatsApp: falling back to clipboard paste for the image")
            import base64
            try:
                context.grant_permissions(["clipboard-read", "clipboard-write"], origin="https://web.whatsapp.com")
            except Exception:
                pass
            with open(path, "rb") as fh:
                b64 = base64.b64encode(fh.read()).decode()
            page.evaluate("""async (b64) => {
                const bin = atob(b64); const arr = new Uint8Array(bin.length);
                for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
                const blob = new Blob([arr], {type: 'image/png'});
                await navigator.clipboard.write([new ClipboardItem({'image/png': blob})]);
            }""", b64)
            page.locator(msg_sel).click()
            page.keyboard.press("Control+v")
            page.wait_for_timeout(3000)

        def click_send():
            for sel in send_selectors:
                try:
                    btn = page.locator(sel).last
                    if btn.is_visible(timeout=3000):
                        btn.click()
                        log(f"WhatsApp: send clicked via {sel}")
                        return
                except Exception:
                    continue
            log("WhatsApp: send button not found; pressing Enter")
            page.keyboard.press("Enter")

        def paste_text(txt):
            page.evaluate("(t) => navigator.clipboard.writeText(t)", txt)
            page.keyboard.press("Control+v")
            page.wait_for_timeout(800)

        send_photo(image_path)
        caption_ok = False
        for sel in caption_selectors:
            try:
                box = page.locator(sel).first
                if box.is_visible():
                    box.click(force=True, timeout=5000)
                    paste_text(text)
                    caption_ok = True
                    log(f"WhatsApp: caption filled via {sel}")
                    break
            except Exception:
                continue
        if not caption_ok:
            log("WhatsApp: caption box not found; the message will be sent as a separate text")
        click_send()
        page.wait_for_timeout(5000)
        if not caption_ok:
            page.wait_for_selector(msg_sel, timeout=15000)
            page.locator(msg_sel).click()
            paste_text(text)
            page.keyboard.press("Enter")
            page.wait_for_timeout(3000)
        log("WhatsApp: picture + message sent")


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
    ap.add_argument("--process-only", action="store_true", help="skip download + git push + WhatsApp; uses --report-file or the file in the folder")
    ap.add_argument("--report-file", help="process this csv instead of Rolling Order Count (T-1).csv")
    ap.add_argument("--download-only", action="store_true", help="only download the MFNAS csv into Folder for Rolling Orders, then stop")
    ap.add_argument("--skip-whatsapp", action="store_true")
    ap.add_argument("--skip-embed", action="store_true")
    ap.add_argument("--skip-push", action="store_true", help="do not git-push the dashboard files (same as ROLLING_DEPLOY=none)")
    a = ap.parse_args()

    slot = SLOTS[0]
    if a.slot and a.slot not in SLOTS:
        log(f"slot {a.slot}: nothing to do - since Version 11.1 Rolling Orders runs once a day at {slot} only "
            f"(remove this cron / Task Scheduler entry)")
        return
    t0 = datetime.strptime(a.date, "%Y-%m-%d") if a.date else run_dates()[0]
    order_date = t0 - timedelta(days=1)
    log(f"T-0={t0:%Y-%m-%d}  ->  Order Date (T-1)={order_date:%Y-%m-%d}")
    with run_lock():
        report = a.report_file or report_path(order_date)
        if not (a.skip_download or a.process_only or a.report_file):
            download_report(t0, report)
        if a.download_only:
            log("download-only: done")
            return
        log(f"report file: {report}  ->  {'FOUND' if os.path.exists(report) else 'NOT FOUND'}")
        if not os.path.exists(report):
            raise SystemExit("no report file - run without --skip-download / --process-only, or pass --report-file")
        cells, stats = build_snapshot(order_date, report)
        log("stats: " + json.dumps(stats, ensure_ascii=False))
        if not cells:
            raise SystemExit(f"report found but 0 usable rows for order date {order_date:%Y-%m-%d} - nothing written")
        data = update_json(order_date, slot, cells, stats)
        if not a.skip_embed:
            embed_in_index(data)
        if not (a.skip_push or a.process_only):
            push_dashboard(order_date, slot, data)       # publish the update to the Rolling Orders tab
        if not (a.skip_whatsapp or a.process_only):
            img = screenshot_table(order_date, slot)
            send_whatsapp(img, whatsapp_text(order_date, slot))


if __name__ == "__main__":
    sys.path.insert(0, HERE)
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
