#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_orders.py  -  KPI Dashboard Version 11.0  "Rolling Orders"

What it does (one run = one time slot):
  1. Logs in to OIX and downloads the Waybill History Status Check report
       - "next week" file : 送貨日期 = T+1 .. T+7           -> "Rolling Order Count (MMM-DD)_next week.csv"
       - "same day"  file : 送貨日期 = T+0 .. T+0           -> "Rolling Order Count (MMM-DD)_same day.csv"
         (only in the 13:00 and 00:00 sessions - see SAME_DAY_SLOTS)
     and saves them to "Folder for Rolling Orders" (same name every run of the day => overwritten,
     so only 2 reports per day are ever kept).
  2. Processes the reports (order-date = T filter, cancelled waybills removed, parent-order grouping,
     waybill de-duplication, district mapping) and counts parent orders / waybills per
     (Expected Delivery Date x district).
  3. Writes rolling_orders.json next to index.html AND embeds it in index.html's
     window.__EMBEDDED_DATA__ block (so the page also works when opened from disk).
  4. Takes a screenshot of the "Rolling Orders" tab table and sends it + the message to the
     WhatsApp group via wa_common (shared WhatsApp Web session).

Usage (WSL):
    python3 rolling_orders.py --slot 06:00            # normal cron call
    python3 rolling_orders.py --slot 00:00            # midnight call: T = YESTERDAY (the session that just ended)
    python3 rolling_orders.py --slot 13:00 --skip-whatsapp
    python3 rolling_orders.py --process-only --next-week-file X.csv [--same-day-file Y.csv] --date 2026-10-14 --slot 13:00 --skip-whatsapp

Credentials (never hard-code them): environment variables OIX_USERNAME / OIX_PASSWORD, or a
"rolling_orders.env" file (KEY=VALUE lines) next to this script.
"""
import argparse
import contextlib
import fcntl
import json
import os
import re
import shutil
import sys
import threading
import time
import traceback
from collections import Counter, defaultdict
from datetime import datetime, timedelta

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.dirname(HERE))   # wa_common.py may sit one level up (…/Whatsapp Session/)


# ----------------------------------------------------------------------------- config
def _load_env_file():
    p = os.path.join(HERE, "rolling_orders.env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


_load_env_file()

OIX_LOGIN_URL = "https://oix.hktvmall.com/hktv-oix/login"
OIX_REPORT_DOWNLOAD_LINK = "https://oix.hktvmall.com/hktv-oix/internal/report/waybill_status_history_report"
OIX_USERNAME = os.environ.get("OIX_USERNAME", "")
OIX_PASSWORD = os.environ.get("OIX_PASSWORD", "")
OIX_HEADLESS = os.environ.get("OIX_HEADLESS", "1") not in ("0", "false", "False")   # --show-browser sets it to 0
OIX_DOWNLOAD_TIMEOUT_MS = int(os.environ.get("OIX_DOWNLOAD_TIMEOUT_MS", "900000"))   # 15 min, same as digimobi script

# WSL paths (C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\...)
KPI_DIR = os.environ.get("KPI_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker")
# The OIX rolling-order CSVs are too large to push to the dashboard, so they live OUTSIDE log-kpi-tracker, one level up:
# C:\Users\chipanl\Downloads\Whatsapp Session\Folder for Rolling Orders
ROLLING_DIR = os.environ.get("ROLLING_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/Folder for Rolling Orders")
PUBLIC_DIR = os.environ.get("KPI_PUBLIC_DIR", os.path.join(KPI_DIR, "public"))   # same folder kpi_pipeline.py serves from
INDEX_HTML = os.environ.get("KPI_INDEX_HTML", os.path.join(PUBLIC_DIR, "index.html"))
ROLLING_JSON = os.environ.get("ROLLING_JSON", os.path.join(PUBLIC_DIR, "rolling_orders.json"))   # next to index.html
LOG_DIR = os.environ.get("ROLLING_LOG_DIR", os.path.join(KPI_DIR, "logs"))
SHOT_DIR = os.environ.get("ROLLING_SHOT_DIR", ROLLING_DIR)

WA_TARGET_GROUP = os.environ.get("WA_TARGET_GROUP", "LOG 區頭 x Head office")
# WA_SESSION_DIR is read by wa_common itself (default /home/chipanl/whatsapp_session_2)

SLOTS = ["06:00", "13:00", "18:00", "00:00"]
SAME_DAY_SLOTS = ["13:00", "00:00"]          # sessions that also download the T-0 ("same day") report
STATUS_BY_SLOT = {"00:00": "RECEIVED_BY_CUSTOMER"}   # 更新狀態 - blank for every other session
STATUS_APPLIES_TO = {"same_day"}             # which report gets the 更新狀態 filter in the 00:00 session
CLIENT_IDS = ["#clientId_1", "#clientId_3", "#clientId_5", "#clientId_6", "#clientId_99999"]  # hktv, offbeat, wm_express, qcs, csec
KEEP_ORDER_DATES = 60                        # how many order dates are kept in rolling_orders.json

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]
# Delivery-map groups NT-YT together with WTH (index.html DD_GROUPS) -> the dashboard has 10 districts
DISTRICT_ALIAS = {"NT-YT": "WTH"}
# Car-line (送貨車號) -> district: EXACTLY kpi_pipeline.py's district_from_carline() (= Domo_Jai.py's
# parse_district_from_carline(), the Excel IFS mapping). First match wins, so the order matters.
def district_from_carline(carline):
    if carline is None or (isinstance(carline, float) and pd.isna(carline)):
        return None
    raw = str(carline).strip()
    up = raw.upper()
    if not raw:
        return None
    if "ETK" in up or "將軍澳" in raw:
        return "ETK"
    if "WH" in up or "CH" in up or "CHA" in up:
        return "WTH"
    if "CTK" in up or "WK" in up:
        return "WTK"
    if "WX" in up:
        return "WTX"
    if "CTW" in up:
        return "NT-TW"
    if "ENH" in up:
        return "ETH"
    if "CTX" in up or "CX" in up or "KX" in up or "KTX" in up:
        return "ETX"
    if "NST" in up:
        return "NT-ST"
    if "NTM" in up:
        return "NT-TM"
    if "ZTS" in up:
        return "NT-TSM"
    if "WTW" in up:
        return "NT-TW"
    return None


OTHER = "Other"      # waybills that cannot be mapped to one of the 10 districts (Macau, ShipEazy, ...)


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


# ----------------------------------------------------------------------------- dates / names
def session_dates(slot, now=None):
    """T-0 of the session. The 00:00 session belongs to the day that just ended."""
    now = now or datetime.now()
    t0 = now.date()
    if slot == "00:00":
        t0 = t0 - timedelta(days=1)
    return datetime(t0.year, t0.month, t0.day)


def mmm_dd(t0):
    return t0.strftime("%b-%d")                       # Oct-13


def report_paths(t0):
    base = f"Rolling Order Count ({mmm_dd(t0)})"
    return (os.path.join(ROLLING_DIR, base + "_next week.csv"),
            os.path.join(ROLLING_DIR, base + "_same day.csv"))


def infer_slot(now=None):
    now = now or datetime.now()
    mins = now.hour * 60 + now.minute
    best = min(SLOTS, key=lambda s: min(abs(mins - int(s[:2]) * 60), 1440 - abs(mins - int(s[:2]) * 60)))
    return best


# ----------------------------------------------------------------------------- OIX download
def download_oix_reports(t0, slot):
    """Single browser / single login, one or two downloads."""
    from playwright.sync_api import sync_playwright
    if not OIX_USERNAME or not OIX_PASSWORD:
        raise RuntimeError("OIX_USERNAME / OIX_PASSWORD not set (env var or rolling_orders.env)")
    os.makedirs(ROLLING_DIR, exist_ok=True)
    next_path, same_path = report_paths(t0)
    jobs = [("next_week", t0 + timedelta(days=1), t0 + timedelta(days=7), next_path)]
    if slot in SAME_DAY_SLOTS:
        jobs.append(("same_day", t0, t0, same_path))
    status_val = STATUS_BY_SLOT.get(slot, "")
    got = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=OIX_HEADLESS, slow_mo=(0 if OIX_HEADLESS else 400))
        try:
            context = browser.new_context(accept_downloads=True)
            page = context.new_page()
            page.on("dialog", lambda d: d.accept())
            log("OIX: login")
            page.goto(OIX_LOGIN_URL)
            page.fill("#username", OIX_USERNAME)
            page.fill("#password", OIX_PASSWORD)
            page.click("form > div.login-button > input")
            page.wait_for_selector("nav > div > ul", timeout=30000)
            for name, d_from, d_to, final_path in jobs:
                st = status_val if name in STATUS_APPLIES_TO else ""
                _download_one(page, name, d_from, d_to, st, final_path)
                got[name] = final_path
        finally:
            try:
                browser.close()
            except Exception:
                pass
    return got


_SET_DELIVERY_DATES_JS = """({sVal, eVal}) => {
    const form = document.getElementById('generateWaybillStatusHistoryReportForm');
    // 1) the two 送貨日期 inputs: #displayDeliveryDateFrom / #displayDeliveryDateTo (label / first-row search is only a fallback)
    let inputs = [];
    const fromEl = document.getElementById('displayDeliveryDateFrom'), toEl = document.getElementById('displayDeliveryDateTo');
    if (fromEl && toEl) inputs = [fromEl, toEl];          // known ids of the 送貨日期 inputs
    const labels = Array.from(form.querySelectorAll('label, span, div, td, th'))
        .filter(e => e.children.length < 3 && /送貨日期/.test((e.textContent||'').trim()) && (e.textContent||'').trim().length < 20);
    for (const lb of (inputs.length >= 2 ? [] : labels)) {
        let node = lb;
        for (let i = 0; i < 4 && node && inputs.length < 2; i++) {
            node = node.parentElement;
            if (!node) break;
            inputs = Array.from(node.querySelectorAll('input[type=text], input:not([type])')).filter(x => x.type !== 'hidden' && x.type !== 'checkbox');
        }
        if (inputs.length >= 2) break;
    }
    // 2) fallback: first two text inputs of the first row of the form (the recorder clicks the calendar buttons there)
    if (inputs.length < 2) {
        const row = form.querySelector(':scope > div:nth-of-type(1)');
        if (row) inputs = Array.from(row.querySelectorAll('input')).filter(x => x.type !== 'hidden' && x.type !== 'checkbox');
    }
    if (inputs.length < 2) return {ok:false, ids: inputs.map(i=>i.id)};
    const pair = [inputs[0], inputs[1]];
    const vals = [sVal, eVal];
    pair.forEach((el, idx) => {
        el.removeAttribute('readonly');
        try { if (window.$ && $(el).datepicker) $(el).datepicker('setDate', vals[idx]); } catch(e) {}
        el.value = vals[idx];
        ['input','change','blur'].forEach(ev => el.dispatchEvent(new Event(ev, {bubbles:true})));
    });
    return {ok:true, ids: pair.map(i=>i.id), values: pair.map(i=>i.value)};
}"""


def _download_one(page, name, d_from, d_to, status, final_path):
    from playwright.sync_api import TimeoutError as PWTimeout
    log(f"OIX: {name}  送貨日期 {d_from:%Y%m%d} -> {d_to:%Y%m%d}  更新狀態={status or '(blank)'}")
    holder = {}
    on_dl = lambda d: holder.setdefault("d", d)
    page.on("download", on_dl)
    log(f"OIX: {name}: will be saved to {final_path}")
    page.goto(OIX_REPORT_DOWNLOAD_LINK)
    page.wait_for_selector("#generateWaybillStatusHistoryReportForm", timeout=30000)
    page.wait_for_selector("#displayDeliveryDateFrom", timeout=15000)
    time.sleep(1)
    res = page.evaluate(_SET_DELIVERY_DATES_JS, {"sVal": d_from.strftime("%Y%m%d"), "eVal": d_to.strftime("%Y%m%d")})
    if not res.get("ok"):
        raise RuntimeError(f"cannot locate the 送貨日期 inputs on the report form: {res}")
    log(f"OIX: delivery-date inputs {res.get('ids')} = {res.get('values')}")
    page.keyboard.press("Escape")
    time.sleep(1)

    missed = []
    for cid in CLIENT_IDS:
        cb = page.locator(cid)
        try:
            cb.wait_for(state="visible", timeout=10000)
            if not cb.is_checked():
                cb.check()
            if not cb.is_checked():
                missed.append(cid)
        except Exception:
            missed.append(cid)
    if missed:
        raise RuntimeError(f"client checkboxes not ticked: {missed}")

    if status:   # 更新狀態 select2 (same way as digimobi_alert_whatsapp.py)
        page.click("[aria-labelledby*='updateStatusSelector']")
        page.click(f"li.select2-results__option:has-text('{status}')")

    log(f"OIX: {name}: submitting the form - OIX now builds the report (can take several minutes) ...")
    page.click("#generateWaybillStatusHistoryReportForm > button", force=True, no_wait_after=True,
               timeout=OIX_DOWNLOAD_TIMEOUT_MS)
    waited = 0
    while "d" not in holder:
        if waited >= OIX_DOWNLOAD_TIMEOUT_MS / 1000:
            raise PWTimeout(f"download did not start within {OIX_DOWNLOAD_TIMEOUT_MS // 1000}s")
        # NOTE: must be page.wait_for_timeout, NOT time.sleep. Playwright's sync API only delivers
        # events (incl. the "download" event that fills `holder`) while the main thread is inside a
        # Playwright call; time.sleep() never pumps them, so the finished download was never seen.
        page.wait_for_timeout(2000)
        waited += 2
        if waited % 20 == 0:
            log(f"OIX: {name}: still waiting for the download to start ({waited}s)")
    dl = holder["d"]
    log(f"OIX: {name}: download started ({dl.suggested_filename}) - receiving the file ...")
    stop = threading.Event()

    def _beat():                      # heartbeat so a big file never looks like a hang
        t = 0
        while not stop.wait(10):
            t += 10
            log(f"OIX: {name}: still receiving the file ({t}s) ...")
    hb = threading.Thread(target=_beat, daemon=True)
    hb.start()
    tmp = final_path + ".tmp"
    try:
        dl.save_as(tmp)               # returns when the whole file has arrived
    finally:
        stop.set()
    err = dl.failure()
    if err:
        raise RuntimeError(f"download failed: {err}")
    os.replace(tmp, final_path)          # same file name each run -> previous report is covered
    log(f"OIX: {name}: SAVED -> {final_path} ({os.path.getsize(final_path) / 1048576:.1f} MB)")
    page.remove_listener("download", on_dl)


# ----------------------------------------------------------------------------- district mapping
def load_estate_map():
    """Estate Code -> district, taken from the Delivery Map data embedded in index.html
    (window.__MAP_DATA__ = {dd:[...], areas:[{e:[[estate, en, zh, ddIdx, ...]]}]}).
    A cached copy (estate_district_map.json) is used when index.html is unavailable."""
    cache = os.path.join(HERE, "estate_district_map.json")
    try:
        idx = next((x for x in (INDEX_HTML, os.path.join(HERE, "index.html")) if os.path.exists(x)), None)
        if not idx:
            log(f"estate map: index.html not found at {INDEX_HTML} or next to the script")
        if idx:
            s = open(idx, encoding="utf-8").read()
            k = "window.__MAP_DATA__="
            a = s.index(k) + len(k)
            D, _ = json.JSONDecoder().raw_decode(s[a:])
            est = {}
            for ar in D["areas"]:
                for e in ar["e"]:
                    d = D["dd"][e[3]]
                    est[e[0]] = DISTRICT_ALIAS.get(d, d)
            json.dump(est, open(cache, "w", encoding="utf-8"))
            return est
    except Exception as ex:
        log(f"estate map from index.html failed ({ex}); trying cache")
    if os.path.exists(cache):
        return json.load(open(cache, encoding="utf-8"))
    log("WARNING: no estate map available - only car-line / area fallback will be used")
    return {}


def _area_key(addr):
    parts = [x.strip() for x in str(addr or "").split(",")]
    return "/".join(parts[:2]) if len(parts) >= 2 and parts[0] in ("香港", "九龍", "新界") else ""


def map_districts(df):
    """df: one row per waybill. Adds column 'district'.
    Order: (1) carline logic (district_from_carline)  (2) Estate Code via delivery-map data  (3) area learned from (1)+(2)  (4) Other."""
    est = load_estate_map()
    d1 = df["送貨車號"].map(district_from_carline)
    d2 = df["Estate Code"].map(est)
    dist = d1.where(d1.notna(), d2)
    # (3) area majority, learned from the confident rows (+ persisted counts from earlier runs)
    learned_path = os.path.join(HERE, "rolling_area_learned.json")
    learned = defaultdict(Counter)
    if os.path.exists(learned_path):
        for k, v in json.load(open(learned_path, encoding="utf-8")).items():
            learned[k].update(v)
    area = df["Address"].map(_area_key)
    for a, d in zip(area[dist.notna()], dist[dist.notna()]):
        if a:
            learned[a][d] += 1
    try:
        json.dump({k: dict(v) for k, v in learned.items()}, open(learned_path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    area_map = {a: c.most_common(1)[0][0] for a, c in learned.items() if c}
    d3 = area.map(area_map)
    out = dist.where(dist.notna(), d3)
    # Macau / ShipEazy etc. have no area token that maps -> Other
    df["district"] = out.fillna(OTHER)
    df["_src"] = pd.Series("carline", index=df.index).where(d1.notna(),
                 pd.Series("estate", index=df.index).where(d2.notna(),
                 pd.Series("area", index=df.index).where(d3.notna(), "other")))
    return df


# ----------------------------------------------------------------------------- OIX report processing
COLS_NEEDED = ["Status History", "Current Status", "Previous Status", "Waybill Number", "Order Number",
               "Parent Order Number", "Order Creation Date", "Expected Delivery", "Estate Code",
               "送貨車號", "Address"]


def read_oix_csv(path):
    df = pd.read_csv(path, skiprows=1, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    if "Waybill Number" not in df.columns:      # tolerate a missing title line
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig", keep_default_na=False)
    missing = [c for c in COLS_NEEDED if c not in df.columns]
    if missing:
        raise ValueError(f"{os.path.basename(path)} is missing columns {missing}")
    return df


def parent_order(order_no, parent_no):
    """Grouping logic of section 2."""
    p, k = (parent_no or "").strip(), (order_no or "").strip()
    if p:
        if p.startswith("EM"):
            return p[1:]
        return p                                  # begins with H (or anything else) -> use directly
    if k.startswith("H"):
        return k[:13]
    if k.startswith("EM"):
        return k[1:]
    return k


_DATE_FORMATS = ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%Y%m%d")


def _date_part(series):
    """Date part of an OIX date/time string. The export format is not always m/d/Y, so every
    known format is tried and the one that parses the most rows wins."""
    head = series.astype(str).str.strip().str.split(r"[ T]", n=1, regex=True).str[0]
    best, best_n = None, -1
    for fmt in _DATE_FORMATS:
        parsed = pd.to_datetime(head, format=fmt, errors="coerce")
        n = int(parsed.notna().sum())
        if n > best_n:
            best, best_n = parsed, n
    return best


def process_report(path, t0):
    """Returns (rows, stats): rows = DataFrame (1 row / waybill) with
    district, parent, delivery (date) for orders created on t0 and not cancelled."""
    df = read_oix_csv(path)
    n_raw = len(df)
    # order-creation date = T-0  (column M, only the date before the time)
    created = _date_part(df["Order Creation Date"])
    df = df[created == pd.Timestamp(t0)]
    n_t0 = len(df)
    if n_t0 == 0:
        raw_vals = read_oix_csv(path)["Order Creation Date"]
        log(f"DIAG {os.path.basename(path)}: 0 rows with Order Creation Date = {t0:%Y-%m-%d}")
        log(f"DIAG sample raw values: {raw_vals.head(3).tolist()}")
        log(f"DIAG parsed {int(created.notna().sum())}/{len(created)} dates; "
            f"min={created.min()} max={created.max()}")
        log(f"DIAG most common dates: {created.dt.strftime('%Y-%m-%d').value_counts().head(5).to_dict()}")
    # remove every row of a waybill that has a status containing "Cancel" in A / C / D
    cancel = (df["Status History"].str.contains("cancel", case=False, na=False)
              | df["Current Status"].str.contains("cancel", case=False, na=False)
              | df["Previous Status"].str.contains("cancel", case=False, na=False))
    bad = set(df.loc[cancel, "Waybill Number"])
    df = df[~df["Waybill Number"].isin(bad)]
    n_nocancel = len(df)
    # parent order grouping, then de-duplicate by waybill
    df = df.assign(parent=[parent_order(k, p) for k, p in zip(df["Order Number"], df["Parent Order Number"])])
    df = df.drop_duplicates("Waybill Number").copy()
    df["delivery"] = _date_part(df["Expected Delivery"])
    df = df[df["delivery"].notna()]
    df = map_districts(df)
    stats = {"rawRows": n_raw, "t0Rows": n_t0, "afterCancelRows": n_nocancel,
             "cancelledWaybills": len(bad), "waybills": len(df),
             "mappedBy": df["_src"].value_counts().to_dict(),
             "other": int((df["district"] == OTHER).sum())}
    return df, stats


def count_cells(rows):
    """{deliveryDate: {district|'OVERALL': [orders, waybills]}} - orders = distinct parent order numbers."""
    cells = {}
    for dte, g in rows.groupby(rows["delivery"].dt.strftime("%Y-%m-%d")):
        cell = {"OVERALL": [int(g["parent"].nunique()), int(g["Waybill Number"].nunique())]}
        for d, gd in g.groupby("district"):
            cell[d] = [int(gd["parent"].nunique()), int(gd["Waybill Number"].nunique())]
        cells[dte] = cell
    return cells


def build_snapshot(t0, slot, next_file, same_file):
    cells, stats = {}, {}
    t0s = t0.strftime("%Y-%m-%d")
    if same_file and os.path.exists(same_file):
        rows, st = process_report(same_file, t0)
        stats["sameDay"] = st
        c = count_cells(rows)
        if t0s in c:
            cells[t0s] = c[t0s]
    if next_file and os.path.exists(next_file):
        rows, st = process_report(next_file, t0)
        stats["nextWeek"] = st
        for k, v in count_cells(rows).items():
            if k == t0s and t0s in cells:
                continue                       # T+0 comes from the "same day" report
            cells[k] = v
    return cells, stats


# ----------------------------------------------------------------------------- JSON + embedding
def update_json(t0, slot, cells, stats):
    data = {}
    if os.path.exists(ROLLING_JSON):
        try:
            data = json.load(open(ROLLING_JSON, encoding="utf-8"))
        except Exception:
            data = {}
    data.setdefault("days", {})
    data.setdefault("snapshots", {})
    t0s = t0.strftime("%Y-%m-%d")
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    data["days"][t0s] = {"slot": slot, "asOf": now, "cells": cells}
    data["snapshots"].setdefault(t0s, {})[slot] = {"asOf": now, "cells": cells}
    keep = sorted(data["days"])[-KEEP_ORDER_DATES:]
    data["days"] = {k: data["days"][k] for k in keep}
    data["snapshots"] = {k: v for k, v in data["snapshots"].items() if k in keep}
    data["districts"] = DISTRICTS
    data["generatedAt"] = now
    data["lastRun"] = {"orderDate": t0s, "slot": slot, "stats": stats}
    tmp = ROLLING_JSON + ".tmp"
    json.dump(data, open(tmp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    os.replace(tmp, ROLLING_JSON)
    log(f"wrote {ROLLING_JSON}")
    return data


def embed_in_index(data):
    """Splice rolling_orders.json into window.__EMBEDDED_DATA__ (between the EMBEDDED_DATA markers)."""
    if not os.path.exists(INDEX_HTML):
        log("index.html not found - embedding skipped")
        return False
    s = open(INDEX_HTML, encoding="utf-8").read()
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
    new_blk = head + "<script>window.__EMBEDDED_DATA__ = " + json.dumps(emb, ensure_ascii=False) + ";</script>" + tail
    s2 = s[:a] + new_blk + s[e:]
    shutil.copy2(INDEX_HTML, INDEX_HTML + ".bak_rolling")
    tmp = INDEX_HTML + ".tmp"
    open(tmp, "w", encoding="utf-8").write(s2)
    os.replace(tmp, INDEX_HTML)
    log("embedded rolling_orders.json into index.html")
    return True


# ----------------------------------------------------------------------------- screenshot + WhatsApp
def screenshot_table(t0, slot):
    from playwright.sync_api import sync_playwright
    os.makedirs(SHOT_DIR, exist_ok=True)
    out = os.path.join(SHOT_DIR, f"Rolling Order Trend ({mmm_dd(t0)})_{slot.replace(':', '')}.png")
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        try:
            pg = b.new_page(viewport={"width": 1700, "height": 1100}, device_scale_factor=1.5)
            pg.goto("file://" + INDEX_HTML, wait_until="domcontentloaded", timeout=120000)
            pg.wait_for_selector(".main-tab[data-target='view-rolling']", timeout=120000)
            pg.click(".main-tab[data-target='view-rolling']")
            pg.wait_for_selector("#rollingTableWrap table", timeout=60000)
            pg.add_style_tag(content='#rollingShot .tbl-dl-bar{display:none !important}')
            time.sleep(1)
            pg.locator("#rollingShot").screenshot(path=out)
        finally:
            b.close()
    log(f"screenshot {out}")
    return out


def whatsapp_text(t0, slot):
    sess = f"{mmm_dd(t0)} Overall" if slot == "00:00" else slot
    return (f"Rolling Order Trend ({mmm_dd(t0)})\n\n"
            f"Attached please find the rolling orders trend as of {sess} above.")


def send_whatsapp(image_path, text):
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
        page.wait_for_selector('div[contenteditable="true"][data-tab="10"]', timeout=10000)

        # attach the picture (Photos & videos input)
        for sel in ['div[aria-label="Attach"]', 'button[aria-label="Attach"]', 'span[data-icon="plus"]',
                    'span[data-icon="attach-menu-plus"]']:
            try:
                btn = page.locator(sel).first
                if btn.is_visible():
                    btn.click()
                    time.sleep(1)
                    break
            except Exception:
                continue
        done = False
        for sel in ['input[accept*="image"]', 'input[type="file"]']:
            try:
                inputs = page.query_selector_all(sel)
                if inputs:
                    inputs[0].set_input_files(image_path)
                    done = True
                    break
            except Exception:
                continue
        if not done:
            raise RuntimeError("could not attach image in WhatsApp Web")
        time.sleep(4)

        # caption (= the message) in the preview; falls back to a separate text message
        caption_ok = False
        for sel in ['div[aria-label*="caption" i][contenteditable="true"]',
                    'div[aria-label*="說明"][contenteditable="true"]',
                    'div[contenteditable="true"][data-tab="10"]']:
            try:
                box = page.locator(sel).last
                if box.is_visible(timeout=3000):
                    box.click()
                    page.evaluate("(t) => navigator.clipboard.writeText(t)", text)
                    page.keyboard.press("Control+v")
                    page.wait_for_timeout(800)
                    caption_ok = True
                    break
            except Exception:
                continue
        sent = False
        for sel in ['span[data-icon="send"]', 'div[aria-label="Send"]', 'div[aria-label="發送"]', 'button[aria-label="Send"]']:
            try:
                b = page.locator(sel).first
                if b.is_visible(timeout=5000):
                    b.click()
                    sent = True
                    break
            except Exception:
                continue
        if not sent:
            page.keyboard.press("Enter")
        time.sleep(6)
        if not caption_ok:
            inp = page.locator('div[contenteditable="true"][data-tab="10"]')
            inp.click()
            page.evaluate("(t) => navigator.clipboard.writeText(t)", text)
            page.keyboard.press("Control+v")
            page.wait_for_timeout(800)
            page.keyboard.press("Enter")
            time.sleep(4)
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
    ap.add_argument("--slot", choices=SLOTS, help="06:00 | 13:00 | 18:00 | 00:00 (default: nearest to now)")
    ap.add_argument("--date", help="override T-0 (YYYY-MM-DD)")
    ap.add_argument("--skip-download", action="store_true", help="use the files already in Folder for Rolling Orders")
    ap.add_argument("--process-only", action="store_true", help="skip download + WhatsApp; needs --next-week-file")
    ap.add_argument("--next-week-file")
    ap.add_argument("--same-day-file")
    ap.add_argument("--show-browser", action="store_true", help="open a visible Chrome for the OIX download (headless=False) so you can watch it")
    ap.add_argument("--download-only", action="store_true", help="only download the OIX report(s) into Folder for Rolling Orders, then stop")
    ap.add_argument("--skip-whatsapp", action="store_true")
    ap.add_argument("--skip-embed", action="store_true")
    a = ap.parse_args()

    if a.show_browser:
        global OIX_HEADLESS
        OIX_HEADLESS = False
    slot = a.slot or infer_slot()
    t0 = datetime.strptime(a.date, "%Y-%m-%d") if a.date else session_dates(slot)
    log(f"slot={slot}  T-0={t0:%Y-%m-%d}")
    with run_lock():
        nxt, same = report_paths(t0)
        if a.next_week_file:
            nxt = a.next_week_file
        if a.same_day_file:
            same = a.same_day_file
        if not (a.skip_download or a.process_only):
            download_oix_reports(t0, slot)
        if a.download_only:
            log("download-only: done")
            return
        for label, f in (("next week", nxt), ("same day", same)):
            log(f"{label} file: {f}  ->  {'FOUND' if os.path.exists(f) else 'NOT FOUND'}")
        cells, stats = build_snapshot(t0, slot, nxt, same)
        log("stats: " + json.dumps(stats, ensure_ascii=False))
        if not cells:
            if not stats:
                raise SystemExit("no report file found - run without --skip-download / --process-only, or pass --next-week-file")
            raise SystemExit(f"report(s) found but 0 usable rows for order date {t0:%Y-%m-%d} "
                             "(wrong --date, or no non-cancelled orders created that day) - nothing written")
        data = update_json(t0, slot, cells, stats)
        if not a.skip_embed:
            embed_in_index(data)
        if not (a.skip_whatsapp or a.process_only):
            img = screenshot_table(t0, slot)
            send_whatsapp(img, whatsapp_text(t0, slot))


if __name__ == "__main__":
    sys.path.insert(0, HERE)
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
