"""
LOG · KPI Tracker — data pipeline
==================================
Pulls the 5 leadership matrices from their real sources (a local Excel export
+ several Tableau reports), applies the business rules from
KPI_Dashboard_Build-up_Script.txt, and writes the result to data.json for the
dashboard to read.

This is a STANDALONE PROGRAM. It does not run in the browser and is not part
of log-kpi-tracker.html — it runs on your Windows machine (Task Scheduler)
and produces the data file the website reads.

SECRETS: never hardcode the Tableau key or a GitHub token in this file.
Put them in a local `.env` file (see .env.example) that stays OUT of git
(it's listed in .gitignore). This script loads them from environment
variables at runtime.

Run modes (matches the two update schedules in the spec):
    python kpi_pipeline.py --section productivity   # 03:00 daily — Excel-based
    python kpi_pipeline.py --section tableau         # 15:00 daily — Tableau-based
    python kpi_pipeline.py --section all             # run everything (for testing)

First-time / debugging:
    python kpi_pipeline.py --section tableau --dump-raw
        Saves the raw CSV Tableau returns for each view into ./raw_dumps/
        WITHOUT touching data.json, so you can check real column names
        before trusting the parsing logic below. Tableau's "view data"
        endpoint returns a FLAT underlying-data table (one row per mark),
        not a visual crosstab — so the "merged cell" issue from looking at
        the dashboard by eye mostly disappears once we read the real data.
        Even so, exact column header text can vary by report; the DEBUG
        RUN is how you confirm it before the first real scheduled run.
"""

import os
import re
import sys
import json
import glob
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd
import requests

try:
    from dotenv import load_dotenv
    load_dotenv()  # reads ./.env if present — never commit that file (see .gitignore)
except ImportError:
    pass  # falls back to real environment variables if python-dotenv isn't installed

# --------------------------------------------------------------------------
# CONFIG — loaded from environment variables. See .env.example.
# --------------------------------------------------------------------------
TABLEAU_SERVER = os.environ.get("TABLEAU_SERVER", "https://inhouse-analytics.hktv.com.hk")
TABLEAU_SITE = os.environ.get("TABLEAU_SITE", "")  # content URL of the site, blank = Default
TABLEAU_PAT_NAME = os.environ.get("TABLEAU_PAT_NAME", "")
TABLEAU_PAT_SECRET = os.environ.get("TABLEAU_PAT_SECRET", "")
API_VERSION = "3.22"

OIX_FOLDER = os.environ.get("OIX_FOLDER", r"C:\Users\chipanl\Downloads\Digimobi Report")
DATA_JSON_PATH = os.environ.get("DATA_JSON_PATH", "./public/data.json")  # public/ is the only folder the Cloudflare Worker serves
HISTORY_PATH = os.environ.get("HISTORY_PATH", "./history.json")
RAW_DUMP_DIR = Path("./raw_dumps")

# Report workbook content URLs (from the URLs in the brief) — used to look
# up each view's numeric id at runtime, since the REST "view data" endpoint
# needs the id, not the content URL.
REPORTS = {
    "logistics_kpi": "LogisticsKPIReport/LogisticsKPI",
    "rfid": "RFIDReport_V3/RFIDReport-expecteddeliverydate",
    "report_a_3r": "MonthlyRP-3RReportA/MonthlyRP-3RSummary",
    "report_b_cs_merchant": "MonthlyRP-CSCancelMerchantPaymentReportB/MonthlyRP-CSCancelMerchantPaymentSummary",
    "report_c_cs_cancel": "MonthlyRP-CSCancelReportC/MonthlyRP-CSCancelSummary",
}

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]
# "NT-YT" isn't one of the 10 official codes but appears in some source data
# and must be folded into NT-TW per the spec.
DISTRICT_MATCH_ORDER = sorted(DISTRICTS + ["NT-YT"], key=len, reverse=True)


def today_hkt():
    # Adjust if the machine running this isn't already on HK time.
    return dt.date.today()


def normalize_district_code(code):
    """NT-YT folds into NT-TW; everything else passes through unchanged."""
    return "NT-TW" if code == "NT-YT" else code


def district_from_party_name(name):
    """Match a 'Final RP' / 'Final Responsible Party' style name (e.g.
    'ETH-盛/望', 'ETH-盛/望UP') to one of the 10 district codes by prefix.
    Longest-code-first so 'NT-TSM' isn't mis-matched as 'NT-T...'."""
    if not isinstance(name, str):
        return None
    for code in DISTRICT_MATCH_ORDER:
        if name.startswith(code):
            return normalize_district_code(code)
    return None


def is_pd_party(name):
    return isinstance(name, str) and name.startswith("PD")


# Replicates the Excel IFS() formula on 送貨車號 (Column P) EXACTLY in order —
# IFS returns the first TRUE condition, so order matters and must not be
# changed. Excel's SEARCH() is case-insensitive, so this matches that.
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


def col(letter):
    """Excel column letter -> 0-based index. A=0, E=4, K=10, L=11, P=15, R=17."""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


# --------------------------------------------------------------------------
# SECTION 1 — Productivity / Staff (Excel-based, runs at 03:00)
# --------------------------------------------------------------------------

def find_oix_file(target_date):
    """OIX_Record_YYYYMMDD.xlsx for the given date (T-1), inside OIX_FOLDER."""
    fname = f"OIX_Record_{target_date.strftime('%Y%m%d')}.xlsx"
    path = os.path.join(OIX_FOLDER, fname)
    if os.path.exists(path):
        return path
    # fall back to a glob in case of a slightly different extension/casing
    matches = glob.glob(os.path.join(OIX_FOLDER, f"OIX_Record_{target_date.strftime('%Y%m%d')}*"))
    if matches:
        return matches[0]
    raise FileNotFoundError(f"Could not find {fname} in {OIX_FOLDER}")


def load_oix(path):
    # Row 1 (index 0) is the report title "Waybill Status History Report" —
    # header=1 skips it and treats the real header row as row 2.
    df = pd.read_excel(path, header=1, dtype=str)
    return df


def process_oix(df):
    """Implements the OIX cleaning + district tagging steps from the spec.
    Returns a cleaned DataFrame with a 'District' column added."""
    c_user, c_addr = col("E"), col("R")
    c_order_no, c_parent = col("K"), col("L")
    c_truck = col("P")

    user = df.iloc[:, c_user].fillna("")
    addr = df.iloc[:, c_addr].fillna("")

    valid_prefixes = ("LF", "LP", "ODS", "VAN")
    remove_mask = addr.str.contains("O2O", na=False) & ~user.str.startswith(valid_prefixes)
    df = df.loc[~remove_mask].copy()

    # Refresh series after row removal (index alignment)
    parent = df.iloc[:, c_parent]
    order_no = df.iloc[:, c_order_no]
    blank_parent = parent.isna() | (parent.astype(str).str.strip() == "")
    filled_parent = order_no.astype(str).str.slice(0, 13)
    df.iloc[:, c_parent] = parent.where(~blank_parent, filled_parent)

    df = df.drop_duplicates(subset=[df.columns[c_parent], df.columns[c_user]])

    df["District"] = df.iloc[:, c_truck].apply(district_from_truck_no)

    unmatched = df["District"].isna().sum()
    if unmatched:
        print(f"  WARNING: {unmatched} row(s) had a truck number that matched none of the "
              f"14 district patterns — they will be silently excluded from every total. "
              f"Check ./raw_dumps or inspect df.iloc[:, {c_truck}] for new/unrecognized formats.")

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
    target_date = today_hkt() - dt.timedelta(days=1)  # T-1
    path = find_oix_file(target_date)
    df = load_oix(path)
    df = process_oix(df)

    hktv_staff = productivity_for_group(df, ("LF", "LP"))
    ods_ratio = productivity_for_group(df, ("ODS", "VAN"))

    return {
        "asOf": target_date.isoformat(),
        "hktvStaff": hktv_staff,
        "odsRatio": ods_ratio,
    }


# --------------------------------------------------------------------------
# SECTION 2 — Tableau REST API helpers
# --------------------------------------------------------------------------

def tableau_signin():
    if not TABLEAU_PAT_NAME or not TABLEAU_PAT_SECRET:
        raise RuntimeError(
            "TABLEAU_PAT_NAME / TABLEAU_PAT_SECRET are not set. "
            "The API key you gave me looks like it may be a combined string — "
            "Tableau Personal Access Tokens are normally issued as two separate "
            "values (a Name you choose, and a Secret Tableau generates). "
            "Check Tableau Server > My Account Settings > Personal Access Tokens "
            "and fill both env vars separately in .env — see chat for details."
        )
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/auth/signin"
    body = {
        "credentials": {
            "personalAccessTokenName": TABLEAU_PAT_NAME,
            "personalAccessTokenSecret": TABLEAU_PAT_SECRET,
            "site": {"contentUrl": TABLEAU_SITE},
        }
    }
    res = requests.post(url, json=body, headers={"Accept": "application/json"}, timeout=30)
    res.raise_for_status()
    data = res.json()
    return data["credentials"]["token"], data["credentials"]["site"]["id"]


def tableau_find_view_id(token, site_id, content_url):
    """Look up a view's numeric/LUID id from its content URL (workbook/view path)."""
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/views"
    headers = {"X-Tableau-Auth": token, "Accept": "application/json"}
    res = requests.get(url, headers=headers, params={"filter": f"contentUrl:eq:{content_url}"}, timeout=30)
    res.raise_for_status()
    views = res.json().get("views", {}).get("view", [])
    if not views:
        raise RuntimeError(f"No view found for contentUrl={content_url}")
    return views[0]["id"]


def tableau_view_data(token, site_id, view_id):
    """Returns the view's underlying data as a pandas DataFrame (flat rows,
    not a visual crosstab — merged-cell issues from the dashboard UI do not
    apply here)."""
    url = f"{TABLEAU_SERVER}/api/{API_VERSION}/sites/{site_id}/views/{view_id}/data"
    headers = {"X-Tableau-Auth": token, "Accept": "text/csv"}
    res = requests.get(url, headers=headers, timeout=60)
    res.raise_for_status()
    from io import StringIO
    return pd.read_csv(StringIO(res.text))


def find_column(df, *keywords):
    """Case-insensitive substring match against column names — real Tableau
    field labels can differ slightly from what's shown on screen, so this
    is more robust than requiring an exact name."""
    for c in df.columns:
        low = str(c).lower()
        if all(k.lower() in low for k in keywords):
            return c
    raise KeyError(f"No column matching {keywords} in {list(df.columns)}")


def dump_raw(name, df):
    RAW_DUMP_DIR.mkdir(exist_ok=True)
    path = RAW_DUMP_DIR / f"{name}.csv"
    df.to_csv(path, index=False)
    print(f"  wrote {path} ({len(df)} rows, columns={list(df.columns)})")


# --------------------------------------------------------------------------
# SECTION 2a — Delay Rate & Poor Rating (same source view, two tables)
# --------------------------------------------------------------------------

def parse_delay_rate(df, month_label):
    """df = underlying data for the 'Delivery On Time Rate' table.
    ASSUMPTION (verify with --dump-raw): columns include a district/zone
    dimension, a month dimension matching `month_label` (e.g. '2026年8月'),
    and a 'Delay %' measure. Adjust the find_column() keywords below once
    you've seen the real column names."""
    month_col = find_column(df, "month") if any("month" in str(c).lower() for c in df.columns) else find_column(df, "year")
    zone_col = find_column(df, "zone") if any("zone" in str(c).lower() for c in df.columns) else find_column(df, "district")
    value_col = find_column(df, "delay")

    cur = df[df[month_col].astype(str).str.contains(re.escape(month_label))]
    per_district = {}
    for d in DISTRICTS:
        rows = cur[cur[zone_col].astype(str) == d]
        per_district[d] = float(rows[value_col].iloc[0]) if len(rows) else None
    overall_rows = cur[cur[zone_col].astype(str).str.lower().isin(["overall", "total", "all"])]
    overall = float(overall_rows[value_col].iloc[0]) if len(overall_rows) else None
    return {"overall": overall, "districts": per_district}


def parse_poor_rating(df, month_label):
    """Same shape as delay rate, different source table/measure."""
    month_col = find_column(df, "month") if any("month" in str(c).lower() for c in df.columns) else find_column(df, "year")
    zone_col = find_column(df, "zone") if any("zone" in str(c).lower() for c in df.columns) else find_column(df, "district")
    value_col = find_column(df, "rating")

    cur = df[df[month_col].astype(str).str.contains(re.escape(month_label))]
    per_district = {}
    for d in DISTRICTS:
        rows = cur[cur[zone_col].astype(str) == d]
        per_district[d] = float(rows[value_col].iloc[0]) if len(rows) else None
    overall_rows = cur[cur[zone_col].astype(str).str.lower().isin(["overall", "total", "all"])]
    overall = float(overall_rows[value_col].iloc[0]) if len(overall_rows) else None
    return {"overall": overall, "districts": per_district}


# --------------------------------------------------------------------------
# SECTION 2b — Missing & Lost Amount (3 reports summed)
# --------------------------------------------------------------------------

def parse_report_a(df, month_label):
    rp_group_col = find_column(df, "rp group")
    final_col = find_column(df, "final")
    total_col = find_column(df, "total")

    rows = df[df[rp_group_col].astype(str).str.strip() == "Bert (Log)"]
    rows = rows[rows[final_col].apply(lambda n: district_from_party_name(n) is not None)]

    per_district = {d: 0.0 for d in DISTRICTS}
    for _, r in rows.iterrows():
        d = district_from_party_name(r[final_col])
        per_district[d] += float(r[total_col])
    overall = sum(per_district.values())
    return {"overall": round(overall, 2), "districts": {d: round(v, 2) for d, v in per_district.items()}}


def parse_report_b(df, month_label):
    rp_group_col = find_column(df, "rp group")
    final_col = find_column(df, "final")
    total_col = find_column(df, "total")

    rows = df[df[rp_group_col].astype(str).str.strip().isin(["Bert (Log)", "Bert (log) - DAMUP", "Bert(Log)", "Bert(log) - DAMUP"])]

    per_district = {d: 0.0 for d in DISTRICTS}
    pd_total = 0.0
    for _, r in rows.iterrows():
        name = r[final_col]
        d = district_from_party_name(name)
        if d:
            per_district[d] += float(r[total_col])
        elif is_pd_party(name):
            pd_total += float(r[total_col])
    overall = sum(per_district.values()) + pd_total
    return {"overall": round(overall, 2), "districts": {d: round(v, 2) for d, v in per_district.items()}}


def parse_report_c(df, month_label):
    rp_group_col = find_column(df, "rp group")
    final_col = find_column(df, "final")
    total_col = find_column(df, "total")

    rows = df[df[rp_group_col].astype(str).str.strip() == "Bert (Log)"]

    per_district = {d: 0.0 for d in DISTRICTS}
    pd_total = 0.0
    for _, r in rows.iterrows():
        name = r[final_col]
        d = district_from_party_name(name)
        if d:
            per_district[d] += float(r[total_col])
        elif is_pd_party(name):
            pd_total += float(r[total_col])
    overall = sum(per_district.values()) + pd_total
    return {"overall": round(overall, 2), "districts": {d: round(v, 2) for d, v in per_district.items()}}


def combine_missing_lost(a, b, c):
    per_district = {d: round(a["districts"][d] + b["districts"][d] + c["districts"][d], 2) for d in DISTRICTS}
    overall = round(a["overall"] + b["overall"] + c["overall"], 2)
    return {"overall": overall, "districts": per_district}


# --------------------------------------------------------------------------
# SECTION 2c — RFID Missing Tote (accumulates within the month)
# --------------------------------------------------------------------------

def parse_rfid(df, target_col_label):
    """df = 'RP Breakdown (7 days)' underlying data. We only want the single
    column matching target_col_label (e.g. '2026年8月14日', the T-4 date)."""
    zone_col = find_column(df, "zone") if any("zone" in str(c).lower() for c in df.columns) else find_column(df, "district")
    date_col = find_column(df, "date")
    value_col = find_column(df, "value") if any("value" in str(c).lower() for c in df.columns) else find_column(df, "count")

    rows = df[df[date_col].astype(str).str.contains(re.escape(target_col_label))]
    per_district = {}
    for d in DISTRICTS:
        match = rows[rows[zone_col].astype(str) == d]
        per_district[d] = float(match[value_col].iloc[0]) if len(match) else 0.0
    overall_rows = rows[rows[zone_col].astype(str).str.lower().isin(["overall", "total", "all"])]
    overall = float(overall_rows[value_col].iloc[0]) if len(overall_rows) else sum(per_district.values())
    return {"overall": overall, "districts": per_district}


# --------------------------------------------------------------------------
# SECTION 3 — history log (for rolling-average forecasts) + prorate forecast
# --------------------------------------------------------------------------

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
        vals = [series[dt_][["districts"]][0].get(d) if False else series[dt_]["districts"].get(d)
                for dt_ in dates if series[dt_]["districts"].get(d) is not None]
        district_avgs[d] = round(sum(vals) / len(vals), 2) if vals else None
    return overall_avg, district_avgs


def prorate_forecast(actual, data_date, total_days):
    if actual is None:
        return None
    return round(actual * (total_days / data_date), 2)


# --------------------------------------------------------------------------
# SECTION 4 — assemble data.json
# --------------------------------------------------------------------------

def load_data_json():
    if os.path.exists(DATA_JSON_PATH):
        with open(DATA_JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"generatedAt": None, "matrices": {}}


def save_data_json(payload):
    payload["generatedAt"] = dt.datetime.now().isoformat(timespec="minutes")
    with open(DATA_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    print(f"Wrote {DATA_JSON_PATH}")


def run_section_productivity():
    result = run_productivity_section()
    today = dt.date.today()
    history = load_history()

    payload = load_data_json()
    matrices = payload.setdefault("matrices", {})

    for key, group in (("hktvStaff", result["hktvStaff"]), ("odsRatio", result["odsRatio"])):
        append_history(history, key, result["asOf"], group["overall"],
                        {d: group["districts"][d]["productivity"] for d in DISTRICTS})
        fc_overall, fc_districts = rolling_average(history, key, 7, today)
        matrices[key] = {
            "actual": {"overall": group["overall"],
                       "districts": {d: group["districts"][d]["productivity"] for d in DISTRICTS}},
            "forecast": {"overall": fc_overall, "districts": fc_districts},
            "asOf": result["asOf"],
        }

    save_history(history)
    save_data_json(payload)


def run_section_tableau(dump_raw_only=False):
    today = dt.date.today()
    month_label = f"{today.year}年{today.month}月"
    total_days = (dt.date(today.year + (today.month == 12), (today.month % 12) + 1, 1) - dt.timedelta(days=1)).day

    token, site_id = tableau_signin()
    view_ids = {k: tableau_find_view_id(token, site_id, v) for k, v in REPORTS.items()}

    df_logistics = tableau_view_data(token, site_id, view_ids["logistics_kpi"])
    df_a = tableau_view_data(token, site_id, view_ids["report_a_3r"])
    df_b = tableau_view_data(token, site_id, view_ids["report_b_cs_merchant"])
    df_c = tableau_view_data(token, site_id, view_ids["report_c_cs_cancel"])
    df_rfid = tableau_view_data(token, site_id, view_ids["rfid"])

    if dump_raw_only:
        print("Dumping raw views for inspection (data.json NOT modified):")
        dump_raw("logistics_kpi", df_logistics)
        dump_raw("report_a_3r", df_a)
        dump_raw("report_b_cs_merchant", df_b)
        dump_raw("report_c_cs_cancel", df_c)
        dump_raw("rfid", df_rfid)
        return

    history = load_history()
    payload = load_data_json()
    matrices = payload.setdefault("matrices", {})

    # --- Delay Rate & Poor Rating (30-day rolling forecast) ---
    delay = parse_delay_rate(df_logistics, month_label)
    poor = parse_poor_rating(df_logistics, month_label)
    for key, val in (("delayRate", delay), ("poorRating", poor)):
        append_history(history, key, today.isoformat(), val["overall"], val["districts"])
        fc_overall, fc_districts = rolling_average(history, key, 30, today)
        matrices[key] = {
            "actual": val,
            "forecast": {"overall": fc_overall, "districts": fc_districts},
            "asOf": today.isoformat(),
        }

    # --- Missing & Lost Amount (prorate forecast) ---
    a = parse_report_a(df_a, month_label)
    b = parse_report_b(df_b, month_label)
    c = parse_report_c(df_c, month_label)
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
    t4 = today - dt.timedelta(days=4)
    t4_label = f"{t4.year}年{t4.month}月{t4.day}日"
    rfid_increment = parse_rfid(df_rfid, t4_label)

    month_key = today.strftime("%Y-%m")
    rfid_state = payload.setdefault("rfidMonthly", {})
    bucket = rfid_state.setdefault(month_key, {"overall": 0.0, "districts": {d: 0.0 for d in DISTRICTS}})
    bucket["overall"] += rfid_increment["overall"]
    for d in DISTRICTS:
        bucket["districts"][d] += rfid_increment["districts"][d]

    matrices["rfidMissingTote"] = {
        "actual": bucket,
        "forecast": {
            "overall": prorate_forecast(bucket["overall"], today.day, total_days),
            "districts": {d: prorate_forecast(bucket["districts"][d], today.day, total_days) for d in DISTRICTS},
        },
        "asOf": today.isoformat(),
        "monthKey": month_key,
    }
    # Keep at most this month + last month visible, per the "show 2 months" requirement.
    keep_keys = sorted(rfid_state.keys())[-2:]
    payload["rfidMonthly"] = {k: rfid_state[k] for k in keep_keys}

    save_history(history)
    save_data_json(payload)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--section", choices=["productivity", "tableau", "all"], required=True)
    parser.add_argument("--dump-raw", action="store_true",
                         help="Tableau section only: save raw CSVs to ./raw_dumps/ without touching data.json")
    args = parser.parse_args()

    if args.section in ("productivity", "all"):
        run_section_productivity()
    if args.section in ("tableau", "all"):
        run_section_tableau(dump_raw_only=args.dump_raw)


if __name__ == "__main__":
    main()
