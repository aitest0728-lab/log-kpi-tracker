#!/usr/bin/env python3
"""
Order Finished Pattern   v1.1.1 (WSL / Linux)      feeds the dashboard tab "Order Finished Pattern"
===================================================================================================
Keep this file in the SAME folder as carline_progress_monitor.py / delay_forecast_cargroup_by_dp.py /
delay_forecast_cargroup.py / kpi_pipeline.py (log-kpi-tracker).

WHAT'S NEW IN v1.1.1
  * OIX downloads use a dedicated Playwright Chromium profile (launch_persistent_context under
    .ofp-browser-profile / FINISH_BROWSER_PROFILE_DIR) so OFP no longer shares a browser with
    carline_express_report / carline_progress_monitor / delay_forecast_*. Credentials still come from OIX.env via
    the existing oix_env / delay_forecast login helpers.
  * No longer waits up to 20 minutes on the other scripts' /tmp locks before starting a download (own browser =
    safe to run in parallel). OFP's own /tmp/order_finished_pattern.lock is unchanged.

WHAT'S NEW IN v1.1.0
  * Downloads the OIX "Overall Order Report" (= Route Resources Planning Report, same download logic as
    delay_forecast_cargroup.download_overall: 3 attempts, size / row-count check, last-good fallback) and the
    Delivery Route Zone ID Report (route -> car group, once per day) next to the two Waybill Status History reports.
  * New row "Outstanding Order Count" = unique parent orders of the Overall Order Report (same exclusions as the finished
    table) that are NOT finished (RECEIVED_BY_CUSTOMER / FAIL_TO_DELIVER) at the report time.
  * Row names: "Sum of hours" -> "Finished Orders", "Daily distinct" -> "Unique Orders Count",
    "Finished more than once" -> "Duplicate Orders".
  * "asOf" is now the time the OIX reports were DOWNLOADED (latest save time of the downloaded files; capped at 24:00 of
    the delivery day), no longer the time the script started. The time of each report is kept in "reportTimes".
  * Week-over-week: the dashboard compares a day with the SAME WEEKDAY LAST WEEK (date - 7). When that day is missing
    in order_finished_pattern.json (or is not final / has no Outstanding data yet) the script BACK-FILLS it from OIX at the
    next run (same browser flow, all reports of that day). --no-backfill switches this off.

WHAT IT DOES (every hour 10:00 ... 23:00 and 00:00)
  1. Downloads the OIX reports of the delivery day and saves the raw files in
        C:\\Users\\chipanl\\Downloads\\Whatsapp Session\\log-kpi-tracker\\Folder for KPI Dashboard
        "Overall Order in Cargroup (Nov-24).xlsx"      Route Resources Planning Report (all orders of the day)
        "Delivered Order in Cargroup (Nov-24).csv"     Waybill Status History RECEIVED_BY_CUSTOMER
        "Failed Order in Cargroup (Nov-24).csv"        Waybill Status History FAIL_TO_DELIVER
        "Carline Mapping (Nov-24).csv"                 Delivery Route Zone ID Report -> carline / car group (HKTV / ODS of routes
                                                       that have not finished anything yet)
     (the files of the day are overwritten every hour = cumulative; files older than RAW_KEEP_DAYS are purged).
     The Shop Code List (Tableau) is refreshed once a day (last good copy is the fallback).
     The Overall Order Report / carline map are SOFT: if they cannot be downloaded the finished table is still written
     (without the Outstanding row for that run).
  2. Builds the table "parent orders finished by hour and district" (HKTV / ODS) with the rules of
     carline_progress_monitor.py:
        * parent order        = column L, else column K (parent_from_oix)
        * HKTV / ODS          = car group (column W "Car Group Code"; blank -> the car group that the same route has in the file)
                                starting with ODS or VAN = ODS, everything else = HKTV   (ODS_PREFIXES)
        * district            = 送貨車號 (route) -> district, the same rule as the dashboard (kpi_pipeline.district_from_carline);
                                routes that map to no district (EDS / express WM*, Macau, shop routes ...) are not counted
        * finished            = RECEIVED_BY_CUSTOMER or FAIL_TO_DELIVER (FAILED_AS_FINISHED, as in the monitor)
        * NOT counted         = shop / outsource pick-up point orders: Delivery Zone in the Shop Code List, zone KN015,
                                waybill starting with R / U, route containing CASE / CK / 舖, JHC* users,
                                routes that are not a T+0 route, the parent orders of EXCLUDE_PARENT_ORDERS
  3. Counting rule (per HKTV / ODS separately)
        * one parent order finished inside an hour range            -> 1 in that hour
        * the same parent order finished separately in 2 hour ranges -> 1 in each hour (Finished Orders counts it twice)
        * ... but when the finish is really ONE delivery that only straddles two hour ranges (the finish records are
          not more than EPISODE_GAP_MIN minutes apart) -> counted ONCE, in the EARLIER hour
        * "Finished Orders" = sum of the hourly counts, "Unique Orders Count" = unique parent orders
          (Finished Orders - Unique Orders Count = "Duplicate Orders", delivered more than once)
        * "Outstanding Order Count" = unique parent orders of the Overall Order Report (same exclusions: cancelled / on-hold
          statuses, KN015 and shop zones, CASE / CK / 舖 routes, T+0 routes only, no district = not counted) that have no
          finish record in the same HKTV / ODS type. HKTV / ODS of an Overall order = car group of its route (Overall
          column if it has one, else the car group the route finished with, else the carline map; unknown = HKTV).
  4. Writes the result of the day into  public/order_finished_pattern.json  (every earlier day stays in the file, so the
     dashboard can show any date), back-fills last week's day when it is missing, re-embeds the dashboard (index.html) and
     runs the deploy hook (soft-fail).

USAGE
  python3 order_finished_pattern.py run                      # download + build (+ back-fill of date-7 when missing)
  python3 order_finished_pattern.py run --skip-download      # rebuild from the raw files already in the folder
  python3 order_finished_pattern.py run --date 2026-11-24    # a given delivery day
  python3 order_finished_pattern.py run --no-backfill        # do not back-fill last week's day
  python3 order_finished_pattern.py run --no-deploy          # no embed / git push
  python3 order_finished_pattern.py install-cron             # 10:00-23:00 hourly + 00:00 (= last run of the previous day)
"""
from __future__ import annotations

import argparse
import fcntl
import json
import logging
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

VERSION = "v1.1.1"
TZ = ZoneInfo("Asia/Hong_Kong")
HERE = Path(__file__).resolve().parent
# delay_forecast_cargroup_by_dp.py / delay_forecast_cargroup.py live in "Whatsapp Session" (one level above log-kpi-tracker),
# the same folder the carline monitor uses (WA_COMMON_DIR)
for _p in (os.environ.get("WA_COMMON_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session"), str(HERE.parent), str(HERE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)
LOG_DIR = Path(os.environ.get("FINISH_LOG_DIR", str(HERE / "logs")))

# ---- folders -------------------------------------------------------------------------------------------------------
# C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker\Folder for KPI Dashboard   (WSL path)
RAW_DIR = Path(os.environ.get("FINISH_RAW_DIR", "/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker/Folder for KPI Dashboard"))
JSON_PATH = Path(os.environ.get("ORDER_FINISHED_PATTERN_PATH", str(HERE / "public" / "order_finished_pattern.json")))
RAW_KEEP_DAYS = int(os.environ.get("FINISH_RAW_KEEP_DAYS", "7"))
HISTORY_KEEP_DAYS = int(os.environ.get("FINISH_HISTORY_KEEP_DAYS", "730"))
BACKFILL_OFFSET_DAYS = int(os.environ.get("FINISH_BACKFILL_OFFSET_DAYS", "7"))   # the comparison day = same weekday last week

# ---- rules (same names / values as carline_progress_monitor.py) --------------------------------------------------------
STATUSES = (("RECEIVED_BY_CUSTOMER", "delivered"), ("FAIL_TO_DELIVER", "failed"))
ODS_PREFIXES = ("ODS", "VAN")                       # car groups of the ODS table (VAN counts as ODS)
EXCLUDE_ZONES = ("KN015",)                          # outsource pick-up point zone (with or without "-n")
EXCLUDE_WAYBILL_PREFIXES = ("R", "U")               # pick-up waybills
EXCLUDE_ROUTE_KEYWORDS = ("CASE", "CK", "舖")       # route names of shop / case routes
EXCLUDE_USER_PREFIXES = ("JHC",)                    # shop staff
EXCLUDE_PARENT_ORDERS = frozenset({
    "H261010521832", "H261024547461", "H261026031108", "H261028538631", "H261029500651", "H261029518629", "H261029520928", "H261029536013",
    "H261030510655", "H261030517670", "H261030525047", "H261030529167", "H261030549928", "H261031513565", "H261031519160", "H261031523336",
    "H261031539679", "H261031552996", "H261031557785", "H261101503158", "H261101512541", "H261101516338"})
EPISODE_GAP_MIN = int(os.environ.get("FINISH_EPISODE_GAP_MIN", "15"))   # finish records of one parent <= this apart = ONE finish
DISTRICTS = ["ETH", "ETK", "ETX", "WTK", "WTH", "WTX", "NT-TW", "NT-ST", "NT-TM", "NT-TSM"]   # column order of the table
TYPES = ("HKTV", "ODS")
# Dedicated Chromium profile for OFP OIX downloads (not shared with express / monitor / forecast).
BROWSER_PROFILE_DIR = Path(os.environ.get(
    "FINISH_BROWSER_PROFILE_DIR", str(HERE / ".ofp-browser-profile")))
# Legacy names kept so older notes / env still parse; OFP no longer blocks on other scripts' locks.
WAIT_MAX_MIN, WAIT_POLL_S = 0, 5

log = logging.getLogger("order_finished_pattern")


# =====================================================================================================================
# SMALL HELPERS (copies of the rules in carline_progress_monitor / delay_forecast_cargroup / kpi_pipeline)
# =====================================================================================================================
def _s(x) -> str:
    return "" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x).strip()


def parent_from_oix(parent_no, order_no) -> str:
    """Column L first, else column K (delay_forecast_cargroup.parent_from_oix)."""
    p = _s(parent_no)
    if p.startswith("EM"):
        return p[1:]
    if p.startswith("H"):
        return p.split("-")[0]
    o = _s(order_no)
    if o.startswith("H"):
        return o[:13]
    if o.startswith("EM"):
        return o[1:]
    return ""


def cargroup_from_code(code) -> str:
    c = _s(code).upper().replace(" ", "")
    return c.split("-", 1)[1] if "-" in c else c


def district_from_carline(carline):
    """送貨車號 -> district (kpi_pipeline.district_from_carline = Domo_Jai.py IFS rule, NT-YT is already WTH)."""
    raw = _s(carline)
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


def is_ods(cg) -> bool:
    return re.sub(r"^\d{8}-", "", _s(cg)).upper().startswith(ODS_PREFIXES)


def mmm_dd(d: datetime) -> str:
    return f"{['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.month - 1]}-{d.day:02d}"


def raw_files(t0: datetime) -> dict[str, Path]:
    tag = mmm_dd(t0)
    return {"overall": RAW_DIR / f"Overall Order in Cargroup ({tag}).xlsx",
            "delivered": RAW_DIR / f"Delivered Order in Cargroup ({tag}).csv",
            "failed": RAW_DIR / f"Failed Order in Cargroup ({tag}).csv",
            "carline_map": RAW_DIR / f"Carline Mapping ({tag}).csv",
            "route_report": RAW_DIR / f"Delivery Route Zone ID Report ({tag}).xlsx",
            "shop_code": RAW_DIR / f"Shop Code List ({tag}).csv"}


SHOP_KEEP = "Shop Code List.csv"                    # permanent copy of the last good list


# =====================================================================================================================
# COMPUTE (no network)
# =====================================================================================================================
def load_events(path: Path, status: str) -> pd.DataFrame:
    """Waybill Status History CSV (banner row + header row) -> one row per status record."""
    raw = pd.read_csv(path, skiprows=1, dtype=str, encoding="utf-8-sig").fillna("")
    raw.columns = [str(c).strip() for c in raw.columns]
    need = ["Status History", "Status Change Date", "User", "Order Number", "Parent Order Number", "送貨車號"]
    miss = [c for c in need if c not in raw.columns]
    if miss:
        raise RuntimeError(f"{path.name}: missing columns {miss}; found {list(raw.columns)}")
    col = lambda *names: next((c for c in raw.columns if c.lower() in names), None)
    wb_c, zone_c, cg_c = col("waybill number", "waybill no"), col("delivery zone"), col("car group code")
    df = pd.DataFrame({
        "Parent": [parent_from_oix(p, o) for p, o in zip(raw["Parent Order Number"], raw["Order Number"])],
        "Status": raw["Status History"].str.split(",").str[0].str.strip().str.upper(),
        "ts": pd.to_datetime(raw["Status Change Date"], errors="coerce"),
        "User": raw["User"].str.strip().str.upper(),
        "Route": raw["送貨車號"].str.strip(),
        "Waybill": raw[wb_c].str.strip().str.upper() if wb_c else "",
        "Zone": raw[zone_c].str.strip().str.upper() if zone_c else "",
        "CG": raw[cg_c].map(cargroup_from_code) if cg_c else "",
    })
    return df[df["Status"] == status]


def _base():
    import delay_forecast_cargroup_by_dp as dp          # same folder; dp.base = delay_forecast_cargroup
    return dp.base


def load_overall_orders(path: Path, t0: datetime, shop_zones, audit: dict) -> pd.DataFrame:
    """Overall Order Report (Route Resources Planning, sheet "Route Resource Data") -> one row per (Parent, Route).

    Same filters as delay_forecast_cargroup.load_overall (cancelled / on-hold statuses, zone KN015, blank / CASE / CK / 舖 routes),
    plus the rules of the finished table: shop zones of the Shop Code List, R / U waybills (when the report has a waybill column),
    parent orders of EXCLUDE_PARENT_ORDERS, T+0 routes only, routes with no district."""
    base = _base()
    df = pd.read_excel(path, sheet_name="Route Resource Data", dtype=str).fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    need = ["District", "Route Name", "Order Id", "Zone ID", "Waybill Status", "Parent Order ID"]
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise RuntimeError(f"{path.name}: Route Resource Data is missing columns {miss}; found {list(df.columns)}")
    audit["ovRows"] = int(len(df))
    low = {c.lower().replace(" ", ""): c for c in df.columns}
    wb_c = next((low[k] for k in ("waybillnumber", "waybillno", "waybillid") if k in low), None)
    cg_c = next((low[k] for k in ("cargroupcode", "cargroup", "cargroupid") if k in low), None)

    route = df["Route Name"].str.strip()
    st = df["Waybill Status"].map(base.norm_waybill_status)
    keep = ~st.isin(base.EXCLUDE_WAYBILL_STATUSES) & ~st.str.contains("CANCEL", regex=False)
    zone = df["Zone ID"].str.strip().str.upper()
    keep &= zone != base.EXCLUDE_ZONE_ID
    keep &= route != ""
    for kw in base.EXCLUDE_ROUTE_KEYWORDS:
        keep &= ~route.str.contains(kw, regex=False)
    audit["ovStatusZoneRoute"] = int((~keep).sum())
    out = pd.DataFrame({
        "Parent": [base.parent_from_overall(p, o) for p, o in zip(df["Parent Order ID"], df["Order Id"])],
        "Route": route, "Zone": zone,
        "Waybill": df[wb_c].str.strip().str.upper() if wb_c else "",
        "CG": df[cg_c].map(cargroup_from_code) if cg_c else "",
    })[keep]

    def drop(mask, key):
        nonlocal out
        audit[key] = int(mask.sum())
        out = out[~mask]

    drop(out["Parent"] == "", "ovNoParent")
    drop(out["Parent"].isin(EXCLUDE_PARENT_ORDERS), "ovListedParents")
    shop = (out["Zone"].isin(shop_zones) if shop_zones else pd.Series(False, index=out.index)) | out["Zone"].str.split("-").str[0].isin(EXCLUDE_ZONES)
    drop(shop, "ovShopPickupZone")
    if wb_c:
        drop(out["Waybill"].str.startswith(EXCLUDE_WAYBILL_PREFIXES), "ovPickupWaybill")
    drop(~out["Route"].str.startswith(f"{t0:%y%m%d}"), "ovNotT0Route")
    out = out.copy()
    out["District"] = out["Route"].map(district_from_carline)
    drop(out["District"].isna(), "ovNoDistrict")
    out = out.drop_duplicates(["Parent", "Route"]).drop(columns=["Zone", "Waybill"]).reset_index(drop=True)
    audit["ovParents"] = int(out["Parent"].nunique())
    return out


def load_carmap(path: Path):
    """Carline Mapping CSV -> (by_district, global) lookup of delay_forecast_cargroup, or None."""
    if not path.exists():
        return None
    try:
        return _base().load_carline_map(str(path))
    except Exception as e:
        log.warning("carline map unreadable (%s: %s) - routes without a finish are treated as HKTV", type(e).__name__, e)
        return None


def compute_outstanding(overall: pd.DataFrame, finished: dict, route_cg: dict, carmap, audit: dict):
    """-> (scheduled, outstanding): {type: {district: n, "_total": n}} of unique parent orders of the Overall Order Report.
    finished = {type: set(parent orders with a finish record)}; a parent order is outstanding when it is not in the set of its type."""
    base = _base() if carmap else None
    ov = overall.copy()
    cg = ov["CG"].where(ov["CG"] != "", ov["Route"].map(route_cg)).fillna("")          # 1) Overall column  2) car group the route finished with
    if carmap:                                                                          # 3) carline map (Default car group of the route)
        by_dist, glob_map = carmap
        miss = cg == ""
        if miss.any():
            cg = cg.where(~miss, [base.map_route_to_cargroup(r, d, by_dist, glob_map) or "" for r, d in zip(ov["Route"], ov["District"])])
    audit["ovTypeUnknown"] = int((cg == "").sum())                                       # unknown = HKTV (same rule as the finished table)
    ov["Type"] = cg.map(lambda c: "ODS" if is_ods(c) else "HKTV")
    ov = ov.drop_duplicates(["Type", "District", "Parent"])
    ov["Open"] = pd.Series([p not in finished.get(t, ()) for t, p in zip(ov["Type"], ov["Parent"])], index=ov.index, dtype=bool)

    def tally(df):
        res = {t: {d: 0 for d in DISTRICTS} for t in TYPES}
        for (t, d), n in df.groupby(["Type", "District"])["Parent"].nunique().items():
            res[t][d] = int(n)
        for t in TYPES:
            res[t]["_total"] = int(df.loc[df["Type"] == t, "Parent"].nunique())
        return res

    return tally(ov), tally(ov[ov["Open"]])


def build_day(delivered: pd.DataFrame, failed: pd.DataFrame, shop_zones, t0: datetime, as_of: datetime,
              overall: pd.DataFrame | None = None, carmap=None, ov_audit: dict | None = None) -> dict:
    """-> the record of one delivery day (see header for the rules)."""
    ev = pd.concat([delivered, failed], ignore_index=True)
    audit = {"rows": int(len(ev))}
    ev = ev[ev["ts"].notna() & (ev["ts"] >= t0) & (ev["ts"] < t0 + timedelta(days=1)) & (ev["Parent"] != "")]
    audit["rowsInDay"] = int(len(ev))

    def drop(mask, key):
        nonlocal ev
        audit[key] = int(mask.sum())
        ev = ev[~mask]

    zone_id = ev["Zone"].str.split("-").str[0]
    drop(ev["Parent"].isin(EXCLUDE_PARENT_ORDERS), "listedParents")
    shop = (ev["Zone"].isin(shop_zones) if shop_zones else pd.Series(False, index=ev.index)) | zone_id.isin(EXCLUDE_ZONES)
    drop(shop, "shopPickupZone")                                                         # shop / outsource pick-up point
    drop(ev["Waybill"].str.startswith(EXCLUDE_WAYBILL_PREFIXES), "pickupWaybill")        # R* / U* pick-up waybills
    drop(ev["Route"].str.upper().str.contains("|".join(re.escape(k) for k in EXCLUDE_ROUTE_KEYWORDS), regex=True), "shopRoute")
    drop(ev["User"].str.startswith(EXCLUDE_USER_PREFIXES), "shopStaff")
    drop(~ev["Route"].str.startswith(f"{t0:%y%m%d}"), "notT0Route")                     # only T+0 routes are in the monitor
    ev = ev.copy()
    ev["District"] = ev["Route"].map(district_from_carline)
    drop(ev["District"].isna(), "noDistrict")                                            # EDS / express / Macau ... routes
    ev = ev.copy()

    if ev.empty:                                            # a day with nothing to count (e.g. a back-fill of a day without data)
        audit["cgFilledFromRoute"] = 0
        known = pd.Series(dtype=object)
        eps = pd.DataFrame({"Type": pd.Series(dtype=object), "Parent": pd.Series(dtype=object), "District": pd.Series(dtype=object),
                            "ts": pd.Series(dtype="datetime64[ns]"), "hour": pd.Series(dtype=int)})
    else:
        # car group: own column W, blank -> the (most frequent) car group of the same route; ODS / VAN prefix = ODS table
        known = ev[ev["CG"] != ""].groupby("Route")["CG"].agg(lambda s: s.mode().iat[0])
        audit["cgFilledFromRoute"] = int((ev["CG"] == "").sum())
        ev["CG"] = ev["CG"].where(ev["CG"] != "", ev["Route"].map(known)).fillna("")
        ev["Type"] = ev["CG"].map(lambda c: "ODS" if is_ods(c) else "HKTV")

        # one finish ("episode") = finish records of the same parent order (per type) that are <= EPISODE_GAP_MIN apart
        ev = ev.sort_values(["Type", "Parent", "ts"], kind="stable")
        gap = ev.groupby(["Type", "Parent"])["ts"].diff().dt.total_seconds().div(60)
        ev["ep"] = ((gap.isna()) | (gap > EPISODE_GAP_MIN)).cumsum()
        eps = ev.groupby("ep").agg(Type=("Type", "first"), Parent=("Parent", "first"), District=("District", "first"), ts=("ts", "first"))
        eps["hour"] = eps["ts"].dt.hour
        # district of a parent order = district of its earliest finish, so that the districts add up to the total
        first_d = eps.groupby(["Type", "Parent"])["District"].transform("first")
        eps["District"] = first_d

    counts = {t: {d: [0] * 24 for d in DISTRICTS} for t in TYPES}
    for (t, d, h), n in eps.groupby(["Type", "District", "hour"]).size().items():
        counts[t][d][int(h)] = int(n)
    distinct = {t: {d: 0 for d in DISTRICTS} for t in TYPES}
    for (t, d), n in eps.groupby(["Type", "District"])["Parent"].nunique().items():
        distinct[t][d] = int(n)
    for t in TYPES:
        distinct[t]["_total"] = int(eps.loc[eps["Type"] == t, "Parent"].nunique())
    audit["parents"] = int(eps["Parent"].nunique())
    audit["finishes"] = int(len(eps))
    rec = {"asOf": as_of.strftime("%Y-%m-%dT%H:%M"), "episodeGapMin": EPISODE_GAP_MIN, "counts": counts, "distinct": distinct}
    if overall is not None:                              # Outstanding Order Count needs the Overall Order Report
        finished = {t: set(eps.loc[eps["Type"] == t, "Parent"]) for t in TYPES}
        rec["scheduled"], rec["outstanding"] = compute_outstanding(overall, finished, known.to_dict(), carmap, ov_audit if ov_audit is not None else audit)
        if ov_audit:
            audit.update(ov_audit)
    rec["audit"] = audit
    return rec


# =====================================================================================================================
# HISTORY JSON
# =====================================================================================================================
def load_history() -> dict:
    try:
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            h = json.load(f)
        if isinstance(h.get("days"), dict):
            return h
    except (OSError, ValueError):
        pass
    return {"days": {}}


def save_day(date_str: str, rec: dict):
    """Adds / replaces ONE day; every earlier day stays in the file (atomic write)."""
    h = load_history()
    h["days"][date_str] = rec
    cutoff = (datetime.now(TZ) - timedelta(days=HISTORY_KEEP_DAYS)).strftime("%Y-%m-%d")
    h["days"] = {k: h["days"][k] for k in sorted(h["days"]) if k >= cutoff}
    h.update({"version": VERSION, "districts": DISTRICTS, "types": list(TYPES), "latestDate": max(h["days"]),
              "generatedAt": datetime.now(TZ).strftime("%Y-%m-%dT%H:%M")})
    JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = JSON_PATH.with_suffix(".json.tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(h, f, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, JSON_PATH)
    log.info("Wrote %s (%d day(s) kept, latest %s)", JSON_PATH, len(h["days"]), h["latestDate"])


# =====================================================================================================================
# DOWNLOAD (OIX = carline monitor / delay forecast functions, Shop Code List = Tableau, once a day)
# =====================================================================================================================
def _ofp_new_page(p, headed: bool = False):
    """Launch Chromium with OFP's own user-data dir (isolated from express / monitor / forecast).

    Returns (context, page, holder) shaped like delay_forecast_cargroup._new_page so the shared
    download_* helpers keep working. Close with context.close().
    """
    BROWSER_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    headless = not headed
    log.info("OIX browser profile: %s (headless=%s)", BROWSER_PROFILE_DIR, headless)
    ctx = p.chromium.launch_persistent_context(
        user_data_dir=str(BROWSER_PROFILE_DIR),
        headless=headless,
        accept_downloads=True,
        viewport={"width": 1400, "height": 950},
        slow_mo=0 if headless else 150,
        args=["--disable-popup-blocking"],
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    holder = {"_ctx": ctx}

    def _watch(pg):
        pg.on("dialog", lambda d: (log.info("  [dialog] %s", (d.message or "")[:120]), d.accept()))
        pg.on("download", lambda d: (log.info("  [download event] %s", d.suggested_filename), holder.__setitem__("dl", d)))
        pg.on("response", lambda r: (r.request.method == "POST" or r.request.resource_type == "document") and log.info(
            "  [net] %s %s %s | type=%s | disp=%s",
            r.status, r.request.method, r.url[:100],
            r.headers.get("content-type", "-"),
            (r.headers.get("content-disposition", "-") or "-")[:60]))
        pg.on("requestfailed", lambda r: log.info("  [net] FAILED %s | %s", r.url[:100], r.failure))

    _watch(page)
    ctx.on("page", lambda pg: (log.info("  [new tab/popup] %s", (pg.url or "")[:100]), _watch(pg)))
    return ctx, page, holder


def download_raw(jobs: list[dict], headed: bool = False):
    """One OIX session for one or more delivery days. jobs = [{"t0": datetime, "files": raw_files(t0)}].
    Per day:  Overall Order Report (soft) -> carline map (soft, once per day) -> RECEIVED_BY_CUSTOMER + FAIL_TO_DELIVER (must succeed).
    The Overall download is delay_forecast_cargroup.download_overall (3 attempts, size / row-count check, last-good fallback).
    Uses OFP's own Chromium profile so it can run in parallel with express / monitor / forecast."""
    import delay_forecast_cargroup_by_dp as dp          # same folder
    base = dp.base
    from playwright.sync_api import sync_playwright
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    base.HEADLESS = not headed
    state = [{"done": set(), "tmp": {k: str(j["files"][k]) + ".part" for k in ("delivered", "failed")}} for j in jobs]
    with sync_playwright() as p:
        for attempt in range(1, base.DOWNLOAD_RETRIES + 1):
            browser = None
            try:
                browser, page, holder = _ofp_new_page(p, headed=headed)
                base._oix_login(page)
                for job, st in zip(jobs, state):
                    t0, files, done = job["t0"], job["files"], st["done"]
                    if "overall" not in done:
                        try:
                            base.download_overall(page, holder, t0, str(files["overall"]))
                            done.add("overall")
                        except Exception as e:                        # soft: the finished table does not need it
                            log.warning("%s: Overall Order Report download failed (%s: %s) - Outstanding Order Count may be missing",
                                        t0.date(), type(e).__name__, e)
                    if "map" not in done and not files["carline_map"].exists():
                        try:
                            base.download_route_zone_report(page, holder, t0, str(files["route_report"]))
                            rows = base.route_report_to_map_rows(str(files["route_report"]), t0)
                            if not rows:
                                raise RuntimeError("carline mapping returned no rows")
                            pd.DataFrame(rows).to_csv(files["carline_map"], index=False, encoding="utf-8-sig")
                            done.add("map")
                        except Exception as e:                        # soft: only used for routes that have not finished anything yet
                            log.warning("%s: carline map download failed (%s: %s)", t0.date(), type(e).__name__, e)
                    for status, key in STATUSES:
                        if key not in done:
                            base.download_waybill_history(page, holder, t0, status, st["tmp"][key])
                            done.add(key)
                break
            except Exception as e:
                log.error("OIX download attempt %d/%d failed: %s: %s", attempt, base.DOWNLOAD_RETRIES, type(e).__name__, e)
                if attempt == base.DOWNLOAD_RETRIES:
                    raise
                time.sleep(30)
            finally:
                try:
                    if browser:
                        browser.close()
                except Exception:
                    pass
    for job, st in zip(jobs, state):                     # replace the day files only after BOTH status downloads succeeded
        for key in ("delivered", "failed"):
            os.replace(st["tmp"][key], job["files"][key])
            log.info("saved %s", job["files"][key])


def report_times(files: dict[str, Path]) -> dict[str, datetime]:
    """Time each raw report was downloaded = save time of the file (HK time, naive)."""
    out = {}
    for k in ("overall", "delivered", "failed"):
        if files[k].exists():
            out[k] = datetime.fromtimestamp(files[k].stat().st_mtime, TZ).replace(tzinfo=None, second=0, microsecond=0)
    return out


def get_shop_zones(files: dict[str, Path], skip_download: bool):
    """Shop Code List: downloaded once a day (Tableau); falls back to the last good copy. None = rule cannot be applied."""
    import delay_forecast_cargroup_by_dp as dp
    keep, path = RAW_DIR / SHOP_KEEP, files["shop_code"]
    if not skip_download and not path.exists():
        try:
            if not dp.download_shop_code_list(str(path)):
                log.warning("Shop Code List download failed")
        except Exception as e:
            log.warning("Shop Code List download failed: %s", e)
    if not path.exists() and keep.exists():
        log.info("using the last saved Shop Code List")
        path = keep
    if not path.exists():
        log.warning("no Shop Code List - shop zones are NOT checked (KN015 / R,U waybills / shop routes still are)")
        return None
    try:
        codes = dp.load_shop_codes(str(path))
    except Exception as e:
        log.warning("Shop Code List unreadable (%s)", e)
        return None
    if path != keep:
        try:
            import shutil
            shutil.copyfile(path, keep)
        except OSError:
            pass
    return codes


def purge_old_raw():
    cutoff = time.time() - RAW_KEEP_DAYS * 86400
    for p in list(RAW_DIR.glob("Delivered Order in Cargroup (*).csv")) + list(RAW_DIR.glob("Failed Order in Cargroup (*).csv")) \
            + list(RAW_DIR.glob("Shop Code List (*).csv")) + list(RAW_DIR.glob("*.part")) \
            + list(RAW_DIR.glob("Overall Order in Cargroup (*).xlsx")) + list(RAW_DIR.glob("Overall Order in Cargroup (*).last_good.*")) \
            + list(RAW_DIR.glob("Carline Mapping (*).csv")) + list(RAW_DIR.glob("Delivery Route Zone ID Report (*).xlsx")):
        try:
            if p.stat().st_mtime < cutoff or p.name.endswith(".part"):
                p.unlink()
        except OSError:
            pass


# =====================================================================================================================
# RUN
# =====================================================================================================================
def build_record(t0: datetime, files: dict[str, Path], shop, now: datetime) -> dict:
    """Raw files of one delivery day -> the record of that day. asOf = latest download time of the reports (capped at 24:00 of the day)."""
    times = report_times(files)
    end = t0 + timedelta(days=1)
    as_of = min(max(times.values()) if times else now, end)
    overall, carmap, ov_audit = None, None, {}
    if files["overall"].exists():
        try:
            overall = load_overall_orders(files["overall"], t0, shop, ov_audit)
            carmap = load_carmap(files["carline_map"])
        except Exception as e:                                       # soft: no Outstanding row for this record
            log.warning("Overall Order Report unreadable (%s: %s) - no Outstanding Order Count", type(e).__name__, e)
            overall = None
    else:
        log.warning("no Overall Order Report for %s - no Outstanding Order Count", t0.date())
    day = build_day(load_events(files["delivered"], "RECEIVED_BY_CUSTOMER"), load_events(files["failed"], "FAIL_TO_DELIVER"),
                    shop, t0, as_of, overall, carmap, ov_audit)
    day["reportTimes"] = {k: v.strftime("%Y-%m-%dT%H:%M") for k, v in times.items()}
    day["final"] = bool(as_of >= end)
    a = day["audit"]
    log.info("%s: finished records %d -> %d parent orders, %d finishes | left out: shop/pick-up zone %d, R/U waybill %d, shop route %d, JHC %d, "
             "not T+0 route %d, no district %d", t0.date(), a["rowsInDay"], a["parents"], a["finishes"], a["shopPickupZone"], a["pickupWaybill"],
             a["shopRoute"], a["shopStaff"], a["notT0Route"], a["noDistrict"])
    for t in TYPES:
        sh = sum(sum(v) for v in day["counts"][t].values())
        un = day["distinct"][t]["_total"]
        o = day.get("outstanding", {}).get(t, {}).get("_total")
        log.info("  %s: Finished Orders %d | Unique Orders Count %d | Duplicate Orders %d | Outstanding Order Count %s",
                 t, sh, un, sh - un, "n/a" if o is None else o)
    if "ovParents" in a:
        log.info("  Overall Order Report: %d rows -> %d parent orders (orders whose type could not be resolved, counted as HKTV: %d rows)", a["ovRows"], a["ovParents"], a.get("ovTypeUnknown", 0))
    return day


def _is_final(date_str: str, rec: dict) -> bool:
    end = (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%dT%H:%M")
    return bool(rec.get("final")) or str(rec.get("asOf", "")) >= end


def backfill_reason(date_str: str) -> str | None:
    """Why the day has to be (re)built from OIX, or None when it is complete in the dashboard data."""
    rec = load_history()["days"].get(date_str)
    if rec is None:
        return "missing in the dashboard data"
    if not _is_final(date_str, rec):
        return "not final (day-end run missing)"
    if "outstanding" not in rec and not rec.get("backfilled"):
        return "no Outstanding Order Count yet"
    return None


def run_backfill(t0: datetime, shop, skip_download: bool, headed: bool) -> bool:
    """Back-fills the comparison day (t0 - BACKFILL_OFFSET_DAYS) when it is missing. Never raises."""
    tb = t0 - timedelta(days=BACKFILL_OFFSET_DAYS)
    key = tb.strftime("%Y-%m-%d")
    why = backfill_reason(key)
    if why is None:
        log.info("back-fill: %s is complete in the dashboard data", key)
        return False
    log.info("back-fill: %s %s - building it from OIX", key, why)
    try:
        files = raw_files(tb)
        if skip_download:
            if not (files["delivered"].exists() and files["failed"].exists()):
                log.info("back-fill: --skip-download and no raw files for %s - skipped", key)
                return False
        else:
            download_raw([{"t0": tb, "files": files}], headed)
        day = build_record(tb, files, shop, datetime.now(TZ).replace(tzinfo=None))
        day["backfilled"] = True
        save_day(key, day)
        return True
    except Exception as e:                                # today's record is already saved - back-fill is retried at the next run
        log.warning("back-fill of %s failed: %s: %s", key, type(e).__name__, e)
        return False


def run_job(date: str | None, skip_download: bool, headed: bool, deploy: bool, backfill: bool = True):
    now = datetime.now(TZ).replace(tzinfo=None)
    today0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
    t0 = datetime.strptime(date, "%Y-%m-%d") if date else (today0 - timedelta(days=1) if now.hour < 6 else today0)   # 00:00 run = the day that just ended
    files = raw_files(t0)
    log.info("=== Order Finished Pattern %s | T+0 %s ===", VERSION, t0.date())
    if not skip_download:
        # Own Chromium profile: start immediately; do not wait on express / monitor / forecast locks.
        download_raw([{"t0": t0, "files": files}], headed)
    for k in ("delivered", "failed"):
        if not files[k].exists():
            raise FileNotFoundError(files[k])
    shop = get_shop_zones(files, skip_download)
    day = build_record(t0, files, shop, now)
    log.info("data as of %s (report download time)", day["asOf"].replace("T", " "))
    save_day(t0.strftime("%Y-%m-%d"), day)
    if not skip_download:
        purge_old_raw()
    if backfill:
        run_backfill(t0, shop, skip_download, headed)
    if deploy:
        try:
            sys.path.insert(0, str(HERE))
            import kpi_pipeline as kp
            kp.update_embedded_data()
            kp.run_deploy_hook()
        except Exception as e:                           # the JSON is already saved; embed / push is soft-fail
            log.warning("embed / deploy skipped: %s: %s", type(e).__name__, e)


def install_cron():
    py, script = sys.executable, Path(__file__).resolve()
    marker = "# OIX_FinishPattern"
    cur = subprocess.run(["crontab", "-l"], capture_output=True, text=True).stdout
    keep = [l for l in cur.splitlines() if marker not in l]
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    for tag, expr in (("day", "0 10-23 * * *"), ("dayend", "0 0 * * *")):
        keep.append(f'{expr} cd "{script.parent}" && "{py}" "{script}" run >> "{LOG_DIR}/order_finished_pattern_{tag}.log" 2>&1 {marker}_{tag}')
    subprocess.run(["crontab", "-"], input="\n".join(keep) + "\n", text=True, check=True)
    print("crontab updated: hourly 10:00-23:00 + 00:00 (last run of the previous day)")


def main():
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        handlers=[logging.FileHandler(LOG_DIR / "order_finished_pattern.log", encoding="utf-8"), logging.StreamHandler()])
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("run", help="download + build once")
    r.add_argument("--date", help="delivery date T+0 (YYYY-MM-DD); default today (00:00-05:59 = yesterday)")
    r.add_argument("--skip-download", action="store_true", help="use the raw files already in the folder")
    r.add_argument("--headed", action="store_true", help="show the browser")
    r.add_argument("--no-deploy", action="store_true", help="write the JSON only (no embed / git push)")
    r.add_argument("--no-backfill", action="store_true", help="do not back-fill the same weekday of last week when it is missing")
    sub.add_parser("install-cron", help="register the hourly jobs")
    a = ap.parse_args()
    if a.cmd == "run":
        lock = open("/tmp/order_finished_pattern.lock", "a")
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            log.warning("another Order Finished Pattern run is still active - skipped")
            return
        run_job(a.date, a.skip_download, a.headed, not a.no_deploy, not a.no_backfill)
    else:
        install_cron()


if __name__ == "__main__":
    main()
