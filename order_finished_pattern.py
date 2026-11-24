#!/usr/bin/env python3
"""
Order Finished Pattern   v1.0.0 (WSL / Linux)      feeds the dashboard tab "Order Finished Pattern"
===================================================================================================
Keep this file in the SAME folder as carline_progress_monitor.py / delay_forecast_cargroup_by_dp.py /
delay_forecast_cargroup.py / kpi_pipeline.py (log-kpi-tracker).

WHAT IT DOES (every hour 10:00 ... 23:00 and 00:00)
  1. Downloads the OIX "Waybill Status History" of the delivery day for TWO statuses only
        RECEIVED_BY_CUSTOMER  and  FAIL_TO_DELIVER
     (same OIX functions as the carline monitor) and saves the raw CSVs in
        C:\\Users\\chipanl\\Downloads\\Whatsapp Session\\log-kpi-tracker\\Folder for KPI Dashboard
     as  "Delivered Order in Cargroup (Nov-24).csv" / "Failed Order in Cargroup (Nov-24).csv"
     (the file of the day is overwritten every hour = cumulative; files older than RAW_KEEP_DAYS are purged).
     The Shop Code List (Tableau) is refreshed once a day (last good copy is the fallback).
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
        * the same parent order finished separately in 2 hour ranges -> 1 in each hour (Sum of hours counts it twice)
        * ... but when the finish is really ONE delivery that only straddles two hour ranges (the finish records are
          not more than EPISODE_GAP_MIN minutes apart) -> counted ONCE, in the EARLIER hour
        * "Sum of hours" = sum of the hourly counts, "Daily distinct" = unique parent orders
          (Sum of hours - Daily distinct = orders that were delivered more than once)
  4. Writes the result of the day into  public/order_finished_pattern.json  (every earlier day stays in the file, so the
     dashboard can show any date), re-embeds the dashboard (index.html) and runs the deploy hook (soft-fail).

USAGE
  python3 order_finished_pattern.py run                      # download + build (what cron runs)
  python3 order_finished_pattern.py run --skip-download      # rebuild from the raw files already in the folder
  python3 order_finished_pattern.py run --date 2026-11-24    # a given delivery day
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

VERSION = "v1.0.1"
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
EPISODE_GAP_MIN = int(os.environ.get("FINISH_EPISODE_GAP_MIN", "30"))   # finish records of one parent <= this apart = ONE finish
DISTRICTS = ["ETH", "ETK", "ETX", "WTK", "WTH", "WTX", "NT-TW", "NT-ST", "NT-TM", "NT-TSM"]   # column order of the table
TYPES = ("HKTV", "ODS")
WAIT_MAX_MIN, WAIT_POLL_S = 20, 20

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
    return {"delivered": RAW_DIR / f"Delivered Order in Cargroup ({tag}).csv",
            "failed": RAW_DIR / f"Failed Order in Cargroup ({tag}).csv",
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


def build_day(delivered: pd.DataFrame, failed: pd.DataFrame, shop_zones, t0: datetime, as_of: datetime) -> dict:
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
    return {"asOf": as_of.strftime("%Y-%m-%dT%H:%M"), "episodeGapMin": EPISODE_GAP_MIN, "counts": counts, "distinct": distinct, "audit": audit}


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
# DOWNLOAD (OIX = carline monitor functions, Shop Code List = Tableau, once a day)
# =====================================================================================================================
def _wait_for_other_jobs():
    deadline = time.time() + WAIT_MAX_MIN * 60
    for path in ("/tmp/carline_progress_monitor.lock", "/tmp/delay_forecast_cargroup_by_dp.lock", "/tmp/delay_forecast_cargroup.lock"):
        while True:
            try:
                f = open(path, "a")
            except OSError:
                break
            try:
                fcntl.flock(f, fcntl.LOCK_EX | fcntl.LOCK_NB)
                fcntl.flock(f, fcntl.LOCK_UN)
                break
            except OSError:
                pass
            finally:
                f.close()
            if time.time() >= deadline:
                log.warning("still locked after %d min: %s - starting anyway", WAIT_MAX_MIN, path)
                break
            time.sleep(WAIT_POLL_S)


def download_raw(t0: datetime, files: dict[str, Path], headed: bool = False):
    import delay_forecast_cargroup_by_dp as dp          # same folder
    base = dp.base
    from playwright.sync_api import sync_playwright
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    base.HEADLESS = not headed
    tmp = {k: str(files[k]) + ".part" for k in ("delivered", "failed")}
    done = set()
    with sync_playwright() as p:
        for attempt in range(1, base.DOWNLOAD_RETRIES + 1):
            browser = None
            try:
                browser, page, holder = base._new_page(p)
                base._oix_login(page)
                for status, key in STATUSES:
                    if key not in done:
                        base.download_waybill_history(page, holder, t0, status, tmp[key])
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
    for key in ("delivered", "failed"):                  # replace the day file only after BOTH downloads succeeded
        os.replace(tmp[key], files[key])
        log.info("saved %s", files[key])


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
            + list(RAW_DIR.glob("Shop Code List (*).csv")) + list(RAW_DIR.glob("*.part")):
        try:
            if p.stat().st_mtime < cutoff or p.name.endswith(".part"):
                p.unlink()
        except OSError:
            pass


# =====================================================================================================================
# RUN
# =====================================================================================================================
def run_job(date: str | None, skip_download: bool, headed: bool, deploy: bool):
    now = datetime.now(TZ).replace(tzinfo=None)
    today0 = now.replace(hour=0, minute=0, second=0, microsecond=0)
    t0 = datetime.strptime(date, "%Y-%m-%d") if date else (today0 - timedelta(days=1) if now.hour < 6 else today0)   # 00:00 run = the day that just ended
    as_of = min(now, t0 + timedelta(days=1))
    files = raw_files(t0)
    log.info("=== Order Finished Pattern %s | T+0 %s | as of %s ===", VERSION, t0.date(), as_of.strftime("%H:%M"))
    if not skip_download:
        _wait_for_other_jobs()
        download_raw(t0, files, headed)
    for k in ("delivered", "failed"):
        if not files[k].exists():
            raise FileNotFoundError(files[k])
    shop = get_shop_zones(files, skip_download)
    day = build_day(load_events(files["delivered"], "RECEIVED_BY_CUSTOMER"), load_events(files["failed"], "FAIL_TO_DELIVER"), shop, t0, as_of)
    a = day["audit"]
    log.info("finished records %d -> %d parent orders, %d finishes | left out: shop/pick-up zone %d, R/U waybill %d, shop route %d, JHC %d, "
             "not T+0 route %d, no district %d", a["rowsInDay"], a["parents"], a["finishes"], a["shopPickupZone"], a["pickupWaybill"],
             a["shopRoute"], a["shopStaff"], a["notT0Route"], a["noDistrict"])
    for t in TYPES:
        sh = sum(sum(v) for v in day["counts"][t].values())
        log.info("  %s: Sum of hours %d | Daily distinct %d", t, sh, day["distinct"][t]["_total"])
    save_day(t0.strftime("%Y-%m-%d"), day)
    if not skip_download:
        purge_old_raw()
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
    sub.add_parser("install-cron", help="register the hourly jobs")
    a = ap.parse_args()
    if a.cmd == "run":
        lock = open("/tmp/order_finished_pattern.lock", "a")
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError:
            log.warning("another Order Finished Pattern run is still active - skipped")
            return
        run_job(a.date, a.skip_download, a.headed, not a.no_deploy)
    else:
        install_cron()


if __name__ == "__main__":
    main()
