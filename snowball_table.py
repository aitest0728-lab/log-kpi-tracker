#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build and refresh public/snowball_table.json for the Snowball Table dashboard tab.

November 2026 onward uses the unique-parent exclude_eds_v1 method:
  - completed delivery dates: OIX Delivered/Failed unique-parent raw CSVs
    (Status Change window D 00:00 through D+1 06:00; FTD is not in Total)
  - open / future delivery dates: unique parents from stored RRP workbooks
    (no new OIX login or download)
  - 2026-11-26 stays zeros (known no-delivery day)

Order-date rows before 2026-11-01 are copied unchanged from the previous
snowball_table.json (or rolling_orders.json on first build).

This module is imported by kpi_pipeline.py --section rolling and also
runnable as: python3 snowball_table.py --build
"""
from __future__ import annotations

import csv
import json
import os
import re
import sys
import traceback
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

HK = ZoneInfo("Asia/Hong_Kong")
METHOD_START = date(2026, 11, 1)
METHOD_VERSION = "exclude_eds_v1"
NO_DELIVERY_DATES = {date(2026, 11, 26)}
DISTRICTS = [
    "ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX",
]
DISTRICT_ALIAS = {"NT-YT": "WTH"}
VALID_HOME = set(DISTRICTS) | {"NT-YT"}
EXCLUDE_ZONE_ID = "KN015"
EXCLUDE_ROUTE_KEYWORDS = ["CK", "舖", "CASE"]
EXCLUDE_WAYBILL_STATUSES = frozenset({
    "TRANSITED_TO_HK", "HOLD_BY_CS", "PACKAGE_DAMAGED", "PACKAGE_DAMAGE", "PACKAGE_LOST",
})
CANCEL_STATUS_KEYWORD = "CANCEL"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def log(msg: str) -> None:
    print(f"[snowball {datetime.now(HK):%H:%M:%S}] {msg}", flush=True)


def today_hk() -> date:
    return datetime.now(HK).date()


def now_hk() -> datetime:
    return datetime.now(HK)


# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

def _env_path(name: str, default: Path) -> Path:
    value = os.environ.get(name)
    return Path(value) if value else default


def detect_tracker_dir() -> Path:
    env = os.environ.get("TRACKER_DIR")
    if env:
        return Path(env)
    here = Path(__file__).resolve().parent
    if (here / "kpi_pipeline.py").is_file():
        return here
    cand = Path("/mnt/c/Users/chipanl/Downloads/Whatsapp Session/log-kpi-tracker")
    if (cand / "kpi_pipeline.py").is_file():
        return cand
    return here


def detect_rolling_project() -> Path:
    env = os.environ.get("ROLLING_PROJECT_DIR")
    if env:
        return Path(env)
    cand = Path("/mnt/c/Users/chipanl/Downloads/wsl-hourly-rrp")
    if cand.is_dir():
        return cand
    box = Path("/workspace/snowball-work/build")
    if box.is_dir():
        return box
    return Path(__file__).resolve().parent


def tracker_dir() -> Path:
    return detect_tracker_dir()


def rolling_project() -> Path:
    return detect_rolling_project()


def public_dir() -> Path:
    return _env_path("PUBLIC_DIR", tracker_dir() / "public")


def snowball_path() -> Path:
    return _env_path("SNOWBALL_TABLE_PATH", public_dir() / "snowball_table.json")


def base_json_path() -> Path:
    env = os.environ.get("SNOWBALL_BASE_JSON")
    if env:
        return Path(env)
    p = snowball_path()
    if p.is_file():
        return p
    return public_dir() / "rolling_orders.json"


def raw_dir() -> Path:
    env = os.environ.get("SNOWBALL_RAW_DIR")
    if env:
        return Path(env)
    return _env_path("CACHE_DIR", rolling_project() / "cache") / "raw"


def rrp_root() -> Path:
    env = os.environ.get("SNOWBALL_RRP_ROOT") or os.environ.get("RRP_ROOT")
    if env:
        return Path(env)
    proj = rolling_project()
    for cand in (proj / "rrp-data", Path("/home/box/oix-reports/route-resource-hourly")):
        if cand.is_dir():
            return cand
    return proj / "rrp-data"


def rrp_cache_path() -> Path:
    return _env_path("CACHE_DIR", rolling_project() / "cache") / "snowball_rrp_by_od.json"


def shop_path() -> Path:
    env = os.environ.get("SHOP_PATH")
    if env:
        return Path(env)
    for cand in (
        rrp_root() / "shop-code-list.csv",
        rolling_project() / "shop-code-list.csv",
        Path("/workspace/snowball-work/build/shop-code-list.csv"),
    ):
        if cand.is_file():
            return cand
    return rolling_project() / "shop-code-list.csv"


def oix_reports_dir() -> Path:
    return _env_path("OIX_REPORTS_DIR", rolling_project() / "oix-reports")


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------

def clean_str(v) -> str:
    if v is None:
        return ""
    return str(v).strip()


def norm_waybill_status(s) -> str:
    return re.sub(r"\s+", "_", clean_str(s).upper())


def parent_from_overall(parent_id, order_id) -> str:
    p = clean_str(parent_id)
    if p.startswith("EM"):
        return p[1:]
    if p.startswith("H"):
        return p.split("-")[0]
    o = clean_str(order_id)
    if o.startswith("H"):
        return o.split("-")[0]
    if o.startswith("EM"):
        return o[1:]
    return ""


def order_date_from_h_em(s) -> date | None:
    s = clean_str(s).upper()
    if not s:
        return None
    if s.startswith("EM"):
        body = s[2:]
    elif s.startswith("H"):
        body = s[1:]
    else:
        return None
    if len(body) < 6 or not body[:6].isdigit():
        return None
    yy, mm, dd = int(body[:2]), int(body[2:4]), int(body[4:6])
    try:
        return date(2000 + yy, mm, dd)
    except ValueError:
        return None


def order_date_from_row(order_id, parent_order_id) -> date | None:
    return order_date_from_h_em(order_id) or order_date_from_h_em(parent_order_id)


def to_date(v) -> date | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = clean_str(v)
    if not s:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        pass
    for fmt in ("%Y%m%d", "%Y/%m/%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def display_district(raw: str) -> str | None:
    d = clean_str(raw).upper()
    if not d:
        return None
    d = DISTRICT_ALIAS.get(d, d)
    if d in DISTRICTS:
        return d
    return None


def overall_name(d: date) -> str:
    return f"Overall Order in Cargroup ({MONTHS[d.month - 1]}-{d.day:02d}).xlsx"


def load_shop_zones(path: Path | None = None) -> set[str]:
    path = Path(path or shop_path())
    if not path.is_file():
        log(f"WARNING: shop code list not found at {path}")
        return set()
    raw = path.read_bytes()
    text = None
    for enc in ("utf-16", "utf-16-le", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            continue
    if text is None:
        return set()
    import io
    header = text.splitlines()[0] if text.splitlines() else ""
    delim = "\t" if "\t" in header else ","
    reader = csv.DictReader(io.StringIO(text), delimiter=delim)
    fields = reader.fieldnames or []
    col = None
    for name in fields:
        key = (name or "").strip().lower().replace(" ", "_")
        if key in {"delivery_zone_id", "delivery_zone", "zone"}:
            col = name
            break
    zones: set[str] = set()
    if col is None:
        return zones
    for row in reader:
        z = (row.get(col) or "").strip().upper()
        if z:
            zones.add(z)
    return zones


def is_shop_or_pickup(zone: str, shop_zones: set[str]) -> bool:
    z = clean_str(zone).upper()
    if not z:
        return False
    if z == EXCLUDE_ZONE_ID or z.startswith(EXCLUDE_ZONE_ID):
        return True
    if z in shop_zones:
        return True
    # KN015* already handled; some lists store variants
    if z.split("-")[0] == EXCLUDE_ZONE_ID:
        return True
    return False


def assign_cross_district(parent_districts, parent_row_counts):
    assigned = {}
    n_cross = 0
    index = {d: i for i, d in enumerate(DISTRICTS)}
    for parent, dists in parent_districts.items():
        live = [d for d in dists if d in index]
        if not live:
            continue
        if len(live) > 1:
            n_cross += 1
            counts = parent_row_counts[parent]
            live.sort(key=lambda d: (-counts[d], index[d]))
        assigned[parent] = live[0]
    return assigned, n_cross


def empty_cell_map():
    return {k: [0, 0] for k in ["OVERALL", *DISTRICTS]}


def add_parent_to_cells(store, order_date: str, district: str, n_waybills: int) -> None:
    block = store.setdefault(order_date, empty_cell_map())
    block[district][0] += 1
    block[district][1] += int(n_waybills or 0)
    block["OVERALL"][0] += 1
    block["OVERALL"][1] += int(n_waybills or 0)


def prune_zero_cells(cells: dict) -> dict:
    out = {}
    for dd, blk in cells.items():
        slim = {}
        for k, v in blk.items():
            if v and (v[0] or v[1]):
                slim[k] = [int(v[0]), int(v[1])]
        if slim:
            # always keep OVERALL if any district is present
            if "OVERALL" not in slim:
                o, w = 0, 0
                for k, v in slim.items():
                    if k != "OVERALL":
                        o += v[0]
                        w += v[1]
                slim = {"OVERALL": [o, w], **slim}
            # include zero districts? original data included all 10 districts even when 0
            # Keep only districts that have values, plus OVERALL. The UI treats missing as NA.
            out[dd] = slim
    return out


# ---------------------------------------------------------------------------
# Raw unique-parent CSVs (completed delivery dates)
# ---------------------------------------------------------------------------

def parents_csv(day: date) -> Path:
    return raw_dir() / f"oix_parents_{day.isoformat()}.csv"


def raw_ready(day: date) -> bool:
    p = parents_csv(day)
    return p.is_file() and p.stat().st_size > 0


def read_raw_parents(day: date) -> list[dict]:
    path = parents_csv(day)
    if not path.is_file():
        return []
    with path.open(encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def cells_from_raw(day: date) -> tuple[dict, dict]:
    """Return (by_order_date cells, stats) for one completed delivery date.

    FTD rows are not in the parents CSV and are not added to Total.
    """
    rows = read_raw_parents(day)
    by_od = {}
    unknown = 0
    before_method = 0
    for rec in rows:
        od = (rec.get("created_date") or "").strip()[:10]
        dist = DISTRICT_ALIAS.get((rec.get("home_district") or "").strip().upper(),
                                 (rec.get("home_district") or "").strip().upper())
        if dist not in DISTRICTS:
            continue
        try:
            nwb = int(float(rec.get("n_waybills") or 0))
        except (TypeError, ValueError):
            nwb = 0
        if not od or od.startswith("19") or len(od) != 10:
            unknown += 1
            continue
        try:
            od_d = date.fromisoformat(od)
        except ValueError:
            unknown += 1
            continue
        if od_d < METHOD_START:
            before_method += 1
        add_parent_to_cells(by_od, od, dist, nwb)
    stats = {
        "delivery": day.isoformat(),
        "parents": len(rows),
        "unknown_created": unknown,
        "before_method_start": before_method,
        "method_rows": len(rows) - unknown,
    }
    return by_od, stats


_EXPORT_UNAVAILABLE = False


def try_export_raw(day: date, ctx=None, force: bool = False):
    """Build a raw CSV from OIX reports already on disk. Never starts a download."""
    global _EXPORT_UNAVAILABLE
    if _EXPORT_UNAVAILABLE:
        return None, ctx
    try:
        proj = rolling_project()
        if str(proj) not in sys.path:
            sys.path.insert(0, str(proj))
        os.environ.setdefault("OIX_REPORTS_DIR", str(oix_reports_dir()))
        os.environ.setdefault("CACHE_DIR", str(_env_path("CACHE_DIR", rolling_project() / "cache")))
        os.environ.setdefault("SHOP_PATH", str(shop_path()))
        import historical_raw as raw
        return raw.export_day(day, ctx=ctx, force=force), ctx
    except Exception as exc:
        log(f"WARNING: could not export raw unique parents for {day}: {type(exc).__name__}: {exc}")
        name = type(exc).__name__ + " " + str(exc)
        if "numpy" in name.lower() or "pandas" in name.lower() or "ImportError" in type(exc).__name__:
            _EXPORT_UNAVAILABLE = True
            log("Further raw exports in this run are skipped (OIX helper import is unavailable)")
        return None, ctx


def report_present(day: date) -> bool:
    tag = f"{MONTHS[day.month - 1]}-{day.day:02d}"
    path = oix_reports_dir() / day.isoformat() / f"Delivered Order in Cargroup ({tag}).csv"
    return path.is_file() and path.stat().st_size > 0


def d1_report_present(day: date) -> bool:
    return report_present(day + timedelta(days=1))


def raw_should_refresh(day: date) -> bool:
    if not raw_ready(day):
        return report_present(day)
    if d1_report_present(day):
        raw_m = parents_csv(day).stat().st_mtime
        tag = f"{MONTHS[day.month - 1]}-{day.day:02d}"
        d1 = day + timedelta(days=1)
        d1_tag = f"{MONTHS[d1.month - 1]}-{d1.day:02d}"
        d1_path = oix_reports_dir() / d1.isoformat() / f"Delivered Order in Cargroup ({d1_tag}).csv"
        if d1_path.is_file() and d1_path.stat().st_mtime > raw_m + 1:
            return True
    return False


# ---------------------------------------------------------------------------
# RRP stored workbooks (open / future delivery dates)
# ---------------------------------------------------------------------------

def find_rrp_xlsx(delivery: date) -> Path | None:
    """Newest stored Overall Order workbook for this delivery date. No download."""
    name = overall_name(delivery)
    candidates: list[Path] = []
    extra = os.environ.get("SNOWBALL_RRP_XLSX_DIR")
    search_roots = [rrp_root() / "latest", rrp_root() / "runs"]
    if extra:
        search_roots.insert(0, Path(extra))
    latest = rrp_root() / "latest" / name
    if latest.is_file():
        candidates.append(latest)
    runs = rrp_root() / "runs"
    if runs.is_dir():
        for run in runs.iterdir():
            if run.is_dir():
                p = run / name
                if p.is_file():
                    candidates.append(p)
    if extra:
        p = Path(extra) / name
        if p.is_file():
            candidates.append(p)
    if not candidates:
        return None
    # Prefer the newest mtime. For today's T+0, later hourly runs do not
    # replace the 13:00 file in latest/, and later run folders usually have
    # no T+0 file, so the 13:00 workbook remains the newest T+0 source.
    candidates.sort(key=lambda p: p.stat().st_mtime)
    return candidates[-1]


def load_rrp_cache() -> dict:
    path = rrp_cache_path()
    if not path.is_file():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_rrp_cache(cache: dict) -> None:
    path = rrp_cache_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".json.part")
    tmp.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    tmp.replace(path)


def count_rrp_workbook(path: Path, delivery: date, shop_zones: set[str]) -> dict:
    """Unique parents by exact order date from a stored RRP workbook.

    Filters match the live RRP snapshot (shop/KN015/route/status/district).
    Does not import the RRP downloader. Waybills = remaining rows after filters.
    """
    import openpyxl

    log(f"Counting RRP {path.name} for {delivery.isoformat()} ({path.stat().st_size / 1e6:.1f} MB)")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if "Route Resource Data" not in wb.sheetnames:
            raise RuntimeError(f"{path.name}: missing sheet 'Route Resource Data'; found {wb.sheetnames}")
        ws = wb["Route Resource Data"]
        rows = ws.iter_rows(values_only=True)
        header = next(rows)
        cols = [clean_str(c) for c in header]
        idx = {c: i for i, c in enumerate(cols)}
        need = ["District", "Route Name", "Order Id", "Zone ID", "Waybill Status", "Parent Order ID"]
        miss = [c for c in need if c not in idx]
        if miss:
            raise RuntimeError(f"{path.name}: missing columns {miss}")
        i_dd = idx.get("Delivery Date")
        i_dist, i_route, i_oid = idx["District"], idx["Route Name"], idx["Order Id"]
        i_zone, i_st, i_pid = idx["Zone ID"], idx["Waybill Status"], idx["Parent Order ID"]
        parent_districts = defaultdict(set)
        parent_rows = defaultdict(Counter)
        parent_od: dict[str, date | None] = {}
        n_raw = 0
        drops = Counter()

        def note_od(parent: str, od: date | None) -> None:
            prev = parent_od.get(parent, None)
            if parent not in parent_od:
                parent_od[parent] = od
            elif od is not None and (prev is None or od < prev):
                parent_od[parent] = od

        for row in rows:
            n_raw += 1
            if i_dd is not None and i_dd < len(row):
                dd = to_date(row[i_dd])
                if dd is None:
                    drops["bad_delivery_date"] += 1
                    continue
                if dd != delivery:
                    drops["wrong_delivery_date"] += 1
                    continue
            st = norm_waybill_status(row[i_st] if i_st < len(row) else "")
            if st in EXCLUDE_WAYBILL_STATUSES or CANCEL_STATUS_KEYWORD in st:
                drops["waybill_status"] += 1
                continue
            zone = clean_str(row[i_zone] if i_zone < len(row) else "").upper()
            if zone == EXCLUDE_ZONE_ID or zone.startswith(EXCLUDE_ZONE_ID):
                drops["kn015"] += 1
                continue
            route = clean_str(row[i_route] if i_route < len(row) else "")
            if route == "":
                drops["blank_route"] += 1
                continue
            if any(kw in route for kw in EXCLUDE_ROUTE_KEYWORDS):
                drops["route_keyword"] += 1
                continue
            if is_shop_or_pickup(zone, shop_zones):
                drops["shop_or_pickup"] += 1
                continue
            dist = display_district(row[i_dist] if i_dist < len(row) else "")
            raw_dist = clean_str(row[i_dist] if i_dist < len(row) else "").upper()
            if raw_dist and raw_dist not in VALID_HOME and DISTRICT_ALIAS.get(raw_dist) is None:
                drops["other_district"] += 1
                continue
            if dist is None:
                drops["other_district"] += 1
                continue
            parent = parent_from_overall(
                row[i_pid] if i_pid < len(row) else "",
                row[i_oid] if i_oid < len(row) else "",
            )
            if not parent:
                drops["no_parent"] += 1
                continue
            od = order_date_from_row(
                row[i_oid] if i_oid < len(row) else "",
                row[i_pid] if i_pid < len(row) else "",
            )
            parent_districts[parent].add(dist)
            parent_rows[parent][dist] += 1
            note_od(parent, od)
            drops["kept"] += 1
    finally:
        wb.close()

    assigned, n_cross = assign_cross_district(parent_districts, parent_rows)
    by_od: dict[str, dict] = {}
    unknown = 0
    for parent, dist in assigned.items():
        od = parent_od.get(parent)
        nwb = int(parent_rows[parent][dist])
        if od is None:
            unknown += 1
            continue
        add_parent_to_cells(by_od, od.isoformat(), dist, nwb)
    log(
        f"  {path.name}: rows {n_raw:,} kept {drops['kept']:,}; "
        f"{len(assigned):,} unique parents; cross={n_cross}; unknown_od={unknown}"
    )
    return {
        "source": str(path),
        "mtime": path.stat().st_mtime,
        "size": path.stat().st_size,
        "n_unique_parents": len(assigned),
        "n_cross": n_cross,
        "n_unknown_od": unknown,
        "n_raw": n_raw,
        "n_kept": int(drops["kept"]),
        "by_order_date": by_od,
    }


def rrp_for_delivery(delivery: date, shop_zones: set[str], cache: dict, force: bool = False) -> dict | None:
    if delivery in NO_DELIVERY_DATES:
        return {
            "source": "no-delivery",
            "mtime": 0,
            "size": 0,
            "n_unique_parents": 0,
            "by_order_date": {},
            "skipped_no_delivery": True,
        }
    path = find_rrp_xlsx(delivery)
    if path is None:
        log(f"No stored RRP workbook for {delivery}")
        return cache.get(delivery.isoformat())
    key = delivery.isoformat()
    prev = cache.get(key) or {}
    if (
        not force
        and prev.get("source") == str(path)
        and prev.get("mtime") == path.stat().st_mtime
        and prev.get("size") == path.stat().st_size
        and "by_order_date" in prev
    ):
        log(f"RRP cache hit for {delivery} ({path.name})")
        return prev
    rec = count_rrp_workbook(path, delivery, shop_zones)
    cache[key] = rec
    return rec


# ---------------------------------------------------------------------------
# Build
# ---------------------------------------------------------------------------

def load_base() -> dict:
    path = base_json_path()
    if not path.is_file():
        return {"days": {}, "districts": DISTRICTS, "snapshots": {}}
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, dict) or "days" not in data:
        return {"days": {}, "districts": DISTRICTS, "snapshots": {}}
    return data


def is_complete_delivery(day: date, today: date) -> bool:
    if day in NO_DELIVERY_DATES:
        return True
    if day >= today:
        return False
    # Yesterday is complete once the D+1 06:00 window has closed. After
    # 06:00 HK we treat it as complete if a report or raw file exists.
    if day == today - timedelta(days=1) and now_hk().hour < 6:
        return False
    return raw_ready(day) or report_present(day)


def merge_into(accum: dict[str, dict], by_od: dict, delivery: date) -> None:
    dd = delivery.isoformat()
    for od, blk in by_od.items():
        rec = accum.setdefault(od, {})
        # later sources for the same (od, dd) overwrite (raw beats a stale RRP
        # if both were somehow present — caller should not pass both)
        rec[dd] = {k: [int(v[0]), int(v[1])] for k, v in blk.items()}


def build_snowball(today: date | None = None, force_completed: bool = False, force_rrp: bool = False) -> dict:
    today = today or today_hk()
    generated = now_hk().strftime("%Y-%m-%dT%H:%M")
    base = load_base()
    days_out = {}
    # Keep pre-METHOD_START order-date rows as they are (old Tableau / parent-order series).
    for od, rec in (base.get("days") or {}).items():
        try:
            od_d = date.fromisoformat(od)
        except ValueError:
            days_out[od] = rec
            continue
        if od_d < METHOD_START:
            days_out[od] = rec

    completed_end = today - timedelta(days=1)
    if now_hk().hour < 6:
        completed_end = today - timedelta(days=2)

    ident = {}
    ctx = None
    nov_cells: dict[str, dict] = {}
    d = METHOD_START
    while d <= completed_end:
        iso = d.isoformat()
        if d in NO_DELIVERY_DATES:
            ident[iso] = {"source": "no-delivery", "parents": 0}
            d += timedelta(days=1)
            continue
        if force_completed or raw_should_refresh(d):
            if report_present(d):
                info, ctx = try_export_raw(d, ctx=ctx, force=True)
                if info:
                    log(f"Raw export {d}: parents={info.get('parents')} ftd={info.get('ftd')}")
        if raw_ready(d):
            by_od, stats = cells_from_raw(d)
            ident[iso] = {"source": "oix_delivered_raw", **stats}
            by_od_nov = {od: blk for od, blk in by_od.items() if od >= METHOD_START.isoformat()}
            merge_into(nov_cells, by_od_nov, d)
        else:
            # Keep cells already stored for this delivery date so an hourly run
            # cannot wipe the November backfill when the raw CSV is not on disk.
            preserved = 0
            if (base.get("method") == METHOD_VERSION) or iso >= METHOD_START.isoformat():
                for od, rec in (base.get("days") or {}).items():
                    if od < METHOD_START.isoformat():
                        continue
                    cell = (rec.get("cells") or {}).get(iso)
                    if cell:
                        nov_cells.setdefault(od, {})[iso] = {
                            k: [int(v[0]), int(v[1])] for k, v in cell.items() if isinstance(v, list) and len(v) >= 2
                        }
                        preserved += 1
            ident[iso] = {"source": "preserved" if preserved else "missing", "parents": 0, "preserved_order_dates": preserved}
            if preserved:
                log(f"No raw CSV for {d}; preserved {preserved} existing order-date cells")
            else:
                log(f"WARNING: no raw unique-parent CSV for completed day {d}")
        d += timedelta(days=1)

    shop_zones = load_shop_zones()
    cache = load_rrp_cache()
    future_start = completed_end + timedelta(days=1)
    future_end = today + timedelta(days=7)
    cache_dirty = False
    d = future_start
    while d <= future_end:
        rec = rrp_for_delivery(d, shop_zones, cache, force=force_rrp)
        iso = d.isoformat()
        if rec is None:
            ident[iso] = {"source": "rrp_missing"}
            d += timedelta(days=1)
            continue
        prev = cache.get(iso)
        if rec is not prev:
            cache_dirty = True
        ident[iso] = {
            "source": "rrp_xlsx" if not rec.get("skipped_no_delivery") else "no-delivery",
            "parents": rec.get("n_unique_parents", 0),
            "file": rec.get("source"),
            "unknown_od": rec.get("n_unknown_od", 0),
        }
        by_od = rec.get("by_order_date") or {}
        by_od_nov = {od: blk for od, blk in by_od.items() if od >= METHOD_START.isoformat()}
        merge_into(nov_cells, by_od_nov, d)
        d += timedelta(days=1)
    if cache_dirty or force_rrp:
        save_rrp_cache(cache)

    slot_today = f"{now_hk().hour:02d}:00"
    for od in sorted(nov_cells):
        cells = prune_zero_cells(nov_cells[od])
        od_d = date.fromisoformat(od)
        complete = od_d < today
        days_out[od] = {
            "slot": "06:00" if complete else slot_today,
            "asOf": generated,
            "cells": cells,
        }

    d = METHOD_START
    while d <= today:
        iso = d.isoformat()
        if iso not in days_out:
            days_out[iso] = {
                "slot": "06:00" if d < today else slot_today,
                "asOf": generated,
                "cells": {},
            }
        d += timedelta(days=1)

    out = {
        "days": dict(sorted(days_out.items())),
        "districts": list(base.get("districts") or DISTRICTS),
        "generatedAt": generated,
        "lastRun": generated,
        "method": METHOD_VERSION,
        "methodFrom": METHOD_START.isoformat(),
        "identity": ident,
    }
    if "snapshots" in base and isinstance(base["snapshots"], dict):
        snaps = {k: v for k, v in base["snapshots"].items() if k < METHOD_START.isoformat()}
        if snaps:
            out["snapshots"] = snaps
    return out


def write_snowball(payload: dict, path: Path | None = None) -> Path:
    path = path or snowball_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".part")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    tmp.replace(path)
    log(f"Wrote {path} ({path.stat().st_size / 1024:.0f} KB, {len(payload.get('days') or {})} order dates)")
    return path


def refresh_snowball_for_rolling() -> dict | None:
    """Hourly hook for kpi_pipeline.py --section rolling. Never starts an OIX download."""
    log("Refreshing Snowball Table from unique-parent raw CSVs and stored RRP snapshots")
    try:
        payload = build_snowball()
        write_snowball(payload)
        return payload
    except Exception as exc:
        log(f"WARNING: snowball refresh failed: {type(exc).__name__}: {exc}")
        traceback.print_exc()
        return None


def summarize(payload: dict, delivery: str) -> dict:
    """Sum unique-parent orders across order dates for one delivery date (Identity A)."""
    totals = Counter()
    by_month = Counter()
    for od, rec in (payload.get("days") or {}).items():
        cell = (rec.get("cells") or {}).get(delivery)
        if not cell:
            continue
        n = (cell.get("OVERALL") or [0, 0])[0]
        totals["all"] += n
        if od >= METHOD_START.isoformat():
            totals["method"] += n
        else:
            totals["legacy"] += n
        by_month[od[:7]] += n
    return {"delivery": delivery, "totals": dict(totals), "by_month": dict(by_month)}


def main(argv: list[str] | None = None) -> int:
    import argparse
    p = argparse.ArgumentParser(description="Build or refresh snowball_table.json")
    p.add_argument("--build", action="store_true", help="Build and write snowball_table.json")
    p.add_argument("--force-completed", action="store_true")
    p.add_argument("--force-rrp", action="store_true")
    p.add_argument("--out", help="Output path (default: SNOWBALL_TABLE_PATH / public/snowball_table.json)")
    args = p.parse_args(argv)
    if not args.build:
        p.print_help()
        return 2
    payload = build_snowball(force_completed=args.force_completed, force_rrp=args.force_rrp)
    write_snowball(payload, Path(args.out) if args.out else None)
    # Print a short identity check for completed November weekdays we care about
    for d in ("2026-11-20", "2026-11-21", "2026-11-22"):
        s = summarize(payload, d)
        ident = (payload.get("identity") or {}).get(d) or {}
        log(
            f"Identity {d}: matrix all={s['totals'].get('all')} method={s['totals'].get('method')} "
            f"legacy={s['totals'].get('legacy')} raw_parents={ident.get('parents')} "
            f"before_method={ident.get('before_method_start')} source={ident.get('source')}"
        )
    # November order-date overall totals (sum of delivery-date cells)
    print("November order-date OVERALL totals:", flush=True)
    for od, rec in (payload.get("days") or {}).items():
        if not od.startswith("2026-11"):
            continue
        total = 0
        for dd, blk in (rec.get("cells") or {}).items():
            total += (blk.get("OVERALL") or [0, 0])[0]
        print(f"  {od}: {total:,}", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
