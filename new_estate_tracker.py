#!/usr/bin/env python3
"""
New-estate tracker (v1.0) — private + public housing that is ready for move-in.

Run through the pipeline:   python kpi_pipeline.py --section newestate [--dry-run] [--limit N]
or standalone:              python new_estate_tracker.py [--dry-run]

What one run does
  1. Collect candidates from every source (see SOURCES):
       watchlist  - the "New Address ..." sheet of the Zone Grouping workbook (your hand-kept list; also the seed)
       bd_op      - Buildings Dept. Monthly Digest Table 5.6 CSV (occupation permits issued -> private housing)
       html_table - any web page with a table (HA / SRPA / seehse ...), configured in new_estate_sources.json
  2. Merge into new_estates.json by a stable key (keeps first_seen, lat/long, history; reports what is NEW / CHANGED).
  3. Check each estate against the Full Open Address List (Raw) -> already opened? (name match)
  4. Estimate the Estate-Code open date, set Status + suggested Action.
  5. Geocode every estate that has no coordinates yet from its ADDRESS (cached in geocode_cache.json,
     HK-polygon validated; providers: CSDI Location Search -> OpenStreetMap Nominatim).
  6. Write New_Estate_Tracker.xlsx (same columns as the "New Address 2026 Q3" sheet + Latitude / Longitude / Status / ...).

Nothing here touches index.html; the dashboard map can read new_estates.json later.
"""
import os, re, sys, json, time, argparse, datetime as dt
from pathlib import Path

import pandas as pd

STATE_PATH     = os.environ.get("NEW_ESTATE_STATE", "./new_estates.json")
OUT_XLSX       = os.environ.get("NEW_ESTATE_XLSX", "./New_Estate_Tracker.xlsx")
SOURCES_CFG    = os.environ.get("NEW_ESTATE_SOURCES", "./new_estate_sources.json")
GEO_CACHE_PATH = os.environ.get("GEOCODE_CACHE", "./geocode_cache.json")
WATCHLIST_XLSX = os.environ.get("NEW_ESTATE_WATCHLIST", "")        # Zone Grouping workbook (sheet name starts 'New Address')
# BD Monthly Digest Table 5.6 (completed new buildings, occupation permit issued) - CSDI WFS, as listed on data.gov.hk
BD_OP_URL      = os.environ.get("BD_OP_CSV_URL", "https://portal.csdi.gov.hk/server/services/common/bd_rcd_1629267205236_397/MapServer/WFSServer?service=wfs&request=GetFeature&typename=BDMD56&outputFormat=CSV")
BD_OP_DAYS     = int(os.environ.get("BD_OP_DAYS", "240"))          # only permits issued within this many days
BD_OP_MIN_UNITS = int(os.environ.get("BD_OP_MIN_UNITS", "20"))
GEO_PROVIDERS  = [p.strip() for p in os.environ.get("GEOCODE_PROVIDERS", "csdi,osm").split(",") if p.strip()]
LOOKAHEAD_DAYS = int(os.environ.get("NEW_ESTATE_LOOKAHEAD_DAYS", "120"))   # 'ready soon' window for Action = add to Raw
OPEN_CODE_LAG_DAYS = int(os.environ.get("OPEN_CODE_LAG_DAYS", "60"))       # occupancy date -> earliest estate-code date
OPEN_CODE_CYCLE_MONTHS = tuple(int(x) for x in os.environ.get("OPEN_CODE_CYCLE_MONTHS", "2,6,10").split(","))
UA = {"User-Agent": "LOG-KPI-Tracker/1.0 (new-estate tracker; contact: logistics team)"}

HK_BBOX = (22.13, 22.57, 113.82, 114.45)


def today_hkt():
    return dt.datetime.now(dt.timezone(dt.timedelta(hours=8))).date()


# ------------------------------------------------------------------ helpers
def norm(s):
    """name key: lower-case, no spaces/punctuation, no phase words."""
    s = str(s or "").lower()
    s = re.sub(r"(第?[一二三四五六七八九十\d]+期|phase\s*\w+|\(.*?\)|（.*?）)", "", s)
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]", "", s)


def parse_date(v):
    """'2026年10月31日' | '2026-10-31' | date -> ISO str or None."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.search(r"(\d{4})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})", s)
    if m:
        try:
            return dt.date(int(m[1]), int(m[2]), int(m[3])).isoformat()
        except ValueError:
            return None
    return None


def zh_date(iso):
    d = dt.date.fromisoformat(iso)
    return f"{d.year}年{d.month}月{d.day}日"


def est_open_code_date(occ_iso):
    """Earliest 1st-of-month in OPEN_CODE_CYCLE_MONTHS that is >= occupancy date + OPEN_CODE_LAG_DAYS."""
    if not occ_iso:
        return None
    t = dt.date.fromisoformat(occ_iso) + dt.timedelta(days=OPEN_CODE_LAG_DAYS)
    for y in (t.year, t.year + 1):
        for m in sorted(OPEN_CODE_CYCLE_MONTHS):
            d = dt.date(y, m, 1)
            if d >= t:
                return d.isoformat()
    return None


def make_key(rec):
    return norm(rec.get("name_zh")) or norm(rec.get("name_en")) or norm(rec.get("address"))


def in_hk_default(lat, lng):
    return HK_BBOX[0] <= lat <= HK_BBOX[1] and HK_BBOX[2] <= lng <= HK_BBOX[3]


# ------------------------------------------------------------------ sources
def find_watchlist(explicit=None):
    """Any .xlsx (this folder, script folder, MAP_EXCEL_FOLDER, parent) that has a sheet named 'New Address…' - file name does not matter."""
    import glob
    from openpyxl import load_workbook
    if explicit and os.path.exists(explicit):
        return explicit
    dirs = [os.getcwd(), os.path.dirname(os.path.abspath(__file__)), os.environ.get("MAP_EXCEL_FOLDER", ""), os.path.dirname(os.getcwd())]
    seen, hits = set(), []
    for d in dirs:
        if not d or not os.path.isdir(d) or d in seen:
            continue
        seen.add(d)
        for f in glob.glob(os.path.join(d, "*.xlsx")):
            if os.path.basename(f).startswith("~$") or os.path.basename(f) == os.path.basename(OUT_XLSX):
                continue
            try:
                wb = load_workbook(f, read_only=True)
                if any(n.lower().startswith("new address") for n in wb.sheetnames):
                    hits.append(f)
                wb.close()
            except Exception:
                pass
    if hits:
        return max(hits, key=os.path.getmtime)
    print("  ⚠️ no workbook with a 'New Address…' sheet found in: " + "; ".join(seen) + "  -> use --watchlist <file>")
    return ""


def src_watchlist(path):
    """The hand-kept 'New Address ...' sheet (same column layout as the Zone Grouping workbook)."""
    if not path or not os.path.exists(path):
        print(f"  (watchlist skipped: {path!r} not found)")
        return []
    xl = pd.ExcelFile(path)
    sheet = next((s for s in xl.sheet_names if s.lower().startswith("new address")), None)
    if not sheet:
        print(f"  (watchlist skipped: no 'New Address…' sheet in {path})")
        return []
    df = xl.parse(sheet, dtype=object)
    cols = list(df.columns)
    out = []
    for _, r in df.iterrows():
        zh, en = r.get(cols[1]), r.get(cols[2])
        if pd.isna(zh) and pd.isna(en):
            continue
        if not isinstance(r.get(cols[0]), (int, float)) or pd.isna(r.get(cols[0])):
            continue                                   # footnote rows
        g = lambda i: (None if i >= len(cols) or pd.isna(r.get(cols[i])) else str(r.get(cols[i])).strip())
        out.append(dict(name_zh=g(1), name_en=g(2), address=g(3), district18=g(4), type=g(5),
                        occupancy_raw=g(6), occupancy_date=parse_date(r.get(cols[6])),
                        source=g(8), remark=g(9), origin="watchlist"))
    return out


def _bd_date(v):
    """LASTUPDATE may be ISO text, '2026年…' or an epoch (s / ms)."""
    iso = parse_date(v)
    if iso:
        return iso
    t = str(v or "").strip()
    if re.fullmatch(r"\d{10,13}", t):
        n = int(t) / (1000 if len(t) == 13 else 1)
        return dt.datetime.fromtimestamp(n, dt.timezone.utc).date().isoformat()
    m = re.search(r"(\d{4})[-/](\d{1,2})", t)
    return f"{m[1]}-{int(m[2]):02d}-01" if m else None


def src_bd_op(session):
    """BD Monthly Digest Table 5.6 via CSDI WFS CSV. Known columns: ADDRESS_EN/TC, SEARCH01..04, NSEARCH01..14 (_EN/_TC),
    LATITUDE, LONGITUDE, LASTUPDATE. Coordinates come straight from BD (no geocoding needed)."""
    from io import StringIO
    try:
        r = session.get(BD_OP_URL, headers=UA, timeout=90)
        r.raise_for_status()
        df = pd.read_csv(StringIO(r.content.decode("utf-8-sig", errors="ignore")), dtype=str).fillna("")
    except Exception as e:
        print(f"  ⚠️ bd_op skipped: {e}")
        return []
    up = {c.upper(): c for c in df.columns}
    c_en, c_tc = up.get("ADDRESS_EN"), up.get("ADDRESS_TC")
    if not (c_en or c_tc):
        print(f"  ⚠️ bd_op: no ADDRESS_EN/ADDRESS_TC column. Columns: {list(df.columns)[:15]}")
        return []
    c_lat, c_lng, c_upd = up.get("LATITUDE"), up.get("LONGITUDE"), up.get("LASTUPDATE")
    txt_cols = [c for c in df.columns if re.match(r"(N?SEARCH\d+_EN|DATASET_EN)$", c.upper())]
    cutoff = today_hkt() - dt.timedelta(days=BD_OP_DAYS)
    year_min = today_hkt().year - 1
    out, shown, n_res = [], 0, 0
    for _, r in df.iterrows():
        blob = " | ".join(str(r[c]) for c in txt_cols if str(r[c]).strip())
        low = blob.lower()
        if not re.search(r"apartment|domestic|residential|composite|flat", low):
            continue
        if re.search(r"single family|guard house", low) and not re.search(r"apartment|composite", low):
            continue
        n_res += 1
        m = re.search(r"/(\d{4})/OP", blob)                       # e.g. NT58/2026/OP -> year of permit
        if m and int(m[1]) < year_min:
            continue
        iso = _bd_date(r[c_upd]) if c_upd else None
        if iso and dt.date.fromisoformat(iso) < cutoff:
            continue
        mu = re.search(r"(\d[\d,]*)\s*(?:domestic\s+)?units", low)
        units = int(mu[1].replace(",", "")) if mu else 0
        if units and units < BD_OP_MIN_UNITS:
            continue
        op = re.search(r"[A-Z]{2}\d+/\d{4}/OP", blob)
        rec = dict(name_zh=None, name_en=(r[c_en] or r[c_tc]).strip(), address=(r[c_tc] or r[c_en]).strip(),
                   district18=None, type="私樓", occupancy_date=iso,
                   source="屋宇署 佔用許可證 (OP issued)" + (f" {op[0]}" if op else ""),
                   remark=(f"{units} units; " if units else "") + "date = BD record update date (actual move-in may be later)",
                   origin="bd_op")
        try:
            la, ln = float(r[c_lat]), float(r[c_lng])
            if in_hk_default(la, ln):
                rec.update(lat=round(la, 6), lng=round(ln, 6), geo_src="BD")
        except (TypeError, ValueError, KeyError):
            pass
        out.append(rec)
        if shown < 3:
            shown += 1
            print(f"    sample: {rec['name_en']} | {iso} | {blob[:140]}")
    print(f"  bd_op: {n_res} residential-looking rows -> {len(out)} within the last {BD_OP_DAYS} days")
    return out


def src_html_tables(session, cfg_path):
    """Generic: new_estate_sources.json = [{"name":"HA intake","url":"...","type":"公屋","table_index":0,
         "columns":{"name_zh":"屋邨","address":"地點","occupancy_date":"入伙日期","remark":"單位數目"}}, ...]"""
    if not os.path.exists(cfg_path):
        return []
    out = []
    for s in json.load(open(cfg_path, encoding="utf-8")):
        try:
            r = session.get(s["url"], headers=UA, timeout=60)
            r.raise_for_status()
            from io import StringIO
            df = pd.read_html(StringIO(r.text))[s.get("table_index", 0)]
        except Exception as e:
            print(f"  ⚠️ {s.get('name', s.get('url'))} skipped: {e}")
            continue
        m = s.get("columns", {})
        for _, r in df.iterrows():
            g = lambda k: (None if k not in m or m[k] not in df.columns or pd.isna(r[m[k]]) else str(r[m[k]]).strip())
            rec = dict(name_zh=g("name_zh"), name_en=g("name_en"), address=g("address") or g("name_zh"),
                       district18=g("district18"), type=s.get("type") or g("type"),
                       occupancy_raw=g("occupancy_date"), occupancy_date=parse_date(g("occupancy_date")),
                       source=s.get("name") or s["url"], remark=g("remark"), origin="html:" + s.get("name", "table"))
            if rec["name_zh"] or rec["name_en"]:
                out.append(rec)
    return out


# ------------------------------------------------------------------ Raw (already-opened) check
def build_raw_index(raw_df):
    """-> (big normalised string of all names, {name_key: code})"""
    names = {}
    if raw_df is not None:
        for _, r in raw_df.iterrows():
            for col in ("zh", "en"):
                k = norm(r.get(col))
                if len(k) >= 3:
                    names.setdefault(k, r.get("code", ""))
    return names


def raw_match(rec, names):
    """exact name or 'candidate name contained in a Raw name' (len>=3) -> Raw estate code or None."""
    zh, en = norm(rec.get("name_zh")), norm((rec.get("name_en") or "").split(",")[0])
    for c in (zh, en):                                   # exact name (either language)
        if len(c) >= 3 and c in names:
            return names[c]
    if len(zh) >= 3:                                     # Chinese name contained in a Raw name (block / phase suffixes)
        for k, code in names.items():
            if zh in k:
                return code
    return None                                          # English: exact only - 'CHESTER' must not match 'CHESTER II' etc.


# ------------------------------------------------------------------ geocoding
def _hk80_to_wgs84(x, y):
    try:
        from pyproj import Transformer
        t = Transformer.from_crs("EPSG:2326", "EPSG:4326", always_xy=True)
        lng, lat = t.transform(float(x), float(y))
        return lat, lng
    except Exception:
        return None


def _geo_csdi(session, q):
    r = session.get("https://geodata.gov.hk/gs/api/v1.0.0/locationSearch", params={"q": q}, headers=UA, timeout=30)
    r.raise_for_status()
    for it in (r.json() or []):
        if "latitude" in it and "longitude" in it:
            yield float(it["latitude"]), float(it["longitude"]), it.get("addressZH") or it.get("nameZH") or ""
        elif "x" in it and "y" in it:
            ll = _hk80_to_wgs84(it["x"], it["y"])
            if ll:
                yield ll[0], ll[1], it.get("addressZH") or it.get("nameZH") or ""


def _geo_osm(session, q):
    time.sleep(1.1)                                   # Nominatim policy: max 1 request / second
    r = session.get("https://nominatim.openstreetmap.org/search",
                    params={"q": q, "format": "jsonv2", "limit": 3, "countrycodes": "hk"}, headers=UA, timeout=30)
    r.raise_for_status()
    for it in r.json():
        yield float(it["lat"]), float(it["lon"]), it.get("display_name", "")


_PROV = {"csdi": _geo_csdi, "osm": _geo_osm}


def geocode(session, rec, cache, in_hk):
    """Address -> (lat, lng, provider, query) using several query variants; cached; HK-validated."""
    addr = (rec.get("address") or "").strip()
    variants = [v for v in dict.fromkeys([
        rec.get("name_en"), addr, (addr + " 香港") if addr else None,
        ((rec.get("name_zh") or "") + " " + addr).strip() or None,
        ((rec.get("name_en") or "") + ", Hong Kong") if rec.get("name_en") else None]) if v]
    for q in variants:
        if q in cache:
            c = cache[q]
            if c:
                return c["lat"], c["lng"], c["prov"], q
            continue
        for p in GEO_PROVIDERS:
            try:
                for lat, lng, _label in _PROV[p](session, q):
                    if in_hk(lat, lng):
                        cache[q] = {"lat": round(lat, 6), "lng": round(lng, 6), "prov": p}
                        return round(lat, 6), round(lng, 6), p, q
            except Exception as e:
                print(f"    geocode {p} failed for {q!r}: {e}")
        cache[q] = None                               # remember misses so we do not hammer the services every run
    return None


# ------------------------------------------------------------------ merge + status
TRACKED = ("name_zh", "name_en", "address", "district18", "type", "occupancy_date", "source", "remark")


def merge(state, cands, today):
    new, changed = [], []
    for c in cands:
        k = make_key(c)
        if not k:
            continue
        cur = state.get(k)
        if cur is None:
            c.update(first_seen=today.isoformat(), last_seen=today.isoformat(), key=k, history=[])
            state[k] = c
            new.append(k)
            continue
        diffs = {f: (cur.get(f), c.get(f)) for f in TRACKED if c.get(f) and c.get(f) != cur.get(f)}
        if c.get("origin") == "watchlist" or not cur.get("occupancy_date"):
            pass
        else:                                          # an automatic source never overwrites a hand-kept date
            diffs.pop("occupancy_date", None)
        if diffs:
            cur["history"].append({"date": today.isoformat(), **{f: v[0] for f, v in diffs.items()}})
            for f, (_, nv) in diffs.items():
                cur[f] = nv
            if "address" in diffs:
                cur.pop("lat", None); cur.pop("lng", None)   # address changed -> geocode again
            changed.append(k)
        cur["last_seen"] = today.isoformat()
    return new, changed


def classify(rec, names, today):
    rec["raw_code"] = raw_match(rec, names)
    occ = rec.get("occupancy_date")
    rec["est_open_code_date"] = est_open_code_date(occ)
    if rec["raw_code"]:
        rec["status"], rec["action"] = "Already in Raw", "已在Raw"
    elif rec.get("origin") == "bd_op" and not rec.get("name_zh"):
        rec["status"], rec["action"] = "OP issued — verify estate name / Raw", "核實"
    elif not occ:
        rec["status"], rec["action"] = "Date TBC", "待確認"
    else:
        d = dt.date.fromisoformat(occ)
        if d <= today:
            rec["status"], rec["action"] = "Occupied — not in Raw", "加入Raw"
        elif (d - today).days <= LOOKAHEAD_DAYS:
            rec["status"], rec["action"] = f"Move-in in {(d - today).days} days", "加入Raw"
        else:
            rec["status"], rec["action"] = "Upcoming", "觀察"


# ------------------------------------------------------------------ output
HEADERS = ["編號", "屋苑/樓宇名稱 (中)", "屋苑/樓宇名稱 (英)", "地址", "地區", "類型", "正式入伙日期",
           "Estimated Open Estate Code Date", "資料來源", "備註", "建議Action",
           "Latitude", "Longitude", "Geocode source", "Status", "Raw estate code", "First seen", "Last updated"]


def write_excel(state, new_keys, changed_keys, path, run_log):
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = "New Address"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True); c.fill = PatternFill("solid", fgColor="D9E1F2")
    order = sorted(state.values(), key=lambda r: (r.get("action") == "已在Raw", r.get("occupancy_date") or "9999", r.get("name_zh") or ""))
    yellow, green = PatternFill("solid", fgColor="FFF2CC"), PatternFill("solid", fgColor="E2EFDA")
    for i, r in enumerate(order, 1):
        occ = zh_date(r["occupancy_date"]) if r.get("occupancy_date") else (r.get("occupancy_raw") or "")
        eoc = zh_date(r["est_open_code_date"]) if r.get("est_open_code_date") else "待確認"
        ws.append([i, r.get("name_zh"), r.get("name_en"), r.get("address"), r.get("district18"), r.get("type"), occ, eoc,
                   r.get("source"), r.get("remark"), r.get("action"), r.get("lat"), r.get("lng"), r.get("geo_src"),
                   r.get("status"), r.get("raw_code"), r.get("first_seen"), r.get("last_seen")])
        fill = green if r["key"] in new_keys else yellow if r["key"] in changed_keys else None
        if fill:
            for c in ws[ws.max_row]:
                c.fill = fill
    for j, w in enumerate([6, 26, 40, 28, 10, 10, 16, 18, 30, 40, 12, 11, 11, 12, 22, 12, 12, 12], 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "C2"
    ws2 = wb.create_sheet("Run log")
    ws2.append(["Green = new this run · Yellow = changed this run"])
    for line in run_log:
        ws2.append([line])
    ws2.column_dimensions["A"].width = 120
    wb.save(path)


# ------------------------------------------------------------------ main entry
def run(raw_loader=None, in_hk=None, dry_run=False, limit=None, watchlist=None):
    import requests
    today = today_hkt()
    in_hk = in_hk or in_hk_default
    print(f"🏘️ New-estate tracker — {today}")
    state = json.load(open(STATE_PATH, encoding="utf-8")) if os.path.exists(STATE_PATH) else {}
    cache = json.load(open(GEO_CACHE_PATH, encoding="utf-8")) if os.path.exists(GEO_CACHE_PATH) else {}
    session = requests.Session()

    wl = find_watchlist(watchlist or WATCHLIST_XLSX)
    if wl:
        print(f"  watchlist file: {wl}")
    cands = src_watchlist(wl)
    print(f"  watchlist: {len(cands)} rows")
    if not dry_run:
        for name, fn in (("bd_op", lambda: src_bd_op(session)), ("web tables", lambda: src_html_tables(session, SOURCES_CFG))):
            got = fn(); print(f"  {name}: {len(got)} rows"); cands += got
    new, changed = merge(state, cands, today)

    try:
        names = build_raw_index(raw_loader() if raw_loader else None)
    except Exception as e:
        print(f"  ⚠️ Raw list not usable ({e}) — 'already in Raw' check skipped this run. Use --map-excel <Full Open Address List.xlsx>.")
        names = {}
    print(f"  Raw index: {len(names):,} names")
    for r in state.values():
        classify(r, names, today)

    todo = [r for r in state.values() if r.get("address") or r.get("name_en") if r.get("lat") is None and r["status"] != "Already in Raw"]
    if limit:
        todo = todo[:limit]
    print(f"  geocoding {len(todo)} estate(s) …" + (" (dry-run: skipped)" if dry_run else ""))
    if not dry_run:
        for r in todo:
            g = geocode(session, r, cache, in_hk)
            if g:
                r["lat"], r["lng"], r["geo_src"], r["geo_query"] = g
            else:
                r["geo_src"] = "NOT FOUND — check address"
                print(f"    ⚠️ no coordinates: {r.get('name_zh') or r.get('name_en')} | {r.get('address')}")

    log = [f"Run {today}: {len(state)} estates tracked · {len(new)} new · {len(changed)} changed"]
    for k in new:
        r = state[k]; log.append(f"NEW  {r.get('name_zh') or r.get('name_en')} — {r.get('type')} — {r.get('occupancy_date') or 'date TBC'} — {r['action']}")
    for k in changed:
        r = state[k]; log.append(f"CHG  {r.get('name_zh') or r.get('name_en')} — {r['history'][-1]}")
    ready = [r for r in state.values() if r["action"] == "加入Raw"]
    log.append(f"Ready / moving in ≤{LOOKAHEAD_DAYS} days and not in Raw: {len(ready)}")
    print("\n".join("  " + l for l in log))

    if not dry_run:
        json.dump(state, open(STATE_PATH, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        json.dump(cache, open(GEO_CACHE_PATH, "w", encoding="utf-8"), ensure_ascii=False)
        write_excel(state, set(new), set(changed), OUT_XLSX, log)
        print(f"  ✅ {OUT_XLSX} · {STATE_PATH}")
    return state


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true"); ap.add_argument("--limit", type=int)
    ap.add_argument("--watchlist"); ap.add_argument("--raw")
    a = ap.parse_args()
    def _raw():
        if not a.raw:
            return None
        df = pd.read_excel(a.raw, dtype=str)
        df.columns = [str(c).strip().upper() for c in df.columns]
        return pd.DataFrame({"code": df.get("ESTATE CODE"), "en": df.get("ESTATE NAME EN"), "zh": df.get("ESTATE NAME ZH")})
    run(_raw, dry_run=a.dry_run, limit=a.limit, watchlist=a.watchlist)
