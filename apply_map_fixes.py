#!/usr/bin/env python3
"""Push verified / geocoded coordinates into index.html (window.__MAP_FIX__ block).

  python apply_map_fixes.py Adjusted_Coordinates_Map_Check.xlsx index.html            # use 'Verified Latitude/Longitude' columns
  python apply_map_fixes.py Adjusted_Coordinates_Map_Check.xlsx index.html --geocode  # also try CSDI -> OSM for rows left blank

--geocode needs internet (run it on your own machine). A geocoded hit is accepted only if it lies inside Hong Kong and within
'Zone limit (km)' of the zone reference centre, so homonym matches elsewhere in HK are rejected.
Existing __MAP_FIX__ entries are kept; new ones override them. Re-run any time; reload the dashboard afterwards.
"""
import sys, re, json, math, time
import pandas as pd

def km(a, b, c, d): return math.hypot((a - c) * 111, (b - d) * 111 * math.cos(math.radians(a)))

def geocode(session, name_en, name_zh):
    UA = {"User-Agent": "LOG-KPI-Tracker/1.0"}
    for q in dict.fromkeys([n for n in (name_zh, name_en, (name_en or '') + ', Hong Kong') if n]):
        try:
            r = session.get("https://geodata.gov.hk/gs/api/v1.0.0/locationSearch", params={"q": q}, headers=UA, timeout=30).json()
            for it in r or []:
                if "latitude" in it: yield float(it["latitude"]), float(it["longitude"]), "csdi"
        except Exception as e: ERR.append(f"CSDI: {e}")
        try:
            time.sleep(1.1)
            for it in session.get("https://nominatim.openstreetmap.org/search", params={"q": q, "format": "jsonv2", "limit": 3, "countrycodes": "hk"}, headers=UA, timeout=30).json():
                yield float(it["lat"]), float(it["lon"]), "osm"
        except Exception as e: ERR.append(f"OSM: {e}")

ERR = []

def main():
    xlsx, html = sys.argv[1], sys.argv[2]
    df = pd.read_excel(xlsx, sheet_name="Adjusted List")
    s = open(html, encoding="utf-8").read()
    m = re.search(r"/\*MAP_FIX_START\*/window\.__MAP_FIX__=(\{.*?\});/\*MAP_FIX_END\*/", s, re.S)
    if not m: sys.exit(f"{html} has no MAP_FIX block - it is not the updated index.html (v37.0). Nothing changed.")
    fix = json.loads(m.group(1))
    sess = None
    if "--geocode" in sys.argv:
        import requests; sess = requests.Session()
    n_v = n_g = 0
    for _, r in df.iterrows():
        code = str(r["Estate Code"])
        la, ln = r["Verified Latitude (fill in)"], r["Verified Longitude (fill in)"]
        if pd.notna(la) and pd.notna(ln):
            fix[code] = [round(float(la), 6), round(float(ln), 6), "verified", "from Adjusted_Coordinates list"]; n_v += 1
        elif sess is not None and code not in fix and pd.notna(r["Zone ref. Latitude"]):
            for gla, gln, prov in geocode(sess, r["Estate Name (EN)"], r["Estate Name (ZH)"]):
                if km(gla, gln, r["Zone ref. Latitude"], r["Zone ref. Longitude"]) <= float(r["Zone limit (km)"] or 2):
                    fix[code] = [round(gla, 6), round(gln, 6), "geocode-" + prov, "name search, validated against zone"]; n_g += 1; break
    if sess is not None and n_g == 0 and ERR: print("geocoder errors (first 3):", *ERR[:3], sep="\n  ")
    block = "/*MAP_FIX_START*/window.__MAP_FIX__=" + json.dumps(fix, ensure_ascii=False) + ";/*MAP_FIX_END*/"
    s = re.sub(r"/\*MAP_FIX_START\*/.*?/\*MAP_FIX_END\*/", lambda _: block, s, count=1, flags=re.S)
    open(html, "w", encoding="utf-8").write(s)
    print(f"verified {n_v} · geocoded {n_g} · total overrides {len(fix)} -> {html}")

if __name__ == "__main__": main()
