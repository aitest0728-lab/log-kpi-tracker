#!/usr/bin/env python3
"""Correct the suspicious map coordinates with the CARTO Location Data Services (LDS) geocoding API.

  python carto_fix_coords.py --test                       # 1) check key + connectivity with one address, nothing is written
  python carto_fix_coords.py                              # 2) geocode every row of the Excel list, write overrides into index.html

Defaults (override with options):
  --xlsx  Adjusted_Coordinates_Map_Check.xlsx     --html index.html     --key-file "Carto Map API Key.txt"
  --host  https://gcp-us-east1.api.carto.com      (EU accounts: https://gcp-europe-west1.api.carto.com ; use the host of your CARTO account)

Needs: pip install pandas openpyxl requests
IMPORTANT: geocoding is NOT covered by a Basemaps (tile) key. You need a CARTO Cloud Native API *access token* and a geocoding provider
(HERE / Google / Mapbox / TomTom / TravelTime) enabled for the organisation. --test tells you which of these is missing.

A result is accepted only if it is inside Hong Kong AND within the row's 'Zone limit (km)' of the delivery-zone reference centre
(columns 'Zone ref. Latitude/Longitude'), so same-name places elsewhere are rejected. Accepted rows are written to the
window.__MAP_FIX__ block of index.html (method 'carto'); everything is logged to carto_geocode_results.csv.
"""
import argparse, csv, json, math, re, sys, time
import requests

HK = (22.13, 22.57, 113.82, 114.45)
def in_hk(la, ln): return HK[0] <= la <= HK[1] and HK[2] <= ln <= HK[3]
def km(a, b, c, d): return math.hypot((a - c) * 111, (b - d) * 111 * math.cos(math.radians(a)))

def read_key(path):
    t = open(path, encoding="utf-8-sig").read().strip()
    try:
        j = json.loads(t)
        for k in ("access_token", "token", "api_key", "apiKey", "key"):
            if isinstance(j, dict) and j.get(k): return str(j[k]).strip()
    except Exception: pass
    toks = re.findall(r"[A-Za-z0-9_\-\.=+/]{16,}", t)
    if not toks: sys.exit(f"Could not find a key in {path}")
    return toks[-1] if len(toks) > 1 else toks[0]

def find_points(o):
    """yield (lat, lng, label) from any LDS-like response shape"""
    if isinstance(o, dict):
        la = o.get("latitude", o.get("lat")); ln = o.get("longitude", o.get("lng", o.get("lon")))
        if isinstance(la, (int, float)) and isinstance(ln, (int, float)):
            yield float(la), float(ln), o.get("formattedAddress") or o.get("formatted_address") or ""; return
        c = o.get("coordinates")
        if isinstance(c, list) and len(c) >= 2 and all(isinstance(x, (int, float)) for x in c[:2]):
            yield float(c[1]), float(c[0]), o.get("formattedAddress") or ""; return        # LDS: longitude first
        for v in o.values(): yield from find_points(v)
    elif isinstance(o, list):
        for v in o: yield from find_points(v)

class Carto:
    def __init__(self, host, token, debug=False):
        self.host, self.debug = host.rstrip("/"), debug
        self.s = requests.Session(); self.s.headers["Authorization"] = "Bearer " + token
    def get(self, path, **params):
        for attempt in range(3):
            r = self.s.get(self.host + path, params=params, timeout=40)
            if r.status_code == 429: time.sleep(2 + attempt * 3); continue
            break
        if self.debug: print("   ", r.status_code, r.url, r.text[:300].replace("\n", " "))
        return r
    def geocode(self, q):
        r = self.get("/v3/lds/geocoding/geocode", address=q)
        if r.status_code != 200: return r.status_code, [], r.text[:300]
        try: return 200, list(find_points(r.json())), ""
        except Exception as e: return 200, [], f"unreadable response: {e}: {r.text[:200]}"

def explain(code, body):
    if code in (401, 403): return ("The key was rejected (HTTP %d). A Basemaps/tile key cannot geocode: create an API access token in CARTO "
                                   "(Developers > Credentials > API access tokens) that allows the LDS API, and make sure a geocoding provider is enabled for your organisation. Server said: %s" % (code, body))
    if code == 400: return "HTTP 400 — often 'Unsupported provider': no geocoding provider is configured for this CARTO organisation. Server said: " + body
    if code == 404: return "HTTP 404 — wrong --host? Use the API host of your CARTO account (gcp-us-east1 / gcp-europe-west1 / ...). Server said: " + body
    return f"HTTP {code}: {body}"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", default="Adjusted_Coordinates_Map_Check.xlsx"); ap.add_argument("--html", default="index.html")
    ap.add_argument("--key-file", default="Carto Map API Key.txt"); ap.add_argument("--host", default="https://gcp-us-east1.api.carto.com")
    ap.add_argument("--test", action="store_true"); ap.add_argument("--debug", action="store_true")
    ap.add_argument("--max-queries", type=int, default=2, help="name variants tried per estate (each uses LDS quota)")
    ap.add_argument("--limit", type=int, help="only the first N rows (trial run)")
    a = ap.parse_args()
    api = Carto(a.host, read_key(a.key_file), a.debug)

    try: st = api.get("/v3/lds/stats")
    except Exception as e: sys.exit(f"Cannot reach {a.host}: {e}\n(proxy / firewall? try --host with your account's region)")
    print("quota check:", st.status_code, st.text[:200].replace("\n", " "))
    code, pts, body = api.geocode("Jumbo Plaza, Sheung Shui, Hong Kong")
    print("test geocode:", code, pts[:2] if pts else body)
    if code != 200 or not pts: sys.exit("\n" + explain(code, body))
    if a.test: print("\nOK — the key can geocode. Run again without --test."); return

    import pandas as pd
    df = pd.read_excel(a.xlsx, sheet_name="Adjusted List")
    if a.limit: df = df.head(a.limit)
    html = open(a.html, encoding="utf-8").read()
    m = re.search(r"/\*MAP_FIX_START\*/window\.__MAP_FIX__=(\{.*?\});/\*MAP_FIX_END\*/", html, re.S)
    if not m: sys.exit(f"{a.html} has no MAP_FIX block — use the updated index.html (v37.0) first. Nothing changed.")
    fix = json.loads(m.group(1)); log = []; ok = 0
    for i, r in df.iterrows():
        code_ = str(r["Estate Code"]); zl = float(r["Zone limit (km)"]) if pd.notna(r["Zone limit (km)"]) else 2.0
        zc = (float(r["Zone ref. Latitude"]), float(r["Zone ref. Longitude"])) if pd.notna(r["Zone ref. Latitude"]) else None
        qs = [q for q in dict.fromkeys([f"{r['Estate Name (ZH)']} 香港", f"{r['Estate Name (EN)']}, Hong Kong", str(r['Estate Name (EN)'])]) if q and "nan" not in q][:a.max_queries]
        got = None; status = "no result inside zone"
        for q in qs:
            c, pts, body = api.geocode(q)
            if c != 200: status = explain(c, body)[:120]; 
            if c in (401, 403): sys.exit("\n" + explain(c, body))
            for la, ln, lab in pts:
                d = km(la, ln, *zc) if zc else 0
                if in_hk(la, ln) and d <= zl: got = (la, ln, d, q, lab); break
            if got: break
            time.sleep(0.15)
        if got:
            fix[code_] = [round(got[0], 6), round(got[1], 6), "carto", f"CARTO LDS: {got[4]}"[:90]]; ok += 1
            log.append([code_, got[3], got[0], got[1], round(got[2], 2), "accepted"])
        else: log.append([code_, "|".join(qs), "", "", "", status])
        if (i + 1) % 20 == 0: print(f"  {i + 1}/{len(df)} rows · {ok} corrected")
    block = "/*MAP_FIX_START*/window.__MAP_FIX__=" + json.dumps(fix, ensure_ascii=False) + ";/*MAP_FIX_END*/"
    open(a.html, "w", encoding="utf-8").write(html[:m.start()] + block + html[m.end():])
    with open("carto_geocode_results.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["Estate Code", "query", "lat", "lng", "km from zone centre", "status"]); w.writerows(log)
    print(f"\nDone: {ok}/{len(df)} corrected via CARTO -> {a.html} (total overrides {len(fix)}); details in carto_geocode_results.csv")
    print("Reload the dashboard; rows not accepted keep their zone-centre approximation.")

if __name__ == "__main__": main()
