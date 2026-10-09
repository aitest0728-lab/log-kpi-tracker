#!/usr/bin/env python3
"""
delay_report.py  (KPI Dashboard v9.1)

Builds the two things the 08:30 "Daily delay rate" WhatsApp post needs, from the dashboard's own files:

  1. build_caption(date, record)      -> the caption text
       Daily update on Oct 6 delay rate by 10區 🙇‍♀️

       Overall 9.4% delay

       Districts with over 10% delay in at least one timeslot:
       • ETX 18.2% - PM 24.4%, AM 21.7%
       ...
  2. render_delay_screenshot(...)     -> PNG of the Delay % tab: header + the T-1 row expanded into AM/PM/EV/EV2
       (same look as the sample: DATE | ETH ... WTX).

No import of kpi_pipeline (that module insists on Tableau credentials at import time), so it can also be run on
its own to test / re-render without touching Tableau or WhatsApp:

    python3 delay_report.py                      # newest day in public/delay_history.json -> logs/delay_report/
    python3 delay_report.py --date 2026-10-06    # a specific day
    python3 delay_report.py --save-login URL     # one-time: log in to a hosted dashboard, keep the session

Screenshot source (SHOT mode):
  * default (local): public/index.html is served from a throw-away localhost web server, so the picture always
    shows exactly the files the pipeline has just written and needs no login at all.
  * DASHBOARD_URL set: the hosted (login-protected) dashboard is opened instead, with the saved login session in
    DASHBOARD_STORAGE_STATE (create it once with --save-login). Only worth it if the picture must come from the
    deployed copy; the local mode shows identical numbers.
"""
import argparse
import functools
import http.server
import json
import os
import sys
import threading
from datetime import date, datetime, timedelta
from pathlib import Path

DISTRICTS = ["ETH", "ETK", "ETX", "NT-ST", "NT-TM", "NT-TSM", "NT-TW", "WTH", "WTK", "WTX"]
SLOTS = ["AM", "PM", "EV", "EV2"]
DEFAULT_TARGET = 4.0            # same flat network-wide MTD delay target the dashboard uses (DELAY_RATE_TARGET)
PINPOINT_THRESHOLD = float(os.environ.get("DELAY_PINPOINT_THRESHOLD", "10"))   # a district is named only if one timeslot is over this %

PUBLIC_DIR = Path(os.environ.get("PUBLIC_DIR", "./public"))
OUT_DIR = Path(os.environ.get("DELAY_REPORT_DIR", "./logs/delay_report"))
DASHBOARD_URL = os.environ.get("DASHBOARD_URL", "").strip()
DASHBOARD_STORAGE_STATE = os.environ.get("DASHBOARD_STORAGE_STATE", "./dashboard_login_state.json")


# --------------------------------------------------------------------------------------------- caption
def _pct(v):
    return f"{v:.1f}%"


def _delay(rec, district, slot):
    return ((rec.get("districts") or {}).get(district) or {}).get(slot, {}).get("delay")


def build_caption(day, rec, threshold=PINPOINT_THRESHOLD):
    """`day` is a date (the T-1 date); `rec` is delay_history.json's daily[day] record.
    A district is pinpointed only when at least ONE of its timeslots (AM/PM/EV/EV2) is over `threshold` % delay; the
    line then names the district's overall delay and every timeslot over the threshold, worst first."""
    overall = ((rec.get("overall") or {}).get("Overall") or {}).get("delay")
    if overall is None:
        raise ValueError(f"No overall delay rate on record for {day}")
    label = f"{day:%b} {day.day}"                      # "Oct 6"
    lines = [f"Daily update on {label} delay rate by 10區 🙇‍♀️", "", f"Overall {_pct(overall)} delay", ""]

    flagged = []
    for d in DISTRICTS:
        hot = [(s, _delay(rec, d, s)) for s in SLOTS]
        hot = sorted([(s, x) for s, x in hot if x is not None and x > threshold], key=lambda t: -t[1])
        if hot:
            flagged.append((d, _delay(rec, d, "Overall"), hot))
    flagged.sort(key=lambda t: -(t[1] if t[1] is not None else t[2][0][1]))      # worst district overall first

    if flagged:
        lines.append(f"Districts with over {threshold:g}% delay in at least one timeslot:")
        for d, v, hot in flagged:
            head = f"• {d} {_pct(v)} - " if v is not None else f"• {d} - "
            lines.append(head + ", ".join(f"{s} {_pct(x)}" for s, x in hot))
    else:
        lines.append(f"No district is over {threshold:g}% delay in any timeslot 🎉")
    return "\n".join(lines)


# ----------------------------------------------------------------------------------------- screenshot
class _Quiet(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *a):
        pass


def _serve(directory):
    handler = functools.partial(_Quiet, directory=str(directory))
    srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv


# Keep only the T-1 row and its sub-rows (AM/PM/EV/EV2 and the bottom 'Last Received' row - all tr.delay-subrow); the Total column (overall delay rate) is kept unless hideTotal.
_TRIM_JS = """([date, hideTotal]) => {
    const rows = [...document.querySelectorAll('#delayTableBody tr')];
    let keep = false;
    for (const tr of rows) {
        if (tr.dataset.date) keep = (tr.dataset.date === date);
        else if (!tr.classList.contains('delay-subrow')) keep = false;
        tr.style.display = keep ? '' : 'none';
    }
    if (hideTotal) {
        for (const tr of document.querySelectorAll('#delayTable tr')) {
            const last = tr.lastElementChild; if (last && tr.style.display !== 'none') last.style.display = 'none';
        }
    }
    document.querySelectorAll('.table-scroll,.table-wrap').forEach(e => { e.style.overflow = 'visible'; e.style.maxHeight = 'none'; });
    document.querySelectorAll('.scroll-fade,.scroll-hint').forEach(e => e.style.display = 'none');
    return rows.filter(tr => tr.style.display !== 'none').length;
}"""


def render_delay_screenshot(day, out_path, hide_total=False, url=None, expect_lr=False):
    """Screenshot the Delay % table with `day` expanded. Raises if the day's row is not on the dashboard.
    expect_lr=True (the day's record has a lastReceived block): the bottom "Last Received" row (date + time per district) must be
    in the picture too, otherwise this raises and prepare() falls back to the data-drawn table, which always includes it."""
    from playwright.sync_api import sync_playwright
    day_s = day.isoformat()
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    url = url or DASHBOARD_URL
    srv = None
    if not url:
        if not (PUBLIC_DIR / "index.html").exists():
            raise FileNotFoundError(f"{PUBLIC_DIR / 'index.html'} not found")
        srv = _serve(PUBLIC_DIR)
        url = f"http://127.0.0.1:{srv.server_address[1]}/index.html"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            try:
                kw = {}
                if DASHBOARD_URL and Path(DASHBOARD_STORAGE_STATE).exists():
                    kw["storage_state"] = DASHBOARD_STORAGE_STATE
                ctx = browser.new_context(device_scale_factor=2, viewport={"width": 1700, "height": 900}, **kw)
                page = ctx.new_page()
                page.goto(url, wait_until="load", timeout=60000)
                page.wait_for_selector('.rail-icon[data-view="delay"]', timeout=60000)
                page.wait_for_timeout(1500)                       # let loadDataSource() finish its fetches
                page.locator('.rail-icon[data-view="delay"]').click()
                page.wait_for_selector("#delayTableBody tr", timeout=30000)
                page.select_option("#delaySlotSelect", "Overall")
                row = page.locator(f'#delayTableBody tr[data-date="{day_s}"]')
                if row.count() == 0:
                    raise RuntimeError(f"The dashboard's Delay % tab has no row for {day_s} "
                                       f"(login page / stale data?). Page title: {page.title()!r}")
                if "delay-row-open" not in (row.first.get_attribute("class") or ""):
                    row.first.click()
                    page.wait_for_selector(f'#delayTableBody tr[data-date="{day_s}"].delay-row-open', timeout=10000)
                n = page.evaluate(_TRIM_JS, [day_s, hide_total])
                need = 6 if expect_lr else 5                      # date row + AM/PM/EV/EV2 (+ Last Received)
                if n < need:
                    raise RuntimeError(f"Expected {need} rows for {day_s}, found {n}"
                                       + (" - the dashboard shows no 'Last Received' row for that day" if expect_lr else ""))
                page.locator("#delayTable").screenshot(path=str(out_path))
            finally:
                browser.close()
    finally:
        if srv:
            srv.shutdown()
    return out_path


def save_login(url):
    """One-time helper: open the hosted dashboard in a visible browser, let the user log in, keep the session."""
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        ctx = browser.new_context()
        ctx.new_page().goto(url)
        input("Log in to the dashboard in the browser window, then press Enter here... ")
        ctx.storage_state(path=DASHBOARD_STORAGE_STATE)
        browser.close()
    print(f"Saved login session to {DASHBOARD_STORAGE_STATE}")


# ------------------------------------------------------------------------------------------ files
def load_day(day):
    path = PUBLIC_DIR / "delay_history.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    rec = (data.get("daily") or {}).get(day.isoformat())
    if not rec:
        raise KeyError(f"{path} has no daily record for {day.isoformat()}")
    return rec


_FALLBACK_CSS = (
    "body{margin:0;background:#fff;font-family:'Noto Sans CJK TC','Microsoft JhengHei',Inter,Arial,sans-serif;font-size:14px}"
    "table{border-collapse:collapse;min-width:1500px}th,td{padding:9px 14px;text-align:right;border:1px solid #e5e7eb}"
    "th{background:#0f172a;color:#cbd5e1;font-size:12px;letter-spacing:.06em}th:first-child,td:first-child{text-align:left}"
    "tr.main td{background:#f1f3f6;font-weight:600}tr.sub td{background:#fafbfc;color:#64748b}tr.sub td:first-child{padding-left:28px;font-weight:600;color:#111}"
    "tr.lr td{background:#fafbfc;color:#64748b;font-size:12.5px;text-align:center;border-top:2px solid #e5e7eb}tr.lr td:first-child{text-align:left;padding-left:28px;font-weight:600;color:#111}")


def render_fallback_table(day, rec, out_path, hide_total=False):
    """Plain re-draw of the same table straight from the record - used only when the dashboard itself has no row
    for `day` (typically the 1st of a month, when T-1 is already last month's closed row)."""
    from playwright.sync_api import sync_playwright
    f = lambda v: "—" if v is None else f"{v:.2f}%"
    head = "<tr><th>DATE</th>" + "".join(f"<th>{d}</th>" for d in DISTRICTS) + ("" if hide_total else "<th>TOTAL</th>") + "</tr>"

    def row(label, slot, cls):
        tot = "" if hide_total else f"<td>{f(((rec.get('overall') or {}).get(slot) or {}).get('delay'))}</td>"
        return (f'<tr class="{cls}"><td>{label}</td>' + "".join(f"<td>{f(_delay(rec, d, slot))}</td>" for d in DISTRICTS) + tot + "</tr>")
    body = row(f"▾ {day.isoformat()}", "Overall", "main") + "".join(row(s, s, "sub") for s in SLOTS)
    lr = rec.get("lastReceived")
    if lr:                                                  # bottom row: date on top, time below (as on the dashboard)
        cell = lambda v: f"<td>{v['date']}<br><b>{v['time']}</b></td>" if v else "<td>—</td>"
        lr_d = lr.get("districts") or {}
        body += ('<tr class="lr"><td>Last Received</td>' + "".join(cell(lr_d.get(d)) for d in DISTRICTS)
                 + ("" if hide_total else cell(lr.get("overall"))) + "</tr>")
    html = f"<html><head><meta charset='utf-8'><style>{_FALLBACK_CSS}</style></head><body><table id='t'>{head}{body}</table></body></html>"
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        try:
            page = browser.new_context(device_scale_factor=2, viewport={"width": 1700, "height": 600}).new_page()
            page.set_content(html)
            page.locator("#t").screenshot(path=str(out_path))
        finally:
            browser.close()
    return Path(out_path)


def prepare(day, rec=None, hide_total=False, threshold=PINPOINT_THRESHOLD):
    """Render screenshot + caption for `day` into OUT_DIR; returns (png_path, caption).
    `rec` (optional) is the day's record; default = delay_history.json's daily[day]."""
    rec = rec or load_day(day)
    caption = build_caption(day, rec, threshold)
    out = Path(OUT_DIR) / f"delay_{day:%Y%m%d}.png"
    try:
        png = render_delay_screenshot(day, out, hide_total=hide_total, expect_lr=bool(rec.get("lastReceived")))
    except Exception as e:                                  # dashboard has no row / page failed to load
        print(f"  ⚠️ dashboard screenshot failed ({e}); drawing the table from the data instead.")
        png = render_fallback_table(day, rec, out, hide_total=hide_total)
    Path(OUT_DIR).mkdir(parents=True, exist_ok=True)
    (Path(OUT_DIR) / f"delay_{day:%Y%m%d}.txt").write_text(caption, encoding="utf-8")
    return png, caption


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", help="YYYY-MM-DD (default: newest day in delay_history.json)")
    ap.add_argument("--hide-total", action="store_true", help="drop the Total column (overall delay rate) from the screenshot")
    ap.add_argument("--save-login", metavar="URL", help="log in to the hosted dashboard once and save the session")
    a = ap.parse_args()
    if a.save_login:
        save_login(a.save_login)
        sys.exit(0)
    if a.date:
        d = datetime.strptime(a.date, "%Y-%m-%d").date()
    else:
        d = date.fromisoformat(max(json.loads((PUBLIC_DIR / "delay_history.json").read_text(encoding="utf-8"))["daily"]))
    png, cap = prepare(d, hide_total=a.hide_total)
    print(cap)
    print(f"\nScreenshot: {png}")
