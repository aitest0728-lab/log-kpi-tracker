#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Download the Tableau layout-A parent map (AI Fetching for Order Date = T-1).

Uses the previously confirmed workbook view from rolling_orders.py v11.14:
  https://inhouse-analytics.hktv.com.hk/#/views/AIFetchingforOrderDateT-1/AIFetchingforOrderDateT-1_1
Login: TABLEAU_USER / TABLEAU_PASS from the KPI .env (never printed).
"""
from __future__ import annotations

import csv
import io
import os
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

HK = ZoneInfo("Asia/Hong_Kong")

KPI_DIR = Path(
    os.environ.get(
        "KPI_DIR",
        r"C:\Users\chipanl\Downloads\Whatsapp Session\log-kpi-tracker",
    )
)
SIGNIN_URL = "https://inhouse-analytics.hktv.com.hk/#/signin"
# Confirmed layout-A view (v11.14). Do not invent a different URL.
PARENT_MAP_URL = os.environ.get(
    "ROLLING_TABLEAU_PARENT_MAP_URL",
    "https://inhouse-analytics.hktv.com.hk/#/views/AIFetchingforOrderDateT-1/AIFetchingforOrderDateT-1_1",
)
VIEW_TITLE = "AI Fetching for Order Date = T-1"
# Delivery-date start that matches the existing map's span (2026-11-01 .. 2026-11-21).
START_DATE = date.fromisoformat(os.environ.get("ROLLING_TABLEAU_PARENT_MAP_START", "2026-11-01"))
SETTLE_S = int(os.environ.get("ROLLING_TABLEAU_SETTLE_WAIT", "15"))
ATTEMPTS = max(1, int(os.environ.get("TABLEAU_DOWNLOAD_ATTEMPTS", "3")))
RECOVERY_S = int(os.environ.get("TABLEAU_RECOVERY_WAIT_S", "20"))
LOAD_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_LOAD_WAIT", "180"))
REFRESH_WAIT_S = int(os.environ.get("ROLLING_TABLEAU_REFRESH_WAIT", "300"))
HEADLESS = os.environ.get("ROLLING_TABLEAU_HEADLESS", "1").strip().lower() not in ("0", "false", "no")

OUT_PRIMARY = Path(
    os.environ.get(
        "TABLEAU_PARENT_MAP_OUT",
        r"C:\Users\chipanl\Downloads\AI Fetching for Order Date = T-1_data.csv",
    )
)
LAYOUT_A_PARENT = "Order (Unlimited & Add-on Group to Parent)"
MIN_BYTES = 200


def log(msg: str) -> None:
    print(f"[{datetime.now(HK):%Y-%m-%d %H:%M:%S} HKT] {msg}", flush=True)


def load_env() -> None:
    """Load KEY=VALUE from env files without overriding a real environment variable."""
    candidates = [
        KPI_DIR / ".env",
        KPI_DIR / "rolling_orders.env",
        Path(__file__).resolve().parent / ".env",
        Path(__file__).resolve().parent.parent / ".env",
    ]
    for path in candidates:
        if not path.is_file():
            continue
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except OSError:
            continue
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = val
        log(f"Loaded env keys from {path.name} (values not logged)")


def _frames(page):
    return [page, *page.frames]


def _fast_click(page, selectors, timeout_ms=5000) -> bool:
    per = max(300, timeout_ms // max(1, len(selectors)))
    end = time.time() + timeout_ms / 1000.0
    while time.time() < end:
        for target in _frames(page):
            for sel in selectors:
                try:
                    target.locator(sel).first.click(timeout=per)
                    return True
                except Exception:
                    continue
    return False


def _smart_click(page, selectors, timeout_sec=15) -> bool:
    end = time.time() + timeout_sec
    while time.time() < end:
        for target in _frames(page):
            for sel in selectors:
                try:
                    loc = target.locator(sel).first
                    if loc.is_visible():
                        loc.click(force=True)
                        return True
                except Exception:
                    continue
        time.sleep(1)
    return False


class VizActivity:
    def __init__(self, page):
        self.count = 0
        self._marked = 0
        page.on("request", self._on_req)

    def _on_req(self, req):
        try:
            url = req.url or ""
        except Exception:
            return
        if "/vizql/" in url:
            self.count += 1

    def mark(self):
        self._marked = self.count

    def refreshed(self) -> bool:
        return self.count > self._marked + 2


def screenshot(page, path: Path, note: str) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        page.screenshot(path=str(path), full_page=False)
        log(f"Screenshot saved: {path} ({note}; url={page.url})")
    except Exception as ex:
        log(f"Screenshot failed ({type(ex).__name__})")


def read_date_box(page) -> str:
    for target in _frames(page):
        try:
            loc = target.locator('[dojoattachpoint="domLowerText"]').first
            if loc.count():
                return (loc.inner_text(timeout=2000) or "").strip()
        except Exception:
            continue
    return ""


def dates_match(shown: str, want: date) -> bool:
    if not shown:
        return False
    norm = shown.lstrip("0").replace("/0", "/")
    want_s = want.strftime("%d/%m/%Y").lstrip("0").replace("/0", "/")
    return norm == want_s


def set_start_date(page, start: date) -> bool:
    want = start.strftime("%d/%m/%Y")
    for attempt in range(1, 4):
        if not _fast_click(
            page,
            ['[dojoattachpoint="domLowerText"]', "div.readoutText.tab-ctrl-formatted-fixedsize"],
            15000,
        ):
            log(f"Start-date box not found (attempt {attempt}/3)")
            time.sleep(5)
            continue
        time.sleep(0.5)
        done = False
        for target in _frames(page):
            try:
                inp = target.locator('input[dojoattachpoint="domLowerInput"]').first
                if not inp.count():
                    continue
                inp.click(timeout=3000, force=True)
                inp.press("Control+A")
                inp.fill(want)
                inp.press("Enter")
                done = True
                break
            except Exception:
                continue
        time.sleep(3)
        shown = read_date_box(page)
        log(f"Start date typed {want} (attempt {attempt}/3) -> box now shows {shown!r}")
        if done and dates_match(shown, start):
            return True
    return False


def view_is_up(page) -> bool:
    try:
        if page.locator('iframe[title="Data Visualisation"]').count():
            return True
        if page.locator("#download, [aria-label='Download']").count():
            return True
        for f in page.frames:
            if f.locator('[dojoattachpoint="domLowerText"]').count():
                return True
            if f.locator("#download, [aria-label='Download'], canvas").count():
                return True
    except Exception:
        return False
    return False


def wait_for_view(page, shot: Path) -> None:
    end = time.time() + LOAD_WAIT_S
    last_goto = time.time()
    while time.time() < end:
        if view_is_up(page):
            log("View loaded")
            return
        if time.time() - last_goto > 90:
            log(f"View not up yet (at {page.url}) - opening it again")
            try:
                page.goto(PARENT_MAP_URL)
            except Exception:
                pass
            last_goto = time.time()
        time.sleep(3)
    screenshot(page, shot, "view timeout")
    raise RuntimeError(f"Tableau view did not load within {LOAD_WAIT_S}s (page at {page.url})")


def check_title(page) -> None:
    seen = []
    try:
        seen.append(page.title() or "")
    except Exception:
        pass
    for target in _frames(page):
        try:
            for loc in target.locator("h1, [title], [data-tb-test-id='view-name']").all():
                t = (loc.inner_text(timeout=1000) or loc.get_attribute("title") or "").strip()
                if t:
                    seen.append(t)
        except Exception:
            continue
    log(f"Page titles/headings seen: {seen[:8]!r}")
    if any(VIEW_TITLE.lower() in (t or "").lower() for t in seen):
        log(f"View title OK ({VIEW_TITLE!r} found)")
    else:
        log(f"WARNING: title {VIEW_TITLE!r} not clearly found; continuing and checking layout after download")


def log_crosstab_sheets(page) -> list[str]:
    names = []
    for target in _frames(page):
        try:
            thumbs = target.locator("[role='dialog'] [role='radio'], [role='dialog'] button, [class*='thumbnail']")
            n = min(thumbs.count(), 20)
            for i in range(n):
                loc = thumbs.nth(i)
                label = (loc.inner_text(timeout=500) or loc.get_attribute("aria-label") or loc.get_attribute("title") or "").strip()
                selected = False
                try:
                    selected = bool(loc.get_attribute("aria-checked") == "true" or loc.get_attribute("aria-selected") == "true")
                except Exception:
                    pass
                if label:
                    names.append(f"{label}{' [selected]' if selected else ''}")
        except Exception:
            continue
        if names:
            break
    log(f"Crosstab sheet thumbnails: {names or '(none listed)'}")
    return names


def press_crosstab_download(page, shot_base: Path):
    selectors = [
        ("export-crosstab-export-Button", lambda r: r.locator('[data-tb-test-id="export-crosstab-export-Button"]').first),
        ("aria-label Download Crosstab", lambda r: r.locator("button[aria-label='Download Crosstab']").first),
        ("dialog button named Download", lambda r: r.locator("[role='dialog']").get_by_role("button", name=re.compile(r"^\s*download", re.I)).last),
        ("aria-label Download in dialog", lambda r: r.locator("[role='dialog'] button[aria-label='Download']").last),
    ]
    last = None
    for label, factory in selectors:
        for target in _frames(page):
            try:
                loc = factory(target)
                if loc.count() == 0:
                    continue
                log(f"Pressing the dialog Download button via {label!r}")
                with page.expect_download(timeout=240000) as info:
                    loc.click(timeout=5000)
                return info.value
            except Exception as ex:
                last = ex
                continue
    screenshot(page, shot_base.with_name(shot_base.stem + "_download_btn.png"), "no download button worked")
    raise RuntimeError(f"no Download button of the Crosstab dialog started a download ({type(last).__name__ if last else 'none'})")


def detect_layout_and_span(path: Path) -> dict:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff") or (len(raw) > 40 and raw[1:2] == b"\x00"):
        text = raw.decode("utf-16")
        enc = "utf-16"
    else:
        text = raw.decode("utf-8-sig", errors="replace")
        enc = "utf-8"
    # Sniff delimiter from header
    first = text.splitlines()[0] if text else ""
    delim = "\t" if first.count("\t") > first.count(",") else ","
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    try:
        header = next(reader)
    except StopIteration:
        raise ValueError("empty CSV")
    header = [h.strip().lstrip("\ufeff") for h in header]
    is_a = LAYOUT_A_PARENT in header and any(h.lower() == "order_number" or h.lower() == "order number" for h in header)
    is_b = any("# of Parent Order" in h for h in header) or (
        "delivery_date" in [h.lower() for h in header] and any("# of Order" in h for h in header)
    )
    layout = "A" if is_a and not is_b else ("B" if is_b else "unknown")
    # Date span for layout A
    od_i = next((i for i, h in enumerate(header) if "order date" in h.lower()), None)
    dd_i = next((i for i, h in enumerate(header) if "delivery" in h.lower() and "date" in h.lower()), None)
    rows = 0
    ods, dds = [], []
    for row in reader:
        rows += 1
        if od_i is not None and od_i < len(row) and row[od_i]:
            ods.append(row[od_i])
        if dd_i is not None and dd_i < len(row) and row[dd_i]:
            dds.append(row[dd_i])
    return {
        "encoding": enc,
        "delimiter": delim,
        "header": header,
        "layout": layout,
        "n_rows": rows,
        "order_date_min": min(ods) if ods else None,
        "order_date_max": max(ods) if ods else None,
        "delivery_date_min": min(dds) if dds else None,
        "delivery_date_max": max(dds) if dds else None,
        "size": path.stat().st_size,
    }


def session(final_path: Path, user: str, password: str) -> Path:
    from playwright.sync_api import sync_playwright

    tmp = final_path.with_suffix(final_path.suffix + ".part")
    shot_base = final_path
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, args=["--disable-popup-blocking"])
        try:
            context = browser.new_context(accept_downloads=True, viewport={"width": 1920, "height": 1080})
            page = context.new_page()
            log("Tableau: login (TABLEAU_USER)")
            page.goto(SIGNIN_URL)
            page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
            page.locator("input[type='text'], input[name='username']").first.fill(user)
            page.locator("input[type='password'], input[name='password']").first.fill(password)
            page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()
            try:
                page.wait_for_function("() => !location.hash.includes('signin')", timeout=60000)
            except Exception:
                log(f"Still on {page.url} after sign-in - continuing")
            time.sleep(5)
            if "#/user/" in (page.url or "") or "#/explore" in (page.url or "") or "signin" not in (page.url or "").lower():
                log(f"Post-login page: {page.url}")
            log(f"Opening parent-map view {PARENT_MAP_URL}")
            page.goto(PARENT_MAP_URL)
            wait_for_view(page, shot_base.with_name(shot_base.stem + "_view_timeout.png"))
            # Forced full reload so the toolbar handlers attach.
            log("Forced full reload of the view")
            page.reload(wait_until="load")
            wait_for_view(page, shot_base.with_name(shot_base.stem + "_reload_timeout.png"))
            if "#/user/" in (page.url or ""):
                log("Bounced back to /#/user/ - opening the view and reloading again")
                page.goto(PARENT_MAP_URL)
                page.reload(wait_until="load")
                wait_for_view(page, shot_base.with_name(shot_base.stem + "_bounce_timeout.png"))
            time.sleep(SETTLE_S)
            check_title(page)
            activity = VizActivity(page)
            activity.mark()
            shown_before = read_date_box(page)
            log(f"Start-date box before typing: {shown_before!r}")
            if not set_start_date(page, START_DATE):
                raise RuntimeError(f"could not set the delivery_date start date to {START_DATE:%d/%m/%Y}")
            log(f"Waiting for the view to refresh (up to {REFRESH_WAIT_S}s) ...")
            t0 = time.time()
            while time.time() - t0 < REFRESH_WAIT_S:
                page.wait_for_timeout(2000)
                if activity.refreshed():
                    log(f"View refreshed after {time.time() - t0:.0f}s ({activity.count} vizql requests)")
                    break
            else:
                log(f"No refresh completion detected - waited the full {REFRESH_WAIT_S}s")
            shown = read_date_box(page)
            if shown and not dates_match(shown, START_DATE):
                raise RuntimeError(f"start date reverted to {shown!r} during the refresh wait")
            # Download -> Crosstab
            if not (_fast_click(page, ["#download", "[aria-label='Download']", "button:has-text('Download')"], 4000) or _smart_click(page, ["#download", "[aria-label='Download']", "button:has-text('Download')"], 15)):
                screenshot(page, shot_base.with_name(shot_base.stem + "_no_download_icon.png"), "download icon missing")
                raise RuntimeError("Download button not found")
            time.sleep(1)
            ct_sel = [
                '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
                "#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
                "span:has-text('Crosstab')",
                "text=Crosstab",
            ]
            opened = _fast_click(page, ct_sel, 4000) or _smart_click(page, ct_sel, 10)
            if not opened:
                # Last resort: keyboard Down, Down, Enter (Image / Data / Crosstab)
                log("Crosstab not found in menu - trying keyboard Down, Down, Enter")
                page.keyboard.press("ArrowDown")
                time.sleep(0.2)
                page.keyboard.press("ArrowDown")
                time.sleep(0.2)
                page.keyboard.press("Enter")
                time.sleep(1)
            time.sleep(2)
            log_crosstab_sheets(page)
            # Do not click a thumbnail: it is a toggle.
            csv_sel = [
                "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
                "label:has-text('CSV')",
                "text=CSV",
            ]
            _fast_click(page, csv_sel, 2000) or _smart_click(page, csv_sel, 5)
            time.sleep(0.5)
            # Verify CSV radio if possible
            try:
                for target in _frames(page):
                    radio = target.locator("input[type='radio'][value='csv'], label:has-text('CSV') input").first
                    if radio.count():
                        checked = radio.is_checked()
                        log(f"CSV radio checked={checked}")
                        if not checked:
                            radio.check()
                        break
            except Exception:
                pass
            dl = press_crosstab_download(page, shot_base)
            suggested = dl.suggested_filename
            log(f"Tableau suggested file name: {suggested}")
            tmp.parent.mkdir(parents=True, exist_ok=True)
            dl.save_as(str(tmp))
        finally:
            browser.close()
    return tmp, suggested if "suggested" in dir() else ""


# press_crosstab_download returns only the download; keep suggested on the session
# Fix: session currently may not return suggested if exception after assign. Handle below.


def download_once(final_path: Path, user: str, password: str) -> tuple[Path, str]:
    # re-implement return of suggested name via a small wrapper using a mutable holder
    holder = {"name": ""}

    def _run():
        from playwright.sync_api import sync_playwright

        tmp = final_path.with_suffix(final_path.suffix + ".part")
        suggested = ""
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=HEADLESS, args=["--disable-popup-blocking"])
            try:
                context = browser.new_context(accept_downloads=True, viewport={"width": 1920, "height": 1080})
                page = context.new_page()
                log("Tableau: login (TABLEAU_USER)")
                page.goto(SIGNIN_URL)
                page.wait_for_selector("input[type='text'], input[name='username']", timeout=30000)
                page.locator("input[type='text'], input[name='username']").first.fill(user)
                page.locator("input[type='password'], input[name='password']").first.fill(password)
                page.locator("button:has-text('Sign In'), [aria-label='Sign In']").first.click()
                try:
                    page.wait_for_function("() => !location.hash.includes('signin')", timeout=60000)
                except Exception:
                    log(f"Still on {page.url} after sign-in - continuing")
                time.sleep(5)
                log(f"Post-login page: {page.url}")
                log(f"Opening parent-map view {PARENT_MAP_URL}")
                page.goto(PARENT_MAP_URL)
                wait_for_view(page, final_path.with_name(final_path.stem + "_view_timeout.png"))
                log("Forced full reload of the view")
                page.reload(wait_until="load")
                wait_for_view(page, final_path.with_name(final_path.stem + "_reload_timeout.png"))
                if "#/user/" in (page.url or ""):
                    log("Bounced back to /#/user/ - opening the view and reloading again")
                    page.goto(PARENT_MAP_URL)
                    page.reload(wait_until="load")
                    wait_for_view(page, final_path.with_name(final_path.stem + "_bounce_timeout.png"))
                time.sleep(SETTLE_S)
                check_title(page)
                activity = VizActivity(page)
                activity.mark()
                shown_before = read_date_box(page)
                log(f"Start-date box before typing: {shown_before!r}")
                if not set_start_date(page, START_DATE):
                    raise RuntimeError(f"could not set the delivery_date start date to {START_DATE:%d/%m/%Y}")
                log(f"Waiting for the view to refresh (up to {REFRESH_WAIT_S}s) ...")
                t0 = time.time()
                while time.time() - t0 < REFRESH_WAIT_S:
                    page.wait_for_timeout(2000)
                    if activity.refreshed():
                        log(f"View refreshed after {time.time() - t0:.0f}s ({activity.count} vizql requests)")
                        break
                else:
                    log(f"No refresh completion detected - waited the full {REFRESH_WAIT_S}s")
                shown = read_date_box(page)
                if shown and not dates_match(shown, START_DATE):
                    raise RuntimeError(f"start date reverted to {shown!r} during the refresh wait")
                if not (
                    _fast_click(page, ["#download", "[aria-label='Download']", "button:has-text('Download')"], 4000)
                    or _smart_click(page, ["#download", "[aria-label='Download']", "button:has-text('Download')"], 15)
                ):
                    screenshot(page, final_path.with_name(final_path.stem + "_no_download_icon.png"), "download icon missing")
                    raise RuntimeError("Download button not found")
                time.sleep(1)
                ct_sel = [
                    '[data-tb-test-id="download-crosstab-Button-MenuItem"]',
                    "#viz-viewer-toolbar-download-menu > div:nth-of-type(3)",
                    "span:has-text('Crosstab')",
                    "text=Crosstab",
                ]
                opened = _fast_click(page, ct_sel, 4000) or _smart_click(page, ct_sel, 10)
                if not opened:
                    log("Crosstab not found in menu - trying keyboard Down, Down, Enter")
                    page.keyboard.press("ArrowDown")
                    time.sleep(0.2)
                    page.keyboard.press("ArrowDown")
                    time.sleep(0.2)
                    page.keyboard.press("Enter")
                    time.sleep(1)
                time.sleep(2)
                log_crosstab_sheets(page)
                csv_sel = [
                    "#export-crosstab-options-dialog-Dialog-BodyWrapper-Dialog-Body-Id label:nth-of-type(2) input",
                    "label:has-text('CSV')",
                    "text=CSV",
                ]
                _fast_click(page, csv_sel, 2000) or _smart_click(page, csv_sel, 5)
                time.sleep(0.5)
                try:
                    for target in _frames(page):
                        radio = target.locator("input[type='radio'][value='csv'], label:has-text('CSV') input").first
                        if radio.count():
                            checked = radio.is_checked()
                            log(f"CSV radio checked={checked}")
                            if not checked:
                                radio.check()
                            break
                except Exception:
                    pass
                dl = press_crosstab_download(page, final_path)
                suggested = dl.suggested_filename
                log(f"Tableau suggested file name: {suggested}")
                tmp.parent.mkdir(parents=True, exist_ok=True)
                dl.save_as(str(tmp))
            finally:
                browser.close()
        holder["name"] = suggested
        return tmp

    return _run(), holder["name"]


def main() -> int:
    load_env()
    # TABLEAU_USER can sign in but gets "Permission Required" on this parent-map
    # workbook (confirmed this run). v11.14 opened the same view with the GMV account.
    user = os.environ.get("TABLEAU_GMV_USER", "").strip()
    password = os.environ.get("TABLEAU_GMV_PASS", "").strip()
    account = "GMV"
    if not user or not password:
        user = os.environ.get("TABLEAU_USER", "").strip()
        password = os.environ.get("TABLEAU_PASS", "").strip()
        account = "TABLEAU_USER"
    if not user or not password:
        log(f"Tableau credentials not set (looked in {KPI_DIR / '.env'} and rolling_orders.env)")
        return 2
    log(f"Using the {account} Tableau account for the parent-map view (password not logged)")

    final_path = OUT_PRIMARY
    last = None
    for n in range(1, ATTEMPTS + 1):
        if n > 1:
            log(f"Retry {n}/{ATTEMPTS} (fresh browser + login)")
            time.sleep(RECOVERY_S)
        try:
            log(f"Attempt {n}/{ATTEMPTS}: download layout-A parent map")
            tmp, suggested = download_once(final_path, user, password)
            size = tmp.stat().st_size if tmp.exists() else 0
            log(f"Downloaded {size} bytes to {tmp.name}")
            if size < MIN_BYTES:
                bad = tmp.with_suffix(tmp.suffix + ".bad")
                tmp.replace(bad)
                raise RuntimeError(f"file smaller than {MIN_BYTES} bytes (kept as {bad.name})")
            info = detect_layout_and_span(tmp)
            log(
                f"Layout={info['layout']} enc={info['encoding']} delim={info['delimiter']!r} "
                f"rows={info['n_rows']:,} header={info['header']!r}"
            )
            log(
                f"Order-date span: {info['order_date_min']} .. {info['order_date_max']}; "
                f"delivery-date span: {info['delivery_date_min']} .. {info['delivery_date_max']}"
            )
            if info["layout"] != "A":
                bad = tmp.with_name(tmp.name + ".bad")
                tmp.replace(bad)
                raise RuntimeError(f"layout mismatch (got {info['layout']}, expected A); kept as {bad}")
            # Replace previous dated copy only after layout check.
            if final_path.exists():
                bak = final_path.with_name(final_path.name + ".bak")
                try:
                    if bak.exists():
                        bak.unlink()
                    final_path.replace(bak)
                    log(f"Previous file moved to {bak.name}")
                except OSError as ex:
                    log(f"Could not keep a .bak ({type(ex).__name__}); overwriting in place")
            tmp.replace(final_path)
            # Also write a sidecar stamp
            stamp = final_path.with_suffix(".export_meta.txt")
            now = datetime.now(HK)
            stamp.write_text(
                "\n".join(
                    [
                        f"exported_at_hkt={now:%Y-%m-%d %H:%M:%S}",
                        f"suggested_name={suggested}",
                        f"path={final_path}",
                        f"size={final_path.stat().st_size}",
                        f"layout={info['layout']}",
                        f"n_rows={info['n_rows']}",
                        f"order_date_span={info['order_date_min']} .. {info['order_date_max']}",
                        f"delivery_date_span={info['delivery_date_min']} .. {info['delivery_date_max']}",
                        f"view_url={PARENT_MAP_URL}",
                        f"start_date={START_DATE.isoformat()}",
                        f"encoding={info['encoding']}",
                    ]
                )
                + "\n",
                encoding="utf-8",
            )
            log(f"OK saved {final_path} ({final_path.stat().st_size} bytes) suggested={suggested!r}")
            return 0
        except Exception as ex:
            last = ex
            log(f"Attempt {n}/{ATTEMPTS} failed: {type(ex).__name__}: {ex}")
    log(f"Parent-map download failed after {ATTEMPTS} sessions: {last}")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
