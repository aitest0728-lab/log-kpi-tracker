#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_orders_backfill.py  -  KPI Dashboard Version 11.11  "Rolling Orders" BACKFILL (whole month)

Re-computes the Rolling Orders figures of every order date from <--from> (default: the 1st of the month of yesterday) to
<--to> (default: yesterday = T-1) with the v11.10 parent-order logic of rolling_orders.py, and writes them into rolling_orders.json
(+ embeds it in index.html and pushes, exactly like the daily job). It uses the SAME processing code as rolling_orders.py
(import), so a backfilled day equals what the 06:00 job would have produced.

What it does
  1. Checks, for every order date in the range, whether
        - the MFNAS csv  "Rolling Order Count (Oct-DD).csv"  is already in "Folder for Rolling Orders", and
        - the day is already in rolling_orders.json,
     and prints that table (use --list to stop here).
  2. Downloads every MISSING MFNAS csv (only for days the Tableau csv covers) (order date D = folder <D+1 as YYYYMMDD> of ECOM-MFNAS_hktvAccumulatedDeliveryLeadTimeReport).
  3. Checks the ONE Tableau report "AI Fetching for Order Date = T-1" that covers the whole range (start date = first day of the
     range, e.g. 01/10/2026; the waybill column is removed so the file stays small). A file already in the folder is used as it is
     (the last order date is NOT forced: days after the latest order date in the file are listed as "not covered"); it is downloaded
     (start date typed, wait for the refresh, Crosstab -> CSV - same steps as rolling_orders.py) only when it is missing or with
     --redownload-tableau. A file the user dropped in by hand is found by its name
     (anything like "AI Fetching for Order Date T-1 ... Oct-1 ...csv" in the folder).
  4. Processes every day IN DATE ORDER (district / CANCEL / KN015 / same-day-slot filters, Order Number -> Parent Order of the same
     delivery date; an order that is not in Tableau is its own parent; a parent that an earlier order date already counted for the
     same delivery date is not counted again - v11.10) and writes the days into rolling_orders.json. The parents every day counted
     are kept in rolling_parent_registry.json.gz, which the 06:00 job continues from - so run this backfill ONCE from the first order
     date whose delivery dates are still open (e.g. --from 2026-10-01) when you install v11.10. Only the "days" of the processed dates are replaced; every other day, the
     snapshots and lastRun are left untouched. A copy of the old json is kept as rolling_orders.json.bak_backfill.
  5. Prints old vs new parent-order totals per order date.
  6. (v11.9) Deletes the csv files it has just written to the json: the MFNAS csv of every written day, and a Tableau csv once no
     failed day still needs it (--keep-files keeps everything; a file given with --tableau-file is never deleted).

Usage (WSL, from the log-kpi-tracker folder):
    python3 rolling_orders_backfill.py --list                       # what is missing? (no download, no write)
    python3 rolling_orders_backfill.py --skip-push                  # whole of this month, check the numbers first
    python3 rolling_orders_backfill.py                              # same + git push to the dashboard
    python3 rolling_orders_backfill.py --from 2026-10-01 --to 2026-10-21
    python3 rolling_orders_backfill.py --only-missing               # only dates that are not in rolling_orders.json yet
    python3 rolling_orders_backfill.py --tableau-file "/path/to/AI Fetching ... (Oct-1).csv"   # use this Tableau csv
    python3 rolling_orders_backfill.py --redownload-tableau --headed   # force a fresh Tableau download, Chrome visible
    python3 rolling_orders_backfill.py --skip-download              # never download, use the files already in the folder
    python3 rolling_orders_backfill.py --skip-tableau-download      # download missing MFNAS csv files, but use the Tableau csv in the folder

Credentials / environment: same as rolling_orders.py (CRAB_NAS_USER / CRAB_NAS_PASS, TABLEAU_GMV_USER / TABLEAU_GMV_PASS,
ROLLING_TABLEAU_REFRESH_WAIT, ROLLING_TABLEAU_LOAD_WAIT, ROLLING_TABLEAU_HEADLESS ...).
"""
import argparse
import glob
import json
import os
import re
import shutil
import sys
import time
import traceback
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import rolling_orders as ro                      # noqa: E402   (same folder; its module-level code only loads the env file)

log = ro.log

if not (hasattr(ro, "read_tableau_csv") and hasattr(ro, "cleanup_files")):          # an older rolling_orders.py sits next to this script
    raise SystemExit("rolling_orders.py in this folder is an older version (it cannot read the comma-separated Tableau csv). "
                     "Replace it with the latest rolling_orders.py (v11.3, the one delivered together with this backfill script).")


# ----------------------------------------------------------------------------- helpers
def _day(s):
    return datetime.strptime(s, "%Y-%m-%d")


def date_range(first, last):
    d = first
    while d <= last:
        yield d
        d += timedelta(days=1)


def load_json():
    if os.path.exists(ro.ROLLING_JSON):
        try:
            return json.load(open(ro.ROLLING_JSON, encoding="utf-8"))
        except Exception as ex:
            log(f"WARNING: {ro.ROLLING_JSON} is unreadable ({ex}) - starting from an empty file")
    return {}


def order_total(cells):
    """Parent orders of one order date = sum over its delivery dates of the OVERALL order count."""
    return sum(v.get("OVERALL", [0, 0])[0] for v in (cells or {}).values())


def delivery_totals(data):
    """{delivery date: (orders, waybills)} = the OVERALL counts summed over every order-date record in the json. Because each record only
    counts the parents / waybills the earlier records did not, these are the distinct Parent Orders / Waybills of that delivery date."""
    tot = {}
    for rec in (data.get("days") or {}).values():
        for dte, cell in (rec.get("cells") or {}).items():
            o, w = cell.get("OVERALL", [0, 0])[:2]
            a = tot.get(dte, (0, 0))
            tot[dte] = (a[0] + o, a[1] + w)
    return tot


def mfnas_present(path):
    return os.path.exists(path) and os.path.getsize(path) > 1024


# ----------------------------------------------------------------------------- MFNAS (missing days only)
def fetch_missing_mfnas(missing):
    """missing = [order dates]. One login for all. Returns {order date: True/False}."""
    res = {}
    nas = ro.NASClient()
    nas.login()
    try:
        for d in missing:
            t0s = (d + timedelta(days=1)).strftime("%Y%m%d")
            final = ro.report_path(d)
            try:
                folder = ro.find_date_folder(nas, t0s)
                files = [f for f in (nas.list_folder(folder, quiet=True) if folder else None) or []
                         if not f["isdir"] and f["name"].lower().endswith(".csv")]
                if not files:
                    log(f"MFNAS: {d:%Y-%m-%d}: no csv in folder {t0s} - cannot backfill this day")
                    res[d] = False
                    continue
                files.sort(key=lambda f: f["name"])
                pick = files[-1]
                if len(files) > 1:
                    log(f"MFNAS: WARNING {len(files)} csv files in {t0s}, using the newest: {pick['name']}")
                log(f"MFNAS: {d:%Y-%m-%d}: {pick['name']} -> {os.path.basename(final)}")
                tmp = final + ".part"
                size = nas.download(pick["path"], tmp)
                if size < 1024:
                    raise RuntimeError(f"downloaded file is only {size} bytes")
                os.replace(tmp, final)
                res[d] = True
            except Exception as ex:
                log(f"MFNAS: {d:%Y-%m-%d}: download failed: {type(ex).__name__}: {ex}")
                res[d] = False
    finally:
        nas.logout()
    return res


# ----------------------------------------------------------------------------- Tableau (one file for the whole range)
def default_tableau_path(first):
    return os.path.join(ro.ROLLING_DIR, f"AI Fetching for Order Date T-1 ({ro.mmm_dd(first)} backfill).csv")


_MON = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


def start_from_name(path, year):
    """Start date written in a Tableau csv name, e.g. "... T-1 (Oct-17).csv" / "(Oct-01 backfill)" -> 17 Oct; None if there is none."""
    norm = re.sub(r"[^a-z0-9]+", "_", os.path.basename(path).lower())
    for m in re.finditer(r"(?<![a-z])(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)_(\d{1,2})(?![0-9])", norm):
        try:
            return datetime(year, _MON[m.group(1)], int(m.group(2)))
        except ValueError:
            continue
    return None


def discover_tableau_files():
    """Every "AI Fetching for Order Date ... T-1" csv in the folder (daily files and backfill files alike)."""
    out = []
    for p in glob.glob(os.path.join(ro.ROLLING_DIR, "*.csv")):
        norm = re.sub(r"[^a-z0-9]+", "_", os.path.basename(p).lower())
        if norm.startswith("ai_fetching_for_order_date"):
            out.append(p)
    return sorted(out)


class TableauFiles:
    """A Tableau report with start date S holds the orders of order dates >= S (delivery date >= S) up to its export time, so a
    month is usually covered by several files (Oct-1 up to Oct 10, Oct-17 up to Oct 22 ...). For every order date the newest file
    that starts on/before it and reaches it is used; days no file covers are reported, never guessed."""

    def __init__(self, year):
        self.year, self.items, self._seen = year, [], set()

    def add(self, path, first, explicit=False):
        key = os.path.realpath(path)
        if key in self._seen:
            return
        self._seen.add(key)
        pm = ro.load_parent_map(path)
        start = start_from_name(path, self.year) or first          # a file without a date in its name is taken to start at the range start
        self.items.append({"path": path, "start": start, "covered": pm.max_order_date, "pm": pm, "mtime": os.path.getmtime(path)})
        log(f"Tableau: {os.path.basename(path)} - start {start:%Y-%m-%d}, order dates up to {pm.max_order_date:%Y-%m-%d}")

    def pick(self, d):
        ok = [t for t in self.items if t["start"] <= d <= t["covered"]]
        return max(ok, key=lambda t: t["mtime"]) if ok else None

    def reach(self):
        return ", ".join(f"{t['path'] and os.path.basename(t['path'])}: {t['start']:%m-%d}..{t['covered']:%m-%d}" for t in self.items)


def prepare_tableau(first, last, explicit, allow_download, force):
    """Returns a TableauFiles. Existing files are used as they are (the last order date is not forced); a file is downloaded
    (start date = first) only when there is none, or with --redownload-tableau."""
    tf = TableauFiles(first.year)
    paths = [explicit] if explicit else discover_tableau_files()
    if not force:
        for p in paths:
            if os.path.exists(p):
                try:
                    tf.add(p, first)
                except Exception as ex:
                    log(f"WARNING: {os.path.basename(p)} cannot be read as a Tableau report ({ex}) - ignored")
    if tf.items and not force:
        return tf
    if not allow_download:
        raise SystemExit(f"No Tableau csv found in {ro.ROLLING_DIR} and the Tableau download is disabled (--skip-download / "
                         f"--skip-tableau-download). Pass --tableau-file with a report whose start date is {first:%d/%m/%Y}.")
    target = explicit or default_tableau_path(first)
    new = target + ".new"
    log(f"Tableau: downloading with start date {first:%d/%m/%Y} ...")
    ro.download_tableau_report(first, new)                       # same steps as the daily job: type date, wait, Crosstab -> CSV
    os.replace(new, target)
    tf.add(target, first)
    log(f"Tableau: saved {target}")
    return tf


# ----------------------------------------------------------------------------- json
def write_days(results, data):
    """Replace only the processed days; snapshots / lastRun stay as they are."""
    data.setdefault("days", {})
    now = datetime.now().strftime("%Y-%m-%dT%H:%M")
    for d, (cells, _st) in results.items():
        data["days"][d.strftime("%Y-%m-%d")] = {"slot": ro.SLOTS[0], "asOf": now, "cells": cells}
    keep = sorted(data["days"])[-ro.KEEP_ORDER_DATES:]
    data["days"] = {k: data["days"][k] for k in keep}
    data["districts"] = ro.DISTRICTS
    data["generatedAt"] = now
    if os.path.exists(ro.ROLLING_JSON):
        shutil.copy2(ro.ROLLING_JSON, ro.ROLLING_JSON + ".bak_backfill")
    tmp = ro.ROLLING_JSON + ".tmp"
    os.makedirs(os.path.dirname(os.path.abspath(ro.ROLLING_JSON)), exist_ok=True)
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, ro.ROLLING_JSON)
    log(f"wrote {ro.ROLLING_JSON} ({len(results)} days replaced)")
    return data


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Backfill the Rolling Orders json for a date range (default: this month so far).")
    ap.add_argument("--from", dest="first", help="first ORDER date YYYY-MM-DD (default: 1st of the month of yesterday)")
    ap.add_argument("--to", dest="last", help="last ORDER date YYYY-MM-DD (default: yesterday = T-1)")
    ap.add_argument("--list", action="store_true", help="only show what is missing, then stop")
    ap.add_argument("--only-missing", action="store_true", help="only process order dates that are not in rolling_orders.json yet")
    ap.add_argument("--redownload", action="store_true", help="download every MFNAS csv again, even when it is in the folder")
    ap.add_argument("--redownload-tableau", action="store_true", help="download the Tableau csv again, even when the file looks fresh")
    ap.add_argument("--skip-download", action="store_true", help="never download (MFNAS or Tableau); use what is in the folder")
    ap.add_argument("--skip-tableau-download", action="store_true",
                    help="do not download the Tableau csv (missing MFNAS csv files are still downloaded); the Tableau file in the folder must reach the last order date")
    ap.add_argument("--tableau-file", help="use this Tableau csv (start date = first day of the range)")
    ap.add_argument("--headed", action="store_true", help="show the Chrome window during the Tableau download")
    ap.add_argument("--keep-files", action="store_true", help="do not delete the csv files after the days were written to the json")
    ap.add_argument("--skip-embed", action="store_true")
    ap.add_argument("--skip-push", action="store_true")
    a = ap.parse_args()
    if a.headed:
        ro.TABLEAU_HEADLESS = False

    today = datetime(*datetime.now().timetuple()[:3])
    last = _day(a.last) if a.last else today - timedelta(days=1)
    first = _day(a.first) if a.first else last.replace(day=1)
    if first > last:
        raise SystemExit(f"--from {first:%Y-%m-%d} is after --to {last:%Y-%m-%d}")
    dates = list(date_range(first, last))
    data = load_json()
    have = set(data.get("days", {}))
    log(f"range: order dates {first:%Y-%m-%d} .. {last:%Y-%m-%d} ({len(dates)} days)")

    # --- 1. what is missing?
    print(f"\n{'order date':<12}{'MFNAS csv':<12}{'in json':<9}")
    for d in dates:
        print(f"{d:%Y-%m-%d}  {'yes' if mfnas_present(ro.report_path(d)) else 'MISSING':<10}  "
              f"{'yes' if d.strftime('%Y-%m-%d') in have else 'MISSING':<9}")
    tfiles = [a.tableau_file] if a.tableau_file else discover_tableau_files()
    print("Tableau csv : " + ("; ".join(os.path.basename(p) for p in tfiles) if tfiles else "MISSING (will be downloaded)") + "\n")
    if a.list:
        return

    todo = [d for d in dates if not (a.only_missing and d.strftime("%Y-%m-%d") in have)]
    if not todo:
        log("nothing to do (every date is already in rolling_orders.json; drop --only-missing to recompute)")
        return

    with ro.run_lock():
        # --- 2. Tableau csv (one file for the whole range). The last order date is NOT forced: days after the latest order date
        failed = {}
        tf = prepare_tableau(first, last, a.tableau_file, allow_download=not (a.skip_download or a.skip_tableau_download),
                             force=a.redownload_tableau)
        for d in todo:
            if tf.pick(d) is None:
                failed[d] = f"no Tableau csv covers this order date ({tf.reach()}) - use --redownload-tableau --from {d:%Y-%m-%d}"
        todo = [d for d in todo if d not in failed]
        if not todo:
            raise SystemExit("no day is covered by a Tableau csv:\n  " + "\n  ".join(f"{d:%Y-%m-%d}: {m}" for d, m in sorted(failed.items())))

        # --- 3. MFNAS csv files
        need = [d for d in todo if a.redownload or not mfnas_present(ro.report_path(d))]
        if need and not a.skip_download:
            log(f"MFNAS: {len(need)} csv file(s) to download: {', '.join(f'{d:%m-%d}' for d in need)}")
            for d, ok in fetch_missing_mfnas(need).items():
                if not ok:
                    failed[d] = "MFNAS csv could not be downloaded"
        for d in todo:
            if d not in failed and not mfnas_present(ro.report_path(d)):
                failed[d] = "MFNAS csv missing (downloads skipped)" if a.skip_download else "MFNAS csv missing"
        todo = [d for d in todo if d not in failed]
        if not todo:
            raise SystemExit("no day has an MFNAS csv - nothing to process: " + "; ".join(f"{d:%m-%d}: {m}" for d, m in failed.items()))

        # --- 4. process (chronological: every day is deduplicated against the days before it)
        results = {}
        registry = ro.ParentRegistry().load()
        registry.quiet = True                        # earlier / later records are re-computed here; one warning at the end
        todo = sorted(todo)
        for d in todo:
            log(f"=== order date {d:%Y-%m-%d} ===")
            try:
                t = tf.pick(d)
                log(f"Tableau csv used: {os.path.basename(t['path'])}")
                rows, st = ro.process_report(ro.report_path(d), d, t["path"], parent_map=t["pm"], registry=registry)
                cells = ro.count_cells(rows)
                if not cells:
                    raise ValueError("0 usable rows")
                results[d] = (cells, st)
            except Exception as ex:
                failed[d] = f"{type(ex).__name__}: {ex}"
                log(f"order date {d:%Y-%m-%d}: FAILED - {ex}")
        old_days = json.loads(json.dumps(data.get("days", {})))      # snapshot of the old values for the summary
        if results:
            data = write_days(results, data)
            registry.save(max(results))
            later = registry.later_dates(max(results))
            if later:
                log(f"WARNING: the registry also holds records of later order dates ({later[0]} .. {later[-1]}); they were deduplicated against "
                    f"the previous history. Re-run the backfill up to yesterday (--to) so every day is consistent.")
            if not (a.keep_files or ro.KEEP_FILES):
                # v11.9 - delete what is now safely in the json: the MFNAS csv of every written day, and a Tableau csv once no day
                # it could serve is still waiting (failed days keep it, so a re-run does not have to download again).
                # A Tableau csv given with --tableau-file is never deleted.
                ro.cleanup_files([ro.report_path(d) for d in results], "written to the json")
                waiting = list(failed)
                for t in tf.items:
                    used = any(tf.pick(d) is t for d in results)
                    needed = any(t["start"] <= d <= t["covered"] for d in waiting)
                    if used and not needed and not a.tableau_file:
                        ro.cleanup_files([t["path"]], "all days it covers are in the json")
            if not a.skip_embed:
                ro.embed_in_index(data)
            if not (a.skip_push):
                ro.push_dashboard(max(results), ro.SLOTS[0], data)

    # --- 5. summary
    print(f"\n{'order date':<12}{'old orders':>12}{'new orders':>12}{'diff':>9}{'not in Tableau':>16}")
    for d in sorted(results):
        new = order_total(results[d][0])
        old = order_total(old_days.get(d.strftime('%Y-%m-%d'), {}).get("cells")) if old_days.get(d.strftime('%Y-%m-%d')) else None
        print(f"{d:%Y-%m-%d}  {('-' if old is None else format(old, ',')):>12}{new:>12,}"
              f"{('' if old is None else format(new - old, '+,')):>9}{results[d][1].get('ordersNotInTableau', 0):>16,}")
    if results:
        tot = delivery_totals(data)
        shown = [k for k in sorted(tot) if first.strftime("%Y-%m-%d") <= k <= last.strftime("%Y-%m-%d")]
        if shown:
            print(f"\n{'delivery date':<14}{'orders (distinct parents)':>28}{'waybills (distinct)':>22}   (all order dates in the json)")
            for k in shown:
                print(f"{k:<14}{tot[k][0]:>28,}{tot[k][1]:>22,}")
    if failed:
        print("\nNOT updated:")
        for d in sorted(failed):
            print(f"  {d:%Y-%m-%d}  {failed[d]}")
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
