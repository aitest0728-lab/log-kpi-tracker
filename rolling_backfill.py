#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rolling_backfill.py  -  KPI Dashboard "Rolling Orders" backfill  (companion of rolling_orders.py v11.1)

Downloads the ECOM-MFNAS Accumulated Delivery Lead Time report of EVERY day from 2026-10-01 up to yesterday's orders and
fills rolling_orders.json with them, so the Rolling Orders tab has the history its benchmark needs
(last-7-days benchmark now, rolling 30-day benchmark once a month of data exists).

How a day is mapped (identical to the 06:00 job):  order date D  <-  MFNAS folder  <report folder>/<D+1 as YYYYMMDD>/
Everything else - download, csv reading (trailing-comma fix), filters, counting, json update, embedding in index.html and
the git push - is reused from rolling_orders.py, so a backfilled day is byte-for-byte what the daily run would have written.

Usage (WSL, same folder as rolling_orders.py; same credentials as the daily job: CRAB_NAS_USER / CRAB_NAS_PASS or rolling_orders.env):
    python3 rolling_backfill.py                       # 2026-10-01 -> yesterday, skips days that are already complete in the json
    python3 rolling_backfill.py --dry-run             # only list what would be done
    python3 rolling_backfill.py --from 2026-10-01 --to 2026-10-13
    python3 rolling_backfill.py --force               # redo days that are already in the json
    python3 rolling_backfill.py --skip-push           # update rolling_orders.json + index.html locally, no git push
    python3 rolling_backfill.py --keep-csv            # keep the downloaded csv files (default: deleted after processing, ~80 MB each)
    python3 rolling_backfill.py --redownload          # ignore a csv that is already in "Folder for Rolling Orders"

Notes
  * A day whose MFNAS folder / csv does not exist (yet) is skipped and listed at the end; the run carries on with the other days.
  * A csv that already sits in "Folder for Rolling Orders" (e.g. Rolling Order Count (Oct-13).csv) is used as it is and never deleted.
  * Only "complete" days (slot 06:00, i.e. written by the MFNAS job) count as done. Old OIX-era partial snapshots (13:00 / 18:00 / 00:00)
    are NOT replaced by a backfill unless their order date is inside the range - they are replaced by the daily run of the next day.
  * Runs queue behind the 06:00 cron job (same lock), so the two never write rolling_orders.json at the same time.
"""
import argparse
import gc
import json
import os
import sys
import time
import traceback
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import rolling_orders as ro          # noqa: E402  (config, NASClient, report processing, json / embed / push)

log = ro.log
DEFAULT_FROM = "2026-10-01"
DONE_SLOT = ro.SLOTS[0]              # "06:00" = a day written by the MFNAS job


def _d(s):
    return datetime.strptime(s, "%Y-%m-%d")


def done_dates():
    """Order dates already complete in rolling_orders.json (slot 06:00)."""
    if not os.path.exists(ro.ROLLING_JSON):
        return set()
    try:
        days = json.load(open(ro.ROLLING_JSON, encoding="utf-8")).get("days", {})
    except Exception:
        return set()
    return {k for k, v in days.items() if (v or {}).get("slot") == DONE_SLOT and (v or {}).get("cells")}


def fetch_one(nas, order_date, final_path):
    """Download the csv of the folder dated order_date+1 to final_path. True = downloaded, False = no folder / no csv there."""
    t0s = (order_date + timedelta(days=1)).strftime("%Y%m%d")
    folder = ro.find_date_folder(nas, t0s)
    if not folder:
        return False
    files = [f for f in (nas.list_folder(folder, quiet=True) or []) if not f["isdir"] and f["name"].lower().endswith(".csv")]
    if not files:
        return False
    files.sort(key=lambda f: f["name"])
    pick = files[-1]                                  # one csv per folder; if several, the newest timestamp wins
    if len(files) > 1:
        log(f"MFNAS: WARNING {len(files)} csv files in {t0s}, using the newest: {pick['name']}")
    log(f"MFNAS: {pick['name']}  ->  {os.path.basename(final_path)}")
    tmp = final_path + ".part"
    size = nas.download(pick["path"], tmp)
    if size < 1024:
        raise RuntimeError(f"downloaded file is only {size} bytes")
    os.replace(tmp, final_path)
    return True


def main():
    ap = argparse.ArgumentParser(description="Backfill rolling_orders.json from the MFNAS reports")
    ap.add_argument("--from", dest="d_from", default=DEFAULT_FROM, help=f"first ORDER date YYYY-MM-DD (default {DEFAULT_FROM})")
    ap.add_argument("--to", dest="d_to", help="last ORDER date YYYY-MM-DD (default yesterday = the order date of today's 06:00 run)")
    ap.add_argument("--force", action="store_true", help="also redo order dates that are already complete in rolling_orders.json")
    ap.add_argument("--redownload", action="store_true", help="download again even if the csv is already in the Rolling Orders folder")
    ap.add_argument("--keep-csv", action="store_true", help="keep the csv files this run downloaded")
    ap.add_argument("--dry-run", action="store_true", help="list the plan, download nothing")
    ap.add_argument("--skip-embed", action="store_true")
    ap.add_argument("--skip-push", action="store_true", help="do not git-push (same as ROLLING_DEPLOY=none)")
    a = ap.parse_args()

    first = _d(a.d_from)
    last = _d(a.d_to) if a.d_to else ro.run_dates()[1]            # run_dates() -> (T-0, T-1)
    if last < first:
        raise SystemExit(f"--to {last:%Y-%m-%d} is before --from {first:%Y-%m-%d}")
    wanted = [first + timedelta(days=i) for i in range((last - first).days + 1)]
    have = done_dates()
    todo = [d for d in wanted if a.force or f"{d:%Y-%m-%d}" not in have]
    log(f"backfill: order dates {first:%Y-%m-%d} -> {last:%Y-%m-%d} ({len(wanted)} days), "
        f"{len(wanted) - len(todo)} already complete, {len(todo)} to do" + (" [--force]" if a.force else ""))
    if a.dry_run:
        for d in todo:
            exists = os.path.exists(ro.report_path(d))
            log(f"  would do order date {d:%Y-%m-%d}  <-  folder {(d + timedelta(days=1)):%Y%m%d}"
                + ("  (csv already local)" if exists and not a.redownload else ""))
        return
    if not todo:
        log("nothing to do")
        return

    ok, missing, failed = [], [], []
    data = None
    with ro.run_lock():
        nas = None
        try:
            for d in todo:
                od = f"{d:%Y-%m-%d}"
                path = ro.report_path(d)
                downloaded = False
                try:
                    if a.redownload or not os.path.exists(path):
                        if nas is None:
                            nas = ro.NASClient()
                            nas.login()
                        try:
                            got = fetch_one(nas, d, path)
                        except Exception as ex:                    # session may have expired on a long run: log in again, retry once
                            log(f"MFNAS: {od} download problem ({ex}) - logging in again and retrying once")
                            nas.logout()
                            nas.login()
                            got = fetch_one(nas, d, path)
                        if not got:
                            log(f"{od}: no MFNAS csv for folder {(d + timedelta(days=1)):%Y%m%d} - skipped")
                            missing.append(od)
                            continue
                        downloaded = True
                    else:
                        log(f"{od}: using the csv already in the folder: {os.path.basename(path)}")
                    rows, st = ro.process_report(path, d)
                    cells = ro.count_cells(rows)
                    del rows
                    if not cells:
                        log(f"{od}: report read but 0 usable rows - nothing written")
                        failed.append((od, "0 usable rows"))
                        continue
                    data = ro.update_json(d, DONE_SLOT, cells, {"report": st, "backfill": True})
                    ok.append(od)
                except Exception as ex:
                    log(f"{od}: FAILED - {ex}")
                    traceback.print_exc()
                    failed.append((od, str(ex)[:200]))
                finally:
                    if downloaded and not a.keep_csv and os.path.exists(path):
                        try:
                            os.remove(path)
                        except OSError:
                            pass
                    gc.collect()
        finally:
            if nas is not None:
                nas.logout()

    if data is not None:
        if not a.skip_embed:
            ro.embed_in_index(data)
        if not a.skip_push:
            ro.push_dashboard(_d(max(ok)), DONE_SLOT, data)

    log("=" * 60)
    log(f"backfill finished: {len(ok)} written, {len(missing)} without a report, {len(failed)} failed")
    if ok:
        log(f"  written : {ok[0]} -> {ok[-1]} ({len(ok)} days)")
    if missing:
        log(f"  no csv  : {', '.join(missing)}")
    for od, why in failed:
        log(f"  failed  : {od}  {why}")
    if missing or failed:
        sys.exit(1)


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        traceback.print_exc()
        sys.exit(1)
