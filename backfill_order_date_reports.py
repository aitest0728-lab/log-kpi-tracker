#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Backfill MFNAS Order Date reports for a range of order dates (T-1)."""
import argparse
import json
import os
import sys
import time
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

from download_order_date_report import (  # noqa: E402
    NAS_REPORT_FOLDER,
    NASClient,
    find_date_folder,
    log,
    report_path,
)

MIN_BYTES = 1024


def daterange(start, end):
    day = start
    while day <= end:
        yield day
        day += timedelta(days=1)


def download_one(nas, order_date, dest, force):
    t0 = order_date + timedelta(days=1)
    t0s = t0.strftime("%Y%m%d")
    if os.path.exists(dest) and (not force) and os.path.getsize(dest) >= MIN_BYTES:
        size = os.path.getsize(dest)
        log(f"SKIP {order_date:%Y-%m-%d} already have {os.path.basename(dest)} ({size / 1048576:.1f} MB)")
        return {"orderDate": order_date.strftime("%Y-%m-%d"), "t0": t0.strftime("%Y-%m-%d"),
                "status": "skipped", "local": dest, "bytes": size}

    folder = find_date_folder(nas, t0s)
    files = [
        f for f in (nas.list_folder(folder, quiet=True) if folder else None) or []
        if not f["isdir"] and f["name"].lower().endswith(".csv")
    ]
    if not files:
        log(f"MISSING {order_date:%Y-%m-%d}: no csv in {folder or NAS_REPORT_FOLDER + '/' + t0s}")
        return {"orderDate": order_date.strftime("%Y-%m-%d"), "t0": t0.strftime("%Y-%m-%d"),
                "status": "missing", "folder": folder}

    files.sort(key=lambda f: f["name"])
    pick = files[-1]
    if len(files) > 1:
        log(f"WARNING {order_date:%Y-%m-%d}: {len(files)} csv files, using {pick['name']}")
    log(f"{order_date:%Y-%m-%d}: {pick['name']} -> {os.path.basename(dest)}")
    os.makedirs(os.path.dirname(dest) or ".", exist_ok=True)
    tmp = dest + ".part"
    size = nas.download(pick["path"], tmp)
    if size < MIN_BYTES:
        raise RuntimeError(f"downloaded file is only {size} bytes")
    os.replace(tmp, dest)
    return {"orderDate": order_date.strftime("%Y-%m-%d"), "t0": t0.strftime("%Y-%m-%d"),
            "status": "downloaded", "remote": pick["name"], "local": dest, "bytes": size}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--order-from", required=True, help="First order date (T-1) YYYY-MM-DD")
    ap.add_argument("--order-to", required=True, help="Last order date (T-1) YYYY-MM-DD")
    ap.add_argument("--force", action="store_true", help="Re-download even if the local file exists")
    ap.add_argument("--manifest", help="Where to write the JSON manifest")
    args = ap.parse_args()

    start = datetime.strptime(args.order_from, "%Y-%m-%d")
    end = datetime.strptime(args.order_to, "%Y-%m-%d")
    if end < start:
        raise SystemExit("order-to is before order-from")

    dates = list(daterange(start, end))
    log(f"backfill {len(dates)} order dates: {start:%Y-%m-%d} .. {end:%Y-%m-%d}")
    nas = NASClient()
    nas.login()
    results = []
    try:
        for order_date in dates:
            dest = report_path(order_date)
            try:
                results.append(download_one(nas, order_date, dest, args.force))
            except Exception as ex:
                log(f"ERROR {order_date:%Y-%m-%d}: {ex}")
                results.append({
                    "orderDate": order_date.strftime("%Y-%m-%d"),
                    "t0": (order_date + timedelta(days=1)).strftime("%Y-%m-%d"),
                    "status": "error",
                    "error": str(ex),
                })
    finally:
        nas.logout()

    summary = {
        "from": args.order_from,
        "to": args.order_to,
        "counts": {
            "downloaded": sum(1 for r in results if r["status"] == "downloaded"),
            "skipped": sum(1 for r in results if r["status"] == "skipped"),
            "missing": sum(1 for r in results if r["status"] == "missing"),
            "error": sum(1 for r in results if r["status"] == "error"),
            "total": len(results),
        },
        "results": results,
        "finishedAt": datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    }
    manifest = args.manifest or os.path.join(os.environ.get("MFNAS_REPORT_DIR", HERE),
                                             f"backfill_{args.order_from}_to_{args.order_to}.json")
    os.makedirs(os.path.dirname(manifest) or ".", exist_ok=True)
    with open(manifest, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=2)
    log(f"manifest {manifest}")
    log("summary " + json.dumps(summary["counts"]))
    if summary["counts"]["error"] or summary["counts"]["missing"]:
        sys.exit(2)


if __name__ == "__main__":
    try:
        main()
    except Exception as ex:
        log(f"ERROR: {ex}")
        sys.exit(1)
