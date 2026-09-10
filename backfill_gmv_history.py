"""
Backfill GMV / Basket Size history from a "Sheet 1.csv" GMV export.

Your normal 15:00 job (parse_gmv() in kpi_pipeline.py) only ever reads ONE
row out of "Sheet 1.csv" — target_date (T-1) — because that's all a fresh
daily download needs. But the file itself is a rolling crosstab that already
contains many days at once (in your Aug export: 2026-08-01 through
2026-09-12, including a few clearly-stale/garbage trailing rows). This
script walks EVERY dated row in that same file and appends each one to
history.json's "gmv" log — using the exact same column layout, district
normalization (normalize_district_code — merges "NT-YT" into "NT-TW"), and
money parsing (parse_money) that parse_gmv() uses for a normal run, so a
backfilled day is computed identically to a normal day. Nothing is
reimplemented here.

What it touches:
  - history.json           ("gmv" key — durable full daily log, same role
                            append_gmv_history() already plays for a normal
                            run; one entry per date, {"overall", "districts"})
  - public/gmv_history.json (recomputed fresh at the end via
                            build_gmv_monthly(), same as a normal run —
                            always a derived file, never hand-edited)

Basket Size for each backfilled date is computed the same way
build_gmv_monthly() always computes it — GMV ÷ that date's Total Parent
Order count, read from history.json's dailyProductivityLog. That means: a
date only gets a real Basket Size once the 03:00 productivity job (or
backfill_oix_history.py) has already produced an entry for it. Any GMV date
with no matching productivity entry still gets its GMV recorded, just with
basketSize: None for that day — same "None rather than a misleading 0 or
divide-by-zero" behavior parse_gmv's caller already relies on elsewhere.

This does NOT touch data.json's live "gmv" matrix (today's displayed
actual/basketSize/asOf) — backfilling the past doesn't change what "today"
shows, and data.json already gets that from the normal 15:00 run. If you
ever do want to force data.json's gmv matrix to refresh from a fuller
history without waiting for tomorrow's run, that's a separate, deliberate
step — not something a backfill should do implicitly.

Usage:
    # Backfill every dated row in the file that falls between --start and
    # --end (inclusive), skipping any date already in history.json's "gmv"
    # log unless --force:
    python backfill_gmv_history.py --start 2026-08-01 --end 2026-08-26

    # Default --end is yesterday, so you can omit it day-to-day:
    python backfill_gmv_history.py --start 2026-08-01

    # Point at a specific CSV instead of the normal REPORT_FOLDER/"Sheet 1.csv"
    # location (e.g. a one-off upload, as opposed to today's live download):
    python backfill_gmv_history.py --start 2026-08-01 --file /path/to/Sheet_1.csv

    # Re-process dates even if already in history.json (e.g. you fixed a
    # district mapping and want to recompute):
    python backfill_gmv_history.py --start 2026-08-01 --force
"""

import argparse
import datetime as dt
import os
import re

import pandas as pd

import kpi_pipeline as k

# Matches the row-0 date-pivot label format parse_gmv() searches for
# ("YYYY年M月D日", no zero-padding — confirmed against the sample export).
_DATE_LABEL_RE = re.compile(r"^(\d{4})年(\d{1,2})月(\d{1,2})日$")


def parse_date_label(label):
    """'2026年8月1日' -> date(2026, 8, 1), or None if it doesn't match /
    isn't a real calendar date (guards against stray/garbage rows like a
    stray far-future date sometimes left in a live export — those simply
    won't match any --start/--end range anyway, so no special-casing
    needed)."""
    m = _DATE_LABEL_RE.match(str(label).strip())
    if not m:
        return None
    y, mo, d = (int(g) for g in m.groups())
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def parse_gmv_all_dates(path):
    """Same layout/columns/parsing as kpi_pipeline.parse_gmv(), but returns
    every dated row in the file as {date: {"overall", "districts"}} instead
    of searching for one target date."""
    raw = pd.read_csv(path, encoding="utf-16", sep="\t", header=None, dtype=str)
    header = raw.iloc[1].tolist()
    data_rows = raw.iloc[2:].reset_index(drop=True)

    col_to_district = {}
    for idx, label in enumerate(header):
        if idx == 0:
            continue
        norm = k.normalize_district_code(str(label).strip())
        if norm in k.DISTRICTS:
            col_to_district[idx] = norm

    missing = [d for d in k.DISTRICTS if d not in col_to_district.values()]
    if missing:
        print(f"  ⚠️ GMV export is missing column(s) for district(s): {missing} "
              f"— those will be treated as $0 for every date.")

    out = {}
    for _, row in data_rows.iterrows():
        date_ = parse_date_label(row.iloc[0])
        if date_ is None:
            continue
        per_district = {d: 0.0 for d in k.DISTRICTS}
        for idx, dist in col_to_district.items():
            per_district[dist] += k.parse_money(row.iloc[idx])
        out[date_] = {"overall": round(sum(per_district.values()), 2),
                       "districts": {d: round(v, 2) for d, v in per_district.items()}}
    return out


def backfill(rows_by_date, start_date, end_date, force=False):
    history = k.load_history()
    gmv_log = history.setdefault("gmv", {})

    processed, skipped, no_row = [], [], []

    d = start_date
    while d <= end_date:
        date_str = d.isoformat()
        if not force and date_str in gmv_log:
            skipped.append(date_str)
            d += dt.timedelta(days=1)
            continue

        gmv_group = rows_by_date.get(d)
        if gmv_group is None:
            no_row.append(date_str)
            d += dt.timedelta(days=1)
            continue

        k.append_gmv_history(history, date_str, gmv_group)
        processed.append(date_str)
        has_orders = date_str in history.get("dailyProductivityLog", {})
        note = "" if has_orders else "  (no dailyProductivityLog entry yet — basketSize will be None for this date)"
        print(f"  {date_str}: OK — GMV overall={gmv_group['overall']}{note}")

        d += dt.timedelta(days=1)

    k.save_history(history)
    k.save_gmv_history(k.build_gmv_monthly(history))

    print(f"\nDone. Processed: {len(processed)}, skipped (already had data): {len(skipped)}, "
          f"no row in file: {len(no_row)}")
    if no_row:
        print("Dates with no matching row in the GMV file:")
        for date_str in no_row:
            print(f"  {date_str}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", required=True, help="YYYY-MM-DD, e.g. 2026-08-01")
    parser.add_argument("--end", help="YYYY-MM-DD — defaults to yesterday")
    parser.add_argument("--file",
                         help="Path to the GMV export CSV — defaults to the normal "
                              "REPORT_FOLDER/'Sheet 1.csv' location a live 15:00 run uses")
    parser.add_argument("--force", action="store_true",
                         help="Reprocess dates even if history.json's \"gmv\" log already has them")
    args = parser.parse_args()

    start_date = dt.date.fromisoformat(args.start)
    end_date = dt.date.fromisoformat(args.end) if args.end else dt.date.today() - dt.timedelta(days=1)
    path = args.file or os.path.join(k.REPORT_FOLDER, k.REPORT_FILES["gmv"])

    print(f"Reading GMV rows from {path!r}...")
    rows_by_date = parse_gmv_all_dates(path)
    print(f"Found {len(rows_by_date)} dated rows in the file.\n")

    print(f"Backfilling {start_date} through {end_date}...\n")
    backfill(rows_by_date, start_date, end_date, force=args.force)


if __name__ == "__main__":
    main()
