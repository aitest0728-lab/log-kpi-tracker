"""
Backfill OIX productivity history for past dates.

Your normal 03:00 job only ever processes "yesterday" (T-1). This script
walks a date range and processes each day the same way — it imports
kpi_pipeline.py and calls its actual find_oix_file / load_oix / process_oix /
productivity_for_group functions directly, so backfilled days are computed
with EXACTLY the same logic as a normal daily run. Nothing is duplicated or
reimplemented here.

What it touches:
  - history.json          (durable log — every backfilled date gets appended,
                            same as append_history()/append_daily_productivity_log()/
                            append_manpower_log() already do for a normal run)
  - public/productivity_history.json  (recomputed fresh at the end, same as
                            a normal run — this is always a derived/trimmed
                            window, never hand-edited)
  - public/manpower_distribution.json (v3.0 §4 — same derived/trimmed-window
                            treatment, using whatever staff list is
                            currently on disk for every backfilled date; see
                            position_map note below)
  - public/data.json's hktvStaff/odsRatio matrices — ONLY refreshed if you
    pass --update-current-forecast (see below). Backfilling the past doesn't
    change what "today's" actual/forecast numbers are by itself; the 7-day
    rolling forecast just becomes more accurate once history.json has more
    real days behind it, which naturally takes effect starting with
    tomorrow's normal 03:00 run. Pass the flag if you want that improved
    forecast to show up on the live dashboard immediately, using today's
    date as the "as of".

Usage:
    # Backfill everything you have files for for, from Aug 6 through
    # yesterday (skips any date already in history.json, unless --force):
    python backfill_oix_history.py --start 2026-08-06

    # Explicit end date instead of defaulting to yesterday:
    python backfill_oix_history.py --start 2026-08-06 --end 2026-08-18

    # Re-process dates even if already in history.json (e.g. you fixed a
    # truck-number pattern and want to recompute):
    python backfill_oix_history.py --start 2026-08-06 --force

    # Also refresh data.json's current hktvStaff/odsRatio forecast using the
    # now-fuller history, instead of waiting for tomorrow's 03:00 run:
    python backfill_oix_history.py --start 2026-08-06 --update-current-forecast
"""

import argparse
import datetime as dt

import kpi_pipeline as k


def backfill(start_date, end_date, force=False):
    history = k.load_history()
    log = history.setdefault("dailyProductivityLog", {})

    # v3.0 §4: same position_map every backfilled day uses, loaded ONCE up
    # front (not the current-dated staff list re-fetched per day — there's
    # no historical staff list per past date, so this is the best available
    # approximation, same as a normal run only ever has "today's" staff
    # list to work with too). None is handled gracefully by process_oix()/
    # productivity_for_group() — see kpi_pipeline.py.
    try:
        position_map = k.load_staff_position_map()
    except FileNotFoundError as e:
        print(f"  ⚠️ {e} — backfilling without Position/Courier/Driver "
              f"classification (§4.1 leader-exclusion won't apply).")
        position_map = None

    processed, skipped, failed = [], [], []

    d = start_date
    while d <= end_date:
        date_str = d.isoformat()
        if not force and date_str in log:
            skipped.append(date_str)
            d += dt.timedelta(days=1)
            continue

        try:
            path = k.find_oix_file(d)
        except FileNotFoundError:
            print(f"  {date_str}: no OIX file found — skipping")
            d += dt.timedelta(days=1)
            continue

        try:
            df = k.load_oix(path)
            df = k.process_oix(df, position_map)

            hktv_staff = k.productivity_for_group(df, ("LF", "LP"), exclude_positions=k.LEADER_EXCLUDE_POSITIONS)
            ods_ratio = k.productivity_for_group(df, ("ODS", "VAN"))
            courier_group = k.manpower_distribution_for_group(df, "courier")
            driver_group = k.manpower_distribution_for_group(df, "driver")

            for key, group in (("hktvStaff", hktv_staff), ("odsRatio", ods_ratio)):
                k.append_history(history, key, date_str, group["overall"],
                                  {dist: group["districts"][dist]["productivity"] for dist in k.DISTRICTS})

            k.append_daily_productivity_log(history, date_str, hktv_staff, ods_ratio)
            k.append_manpower_log(history, date_str, courier_group, driver_group)
            processed.append(date_str)
            print(f"  {date_str}: OK — hktvStaff overall={hktv_staff['overall']}, "
                  f"odsRatio overall={ods_ratio['overall']}")
        except Exception as e:
            failed.append((date_str, str(e)))
            print(f"  {date_str}: FAILED — {e}")

        d += dt.timedelta(days=1)

    k.save_history(history)
    k.save_productivity_history(k.trimmed_daily_productivity(history, k.DAILY_PRODUCTIVITY_KEEP_DAYS))
    k.save_manpower_history(k.trimmed_manpower_distribution(history, k.DAILY_PRODUCTIVITY_KEEP_DAYS))

    print(f"\nDone. Processed: {len(processed)}, skipped (already had data): {len(skipped)}, "
          f"failed: {len(failed)}")
    if failed:
        print("Failed dates:")
        for date_str, err in failed:
            print(f"  {date_str}: {err}")

    return history


def refresh_current_forecast(history):
    """Optional: recompute today's hktvStaff/odsRatio actual+forecast in
    data.json using the now-enriched history, instead of waiting for
    tomorrow's normal 03:00 run to pick it up."""
    payload = k.load_data_json()
    matrices = payload.setdefault("matrices", {})
    log = history.get("dailyProductivityLog", {})
    if not log:
        print("No history to refresh from — skipping --update-current-forecast.")
        return

    latest_date = sorted(log.keys())[-1]
    for key in ("hktvStaff", "odsRatio"):
        series = history.get(key, {})
        if latest_date not in series:
            continue
        latest_entry = series[latest_date]
        fc_overall, fc_districts = k.rolling_average(history, key, 7, dt.date.today())
        matrices[key] = {
            "actual": {"overall": latest_entry["overall"], "districts": latest_entry["districts"]},
            "forecast": {"overall": fc_overall, "districts": fc_districts},
            "asOf": latest_date,
        }
    k.save_data_json(payload)
    print(f"Refreshed data.json's hktvStaff/odsRatio forecast using history through {latest_date}.")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", required=True, help="YYYY-MM-DD, e.g. 2026-08-06")
    parser.add_argument("--end", help="YYYY-MM-DD — defaults to yesterday")
    parser.add_argument("--force", action="store_true",
                         help="Reprocess dates even if history.json already has them")
    parser.add_argument("--update-current-forecast", action="store_true",
                         help="Also refresh data.json's live hktvStaff/odsRatio forecast now, "
                              "instead of waiting for tomorrow's 03:00 run")
    args = parser.parse_args()

    start_date = dt.date.fromisoformat(args.start)
    end_date = dt.date.fromisoformat(args.end) if args.end else dt.date.today() - dt.timedelta(days=1)

    print(f"Backfilling {start_date} through {end_date}...\n")
    history = backfill(start_date, end_date, force=args.force)

    if args.update_current_forecast:
        refresh_current_forecast(history)


if __name__ == "__main__":
    main()
