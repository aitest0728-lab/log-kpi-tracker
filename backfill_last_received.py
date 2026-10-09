#!/usr/bin/env python3
"""
backfill_last_received.py — backfill the "Last Received by Customer" time for past days (default: 1 Oct .. T-1).

Put this file in the SAME folder as kpi_pipeline.py and run it from there:

    python backfill_last_received.py                          # every day from 1st of this month to T-1 that has no value yet
    python backfill_last_received.py --start 2026-10-01 --end 2026-10-07
    python backfill_last_received.py --force                  # recompute days that already have a value
    python backfill_last_received.py --dry-run                # compute + print only, write nothing

It reuses kpi_pipeline.compute_last_received() (same OIX_Record_YYYYMMDD logic as the 08:30 job), so results are identical to
`kpi_pipeline.py --section lastreceived --date D` run once per day — but history / delay_history.json / index.html are
written ONCE at the end instead of 7+ times. Days whose OIX_Record file is missing or empty are skipped and listed.
"""
import argparse
import datetime as dt
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import kpi_pipeline as kp  # noqa: E402


def parse_day(s):
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


def main():
    ap = argparse.ArgumentParser(description="Backfill Last Received by Customer time for several days.")
    ap.add_argument("--start", type=parse_day, help="first day (default: 1st of the current month)")
    ap.add_argument("--end", type=parse_day, help="last day (default: T-1)")
    ap.add_argument("--force", action="store_true", help="recompute days that already have a value")
    ap.add_argument("--dry-run", action="store_true", help="compute and print only; write nothing")
    args = ap.parse_args()

    t1 = kp.today_hkt() - dt.timedelta(days=1)
    end = args.end or t1
    start = args.start or end.replace(day=1)
    if start > end:
        raise SystemExit(f"--start {start} is after --end {end}")

    history = kp.load_history()
    existing = history.get("lastReceivedDaily", {})

    days = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    done, skipped_have, failed = [], [], []
    for day in days:
        key = day.isoformat()
        if key in existing and not args.force:
            skipped_have.append(key)
            continue
        print(f"\n== {key} ==")
        # apply_last_received is soft-fail: prints the reason and returns False on a missing file / no rows
        if kp.apply_last_received(history, day):
            done.append(key)
        else:
            failed.append(key)

    print("\n---------------- summary ----------------")
    print(f"filled   ({len(done)}): {', '.join(done) or '-'}")
    print(f"already had a value, left alone ({len(skipped_have)}): {', '.join(skipped_have) or '-'}  (use --force to redo)")
    print(f"skipped / no data ({len(failed)}): {', '.join(failed) or '-'}")

    if not done:
        print("Nothing new to write.")
        return
    if args.dry_run:
        print("--dry-run: nothing written.")
        return

    # Same persistence steps as kpi_pipeline.run_last_received_section(), done once for all days.
    old = {}
    if os.path.exists(kp.DELAY_HISTORY_PATH):
        try:
            with open(kp.DELAY_HISTORY_PATH, "r", encoding="utf-8") as f:
                old = json.load(f)
        except (OSError, ValueError):
            old = {}
    kp.save_history(history)
    kp.save_delay_history(kp.build_delay_monthly(
        history, None, None, None, mtd_override=old.get("mtd") or {"overall": {}, "districts": {}}))
    kp.update_embedded_data()
    print("✅ history, delay_history.json and the dashboard snapshot (index.html) updated. Deploy as usual.")


if __name__ == "__main__":
    main()
