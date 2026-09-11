#!/usr/bin/env bash
# Linux/Mac equivalent of run_tableau_1400.bat, for hosts using cron instead
# of Windows Task Scheduler.
# Intended crontab entry (see setup steps in chat):
#   0 14 * * * /path/to/repo/run_tableau_1400.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -uo pipefail
# (dropped -e so a failed git commit/push below doesn't kill the script
# before the explicit success/failure line gets logged)

cd "$(dirname "${BASH_SOURCE[0]}")"

echo "" >> pipeline_log.txt
echo "==== run_tableau_1400 started $(date '+%Y-%m-%d %H:%M:%S') ====" >> pipeline_log.txt

python3 kpi_pipeline.py --section tableau >> pipeline_log.txt 2>&1

# Fix: this relied entirely on the crontab's own `>> pipeline_log.txt 2>&1`
# wrapping the whole script to capture git's output — fine if the cron entry
# is set up exactly as documented above, but silent (same failure mode as
# the .bat had) if this script is ever run manually or via a different
# scheduler that doesn't redirect the whole invocation. Logging explicitly
# here makes it self-contained either way, and checking exit status turns a
# failed push (data generated locally but never deployed) into a loud,
# visible log line instead of nothing.
#
# productivity_history.json isn't touched by this section (only the 03:00
# productivity job writes it) — included here too just so this commit
# doesn't miss it if the morning job's own commit ever failed silently.
echo "  Committing and pushing dashboard data..." >> pipeline_log.txt

# Fix: `git add` is atomic per invocation — if any one pathspec below
# doesn't exist on disk, NOTHING in the command gets staged, not even the
# files that were fine (this is what caused the 2026-08-28 15:00 run to
# silently skip deploying: public/manpower_distribution.json is only ever
# written by the 03:00 productivity job, so on a run where that job hasn't
# produced it yet, git add failed for ALL five files, commit had nothing to
# commit, and push then reported "Everything up-to-date" even though fresh
# data.json/gmv_history.json changes were sitting uncommitted locally).
# Adding files one at a time makes a missing file a loud skip line instead
# of silently blocking everything else.
# Fix: public/delay_history.json and public/other_aspects_history.json were
# missing from this list entirely - kpi_pipeline.py has been writing both
# correctly on every run ("Wrote ./public/delay_history.json" /
# "Wrote ./public/other_aspects_history.json" in pipeline_log.txt), but
# since they were never in this pathspec list, `git add` never staged them,
# so they sat as "Untracked files" forever and never reached the deployed
# dashboard - the Delay % / Other Aspects tabs had nothing to fetch even
# though the pipeline was generating their data correctly the whole time.
for f in public/data.json public/productivity_history.json public/manpower_distribution.json public/gmv_history.json public/delay_history.json public/other_aspects_history.json history.json; do
    if [ -f "$f" ]; then
        git add "$f" >> pipeline_log.txt 2>&1
    else
        echo "  (skip) $f not found on disk — not staging" >> pipeline_log.txt
    fi
done

git commit -m "Auto-update: Tableau data $(date '+%Y-%m-%d %H:%M:%S')" >> pipeline_log.txt 2>&1
commit_status=$?
if [ $commit_status -ne 0 ]; then
    echo "  i git commit reported an error (exit $commit_status) — if this just means \"nothing to commit\", the next step still pushes any earlier unpushed commit. If it's a real error, check the git output just above this line." >> pipeline_log.txt
fi

git push >> pipeline_log.txt 2>&1
push_status=$?
if [ $push_status -ne 0 ]; then
    echo "  X git push FAILED (exit $push_status) — kpi_pipeline.py's data was generated locally but was NOT deployed to the dashboard. Check the git output just above this line, and check credentials/auth for the account this cron job runs as." >> pipeline_log.txt
else
    echo "  OK git push succeeded — dashboard data deployed." >> pipeline_log.txt
fi
