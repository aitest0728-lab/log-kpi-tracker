#!/usr/bin/env bash
# Linux/Mac equivalent of run_productivity_0300.bat, for hosts using cron
# instead of Windows Task Scheduler. Same two steps, same files touched.
# Intended crontab entry (see setup steps in chat):
#   0 3 * * * /path/to/repo/run_productivity_0300.sh >> /path/to/repo/pipeline_log.txt 2>&1
# Fix: dropped -e (matching run_tableau_1400.sh) so a failed git
# commit/push below doesn't kill the script before an explicit
# success/failure line gets logged.
set -uo pipefail

# Resolves to this script's own folder regardless of where cron's default
# working directory is — same purpose as the .bat's "cd /d %~dp0".
cd "$(dirname "${BASH_SOURCE[0]}")"

echo "" >> pipeline_log.txt
echo "==== run_productivity_0300 started $(date '+%Y-%m-%d %H:%M:%S') ====" >> pipeline_log.txt

python3 kpi_pipeline.py --section productivity >> pipeline_log.txt 2>&1

# Fix: same atomic-git-add issue as run_tableau_1400.sh — stage files one
# at a time so a missing file (e.g. very first run, before
# productivity_history.json exists) skips loudly instead of silently
# blocking every other file from being staged.
echo "  Committing and pushing dashboard data..." >> pipeline_log.txt
for f in public/data.json public/productivity_history.json public/manpower_distribution.json history.json; do
    if [ -f "$f" ]; then
        git add "$f" >> pipeline_log.txt 2>&1
    else
        echo "  (skip) $f not found on disk — not staging" >> pipeline_log.txt
    fi
done

git commit -m "Auto-update: productivity data $(date '+%Y-%m-%d %H:%M:%S')" >> pipeline_log.txt 2>&1
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
