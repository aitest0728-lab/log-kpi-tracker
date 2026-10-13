#!/usr/bin/env bash
# KPI Dashboard - Daily Cost Report job (WSL / Linux cron), Tuesday + Friday.
#   1. kpi_pipeline.py --section costreport   download the newest "Daily Cost Report" from the WhatsApp group, update the dashboard files
#   2. git add / commit / push                deploy (same files and method as run_tableau_0815.sh)
# The 08:15 job has already pushed by now, so this push is what puts the cost report data online.
# Crontab (leave 09:00 - or move to 09:05 if the report is often posted a little late):
#   CRON_TZ=Asia/Hong_Kong
#   0 9 * * 2,5 /path/to/repo/run_costreport_0900.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -uo pipefail
export TZ=Asia/Hong_Kong PYTHONIOENCODING=utf-8 LANG=C.UTF-8
PYTHON="${PYTHON:-python3}"
cd "$(dirname "${BASH_SOURCE[0]}")" || exit 1
LOG=pipeline_log.txt

echo "" >> "$LOG"
echo "==== run_costreport_0900 started $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"

# DEPLOY_CMD=none: the push is done below, so the pipeline's own deploy hook must not push a second time.
if ! DEPLOY_CMD=none "$PYTHON" kpi_pipeline.py --section costreport >> "$LOG" 2>&1; then
    echo "  X cost report update FAILED - nothing pushed. Re-run by hand:  $PYTHON kpi_pipeline.py --section costreport" >> "$LOG"
    exit 1
fi

echo "  Committing and pushing dashboard data..." >> "$LOG"
for f in public/data.json public/productivity_history.json public/manpower_distribution.json public/gmv_history.json public/delay_history.json public/other_aspects_history.json public/index.html history.json; do
    if [ -f "$f" ]; then git add "$f" >> "$LOG" 2>&1; else echo "  (skip) $f not found - not staging" >> "$LOG"; fi
done
git commit -m "Auto-update: Daily Cost Report $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG" 2>&1 \
    || echo "  i git commit reported an error - if it only means 'nothing to commit' the push below still runs." >> "$LOG"
if git push >> "$LOG" 2>&1; then
    echo "  OK git push succeeded - cost report data deployed." >> "$LOG"
else
    echo "  X git push FAILED - dashboard NOT updated online (data is correct locally)." >> "$LOG"
    exit 2
fi
echo "==== run_costreport_0900 finished $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"
