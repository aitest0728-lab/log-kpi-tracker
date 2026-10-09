#!/usr/bin/env bash
# KPI Dashboard - 08:30 daily job (WSL / Linux cron). Replaces BOTH run_delay_0830.sh and run_tableau_1400.*
#   1. kpi_pipeline.py --section tableau      all Tableau downloads (includes the T-1 delay rate)
#   2. git add / commit / push                deploy the refreshed dashboard
#   3. kpi_pipeline.py --section delayreport  Delay % screenshot + caption -> WhatsApp group
# Step 1 failing (stale / missing download) stops everything: wrong numbers are never deployed or posted.
# Step 2 failing is logged loudly but does not stop step 3.
# Crontab:
#   CRON_TZ=Asia/Hong_Kong
#   30 8 * * * /path/to/repo/run_tableau_0830.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -uo pipefail
export TZ=Asia/Hong_Kong PYTHONIOENCODING=utf-8 LANG=C.UTF-8
PYTHON="${PYTHON:-python3}"
cd "$(dirname "${BASH_SOURCE[0]}")" || exit 1
LOG=pipeline_log.txt

echo "" >> "$LOG"
echo "==== run_tableau_0830 started $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"

"$PYTHON" kpi_pipeline.py --section tableau >> "$LOG" 2>&1
if [ $? -ne 0 ]; then
    echo "  X tableau update FAILED - nothing deployed, nothing sent to WhatsApp. See the error above." >> "$LOG"
    exit 1
fi

echo "  Committing and pushing dashboard data..." >> "$LOG"
for f in public/data.json public/productivity_history.json public/manpower_distribution.json public/gmv_history.json public/delay_history.json public/other_aspects_history.json public/index.html history.json; do
    if [ -f "$f" ]; then git add "$f" >> "$LOG" 2>&1; else echo "  (skip) $f not found - not staging" >> "$LOG"; fi
done
git commit -m "Auto-update: Tableau data $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG" 2>&1 \
    || echo "  i git commit reported an error - if it only means 'nothing to commit' the push below still runs." >> "$LOG"
if git push >> "$LOG" 2>&1; then
    echo "  OK git push succeeded - dashboard data deployed." >> "$LOG"
else
    echo "  X git push FAILED - dashboard NOT updated online (data is correct locally). Continuing with WhatsApp." >> "$LOG"
fi

if command -v xvfb-run >/dev/null 2>&1; then
    WA_HIDE_MODE=off xvfb-run -a "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
else
    "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
fi
if [ $? -ne 0 ]; then
    echo "  X WhatsApp post FAILED - re-run by hand:  $PYTHON kpi_pipeline.py --section delayreport   (add --force to repost)" >> "$LOG"
    exit 2
fi
echo "==== run_tableau_0830 finished $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"
