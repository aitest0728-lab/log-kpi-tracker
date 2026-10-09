#!/usr/bin/env bash
# KPI Dashboard v9.1 - 08:30 daily Delay Rate job (WSL / Linux cron).
#   1. kpi_pipeline.py --section delay        T-1 delay report only -> dashboard files
#   2. git add / commit / push                deploy the refreshed dashboard
#   3. kpi_pipeline.py --section delayreport  screenshot of the Delay % tab + caption -> WhatsApp group
# Step 1 failing (stale / missing download) stops everything: wrong numbers are never deployed or posted.
# Step 2 failing is logged loudly but does not stop step 3 (the figures themselves are correct).
# Crontab (see README / chat):
#   CRON_TZ=Asia/Hong_Kong
#   30 8 * * * /path/to/repo/run_delay_0830.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -uo pipefail
export TZ=Asia/Hong_Kong PYTHONIOENCODING=utf-8 LANG=C.UTF-8
PYTHON="${PYTHON:-python3}"
cd "$(dirname "${BASH_SOURCE[0]}")" || exit 1
LOG=pipeline_log.txt

{
echo ""
echo "==== run_delay_0830 started $(date '+%Y-%m-%d %H:%M:%S') ===="
} >> "$LOG"

"$PYTHON" kpi_pipeline.py --section delay >> "$LOG" 2>&1
if [ $? -ne 0 ]; then
    echo "  X delay update FAILED - nothing deployed, nothing sent to WhatsApp. See the error above." >> "$LOG"
    exit 1
fi

# only the files this job touches; added one by one so a missing file is a loud skip, not a silent all-or-nothing failure
echo "  Committing and pushing dashboard data..." >> "$LOG"
for f in public/data.json public/delay_history.json public/index.html history.json; do
    if [ -f "$f" ]; then git add "$f" >> "$LOG" 2>&1; else echo "  (skip) $f not found - not staging" >> "$LOG"; fi
done
git commit -m "Auto-update: T-1 delay rate $(date '+%Y-%m-%d %H:%M:%S')" >> "$LOG" 2>&1 \
    || echo "  i git commit reported an error - if it only means 'nothing to commit' the push below still runs." >> "$LOG"
if git push >> "$LOG" 2>&1; then
    echo "  OK git push succeeded - dashboard data deployed." >> "$LOG"
else
    echo "  X git push FAILED - dashboard NOT updated online (data is correct locally). Continuing with WhatsApp." >> "$LOG"
fi

# cron has no display: give WhatsApp's Chromium a virtual one (same xvfb-run the other jobs use); WA_HIDE_MODE=off because
# xvfb-run already supplies it (wa_send's own Xvfb cannot start under WSL anyway)
if command -v xvfb-run >/dev/null 2>&1; then
    WA_HIDE_MODE=off xvfb-run -a "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
else
    "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
fi
if [ $? -ne 0 ]; then
    echo "  X WhatsApp post FAILED - re-run by hand:  $PYTHON kpi_pipeline.py --section delayreport   (add --force to repost)" >> "$LOG"
    exit 2
fi
echo "==== run_delay_0830 finished $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"
