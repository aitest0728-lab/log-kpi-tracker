#!/usr/bin/env bash
# KPI Dashboard - 08:15 daily job (WSL / Linux cron). Was run_tableau_0830.sh (started 08:30; moved 15 minutes earlier in v11.2).
#   1. kpi_pipeline.py --section tableau      all Tableau downloads (includes the T-1 delay rate)
#        - every report: 3 download attempts, then automatic RECOVERY sessions for any report that still failed (inside kpi_pipeline.py)
#        - if step 1 still fails: re-run it (up to TABLEAU_TRIES times in total, TABLEAU_RETRY_WAIT seconds apart) with --resume,
#          which only fetches the reports that are still missing instead of downloading everything again
#        - last resort: `--section delay` (just the two T-1 delay sheets) so the delay-rate figures and the WhatsApp post still go out;
#          the rest of the dashboard then keeps yesterday's figures and the run ends with exit code 3
#   2. git add / commit / push                deploy the refreshed dashboard
#   3. kpi_pipeline.py --section delayreport  Delay % screenshot + caption -> WhatsApp group
# Stale / missing data is still never deployed or posted: if even the last resort fails nothing is deployed or sent (exit 1).
# Step 2 failing is logged loudly but does not stop step 3.
# Crontab:
#   CRON_TZ=Asia/Hong_Kong
#   15 8 * * * /path/to/repo/run_tableau_0815.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -uo pipefail
export TZ=Asia/Hong_Kong PYTHONIOENCODING=utf-8 LANG=C.UTF-8
PYTHON="${PYTHON:-python3}"
TABLEAU_TRIES="${TABLEAU_TRIES:-3}"            # whole-step attempts (1st = full run, later ones = --resume)
TABLEAU_RETRY_WAIT="${TABLEAU_RETRY_WAIT:-90}" # seconds between those attempts
cd "$(dirname "${BASH_SOURCE[0]}")" || exit 1
LOG=pipeline_log.txt
RC=0

echo "" >> "$LOG"
echo "==== run_tableau_0815 started $(date '+%Y-%m-%d %H:%M:%S') ====" >> "$LOG"

step1_ok=0
for try in $(seq 1 "$TABLEAU_TRIES"); do
    if [ "$try" -eq 1 ]; then
        "$PYTHON" kpi_pipeline.py --section tableau >> "$LOG" 2>&1
    else
        echo "  >> tableau update attempt $try/$TABLEAU_TRIES (--resume: only the missing reports) $(date '+%H:%M:%S')" >> "$LOG"
        "$PYTHON" kpi_pipeline.py --section tableau --resume >> "$LOG" 2>&1
    fi
    if [ $? -eq 0 ]; then step1_ok=1; break; fi
    echo "  ! tableau update attempt $try/$TABLEAU_TRIES FAILED $(date '+%H:%M:%S')" >> "$LOG"
    if [ "$try" -lt "$TABLEAU_TRIES" ]; then sleep "$TABLEAU_RETRY_WAIT"; fi
done

if [ "$step1_ok" -ne 1 ]; then
    echo "  ! full tableau update failed $TABLEAU_TRIES times - last resort: T-1 delay rate only (--section delay)" >> "$LOG"
    if "$PYTHON" kpi_pipeline.py --section delay >> "$LOG" 2>&1; then
        echo "  ! delay-only update OK - the other dashboard figures were NOT refreshed today (exit code 3 at the end)." >> "$LOG"
        RC=3
    else
        echo "  X tableau update FAILED - nothing deployed, nothing sent to WhatsApp. See the error above." >> "$LOG"
        exit 1
    fi
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
    "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
else
    "$PYTHON" kpi_pipeline.py --section delayreport >> "$LOG" 2>&1
fi
if [ $? -ne 0 ]; then
    echo "  X WhatsApp post FAILED - re-run by hand:  $PYTHON kpi_pipeline.py --section delayreport   (add --force to repost)" >> "$LOG"
    exit 2
fi
echo "==== run_tableau_0815 finished $(date '+%Y-%m-%d %H:%M:%S') (exit $RC) ====" >> "$LOG"
exit "$RC"
