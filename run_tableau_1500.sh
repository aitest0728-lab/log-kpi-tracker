#!/usr/bin/env bash
# Linux/Mac equivalent of run_tableau_1500.bat, for hosts using cron instead
# of Windows Task Scheduler.
# Intended crontab entry (see setup steps in chat):
#   0 15 * * * /path/to/repo/run_tableau_1500.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -euo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"

python3 kpi_pipeline.py --section tableau

# productivity_history.json isn't touched by this section (only the 03:00
# productivity job writes it) — included here too just so this commit
# doesn't miss it if the morning job's own commit ever failed silently.
git add public/data.json public/productivity_history.json public/manpower_distribution.json public/gmv_history.json history.json
git commit -m "Auto-update: Tableau data $(date '+%Y-%m-%d %H:%M:%S')" || true
git push
