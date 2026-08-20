#!/usr/bin/env bash
# Linux/Mac equivalent of run_productivity_0300.bat, for hosts using cron
# instead of Windows Task Scheduler. Same two steps, same files touched.
# Intended crontab entry (see setup steps in chat):
#   0 3 * * * /path/to/repo/run_productivity_0300.sh >> /path/to/repo/pipeline_log.txt 2>&1
set -euo pipefail

# Resolves to this script's own folder regardless of where cron's default
# working directory is — same purpose as the .bat's "cd /d %~dp0".
cd "$(dirname "${BASH_SOURCE[0]}")"

python3 kpi_pipeline.py --section productivity

git add public/data.json public/productivity_history.json history.json
git commit -m "Auto-update: productivity data $(date '+%Y-%m-%d %H:%M:%S')" || true
git push
