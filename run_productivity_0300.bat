@echo off
REM Scheduled at 03:00 daily via Windows Task Scheduler.
REM Set "Start in" to this folder (your cloned repo) when creating the task.

cd /d %~dp0
py kpi_pipeline.py --section productivity >> pipeline_log.txt 2>&1

git add public/data.json public/productivity_history.json history.json
git commit -m "Auto-update: productivity data %date% %time%"
git push
