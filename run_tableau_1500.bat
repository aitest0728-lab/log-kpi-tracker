@echo off
REM Scheduled at 15:00 daily via Windows Task Scheduler.
REM Set "Start in" to this folder (your cloned repo) when creating the task.

cd /d %~dp0
py kpi_pipeline.py --section tableau >> pipeline_log.txt 2>&1

git add public/data.json history.json
git commit -m "Auto-update: Tableau data %date% %time%"
git push
