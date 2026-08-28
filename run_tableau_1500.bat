@echo off
REM Scheduled at 15:00 daily via Windows Task Scheduler.
REM Set "Start in" to this folder (your cloned repo) when creating the task.

REM Force UTF-8 so Python printing emoji/Chinese text doesn't crash with
REM UnicodeEncodeError when Task Scheduler redirects output to a file
REM instead of a real console (which defaults to the system's ANSI codepage).
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d %~dp0
py kpi_pipeline.py --section tableau >> pipeline_log.txt 2>&1

REM productivity_history.json isn't touched by this section (only the 03:00
REM productivity job writes it) - included here too just so this commit
REM doesn't miss it if the morning job's own commit ever failed silently.
git add public/data.json public/productivity_history.json public/manpower_distribution.json public/gmv_history.json history.json
git commit -m "Auto-update: Tableau data %date% %time%"
git push
