@echo off
REM ============================================================
REM  KPI Dashboard - Temporary delivery points refresh (09:45 daily)
REM  1) reads "Tracking of Delivery Address Status" (written by address_tracking.py, 09:30)
REM  2) embeds the open changes into index.html  (kpi_pipeline.py --section maptemp)
REM  3) kpi_pipeline.py then runs DEPLOY_CMD to push the dashboard
REM  Put this file in the same folder as kpi_pipeline.py and index.html.
REM ============================================================
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
set PYTHONUTF8=1

REM --- Push command. If you already set DEPLOY_CMD as a system/user environment
REM     variable (used by your 09:00 cost-report job) it is kept as is.
REM     Otherwise this default does: git add -> commit (only if something changed) -> push.
if not defined DEPLOY_CMD set "DEPLOY_CMD=git add -A && (git diff --cached --quiet || git commit -m "temp delivery points auto update") && git push"

echo. >> pipeline_log.txt
echo ===== %date% %time% : maptemp 09:45 ===== >> pipeline_log.txt
py kpi_pipeline.py --section maptemp >> pipeline_log.txt 2>&1
set RC=%ERRORLEVEL%
echo ===== finished, exit code %RC% ===== >> pipeline_log.txt
exit /b %RC%
