@echo off
REM Scheduled at 03:00 daily via Windows Task Scheduler.
REM Set "Start in" to this folder (your cloned repo) when creating the task.

REM Force UTF-8 so Python printing emoji/Chinese text doesn't crash with
REM UnicodeEncodeError when Task Scheduler redirects output to a file
REM instead of a real console (which defaults to the system's ANSI codepage).
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d %~dp0

echo. >> pipeline_log.txt
echo ==== run_productivity_0300 started %date% %time% ==== >> pipeline_log.txt

py kpi_pipeline.py --section productivity >> pipeline_log.txt 2>&1

REM Fix: previously only the `py` line above was redirected into
REM pipeline_log.txt, and git ran with no logging or error checking at all
REM (same silent-failure risk described in run_tableau_1400.bat). Also
REM fixes the same atomic-git-add issue: staging files one at a time so a
REM missing file (e.g. very first run, before productivity_history.json
REM exists) skips loudly instead of blocking every other file.
echo   Committing and pushing dashboard data... >> pipeline_log.txt
for %%f in (public\data.json public\productivity_history.json public\manpower_distribution.json history.json) do (
  if exist "%%f" (
    git add "%%f" >> pipeline_log.txt 2>&1
  ) else (
    echo   ^(skip^) %%f not found on disk - not staging >> pipeline_log.txt
  )
)

git commit -m "Auto-update: productivity data %date% %time%" >> pipeline_log.txt 2>&1
if errorlevel 1 (
  echo   i git commit reported an error ^(errorlevel %errorlevel%^) - if this just means "nothing to commit", the next line still pushes any earlier unpushed commit. If it's a real error, check the git output just above this line. >> pipeline_log.txt
)
git push >> pipeline_log.txt 2>&1
if errorlevel 1 (
  echo   X git push FAILED - kpi_pipeline.py's data was generated locally but was NOT deployed to the dashboard. Check the git output just above this line, and check credentials/auth for the account Task Scheduler runs this task as. >> pipeline_log.txt
) else (
  echo   OK git push succeeded - dashboard data deployed. >> pipeline_log.txt
)
