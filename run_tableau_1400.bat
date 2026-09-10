@echo off
REM Scheduled at 14:00 daily via Windows Task Scheduler.
REM Set "Start in" to this folder (your cloned repo) when creating the task.

REM Force UTF-8 so Python printing emoji/Chinese text doesn't crash with
REM UnicodeEncodeError when Task Scheduler redirects output to a file
REM instead of a real console (which defaults to the system's ANSI codepage).
chcp 65001 >nul
set PYTHONIOENCODING=utf-8

cd /d %~dp0

echo. >> pipeline_log.txt
echo ==== run_tableau_1400 started %date% %time% ==== >> pipeline_log.txt

py kpi_pipeline.py --section tableau >> pipeline_log.txt 2>&1

REM Fix: previously only the `py` line above was redirected into
REM pipeline_log.txt - these three git commands ran with no logging at all.
REM Under Task Scheduler ("run whether user is logged on or not"), a failed
REM git push (expired credentials, Credential Manager unable to prompt in a
REM non-interactive session, network blip, etc.) would fail completely
REM silently: kpi_pipeline.py generates correct data locally, nothing ever
REM reaches the dashboard, and there's no record anywhere of why. Redirecting
REM these too, plus checking errorlevel after push, makes that failure mode
REM visible instead of invisible.
REM
REM productivity_history.json isn't touched by this section (only the 03:00
REM productivity job writes it) - included here too just so this commit
REM doesn't miss it if the morning job's own commit ever failed silently.
echo   Committing and pushing dashboard data... >> pipeline_log.txt

REM Fix: `git add` is atomic per invocation - if any one pathspec below
REM doesn't exist on disk, NOTHING in the command gets staged, not even the
REM files that were fine (this is what caused the 2026-08-28 15:00 run to
REM silently skip deploying: public/manpower_distribution.json is only ever
REM written by the 03:00 productivity job, so on a run where that job
REM hasn't produced it yet, git add failed for ALL five files, commit had
REM nothing to commit, and push then reported "Everything up-to-date" even
REM though fresh data.json/gmv_history.json changes were sitting
REM uncommitted locally). Adding files one at a time makes a missing file
REM a loud skip line instead of silently blocking everything else.
for %%f in (public\data.json public\productivity_history.json public\manpower_distribution.json public\gmv_history.json history.json) do (
  if exist "%%f" (
    git add "%%f" >> pipeline_log.txt 2>&1
  ) else (
    echo   ^(skip^) %%f not found on disk - not staging >> pipeline_log.txt
  )
)

git commit -m "Auto-update: Tableau data %date% %time%" >> pipeline_log.txt 2>&1
if errorlevel 1 (
  echo   i git commit reported an error ^(errorlevel %errorlevel%^) - if this just means "nothing to commit", the next line still pushes any earlier unpushed commit. If it's a real error, check the git output just above this line. >> pipeline_log.txt
)
git push >> pipeline_log.txt 2>&1
if errorlevel 1 (
  echo   X git push FAILED - kpi_pipeline.py's data was generated locally but was NOT deployed to the dashboard. Check the git output just above this line, and check credentials/auth for the account Task Scheduler runs this task as. >> pipeline_log.txt
) else (
  echo   OK git push succeeded - dashboard data deployed. >> pipeline_log.txt
)
