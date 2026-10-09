@echo off
REM Run this ONCE (double-click). It registers the daily 09:45 Task Scheduler job.
setlocal
set "BAT=%~dp0run_maptemp_0945.bat"
if not exist "%BAT%" (
  echo Cannot find %BAT%
  pause
  exit /b 1
)
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
 "$a = New-ScheduledTaskAction -Execute '%BAT%' -WorkingDirectory '%~dp0'; " ^
 "$t = New-ScheduledTaskTrigger -Daily -At 09:45; " ^
 "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -MultipleInstances IgnoreNew; " ^
 "Register-ScheduledTask -TaskName 'KPI Temp Delivery Points 0945' -Action $a -Trigger $t -Settings $s -Description 'kpi_pipeline.py --section maptemp then push dashboard' -Force"
if errorlevel 1 (
  echo Failed - try right-click ^> Run as administrator.
) else (
  echo Done. Task "KPI Temp Delivery Points 0945" will run daily at 09:45.
  echo Test it now with:  schtasks /Run /TN "KPI Temp Delivery Points 0945"
)
pause
