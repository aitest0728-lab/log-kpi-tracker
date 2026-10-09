@echo off
REM v8.0 - Daily Cost Report job (Tue + Fri 09:00). Downloads the newest
REM "Daily Cost Report_YYYYMM_Last Update_MMM DD.xlsx" from the WhatsApp group
REM and covers the old manpower / Cost per Order / Productivity data.
cd /d "%~dp0"
python kpi_pipeline.py --section costreport >> cost_report_job.log 2>&1
