@echo off
cd /d C:\AI\asagake
for /f "tokens=1 delims=:" %%H in ('time /t') do set CUR_HOUR=%%H
for /f "tokens=2 delims=:" %%M in ('time /t') do set CUR_MIN=%%M
set TARGET=16:30
for /f "tokens=1,2 delims=:" %%a in ("%TARGET%") do (
  set TH=%%a
  set TM=%%b
)
setlocal enabledelayedexpansion
for /f "tokens=1,2 delims=: " %%a in ('time /t') do set NOW=%%a:%%b
set NOW=%time:~0,5%
if "%NOW%"=="" set NOW=00:00

powershell -NoProfile -Command "while(([datetime]::Now.TimeOfDay) -lt [TimeSpan]::Parse('16:30:00')){Start-Sleep -Seconds 30}"
endlocal

C:\Python313\python.exe scripts\nightly_build_candidates.py --excel SHINSOKU.xlsm --excel-summary --jobs 8

