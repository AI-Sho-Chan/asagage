@echo off
setlocal
cd /d C:\AI\asagake
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\AI\asagake\scripts\run_daily_replay.ps1
exit /b %ERRORLEVEL%

