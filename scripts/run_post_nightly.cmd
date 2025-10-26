@echo off
cd /d C:\AI\asagake
"%~dp0..\..\Python313\python.exe" scripts\post_nightly_tasks.py
if %errorlevel% neq 0 (
  rem fallback to PATH python
  python scripts\post_nightly_tasks.py
)

