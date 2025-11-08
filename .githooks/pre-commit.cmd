@echo off
REM Guard against reintroducing openpyxl-based writes to SHINSOKU.xlsm
python scripts\guard_no_openpyxl_xlsm.py
if errorlevel 1 (
  exit /b 1
)
exit /b 0

