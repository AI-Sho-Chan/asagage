@echo off
setlocal

set "WB=C:\AI\asagake\ASAGAKE.xlsm"
if not exist "%WB%" (
  echo Workbook not found: %WB%
  exit /b 1
)

rem Open in a dedicated Excel process (isolates ASAGAKE timer VBA)
start "" excel.exe /x "%WB%"

