param(
  [string]$WorkbookPath = "C:\AI\asagake\ASAGAKE.xlsm"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $WorkbookPath)) {
  throw "Workbook not found: $WorkbookPath"
}

# Open ASAGAKE in a dedicated Excel *process* so its timer VBA (AutoTickV2) does not block other workbooks.
# /x forces a new instance.
Start-Process -FilePath "excel.exe" -ArgumentList @("/x", $WorkbookPath)

