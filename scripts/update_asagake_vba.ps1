Param(
  [string]$WorkbookPath = "C:\AI\asagake\ASAGAKE.xlsm",
  [string]$ModuleBasPath = "C:\AI\asagake\excel\AutoTraderAdvanced.bas",
  [switch]$Force
)

$ErrorActionPreference = "Stop"

function New-Timestamp {
  (Get-Date).ToString("yyyyMMdd_HHmmss")
}

if (!(Test-Path -LiteralPath $WorkbookPath)) {
  throw "Workbook not found: $WorkbookPath"
}
if (!(Test-Path -LiteralPath $ModuleBasPath)) {
  throw "Module BAS not found: $ModuleBasPath"
}

# Guard: don't touch if Excel already has this workbook open unless -Force.
$runningExcel = $null
try {
  $runningExcel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
  $runningExcel = $null
}
if ($runningExcel -ne $null) {
  foreach ($wb in $runningExcel.Workbooks) {
    if ($wb.FullName -ieq $WorkbookPath) {
      if (-not $Force) {
        throw "ASAGAKE.xlsm is open in Excel. Close it first, or rerun with -Force."
      }
    }
  }
}

# Backup first
$stamp = New-Timestamp
$dir = Split-Path -Parent $WorkbookPath
$base = Split-Path -LeafBase $WorkbookPath
$ext = Split-Path -Leaf $WorkbookPath
$backupPath = Join-Path $dir ("{0}_backup_{1}.xlsm" -f $base, $stamp)
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
Write-Host "Backup created: $backupPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
  $wb = $excel.Workbooks.Open($WorkbookPath, $null, $false)

  # Access to VBProject requires Excel setting:
  # Trust Center -> Macro Settings -> "Trust access to the VBA project object model"
  try {
    $vbproj = $wb.VBProject
  } catch {
    throw "Cannot access VBProject. Enable 'Trust access to the VBA project object model' in Excel Trust Center, then rerun."
  }

  $components = $vbproj.VBComponents
  $targetName = "AutoTraderAdvanced"

  # Remove existing module (if any)
  foreach ($c in @($components)) {
    if ($c.Name -eq $targetName) {
      $components.Remove($c)
      break
    }
  }

  # Import .bas
  $imported = $components.Import($ModuleBasPath)
  if ($imported.Name -ne $targetName) {
    $imported.Name = $targetName
  }

  $wb.Save()
  Write-Host "Updated VBA module '$targetName' from: $ModuleBasPath"
} finally {
  try { $wb.Close($true) } catch {}
  $excel.Quit() | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Done."

