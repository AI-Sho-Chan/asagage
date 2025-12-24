Param(
  [string]$WorkbookPath = "C:\AI\asagake\ASAGAKE.xlsm",
  [string]$ModuleBasPath = "C:\AI\asagake\excel\AutoTraderAdvanced.bas",
  [switch]$Force
)

$ErrorActionPreference = "Stop"

function New-Timestamp {
  (Get-Date).ToString("yyyyMMdd_HHmmss")
}

function New-TempBasWithoutBom {
  param(
    [Parameter(Mandatory = $true)][string]$SourcePath
  )

  $bytes = [System.IO.File]::ReadAllBytes($SourcePath)
  $outBytes = $bytes

  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    # UTF-8 BOM -> strip only
    $outBytes = $bytes[3..($bytes.Length - 1)]
  } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    # UTF-16LE BOM -> decode and re-encode as UTF-8 (no BOM)
    $text = [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2)
    $outBytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    # UTF-16BE BOM -> decode and re-encode as UTF-8 (no BOM)
    $enc = [System.Text.Encoding]::BigEndianUnicode
    $text = $enc.GetString($bytes, 2, $bytes.Length - 2)
    $outBytes = [System.Text.Encoding]::UTF8.GetBytes($text)
  }

  $tmp = Join-Path $env:TEMP ("asagake_vba_import_{0}.bas" -f (New-Timestamp))
  [System.IO.File]::WriteAllBytes($tmp, $outBytes)
  return $tmp
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
$base = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath)
$backupPath = Join-Path $dir ("{0}_backup_{1}.xlsm" -f $base, $stamp)
Copy-Item -LiteralPath $WorkbookPath -Destination $backupPath -Force
Write-Host "Backup created: $backupPath"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$tmpBas = $null
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

  # Import .bas (strip/convert BOM to avoid invisible char issues in VBE)
  $tmpBas = New-TempBasWithoutBom -SourcePath $ModuleBasPath
  $imported = $components.Import($tmpBas)
  if ($imported.Name -ne $targetName) {
    $imported.Name = $targetName
  }

  $wb.Save()
  Write-Host "Updated VBA module '$targetName' from: $ModuleBasPath"
} finally {
  if ($tmpBas -ne $null) {
    try { Remove-Item -LiteralPath $tmpBas -Force } catch {}
  }
  try { $wb.Close($true) } catch {}
  $excel.Quit() | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

Write-Host "Done."
