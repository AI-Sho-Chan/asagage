param(
  [string]$WorkbookPath = "C:\AI\asagake\ASAGAKE.xlsm",
  [string]$DashboardSheet = "NewDashboardV2",
  [string]$OrdersSheet = "Orders",
  [int]$MaxRows = 600
)

$ErrorActionPreference = "Stop"

function Get-HeaderMap {
  param(
    $Worksheet,
    [int]$HeaderRow,
    [int]$LastCol
  )

  $hdr = $Worksheet.Range($Worksheet.Cells($HeaderRow, 1), $Worksheet.Cells($HeaderRow, $LastCol)).Value2
  $headers = @()
  for ($c = 1; $c -le $LastCol; $c++) {
    $h = $hdr[1, $c]
    if ($null -eq $h) { $headers += "" } else { $headers += [string]$h }
  }

  $map = @{}
  for ($i = 0; $i -lt $headers.Count; $i++) {
    $name = $headers[$i]
    if (-not $name) { continue }
    if (-not $map.ContainsKey($name)) {
      $map[$name] = $i + 1
    }
  }
  return $map
}

try {
  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
  Write-Output "NO_ACTIVE_EXCEL: Excel is not running. Open ASAGAKE.xlsm and run again."
  exit 2
}

$wb = $null
foreach ($w in @($excel.Workbooks)) {
  if ($w.FullName -eq $WorkbookPath) { $wb = $w; break }
}
if (-not $wb) {
  Write-Output "WORKBOOK_NOT_OPEN: $WorkbookPath"
  exit 2
}

$ws = $wb.Worksheets.Item($DashboardSheet)
$used = $ws.UsedRange
$lastRow = $used.Row + $used.Rows.Count - 1
$lastCol = $used.Column + $used.Columns.Count - 1
if ($lastRow -gt $MaxRows) { $lastRow = $MaxRows }

$headerRow = 5
$dataStartRow = 6
$map = Get-HeaderMap -Worksheet $ws -HeaderRow $headerRow -LastCol $lastCol

if (-not $map.ContainsKey("J_th")) {
  Write-Output "HEADER_NOT_FOUND: J_th"
  exit 2
}

$gapThreshold = $ws.Range("G2").Value2

Write-Output ("Workbook: {0}" -f $WorkbookPath)
Write-Output ("Rows: {0} (data {1}-{0}) Cols: {2}" -f $lastRow, $dataStartRow, $lastCol)
Write-Output ("Index B2/C2/E2/F2: {0} / {1} / {2} / {3}" -f $ws.Range("B2").Text, $ws.Range("C2").Text, $ws.Range("E2").Text, $ws.Range("F2").Text)
Write-Output ("TrendBox B3: {0}" -f $ws.Range("B3").Text)
Write-Output ("TrendBox D3: {0}" -f $ws.Range("D3").Text)
Write-Output ("Gap threshold (G2): {0}" -f $gapThreshold)

$rows = $lastRow - $headerRow
$colJ = $map["J_th"]
$colGap = $map["Gap_bp"]
$colTicker = $map["Ticker"]
$colEntrySide = $map["EntrySide"]
$colPolicy = $map["trend_allowed_policy"]
$colDriverAllow = $map["driver_allowed_side"]

$jVals = $ws.Range($ws.Cells($dataStartRow, $colJ), $ws.Cells($lastRow, $colJ)).Value2
$gapVals = $ws.Range($ws.Cells($dataStartRow, $colGap), $ws.Cells($lastRow, $colGap)).Value2
$tikVals = $ws.Range($ws.Cells($dataStartRow, $colTicker), $ws.Cells($lastRow, $colTicker)).Value2
$entryVals = $ws.Range($ws.Cells($dataStartRow, $colEntrySide), $ws.Cells($lastRow, $colEntrySide)).Value2
$polVals = $ws.Range($ws.Cells($dataStartRow, $colPolicy), $ws.Cells($lastRow, $colPolicy)).Value2
$drvVals = $ws.Range($ws.Cells($dataStartRow, $colDriverAllow), $ws.Cells($lastRow, $colDriverAllow)).Value2

$xlErrValue = -2146826273
$counts = @{
  BAN   = 0
  ERR   = 0
  NUM   = 0
  BLANK = 0
  OTHER = 0
}

$banRows = @()
$activeRows = 0
for ($i = 1; $i -le $rows; $i++) {
  $tickerHere = $tikVals[$i, 1]
  if ($null -eq $tickerHere -or $tickerHere -eq "") { continue }
  $activeRows++
  $v = $jVals[$i, 1]
  if ($null -eq $v -or $v -eq "") { $counts.BLANK++; continue }
  if ($v -is [string]) {
    if ($v -eq "BAN") {
      $counts.BAN++
      $banRows += $i
      continue
    }
    $counts.OTHER++; continue
  }
  if ($v -is [double] -or $v -is [int]) {
    if ($v -eq $xlErrValue) { $counts.ERR++; continue }
    if ($v -is [int] -and $v -lt 0) { $counts.ERR++; continue }
    $counts.NUM++; continue
  }
  $counts.OTHER++
}

Write-Output ("Active tickers rows: {0}" -f $activeRows)
Write-Output ("J_th summary (active rows only): NUM={0} BAN={1} ERR={2} BLANK={3} OTHER={4}" -f $counts.NUM, $counts.BAN, $counts.ERR, $counts.BLANK, $counts.OTHER)

if ($banRows.Count -gt 0) {
  Write-Output "BAN rows (why):"
  foreach ($i in $banRows) {
    $row = $i + ($dataStartRow - 1)
    $ticker = $tikVals[$i, 1]
    $gap = $gapVals[$i, 1]
    $entry = $entryVals[$i, 1]
    $pol = $polVals[$i, 1]
    $drv = $drvVals[$i, 1]

    $gapBan = $false
    if ($gap -is [double] -or $gap -is [int]) {
      if ([math]::Abs([double]$gap) / 100 -gt [double]$gapThreshold) { $gapBan = $true }
    }

    $dirBan = $false
    if ($pol -ne $null -and $pol.ToString().ToUpper() -eq "ALIGNED_ONLY") {
      if ($drv -ne $null -and $drv.ToString().ToUpper() -ne "BOTH") {
        if ($entry -ne $null -and $drv.ToString().ToUpper() -ne $entry.ToString().ToUpper()) { $dirBan = $true }
      }
    }

    $reasons = @()
    if ($gapBan) { $reasons += "GAP_TOO_BIG" }
    if ($dirBan) { $reasons += "DIRECTION_MISMATCH" }
    if ($reasons.Count -eq 0) { $reasons += "OTHER" }

    Write-Output ("  row={0} ticker={1} gap_bp={2} entry={3} policy={4} driver_allowed={5} reason={6}" -f $row, $ticker, $gap, $entry, $pol, $drv, ($reasons -join "+"))
  }
}

try {
  $ow = $wb.Worksheets.Item($OrdersSheet)
  $usedO = $ow.UsedRange
  $lastO = $usedO.Row + $usedO.Rows.Count - 1
  if ($lastO -ge 2) {
    $start = [math]::Max(2, $lastO - 9)
    $rng = $ow.Range($ow.Cells($start, 1), $ow.Cells($lastO, 8)).Value2
    Write-Output "Orders tail (ts,ticker,side,mode,status,note):"
    for ($r = 1; $r -le ($lastO - $start + 1); $r++) {
      Write-Output ("  {0} | {1} | {2} | {3} | {4} | {5}" -f $rng[$r, 1], $rng[$r, 2], $rng[$r, 3], $rng[$r, 6], $rng[$r, 7], $rng[$r, 8])
    }
  } else {
    Write-Output "Orders: (empty)"
  }
} catch {
  Write-Output ("Orders read failed: {0}" -f $_.Exception.Message)
}
