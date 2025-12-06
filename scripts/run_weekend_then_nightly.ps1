param(
  [int]$Jobs = 8,
  [int]$BayesTrialsWeekend = 36,
  [int]$BayesTrialsNightly = 24,
  [int]$WeekendTimeoutMinutes = 240,
  [int]$NightlyTimeoutMinutes = 120,
  [int]$PlanTimeoutMinutes = 45,
  [int]$MaxPlanRetries = 2
)

$ErrorActionPreference = 'Stop'
$Repo = 'C:/AI/asagake'
$Py = 'C:/Python313/python.exe'
$StatusPath = Join-Path $Repo 'logs/nightly_status.txt'
$SeqLog = Join-Path $Repo 'logs/sequential_runner.log'

function Write-SeqLog {
  param([string]$Message)
  $line = "[{0}] {1}" -f (Get-Date).ToString('s'), $Message
  $line | Out-File -FilePath $SeqLog -Append -Encoding utf8
}

function Kill-OldRuns {
  try {
    Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*scripts/nightly_build_candidates.py*' } | ForEach-Object {
      Write-SeqLog ("Killing PID {0}: {1}" -f $_.ProcessId, $_.CommandLine)
      Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
  } catch {}
}

function Start-PythonJob {
  param(
    [string]$Label,
    [string[]]$Arguments,
    [int]$TimeoutMinutes
  )
  Write-SeqLog "$Label starting (timeout ${TimeoutMinutes}m)"
  $proc = Start-Process -FilePath $Py -ArgumentList $Arguments -WorkingDirectory $Repo -NoNewWindow -PassThru
  $timeoutMs = [Math]::Max(60000, [int]($TimeoutMinutes * 60000))
  $exited = $proc.WaitForExit($timeoutMs)
  if (-not $exited) {
    Write-SeqLog "$Label timed out after $TimeoutMinutes minutes; killing PID $($proc.Id)"
    try { Stop-Process -Id $proc.Id -Force } catch {}
    return @{ ExitCode = -999; TimedOut = $true }
  }
  $code = $proc.ExitCode
  Write-SeqLog "$Label exited with code $code"
  return @{ ExitCode = $code; TimedOut = $false }
}

function Parse-StatusFile {
  if (-not (Test-Path $StatusPath)) { return $null }
  $dict = @{}
  foreach ($line in Get-Content $StatusPath) {
    if ($line -match '^([^=]+)=(.*)$') {
      $dict[$matches[1]] = $matches[2]
    }
  }
  if ($dict.Count -eq 0) { return $null }
  $plans = @()
  if ($dict.ContainsKey('plans')) {
    $plans = $dict['plans'].Split(',') | Where-Object { $_ }
  }
  $completed = 0
  if ($dict.ContainsKey('completed_plans')) {
    [void][int]::TryParse($dict['completed_plans'], [ref]$completed)
  }
  return [pscustomobject]@{
    RunType = $dict['run_type']
    TargetDate = $dict['target_date']
    Plans = $plans
    Completed = $completed
  }
}

function Get-WeekendPlanOrder {
  $order = @()
  foreach ($window in 'AM15','AM0930','AM0945','AM1015','AM1030') {
    $order += "${window}_j-only"
    $order += "${window}_j-cross"
  }
  $order += 'MID1030_j-cross','PM1230_j-cross'
  return $order
}

function Get-RemainingPlans {
  param($statusInfo)
  if ($null -eq $statusInfo -or -not $statusInfo.Plans -or $statusInfo.Plans.Count -eq 0) {
    return (Get-WeekendPlanOrder)
  }
  $plans = $statusInfo.Plans
  $completed = [Math]::Min($statusInfo.Completed, $plans.Count)
  if ($completed -ge $plans.Count) { return @() }
  return $plans[$completed..($plans.Count - 1)]
}

function Get-WeekendBaseArgs {
  param([string]$TargetDate)
  return @(
    'scripts/nightly_build_candidates.py',
    '--universe-mode','yahoo-top','--universe-size','200',
    '--lookback','60','--chunk-days','5','--train-days','12','--forward-days','4',
    '--min-train-trades','10','--min-forward-trades','2','--forward-pf-min','1.3','--min-forward-winrate','0.60',
    '--gap-guard-abs-bp','80','--gap-guard-dir-bp','40','--slipbp','4','--feebp','4',
    '--liquidity-quantile','0.3','--jobs',"$Jobs",'--run-type','weekend','--plan-profile','weekend',
    '--enable-asha','--enable-bayes','--bayes-trials',"$BayesTrialsWeekend",'--bayes-timeout','600',
    '--mask-ineffective','--mask-window','20','--mask-threshold','1.05','--cache-refresh-weekend',
    '--enable-rd-windows','--enable-market-features','--headless','--coeff-history-days','5',
    '--target-date', $TargetDate
  )
}

function Invoke-WeekendPlan {
  param(
    [string]$PlanTag,
    [string]$TargetDate
  )
  $args = Get-WeekendBaseArgs -TargetDate $TargetDate
  $args += @('--plan-focus', $PlanTag)
  $result = Start-PythonJob "Weekend plan $PlanTag" $args $PlanTimeoutMinutes
  return $result.ExitCode -eq 0
}

function Invoke-WeekendPlanWithRetry {
  param(
    [string]$PlanTag,
    [string]$TargetDate
  )
  for ($attempt = 1; $attempt -le ($MaxPlanRetries + 1); $attempt++) {
    Write-SeqLog "Retrying $PlanTag (attempt $attempt)"
    if (Invoke-WeekendPlan -PlanTag $PlanTag -TargetDate $TargetDate) {
      Write-SeqLog "$PlanTag completed on attempt $attempt"
      return $true
    }
  }
  Write-SeqLog "$PlanTag failed after $($MaxPlanRetries + 1) attempts"
  return $false
}

function Run-WeekendBatch {
  $targetDate = (Get-Date).ToString('yyyyMMdd')
  $args = Get-WeekendBaseArgs -TargetDate $targetDate
  $result = Start-PythonJob 'WEEKEND batch' $args $WeekendTimeoutMinutes
  if ($result.ExitCode -eq 0) {
    return $true
  }
  Write-SeqLog "WEEKEND batch failed with code $($result.ExitCode). Attempting per-plan recovery."
  $statusInfo = Parse-StatusFile
  if ($null -eq $statusInfo -or $statusInfo.RunType -ne 'weekend') {
    $statusInfo = [pscustomobject]@{ RunType='weekend'; TargetDate=$targetDate; Plans=Get-WeekendPlanOrder; Completed=0 }
  }
  if (-not $statusInfo.TargetDate) { $statusInfo.TargetDate = $targetDate }
  $remaining = Get-RemainingPlans -statusInfo $statusInfo
  if ($remaining.Count -eq 0) { $remaining = Get-WeekendPlanOrder }
  foreach ($plan in $remaining) {
    if (-not (Invoke-WeekendPlanWithRetry -PlanTag $plan -TargetDate $statusInfo.TargetDate)) {
      Write-SeqLog "Per-plan recovery failed at $plan"
      return $false
    }
  }
  return $true
}

function Get-NightlyArgs {
  param([string]$TargetDate)
  return @(
    'scripts/nightly_build_candidates.py',
    '--excel','ASAGAKE.xlsm','--excel-summary',
    '--jobs',"$Jobs",'--universe-mode','yahoo-top','--universe-size','150',
    '--run-type','weekday','--plan-profile','weekday',
    '--enable-asha','--enable-bayes','--bayes-trials',"$BayesTrialsNightly",'--bayes-timeout','300',
    '--mask-ineffective','--enable-market-features','--min-forward-winrate','0.60','--headless','--coeff-history-days','5',
    '--target-date',$TargetDate
  )
}

function Run-NightlyBatch {
  $targetDate = (Get-Date).ToString('yyyyMMdd')
  $args = Get-NightlyArgs -TargetDate $targetDate
  $result = Start-PythonJob 'NIGHTLY batch' $args $NightlyTimeoutMinutes
  return $result.ExitCode -eq 0
}

function Repair-AsagakeWorkbook {
  try {
    Write-SeqLog 'Running auto workbook repair'
    & $Py @('scripts/auto_repair_asagake.py','--excel','C:/AI/asagake/ASAGAKE.xlsm') | Out-Null
  }
  catch {
    Write-SeqLog ("auto repair failed: {0}" -f $_.Exception.Message)
  }
}

Write-SeqLog 'Sequential runner started'
Kill-OldRuns
$weekendOk = Run-WeekendBatch
if (-not $weekendOk) {
  Write-SeqLog 'Sequential runner aborting: weekend batch did not complete'
  exit 1
}
$nightlyOk = Run-NightlyBatch
if (-not $nightlyOk) {
  Write-SeqLog 'Sequential runner aborting: nightly batch failed'
  exit 1
}
Repair-AsagakeWorkbook
Write-SeqLog 'Sequential runner completed'
try {
  powershell -NoProfile -ExecutionPolicy Bypass -File scripts\register_boardlogger_task.ps1 | Out-Null
} catch {}

