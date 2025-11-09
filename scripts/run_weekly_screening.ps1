param(
  [string]$DateTag = (Get-Date -Format 'yyyyMMdd'),
  [int]$TopUniverse = 200,
  [int]$TargetTop = 50
)

$ErrorActionPreference = 'Stop'
$Repo = "C:\\AI\\asagake"
$Python = 'C:\\Python313\\python.exe'
$Status = Join-Path $Repo 'logs\weekly_screening_status.txt'

function Write-Status([string]$step, [string]$msg) {
  $now = Get-Date
  @(
    "updated=$($now.ToString('s'))",
    "step=$step",
    "message=$msg"
  ) | Out-File -FilePath $Status -Append -Encoding utf8
}

$targetDate = [datetime]::ParseExact($DateTag, 'yyyyMMdd', $null)
$maxBacktrack = 10
while ($maxBacktrack -gt 0) {
  $checkArgs = @('scripts/check_trading_day.py', '--date', $targetDate.ToString('yyyy-MM-dd'))
  $checkProcess = Start-Process -FilePath $Python -ArgumentList $checkArgs -WorkingDirectory $Repo -NoNewWindow -PassThru -Wait
  if ($checkProcess.ExitCode -eq 0) { break }
  $targetDate = $targetDate.AddDays(-1)
  $maxBacktrack--
}
if ($maxBacktrack -le 0) {
  Write-Status 'error' "could not find prior trading day for $DateTag"
  throw "no trading day found"
}

$resolvedTag = $targetDate.ToString('yyyyMMdd')
$OutRoot = "C:\\AI\\asagake\\output\\bt30\\WEEKLY_$resolvedTag"
if ($resolvedTag -ne $DateTag) {
  Write-Status 'info' "resolved trading day $resolvedTag"
}

try {
  New-Item -ItemType Directory -Force -Path $OutRoot | Out-Null
  Write-Status 'start' "weekly screening begin"

$args = @(
    'scripts/nightly_build_candidates.py',
    '--excel','ASAGAKE.xlsm',
    '--base-out',$OutRoot,
    '--run-type','weekend','--plan-profile','weekend',
    '--target-date',$resolvedTag,
    '--universe-mode','yahoo-top','--universe-size',$TopUniverse,
    '--lookback','60','--chunk-days','5','--train-days','12','--forward-days','4',
    '--min-train-trades','12','--min-forward-trades','5','--forward-pf-min','1.3','--min-forward-winrate','0.60','--min-forward-ci','0.65',
    '--gap-guard-abs-bp','80.0','--gap-guard-dir-bp','40.0','--slipbp','4.0','--feebp','4.0',
    '--liquidity-quantile','0.3','--jobs','6','--enable-asha','--enable-bayes','--bayes-trials','24','--bayes-timeout','600',
    '--mask-ineffective','--mask-window','20','--mask-threshold','1.05',
    '--enable-market-features','--excel-summary','--analysis-ledger','--refine-quick-grid'
)
  $p = Start-Process -FilePath $Python -ArgumentList $args -WorkingDirectory $Repo -NoNewWindow -PassThru -Wait
  if ($p.ExitCode -ne 0) { throw "weekly coarse/refine failed ($($p.ExitCode))" }

  # Aggregate to weekly candidates (~Top50 by score with filters: win>=0.70, pf>=1.30, exp_bp>0)
  $weeklyOut = Join-Path $Repo "output/excel/weekly_candidates_$resolvedTag.csv"
  $aggArgs = @(
    'tools/aggregate_weekly_candidates.py',
    '--date',$resolvedTag,
    '--target-top',"$TargetTop",
    '--output',$weeklyOut
  )
  $aggJson = & $Python $aggArgs | ConvertFrom-Json
  if ($LASTEXITCODE -ne 0) { throw "aggregate weekly failed ($LASTEXITCODE)" }

  if ($aggJson -and $aggJson.written) {
    $written = [System.IO.Path]::GetFullPath($aggJson.written)
    $latest = Join-Path $Repo 'output/excel/weekly_candidates_latest.csv'
    Copy-Item $written $latest -Force
    # 騾ｱ譛ｫ邨先棡繧堤峩縺｡縺ｫ鄙悟霧讌ｭ譌･蛟呵｣懊↓蜿肴丐
    Copy-Item $written (Join-Path $Repo 'output/excel/candidates_nextday.csv') -Force
    Write-Status 'aggregate' "weekly candidates: $($aggJson.rows) rows -> $written"
  }

  # Pre-fetch minute cache for listed tickers to ensure 'ts' exists
  $cacheArgs = @(
    'tools/update_minute_cache.py',
    '--codes-file',$weeklyOut,
    '--history-days','8'
  )
  & $Python $cacheArgs
  if ($LASTEXITCODE -ne 0) { throw "minute cache update failed ($LASTEXITCODE)" }

  # Compute dashboard coefficients with guarded merge
  $coeffArgs = @(
    'tools/compute_dashboard_coeffs.py',
    '--codes-file',$weeklyOut,
    '--history-days','8',
    '--save-dated'
  )
  & $Python $coeffArgs
  if ($LASTEXITCODE -ne 0) { throw "dashboard coeff calc failed ($LASTEXITCODE)" }

  # 成果に基づくルール自動更新（AM1000 SELL の自動拡張可否）
  try {
    & $Python @('tools/auto_rules_from_results.py','--summary','analysis/session_mode_summary.csv','--rules','state/strategy_rules.ini') | Out-Null
    Write-Status 'rules' 'auto_rules_from_results completed'
  }
  catch {
    Write-Status 'rules_error' $_.Exception.Message
  }

  Write-Status 'completed' "weekly screening completed"
  try {
    & "$Repo\scripts\run_trade_analysis.ps1" | Out-Null
    Write-Status 'analysis' "run_trade_analysis completed"
  }
  catch {
    Write-Status 'analysis_error' $_.Exception.Message
  }
}
catch {
  Write-Status 'error' ($_.Exception.Message)
  throw
}

try {
  & C:\\Python313\\python.exe scripts\\register_boardlogger_task.ps1 | Out-Null
  Write-Status 'boardlogger' 'registered'
} catch { Write-Status 'boardlogger_error' .Exception.Message }

