param(
  [int]$Jobs = 8,
  [int]$BayesTrialsWeekend = 36,
  [int]$BayesTrialsNightly = 24
)

$ErrorActionPreference = 'Stop'
$Repo = "C:\AI\asagake"
$Py = "C:\\Python313\\python.exe"
$Status = Join-Path $Repo 'logs\nightly_status.txt'
$SeqLog = Join-Path $Repo 'logs\sequential_runner.log'

function Write-SeqLog([string]$msg) {
  $line = "[{0}] {1}" -f (Get-Date).ToString('s'), $msg
  $line | Out-File -FilePath $SeqLog -Append -Encoding utf8
}

function Kill-OldRuns() {
  try {
    Get-CimInstance Win32_Process | Where-Object { $_.CommandLine -like '*scripts/nightly_build_candidates.py*' } | ForEach-Object {
      Write-SeqLog ("Killing PID {0}: {1}" -f $_.ProcessId, $_.CommandLine)
      Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }
  } catch {}
}

function Run-Weekend() {
  Write-SeqLog 'Starting WEEKEND batch (10 windows + R&D windows)'
  $args = @(
    'scripts/nightly_build_candidates.py',
    '--universe-mode','yahoo-top','--universe-size','200',
    '--lookback','60','--chunk-days','5','--train-days','12','--forward-days','4',
    '--min-train-trades','10','--min-forward-trades','2','--forward-pf-min','1.3',
    '--gap-guard-abs-bp','80','--gap-guard-dir-bp','40','--slipbp','4','--feebp','4',
    '--liquidity-quantile','0.5','--jobs',"$Jobs",'--run-type','weekend','--plan-profile','weekend',
    '--enable-asha','--enable-bayes','--bayes-trials',"$BayesTrialsWeekend",'--bayes-timeout','600',
    '--mask-ineffective','--mask-window','20','--mask-threshold','1.05','--cache-refresh-weekend',
    '--enable-rd-windows','--enable-market-features'
  )
  $p = Start-Process -FilePath $Py -ArgumentList $args -WorkingDirectory $Repo -PassThru -Wait
  Write-SeqLog ("WEEKEND batch exit code: {0}" -f $p.ExitCode)
}

function Run-Nightly() {
  Write-SeqLog 'Starting NIGHTLY batch (weekday mode)'
  $args = @(
    'scripts/nightly_build_candidates.py',
    '--excel','SHINSOKU.xlsm','--excel-summary',
    '--jobs',"$Jobs",'--universe-mode','yahoo-top','--universe-size','150',
    '--run-type','weekday','--plan-profile','weekday',
    '--enable-asha','--enable-bayes','--bayes-trials',"$BayesTrialsNightly",'--bayes-timeout','300',
    '--mask-ineffective',
    '--mask-ineffective','--enable-market-features'
  )
  $p = Start-Process -FilePath $Py -ArgumentList $args -WorkingDirectory $Repo -PassThru -Wait
  Write-SeqLog ("NIGHTLY batch exit code: {0}" -f $p.ExitCode)
}

Write-SeqLog 'Sequential runner started'
Kill-OldRuns
Run-Weekend
Run-Nightly
Write-SeqLog 'Sequential runner completed'
