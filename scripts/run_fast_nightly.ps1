param(
  [string]$DateTag = (Get-Date -Format 'yyyyMMdd'),
  [string]$UniverseFile = "C:\\AI\\asagake\\output\\bt30\\NIGHTLY_20251030\\universe_amt_top_300_2025-10-30.csv",
  [string]$OutRoot = "C:\\AI\\asagake\\output\\bt30\\NIGHTLY_${((Get-Date).ToString('yyyyMMdd'))}_FAST",
  [int]$TopK = 6
)

$ErrorActionPreference = 'Stop'
$Repo = "C:\\AI\\asagake"
$Status = "C:\\AI\\asagake\\logs\\fast_nightly_status.txt"

if ([string]::IsNullOrWhiteSpace($DateTag)) {
  $DateTag = (Get-Date -Format 'yyyyMMdd')
}

if (-not $PSBoundParameters.ContainsKey('OutRoot') -or [string]::IsNullOrWhiteSpace($OutRoot)) {
  $OutRoot = "C:\\AI\\asagake\\output\\bt30\\NIGHTLY_{0}_FAST" -f $DateTag
}

function Write-Status([string]$step, [string]$msg) {
  $now = Get-Date
  $lines = @(
    "updated=$($now.ToString('s'))",
    "step=$step",
    "message=$msg"
  )
  $lines | Out-File -FilePath $Status -Append -Encoding utf8
}

function Resolve-UniverseFile([string]$path) {
  if (Test-Path $path) {
    return $path
  }
  $pattern = Join-Path $Repo 'output\bt30\NIGHTLY_*\universe_amt_top_300_*.csv'
  $candidate = Get-ChildItem $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  if (-not $candidate) {
    throw "Universe file not found. Expected $path and no fallback matched."
  }
  Write-Status 'universe-fallback' $candidate.FullName
  return $candidate.FullName
}

function Run-Coarse([string]$label, [string]$sessEnd, [string]$sig) {
  $outCoarse = Join-Path $OutRoot ("RUN_coarse_{0}_{1}" -f $label,$sig)
  New-Item -ItemType Directory -Force -Path $outCoarse | Out-Null
  $args = @(
    'scripts/bt_opt30_forward.py',
    '--excel','SHINSOKU.xlsm',
    '--outdir',$outCoarse,
    '--mode','coarse',
    '--signal-mode',$sig,
    '--session-start','09:00','--session-end',$sessEnd,
    '--lookback','60','--chunk-days','5','--train-days','12','--forward-days','4',
    '--min-train-trades','8','--min-forward-trades','2','--forward-pf-min','0.9',
    '--gap-guard-abs-bp','80.0','--gap-guard-dir-bp','40.0','--slipbp','4.0','--feebp','4.0',
    '--liquidity-quantile','0.5','--repeat-mask-threshold','10','--jobs','8','--use-local-raw','--run-type','weekday',
    '--enable-asha','--mask-ineffective','--mask-window','20','--mask-threshold','1.05','--mask-keep-j-min','1.35',
    '--codes-file',$UniverseFile,'--excel-summary'
  )
  $args += '--no-low-priority'
  # 市場×J調整と動的TP/SLをON
  $args += @('--enable-market-features','--market-adjust-j','--market-j-delta-up','0.10','--market-j-delta-down','0.10','--dynamic-risk-j','--tp-per-j','0.15','--sl-per-j','0.10')
  Write-Status "coarse-$label-$sig" "start"
  $p = Start-Process -FilePath 'C:\\Python313\\python.exe' -ArgumentList $args -WorkingDirectory $Repo -NoNewWindow -PassThru -Wait
  if ($p.ExitCode -ne 0) { throw "coarse failed: $label $sig ($($p.ExitCode))" }
  Write-Status "coarse-$label-$sig" "done"
  return $outCoarse
}

function Select-TopCodes([string]$topFile, [int]$k) {
  $out = [System.IO.Path]::ChangeExtension($topFile, $null) + ("_TOP{0}_codes.csv" -f $k)
$py = @'
import pandas as pd, sys
inp, k, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
df = pd.read_csv(inp)
if df.empty:
    open(out,'w').write('code\n')
    sys.exit(0)
cols = {c.lower():c for c in df.columns}
code_col = cols.get('code') or cols.get('ticker')
if code_col is None:
    raise SystemExit('no code/ticker in candidates')
pf_col = cols.get('forward_pf_eff'); tr_col = cols.get('forward_trades')
win_col = cols.get('forward_winrate')
if pf_col and win_col:
    df = df.sort_values([pf_col, win_col, tr_col or pf_col], ascending=[False, False, False])
elif pf_col:
    df = df.sort_values([pf_col, tr_col or pf_col], ascending=[False, False])
elif win_col:
    df = df.sort_values([win_col], ascending=False)
small = df[[code_col]].head(k).rename(columns={code_col:'code'})
small.to_csv(out, index=False)
print(out)
'@
  & 'C:\\Python313\\python.exe' -c $py $topFile $k $out | Out-Null
  return $out
}

function Run-Refine([string]$label, [string]$sessEnd, [string]$sig, [string]$codesFile, [bool]$analysisLedger, [bool]$marketFeatures) {
  $outRefine = Join-Path $OutRoot ("RUN_refine_{0}_{1}" -f $label,$sig)
  New-Item -ItemType Directory -Force -Path $outRefine | Out-Null
  $args = @(
    'scripts/bt_opt30_forward.py',
    '--excel','SHINSOKU.xlsm',
    '--outdir',$outRefine,
    '--mode','refine',
    '--signal-mode',$sig,
    '--session-start','09:00','--session-end',$sessEnd,
    '--lookback','60','--chunk-days','5','--train-days','12','--forward-days','4',
    '--min-train-trades','10','--min-forward-trades','3','--forward-pf-min','1.1',
    '--gap-guard-abs-bp','80.0','--gap-guard-dir-bp','40.0','--slipbp','4.0','--feebp','4.0',
    '--liquidity-quantile','0.5','--repeat-mask-threshold','10','--jobs','8','--codes-file',$codesFile,
    '--use-local-raw','--run-type','weekday','--enable-bayes','--bayes-trials','16','--bayes-timeout','600','--excel-summary','--quick-grid','--optimize-io'
  )
  $args += '--no-low-priority'
  if ($marketFeatures) { $args += '--enable-market-features' }
  if ($marketFeatures) { $args += @('--market-adjust-j','--market-j-delta-up','0.10','--market-j-delta-down','0.10','--dynamic-risk-j','--tp-per-j','0.15','--sl-per-j','0.10') }
  if ($analysisLedger) { $args += '--analysis-ledger' }
  Write-Status "refine-$label-$sig" "start"
  $p = Start-Process -FilePath 'C:\\Python313\\python.exe' -ArgumentList $args -WorkingDirectory $Repo -NoNewWindow -PassThru -Wait
  if ($p.ExitCode -ne 0) { throw "refine failed: $label $sig ($($p.ExitCode))" }
  Write-Status "refine-$label-$sig" "done"
  return $outRefine
}

try {
  $UniverseFile = Resolve-UniverseFile $UniverseFile

  # Orchestrate 4 plans from scratch
  New-Item -ItemType Directory -Force -Path $OutRoot | Out-Null
  Write-Status 'start' "fast nightly begin"

  $plans = @(
    @{label='AM0930'; end='09:30'; sig='j-only';  analysis=$false; market=$true},
    @{label='AM0930'; end='09:30'; sig='j-cross'; analysis=$true;  market=$true},
    @{label='AM0945'; end='09:45'; sig='j-only';  analysis=$false; market=$false},
    @{label='AM0945'; end='09:45'; sig='j-cross'; analysis=$false; market=$false}
  )

  foreach ($pl in $plans) {
    $coarseDir = Run-Coarse -label $pl.label -sessEnd $pl.end -sig $pl.sig
    $topFile = Join-Path $coarseDir '_TOP_CANDIDATES.csv'
    if (!(Test-Path $topFile)) { Write-Status "coarse-$($pl.label)-$($pl.sig)" 'no candidates file'; continue }
    $codesTop = Select-TopCodes -topFile $topFile -k $TopK
    Run-Refine -label $pl.label -sessEnd $pl.end -sig $pl.sig -codesFile $codesTop -analysisLedger $pl.analysis -marketFeatures $pl.market | Out-Null
  }

  Write-Status 'completed' "fast nightly completed"
}
catch {
  Write-Status 'error' ($_.Exception.Message)
  throw
}
