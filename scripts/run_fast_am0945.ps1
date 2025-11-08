param(
  [string]$UniverseFile = "C:\\AI\\asagake\\output\\bt30\\NIGHTLY_20251030\\universe_amt_top_150_2025-10-30.csv",
  [string]$OutRoot = "C:\\AI\\asagake\\output\\bt30\\NIGHTLY__FAST",
  [int]$TopK = 12
)

$ErrorActionPreference = 'Stop'
$Repo = "C:\\AI\\asagake"
$Status = Join-Path $Repo 'logs\\fast_nightly_status.txt'

function Write-Status([string]$step, [string]$msg) {
  $now = Get-Date
  $lines = @(
    "updated=$($now.ToString('s'))",
    "step=$step",
    "message=$msg"
  )
  $lines | Out-File -FilePath $Status -Append -Encoding utf8
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
    '--min-train-trades','12','--min-forward-trades','3','--forward-pf-min','1.1',
    '--gap-guard-abs-bp','80.0','--gap-guard-dir-bp','40.0','--slipbp','4.0','--feebp','4.0',
    '--liquidity-quantile','0.5','--jobs','8','--use-local-raw','--run-type','weekday',
    '--enable-asha','--mask-ineffective','--mask-window','20','--mask-threshold','1.15','--mask-keep-j-min','1.35',
    '--codes-file',$UniverseFile,'--excel-summary',
    '--enable-market-features','--market-adjust-j','--market-j-delta-up','0.10','--market-j-delta-down','0.10',
    '--dynamic-risk-j','--tp-per-j','0.15','--sl-per-j','0.10'
  )
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
if pf_col:
    df = df.sort_values([pf_col, tr_col or pf_col], ascending=[False, False])
small = df[[code_col]].head(k).rename(columns={code_col:'code'})
small.to_csv(out, index=False)
print(out)
'@
  & 'C:\\Python313\\python.exe' -c $py $topFile $k $out | Out-Null
  return $out
}

function Run-Refine([string]$label, [string]$sessEnd, [string]$sig, [string]$codesFile) {
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
    '--min-train-trades','15','--min-forward-trades','5','--forward-pf-min','1.3',
    '--gap-guard-abs-bp','80.0','--gap-guard-dir-bp','40.0','--slipbp','4.0','--feebp','4.0',
    '--liquidity-quantile','0.5','--jobs','8','--codes-file',$codesFile,
    '--use-local-raw','--run-type','weekday','--enable-bayes','--bayes-trials','16','--bayes-timeout','600','--excel-summary',
    '--enable-market-features','--market-adjust-j','--market-j-delta-up','0.10','--market-j-delta-down','0.10',
    '--dynamic-risk-j','--tp-per-j','0.15','--sl-per-j','0.10'
  )
  Write-Status "refine-$label-$sig" "start"
  $p = Start-Process -FilePath 'C:\\Python313\\python.exe' -ArgumentList $args -WorkingDirectory $Repo -NoNewWindow -PassThru -Wait
  if ($p.ExitCode -ne 0) { throw "refine failed: $label $sig ($($p.ExitCode))" }
  Write-Status "refine-$label-$sig" "done"
  return $outRefine
}

Write-Status 'start-AM0945' 'begin AM0945 two plans'

foreach ($pl in @(
  @{label='AM0945'; end='09:45'; sig='j-only'},
  @{label='AM0945'; end='09:45'; sig='j-cross'}
)) {
  $coarseDir = Run-Coarse -label $pl.label -sessEnd $pl.end -sig $pl.sig
  $topFile = Join-Path $coarseDir '_TOP_CANDIDATES.csv'
  if (!(Test-Path $topFile)) { Write-Status "coarse-$($pl.label)-$($pl.sig)" 'no candidates file'; continue }
  $codesTop = Select-TopCodes -topFile $topFile -k $TopK
  if (Test-Path $codesTop) {
    Run-Refine -label $pl.label -sessEnd $pl.end -sig $pl.sig -codesFile $codesTop | Out-Null
  } else {
    Write-Status "refine-$($pl.label)-$($pl.sig)" 'skipped (no codes)'
  }
}

Write-Status 'completed-AM0945' 'AM0945 two plans completed'

