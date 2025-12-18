param(
  [string]$DateTag = (Get-Date -Format "yyyyMMdd"),
  [double]$Nominal = 10000000,
  [int]$MaxAttempts = 12,
  [int]$RetrySleepSeconds = 600
)

$ErrorActionPreference = "Stop"

$repo = "C:/AI/asagake"
$python = "C:/Python313/python.exe"

Set-Location $repo

if (-not $DateTag) {
  $DateTag = (Get-Date -Format "yyyyMMdd")
}

$logDir = Join-Path $repo "logs"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logPath = Join-Path $logDir ("daily_replay_task_{0}.log" -f $DateTag)

function Write-Log {
  param([string]$Message)
  $ts = Get-Date -Format s
  ("[{0}] {1}" -f $ts, $Message) | Out-File -FilePath $logPath -Append -Encoding utf8
}

Write-Log ("=== DailyReplay start (date={0}) ===" -f $DateTag)

$summaryPath = Join-Path $repo ("analysis/daily_replay_{0}.json" -f $DateTag)

$simulateOk = $false
$lastSimulateOutput = ""
$attempt = 1
while (-not $simulateOk -and $attempt -le $MaxAttempts) {
  try {
    Write-Log ("simulate attempt {0}/{1}" -f $attempt, $MaxAttempts)

    $pyArgs = @(
      "tools/simulate_daily_replay.py",
      "--date", $DateTag,
      "--nominal", [string][int][math]::Round($Nominal)
    )

    $out = & $python @pyArgs 2>&1
    $lastSimulateOutput = ($out | Out-String)
    $out | Out-File -FilePath $logPath -Append -Encoding utf8
    if ($LASTEXITCODE -ne 0) {
      throw "simulate_daily_replay.py failed (exit=$LASTEXITCODE)"
    }

    $simulateOk = $true
  } catch {
    Write-Log ("[warn] simulate failed: {0}" -f $_.Exception.Message)
    if ($attempt -lt $MaxAttempts) {
      Write-Log ("sleep {0}s then retry" -f $RetrySleepSeconds)
      Start-Sleep -Seconds $RetrySleepSeconds
    }
    $attempt += 1
  }
}

if (-not $simulateOk) {
  Write-Log ("[error] simulate did not succeed after {0} attempts; writing placeholder summary" -f $MaxAttempts)
  $placeholder = @{
    date = $DateTag
    status = "NOT_READY"
    attempts = $MaxAttempts
    note = "Yahoo 1m data may not be ready yet; rerun later."
  } | ConvertTo-Json -Depth 4
  $placeholder | Out-File -FilePath $summaryPath -Encoding utf8
}

# If `analysis/daily_replay_<date>.json` and `state/smtp.json` exist, send summary email.
try {
  $smtpPath = Join-Path $repo "state/smtp.json"

  if ((Test-Path $summaryPath) -and (Test-Path $smtpPath)) {
    $summary = Get-Content $summaryPath -Raw | ConvertFrom-Json
    $smtpConf = Get-Content $smtpPath -Raw | ConvertFrom-Json

    $cred = New-Object System.Management.Automation.PSCredential(
      $smtpConf.user,
      (ConvertTo-SecureString $smtpConf.pass -AsPlainText -Force)
    )

    $dailyTradesPath = Join-Path $repo ("analysis/daily_trades_{0}.csv" -f $DateTag)

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(("ASAGAKE DailyReplay（仮想売買）レポート: {0}" -f $DateTag))
    $lines.Add("")

    if ($summary.status -eq "NOT_READY") {
      $lines.Add("結果: まだデータが揃っていない可能性があり、今回の自動実行では完了できませんでした。")
      $lines.Add(("補足: {0}" -f $summary.note))
      $lines.Add("対策: 自動で再試行します（次の実行または手動再実行で更新されます）。")
    } else {
      $trades = [int]($summary.trades)
      $pnlYen = [double]($summary.pnl_yen)
      $pnlBpMean = $summary.pnl_bp_mean

      $lines.Add(("結果サマリ: 取引 {0} 回 / 合計損益 {1:N0} 円" -f $trades, $pnlYen))
      if ($null -ne $pnlBpMean) {
        $lines.Add(("（平均損益: {0:N1} bp）" -f [double]$pnlBpMean))
      }

      if ($summary.PSObject.Properties.Name -contains "diag_skip_trend_mismatch") {
        $lines.Add(("見送り: 方向不一致 {0} 件 / シグナルなし {1} 件" -f $summary.diag_skip_trend_mismatch, $summary.diag_no_signal))
      }

      $lines.Add("")
      $lines.Add("クラス別（強/標準/デモ）:")
      if ($summary.PSObject.Properties.Name -contains "LIVE_STRONG_trades") {
        $lines.Add(("  LIVE_STRONG: {0}回, {1:N0}円" -f $summary.LIVE_STRONG_trades, $summary.LIVE_STRONG_pnl_yen))
        $lines.Add(("  LIVE_BASE:   {0}回, {1:N0}円" -f $summary.LIVE_BASE_trades, $summary.LIVE_BASE_pnl_yen))
        $lines.Add(("  DEMO_ONLY:   {0}回, {1:N0}円" -f $summary.DEMO_ONLY_trades, $summary.DEMO_ONLY_pnl_yen))
      } else {
        $lines.Add("  （この日の詳細分類がありませんでした）")
      }

      if (Test-Path $dailyTradesPath) {
        $tradesRows = Import-Csv $dailyTradesPath
        if ($tradesRows.Count -gt 0) {
          foreach ($row in $tradesRows) {
            if ($row.pnl_yen -ne $null) {
              $row.pnl_yen = [double]$row.pnl_yen
            }
            if ($row.pnl_bp -ne $null) {
              $row.pnl_bp = [double]$row.pnl_bp
            }
            if ($row.budget_factor -ne $null) {
              $row.budget_factor = [double]$row.budget_factor
            }
          }

          $lines.Add("")
          $lines.Add("負けが大きかった順（上位5件）:")
          $losses = $tradesRows | Sort-Object -Property pnl_yen | Select-Object -First 5
          foreach ($r in $losses) {
            $lines.Add(("  {0} {1} {2} {3}: {4:N0}円 ({5:N1}bp) [{6}] x{7}" -f $r.code, $r.session, $r.signal_mode, $r.side, $r.pnl_yen, $r.pnl_bp, $r.live_demo_class, $r.budget_factor))
          }

          $lines.Add("")
          $lines.Add("勝ちが大きかった順（上位5件）:")
          $wins = $tradesRows | Sort-Object -Property pnl_yen -Descending | Select-Object -First 5
          foreach ($r in $wins) {
            $lines.Add(("  {0} {1} {2} {3}: {4:N0}円 ({5:N1}bp) [{6}] x{7}" -f $r.code, $r.session, $r.signal_mode, $r.side, $r.pnl_yen, $r.pnl_bp, $r.live_demo_class, $r.budget_factor))
          }

          $minLoss = ($tradesRows | Measure-Object -Property pnl_yen -Minimum).Minimum
          $lines.Add("")
          $lines.Add("かんたんな解説:")
          if ($minLoss -lt 0 -and [Math]::Abs($minLoss) -gt [Math]::Abs($pnlYen) * 0.6) {
            $lines.Add("  1つの大きな負けが、1日の結果をほぼ決めています（大負けを減らすと安定します）。")
          } else {
            $lines.Add("  いくつかの勝ち/負けの合計で結果が決まっています（大負けの有無を確認してください）。")
          }
          $lines.Add("  ※これは「取引終了後に、Yahooの1分足データで再現した仮想売買」です。Excelの場中ログと一致しないことがあります。")
        }
      } else {
        $lines.Add("")
        $lines.Add(("注意: {0} が見つからないため、上位の勝ち/負け内訳は省略します。" -f $dailyTradesPath))
      }
    }

    $body = ($lines -join [Environment]::NewLine)

    $mailParams = @{
      To         = "shouichi.ikeda@gmail.com"
      From       = $smtpConf.user
      Subject    = ("ASAGAKE DailyReplay {0}" -f $DateTag)
      Body       = $body
      SmtpServer = $smtpConf.host
      Port       = [int]$smtpConf.port
      UseSsl     = $true
      Credential = $cred
    }

    Send-MailMessage @mailParams
    Write-Log "Mail sent to shouichi.ikeda@gmail.com"
  } else {
    Write-Log "No smtp.json or summary not found; skip mail."
  }
} catch {
  Write-Log ("Mail send failed: {0}" -f $_.Exception.Message)
}

Write-Log "=== DailyReplay end ==="
