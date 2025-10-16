# ===== 繝ｦ繝ｼ繧ｶ繝ｼ險ｭ螳・=====
$BASE = "C:\AI\asagake"
$SCR  = "$BASE\scripts"
$PY   = "py"
$DAY  = (Get-Date).ToString("yyyy-MM-dd")

# ===== 1. 迚ｹ蠕ｴ驥冗函謌・=====
& $PY "$SCR\build_features.py"
if ($LASTEXITCODE -ne 0) { Write-Host "build_features.py 螟ｱ謨・; exit 1 }

# ===== 2. 繝舌ャ繧ｯ繝・せ繝茨ｼ医＠縺阪＞蛟､縺ｮ荳頑嶌縺堺ｾ具ｼ・=====
$bt = Get-Content "$SCR\backtest_speed_kairi.py" -Raw
$bt = $bt -replace 'theta_J\s*=\s*[\d\.]+', 'theta_J = 0.35'
$bt = $bt -replace 'lam\s*=\s*[\d\.]+',     'lam     = 0.8'
$bt = $bt -replace 'kATR\s*=\s*[\d\.]+',    'kATR    = 2.0'
$bt = $bt -replace 'Tmax\s*=\s*\d+',        'Tmax    = 60'
$bt = $bt -replace 'use_session_filter\s*=\s*\w+', 'use_session_filter = $True'
$bt | Set-Content "$SCR\_bt_tmp.py" -Encoding UTF8

& $PY "$SCR\_bt_tmp.py"
if ($LASTEXITCODE -ne 0) { Write-Host "backtest 螟ｱ謨・; exit 1 }

# ===== 3. 結果確認（最も新しい _SUMMARY.csv を開く） =====
$latest = Get-ChildItem "$BASE\data\proc\features_1m" -Recurse -Filter _SUMMARY.csv |
  Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -ne $latest) {
  Write-Host "SUMMARY: $($latest.FullName)"
  Start-Process $latest.FullName
} else {
  Write-Host "SUMMARYが見つかりません"
}
