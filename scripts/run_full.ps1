# 取得(1m) → 特徴量 → しきい最適化 → 推論 → (任意)Excelマクロで取込/発注
$BASE = "C:\AI\asagake"
$SCR  = "$BASE\scripts"
$PY   = "py"
$DAY  = (Get-Date).ToString("yyyy-MM-dd")

# 1) 1分足取得（必要に応じてコメントアウト）
& $PY "$SCR\fetch_yahoo_1m.py"

# 2) 特徴量生成
& $PY "$SCR\build_features.py"

# 3) しきい最適化
& $PY "$SCR\optimize_thresholds.py"

# 4) 推論（top_candidates.csv生成）
& $PY "$SCR\infer_reversion.py"

# 5) (任意) Excelを叩いて取り込み＆発注
# $xl = New-Object -ComObject Excel.Application
# $xl.Visible = $true
# $wb = $xl.Workbooks.Open("$BASE\excel\ASAGAKE_template_30_safe.xlsm")
# $xl.Run("ImportTopCandidates")
# $xl.Run("PlaceOrdersNow")  # 実弾はコメント解除の前に慎重に
# $wb.Save()
# $wb.Close($false)
# $xl.Quit()
