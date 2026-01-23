# Handover 2026-01-24 (fix: Import Candidates が1件になる)

## 背景
- ASAGAKE の `Import Candidates` 実行で「1銘柄/1プランしか取り込まれない」事象が散発。
- `logs/vba_events.log` には `ImportCandidatesV2 done imported=1` が記録されていた。
- `output/excel/candidates_nextday.csv.backup_*` を確認すると、実際に **rows=1** や **ヘッダのみ（rows=0相当）** のバックアップが存在。

結論として、Import 側の不具合というより **Import 元の `candidates_nextday.csv` が極端に小さい内容で上書きされる日がある** のが主因。

## 対応（再発防止）
### 1) nightly 側の上書きガード強化
- `scripts/run_nightly_candidates.ps1` で `tools/aggregate_candidates_today.py` 呼び出し時に `--fallback-min-rows 10` を付与。
- `tools/aggregate_candidates_today.py` のデフォルト `FALLBACK_MIN_ROWS_DEFAULT` を **5 → 10** に引き上げ。
  - 上流の出力欠損/部分書き込み等で候補が「小さすぎる」場合に、`candidates_nextday_last_good.csv` を維持して事故を避ける目的。

### 2) Excel Import 側の last_good フォールバック追加
- `excel/AutoTraderAdvanced.bas` の `ImportCandidatesV2` にて、
  - `candidates_nextday.csv` が小さすぎる場合、`candidates_nextday_last_good.csv` を読み直して Import を継続。
  - last_good でも不足なら Import を中断（既存のダッシュボードを消さない）。

### 3) 朝8:50タスク側の安全策
- `Step1_Morning.vbs` の最小行数を **5 → 10** に引き上げ。
- `candidates_nextday.csv` が小さすぎる場合、`candidates_nextday_last_good.csv` から復元してから Import を試みる。
- さらに重要な安全策として、**SYSTEM/非対話セッションでは実行を拒否**（見えないExcelが起動してファイルをロックする事故を防ぐ）。

## 期待する効果
- 「nightly が失敗/部分出力 → candidates_nextday が 0〜1行に縮む → 翌朝 Import が 1件」事故の頻度を大きく減らす。
- 朝タスクでExcelが見えないのにASAGAKEだけロックする事故（他ユーザー使用中）を減らす。

## 次の確認（運用）
- `output/excel/candidates_nextday.csv` の行数・更新時刻を daily で確認。
- `logs/vba_events.log` の `ImportCandidatesV2` が `fallback_last_good_used` を出していないかを確認（出ていれば “元CSVが小さかった日” のサイン）。

