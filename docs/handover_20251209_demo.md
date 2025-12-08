# Handover 2025-12-09 (Demo preplace & nightly準備)

## 1. Demoモードの preplace ロギングを強化
- 変更ファイル: `excel/AutoTraderAdvanced.bas`
- `RefreshTrendsV2` 内で `RunStatusV2=DEMO_RUNNING` の場合、`MarkPendingPreplaceAsOrderedDemo` を呼び、Orders シートの PENDING 行を `mode=preplace_demo` / `status=ORDERED` に更新してログを残す。
- 判定ヘルパーとして `IsDemoMode` を追加。

## 2. BudgetFactor / Live-Demo分類の取り込み
- 変更ファイル: `tools/aggregate_candidates_today.py`, `analysis/summarize_nightly_candidates.py`
- CSVに `BudgetFactor_row` (2.0/1.0/0.5) と Live/Demo クラスを付与してダッシュボードに渡す。
- ダッシュボードのヘッダに `BudgetFactor_row` を追加済み。ImportCandidatesV2 で取り込める。

## 3. トレンドフィルタの運用方針メモ
- `AllowedSide` は基本 BOTH。
- `trend_allowed_policy` は基本 BOTH/空。順張りだけに絞るときのみ ALIGNED_ONLY を使う。
- `Bias_bp` は 1.0 以上に設定しないと BAN 多発。

## 4. ローカル nightly のエラー
- `logs/nightly_py_error_20251208_163001.log` で `ModuleNotFoundError: numpy`。ローカル nightly を再開するなら `pip install -r requirements.txt` または `pip install numpy` を実行。

## 5. 反映・確認TODO
- `AutoTraderAdvanced.bas` を Excel にインポートし直し、`SetupDashboardUIV2` を実行してヘッダ/ボタンを再生成。
- DEMO_RUNNING で `RefreshTrendsV2` 実行 → Orders に preplace_demo/ORDERED が残るか軽く確認。
- （必要なら）ローカル nightly 用に依存ライブラリをインストール。
