# Todo

## Done
- [x] `output/excel/candidates_nextday.csv` を Excel（ASAGAKE.xlsm）に取り込める状態にした（同一銘柄の複数プランも許容）。
- [x] 候補CSVに `BudgetFactor_row` / `live_demo_class` / `NKY_AllowedSide` / `TOPIX_AllowedSide` / `GapBanPct` / `NoTradeMin` を含めるように整備。
- [x] DailyReplay（取引後の仮想売買）を平日 18:00 に自動実行し、メール送信する仕組みを運用に乗せた（Windows側）。
- [x] DailyReplayの取引ログに `exit_reason`（利確/損切り/引け）などを出し、メール本文にも内訳と上位負けトレードの「負け方」を出すように改善。
- [x] 週末VMの週末バッチを「金曜 16:30 JST に自動起動 → 完了後に自動停止」へ安定化（`docs/handover_20251216.md`）。
- [x] 週末バッチを差分更新（新規＋異常＋月次リセットのみフル探索）できるようにした（`docs/handover_20251220.md`）。
- [x] 「Top200に一度入った銘柄」を永久保存対象にし、1分足データを毎日育てる仕組みを実装（`tools/build_top200_ever_universe.py`, `scripts/run_update_regulars_1m.ps1`）。
- [x] 週末VMで Optuna を利用可能にし、`--enable-bayes` が実際に効くようにした（refineの高速化目的）。
- [x] 候補が0件のときに `candidates_nextday.csv` が空で上書きされないように保護（前回の候補を保持し、診断ログを出す）。

## Next（優先）
- [x] 週末差分更新の前提ファイル `output/excel/candidates_for_YYYYMMDD.csv` が空にならないようにする（`scripts/nightly_build_candidates.py` のスナップショット生成位置の不具合を修正済み）
- [x] 2025/12/19 の `candidates_for_20251219.csv` がヘッダのみ問題を復旧（VMで `candidates_nextday.csv` をコピーして暫定復旧）
- [x] 取引時間中に邪魔になるため、5分スナップショット/1分足更新系のタスクは一旦停止（必要になれば後で復活）
- [ ] 次の金曜に週末バッチが自動起動・完走したかを確認（ログの見方を固定化）
- [ ] `abnormal_codes_latest.csv` のアップロードが失敗していないか、DailyReplayログで確認（失敗時はVMの異常時フル探索が効かない）

## Later（改善）
- [ ] `scripts/run_update_regulars_1m.ps1`（Top200常連の1分足更新）を再開する場合は、タスクの実行時刻/負荷/アップロード先を再点検してから有効化する
- [ ] VM が `gs://asagage-weekend-output/yahoo_1m_regulars` を取り込めているか確認（`vm_run_weekend_only.sh` の rsync ログ）
- [ ] GapBanPct / NoTradeMin の適正値を、DailyReplay の複数日（例: 20営業日）で検証して提案する（実装は別途判断）。
- [ ] 時間帯の切り方（区切る/区切らない/粗い区切り）を“同一条件”で比較する（比較表を `analysis/` に出力）。
- [ ] ABテスト（A: coarse=5/refine=5 vs B: coarse=3/refine=5）を20営業日で比較し、差が出るかを判断する（B候補は週末VMで「別出力」にして現行を汚さない）。
- [ ] Live/Demo 判定や BudgetFactor の基準を、実績（DailyReplay/DEMO）に合わせて自動提案できるようにする。
- [ ] Orders（負け/勝ちの原因）を定期的に棚卸しして、改善候補（除外銘柄・守り条件）を提案できる形にする（例: 5801.Tは“戻らない系”としてDEMO_ONLY/予算減、8105.Tは“急変で即死しやすい系”としてGapBan/NoTradeMin調整対象）。
