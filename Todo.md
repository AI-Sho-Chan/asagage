# Todo

## Done
- [x] `output/excel/candidates_nextday.csv` を Excel（ASAGAKE.xlsm）に取り込める状態にした（同一銘柄の複数プランも許容）。
- [x] 候補CSVに `BudgetFactor_row` / `live_demo_class` / `NKY_AllowedSide` / `TOPIX_AllowedSide` / `GapBanPct` / `NoTradeMin` を含めるように整備。
- [x] DailyReplay（取引後の仮想売買）を平日 18:00 に自動実行し、メール送信する仕組みを運用に乗せた（Windows側）。
- [x] 週末VMの週末バッチを「金曜 16:30 JST に自動起動 → 完了後に自動停止」へ安定化（`docs/handover_20251216.md`）。
- [x] 週末バッチを差分更新（新規＋異常＋月次リセットのみフル探索）できるようにした（`docs/handover_20251220.md`）。

## Next（優先）
- [ ] Top200「常連」銘柄の 1分足データを“永久保存対象”として毎日育てる（`scripts/run_update_regulars_1m.ps1`）
- [ ] Top200常連の 1分足を VM 側でも活用できるように同期（Windows → GCS → VM）
- [ ] 次の金曜に週末バッチが自動起動・完走したかを確認（ログの見方を固定化）
- [ ] `abnormal_codes_latest.csv` のアップロードが失敗していないか、DailyReplayログで確認（失敗時はVMの異常時フル探索が効かない）

## Later（改善）
- [ ] GapBanPct / NoTradeMin の適正値を、DailyReplay の複数日（例: 20営業日）で検証して提案する（実装は別途判断）。
- [ ] 時間帯の切り方（区切る/区切らない/粗い区切り）を“同一条件”で比較する（比較表を `analysis/` に出力）。
- [ ] Live/Demo 判定や BudgetFactor の基準を、実績（DailyReplay/DEMO）に合わせて自動提案できるようにする。
