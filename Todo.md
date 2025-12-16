# Todo

## Done
- [x] Excel 取り込みの正は `output/excel/candidates_nextday.csv`
- [x] 同一銘柄でも「強いプランは複数採用」を許容（候補CSVは1銘柄1プランに縮約しない）
- [x] DailyReplay（取引後の疑似売買）を追加し、メール送信できる状態にした
- [x] 週末VMが金曜16:30に走らなかった原因を特定（`.sh`の改行がWindows形式になり実行失敗）
- [x] 取引中に邪魔になるスナップショット系タスクは停止（必要なら復活できるようにする）

## Next（優先）
- [ ] Windowsの `Asagake-DailyReplay` を「平日18:00」「ログオンしてなくても動く」設定にする（PCスリープ/未ログインで止まるのを防ぐ）
- [ ] 週末VM側は `~/asagage/vm_run_weekend_only.sh` を正とし、Windowsから `.sh` をscp上書きしない運用に固定
- [ ] Top200常連の1分足データを毎日育てる（`scripts/run_update_regulars_1m.ps1`）

## Later（検証・改善）
- [ ] 時間帯を区切る/区切らない（M0〜M3）を、同じ前提で20営業日ぶん集計して比較（前提ルールは `analysis/method_comparison_schema.md`）
- [ ] GapBanPct / NoTradeMin の適正値をDailyReplayで検証（実運用へ反映するかは別判断）

