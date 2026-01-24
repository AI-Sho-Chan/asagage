# Handover 2026-01-24（修正：Import Candidates が 1 件になる）

## 何が起きていたか
- ASAGAKE の `Import Candidates` を押すと、候補が 1 銘柄/1プランしか入らないことがあった。
- `logs/vba_events.log` に `ImportCandidatesV2 done imported=1` が残っていた。

## 原因の候補（複数パターン）
1) `output/excel/candidates_nextday.csv` が **小さすぎる**（生成途中/上書き途中/0行に近い）
2) Import 処理が **途中で止まっている**（エラーが握りつぶされている）
3) 朝の自動処理（タスクスケジューラ）で Excel を見えない形で起動していて、ファイル更新やロックが絡んでいる

## 対応（VBA側）
- `ImportCandidatesV2` に「小さすぎるCSVを読み込んだ時に、ダッシュボードを空にしない」保護を入れる。
- CSVが小さい場合は、最後に成功した候補（`candidates_nextday_last_good.csv`）へフォールバックして Import を続行する。
- Import 終了時に `approxRecords / parsedRecords / nonEmptyLines / size / mtime` を `vba_events.log` に必ず出す（切り分け用）。

## 対応（運用側：切り分け手順）
1) Import 直後に `logs/vba_events.log` を開く
2) `ImportCandidatesV2 done ...` の行を確認し、
   - imported が 1 なのか
   - parsedRecords が 1 なのか（CSVが小さい/壊れている）
   - parsedRecords は多いのに imported が 1 なのか（途中終了の疑い）
   を判断する

## 再発防止
- `tools/aggregate_candidates_today.py` が `candidates_nextday.csv` を生成する際、生成失敗で 0〜1行になりそうなら「前回のlast_goodを残す」運用にする。
- 朝の自動起動（タスクスケジューラ）が SYSTEM でExcelを立ち上げている場合、ユーザーからは見えず、かつファイルロックになるので要注意（ユーザー実行に寄せる）。

