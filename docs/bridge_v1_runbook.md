# Bridge v1 Runbook（DEMO）

この文書は、Bridge v1 を使って **「Pythonが orders_cmd を出す → Excel(DEMO)が消化する」** を安定運用するための手順書です。  
既存の `output/excel/candidates_nextday.csv` と Orders シート運用は壊さず、追加機能として動きます。

## 1. 使うファイル（固定）

- Excelが読む（inbox）
  - `output/excel/inbox/orders_cmd_YYYYMMDD.csv`（OC.v1 / atomic置換）
- Excelが書く（outbox）
  - `output/excel/outbox/market_snapshots_YYYYMMDD.csv`（MS.v1 / 追記のみ）
  - `output/excel/outbox/execution_events_YYYYMMDD.csv`（EE.v1 / 追記のみ）
- Pythonが書く（分析）
  - `analysis/bridge_health_YYYYMMDD.csv` / `.txt`
  - `analysis/decision_trace_YYYYMMDD.csv`（DT.v1）

## 2. DEMO疎通（最小の手順）

### 2.1 事前準備（候補CSV）
- `output/excel/candidates_nextday.csv` が存在することを確認（空だとダッシュボードに候補が入りません）。
- `tools/aggregate_candidates_today.py` は、候補CSVの右端に `candidate_id` 等の列を追加して出力します（既存列は変更しません）。

### 2.2 Excel側（DEMO開始）
1) `ASAGAKE.xlsm` を開く  
2) `DEMO Start`（または `RunStatusV2=DEMO_RUNNING`）にする  
3) 数秒待つ（AutoTickV2 が動いている状態）

期待される挙動（DEMO中）:
- `output/excel/outbox/market_snapshots_YYYYMMDD.csv` が増えていく
- `output/excel/outbox/execution_events_YYYYMMDD.csv` が増えていく（orders_cmdを処理した時）

### 2.3 Python側（orders_cmd を1件だけ投げる）
PowerShell（`C:\AI\asagake`）で:
- `python tools/bridge_smoketest_orders.py --date YYYYMMDD --run-id <ExcelRunId> --ticker 7203 --side BUY --qty 100 --limit-price 1000`

期待される挙動:
- Excelが `orders_cmd_YYYYMMDD.csv` を読み、Ordersシートに `mode=bridge_cmd` の行が1行増える
- `execution_events_YYYYMMDD.csv` に `ACK`（または `REJECT`）が追記される

## 3. Healthcheck（照合）

PowerShell（`C:\AI\asagake`）で:
- `python tools/bridge_healthcheck.py --date YYYYMMDD --base-dir C:\\AI\\asagake`

結果:
- `analysis/bridge_health_YYYYMMDD.txt`：人間向けサマリ
- `analysis/bridge_health_YYYYMMDD.csv`：機械向けメトリクス

終了コード:
- `0`：致命的な問題なし
- `2`：致命的（ファイル欠損 / cmd_seq重複 / ACK重複など）

## 4. よくあるトラブルと対処

### 4.1 outbox が増えない
- Excelが `DEMO_RUNNING` になっているか確認
- `AutoTickV2` が停止していないか確認（VBAエラーで止まることがあります）

### 4.2 orders_cmd を置いたのに Orders に反映されない
- `output/excel/inbox/orders_cmd_YYYYMMDD.csv` の日付が今日と一致しているか
- `cmd_seq` が前回より大きいか（Excelは `last_cmd_seq` 以下を無視します）

### 4.3 Excel再起動後に同じcmdが再実行される（cmd_seq巻き戻り）
- 現在は **WorkbookのName（非表示）** に `last_cmd_seq` を保存します
- Healthcheck の `ack_like_duplicate_cmd_seq` が増える場合は巻き戻り疑い

Nameの保存場所:
- `BridgeLastCmdSeqV1_YYYYMMDD`（非表示 / ブック内）
- `BridgeEventSeqV1_YYYYMMDD`（非表示 / ブック内）
- `BridgeRunIdV1_YYYYMMDD`（非表示 / ブック内）

## 5. 1日1回の運用（推奨）

1) 取引時間中（DEMO）: Excelで outbox が生成される
2) 夕方: `bridge_healthcheck` を実行して異常がないか確認
3) 異常があれば `analysis/bridge_health_YYYYMMDD.txt` を見て原因を切り分ける

