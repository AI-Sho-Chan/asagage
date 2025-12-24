# Bridge.v1（Excel/VBAを薄くするための入出力CSV）実装計画（DEMOは最小実装済み）

この文書は「今動いているASAGAKE.xlsm（Import Candidates / Ordersシート）」を壊さずに、あとから **ファイルI/Oの形だけ** 追加するための計画です。
2025-12-24 時点で、`excel/AutoTraderAdvanced.bas` に **DEMO向けの最小Bridge実装**（MarketSnapshotの追記、OrdersCmdのポーリング、ExecutionEventsの追記）を追加しました。LIVE側は次工程です。

## 1. ディレクトリ（既存のoutput配下に追加）
- Excelが読む: `output/excel/inbox/`
- Excelが書く: `output/excel/outbox/`

## 2. ファイル3本（最小）

### 2.1 Excel → Python（追記）: MarketSnapshot.v1
- パス: `output/excel/outbox/market_snapshots_YYYYMMDD.csv`
- 目的: RSSで実際に見えている値をそのまま保存（後から監査できる）
- ルール: 1行=1スナップショット、追記のみ（append-only）
- 文字コード: UTF-8 with BOM（Excelで文字化けしにくい）

**VBA疑似コード**
1) タイマー（例: 1〜5秒）で定期実行
2) 対象行（候補行）をループし、ticker/Last/Bid/Ask/VWAP/出来高などを読む
3) `output/excel/outbox/market_snapshots_YYYYMMDD.csv` に1行追記
4) 取れない場合は `data_quality=MISSING` で残す

### 2.2 Python → Excel（丸ごと置換）: OrdersCmd.v1
- パス: `output/excel/inbox/orders_cmd_YYYYMMDD.csv`
- 目的: Python側の判断結果（PLACE/MODIFY/CANCEL）をExcelに渡す
- ルール: **原子更新**（tmpに書いてからリネーム/置換）
- Excel側の読み方: `cmd_seq` が前回処理済みより大きい行だけ実行

**VBA疑似コード**
1) `orders_cmd_YYYYMMDD.csv` が存在すれば読み込む
2) シートの隠しセル（例: `state/adapter_state.json` でも可）に保存している `last_cmd_seq` を読む
3) `cmd_seq > last_cmd_seq` の行だけ処理する
4) PLACE/MODIFY/CANCEL を RSS注文関数（または既存ロジック）に流す
5) 成功/失敗を ExecEvents.v1 に追記
6) `last_cmd_seq` を更新

### 2.3 Excel → Python（追記）: ExecEvents.v1
- パス: `output/excel/outbox/execution_events_YYYYMMDD.csv`
- 目的: Excelが「何を送ったか」「結果どうなったか」を機械的に回収
- ルール: 追記のみ（append-only）

**VBA疑似コード**
1) PLACE/MODIFY/CANCEL を送った直後に `SENT` を1行追記
2) 受付/拒否が分かるなら `ACK` / `REJECT` を追記
3) 約定が取れるなら `FILL` / `PARTIAL_FILL` を追記
4) キャンセル等も `CANCELLED` / `EXPIRED` を追記

## 3. Ordersシートは「人が見るビュー」
今回の方針は「Ordersシートを壊さず、右側に列を追加してJOINできる状態にする」です。
（追加のみ。既存列の削除/並べ替えはしない）

**追加推奨列（右側に）**
- `run_id`
- `candidate_id`
- `decision_id`
- `client_order_id`
- `cmd_seq`
- `broker_order_id`
- `exec_event`
- `fill_qty`
- `fill_price`
- `deny_reasons`
- `engine_version`

## 4. ポーリング間隔の考え方（簡単）
- 最初は「5秒に1回」で十分（壊れにくい）
- 将来、本番に寄せるなら「1秒」または「イベント駆動」に近づける（ただしRSSの更新頻度が限界）

## 5. 次工程（VBA実装に入る前に決めること）
- どの列（ticker/価格/VWAP等）をMarketSnapshotに出すか（RSSで取得可能な範囲）
- 注文の責務分割（ExcelがTP/SL/Trailまで持つのか、Pythonが持つのか）
- `run_id` の採番規則（PC名・LIVE/DEMO・連番）
