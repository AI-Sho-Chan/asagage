# Bridge v1 / DecisionTrace v1 実装検証（監査用エビデンス）

作成日: 2025-12-24  
対象リポジトリ: `C:\AI\asagake`  
対象ブランチ: `master`  

この文書は、DT.v1 / Bridge v1 が「実装できた」と言ってよいかを判定するための **YES/NO＋根拠** を、再現可能な形で固定したものです。

---

## 0. 対象ファイル一覧

**スキーマ（YAMLだが中身はJSON＝YAML互換）**
- `schemas/decision_trace_dt_v1.yaml`
- `schemas/market_snapshot_ms_v1.yaml`
- `schemas/orders_cmd_oc_v1.yaml`
- `schemas/execution_events_ee_v1.yaml`

**Python I/O**
- `src/asagake_io/csv_schemas.py`
- `src/asagake_io/csv_writer.py`
- `src/asagake_io/atomic_writer.py`
- `src/asagake_io/validator.py`

**DailyReplay（DT追記対応済み）**
- `tools/simulate_daily_replay.py`

**テスト**
- `tests/test_bridge_v1_schemas.py`

**VBA実装計画（仕様書のみ・実装は次工程）**
- `docs/bridge_v1_vba_plan.md`

---

## A. スキーマ（YAML）整合性

### A1) schema_version の固定値明記（DT/MS/OC/EE）
判定: **YES**

根拠（各yaml先頭20行の抜粋）

`schemas/decision_trace_dt_v1.yaml`:
- 2行目に `"schema_version": "DT.v1"` が存在
- 3〜6行目に `primary_key: ["run_id","event_seq"]` が存在

`schemas/market_snapshot_ms_v1.yaml`:
- 2行目に `"schema_version": "MS.v1"` が存在
- 3〜7行目に `primary_key: ["run_id","snap_ts","ticker"]` が存在

`schemas/orders_cmd_oc_v1.yaml`:
- 2行目に `"schema_version": "OC.v1"` が存在
- 3〜6行目に `primary_key: ["run_id","cmd_seq"]` が存在

`schemas/execution_events_ee_v1.yaml`:
- 2行目に `"schema_version": "EE.v1"` が存在
- 3〜6行目に `primary_key: ["run_id","event_seq"]` が存在

### A2) enum定義（DT event_type / OC action/side/order_type）
判定: **YES**

根拠（該当箇所の行番号は `.yaml` 内に含まれる番号）

- DT event_type: `schemas/decision_trace_dt_v1.yaml:0068-0076`
- OC action: `schemas/orders_cmd_oc_v1.yaml:0033-0041`
- OC side: `schemas/orders_cmd_oc_v1.yaml:0050-0058`
- OC order_type: `schemas/orders_cmd_oc_v1.yaml:0066-0074`

### A3) primary_key（一意キー設計）の明記
判定: **YES**

根拠: A1の先頭20行に `primary_key` が明記されている。

### A4) 後方互換方針（列追加は右側のみ・既存列名不変）を明文化
判定: **YES**

根拠（明文化の所在）
- `docs/bridge_v1_vba_plan.md:50-53`
  - `:50` 「Ordersシートを壊さず、右側に列を追加…」
  - `:51` 「追加のみ。既存列の削除/並べ替えはしない」
  - `:53` 「追加推奨列（右側に）」

---

## B. CSV Writer / Atomic Writer / Validator

### B1) csv_writer.py がヘッダ二重書きを起こさない
判定: **YES**

根拠1（pytest）
- `pytest -q tests/test_bridge_v1_schemas.py`
  - 結果: `5 passed in 0.17s`
  - `tests/test_bridge_v1_schemas.py:37-81` が「2回appendしてもヘッダが1回のみ」を検証

根拠2（手動サンプル）
- `tmp/dt_writer_sample.csv` を writer で2回 append し、`header_count=1` を確認

### B2) writer のエンコーディングが UTF-8 with BOM（utf-8-sig）
判定: **YES（新規作成時は必ずBOM、追記時はBOM重複を避ける）**

根拠（実装）
- `src/asagake_io/csv_writer.py:23-37`
  - 新規: `encoding_new_file="utf-8-sig"`, `mode="w"`
  - 追記: `encoding_append="utf-8"`, `mode="a"`

根拠（ファイル先頭3バイト）
- `tmp/dt_writer_sample.csv` 先頭3バイトが `EF BB BF`（= 239,187,191）

### B3) atomic_writer.py が temp→os.replace の原子更新
判定: **YES**

根拠（実装）
- `src/asagake_io/atomic_writer.py:17-25`
  - `tmp` に書いてから `os.replace(tmp, path)`

追加の簡易検証（Windows）
- `tmp/orders_cmd_atomic_test.csv` を50回連続で更新し、毎回CSVとして読めることを確認
  - 結果: `atomic_update_ok: True`

### B4) validator の列チェック方針（必須列+追加列OKか）
判定: **追加列OK（allow_extra=True がデフォルト）**

根拠（実装）
- `src/asagake_io/validator.py:18-49`
  - `allow_extra: bool = True`
  - `allow_extra=False` のときだけ “Unexpected column” をエラーにする

### B5) 数値フォーマット（桁区切り無し・小数点は .）
判定: **YES**

根拠（出力例）
- `tmp/dt_writer_sample.csv` に `2810.5`, `1.0` 等
- `tmp/orders_cmd_sample.csv` に `2812.8`, `2809.2` 等

---

## C. DailyReplay（simulate_daily_replay.py）の DT 出力が監査に使える最低限レベルか

### C1) DTに最低限の event_type が出る
判定: **YES**

実行（例）
- `C:\Python313\python.exe tools\simulate_daily_replay.py --date 20251219 --label dt --decision-trace --run-id RUN_TEST_20251219 --engine-version EV_TEST`

根拠（event_type件数: run_id=RUN_TEST_20251219）
- `MARKET_SNAPSHOT: 18`
- `FEATURES_COMPUTED: 18`
- `FILTER_EVAL: 18`
- `DECISION: 8`
- `EXIT: 8`

### C2) event_seq が run 内で単調増加
判定: **YES**

根拠
- `event_seq_strictly_increasing: True` を確認

### C3) --run-id / --engine-version が全行に反映
判定: **YES**

根拠
- DT先頭3行で `run_id=RUN_TEST_20251219`, `engine_version=EV_TEST` が一致

### C4) candidate_id が埋まっているか（合否直結）
判定: **YES**

根拠
- `candidate_id_nonempty: 70 / 70`

### C5) daily_trades_YYYYMMDD.csv が DT追加後も壊れていない
判定: **YES**

根拠（DT無し vs DTありで一致）
- `analysis/daily_trades_20251219_nodt.csv` と `analysis/daily_trades_20251219_dt.csv` を比較
  - 行数: `7` vs `7`（一致）
  - `pnl_yen` 合計: `82399.63983261558` vs `82399.63983261558`（一致）

---

## D. テスト（最低限の品質担保）

### D1) pytest -q（全テスト）が通るか
判定: **NO（ローカルではタイムアウト）**

根拠
- `pytest -q` が 10分でタイムアウト（DT/Bridge以外の重いテストが含まれる可能性）

### D2) Bridge v1 の受け入れに必要なテストがあるか
判定: **部分的にYES**

根拠（通過したテスト）
- `pytest -q tests/test_bridge_v1_schemas.py` → `5 passed in 0.17s`

カバー範囲
- スキーマyamlとPython列順の一致
- append writer の「二重ヘッダ防止」＋ validator の基本チェック

未カバー（pytest内）
- `atomic_writer.py` の原子更新自体は pytest では未カバー  
  ※ただし手動で `atomic_update_ok: True` の簡易検証は実施

---

## 変更履歴（この検証に関連する追加修正）

- `tools/simulate_daily_replay.py` で tz-aware timestamp を含む場合に `minutes_from_open` 計算が落ちる問題を修正  
  - 修正内容: tzinfo を落として差分を計算（naive/aware subtraction回避）
  - コミット: `6a3a94dd91`

