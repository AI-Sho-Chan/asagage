# 監査用インプット一式（VWAP回帰・逆張りスキャルピング / ASAGAKE）

この文書は「別のAI（第三者）に、ASAGAKEシステムを“丸ごと”監査してもらう」ための最低限〜網羅情報です。  
**専門用語はなるべく避け、必要な用語は最後に説明**します。

> 注意（重要）
> - 実運用の損益を保証するものではありません。
> - SMTP等のパスワードは **この文書には含めません**（`state/smtp.json` は監査に出す前に必ず伏せてください）。
> - データ取得元（Yahoo等）の仕様変更・取得制限の影響を受けます。

---

## 0. これを読めば何が分かるか（監査観点）

- このシステムが狙っている「儲け方の仮説」と、売買ルール（いつ入って、いつ出るか）
- Excel（楽天RSS）を使う理由と、動かし方（Live/Demo）
- 週末バッチ/ナイトバッチ/日次リプレイ（仮想売買）の役割分担
- バックテスト（週末バッチのWalk-Forward）が何を保証していて、何が弱点か
- 現時点の成績（十分ではない点も含めて）と、いま残っている課題

---

## 1. 投資アイディア（仮説）と、ワークすると考える根拠

### 1.1 仮説（何を狙うか）
- 株価は日中の短い時間では「行き過ぎ」やすく、行き過ぎた後に **VWAP（出来高加重平均価格）へ戻りやすい**場面がある。
- その「行き過ぎ」を数値化して（J値）、一定の水準を超えたら逆方向に入る（逆張り）。

### 1.2 なぜ勝てる可能性があるのか（想定ロジック）
- 板（気配）とVWAPを使い「行き過ぎ価格」に先回り指値を置くことで、約定後すぐ戻る局面を取りに行く。
- ただし、行き過ぎが“戻らない”場面（トレンドが強い日、ギャップが大きい日など）では負けやすいので、
  - 「朝一は触りにくい」（NoTradeMin）
  - 「ギャップが大きいときはやらない」（GapBan）
  - 「日経/TOPIXの方向と反対はやらない」（方向フィルタ）
  といった“守り”を入れる。

---

## 2. 売買ルール（具体）

### 2.1 主要指標（J）
- **ベースのJ**：`(価格 - VWAP) / ATR` を中心に、補正（バイアス補正・ギャップ補正等）を加えたもの。
- **J_th（しきい値）**：候補ごとに最適化された「どの程度の行き過ぎで仕掛けるか」。
  - Excel上では `J_th` が `BAN` の場合、その行は取引しない（実質的に無効化）。

### 2.2 エントリー（先回り指値）
- 仕掛けは「Jが閾値を超えたら成行」ではなく、**閾値に到達しそうな価格に先回りで指値**を置く。
- 先回り指値は、J_thやVWAP/ATRなどから動的に計算される。
- Demoの場合は実注文は出さず、Ordersシートにログとして残す。

### 2.3 決済（利確/損切り/トレーリング）
- 利確（TP）と損切り（SL）をあらかじめ置く（またはDemoログとして生成する）。
- 追加でトレーリング（Trail）を設定する場合がある。

### 2.4 取引しない条件（主なもの）
- **NoTradeMin**：寄り（09:00）から一定分は仕掛けない。
- **GapBanPct / GapBan**：前日からのギャップが大きいときに見送る。
- **方向フィルタ（市場トレンド同期）**：
  - 日経平均（NKY）やTOPIXの「上げ/下げ」判定と、仕掛け方向が一致する場合だけ許可する（=ALIGNED_ONLY）。

---

## 3. システム構成（楽天RSS利用の詳細）

### 3.1 全体像（ざっくり）
- **Windows PC**
  - Excel：`C:\AI\asagake\ASAGAKE.xlsm`
  - VBA：`excel/AutoTraderAdvanced.bas` を取り込んで動かす
  - 取引後の仮想売買（DailyReplay）：18:00に自動実行してメール送信
- **週末VM（GCP）**
  - 週末バッチ（重い最適化）を金曜に自動実行
  - 結果をGCSへ同期（Windows側が回収）

### 3.2 Excel/VBA（実取引の土台）
- メインVBA：`excel/AutoTraderAdvanced.bas`
- Live/Demoの区別：
  - `LIVE_RUNNING`：実注文（楽天RSS注文）
  - `DEMO_RUNNING`：Ordersシートにログのみ（実注文なし）
- 重要：Demoは「Excel内で擬似約定/擬似決済」を行うため、実際の板更新頻度やチェック間隔の影響を受ける。

### 3.3 VBA反映（更新手順）
- 更新スクリプト：`scripts/update_asagake_vba.ps1`
  - `excel/AutoTraderAdvanced.bas` の内容を `ASAGAKE.xlsm` に差し替える。
  - 実行時にバックアップ（`ASAGAKE_backup_YYYYMMDD_HHMMSS.xlsm`）を作る。

---

## 4. 候補（銘柄・時間帯・パラメータ）の選定方法

### 4.1 候補CSV（Excelが読むファイル）
- Excelが取り込むのは **`output/excel/candidates_nextday.csv`**。
- 例：列（主要なもの）
  - `Ticker`（銘柄）
  - `session`（時間帯。例：AM0930）
  - `SignalMode`（j-only / j-cross）
  - `J_th`, `TPk`, `SLk`, `TMAX`, `ATR_n`
  - `BudgetFactor_row`（2.0/1.0/0.5）
  - `live_demo_class`（LIVE_STRONG / LIVE_BASE / DEMO_ONLY）
  - `NoTradeMin`, `GapBanPct`
  - `trend_driver`, `trend_window`, `trend_bp_th`, `trend_allowed_policy`

### 4.2 週末バッチ（重い最適化 / Walk-Forward）
- 目的：市場の流動性上位銘柄（Top300等）から「勝てそうな銘柄×時間帯×パラメータ」を探す。
- 実体：VMで `scripts/nightly_build_candidates.py` が `scripts/bt_opt30_forward.py` を多数回呼ぶ。
- 週末VM起動スクリプト：`vm_run_weekend_only.sh`
  - Universe作成（TopVol）
  - 1分足キャッシュ更新
  - coarse（粗探索）→ refine（詳細探索）
  - `output/excel/NIGHTLY_YYYYMMDD/**/candidates_*.csv` を生成
  - `tools/aggregate_candidates_today.py` で `output/excel/candidates_nextday.csv` にまとめる

### 4.3 時間帯を区切る理由（現状の設計）
- 9:00直後・前場・後場で値動きが違うため、「同じ銘柄でも効く設定が違う」前提で時間帯ごとに探索している。
- 監査ポイント：
  - 区切りすぎると“当たり外れ”が増える（同じ銘柄でも毎週結論が揺れやすい）。
  - 区切らないと平均化され、短い時間帯で効く設定を取り逃がす可能性がある。

---

## 5. ナイトバッチ / 日次リプレイ（仮想売買）の役割

### 5.1 ナイトバッチ（候補生成の平日版）
- 目的：平日の夕方に軽めに候補を更新（週末ほど重くない）。
- 実体：Windows側で `scripts/run_nightly_candidates.ps1`（環境依存があり、停止/不安定な時期がある）。

### 5.2 日次リプレイ（DailyReplay：18:00の仮想売買）
- 目的：当日の候補で「もし売買していたらどうだったか」を、Yahooの1分足で再現して日次で見える化する。
- 実体：
  - 実行：`scripts/run_daily_replay.ps1`
  - 中身：`tools/simulate_daily_replay.py`
  - 出力：
    - `analysis/daily_trades_YYYYMMDD.csv`（トレード明細）
    - `analysis/daily_replay_YYYYMMDD_mail.txt`（日本語要約）
  - Windowsタスク：`Asagake-DailyReplay`（平日18:00）

---

## 6. バックテスト（週末バッチ）の妥当性と弱点

### 6.1 強み（何が良いか）
- 過去の一定期間を「学習（train）」と「検証（forward）」に分け、未来を見ない形に近づけている（Walk-Forward）。
- 取引回数が少なすぎるものを弾くなど、最低限の足切りがある。

### 6.2 弱点（監査で必ず指摘される点）
- 探索する組合せが多く、偶然当たっただけの“偽の勝ち設定”が混じる可能性（過剰最適化）。
- `forward_pf_eff=999` のように「損失0」や「回数が少ない」時に数字が跳ねる場合があるため、見かけの数値に注意が必要。
- Yahooの1分足取得制限により、履歴が欠けやすい（=検証が不安定になる）。

---

## 7. 現在のステータス（2025-12-23時点）

- パイプライン（週末VM → candidates_nextday → Excel import → DEMOログ → 18:00 DailyReplay → メール）は一通り動く状態。
- ただし「結論として儲かる」と言い切れる段階ではない（次章参照）。
- ABテスト（coarse回数を減らす案）：B（coarse=3/refine=5）をVMで別出力で実行中（完走待ち）。

---

## 8. これまでのパフォーマンス（不十分である点も明記）

### 8.1 DailyReplayの実績（注意：まだ日数が少ない）
- `analysis/daily_trades_YYYYMMDD.csv`（サフィックス無し）のみを集計すると：
  - 対象日数：17日（2025-11-17〜2025-12-23）
  - 総トレード数：230
  - 合計損益：+227,172円
  - 1日平均：+13,363円
  - 勝率（単純平均）：0.583
  - 最悪日：2025-12-03（-411,652円）
  - 最良日：2025-11-17（+643,970円）
- 重要：日数が少なく、ブレが大きい。現時点では「儲かる/儲からない」を断定できない。

### 8.2 ExcelのDEMOログとの関係
- ExcelのDEMOは「一定間隔チェック」等の仕様があり、DailyReplay（1分足）と完全一致しないことがある。
- 監査では「どちらが“真実に近い”のか」を明確にする必要がある（現状はDailyReplayを主軸に改善中）。

---

## 9. 現在の課題（重要な順）

1. **候補が週ごとに揺れやすい**（同じ銘柄でも時間帯/パラメータが変わる）
2. **週末バッチの所要時間が長い**（VMコスト/運用負荷）
3. **Yahoo 1分足の欠損**（履歴不足でテストが不安定）
4. **Excel DEMOの約定判定が粗い**（チェック間隔/板更新頻度の制約）
5. **本番（LIVE）移行前の確認がまだ不足**
   - 発注・訂正・キャンセル・約定・決済の流れが「事故なく」回るか

---

## 10. 監査担当（別AI）への依頼事項（チェックリスト）

### 10.1 ロジック監査
- 逆張りが“戻らない日”に負ける構造になっていないか（守り条件で抑えられているか）
- trendフィルタ（ALIGNED_ONLY）の整合性：実装が意図どおりか
- TP/SL/Trail/GAPBAN/NoTradeMin の実運用とテスト（週末・DailyReplay・Excel）の整合性

### 10.2 バックテスト監査
- Walk-Forwardの切り方（train/forward）と、採用基準の妥当性
- 探索数に対する偽陽性（偶然当たり）のリスク評価
- サンプル数不足（forward_tradesが少ない）の扱い

### 10.3 運用監査
- 週末VM自動実行の確実性（スケジュール/ログ/停止）
- 毎日18:00のDailyReplayメールが落ちない構成（リトライ/ログ）
- “監査に必要な証跡”（ログ、CSV、設定）が残るか

---

## 11. 重要ファイル一覧（監査に必須）

### Excel/VBA
- `C:\AI\asagake\ASAGAKE.xlsm`（本体）
- `excel/AutoTraderAdvanced.bas`（VBAソース）
- `scripts/update_asagake_vba.ps1`（VBA取り込み）
- `state/strategy_rules.ini`（運用パラメータ）

### 候補生成（週末/ナイト）
- `scripts/nightly_build_candidates.py`
- `scripts/bt_opt30_forward.py`
- `tools/aggregate_candidates_today.py`
- `vm_run_weekend_only.sh`（VMの週末バッチ）
- VMログ：`~/cloud_logs/weekend_YYYYMMDD_*.log`

### 日次仮想売買（18:00）
- `tools/simulate_daily_replay.py`
- `scripts/run_daily_replay.ps1`
- `analysis/daily_trades_YYYYMMDD.csv`
- `analysis/daily_replay_YYYYMMDD_mail.txt`

### 運用ログ
- `logs/vba_events.log`（Excel側イベント/エラー）
- `docs/handover_*.md`（変更履歴）
- `Todo.md`（今後の作業）

---

## 12. 用語（最小）

- **VWAP**：1日中の平均価格（出来高で重みづけした平均）。
- **ATR**：値動きの大きさ（ざっくり「普段どれくらい動くか」）。
- **J**：VWAPからの「行き過ぎ度合い」を表す指標（単位なし）。
- **Walk-Forward**：過去を「学習期間」と「検証期間」に分けて、未来を見ない形に近づける検証方法。
- **coarse/refine**：粗く探す→良さそうな候補だけ細かく詰める二段階探索。

