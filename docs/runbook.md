# Asagake RSS Runbook

## 1. 足種・セッション日
- 足種: `Settings!B4` で指定（例: `1M`）。
- セッション日: RssChart の出力列(E/F…)に記録。休場日検証が必要なら `Settings!B5` に任意日を入れて式で参照。

## 2. Bars の再構成と #VALUE! 復旧
1) Alt+F8 → `RebuildBarsAll` 実行  
2) Alt+F8 → `FixCalcAndRefresh` 実行  
3) Alt+F8 → `NudgeRssTriggers` 実行  
4) Ctrl+Alt+F9（全再計算）

## 3. A2（トリガー式）の基本
- A2 の式は 12 列単位のブロックで銘柄を切り替える。A2 を M2→Y2→AK2…へ **12 列おき** に横コピー。
- A2 サンプル（Dashboard!A 列の銘柄コードを順次参照）:
  =LET(
    blk, QUOTIENT(COLUMN()-COLUMN($A$2),12),
    code, INDEX(Dashboard!$A:$A, 2+blk),
    RssChart($B$2:$K$2, IFERROR(TEXT(code,"0"), code&""), Settings!$B$4, 20)
  )

## 4. 典型トラブル → 対処
- ヘッダーだけ英語でデータ空: 2)の復旧手順を実施。
- “スマート引用符” 混入でコンパイルエラー: 文字化け行を削除し、ASCII の " ' のみで貼り直し。必要なら PowerShell で .bas を生成してインポート。
- すべて同じ銘柄になる: A2 の LET 中 `blk, QUOTIENT(COLUMN()-COLUMN($A$2),12)` が入っているか確認。

## 5. マクロ起動（監視と紙トレ）
- Alt+F8 → `StartWatchlistAutoReload`（5秒ごとに watchlist.txt 監視）
- Alt+F8 → `RepairAndRebuild`（計算復旧＋Bars再構成＋入れ直し）
- Alt+F8 → `StartPaperTrader`（紙トレ自動化）
- 停止: Alt+F8 → `StopPaperTrader`

## 6. ログ/成果物
- 取引ログ: `excel\trades\YYYY-MM-DD.csv`
- イベントログ: `excel\logs\YYYY-MM-DD.log`
