# Asagake ダッシュボードの見方と使い方（要点）

- **Settings!B2**: watchlist.txt の絶対パス、**B4**: しきい値（例 0.6）
- **Dashboard**:
  - A列: 銘柄コード（watchlistから自動読込）
  - B列: 銘柄名（RSSの銘柄名称）
  - H列: **AVWAP（当日始値アンカーのVWAP）**
  - I列: **ATR(5)**（1分足、期間5）
  - J列: **乖離(ATR単位)** = (現在値-AVWAP)/ATR
  - J≥+しきい値 ⇒ **ショート候補**、J≤-しきい値 ⇒ **ロング候補**
  - K/L: 利確/損切りの目安（AVWAP到達／ATR×係数）
- **Bars**: 20ブロックの1分足。各ブロックの A2 が RssChart トリガー。E列=日付、F=時刻、G=始値～K=出来高。
- **アラート&自動記録**:
  - StartAlerts : 2秒ごとスキャンで J の色付け＋音
  - StartRecording : Bars の**最新1本だけ**を CSV に追記（1分毎）
  - CSV: excel\data\intraday\YYYY-MM-DD\{コード}.csv（ヘッダー: date,time,open,high,low,close,volume）
- **紙上トレード**（デモ自動売買）:
  - しきい値越えでエントリー、**目標=AVWAP**, **損切=ATR×係数**、**タイムストップ=9:15**
  - ログ: excel\trades\YYYY-MM-DD.csv（entry/exit/方向/損益/手数料/スリッページ含む）
