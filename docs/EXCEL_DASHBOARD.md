# Excel コックピット構築手順（MarketSpeed II RSS 前提）

## 1) 準備
- MarketSpeed II にログインし、Excel の「楽天 RSS」アドインを有効化
- 本ブックはマクロ有効形式（.xlsm）で保存

## 2) シート構成
- Settings: B2に watchlist.txt の絶対パス、B3に乖離しきい値(初期0.6)
- Dashboard: A2:A11 銘柄コード（自動読込）、C以降に単項目
  - C: =RssMarket($A2,"現在値")
  - D: =RssMarket($A2,"始値")
  - E: =RssMarket($A2,"高値")
  - F: =RssMarket($A2,"安値")
  - G: =RssMarket($A2,"出来高")
  - H: AVWAP(後述), I: ATR(5), J: 乖離=(C-H)/I
- Bars(非表示可): 各銘柄行に 1分足を RssChart スピル
  - 例: Bars!B2 に =RssChart(1, Dashboard!$A2, "1M", 20)

## 3) AVWAP/ATR(5) の数式（例）
- AVWAP(H2):
  =LET(blk, Bars!B3:K22, d, INDEX(blk,,4), c, INDEX(blk,,9), v, INDEX(blk,,10),
       rows, FILTER(blk, d=TODAY()),
       IF(SUM(INDEX(rows,,10))>0, SUM(INDEX(rows,,9)*INDEX(rows,,10))/SUM(INDEX(rows,,10)), ""))

- ATR(5)(I2):
  =LET(blk, Bars!B3:K22, rows, FILTER(blk, INDEX(blk,,4)=TODAY()),
       Hh, INDEX(rows,,7), Ll, INDEX(rows,,8), Cl, INDEX(rows,,9),
       Prev, VSTACK(NA(), TAKE(Cl, ROWS(Cl)-1)),
       TR, BYROW(SEQUENCE(ROWS(Cl)), LAMBDA(r, MAX(Hh[r]-Ll[r], ABS(Hh[r]-Prev[r]), ABS(Ll[r]-Prev[r])))),
       IF(ROWS(TR)>=5, AVERAGE(TAKE(TR,5*-1)),""))

- 乖離(J2): =IFERROR( (C2 - H2) / I2 , "")

## 4) 条件付き書式（Dashboard!J2:J11）
- ルール1: =J2>=Settings!$B$3 → 赤系
- ルール2: =J2<=-Settings!$B$3 → 緑系

## 5) 運用
- ブックを開くと Auto_Open で watchlist.txt を読み込み
- MarketSpeed II を起動し続けること（RSS更新のため）
