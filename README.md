# Asagake (Morning Rush)
寄り前 AOI × 寄り直後 AVWAP/ATR スキャルの実運用支援ツール。

## 使い方（Mock/TDD）
- venv / 依存セットアップ
- `python scripts/run_screener.py --provider mock` で `out/watchlist.txt` を生成
- Excel コックピットを開き `Settings!B2` に watchlist パスを設定 → 自動読込 → 監視

## kabu API（開通後）
- ランキングAPIで一次抽出 → 50銘柄×バッチで登録/PUSH収集 → AOI → watchlist 出力
