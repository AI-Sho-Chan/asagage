from __future__ import annotations
from typing import List, Dict, Tuple
from .base import BaseProvider

# NOTE: kabu API 制約（厳守）
# - 情報系リクエストは ~10 req/sec 上限
# - WebSocket PUSH は「銘柄登録」済みの最大 50 銘柄が配信対象
# - 寄り前は最良気配が欠測（特別気配等）し得る → 欠測は無効サンプルとして捌く
# 実装方針（API開通後に実装）:
# 1) ランキングAPIで ~200 銘柄に一次抽出（売買代金/出来高など）
# 2) 50銘柄×バッチで /unregister/all → /register(≤50) → WebSocket受信(~60秒) → 次バッチ
# 3) 8:59:30 までに全バッチ走査 → AOI 時系列→選抜

class KabuProvider(BaseProvider):
    def __init__(self, host: str = "http://localhost:18080", token: str | None = None):
        self.host = host
        self.token = token  # TODO: .env から読む

    def get_liquidity_universe(self, max_n: int = 200) -> List[str]:
        raise NotImplementedError  # ランキングAPIで実装

    def collect_aoi_input(self, symbols: List[str]) -> Dict[str, Tuple[list[float], list[float]]]:
        raise NotImplementedError  # バッチ登録 + WebSocket PUSH 実装
