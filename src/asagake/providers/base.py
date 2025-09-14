from __future__ import annotations
from typing import List, Dict, Tuple

class BaseProvider:
    def get_liquidity_universe(self, max_n: int = 200) -> List[str]:
        """ランキング等で 200 銘柄程度に一次抽出"""
        raise NotImplementedError

    def collect_aoi_input(self, symbols: List[str]) -> Dict[str, Tuple[list[float], list[float]]]:
        """{symbol: ([bid_qty_series],[ask_qty_series])} を 8:55-9:00 相当で収集"""
        raise NotImplementedError
