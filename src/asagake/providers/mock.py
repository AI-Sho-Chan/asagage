from __future__ import annotations
from typing import List, Dict, Tuple
import random
from .base import BaseProvider

class MockProvider(BaseProvider):
    def __init__(self, seed: int = 42):
        random.seed(seed)

    def get_liquidity_universe(self, max_n: int = 200) -> List[str]:
        # 実運用ではランキングAPIで抽出。ここでは固定の大型株サンプルで代用。
        base = ["7203","9984","6758","8058","6861","8035","9432","4502","6273","8766",
                "6902","6367","6098","6095","6920","4503","4063","4568","4507","2413"]
        return base[:max_n]

    def collect_aoi_input(self, symbols: List[str]) -> Dict[str, Tuple[list[float], list[float]]]:
        out: Dict[str, Tuple[list[float], list[float]]] = {}
        for s in symbols:
            bid, ask = [], []
            bias = random.uniform(-0.8, 0.8)  # 不均衡方向
            for _ in range(25):               # 約4分間で ~25サンプル想定
                base = random.uniform(5000, 9000)
                # 欠測の疑似（特別気配等）
                if random.random() < 0.05:
                    bid.append(None); ask.append(None); continue
                b = base * (1 + max(0, bias)) * random.uniform(0.9, 1.1)
                a = base * (1 - min(0, bias)) * random.uniform(0.9, 1.1)
                bid.append(b); ask.append(a)
            out[s] = (bid, ask)
        return out
