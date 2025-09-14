from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List, Callable, Optional, Tuple
import math
from .aoi import compute_aoi_series, summarize_aoi

@dataclass
class SelectionParams:
    min_abs_aoi: float = 0.40
    max_sigma: float = 0.10
    top_k: int = 10
    min_samples: int = 8  # 欠測耐性：有効点がこの数未満なら除外

def select_symbols(
    aoi_inputs: Dict[str, Tuple[list[float], list[float]]],
    params: SelectionParams,
    tie_break_score: Optional[Callable[[str], float]] = None,
) -> List[str]:
    scored: List[tuple[str, float, float]] = []
    for sym, (bid_list, ask_list) in aoi_inputs.items():
        series = compute_aoi_series(bid_list, ask_list)
        if len(series) < params.min_samples:
            continue
        m = summarize_aoi(series)
        if math.isnan(m.latest) or math.isnan(m.sigma):
            continue
        if abs(m.latest) < params.min_abs_aoi or m.sigma > params.max_sigma:
            continue
        base = abs(m.latest)
        bonus = tie_break_score(sym) if tie_break_score else 0.0
        scored.append((sym, base, bonus))
    scored.sort(key=lambda x: (x[1], x[2]), reverse=True)
    return [s for s, _, _ in scored[: params.top_k]]
