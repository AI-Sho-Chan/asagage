from dataclasses import dataclass
from typing import Iterable, List, Optional
import numpy as np

@dataclass
class AOIMetrics:
    latest: float
    sigma: float
    n: int

def compute_aoi_series(bid_qty: Iterable[Optional[float]],
                       ask_qty: Iterable[Optional[float]]) -> List[float]:
    """欠測(None)やゼロ合計は捨象して AOI=(Bid-Ask)/(Bid+Ask) を列挙"""
    out: List[float] = []
    for b, a in zip(bid_qty, ask_qty):
        if b is None or a is None:
            continue
        s = b + a
        if s == 0:
            continue
        out.append((b - a) / s)
    return out

def summarize_aoi(series: List[float]) -> AOIMetrics:
    if not series:
        return AOIMetrics(np.nan, np.nan, 0)
    arr = np.asarray(series, dtype=float)
    return AOIMetrics(latest=float(arr[-1]), sigma=float(arr.std(ddof=0)), n=int(arr.size))
