from dataclasses import dataclass

@dataclass
class ScreenerConfig:
    out_watchlist_path: str = "out/watchlist.txt"
    aoi_min_abs: float = 0.40
    aoi_max_sigma: float = 0.10
    top_k: int = 10
    min_samples: int = 8
