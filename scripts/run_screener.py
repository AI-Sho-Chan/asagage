from __future__ import annotations
import argparse
from pathlib import Path
from asagake.config import ScreenerConfig
from asagake.screener import SelectionParams, select_symbols
from asagake.providers.mock import MockProvider
try:
    from asagake.providers.kabu import KabuProvider
except Exception:
    KabuProvider = None  # 未開通時は None

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", default="out/watchlist.txt")
    parser.add_argument("--top-k", type=int, default=10)
    parser.add_argument("--min-aoi", type=float, default=0.40)
    parser.add_argument("--max-sigma", type=float, default=0.10)
    parser.add_argument("--min-samples", type=int, default=8)
    parser.add_argument("--provider", choices=["mock","kabu"], default="mock")
    args = parser.parse_args()

    cfg = ScreenerConfig(out_watchlist_path=args.out,
                         aoi_min_abs=args.min_aoi,
                         aoi_max_sigma=args.max_sigma,
                         top_k=args.top_k,
                         min_samples=args.min_samples)

    # Provider 選択
    prov = MockProvider() if (args.provider=="mock" or KabuProvider is None) else KabuProvider()

    # 一次抽出（ランキング想定 → Mock は固定）
    universe = prov.get_liquidity_universe(max_n=200)

    # AOI 入力収集（Mockは疑似）
    aoi_inputs = prov.collect_aoi_input(universe)

    params = SelectionParams(min_abs_aoi=cfg.aoi_min_abs,
                             max_sigma=cfg.aoi_max_sigma,
                             top_k=cfg.top_k,
                             min_samples=cfg.min_samples)
    selected = select_symbols(aoi_inputs, params)

    Path("out").mkdir(exist_ok=True, parents=True)
    Path(cfg.out_watchlist_path).write_text("\n".join(selected), encoding="utf-8")
    print(f"[OK] wrote {len(selected)} symbols -> {cfg.out_watchlist_path}")
    for s in selected:
        print("  -", s)

if __name__ == "__main__":
    main()
