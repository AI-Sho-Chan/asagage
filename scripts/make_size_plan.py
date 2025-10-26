import argparse
import datetime as dt
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd

# Defaults (can later be tuned based on board logs)
TOTAL_CAP = 100_000_000  # yen
MIN_NOMINAL = 2_000_000  # yen
MAX_NOMINAL = 20_000_000  # yen


def score_row(r: pd.Series) -> float:
    # combine forward metrics into 0..1 score
    wr = float(r.get("forward_winrate", 0.5) or 0.5)
    pf = float(r.get("forward_pf_eff", 1.0) or 1.0)
    ci = float(r.get("forward_exp_boot_low", 0.0) or 0.0)
    # normalize
    wrn = max(0.0, min(1.0, (wr - 0.5) / (0.8 - 0.5)))
    pfn = max(0.0, min(1.0, (pf - 1.0) / (3.0 - 1.0)))
    cin = max(0.0, min(1.0, (ci - 0.0) / (5.0 - 0.0)))
    return float(0.45 * wrn + 0.45 * pfn + 0.10 * cin)


def make_size_plan(cand_path: Path, out_path: Path) -> Path:
    df = pd.read_csv(cand_path)
    if df.empty:
        raise SystemExit("no candidates in " + str(cand_path))
    # score and nominal preference
    df["score"] = df.apply(score_row, axis=1)
    df["nominal_pref"] = MIN_NOMINAL + df["score"] * (MAX_NOMINAL - MIN_NOMINAL)

    # approximate occupancy by expected trades per day
    trades_fwd = df.get("forward_trades", pd.Series([1] * len(df)))
    df["trades_per_day"] = trades_fwd.astype(float) / 5.0
    # if forward_avg_bars available
    minutes_map = {"AM0930": 30, "AM0945": 45, "AM1000": 60, "AM1015": 75, "AM1030": 90}
    df["minutes"] = df["session"].map(lambda s: minutes_map.get(str(s), 60))
    favg = df.get("ForwardAvgBars", df.get("forward_avg_bars", pd.Series([10.0] * len(df))))
    df["occ"] = np.minimum(1.0, (favg.astype(float) * df["trades_per_day"]) / df["minutes"].astype(float))

    # scale to total cap
    total_occ_nom = float((df["nominal_pref"] * df["occ"]).sum())
    scale = TOTAL_CAP / total_occ_nom if total_occ_nom > 0 else 1.0
    df["nominal"] = np.clip(df["nominal_pref"] * scale, MIN_NOMINAL, MAX_NOMINAL)

    out = pd.DataFrame(
        {
            "Ticker": df["Ticker"],
            "session": df["session"],
            "plan_tag": df.get("plan_tag", ""),
            "signal_mode": df["SignalMode"],
            "size_multiplier": (df["nominal"] / ((MIN_NOMINAL + MAX_NOMINAL) / 2.0)).round(3),
            "nominal_yen": df["nominal"].round(0).astype(int),
            "nominal_min": int(MIN_NOMINAL),
            "nominal_max": int(MAX_NOMINAL),
            "score": df["score"].round(3),
        }
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(out_path, index=False, encoding="utf-8-sig")
    return out_path


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidates", default="output/excel/candidates_nextday.csv")
    ap.add_argument(
        "--out",
        default=f"output/excel/size_plan/size_plan_{dt.datetime.now():%Y%m%d}.csv",
    )
    args = ap.parse_args()
    path = make_size_plan(Path(args.candidates), Path(args.out))
    print("written:", path)


if __name__ == "__main__":
    main()
