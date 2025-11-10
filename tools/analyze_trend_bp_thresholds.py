#!/usr/bin/env python3
"""
Compute performance metrics under different gap/Trend bp thresholds.

We currently treat the per-trade gap_bp from analysis/all_trades_snapshot.csv
as a proxy for the NKY trend magnitude.  This isn't perfect, but gives a
repeatable way to compare filtering thresholds until we have richer NKY data.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

import pandas as pd


def summarize(df: pd.DataFrame) -> dict[str, float]:
    if df.empty:
        return {"trades": 0, "win_rate": 0.0, "avg_bp": 0.0, "sum_bp": 0.0}
    wins = (df["pnl_bp"] > 0).sum()
    return {
        "trades": int(len(df)),
        "win_rate": wins / len(df),
        "avg_bp": float(df["pnl_bp"].mean()),
        "sum_bp": float(df["pnl_bp"].sum()),
    }


def analyze(trades: pd.DataFrame, thresholds: Iterable[int]) -> pd.DataFrame:
    records: list[dict[str, float | str | bool]] = []
    trades = trades.copy()
    trades["gap_abs"] = trades["gap_bp"].abs()
    for th in thresholds:
        mask = trades["gap_abs"] >= th
        for include_flat in (True, False):
            current_mask = mask.copy()
            if not include_flat:
                current_mask &= trades["market_bias"] != "neutral"
            df = trades[current_mask]
            stats = summarize(df)
            stats.update(
                {
                    "threshold_bp": th,
                    "include_flat": include_flat,
                }
            )
            records.append(stats)
    return pd.DataFrame(records)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument(
        "--trades",
        default=Path("analysis/all_trades_snapshot.csv"),
        type=Path,
    )
    ap.add_argument(
        "--thresholds",
        default="15,20,25,30",
        help="Comma-separated bp thresholds",
    )
    ap.add_argument(
        "--output",
        default=Path("analysis/trend_bp_thresholds.csv"),
        type=Path,
    )
    args = ap.parse_args()

    thresholds = [int(x) for x in args.thresholds.split(",") if x.strip()]
    if not thresholds:
        raise SystemExit("No thresholds provided")

    trades = pd.read_csv(args.trades)
    if "gap_bp" not in trades.columns or "market_bias" not in trades.columns:
        raise SystemExit("analysis/all_trades_snapshot.csv missing required columns")

    df = analyze(trades, thresholds)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(args.output, index=False)
    print(f"written {args.output} ({len(df)} rows)")


if __name__ == "__main__":
    main()
