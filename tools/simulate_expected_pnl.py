#!/usr/bin/env python3
"""Compute expected daily P&L from candidates CSV files.

Usage:
    python tools/simulate_expected_pnl.py --budget 10000000 --current output/excel/candidates_nextday.csv [--baseline path]

Returns JSON with expected_yen, per_position and optional delta vs baseline.
"""

import argparse
import json
from pathlib import Path
from typing import Dict

import pandas as pd


def load_candidates(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(path)
    df = pd.read_csv(path)
    # normalise column names to lower
    df.columns = [c.strip() for c in df.columns]
    return df


def expected_pnl(df: pd.DataFrame, budget: float) -> Dict[str, float]:
    if df.empty:
        return {"positions": 0, "expected_yen": 0.0, "per_position": 0.0}

    # Use forward_exp_boot_mean (bp) as expected return proxy; fall back to forward_pf_eff*100?
    exp_col = None
    for cand in ("forward_exp_boot_mean", "ExpBootMean", "exp_boot_mean"):
        if cand in df.columns:
            exp_col = cand
            break
    if exp_col is None:
        raise ValueError("forward_exp_boot_mean column not found")

    exp_bp = pd.to_numeric(df[exp_col], errors="coerce").fillna(0.0)
    positions = len(exp_bp)
    alloc = budget / positions if positions else 0.0
    expected_yen = float((alloc * exp_bp / 10000.0).sum())
    return {
        "positions": positions,
        "expected_yen": expected_yen,
        "per_position": expected_yen / positions if positions else 0.0,
    }


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--budget", type=float, default=10_000_000)
    ap.add_argument("--current", type=Path, required=True)
    ap.add_argument("--baseline", type=Path)
    args = ap.parse_args()

    result = {"budget": args.budget}

    cur_df = load_candidates(args.current)
    cur_stats = expected_pnl(cur_df, args.budget)
    result["current"] = cur_stats

    if args.baseline:
        base_df = load_candidates(args.baseline)
        base_stats = expected_pnl(base_df, args.budget)
        result["baseline"] = base_stats
        result["delta"] = {
            "expected_yen": cur_stats["expected_yen"] - base_stats["expected_yen"],
            "per_position": cur_stats["per_position"] - base_stats["per_position"],
        }

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
