#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Optional

import numpy as np
import pandas as pd


DEFAULT_SOURCE = "analysis/all_trades_snapshot.csv"


def _safe_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def _bucket_overshoot(ratio: float) -> str:
    if not np.isfinite(ratio):
        return "unknown"
    if ratio < 1.0:
        return "<1.0"
    if ratio < 1.2:
        return "1.0-1.2"
    if ratio < 1.5:
        return "1.2-1.5"
    if ratio < 2.0:
        return "1.5-2.0"
    return ">=2.0"


def _bucket_bars(bars: float) -> str:
    if not np.isfinite(bars):
        return "unknown"
    if bars <= 2:
        return "0-2"
    if bars <= 5:
        return "3-5"
    if bars <= 10:
        return "6-10"
    if bars <= 20:
        return "11-20"
    return "21+"


def summarize_group(df: pd.DataFrame, keys: Iterable[str]) -> pd.DataFrame:
    pnl_bp = _safe_numeric(df["pnl_bp"])
    win = df["win"].astype(bool)
    out = (
        df.assign(pnl_bp=pnl_bp, win=win)
        .groupby(list(keys), dropna=False)
        .agg(
            trades=("pnl_bp", "count"),
            winrate=("win", "mean"),
            pnl_bp_sum=("pnl_bp", "sum"),
            pnl_bp_mean=("pnl_bp", "mean"),
            pnl_bp_median=("pnl_bp", "median"),
        )
        .reset_index()
        .sort_values(["pnl_bp_mean", "trades"], ascending=[True, False])
    )
    return out


def print_table(title: str, df: pd.DataFrame, limit: int = 15) -> None:
    print("")
    print(f"== {title} ==")
    if df.empty:
        print("(none)")
        return
    with pd.option_context("display.max_columns", 100, "display.width", 200):
        print(df.head(limit).to_string(index=False))


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default=DEFAULT_SOURCE, help=f"CSV to analyze (default: {DEFAULT_SOURCE})")
    ap.add_argument("--min-trades", type=int, default=50, help="Minimum trades to show in group summaries")
    ap.add_argument("--top-n", type=int, default=20, help="Top-N wins/losses to show")
    ap.add_argument("--out-prefix", default="", help="If set, write CSV outputs with this prefix")
    args = ap.parse_args()

    src = Path(args.source)
    if not src.exists():
        raise SystemExit(f"source not found: {src}")

    usecols = [
        "code",
        "date",
        "ts",
        "ts_exit",
        "mode",
        "signal_mode",
        "session",
        "side",
        "pnl_bp",
        "pnl_yen",
        "bars",
        "J_th",
        "j_abs",
        "gap_bucket",
        "gap_bp",
        "market_bias",
        "market_ret_5m",
        "vol_spike",
        "repeat_index",
        "forward_pf_eff",
        "forward_exp_bp",
        "win",
    ]

    df = pd.read_csv(src, usecols=lambda c: c in set(usecols))
    if df.empty:
        raise SystemExit("empty source")

    if "win" not in df.columns:
        df["win"] = _safe_numeric(df.get("pnl_bp", pd.Series([], dtype=float))) > 0
    df["pnl_bp"] = _safe_numeric(df.get("pnl_bp", pd.Series([], dtype=float)))
    df["bars"] = _safe_numeric(df.get("bars", pd.Series([], dtype=float)))
    df["J_th"] = _safe_numeric(df.get("J_th", pd.Series([], dtype=float)))
    df["j_abs"] = _safe_numeric(df.get("j_abs", pd.Series([], dtype=float)))
    df["repeat_index"] = _safe_numeric(df.get("repeat_index", pd.Series([], dtype=float)))
    df["vol_spike"] = _safe_numeric(df.get("vol_spike", pd.Series([], dtype=float)))
    df["gap_bp"] = _safe_numeric(df.get("gap_bp", pd.Series([], dtype=float)))
    df["market_ret_5m"] = _safe_numeric(df.get("market_ret_5m", pd.Series([], dtype=float)))

    # Derived: overshoot ratio = |J| / J_th at entry.
    ratio = df["j_abs"] / df["J_th"].replace(0, np.nan)
    df["j_overshoot"] = ratio
    df["overshoot_bucket"] = ratio.apply(_bucket_overshoot)
    df["bars_bucket"] = df["bars"].apply(_bucket_bars)
    df["repeat_bucket"] = df["repeat_index"].apply(
        lambda x: "unknown"
        if not np.isfinite(x)
        else ("1" if x <= 1 else ("2" if x == 2 else ("3" if x == 3 else "4+")))
    )

    # Overall
    total = int(df["pnl_bp"].notna().sum())
    wins = int((df["pnl_bp"] > 0).sum())
    losses = int((df["pnl_bp"] < 0).sum())
    print(f"ASAGAKE trade-pattern report from {src}")
    print(f"trades={total} wins={wins} losses={losses} winrate={wins/total:.3f}")
    print(
        f"pnl_bp_sum={df['pnl_bp'].sum():.1f} pnl_bp_mean={df['pnl_bp'].mean():.3f} pnl_bp_median={df['pnl_bp'].median():.3f}"
    )

    # Worst / best trades
    keep_cols = [
        "date",
        "code",
        "session",
        "signal_mode",
        "side",
        "mode",
        "pnl_bp",
        "bars",
        "gap_bucket",
        "gap_bp",
        "market_bias",
        "vol_spike",
        "repeat_index",
        "j_abs",
        "J_th",
        "j_overshoot",
    ]
    worst = df.sort_values("pnl_bp").head(args.top_n)[keep_cols]
    best = df.sort_values("pnl_bp", ascending=False).head(args.top_n)[keep_cols]
    print_table(f"Worst {args.top_n} trades (pnl_bp)", worst, limit=args.top_n)
    print_table(f"Best {args.top_n} trades (pnl_bp)", best, limit=args.top_n)

    # Group summaries
    by_plan = summarize_group(df, ["session", "signal_mode"])
    by_plan = by_plan[by_plan["trades"] >= args.min_trades].sort_values(
        ["pnl_bp_mean", "trades"], ascending=[True, False]
    )
    print_table(f"By session x signal (trades>={args.min_trades})", by_plan, limit=25)

    by_gap = summarize_group(df, ["gap_bucket"])
    by_gap = by_gap[by_gap["trades"] >= args.min_trades]
    print_table(f"By gap bucket (trades>={args.min_trades})", by_gap, limit=25)

    if "market_bias" in df.columns:
        by_bias = summarize_group(df, ["market_bias"])
        by_bias = by_bias[by_bias["trades"] >= args.min_trades]
        print_table(f"By market bias (trades>={args.min_trades})", by_bias, limit=25)

    by_overshoot = summarize_group(df, ["overshoot_bucket"])
    by_overshoot = by_overshoot[by_overshoot["trades"] >= args.min_trades]
    print_table(f"By |J|/J_th bucket (trades>={args.min_trades})", by_overshoot, limit=25)

    by_bars = summarize_group(df, ["bars_bucket"])
    by_bars = by_bars[by_bars["trades"] >= args.min_trades]
    print_table(f"By holding bars bucket (trades>={args.min_trades})", by_bars, limit=25)

    by_repeat = summarize_group(df, ["repeat_bucket"])
    by_repeat = by_repeat[by_repeat["trades"] >= args.min_trades]
    print_table(f"By repeat index bucket (trades>={args.min_trades})", by_repeat, limit=25)

    if args.out_prefix:
        out_prefix = Path(args.out_prefix)
        out_prefix.parent.mkdir(parents=True, exist_ok=True)
        worst.to_csv(str(out_prefix) + "_worst.csv", index=False, encoding="utf-8-sig")
        best.to_csv(str(out_prefix) + "_best.csv", index=False, encoding="utf-8-sig")
        by_plan.to_csv(str(out_prefix) + "_by_plan.csv", index=False, encoding="utf-8-sig")
        by_gap.to_csv(str(out_prefix) + "_by_gap.csv", index=False, encoding="utf-8-sig")
        by_overshoot.to_csv(str(out_prefix) + "_by_overshoot.csv", index=False, encoding="utf-8-sig")
        by_bars.to_csv(str(out_prefix) + "_by_bars.csv", index=False, encoding="utf-8-sig")
        by_repeat.to_csv(str(out_prefix) + "_by_repeat.csv", index=False, encoding="utf-8-sig")
        if "market_bias" in df.columns:
            by_bias.to_csv(str(out_prefix) + "_by_market_bias.csv", index=False, encoding="utf-8-sig")
        print(f"\nWrote CSV outputs with prefix: {out_prefix}")


if __name__ == "__main__":
    main()
