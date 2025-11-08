#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


def build_stats(ledger: Path, output: Path, min_count: int) -> None:
    if not ledger.exists():
        raise SystemExit(f"Ledger not found: {ledger}")
    df = pd.read_csv(ledger)
    required = {"code", "session", "j_abs", "J_th"}
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise SystemExit(f"Missing columns in ledger: {', '.join(missing)}")

    df = df.copy()
    df["code"] = df["code"].astype(str).str.upper().str.strip()
    df["session"] = df["session"].astype(str).str.upper().str.strip()
    jth = pd.to_numeric(df["J_th"], errors="coerce")
    jabs = pd.to_numeric(df["j_abs"], errors="coerce")
    ratio = np.where(np.abs(jth) > 0, np.abs(jabs) / np.abs(jth), np.nan)
    df["ratio"] = ratio
    df = df.replace([np.inf, -np.inf], np.nan).dropna(subset=["ratio"])

    grouped = (
        df.groupby(["code", "session"])["ratio"]
        .agg(count="count", ratio_mu="mean", ratio_sigma="std")
        .reset_index()
    )
    grouped = grouped[grouped["count"] >= max(1, min_count)]
    grouped["ratio_sigma"] = grouped["ratio_sigma"].fillna(0.05)

    output.parent.mkdir(parents=True, exist_ok=True)
    grouped.to_csv(output, index=False)
    print(f"wrote {output} ({len(grouped)} rows)")


def main() -> None:
    ap = argparse.ArgumentParser(description="Build per-ticker J ratio statistics.")
    ap.add_argument("--ledger", type=Path, default=Path("analysis/all_trades_snapshot.csv"))
    ap.add_argument("--output", type=Path, default=Path("state/j_stats.csv"))
    ap.add_argument("--min-count", type=int, default=12, help="Minimum samples per (code,session)")
    args = ap.parse_args()
    build_stats(args.ledger, args.output, args.min_count)


if __name__ == "__main__":
    main()

