#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd


def build_stats(
    ledger_path: Path,
    output: Path,
    min_count: int,
    rules_path: Path,
    flat_quantile: float,
    trend_quantile: float,
    sigma_floor: float,
) -> None:
    if not ledger_path.exists():
        raise SystemExit(f"Ledger not found: {ledger_path}")

    ledger = pd.read_csv(ledger_path)
    required = {"code", "session", "j_abs", "J_th"}
    missing = [col for col in required if col not in ledger.columns]
    if missing:
        raise SystemExit(f"Missing columns in ledger: {', '.join(missing)}")

    data = ledger.copy()
    data["code"] = data["code"].astype(str).str.upper().str.strip()
    data["session"] = data["session"].astype(str).str.upper().str.strip()
    jth = pd.to_numeric(data["J_th"], errors="coerce")
    jabs = pd.to_numeric(data["j_abs"], errors="coerce")
    data["ratio"] = np.where(np.abs(jth) > 0, np.abs(jabs) / np.abs(jth), np.nan)
    data = data.replace([np.inf, -np.inf], np.nan).dropna(subset=["ratio"])

    grouped = (
        data.groupby(["code", "session"])["ratio"]
        .agg(count="count", ratio_mu="mean", ratio_sigma="std")
        .reset_index()
    )
    grouped = grouped[grouped["count"] >= max(1, min_count)]
    grouped["ratio_sigma"] = grouped["ratio_sigma"].fillna(sigma_floor)

    output.parent.mkdir(parents=True, exist_ok=True)
    grouped.to_csv(output, index=False)
    print(f"wrote {output} ({len(grouped)} rows)")

    update_bb_rules(
        ledger=data,
        rules_path=rules_path,
        flat_quantile=flat_quantile,
        trend_quantile=trend_quantile,
        sigma_floor=sigma_floor,
    )


def update_bb_rules(
    ledger: pd.DataFrame,
    rules_path: Path,
    flat_quantile: float,
    trend_quantile: float,
    sigma_floor: float,
) -> None:
    if not rules_path.exists():
        return
    trend_col = "nky_window_trend" if "nky_window_trend" in ledger.columns else None
    if trend_col is None:
        return

    trend_series = ledger[trend_col].astype(str).str.lower()
    ratio_series = ledger["ratio"]

    flat_mask = trend_series == "flat"
    trend_mask = trend_series != "flat"

    flat_ratio = ratio_series[flat_mask].dropna()
    trend_ratio = ratio_series[trend_mask].dropna()

    flat_k = compute_k(flat_ratio, flat_quantile, sigma_floor)
    trend_k = compute_k(trend_ratio, trend_quantile, sigma_floor)

    rules_text = rules_path.read_text(encoding="utf-8").splitlines()
    rules_text = upsert_rule(rules_text, "bb_flat_k", f"{flat_k:.4f}")
    rules_text = upsert_rule(rules_text, "bb_trend_k", f"{trend_k:.4f}")
    rules_path.write_text("\n".join(rules_text) + "\n", encoding="utf-8")
    print(f"Updated {rules_path} with bb_flat_k={flat_k:.4f}, bb_trend_k={trend_k:.4f}")


def compute_k(series: pd.Series, target_quantile: float, sigma_floor: float) -> float:
    if series.empty:
        return 1.0
    mu = float(series.mean())
    sigma = float(series.std(ddof=0))
    sigma = max(sigma, sigma_floor)
    target = float(series.quantile(target_quantile))
    k = (target - mu) / sigma if sigma > 0 else 1.0
    if not np.isfinite(k):
        k = 1.0
    return max(0.1, k)


def upsert_rule(lines: list[str], key: str, value: str) -> list[str]:
    key_lower = key.lower()
    updated: list[str] = []
    found = False
    for line in lines:
        striped = line.strip()
        if striped.lower().startswith(key_lower + "="):
            updated.append(f"{key}={value}")
            found = True
        else:
            updated.append(line)
    if not found:
        updated.append(f"{key}={value}")
    return updated


def main() -> None:
    ap = argparse.ArgumentParser(description="Build per-ticker J ratio statistics.")
    ap.add_argument("--ledger", type=Path, default=Path("analysis/all_trades_snapshot.csv"))
    ap.add_argument("--output", type=Path, default=Path("state/j_stats.csv"))
    ap.add_argument("--min-count", type=int, default=12, help="Minimum samples per (code,session)")
    ap.add_argument("--rules", type=Path, default=Path("state/strategy_rules.ini"))
    ap.add_argument("--target-flat-quantile", type=float, default=0.80)
    ap.add_argument("--target-trend-quantile", type=float, default=0.90)
    ap.add_argument("--sigma-floor", type=float, default=0.05)
    args = ap.parse_args()
    build_stats(
        ledger_path=args.ledger,
        output=args.output,
        min_count=args.min_count,
        rules_path=args.rules,
        flat_quantile=args.target_flat_quantile,
        trend_quantile=args.target_trend_quantile,
        sigma_floor=args.sigma_floor,
    )


if __name__ == "__main__":
    main()
