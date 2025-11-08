#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Iterable, List

import numpy as np
import pandas as pd


def _decode_series(value: object, cast=float) -> List[float]:
    if isinstance(value, str) and value.strip():
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            return []
        if isinstance(parsed, Iterable):
            out: List[float] = []
            for item in parsed:
                try:
                    out.append(cast(item))
                except Exception:
                    continue
            return out
    return []


def _plan_from_path(path: Path) -> str:
    name = path.name
    if name.startswith("RUN_"):
        name = name[4:]
    return name


def build_detail_frame(run_root: Path) -> pd.DataFrame:
    records: List[dict] = []
    for summary_path in sorted(run_root.rglob("_SUMMARY_FORWARD.csv")):
        plan = _plan_from_path(summary_path.parent)
        try:
            df = pd.read_csv(summary_path)
        except Exception:
            continue
        if df.empty:
            continue

        for _, row in df.iterrows():
            slice_pf = _decode_series(row.get("forward_slice_pf_series"), float)
            slice_exp = _decode_series(row.get("forward_slice_exp_series"), float)
            slice_trades = _decode_series(row.get("forward_slice_trades"), int)
            total = int(row.get("forward_slices_total", 0))
            passed = int(row.get("forward_slices_pass", 0))
            pass_ratio = (passed / total) if total else 0.0
            pf_min = float(np.min(slice_pf)) if slice_pf else np.nan
            pf_mean = float(np.mean(slice_pf)) if slice_pf else np.nan
            pf_med = float(np.median(slice_pf)) if slice_pf else np.nan
            pf_std = float(np.std(slice_pf, ddof=0)) if len(slice_pf) > 1 else (0.0 if slice_pf else np.nan)
            exp_mean = float(np.mean(slice_exp)) if slice_exp else np.nan
            exp_min = float(np.min(slice_exp)) if slice_exp else np.nan
            trades_median = float(np.median(slice_trades)) if slice_trades else np.nan
            trades_min = float(np.min(slice_trades)) if slice_trades else np.nan
            worst_idx = int(np.argmin(slice_pf)) if slice_pf else -1
            worst_pf = slice_pf[worst_idx] if slice_pf else np.nan

            records.append(
                {
                    "plan": plan,
                    "code": row.get("code"),
                    "signal_mode": row.get("signal_mode"),
                    "forward_trades": row.get("forward_trades"),
                    "forward_pf_eff": row.get("forward_pf_eff"),
                    "forward_exp_bp": row.get("forward_exp_bp"),
                    "slices_total": total,
                    "slices_pass": passed,
                    "pass_ratio": pass_ratio,
                    "slice_pf_min": pf_min,
                    "slice_pf_mean": pf_mean,
                    "slice_pf_median": pf_med,
                    "slice_pf_std": pf_std,
                    "slice_pf_worst": worst_pf,
                    "slice_pf_worst_index": worst_idx,
                    "slice_exp_mean": exp_mean,
                    "slice_exp_min": exp_min,
                    "slice_trades_median": trades_median,
                    "slice_trades_min": trades_min,
                }
            )
    return pd.DataFrame.from_records(records)


def build_plan_summary(detail_df: pd.DataFrame, *, pf_threshold: float, pass_threshold: float) -> pd.DataFrame:
    if detail_df.empty:
        return pd.DataFrame(
            columns=[
                "plan",
                "codes",
                "avg_pass_ratio",
                "min_pass_ratio",
                "avg_slice_pf_min",
                "median_slice_pf_min",
                "share_slice_pf_min_ge_threshold",
                "share_pass_ratio_ge_threshold",
                "avg_slice_exp_mean",
                "median_slice_trades",
            ]
        )

    def ratio_ge(series: pd.Series, threshold: float) -> float:
        mask = series.dropna()
        if mask.empty:
            return 0.0
        return float((mask >= threshold).mean())

    grouped = detail_df.groupby("plan", dropna=False)
    summary = pd.DataFrame(
        {
            "codes": grouped["code"].count(),
            "avg_pass_ratio": grouped["pass_ratio"].mean(),
            "min_pass_ratio": grouped["pass_ratio"].min(),
            "avg_slice_pf_min": grouped["slice_pf_min"].mean(),
            "median_slice_pf_min": grouped["slice_pf_min"].median(),
            "share_slice_pf_min_ge_threshold": grouped["slice_pf_min"].apply(lambda s: ratio_ge(s, pf_threshold)),
            "share_pass_ratio_ge_threshold": grouped["pass_ratio"].apply(lambda s: ratio_ge(s, pass_threshold)),
            "avg_slice_exp_mean": grouped["slice_exp_mean"].mean(),
            "median_slice_trades": grouped["slice_trades_median"].median(),
        }
    ).reset_index()
    return summary


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", required=True, help="Path to NIGHTLY_xxxx directory")
    parser.add_argument("--output", required=True, help="CSV file for per-code walk-forward metrics")
    parser.add_argument("--plan-output", help="CSV file for per-plan summary")
    parser.add_argument("--pf-threshold", type=float, default=1.0, help="PF threshold for stability ratio (default 1.0)")
    parser.add_argument(
        "--pass-threshold",
        type=float,
        default=0.75,
        help="Slice pass ratio threshold for stability ratio (default 0.75)",
    )
    args = parser.parse_args()

    run_root = Path(args.run_root)
    detail_df = build_detail_frame(run_root)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    detail_df.to_csv(output_path, index=False)

    if args.plan_output:
        plan_df = build_plan_summary(
            detail_df,
            pf_threshold=float(args.pf_threshold),
            pass_threshold=float(args.pass_threshold),
        )
        plan_path = Path(args.plan_output)
        plan_path.parent.mkdir(parents=True, exist_ok=True)
        plan_df.to_csv(plan_path, index=False)


if __name__ == "__main__":
    main()
