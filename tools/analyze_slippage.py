#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

DATA_ROOT = Path("data/raw/yahoo_1m")


def _plan_from_path(path: Path) -> str:
    name = path.name
    if name.startswith("RUN_"):
        name = name[4:]
    return name


def _load_minute_frame(code: str, trade_date: str, cache: Dict[Tuple[str, str], pd.DataFrame]) -> Optional[pd.DataFrame]:
    key = (code, trade_date)
    if key in cache:
        return cache[key]
    parquet_path = DATA_ROOT / code / f"{trade_date}.parquet"
    if not parquet_path.exists():
        return None
    try:
        df = pd.read_parquet(parquet_path)
    except Exception:
        return None
    if not isinstance(df.index, pd.DatetimeIndex):
        df.index = pd.to_datetime(df.index)
    if df.index.tz is None:
        df.index = df.index.tz_localize("Asia/Tokyo")
    else:
        df.index = df.index.tz_convert("Asia/Tokyo")
    cache[key] = df
    return df


def _iter_trade_details(value: object) -> Iterable[dict]:
    if isinstance(value, str) and value.strip().startswith("["):
        try:
            parsed = json.loads(value)
        except json.JSONDecodeError:
            return []
        if isinstance(parsed, list):
            return parsed
    return []


def collect_slippage_records(run_root: Path) -> pd.DataFrame:
    cache: Dict[Tuple[str, str], pd.DataFrame] = {}
    records: List[dict] = []
    for compare_path in sorted(run_root.rglob("_COMPARE.csv")):
        plan = _plan_from_path(compare_path.parent)
        try:
            df = pd.read_csv(compare_path)
        except Exception:
            continue
        if df.empty:
            continue
        for _, row in df.iterrows():
            details_iter = _iter_trade_details(row.get("forward_trade_details"))
            for detail in details_iter:
                ts_str = detail.get("ts")
                code = detail.get("code") or row.get("code")
                side = detail.get("side")
                if not ts_str or not code or side not in {"BUY", "SELL"}:
                    continue
                try:
                    ts = pd.Timestamp(ts_str)
                except Exception:
                    continue
                if ts.tzinfo is None:
                    ts = ts.tz_localize("Asia/Tokyo")
                else:
                    ts = ts.tz_convert("Asia/Tokyo")
                trade_date = ts.strftime("%Y-%m-%d")
                minute_df = _load_minute_frame(str(code), trade_date, cache)
                if minute_df is None or ts not in minute_df.index:
                    continue
                row_bar = minute_df.loc[ts]
                try:
                    locator = minute_df.index.get_loc(ts)
                except KeyError:
                    continue
                next_row = minute_df.iloc[locator + 1] if locator + 1 < len(minute_df.index) else None

                entry_price = float(row_bar["close"])
                high = float(row_bar["high"])
                low = float(row_bar["low"])
                volume = float(row_bar.get("volume", 0.0))

                if side == "BUY":
                    intra_diff = max(0.0, high - entry_price)
                    next_diff = 0.0
                    if next_row is not None:
                        next_open = float(next_row["open"])
                        next_diff = max(0.0, next_open - entry_price)
                else:
                    intra_diff = max(0.0, entry_price - low)
                    next_diff = 0.0
                    if next_row is not None:
                        next_open = float(next_row["open"])
                        next_diff = max(0.0, entry_price - next_open)

                intra_bp = (intra_diff / entry_price) * 10000 if entry_price else 0.0
                next_bp = (next_diff / entry_price) * 10000 if entry_price else 0.0

                records.append(
                    {
                        "plan": plan,
                        "code": code,
                        "side": side,
                        "ts": ts.isoformat(),
                        "entry_price": entry_price,
                        "volume": volume,
                        "intra_adverse_bp": intra_bp,
                        "next_adverse_bp": next_bp,
                        "gap_bp": detail.get("gap_bp"),
                    }
                )
    return pd.DataFrame.from_records(records)


def build_plan_summary(detail_df: pd.DataFrame, *, threshold_bp: float) -> pd.DataFrame:
    if detail_df.empty:
        return pd.DataFrame(
            columns=[
                "plan",
                "observations",
                "median_intra_adverse_bp",
                "p95_intra_adverse_bp",
                "median_next_adverse_bp",
                "p95_next_adverse_bp",
                "share_next_adverse_gt_threshold",
                "median_volume",
            ]
        )

    grouped = detail_df.groupby("plan", dropna=False)
    summary = pd.DataFrame(
        {
            "observations": grouped.size(),
            "median_intra_adverse_bp": grouped["intra_adverse_bp"].median(),
            "p95_intra_adverse_bp": grouped["intra_adverse_bp"].quantile(0.95, interpolation="linear"),
            "median_next_adverse_bp": grouped["next_adverse_bp"].median(),
            "p95_next_adverse_bp": grouped["next_adverse_bp"].quantile(0.95, interpolation="linear"),
            "share_next_adverse_gt_threshold": grouped["next_adverse_bp"].apply(
                lambda s: float((s > threshold_bp).mean()) if not s.empty else 0.0
            ),
            "median_volume": grouped["volume"].median(),
        }
    ).reset_index()
    return summary


def build_recommendations(plan_df: pd.DataFrame) -> pd.DataFrame:
    if plan_df.empty:
        return pd.DataFrame(columns=["plan", "session", "entry_buffer_bp", "size_multiplier"])

    def calculate_size_multiplier(share_gt10: float) -> float:
        if share_gt10 >= 0.1:
            return 0.75
        if share_gt10 >= 0.05:
            return 0.9
        return 1.0

    records: List[dict] = []
    for _, row in plan_df.iterrows():
        plan = str(row.get("plan", ""))
        session = ""
        if "_" in plan:
            parts = plan.split("_")
            for part in parts:
                if part.startswith("AM") or part.startswith("PM"):
                    session = part
                    break
        buffer_bp = float(row.get("p95_next_adverse_bp", 0.0))
        share_gt10 = float(row.get("share_next_adverse_gt_threshold", 0.0))
        records.append(
            {
                "plan": plan,
                "session": session,
                "entry_buffer_bp": round(buffer_bp, 2),
                "size_multiplier": round(calculate_size_multiplier(share_gt10), 2),
            }
        )
    return pd.DataFrame.from_records(records)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-root", required=True, help="Path to NIGHTLY_xxxx directory")
    parser.add_argument("--output", required=True, help="CSV file for per-trade slippage metrics")
    parser.add_argument("--plan-output", help="CSV summary aggregated per plan")
    parser.add_argument(
        "--threshold-bp",
        type=float,
        default=10.0,
        help="Basis-point threshold for adverse gap ratio (default 10bp)",
    )
    parser.add_argument("--recommend-output", help="CSV with recommended buffers per plan")
    args = parser.parse_args()

    run_root = Path(args.run_root)
    detail_df = collect_slippage_records(run_root)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    detail_df.to_csv(output_path, index=False)

    if args.plan_output:
        plan_df = build_plan_summary(detail_df, threshold_bp=float(args.threshold_bp))
        plan_path = Path(args.plan_output)
        plan_path.parent.mkdir(parents=True, exist_ok=True)
        plan_df.to_csv(plan_path, index=False)

        if args.recommend_output:
            rec_df = build_recommendations(plan_df)
            rec_path = Path(args.recommend_output)
            rec_path.parent.mkdir(parents=True, exist_ok=True)
            rec_df.to_csv(rec_path, index=False)


if __name__ == "__main__":
    main()
