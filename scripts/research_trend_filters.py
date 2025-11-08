#!/usr/bin/env python3
"""
Rough-cut analysis for index driver & direction-aligned filters.

Loads hypothesis trade ledgers, limits to recent trades, and compares
Nikkei (現物), Nikkei先物（CME/USDミニ）, and TOPIX ETFを方向判定ドライバとして評価する。

Outputs:
  - analysis/trend_alignment_summary.csv
  - analysis/trend_alignment_improvement.csv
  - analysis/trend_ticker_preference.csv
  - analysis/trend_driver_overview.json
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Dict, Iterable, List
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import numpy as np
import pandas as pd

from analyze_trades import BT_ROOT, fetch_daily_prices, fetch_intraday_prices, load_ledger_frames


# Yahoo Finance symbols for each driver.
DRIVERS = {
    "NKY": {"symbol": "^N225"},       # 現物 Nikkei 225
    "NKY_F": {"symbol": "NIY=F"},     # CME Nikkei 225 futures (USD)
    "TOPIX": {"symbol": "1306.T"},    # TOPIX ETF proxy
}

TREND_TYPES = ("day", "window")


def classify_direction(ret_bp: float, threshold: float) -> str:
    if np.isnan(ret_bp):
        return "missing"
    if ret_bp > threshold:
        return "up"
    if ret_bp < -threshold:
        return "down"
    return "flat"


def ensure_utc(series: pd.Series) -> pd.Series:
    ser = pd.to_datetime(series, utc=True)
    return ser


def attach_day_trend(
    trades: pd.DataFrame, daily: pd.DataFrame, driver: str, threshold: float
) -> None:
    prefix = f"{driver}_day"
    if daily.empty:
        trades[f"{prefix}_ret_bp"] = np.nan
        trades[f"{prefix}_trend"] = "missing"
        return

    day = daily.copy()
    if "open" not in day.columns or "close" not in day.columns:
        trades[f"{prefix}_ret_bp"] = np.nan
        trades[f"{prefix}_trend"] = "missing"
        return

    day["ret_bp"] = (day["close"] - day["open"]) / day["open"] * 10000
    trades[f"{prefix}_ret_bp"] = trades["date"].map(day["ret_bp"])
    trades[f"{prefix}_trend"] = trades[f"{prefix}_ret_bp"].apply(
        lambda x: classify_direction(x, threshold)
    )


def attach_window_trend(
    trades: pd.DataFrame, intraday: pd.Series, driver: str, threshold: float
) -> None:
    prefix = f"{driver}_window"
    if intraday.empty:
        trades[f"{prefix}_ret_bp"] = np.nan
        trades[f"{prefix}_trend"] = "missing"
        trades[f"{prefix}_alignment"] = "missing"
        return

    series = intraday.copy()
    if series.index.tz is None:
        series = series.tz_localize("UTC")
    else:
        series = series.tz_convert("UTC")

    start_px = series.reindex(trades["ts"], method="ffill").to_numpy()
    end_px = series.reindex(trades["ts_exit"], method="ffill").to_numpy()
    valid = (start_px > 0) & (end_px > 0)
    window_ret = np.full(len(trades), np.nan)
    window_ret[valid] = (end_px[valid] - start_px[valid]) / start_px[valid] * 10000

    trend_col = f"{prefix}_trend"
    trades[f"{prefix}_ret_bp"] = window_ret
    trades[trend_col] = trades[f"{prefix}_ret_bp"].apply(
        lambda x: classify_direction(x, threshold)
    )

    align_col = f"{prefix}_alignment"

    def _label(trend: str, side: str) -> str:
        if pd.isna(side) or trend in {"missing", "flat"}:
            return trend
        if (trend == "up" and side == "BUY") or (trend == "down" and side == "SELL"):
            return "aligned"
        return "counter"

    trades[align_col] = [
        _label(trend, side) for trend, side in zip(trades[trend_col], trades["side"])
    ]


def build_daily_from_intraday(
    intraday: pd.Series,
    open_time: time,
    close_time: time,
    tz: str = "Asia/Tokyo",
) -> pd.DataFrame:
    if intraday.empty:
        return pd.DataFrame(columns=["open", "close"])
    series = intraday.copy()
    if series.index.tz is None:
        series = series.tz_localize("UTC")
    series = series.tz_convert(tz)
    df = series.to_frame("price")
    df["date"] = df.index.date
    df["time"] = df.index.time

    records = []
    for date, group in df.groupby("date"):
        open_candidates = group[group["time"] >= open_time]
        close_candidates = group[group["time"] <= close_time]
        open_px = open_candidates["price"].iloc[0] if not open_candidates.empty else np.nan
        close_px = close_candidates["price"].iloc[-1] if not close_candidates.empty else np.nan
        records.append({"date": date, "open": open_px, "close": close_px})

    out = pd.DataFrame(records).set_index("date")
    return out


def attach_day_alignment(trades: pd.DataFrame, driver: str) -> None:
    trend_col = f"{driver}_day_trend"
    align_col = f"{driver}_day_alignment"

    def _label(trend: str, side: str) -> str:
        if pd.isna(side) or trend in {"missing", "flat"}:
            return trend
        if (trend == "up" and side == "BUY") or (trend == "down" and side == "SELL"):
            return "aligned"
        return "counter"

    trades[align_col] = [
        _label(trend, side) for trend, side in zip(trades[trend_col], trades["side"])
    ]


def calc_stats(subset: pd.DataFrame) -> Dict[str, float]:
    if subset.empty:
        return {"count": 0, "mean_bp": np.nan, "median_bp": np.nan, "win_rate": np.nan, "pf": np.nan, "std_bp": np.nan}
    pnl = subset["pnl_bp"].to_numpy()
    count = len(pnl)
    mean_bp = float(np.mean(pnl))
    median_bp = float(np.median(pnl))
    std_bp = float(np.std(pnl, ddof=0))
    wins = float((pnl > 0).sum()) / count if count else np.nan
    pos = pnl[pnl > 0].sum()
    neg = pnl[pnl < 0].sum()
    pf = float(pos / abs(neg)) if neg < 0 else np.nan
    return {
        "count": count,
        "mean_bp": mean_bp,
        "median_bp": median_bp,
        "win_rate": wins,
        "pf": pf,
        "std_bp": std_bp,
    }


def summarize_alignment(trades: pd.DataFrame, min_count: int) -> pd.DataFrame:
    records: List[Dict[str, object]] = []
    for driver in DRIVERS:
        for trend_type in TREND_TYPES:
            trend_col = f"{driver}_{trend_type}_trend"
            align_col = f"{driver}_{trend_type}_alignment"
            if trend_type == "day":
                align_col = f"{driver}_day_alignment"

            base_stats = calc_stats(trades)
            for policy in ("all", "aligned", "counter", "flat"):
                if policy == "all":
                    subset = trades
                elif policy == "flat":
                    subset = trades[trades[trend_col] == "flat"]
                else:
                    subset = trades[trades[align_col] == policy]

                stats = calc_stats(subset)
                stats = {k: (None if np.isnan(v) else v) for k, v in stats.items()}

                records.append(
                    {
                        "driver": driver,
                        "trend_type": trend_type,
                        "policy": policy,
                        **stats,
                        "delta_vs_all": None
                        if stats["mean_bp"] is None or np.isnan(base_stats["mean_bp"])
                        else stats["mean_bp"] - base_stats["mean_bp"],
                    }
                )
    return pd.DataFrame(records)


def summarize_ticker_preferences(
    trades: pd.DataFrame, min_count: int
) -> pd.DataFrame:
    records: List[Dict[str, object]] = []
    for code, df_code in trades.groupby("code"):
        best_choice = None
        for driver in DRIVERS:
            for trend_type in TREND_TYPES:
                align_col = (
                    f"{driver}_{trend_type}_alignment"
                    if trend_type == "window"
                    else f"{driver}_day_alignment"
                )
                subset = df_code[df_code[align_col] == "aligned"]
                if len(subset) < min_count:
                    continue
                stats = calc_stats(subset)
                mean_bp = stats["mean_bp"]
                if np.isnan(mean_bp):
                    continue
                candidate = {
                    "driver": driver,
                    "trend_type": trend_type,
                    "mean_bp": mean_bp,
                    "count": stats["count"],
                    "win_rate": stats["win_rate"],
                    "pf": stats["pf"],
                }
                if (
                    best_choice is None
                    or mean_bp > best_choice["mean_bp"]
                    or (
                        np.isclose(mean_bp, best_choice["mean_bp"])
                        and stats["count"] > best_choice["count"]
                    )
                ):
                    best_choice = candidate
        if best_choice:
            records.append({"code": code, **best_choice})
    return pd.DataFrame(records)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Compare trend drivers for direction filters.")
    parser.add_argument("--bt-root", type=Path, default=BT_ROOT, help="Backtest root with hypothesis_trade_ledger.csv files.")
    parser.add_argument("--output-dir", type=Path, default=Path("analysis"), help="Directory to store analysis CSV/JSON.")
    parser.add_argument("--lookback-days", type=int, default=30)
    parser.add_argument("--top-n", type=int, default=100)
    parser.add_argument("--sessions", type=str, default="AM15,AM0930,PM1")
    parser.add_argument("--dir-threshold-bp", type=float, default=15.0)
    parser.add_argument("--min-count", type=int, default=8)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    args.output_dir.mkdir(exist_ok=True, parents=True)

    trades = load_ledger_frames(args.bt_root)
    trades["ts"] = ensure_utc(trades["ts"])
    trades["ts_exit"] = ensure_utc(trades["ts_exit"])
    trades["date"] = pd.to_datetime(trades["date"]).dt.date
    trades["session"] = trades["session"].astype(str)
    trades["side"] = trades["side"].astype(str).str.upper()

    max_date = trades["date"].max()
    min_date = (pd.Timestamp(max_date) - timedelta(days=args.lookback_days)).date()
    trades = trades[(trades["date"] >= min_date) & (trades["date"] <= max_date)]

    sessions = {s.strip().upper() for s in args.sessions.split(",") if s.strip()}
    if sessions:
        trades = trades[trades["session"].str.upper().isin(sessions)]

    top_codes = (
        trades["code"].value_counts().head(args.top_n).index.tolist()
        if args.top_n > 0
        else trades["code"].unique()
    )
    trades = trades[trades["code"].isin(top_codes)].reset_index(drop=True)

    start_ts = pd.Timestamp(min(trades["ts"])).tz_convert("UTC") - timedelta(days=2)
    end_ts = pd.Timestamp(max(trades["ts_exit"])).tz_convert("UTC") + timedelta(days=2)

    for driver, cfg in DRIVERS.items():
        symbol = cfg["symbol"]
        intraday = fetch_intraday_prices(symbol, start_ts, end_ts)
        if driver == "NKY_F":
            daily = build_daily_from_intraday(
                intraday,
                open_time=time(hour=8, minute=45),
                close_time=time(hour=15, minute=15),
            )
        else:
            daily = fetch_daily_prices(symbol, start_ts, end_ts)
        attach_day_trend(trades, daily, driver, args.dir_threshold_bp)
        attach_day_alignment(trades, driver)
        attach_window_trend(trades, intraday, driver, args.dir_threshold_bp)

    alignment_df = summarize_alignment(trades, args.min_count)
    alignment_path = args.output_dir / "trend_alignment_summary.csv"
    alignment_df.to_csv(alignment_path, index=False)
    improvement_path = args.output_dir / "trend_alignment_improvement.csv"
    alignment_df[alignment_df["policy"] == "aligned"].to_csv(improvement_path, index=False)

    ticker_pref_df = summarize_ticker_preferences(trades, args.min_count)
    ticker_pref_path = args.output_dir / "trend_ticker_preference.csv"
    ticker_pref_df.to_csv(ticker_pref_path, index=False)

    overview = {
        "lookback_days": args.lookback_days,
        "top_codes": len(top_codes),
        "sessions": sorted(sessions),
        "dir_threshold_bp": args.dir_threshold_bp,
        "sample_trades": len(trades),
        "driver_counts": {
            driver: {
                "day_missing": int((trades[f"{driver}_day_trend"] == "missing").sum()),
                "window_missing": int((trades[f"{driver}_window_trend"] == "missing").sum()),
            }
            for driver in DRIVERS
        },
    }
    overview_path = args.output_dir / "trend_driver_overview.json"
    overview_path.write_text(json.dumps(overview, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"wrote {alignment_path}, {improvement_path}, {ticker_pref_path}, {overview_path}")


if __name__ == "__main__":
    main()
