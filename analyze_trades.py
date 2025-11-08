#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import math
from datetime import timedelta
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import numpy as np
import pandas as pd
import yfinance as yf

BT_ROOT = Path("output/bt30")
OUTPUT_DIR = Path("analysis")
STATE_RULES_PATH = Path("state/strategy_rules.ini")
SYMBOL = "^N225"
TZ = "Asia/Tokyo"
DIR_THRESHOLD_BP = 5.0


def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        new_cols = []
        for col in df.columns:
            parts = [str(part).strip() for part in col if str(part).strip()]
            new_cols.append("_".join(parts))
        df.columns = new_cols
    return df


def load_ledger_frames(root: Path) -> pd.DataFrame:
    ledger_paths = sorted(root.rglob("hypothesis_trade_ledger.csv"))
    if not ledger_paths:
        raise SystemExit(f"No hypothesis_trade_ledger.csv files found under {root}")

    frames: List[pd.DataFrame] = []
    for path in ledger_paths:
        rel = path.relative_to(root)
        parts = rel.parts
        if parts[0] == "WEEKLY_":
            batch_folder = parts[1] if len(parts) > 1 else "UNKNOWN"
            batch_kind = "weekly"
        else:
            batch_folder = parts[0]
            batch_kind = "nightly"
        df_part = pd.read_csv(path)
        df_part["batch_folder"] = batch_folder
        df_part["batch_kind"] = batch_kind
        df_part["run_path"] = str(rel.parent)
        frames.append(df_part)

    df = pd.concat(frames, ignore_index=True)
    df["date"] = pd.to_datetime(df["date"]).dt.date
    df["ts"] = pd.to_datetime(df["ts"], utc=True)
    df["bars"] = pd.to_numeric(df["bars"], errors="coerce").fillna(0)
    df["ts_exit"] = df["ts"] + pd.to_timedelta(df["bars"], unit="m")
    df = df.dropna(subset=["pnl_bp"])
    df["win"] = df["pnl_bp"] > 0
    return df


def fetch_daily_prices(symbol: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.DataFrame:
    data = yf.download(
        symbol,
        start=start.strftime("%Y-%m-%d"),
        end=end.strftime("%Y-%m-%d"),
        interval="1d",
        auto_adjust=False,
        progress=False,
    )
    if data.empty:
        return data
    data = flatten_columns(data)
    data.index = pd.to_datetime(data.index)
    data["date"] = data.index.date
    columns = {col: col.split("_")[0].lower() for col in data.columns}
    data = data.rename(columns=columns)
    return data.set_index("date")


def fetch_intraday_prices(symbol: str, start: pd.Timestamp, end: pd.Timestamp) -> pd.Series:
    data = yf.download(
        symbol,
        start=start.strftime("%Y-%m-%d"),
        end=end.strftime("%Y-%m-%d"),
        interval="5m",
        auto_adjust=False,
        progress=False,
    )
    if data.empty:
        return pd.Series(dtype=float)
    data = flatten_columns(data)
    idx = pd.to_datetime(data.index)
    if idx.tz is None:
        idx = idx.tz_localize("UTC")
    columns = {col: col.split("_")[0] for col in data.columns}
    data = data.rename(columns=columns)
    close = pd.Series(data["Close"].values, index=idx, name="Close")
    close = close.tz_convert(TZ)
    return close


def classify_direction(ret_bp: float) -> str:
    if pd.isna(ret_bp):
        return "missing"
    if ret_bp > DIR_THRESHOLD_BP:
        return "up"
    if ret_bp < -DIR_THRESHOLD_BP:
        return "down"
    return "flat"


def attach_day_trend(df: pd.DataFrame, daily: pd.DataFrame) -> None:
    if daily.empty:
        df["nky_day_ret_bp"] = np.nan
        df["nky_day_trend"] = "missing"
        return
    daily = daily.copy()
    daily["ret_bp"] = (daily["close"] - daily["open"]) / daily["open"] * 10000
    daily["trend"] = daily["ret_bp"].apply(classify_direction)
    df["nky_day_ret_bp"] = df["date"].map(daily["ret_bp"])
    df["nky_day_trend"] = df["date"].map(daily["trend"])


def attach_window_ret(df: pd.DataFrame, intraday_close: pd.Series) -> None:
    if intraday_close.empty:
        df["nky_window_ret_bp"] = np.nan
        df["nky_window_trend"] = "missing"
        df["nky_window_alignment"] = "missing"
        return

    start_px = intraday_close.reindex(df["ts"], method="ffill").to_numpy()
    end_px = intraday_close.reindex(df["ts_exit"], method="ffill").to_numpy()
    valid = (start_px > 0) & (end_px > 0)
    window_ret = np.full(len(df), np.nan)
    window_ret[valid] = (end_px[valid] - start_px[valid]) / start_px[valid] * 10000
    df["nky_window_ret_bp"] = window_ret
    df["nky_window_trend"] = df["nky_window_ret_bp"].apply(classify_direction)

    def label_alignment(row: pd.Series) -> str:
        trend = row["nky_window_trend"]
        side = row["side"]
        if pd.isna(side) or trend in {"missing", "flat"}:
            return trend
        if (trend == "up" and side == "BUY") or (trend == "down" and side == "SELL"):
            return "aligned"
        return "counter"

    df["nky_window_alignment"] = df.apply(label_alignment, axis=1)


def build_summary(sub: pd.DataFrame) -> Dict[str, float]:
    n = len(sub)
    if n == 0:
        return {
            "count": 0,
            "mean": np.nan,
            "std": np.nan,
            "median": np.nan,
            "win_rate": np.nan,
            "se_mean": np.nan,
            "ci95_low": np.nan,
            "ci95_high": np.nan,
        }
    mean = sub["pnl_bp"].mean()
    std = sub["pnl_bp"].std(ddof=1)
    median = sub["pnl_bp"].median()
    win_rate = sub["win"].mean()
    se_mean = std / math.sqrt(n) if n > 1 else float("nan")
    ci95_low = mean - 1.96 * se_mean if n > 1 else float("nan")
    ci95_high = mean + 1.96 * se_mean if n > 1 else float("nan")
    return {
        "count": int(n),
        "mean": mean,
        "std": std,
        "median": median,
        "win_rate": win_rate,
        "se_mean": se_mean,
        "ci95_low": ci95_low,
        "ci95_high": ci95_high,
    }


def mean_diff_stats(a: pd.DataFrame, b: pd.DataFrame) -> Dict[str, float]:
    na, nb = len(a), len(b)
    mean_a, mean_b = a["pnl_bp"].mean(), b["pnl_bp"].mean()
    var_a, var_b = a["pnl_bp"].var(ddof=1), b["pnl_bp"].var(ddof=1)
    if na > 1 and nb > 1:
        se = math.sqrt(var_a / na + var_b / nb)
        z = (mean_b - mean_a) / se if se > 0 else float("inf")
        p = math.erfc(abs(z) / math.sqrt(2))
    else:
        z = float("nan")
        p = float("nan")
    return {
        "mean_diff_sell_minus_buy": mean_b - mean_a,
        "z_score": z,
        "p_value_approx": p,
    }


def summarize_group(df: pd.DataFrame, by: Iterable[str], label: str) -> pd.DataFrame:
    rows: List[Dict[str, float]] = []
    for keys, subset in df.groupby(list(by)):
        if not isinstance(keys, tuple):
            keys = (keys,)
        summary = build_summary(subset)
        for name, value in zip(by, keys):
            summary[name] = value
        rows.append(summary)
    return pd.DataFrame(rows)


def summarize_tp_sl(df: pd.DataFrame) -> pd.DataFrame:
    records: List[Dict[str, float]] = []
    for side, subset in df.groupby("side"):
        tp_med = subset["TPk"].median()
        sl_med = subset["SLk"].median()
        ratio = (subset["TPk"] / subset["SLk"]).replace([np.inf, -np.inf], np.nan).median()
        records.append(
            {
                "side": side,
                "median_TPk": tp_med,
                "median_SLk": sl_med,
                "median_TP_over_SL": ratio,
            }
        )
    return pd.DataFrame(records)


def save_json_overview(df: pd.DataFrame, side_groups: Dict[str, Dict[str, float]], diff_stats: Dict[str, float]) -> None:
    overview = {
        "overall_side": side_groups,
        "diff_stats": diff_stats,
        "nky_day_trend_counts": df["nky_day_trend"].value_counts().to_dict(),
        "nky_window_trend_counts": df["nky_window_trend"].value_counts().to_dict(),
        "batch_counts": [
            {"batch_kind": k[0], "side": k[1], "count": int(v)}
            for k, v in df.groupby(["batch_kind", "side"]).size().to_dict().items()
        ],
    }
    with (OUTPUT_DIR / "overview_stats.json").open("w", encoding="utf-8") as f:
        json.dump(overview, f, ensure_ascii=False, indent=2)


def write_strategy_rules(trades: pd.DataFrame, batch_side_df: pd.DataFrame, session_mode_df: pd.DataFrame) -> None:
    STATE_RULES_PATH.parent.mkdir(exist_ok=True, parents=True)

    def _mean_for(df: pd.DataFrame, **filters: str) -> float | None:
        if df.empty or "mean" not in df.columns:
            return None
        sel = df
        for col, target in filters.items():
            sel = sel[sel[col].astype(str).str.upper() == target.upper()]
        if sel.empty:
            return None
        value = pd.to_numeric(sel["mean"], errors="coerce").dropna()
        if value.empty:
            return None
        return float(value.iloc[0])

    weekly_sell_mean = _mean_for(batch_side_df, batch_kind="weekly", side="SELL")
    weekly_sell_rule = "disable" if (weekly_sell_mean is None or weekly_sell_mean < 30) else "allow"

    jcross_sell_mean = _mean_for(session_mode_df, session="AM15", signal_mode="j-cross", side="SELL")
    jcross_rule = "1" if (jcross_sell_mean is None or jcross_sell_mean < 35) else "0"

    lines = [
        f"weekly_sell={weekly_sell_rule}",
        f"jcross_sell_require_nky_down={jcross_rule}",
        "jcross_sell_min_gap_bp=20",
        "nky_initial_bp=10",
        "nky_steady_bp=15",
        "alert_cooldown_min=10",
        "bb_flat_k=1",
        "bb_trend_k=1.3",
        "bb_min_samples=12",
        "bb_sigma_floor=0.05",
    ]
    STATE_RULES_PATH.write_text("\n".join(lines) + "\n", encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Aggregate trade ledgers and NKY trends.")
    parser.add_argument("--bt-root", type=Path, default=BT_ROOT)
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument("--symbol", default=SYMBOL)
    parser.add_argument("--dir-threshold", type=float, default=DIR_THRESHOLD_BP)
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    global DIR_THRESHOLD_BP, OUTPUT_DIR
    DIR_THRESHOLD_BP = args.dir_threshold
    OUTPUT_DIR = args.output_dir
    OUTPUT_DIR.mkdir(exist_ok=True, parents=True)

    trades = load_ledger_frames(args.bt_root)
    trades.to_csv(OUTPUT_DIR / "all_trades_snapshot.csv", index=False)

    start_date = pd.Timestamp(min(trades["date"])) - timedelta(days=2)
    end_date = pd.Timestamp(max(trades["date"])) + timedelta(days=2)
    daily = fetch_daily_prices(args.symbol, start_date, end_date)
    attach_day_trend(trades, daily)

    intraday_close = fetch_intraday_prices(args.symbol, start_date, end_date)
    attach_window_ret(trades, intraday_close)

    side_groups = {
        side: build_summary(trades[trades["side"] == side])
        for side in sorted(trades["side"].dropna().unique())
    }
    diff_stats = mean_diff_stats(
        trades[trades["side"] == "BUY"], trades[trades["side"] == "SELL"]
    )

    batch_side_df = summarize_group(trades, ["batch_kind", "side"], "batch_kind")
    day_trend_df = summarize_group(trades, ["nky_day_trend", "side"], "nky_day_trend")
    window_trend_df = summarize_group(trades, ["nky_window_trend", "side"], "nky_window_trend")
    window_align_df = summarize_group(
        trades, ["nky_window_alignment", "side"], "nky_window_alignment"
    )
    session_mode_df = summarize_group(trades, ["session", "signal_mode", "side"], "session")
    tp_sl_df = summarize_tp_sl(trades)

    batch_side_df.to_csv(OUTPUT_DIR / "batch_side_summary.csv", index=False)
    day_trend_df.to_csv(OUTPUT_DIR / "day_trend_summary.csv", index=False)
    window_trend_df.to_csv(OUTPUT_DIR / "window_trend_summary.csv", index=False)
    window_align_df.to_csv(OUTPUT_DIR / "window_alignment_summary.csv", index=False)
    session_mode_df.to_csv(OUTPUT_DIR / "session_mode_summary.csv", index=False)
    tp_sl_df.to_csv(OUTPUT_DIR / "tp_sl_ratio_summary.csv", index=False)
    am15 = pd.DataFrame()
    if "session" in trades.columns:
        session_upper = trades["session"].astype(str).str.upper()
        am15_mask = session_upper == "AM15"
        am15 = trades[am15_mask]
    if not am15.empty:
        am15_summary = (
            am15.groupby(["nky_window_alignment", "side"])["pnl_bp"]
            .agg(count="count", mean="mean", median="median")
            .reset_index()
        )
        am15_summary.to_csv(OUTPUT_DIR / "am15_alignment_summary.csv", index=False)

    save_json_overview(trades, side_groups, diff_stats)
    write_strategy_rules(trades, batch_side_df, session_mode_df)
    print("analysis complete -> analysis/*.csv, overview_stats.json")


if __name__ == "__main__":
    main()
