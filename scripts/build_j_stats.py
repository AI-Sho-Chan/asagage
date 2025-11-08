#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import datetime as dt

import numpy as np
import pandas as pd


def select_trend_series(df: pd.DataFrame) -> pd.Series:
    for col in (
        "driver_window_trend",
        "trend_window",
        "driver_day_trend",
        "NKY_window_trend",
        "nky_window_trend",
    ):
        if col in df.columns:
            return df[col].astype(str).str.lower()
    if len(df) == 0:
        return pd.Series(dtype=str)
    return pd.Series(["flat"] * len(df))


def build_stats(
    ledger_path: Path,
    output: Path,
    min_count: int,
    rules_path: Path,
    flat_quantile: float,
    trend_quantile: float,
    sigma_floor: float,
    log_dir: Path | None,
    log_days: int,
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

    ratio_frames: list[pd.DataFrame] = []
    ledger_frame = data[["code", "session", "ratio"]].copy()
    ledger_frame["trend_label"] = select_trend_series(data)
    ledger_frame["source"] = "ledger"
    ratio_frames.append(ledger_frame)

    log_frame = load_dashboard_ratios(log_dir, log_days)
    if log_frame is not None and not log_frame.empty:
        ratio_frames.append(log_frame)

    combined = pd.concat(ratio_frames, ignore_index=True)
    combined = combined.replace([np.inf, -np.inf], np.nan).dropna(subset=["ratio"])

    grouped = (
        combined.groupby(["code", "session"])["ratio"]
        .agg(count="count", ratio_mu="mean", ratio_sigma="std")
        .reset_index()
    )
    grouped = grouped[grouped["count"] >= max(1, min_count)]
    grouped["ratio_sigma"] = grouped["ratio_sigma"].fillna(sigma_floor)

    output.parent.mkdir(parents=True, exist_ok=True)
    grouped.to_csv(output, index=False)
    print(f"wrote {output} ({len(grouped)} rows)")

    update_bb_rules(
        dataset=combined,
        rules_path=rules_path,
        flat_quantile=flat_quantile,
        trend_quantile=trend_quantile,
        sigma_floor=sigma_floor,
    )


def update_bb_rules(
    dataset: pd.DataFrame,
    rules_path: Path,
    flat_quantile: float,
    trend_quantile: float,
    sigma_floor: float,
) -> None:
    if not rules_path.exists():
        return
    if dataset.empty:
        return
    if "trend_label" not in dataset.columns:
        return

    trend_series = dataset["trend_label"].astype(str).str.lower()
    ratio_series = dataset["ratio"]

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


def load_dashboard_ratios(log_dir: Path | None, max_days: int) -> pd.DataFrame:
    if log_dir is None or not log_dir.exists():
        return pd.DataFrame()
    rows: list[pd.DataFrame] = []
    cutoff_date: dt.date | None = None
    if max_days > 0:
        cutoff_date = dt.date.today() - dt.timedelta(days=max_days)
    for day_dir in sorted(log_dir.iterdir()):
        if not day_dir.is_dir():
            continue
        try:
            day_val = dt.datetime.strptime(day_dir.name, "%Y%m%d").date()
        except ValueError:
            continue
        if cutoff_date and day_val < cutoff_date:
            continue
        for csv_path in sorted(day_dir.glob("dashboard_j_*.csv")):
            try:
                df = pd.read_csv(csv_path)
            except Exception:
                continue
            if not {"Ticker", "J", "J_th"}.issubset(df.columns):
                continue
            ticker_series = df["Ticker"].astype(str).str.upper().str.strip()
            if "session" in df.columns:
                session_series = df["session"]
            else:
                session_series = pd.Series([""] * len(df))
            session_series = session_series.astype(str).str.upper().str.strip()
            j_vals = pd.to_numeric(df["J"], errors="coerce")
            jth_vals = pd.to_numeric(df["J_th"], errors="coerce")
            ratio = np.where(np.abs(jth_vals) > 0, np.abs(j_vals) / np.abs(jth_vals), np.nan)
            frame = pd.DataFrame(
                {
                    "code": ticker_series,
                    "session": session_series,
                    "ratio": ratio,
                    "trend_label": select_trend_series(df),
                    "source": "dashboard",
                }
            )
            rows.append(frame)
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="Build per-ticker J ratio statistics.")
    ap.add_argument("--ledger", type=Path, default=Path("analysis/all_trades_snapshot.csv"))
    ap.add_argument("--output", type=Path, default=Path("state/j_stats.csv"))
    ap.add_argument("--min-count", type=int, default=12, help="Minimum samples per (code,session)")
    ap.add_argument("--rules", type=Path, default=Path("state/strategy_rules.ini"))
    ap.add_argument("--target-flat-quantile", type=float, default=0.80)
    ap.add_argument("--target-trend-quantile", type=float, default=0.90)
    ap.add_argument("--sigma-floor", type=float, default=0.05)
    ap.add_argument("--logs", type=Path, default=Path("output/j_logs"))
    ap.add_argument("--log-days", type=int, default=7)
    args = ap.parse_args()
    build_stats(
        ledger_path=args.ledger,
        output=args.output,
        min_count=args.min_count,
        rules_path=args.rules,
        flat_quantile=args.target_flat_quantile,
        trend_quantile=args.target_trend_quantile,
        sigma_floor=args.sigma_floor,
        log_dir=args.logs,
        log_days=args.log_days,
    )


if __name__ == "__main__":
    main()
