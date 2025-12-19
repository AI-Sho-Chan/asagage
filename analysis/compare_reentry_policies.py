#!/usr/bin/env python3
"""
Compare DailyReplay outcomes under different "same-ticker re-entry" policies.

Goal (non-technical):
  See whether relaxing the "max trades per ticker per day" rule changes results,
  using the same candidate set across multiple past days.

This script:
  - Uses the current candidates CSV (default: output/excel/candidates_nextday.csv)
  - Picks the most recent N days with enough local 1-minute data coverage
  - Runs DailyReplay twice per day:
      A) Baseline: max 2 trades per ticker/day
      B) Proposed: no max trade cap, but once a ticker loses, stop trading it that day
  - Writes: analysis/reentry_policy_compare.csv and analysis/reentry_policy_summary.json
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import importlib.util
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import pandas as pd

_ROOT = Path(__file__).resolve().parents[1]
_SIM_PATH = _ROOT / "tools" / "simulate_daily_replay.py"
_spec = importlib.util.spec_from_file_location("simulate_daily_replay", _SIM_PATH)
if _spec is None or _spec.loader is None:
    raise SystemExit(f"failed to load module spec: {_SIM_PATH}")
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)  # type: ignore[attr-defined]

DATA_ROOT = _mod.DATA_ROOT
load_candidates = _mod.load_candidates
simulate_day = _mod.simulate_day
normalize_columns = _mod.normalize_columns


def _parse_ymd(stem: str) -> dt.date | None:
    try:
        return dt.datetime.strptime(stem, "%Y-%m-%d").date()
    except ValueError:
        return None


def list_local_days_for_code(code: str) -> List[dt.date]:
    root = DATA_ROOT / code
    if not root.exists():
        return []
    days: List[dt.date] = []
    for p in root.glob("????-??-??.parquet"):
        d = _parse_ymd(p.stem)
        if d is not None:
            days.append(d)
    return sorted(set(days))


def pick_days(
    *,
    tickers: List[str],
    days: int,
    min_coverage: float,
    max_lookback_days: int,
) -> List[dt.date]:
    today = dt.date.today()
    earliest = today - dt.timedelta(days=max_lookback_days)

    counts: Dict[dt.date, int] = {}
    for code in tickers:
        for d in list_local_days_for_code(code):
            if d < earliest or d > today:
                continue
            counts[d] = counts.get(d, 0) + 1

    if not counts:
        return []

    total = max(1, len(tickers))
    ranked = sorted(counts.items(), key=lambda kv: kv[0], reverse=True)
    picked: List[dt.date] = []
    for d, c in ranked:
        if c / total >= min_coverage:
            picked.append(d)
        if len(picked) >= days:
            break
    return picked


def run_policy(
    cand_path: Path,
    trading_day: dt.date,
    *,
    nominal: float,
    cooldown_minutes: int,
    max_trades_per_ticker: int,
    stop_after_loss: bool,
) -> Tuple[pd.DataFrame, Dict[str, object]]:
    return simulate_day(
        cand_path,
        trading_day,
        nominal,
        cooldown_minutes=cooldown_minutes,
        max_trades_per_ticker=max_trades_per_ticker,
        stop_after_loss=stop_after_loss,
    )


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Compare re-entry policies using DailyReplay.")
    ap.add_argument(
        "--candidates",
        type=Path,
        default=Path("output/excel/candidates_nextday.csv"),
        help="Candidates CSV to use for all days (default: output/excel/candidates_nextday.csv)",
    )
    ap.add_argument("--days", type=int, default=20, help="Number of past days to test (default 20)")
    ap.add_argument(
        "--min-coverage",
        type=float,
        default=0.6,
        help="Min fraction of tickers with local 1m data for the day to be included (default 0.6)",
    )
    ap.add_argument(
        "--max-lookback-days",
        type=int,
        default=120,
        help="Search window for available days (default 120 calendar days)",
    )
    ap.add_argument("--nominal", type=float, default=10_000_000.0, help="Budget per plan in yen (default 10,000,000)")
    ap.add_argument("--cooldown-minutes", type=int, default=5, help="Cooldown minutes (default 5)")
    ap.add_argument(
        "--baseline-max-trades",
        type=int,
        default=2,
        help="Baseline max trades per ticker/day (default 2)",
    )
    ap.add_argument(
        "--out-csv",
        type=Path,
        default=Path("analysis/reentry_policy_compare.csv"),
        help="Output CSV path",
    )
    ap.add_argument(
        "--out-json",
        type=Path,
        default=Path("analysis/reentry_policy_summary.json"),
        help="Output JSON summary path",
    )
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    cand_path = args.candidates
    if not cand_path.exists():
        raise SystemExit(f"candidates file not found: {cand_path}")

    cand_raw = load_candidates(cand_path)
    cand_df = normalize_columns(cand_raw)
    tickers = sorted({str(t).strip() for t in cand_df["code"].astype(str).tolist() if str(t).strip()})
    if not tickers:
        raise SystemExit("no tickers found in candidates")

    picked_days = pick_days(
        tickers=tickers,
        days=int(args.days),
        min_coverage=float(args.min_coverage),
        max_lookback_days=int(args.max_lookback_days),
    )
    if not picked_days:
        raise SystemExit(
            "no days found with enough local 1m data coverage; "
            "consider lowering --min-coverage or increasing --max-lookback-days"
        )

    rows: List[Dict[str, object]] = []
    deltas: List[Dict[str, object]] = []

    for d in picked_days:
        base_trades, base_sum = run_policy(
            cand_path,
            d,
            nominal=float(args.nominal),
            cooldown_minutes=int(args.cooldown_minutes),
            max_trades_per_ticker=int(args.baseline_max_trades),
            stop_after_loss=False,
        )
        alt_trades, alt_sum = run_policy(
            cand_path,
            d,
            nominal=float(args.nominal),
            cooldown_minutes=int(args.cooldown_minutes),
            max_trades_per_ticker=0,  # no cap
            stop_after_loss=True,
        )

        def _row(policy: str, summary: Dict[str, object]) -> Dict[str, object]:
            return {
                "date": d.strftime("%Y-%m-%d"),
                "policy": policy,
                "trades": int(summary.get("trades", 0) or 0),
                "pnl_yen": float(summary.get("pnl_yen", 0.0) or 0.0),
                "pnl_bp_mean": float(summary.get("pnl_bp_mean", 0.0) or 0.0),
                "cooldown_minutes": int(summary.get("cooldown_minutes", args.cooldown_minutes) or args.cooldown_minutes),
                "max_trades_per_ticker": int(summary.get("max_trades_per_ticker", 0) or 0),
                "stop_after_loss": bool(summary.get("stop_after_loss", False)),
                "diag_skip_trend_mismatch": int(summary.get("diag_skip_trend_mismatch", 0) or 0),
                "diag_no_signal": int(summary.get("diag_no_signal", 0) or 0),
                "diag_skip_gapban": int(summary.get("diag_skip_gapban", 0) or 0),
            }

        rows.append(_row("baseline_max2", base_sum))
        rows.append(_row("no_cap_stop_after_loss", alt_sum))

        if float(base_sum.get("pnl_yen", 0.0) or 0.0) != float(alt_sum.get("pnl_yen", 0.0) or 0.0) or int(
            base_sum.get("trades", 0) or 0
        ) != int(alt_sum.get("trades", 0) or 0):
            deltas.append(
                {
                    "date": d.strftime("%Y-%m-%d"),
                    "baseline_trades": int(base_sum.get("trades", 0) or 0),
                    "alt_trades": int(alt_sum.get("trades", 0) or 0),
                    "baseline_pnl_yen": float(base_sum.get("pnl_yen", 0.0) or 0.0),
                    "alt_pnl_yen": float(alt_sum.get("pnl_yen", 0.0) or 0.0),
                    "delta_pnl_yen": float(alt_sum.get("pnl_yen", 0.0) or 0.0)
                    - float(base_sum.get("pnl_yen", 0.0) or 0.0),
                }
            )

    out_df = pd.DataFrame(rows).sort_values(["date", "policy"])
    args.out_csv.parent.mkdir(parents=True, exist_ok=True)
    out_df.to_csv(args.out_csv, index=False, encoding="utf-8-sig")

    summary_df = out_df.groupby("policy", as_index=False).agg(
        days=("date", "nunique"),
        trades=("trades", "sum"),
        pnl_yen=("pnl_yen", "sum"),
        pnl_yen_mean=("pnl_yen", "mean"),
        pnl_yen_median=("pnl_yen", "median"),
    )
    summary = {
        "candidates": str(cand_path),
        "tested_days": [d.strftime("%Y-%m-%d") for d in picked_days],
        "unique_tickers": len(tickers),
        "days_requested": int(args.days),
        "min_coverage": float(args.min_coverage),
        "cooldown_minutes": int(args.cooldown_minutes),
        "baseline_max_trades": int(args.baseline_max_trades),
        "summary_by_policy": summary_df.to_dict(orient="records"),
        "days_with_any_difference": len(deltas),
        "differences": deltas,
    }
    args.out_json.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"wrote {args.out_csv}")
    print(f"wrote {args.out_json}")
    print("summary:")
    for row in summary["summary_by_policy"]:
        print(
            f"  {row['policy']}: days={row['days']} trades={row['trades']} pnl_yen={row['pnl_yen']:.0f} "
            f"(mean/day={row['pnl_yen_mean']:.0f}, median/day={row['pnl_yen_median']:.0f})"
        )
    print(f"days with any difference: {len(deltas)}")


if __name__ == "__main__":
    main()
