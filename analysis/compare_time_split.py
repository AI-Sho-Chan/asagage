#!/usr/bin/env python3
"""
Compare "time-split" (M3) vs "no split" (M0) using the same candidate list.

This is a quick experiment to answer:
  - If we remove the per-session time window and allow entries all day,
    does the realized PnL (replay on 1m data) get better or worse?

Common replay rules (user-approved baseline):
  - One position per ticker at a time
  - Re-entry allowed only after exit
  - Cooldown: 5 minutes
  - Max trades per ticker per day: 2

Notes:
  - This script is for comparison, not production trading.
  - It uses the same candidates for all dates, so it is *not* a
    strict "no-lookahead" evaluation.
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable, List, Tuple

import numpy as np
import pandas as pd

import sys

# Ensure repo root is on sys.path so we can import `tools/*` when executed as a script.
REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

from tools.simulate_daily_replay import simulate_day


DATA_ROOT = Path("data/raw/yahoo_1m")
OUT_DIR = Path("analysis")


def _iter_available_dates(code: str) -> List[dt.date]:
    code_dir = DATA_ROOT / code
    if not code_dir.exists():
        return []
    dates: List[dt.date] = []
    for p in sorted(code_dir.glob("*.parquet")):
        try:
            dates.append(dt.date.fromisoformat(p.stem))
        except ValueError:
            continue
    return dates


def _as_yyyymmdd(day: dt.date) -> str:
    return day.strftime("%Y%m%d")


def _bootstrap_ci(values: np.ndarray, n: int = 5000, seed: int = 1) -> Tuple[float, float]:
    if values.size == 0:
        return (float("nan"), float("nan"))
    rng = np.random.default_rng(seed)
    means = np.empty(n, dtype=float)
    for i in range(n):
        sample = rng.choice(values, size=values.size, replace=True)
        means[i] = float(np.mean(sample))
    lo = float(np.quantile(means, 0.025))
    hi = float(np.quantile(means, 0.975))
    return (lo, hi)


def _make_m0_candidates(df_m3: pd.DataFrame) -> pd.DataFrame:
    df = df_m3.copy()
    if "session" in df.columns:
        df["session"] = ""
    return df


def run_compare(
    candidates_path: Path,
    dates: Iterable[dt.date],
    nominal: float,
    out_csv: Path,
) -> pd.DataFrame:
    df_m3 = pd.read_csv(candidates_path)
    df_m0 = _make_m0_candidates(df_m3)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    tmp_m0 = OUT_DIR / "tmp_candidates_m0.csv"
    df_m0.to_csv(tmp_m0, index=False, encoding="utf-8-sig")

    rows = []
    for day in dates:
        date_tag = _as_yyyymmdd(day)
        try:
            _, s3 = simulate_day(candidates_path, day, nominal)
        except Exception as exc:  # pragma: no cover
            rows.append({"date_tag": date_tag, "variant": "M3", "error": repr(exc)})
            continue

        try:
            _, s0 = simulate_day(tmp_m0, day, nominal)
        except Exception as exc:  # pragma: no cover
            rows.append({"date_tag": date_tag, "variant": "M0", "error": repr(exc)})
            continue

        rows.append(
            {
                "date_tag": date_tag,
                "pnl_yen_M3": float(s3.get("pnl_yen", 0.0)),
                "trades_M3": int(s3.get("trades", 0)),
                "pnl_yen_M0": float(s0.get("pnl_yen", 0.0)),
                "trades_M0": int(s0.get("trades", 0)),
            }
        )

    out = pd.DataFrame(rows)
    if not out.empty and "pnl_yen_M3" in out.columns and "pnl_yen_M0" in out.columns:
        out["diff_M3_minus_M0"] = out["pnl_yen_M3"] - out["pnl_yen_M0"]
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(out_csv, index=False, encoding="utf-8-sig")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="Compare M3 (time-split) vs M0 (no split).")
    ap.add_argument(
        "--candidates",
        type=Path,
        default=Path("output/excel/candidates_nextday.csv"),
        help="Candidates CSV (default: output/excel/candidates_nextday.csv)",
    )
    ap.add_argument("--nominal", type=float, default=10_000_000.0)
    ap.add_argument("--days", type=int, default=20, help="How many recent dates to compare (default 20).")
    ap.add_argument(
        "--anchor-code",
        default="",
        help="Ticker to pick available dates from (default: first ticker in candidates).",
    )
    ap.add_argument(
        "--out",
        type=Path,
        default=Path("analysis/time_split_comparison.csv"),
        help="Output CSV path",
    )
    args = ap.parse_args()

    if not args.candidates.exists():
        raise SystemExit(f"Candidates not found: {args.candidates}")

    df = pd.read_csv(args.candidates)
    if df.empty:
        raise SystemExit(f"Candidates is empty: {args.candidates}")

    code_col = "Ticker" if "Ticker" in df.columns else ("code" if "code" in df.columns else None)
    if not code_col:
        raise SystemExit("Candidates must contain Ticker or code column.")

    anchor = str(args.anchor_code).strip() or str(df[code_col].iloc[0]).strip()
    available = _iter_available_dates(anchor)
    if not available:
        raise SystemExit(f"No 1m parquet dates found for {anchor} under {DATA_ROOT}.")

    dates = available[-int(args.days) :]
    out = run_compare(args.candidates, dates, args.nominal, args.out)

    ok = out.dropna(subset=["diff_M3_minus_M0"]) if "diff_M3_minus_M0" in out.columns else pd.DataFrame()
    if ok.empty:
        print(f"Wrote {args.out} (no comparable rows)")
        return

    diffs = ok["diff_M3_minus_M0"].to_numpy(dtype=float)
    mean = float(np.mean(diffs))
    med = float(np.median(diffs))
    lo, hi = _bootstrap_ci(diffs, n=5000, seed=1)
    win = int(np.sum(diffs > 0))
    loss = int(np.sum(diffs < 0))
    tie = int(np.sum(diffs == 0))

    print(f"Wrote {args.out} (rows={len(ok)})")
    print(f"diff (M3 - M0) mean={mean:.0f} yen, median={med:.0f} yen, 95%CI=[{lo:.0f}, {hi:.0f}]")
    print(f"days M3 better/worse/tie: {win}/{loss}/{tie}")


if __name__ == "__main__":
    main()
