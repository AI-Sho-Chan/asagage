import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd

ROOT = Path("output/excel")
OUT_DIR = Path("analysis")


@dataclass(frozen=True)
class TradeRules:
    max_trades_per_ticker_per_day: int
    cooldown_minutes: int
    one_position_per_ticker: bool


@dataclass(frozen=True)
class Filters:
    min_j_th: float
    min_forward_pf: float
    min_forward_trades: int
    min_win_ci_low: float


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description=(
            "Compare time segmentation methods (M0..M3) using existing NIGHTLY candidate CSVs."
        )
    )
    ap.add_argument(
        "--root",
        type=Path,
        default=ROOT,
        help="Root folder that contains NIGHTLY_YYYYMMDD directories (default: output/excel).",
    )
    ap.add_argument(
        "--date-tags",
        nargs="*",
        default=[],
        help="Optional list of YYYYMMDD. When omitted, auto-detects under root.",
    )
    ap.add_argument(
        "--min-date",
        type=str,
        default="",
        help="Optional minimum YYYYMMDD (inclusive) when auto-detecting date tags.",
    )
    ap.add_argument(
        "--max-date",
        type=str,
        default="",
        help="Optional maximum YYYYMMDD (inclusive) when auto-detecting date tags.",
    )
    ap.add_argument("--out-prefix", type=str, default="method_cmp")

    ap.add_argument("--min-j-th", type=float, default=0.8)
    ap.add_argument("--min-forward-pf", type=float, default=1.3)
    ap.add_argument("--min-forward-trades", type=int, default=5)
    ap.add_argument("--min-win-ci-low", type=float, default=0.65)

    ap.add_argument(
        "--max-trades-per-ticker-per-day",
        type=int,
        default=2,
        help="Comparison rule: cap selected combos per ticker per day (default: 2).",
    )
    ap.add_argument(
        "--cooldown-minutes",
        type=int,
        default=5,
        help="Comparison rule: cooldown minutes (documented only; selection cap approximates it).",
    )
    ap.add_argument(
        "--bootstrap-iters",
        type=int,
        default=5000,
        help="Number of bootstrap resamples for day-level paired differences.",
    )
    ap.add_argument("--seed", type=int, default=7)
    return ap.parse_args()


def iter_date_tags(root: Path, min_date: str, max_date: str) -> List[str]:
    date_tags: List[str] = []
    for path in root.glob("NIGHTLY_*"):
        m = re.match(r"^NIGHTLY_(\d{8})$", path.name)
        if not m:
            continue
        tag = m.group(1)
        if min_date and tag < min_date:
            continue
        if max_date and tag > max_date:
            continue
        date_tags.append(tag)
    return sorted(date_tags)


def _read_csv(path: Path) -> Optional[pd.DataFrame]:
    try:
        return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        try:
            return pd.read_csv(path)
        except Exception:
            return None


def collect_plan_frames(root: Path, date_tag: str) -> List[pd.DataFrame]:
    nightly = root / f"NIGHTLY_{date_tag}"
    frames: List[pd.DataFrame] = []
    if not nightly.exists():
        return frames

    for plan_dir in sorted([p for p in nightly.iterdir() if p.is_dir()]):
        # Prefer a file that matches the date, otherwise fall back to newest candidates_*.csv.
        direct = plan_dir / f"candidates_{date_tag}.csv"
        candidate_files: List[Path] = []
        if direct.exists():
            candidate_files = [direct]
        else:
            candidate_files = sorted(plan_dir.glob("candidates_*.csv"), key=lambda p: p.stat().st_mtime)
            if candidate_files:
                candidate_files = [candidate_files[-1]]

        for path in candidate_files:
            df = _read_csv(path)
            if df is None or df.empty:
                continue
            df = df.copy()
            df["date_tag"] = date_tag
            df["plan_tag"] = plan_dir.name
            frames.append(df)
    return frames


def _colmap(df: pd.DataFrame) -> Dict[str, str]:
    return {c.lower(): c for c in df.columns}


def _col(cols: Dict[str, str], name: str) -> Optional[str]:
    return cols.get(name.lower())


def _num(df: pd.DataFrame, col: Optional[str]) -> pd.Series:
    if not col or col not in df.columns:
        return pd.Series(np.nan, index=df.index)
    return pd.to_numeric(df[col], errors="coerce")


def compute_score(df: pd.DataFrame) -> pd.Series:
    cols = _colmap(df)
    pf = _num(df, _col(cols, "forward_pf_eff")).fillna(0.0).replace([np.inf, -np.inf], 0.0)
    win = _num(df, _col(cols, "forward_winrate")).fillna(0.0).replace([np.inf, -np.inf], 0.0)
    trades = _num(df, _col(cols, "forward_trades")).fillna(0.0).replace([np.inf, -np.inf], 0.0)
    dd = _num(df, _col(cols, "MaxDD")).fillna(0.0).replace([np.inf, -np.inf], 0.0)
    dd = dd.clip(lower=0.0)

    score = pf * (win.clip(0, 1) ** 1.2) * np.log1p(trades.clip(lower=0.0) + 1.0) / (
        1.0 + dd / 1000.0
    )

    if "CorrTOPIX" in df.columns:
        corr = pd.to_numeric(df["CorrTOPIX"], errors="coerce").fillna(0.0).clip(-1, 1)
        score = score * (1.0 + corr * 0.05)
    if "VWAP_revert_prob" in df.columns:
        vprob = pd.to_numeric(df["VWAP_revert_prob"], errors="coerce").fillna(0.0).clip(0, 1)
        score = score * (1.0 + vprob * 0.1)

    return score


def apply_quality_filters(df: pd.DataFrame, filters: Filters) -> pd.DataFrame:
    cols = _colmap(df)
    j_th = _num(df, _col(cols, "J_th"))
    pf = _num(df, _col(cols, "forward_pf_eff"))
    trades = _num(df, _col(cols, "forward_trades"))
    ci_low = _num(df, _col(cols, "forward_win_ci_low"))

    mask = pd.Series(True, index=df.index)
    if not j_th.isna().all():
        mask &= j_th.fillna(0.0) >= filters.min_j_th
    if not pf.isna().all():
        mask &= pf.fillna(0.0) >= filters.min_forward_pf
    if not trades.isna().all():
        mask &= trades.fillna(0.0) >= float(filters.min_forward_trades)
    if not ci_low.isna().all():
        mask &= ci_low.fillna(0.0) >= filters.min_win_ci_low

    return df[mask].copy()


def expected_bp(df: pd.DataFrame) -> pd.Series:
    cols = _colmap(df)
    col = _col(cols, "forward_exp_boot_mean") or _col(cols, "forward_exp_bp")
    series = _num(df, col).fillna(0.0).replace([np.inf, -np.inf], 0.0)
    return series


def session_from_plan(plan_tag: str) -> str:
    # e.g. AM0930_j-only -> AM0930
    return plan_tag.split("_", 1)[0] if plan_tag else ""


def bucket_m2(session: str) -> str:
    # Simple split: morning sessions start with "AM" or "MID", afternoon with "PM".
    if session.startswith("PM"):
        return "PM"
    if session.startswith("MID"):
        return "AM"
    return "AM" if session.startswith("AM") else "AM"


def build_method_selection(
    df: pd.DataFrame,
    *,
    method: str,
    max_per_ticker: int,
) -> pd.DataFrame:
    cols = _colmap(df)
    ticker_col = _col(cols, "Ticker") or _col(cols, "code") or _col(cols, "ticker")
    if not ticker_col or ticker_col not in df.columns:
        return df.iloc[0:0].copy()

    df = df.copy()
    df["_ticker"] = df[ticker_col].astype(str).str.upper()
    df["_session"] = df["plan_tag"].map(session_from_plan)
    df["_score"] = compute_score(df)
    df["_exp_bp"] = expected_bp(df)

    if method == "M0":
        # One combo per ticker.
        return (
            df.sort_values(["_ticker", "_score"], ascending=[True, False])
            .groupby("_ticker", as_index=False)
            .head(1)
            .drop(columns=["_ticker"])
        )

    if method == "M2":
        df["_bucket"] = df["_session"].map(bucket_m2)
        # Best per ticker per bucket (AM/PM), then cap total (should already be <=2).
        best = (
            df.sort_values(["_ticker", "_bucket", "_score"], ascending=[True, True, False])
            .groupby(["_ticker", "_bucket"], as_index=False)
            .head(1)
        )
        best = (
            best.sort_values(["_ticker", "_score"], ascending=[True, False])
            .groupby("_ticker", as_index=False)
            .head(max_per_ticker)
        )
        return best.drop(columns=["_ticker"])

    if method == "M1":
        # Open is noisy: drop AM15 sessions then behave like M3.
        filtered = df[~df["_session"].str.startswith("AM15")].copy()
        return build_method_selection(filtered, method="M3", max_per_ticker=max_per_ticker)

    if method == "M3":
        # Fine split (existing). For fair comparison, cap per ticker (max trades/day).
        return (
            df.sort_values(["_ticker", "_score"], ascending=[True, False])
            .groupby("_ticker", as_index=False)
            .head(max_per_ticker)
            .drop(columns=["_ticker"])
        )

    raise ValueError(f"Unknown method: {method}")


def per_ticker_summary(selected: pd.DataFrame) -> pd.DataFrame:
    if selected.empty:
        return pd.DataFrame(columns=["date_tag", "method", "Ticker", "exp_bp_sum", "combos"])
    cols = _colmap(selected)
    ticker_col = _col(cols, "Ticker") or _col(cols, "code") or _col(cols, "ticker")
    df = selected.copy()
    df["Ticker"] = df[ticker_col].astype(str).str.upper()
    df["exp_bp"] = df["_exp_bp"] if "_exp_bp" in df.columns else expected_bp(df)
    out = (
        df.groupby(["date_tag", "method", "Ticker"], as_index=False)
        .agg(exp_bp_sum=("exp_bp", "sum"), combos=("Ticker", "size"))
        .copy()
    )
    return out


def bootstrap_ci_of_mean(diff_values: np.ndarray, iters: int, rng: np.random.Generator) -> Tuple[float, float, float]:
    if len(diff_values) == 0:
        return (0.0, 0.0, 0.0)
    means = []
    n = len(diff_values)
    for _ in range(iters):
        sample = diff_values[rng.integers(0, n, size=n)]
        means.append(float(np.mean(sample)))
    means = np.array(means)
    return (float(np.mean(diff_values)), float(np.percentile(means, 2.5)), float(np.percentile(means, 97.5)))


def main() -> int:
    args = parse_args()
    rng = np.random.default_rng(args.seed)

    rules = TradeRules(
        max_trades_per_ticker_per_day=int(args.max_trades_per_ticker_per_day),
        cooldown_minutes=int(args.cooldown_minutes),
        one_position_per_ticker=True,
    )
    filters = Filters(
        min_j_th=float(args.min_j_th),
        min_forward_pf=float(args.min_forward_pf),
        min_forward_trades=int(args.min_forward_trades),
        min_win_ci_low=float(args.min_win_ci_low),
    )

    date_tags = [t.strip() for t in args.date_tags if t.strip()]
    if not date_tags:
        date_tags = iter_date_tags(args.root, args.min_date, args.max_date)
    if not date_tags:
        raise SystemExit("No date tags found under root.")

    methods = ["M0", "M1", "M2", "M3"]

    all_selected_rows: List[pd.DataFrame] = []
    all_ticker_rows: List[pd.DataFrame] = []
    day_method_rows: List[Dict[str, object]] = []

    for tag in date_tags:
        frames = collect_plan_frames(args.root, tag)
        if not frames:
            continue
        raw = pd.concat(frames, ignore_index=True)
        raw = apply_quality_filters(raw, filters)
        if raw.empty:
            continue

        for method in methods:
            selected = build_method_selection(
                raw, method=method, max_per_ticker=rules.max_trades_per_ticker_per_day
            )
            if selected.empty:
                continue
            selected["method"] = method
            all_selected_rows.append(selected)

            ticker_df = per_ticker_summary(selected)
            all_ticker_rows.append(ticker_df)

            exp_sum = float(ticker_df["exp_bp_sum"].sum()) if not ticker_df.empty else 0.0
            exp_mean = float(ticker_df["exp_bp_sum"].mean()) if not ticker_df.empty else 0.0
            day_method_rows.append(
                {
                    "date_tag": tag,
                    "method": method,
                    "tickers": int(ticker_df["Ticker"].nunique()) if not ticker_df.empty else 0,
                    "combos": int(ticker_df["combos"].sum()) if not ticker_df.empty else 0,
                    "exp_bp_sum": exp_sum,
                    "exp_bp_per_ticker_mean": exp_mean,
                }
            )

    if not day_method_rows:
        raise SystemExit("No candidates loaded after filtering; nothing to compare.")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_prefix = args.out_prefix.strip() or "method_cmp"
    summary_path = OUT_DIR / f"{out_prefix}_summary.csv"
    diffs_path = OUT_DIR / f"{out_prefix}_diffs.csv"

    day_method_df = pd.DataFrame(day_method_rows).sort_values(["date_tag", "method"])
    day_method_df.to_csv(summary_path, index=False, encoding="utf-8-sig")

    ticker_df_all = pd.concat(all_ticker_rows, ignore_index=True)

    # Paired differences at day level (more conservative than per-ticker).
    diffs: List[Dict[str, object]] = []
    pivot = day_method_df.pivot_table(
        index="date_tag", columns="method", values="exp_bp_per_ticker_mean", aggfunc="first"
    )
    pairs = [("M3", "M0"), ("M3", "M2"), ("M2", "M0"), ("M1", "M3")]
    for a, b in pairs:
        if a not in pivot.columns or b not in pivot.columns:
            continue
        diff = (pivot[a] - pivot[b]).dropna().to_numpy(dtype=float)
        mean, lo, hi = bootstrap_ci_of_mean(diff, int(args.bootstrap_iters), rng)
        diffs.append(
            {
                "metric": "exp_bp_per_ticker_mean",
                "method_a": a,
                "method_b": b,
                "n_days": int(len(diff)),
                "mean_diff": mean,
                "ci_low": lo,
                "ci_high": hi,
            }
        )

    diffs_df = pd.DataFrame(diffs)
    diffs_df.to_csv(diffs_path, index=False, encoding="utf-8-sig")

    # Print a short human-readable summary to stdout.
    by_method = (
        day_method_df.groupby("method", as_index=False)
        .agg(
            n_days=("date_tag", "nunique"),
            exp_bp_per_ticker_mean=("exp_bp_per_ticker_mean", "mean"),
            exp_bp_sum_mean=("exp_bp_sum", "mean"),
            tickers_mean=("tickers", "mean"),
            combos_mean=("combos", "mean"),
        )
        .sort_values("exp_bp_per_ticker_mean", ascending=False)
    )
    print("Time segmentation comparison (expected bp proxy):")
    print(by_method.to_string(index=False))
    print("")
    print("Paired day-level differences (A - B) with 95% CI (bootstrap):")
    if diffs_df.empty:
        print("(no diffs computed)")
    else:
        print(diffs_df.to_string(index=False))
    print("")
    print(f"Wrote: {summary_path}")
    print(f"Wrote: {diffs_path}")
    print(
        "Comparison rules (applied equally): "
        f"1 ticker = {rules.max_trades_per_ticker_per_day} trades/day max, cooldown={rules.cooldown_minutes}min, "
        "simultaneous=1 position/ticker."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

