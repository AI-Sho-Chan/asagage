import argparse
from pathlib import Path
from typing import Iterable, List, Tuple

import pandas as pd


def _to_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def load_candidates_for_date(date_tag: str) -> Tuple[pd.DataFrame, List[pd.DataFrame]]:
    """Load per-plan candidate CSVs for a given NIGHTLY_YYYYMMDD run."""
    root = Path("output/excel") / f"NIGHTLY_{date_tag}"
    if not root.exists():
        return pd.DataFrame(), []

    plan_frames: List[pd.DataFrame] = []
    for path in sorted(root.glob("*/candidates_*.csv")):
        plan_tag = path.parent.name
        try:
            frame = pd.read_csv(path)
        except Exception:
            continue
        if frame.empty:
            continue
        frame["plan_tag"] = plan_tag
        frame["date_tag"] = date_tag
        plan_frames.append(frame)

    if not plan_frames:
        return pd.DataFrame(), []

    all_frame = pd.concat(plan_frames, ignore_index=True)
    return all_frame, plan_frames


def summarize_per_plan(frame: pd.DataFrame, date_tag: str) -> pd.DataFrame:
    """Aggregate forward metrics per (date_tag, plan_tag)."""
    if frame.empty or "plan_tag" not in frame.columns:
        return pd.DataFrame()

    group = frame.groupby("plan_tag", dropna=False)

    def numeric_mean(series: Iterable[float]) -> float:
        series_num = _to_numeric(pd.Series(list(series))).dropna()
        return float(series_num.mean()) if not series_num.empty else float("nan")

    plans: List[str] = []
    rows: List[int] = []
    unique_tickers: List[int] = []
    pf_mean: List[float] = []
    pf_median: List[float] = []
    win_mean: List[float] = []
    win_median: List[float] = []
    trades_mean: List[float] = []
    trades_median: List[float] = []

    for plan_tag, sub in group:
        plans.append(str(plan_tag))
        rows.append(len(sub))
        if "Ticker" in sub.columns:
            unique_tickers.append(sub["Ticker"].astype(str).nunique())
        else:
            unique_tickers.append(len(sub))

        pf = _to_numeric(sub.get("forward_pf_eff", pd.Series(dtype=float)))
        win = _to_numeric(sub.get("forward_winrate", pd.Series(dtype=float)))
        trades = _to_numeric(sub.get("forward_trades", pd.Series(dtype=float)))

        pf_mean.append(float(pf.mean()) if not pf.empty else float("nan"))
        pf_median.append(float(pf.median()) if not pf.empty else float("nan"))
        win_mean.append(float(win.mean()) if not win.empty else float("nan"))
        win_median.append(float(win.median()) if not win.empty else float("nan"))
        trades_mean.append(float(trades.mean()) if not trades.empty else float("nan"))
        trades_median.append(float(trades.median()) if not trades.empty else float("nan"))

    summary = pd.DataFrame(
        {
            "date_tag": date_tag,
            "plan_tag": plans,
            "rows": rows,
            "unique_tickers": unique_tickers,
            "forward_pf_eff_mean": pf_mean,
            "forward_pf_eff_median": pf_median,
            "forward_winrate_mean": win_mean,
            "forward_winrate_median": win_median,
            "forward_trades_mean": trades_mean,
            "forward_trades_median": trades_median,
        }
    )
    return summary


def summarize_strong_combos(frame: pd.DataFrame, date_tag: str) -> pd.DataFrame:
    """Extract strong combos (rough WF edge) for inspection + Live/Demo提案.

    抽出条件（後で調整可能）:
      - forward_pf_eff >= 1.3
      - forward_winrate >= 0.6
      - forward_trades >= 10

    さらに、各コンボに対して
      - live_demo_class: LIVE_STRONG / LIVE_BASE / DEMO_ONLY
      - budget_factor:   2.0 / 1.0 / 0.5
    を付与する（後で Excel 側の BudgetPerTicker × factor で数量調整に使えるようにする想定）。
    """
    if frame.empty:
        return pd.DataFrame()

    pf = _to_numeric(frame.get("forward_pf_eff", pd.Series(dtype=float)))
    win = _to_numeric(frame.get("forward_winrate", pd.Series(dtype=float)))
    trades = _to_numeric(frame.get("forward_trades", pd.Series(dtype=float)))

    mask = (pf >= 1.3) & (win >= 0.6) & (trades >= 10)
    strong = frame.loc[mask].copy()
    if strong.empty:
        return strong

    strong["forward_pf_eff"] = pf.loc[mask]
    strong["forward_winrate"] = win.loc[mask]
    strong["forward_trades"] = trades.loc[mask]

    # BudgetFactor/Live-Demo クラス付け
    live_demo_class: List[str] = []
    budget_factor: List[float] = []

    for pf_i, win_i, tr_i in zip(
        strong["forward_pf_eff"], strong["forward_winrate"], strong["forward_trades"]
    ):
        # 999 などの上限値は、過剰に効き過ぎないようにクリップして判定
        pf_clip = min(max(float(pf_i), 1.0), 5.0)

        # NOTE:
        # - まずは pf / winrate / trades のみでクラス分けし、
        #   forward_exp_bp などの期待値指標は今後の検証結果を見ながら閾値に組み込む。
        if tr_i >= 30 and win_i >= 0.7 and pf_clip >= 1.8:
            live_demo_class.append("LIVE_STRONG")
            budget_factor.append(2.0)
        elif tr_i >= 15 and win_i >= 0.6 and pf_clip >= 1.3:
            live_demo_class.append("LIVE_BASE")
            budget_factor.append(1.0)
        else:
            # 条件自体は strong 抽出を満たしているが、サンプルや期待値がやや弱い場合
            live_demo_class.append("DEMO_ONLY")
            budget_factor.append(0.5)

    strong["live_demo_class"] = live_demo_class
    strong["budget_factor"] = budget_factor

    strong = strong.sort_values(
        ["budget_factor", "forward_pf_eff", "forward_trades", "forward_winrate"],
        ascending=[False, False, False, False],
    )

    keep_columns = [
        "date_tag",
        "plan_tag",
        "Ticker",
        "SignalMode",
        "session"
        if "session" in strong.columns
        else "session_label"
        if "session_label" in strong.columns
        else None,
        "forward_pf_eff",
        "forward_winrate",
        "forward_trades",
        "live_demo_class",
        "budget_factor",
    ]
    keep_columns = [col for col in keep_columns if col is not None]

    return strong[keep_columns]


def infer_available_dates() -> List[str]:
    """Infer available NIGHTLY_YYYYMMDD folders under output/excel."""
    base = Path("output/excel")
    date_tags: List[str] = []
    for path in sorted(base.glob("NIGHTLY_*")):
        name = path.name
        if name.startswith("NIGHTLY_"):
            date_tags.append(name[len("NIGHTLY_") :])
    return date_tags


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Summarize NIGHTLY_* candidate CSVs per plan and extract strong combos.\n"
            "Default is to scan all NIGHTLY_YYYYMMDD folders under output/excel."
        )
    )
    parser.add_argument(
        "--date-tags",
        nargs="+",
        help="Target date_tag values (YYYYMMDD). If omitted, infer from NIGHTLY_* folders.",
    )
    parser.add_argument(
        "--output-plan-summary",
        type=Path,
        default=Path("analysis/nightly_plan_summary.csv"),
        help="Where to write per-plan summary CSV.",
    )
    parser.add_argument(
        "--output-strong-combos",
        type=Path,
        default=Path("analysis/nightly_strong_candidates.csv"),
        help="Where to write strong-combo CSV.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    date_tags: List[str]

    if args.date_tags:
        date_tags = list(dict.fromkeys(args.date_tags))  # preserve order, de-duplicate
    else:
        date_tags = infer_available_dates()

    if not date_tags:
        print("No NIGHTLY_YYYYMMDD folders found under output/excel.")
        return

    plan_summaries: List[pd.DataFrame] = []
    strong_summaries: List[pd.DataFrame] = []

    for date_tag in date_tags:
        all_frame, _ = load_candidates_for_date(date_tag)
        if all_frame.empty:
            continue

        plan_summary = summarize_per_plan(all_frame, date_tag)
        if not plan_summary.empty:
            plan_summaries.append(plan_summary)

        strong = summarize_strong_combos(all_frame, date_tag)
        if not strong.empty:
            strong_summaries.append(strong)

    if plan_summaries:
        combined_plans = pd.concat(plan_summaries, ignore_index=True)
        args.output_plan_summary.parent.mkdir(parents=True, exist_ok=True)
        combined_plans.to_csv(args.output_plan_summary, index=False, encoding="utf-8-sig")
        print(f"Wrote plan summary: {args.output_plan_summary} (rows={len(combined_plans)})")

    if strong_summaries:
        combined_strong = pd.concat(strong_summaries, ignore_index=True)
        args.output_strong_combos.parent.mkdir(parents=True, exist_ok=True)
        combined_strong.to_csv(args.output_strong_combos, index=False, encoding="utf-8-sig")
        print(f"Wrote strong candidates: {args.output_strong_combos} (rows={len(combined_strong)})")


if __name__ == "__main__":
    main()
