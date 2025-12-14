import argparse
import glob
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import numpy as np
import pandas as pd

ROOT = Path("output/excel")


@dataclass
class Thresholds:
    min_j: float
    min_win_ci: float
    min_pf: float
    min_trades: int
    win_power: float
    dd_scale: float
    default_gapban_pct: float
    default_no_trade_min: int


@dataclass
class AggregateSummary:
    rows_in: int
    rows_filtered_out: int
    unique_tickers: int
    avg_forward_pf: Optional[float]
    avg_forward_winrate: Optional[float]
    avg_expected_bp: Optional[float]

    def to_json(self) -> Dict[str, object]:
        return {
            "rows_in": self.rows_in,
            "rows_filtered_out": self.rows_filtered_out,
            "unique_tickers": self.unique_tickers,
            "avg_forward_pf": self.avg_forward_pf,
            "avg_forward_winrate": self.avg_forward_winrate,
            "avg_expected_bp": self.avg_expected_bp,
        }


def collect_candidate_paths() -> List[Path]:
    """Return candidate CSV paths sorted oldest→newest so later files win."""
    patterns = [
        ROOT / "NIGHTLY_*" / "*" / "candidates_*.csv",
        ROOT / "candidates_*.csv",
    ]

    paths: List[Path] = []
    for pat in patterns:
        for path_str in glob.glob(str(pat)):
            paths.append(Path(path_str))

    return sorted(paths)


def _column_lookup(df: pd.DataFrame) -> Dict[str, str]:
    return {c.lower(): c for c in df.columns}


def _col(cols: Dict[str, str], name: str) -> Optional[str]:
    return cols.get(name.lower())


def _num(df: pd.DataFrame, column: Optional[str]) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[column], errors="coerce").fillna(0.0)


def _avg(df: pd.DataFrame, column: Optional[str]) -> Optional[float]:
    if not column or column not in df.columns:
        return None
    series = pd.to_numeric(df[column], errors="coerce").dropna()
    if series.empty:
        return None
    return float(series.mean())


def _ensure_live_demo_fields(df: pd.DataFrame) -> pd.DataFrame:
    cols = _column_lookup(df)
    pf_col = _col(cols, "forward_pf_eff")
    win_col = _col(cols, "forward_winrate")
    trades_col = _col(cols, "forward_trades")

    if not (pf_col and win_col and trades_col):
        df["BudgetFactor_row"] = 1.0
        df["live_demo_class"] = "LIVE_BASE"
        return df

    pf = _num(df, pf_col).clip(lower=1.0, upper=5.0)
    win = _num(df, win_col)
    trades = _num(df, trades_col)

    live_strong = (trades >= 30) & (win >= 0.7) & (pf >= 1.8)
    live_base = (trades >= 15) & (win >= 0.6) & (pf >= 1.3) & ~live_strong

    df["BudgetFactor_row"] = np.where(live_strong, 2.0, np.where(live_base, 1.0, 0.5))
    df["live_demo_class"] = np.where(
        live_strong, "LIVE_STRONG", np.where(live_base, "LIVE_BASE", "DEMO_ONLY")
    )
    return df


def _ensure_allowed_side_fields(df: pd.DataFrame) -> pd.DataFrame:
    cols = _column_lookup(df)
    nky_col = cols.get("nky_allowedside") or cols.get("nky_allowed_side")
    topix_col = cols.get("topix_allowedside") or cols.get("topix_allowed_side")

    df["NKY_AllowedSide"] = df[nky_col].fillna("BOTH") if nky_col else "BOTH"
    df["TOPIX_AllowedSide"] = df[topix_col].fillna("BOTH") if topix_col else "BOTH"
    return df


def _ensure_gapban_notrade_fields(df: pd.DataFrame, thresholds: Thresholds) -> pd.DataFrame:
    cols = _column_lookup(df)
    gap_col = (
        cols.get("gapbanpct")
        or cols.get("gapbanpc")
        or cols.get("gap_ban_pct")
        or cols.get("gap_ban_pc")
    )
    notrade_col = cols.get("notrademin") or cols.get("no_trade_min") or cols.get("notrade_min")

    df["GapBanPct"] = _num(df, gap_col) if gap_col else float(thresholds.default_gapban_pct)
    df["NoTradeMin"] = _num(df, notrade_col) if notrade_col else float(thresholds.default_no_trade_min)
    return df


def aggregate_frames(frames: Iterable[pd.DataFrame], thresholds: Thresholds) -> Optional[pd.DataFrame]:
    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        return None

    df = pd.concat(frames, ignore_index=True)
    cols = _column_lookup(df)

    if "maxdd" not in cols and "max_dd" in cols:
        df = df.rename(columns={cols["max_dd"]: "MaxDD"})
        cols = _column_lookup(df)
    if "MaxDD" not in df.columns:
        df["MaxDD"] = 0.0

    j_col = _col(cols, "J_th")
    if j_col:
        df = df[_num(df, j_col) >= thresholds.min_j]

    win_ci_col = _col(cols, "forward_win_ci_low")
    if win_ci_col:
        df = df[_num(df, win_ci_col) >= thresholds.min_win_ci]

    pf_col = _col(cols, "forward_pf_eff")
    if pf_col:
        df = df[_num(df, pf_col) >= thresholds.min_pf]

    trades_col = _col(cols, "forward_trades")
    if trades_col:
        df = df[_num(df, trades_col) >= thresholds.min_trades]

    if df.empty:
        return df

    cols = _column_lookup(df)
    pf = _num(df, _col(cols, "forward_pf_eff"))
    win = _num(df, _col(cols, "forward_winrate"))
    trades = _num(df, _col(cols, "forward_trades"))
    dd = _num(df, "MaxDD")
    score = pf * np.power(win, thresholds.win_power) * np.log1p(trades) / (1.0 + (dd / thresholds.dd_scale))
    df["_score"] = score

    df = _ensure_live_demo_fields(df)
    df = _ensure_allowed_side_fields(df)
    df = _ensure_gapban_notrade_fields(df, thresholds)

    cols = _column_lookup(df)
    ticker_col = _col(cols, "ticker")
    if ticker_col:
        df = df.sort_values([ticker_col, "_score"], ascending=[True, False])
    else:
        df = df.sort_values("_score", ascending=False)

    df.drop(columns=["_score"], inplace=True, errors="ignore")
    return df


def build_summary(df: pd.DataFrame, rows_before: int) -> AggregateSummary:
    cols = _column_lookup(df)
    ticker_col = _col(cols, "ticker")
    return AggregateSummary(
        rows_in=rows_before,
        rows_filtered_out=rows_before - len(df),
        unique_tickers=len(df[ticker_col].unique()) if ticker_col and ticker_col in df.columns else len(df),
        avg_forward_pf=_avg(df, _col(cols, "forward_pf_eff")),
        avg_forward_winrate=_avg(df, _col(cols, "forward_winrate")),
        avg_expected_bp=_avg(df, _col(cols, "forward_exp_boot_mean")),
    )


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Aggregate candidates into candidates_nextday.csv with quality filters."
    )
    ap.add_argument("--output", type=Path, default=ROOT / "candidates_nextday.csv")
    ap.add_argument("--min-j", type=float, default=0.8)
    ap.add_argument("--min-win-ci", type=float, default=0.70)
    ap.add_argument("--min-pf", type=float, default=1.30)
    ap.add_argument("--min-trades", type=int, default=5)
    ap.add_argument("--win-power", type=float, default=1.2)
    ap.add_argument("--dd-scale", type=float, default=1000.0)
    ap.add_argument("--default-gapban-pct", type=float, default=3.0)
    ap.add_argument("--default-no-trade-min", type=int, default=5)
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    thresholds = Thresholds(
        min_j=args.min_j,
        min_win_ci=args.min_win_ci,
        min_pf=args.min_pf,
        min_trades=args.min_trades,
        win_power=args.win_power,
        dd_scale=args.dd_scale,
        default_gapban_pct=args.default_gapban_pct,
        default_no_trade_min=args.default_no_trade_min,
    )

    frames: List[pd.DataFrame] = []
    for path in collect_candidate_paths():
        try:
            frames.append(pd.read_csv(path))
        except Exception:
            continue

    out = args.output
    out.parent.mkdir(parents=True, exist_ok=True)

    if not frames:
        pd.DataFrame().to_csv(out, index=False, encoding="utf-8-sig")
        print(json.dumps({"written": str(out), "rows": 0}))
        return

    combined = aggregate_frames(frames, thresholds)
    if combined is None:
        pd.DataFrame().to_csv(out, index=False, encoding="utf-8-sig")
        print(json.dumps({"written": str(out), "rows": 0}))
        return

    summary = build_summary(combined, sum(len(frame) for frame in frames))
    combined.to_csv(out, index=False, encoding="utf-8-sig")

    payload = {"written": str(out), "rows": int(len(combined))}
    payload.update(summary.to_json())
    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()

