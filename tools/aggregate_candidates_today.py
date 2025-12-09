import argparse
import glob
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import numpy as np
import pandas as pd

root = Path("output/excel")


@dataclass
class Thresholds:
    min_j: float = 0.8
    min_win_ci: float = 0.65
    min_pf: float = 1.20
    min_trades: int = 3
    win_power: float = 1.2
    dd_scale: float = 1000.0


@dataclass
class AggregateSummary:
    rows_in: int
    rows_filtered: int
    unique_tickers: int
    avg_forward_pf: Optional[float]
    avg_forward_winrate: Optional[float]
    avg_expected_bp: Optional[float]

    def to_json(self) -> Dict[str, object]:
        return {
            "rows_in": self.rows_in,
            "rows_filtered_out": self.rows_filtered,
            "unique_tickers": self.unique_tickers,
            "avg_forward_pf": self.avg_forward_pf,
            "avg_forward_winrate": self.avg_forward_winrate,
            "avg_expected_bp": self.avg_expected_bp,
        }


def collect_candidate_paths() -> List[Path]:
    """Return candidate CSV paths sorted oldest→newest so later files win."""
    patterns = [
        root / "NIGHTLY_*" / "*" / "candidates_*.csv",
        root / "candidates_*.csv",
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

    pf = _num(df, pf_col)
    win = _num(df, _col(cols, "forward_winrate"))
    trades = _num(df, trades_col)
    dd = _num(df, "MaxDD")
    score = pf * np.power(win, thresholds.win_power) * np.log1p(trades) / (1.0 + (dd / thresholds.dd_scale))
    df["_score"] = score

    # BudgetFactor_row: Live/Demo 分類とロット倍率のたたき台。
    # nightly_summarize 側と同じロジックで、pf/win/trades から 2.0 / 1.0 / 0.5 を割り当てる。
    if pf_col and trades_col:
        pf_clip = pf.clip(lower=1.0, upper=5.0)
        live_strong = (trades >= 30) & (win >= 0.7) & (pf_clip >= 1.8)
        live_base = (trades >= 15) & (win >= 0.6) & (pf_clip >= 1.3) & ~live_strong
        df["BudgetFactor_row"] = np.where(
            live_strong,
            2.0,
            np.where(live_base, 1.0, 0.5),
        )
        df["live_demo_class"] = np.where(
            live_strong,
            "LIVE_STRONG",
            np.where(live_base, "LIVE_BASE", "DEMO_ONLY"),
        )
    else:
        df["BudgetFactor_row"] = 1.0
        df["live_demo_class"] = "LIVE_BASE"

    # AllowedSide (NKY/TOPIX) が無い場合は BOTH をデフォルトにする
    cols = _column_lookup(df)
    if "nky_allowedside" not in cols:
        df["NKY_AllowedSide"] = "BOTH"
    else:
        df["NKY_AllowedSide"] = df[cols["nky_allowedside"]].fillna("BOTH")
    if "topix_allowedside" not in cols:
        df["TOPIX_AllowedSide"] = "BOTH"
    else:
        df["TOPIX_AllowedSide"] = df[cols["topix_allowedside"]].fillna("BOTH")

    # 以前は「1銘柄につきスコア最大の 1 プランだけ」を残していたが、
    # 強いコンボが複数ある銘柄もすべて採用したいので、
    # ここではスコアでソートするだけに留めて drop_duplicates は行わない。
    ticker_col = _col(cols, "ticker")
    if ticker_col:
        df = df.sort_values([ticker_col, "_score"], ascending=[True, False])
    else:
        df = df.sort_values("_score", ascending=False)

    df.drop(columns=["_score"], inplace=True, errors="ignore")
    return df


def build_summary(df: pd.DataFrame, rows_before: int) -> AggregateSummary:
    cols = _column_lookup(df)
    pf_col = _col(cols, "forward_pf_eff")
    win_col = _col(cols, "forward_winrate")
    exp_col = _col(cols, "forward_exp_boot_mean")
    ticker_col = _col(cols, "ticker")

    def avg(col: Optional[str]) -> Optional[float]:
        if not col or col not in df.columns:
            return None
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        if series.empty:
            return None
        return float(series.mean())

    return AggregateSummary(
        rows_in=rows_before,
        rows_filtered=rows_before - len(df),
        unique_tickers=len(df[ticker_col].unique()) if ticker_col and ticker_col in df.columns else len(df),
        avg_forward_pf=avg(pf_col),
        avg_forward_winrate=avg(win_col),
        avg_expected_bp=avg(exp_col),
    )


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Aggregate candidates into candidates_nextday.csv with quality filters.")
    ap.add_argument("--output", type=Path, default=root / "candidates_nextday.csv")
    ap.add_argument("--min-j", type=float, default=0.8)
    ap.add_argument("--min-win-ci", type=float, default=0.70)
    ap.add_argument("--min-pf", type=float, default=1.30)
    ap.add_argument("--min-trades", type=int, default=5)
    ap.add_argument("--win-power", type=float, default=1.2)
    ap.add_argument("--dd-scale", type=float, default=1000.0)
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
    )

    frames: List[pd.DataFrame] = []
    for path in collect_candidate_paths():
        try:
            frames.append(pd.read_csv(path))
        except Exception:
            continue

    out = args.output
    if not frames:
        out.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_csv(out, index=False, encoding="utf-8-sig")
        print(json.dumps({"written": str(out), "rows": 0}))
        return

    combined = aggregate_frames(frames, thresholds)
    if combined is None:
        out.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame().to_csv(out, index=False, encoding="utf-8-sig")
        print(json.dumps({"written": str(out), "rows": 0}))
        return

    summary = build_summary(combined, sum(len(frame) for frame in frames))
    out.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(out, index=False, encoding="utf-8-sig")
    payload = {"written": str(out), "rows": int(len(combined))}
    payload.update(summary.to_json())
    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
