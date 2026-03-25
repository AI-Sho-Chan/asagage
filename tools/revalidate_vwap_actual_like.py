from __future__ import annotations

import argparse
import datetime as dt
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Sequence

import numpy as np
import pandas as pd

import simulate_daily_replay as base


DEFAULT_ORIGINAL_ROOT = Path(r"C:\AI\asagake")
DEFAULT_CANDIDATES = DEFAULT_ORIGINAL_ROOT / "output" / "excel" / "candidates_nextday.csv"
DEFAULT_DATA_ROOT = DEFAULT_ORIGINAL_ROOT / "data" / "raw" / "yahoo_1m"
DEFAULT_ANALYSIS_DIR = Path("analysis")
DEFAULT_REPORTS_DIR = Path("reports")
DEFAULT_NOMINAL = 10_000_000.0
DEFAULT_BIAS_BP = 1.0
DEFAULT_MAX_HOLD_MIN = 30.0
DEFAULT_DAILY_ENTRY_CAP = 20
DEFAULT_COST_BP = 8.0


SCENARIOS = [
    {
        "name": "baseline_current_all",
        "engine": "baseline",
        "include": None,
        "exclude": None,
    },
    {
        "name": "baseline_current_live_strong",
        "engine": "baseline",
        "include": {"LIVE_STRONG"},
        "exclude": None,
    },
    {
        "name": "actual_like_all",
        "engine": "actual_like",
        "include": None,
        "exclude": None,
    },
    {
        "name": "actual_like_live_strong",
        "engine": "actual_like",
        "include": {"LIVE_STRONG"},
        "exclude": None,
    },
    {
        "name": "actual_like_live_base_only",
        "engine": "actual_like",
        "include": {"LIVE_BASE"},
        "exclude": None,
    },
    {
        "name": "actual_like_no_live_base",
        "engine": "actual_like",
        "include": None,
        "exclude": {"LIVE_BASE"},
    },
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Re-validate the VWAP reversion idea with an actual-like passive-limit "
            "replay and compare it with the current DailyReplay baseline."
        )
    )
    parser.add_argument("--candidates", type=Path, default=DEFAULT_CANDIDATES)
    parser.add_argument("--data-root", type=Path, default=DEFAULT_DATA_ROOT)
    parser.add_argument("--start-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--end-date", required=True, help="YYYY-MM-DD")
    parser.add_argument("--analysis-dir", type=Path, default=DEFAULT_ANALYSIS_DIR)
    parser.add_argument("--reports-dir", type=Path, default=DEFAULT_REPORTS_DIR)
    parser.add_argument("--label", default="")
    parser.add_argument("--nominal", type=float, default=DEFAULT_NOMINAL)
    parser.add_argument("--bias-bp", type=float, default=DEFAULT_BIAS_BP)
    parser.add_argument("--max-hold-min", type=float, default=DEFAULT_MAX_HOLD_MIN)
    parser.add_argument("--daily-entry-cap", type=int, default=DEFAULT_DAILY_ENTRY_CAP)
    parser.add_argument("--cost-bp", type=float, default=DEFAULT_COST_BP)
    return parser.parse_args()


def _parse_date(text: str) -> dt.date:
    return dt.datetime.strptime(text, "%Y-%m-%d").date()


def _business_dates(start: dt.date, end: dt.date) -> List[dt.date]:
    out: List[dt.date] = []
    cur = start
    while cur <= end:
        if cur.weekday() < 5:
            out.append(cur)
        cur += dt.timedelta(days=1)
    return out


def _pf(series: pd.Series) -> float:
    pos = float(series[series > 0].sum())
    neg = float(series[series < 0].sum())
    if neg < 0:
        return pos / abs(neg)
    return float("inf") if pos > 0 else 0.0


def _cols(df: pd.DataFrame) -> Dict[str, str]:
    return {str(c).lower(): str(c) for c in df.columns}


def _text_col(df: pd.DataFrame, name: str, default: str = "") -> pd.Series:
    cols = _cols(df)
    col = cols.get(name.lower())
    if col and col in df.columns:
        return df[col].fillna(default).astype(str)
    return pd.Series(default, index=df.index, dtype="object")


def _num_col(df: pd.DataFrame, name: str, default: float = 0.0) -> pd.Series:
    cols = _cols(df)
    col = cols.get(name.lower())
    if col and col in df.columns:
        return pd.to_numeric(df[col], errors="coerce").fillna(default)
    return pd.Series(default, index=df.index, dtype="float64")


def _normalize_actual_like(raw: pd.DataFrame) -> pd.DataFrame:
    out = base.normalize_columns(raw)
    out["BiasSlope_row"] = _num_col(raw, "BiasSlope_row", 0.1)
    out["GapSlope_row"] = _num_col(raw, "GapSlope_row", 0.2)
    out["CorrSlope_row"] = _num_col(raw, "CorrSlope_row", 0.05)
    out["CorrNKY"] = _num_col(raw, "CorrNKY", 0.0)
    out["CorrTOPIX"] = _num_col(raw, "CorrTOPIX", 0.0)
    out["candidate_id"] = _text_col(raw, "candidate_id", "")
    missing = out["candidate_id"].str.len() == 0
    if missing.any():
        out.loc[missing, "candidate_id"] = out[missing].apply(
            lambda r: base._make_candidate_id(r.to_dict()), axis=1
        )
    return out


def _apply_class_filter(
    raw: pd.DataFrame, include: set[str] | None, exclude: set[str] | None
) -> pd.DataFrame:
    if raw.empty:
        return raw.copy()
    cls = _text_col(raw, "live_demo_class", "LIVE_BASE").str.upper()
    mask = pd.Series(True, index=raw.index)
    if include:
        mask &= cls.isin({c.upper() for c in include})
    if exclude:
        mask &= ~cls.isin({c.upper() for c in exclude})
    return raw.loc[mask].copy()


def _allow(allowed_side: str, signal_side: str) -> bool:
    side = (allowed_side or "").strip().upper()
    if side in {"", "BOTH"}:
        return True
    return side == signal_side


def _driver_corr(row: pd.Series) -> float:
    driver = str(row.get("trend_driver") or "NKY").strip().upper()
    if driver in {"TOPIX", "TOPX"}:
        return float(row.get("CorrTOPIX", 0.0) or 0.0)
    return float(row.get("CorrNKY", 0.0) or 0.0)


def _actual_like_limit(base_price: float, side: str, adj_jth: float) -> float:
    if not (np.isfinite(base_price) and base_price > 0):
        return float("nan")
    if side == "BUY":
        return base_price - 0.001 * abs(adj_jth) * base_price
    return base_price + 0.001 * abs(adj_jth) * base_price


def _select_base_price(entry_row: pd.Series, prev_close: float | None) -> float:
    vwap = float(entry_row.get("vwap", np.nan))
    if np.isfinite(vwap) and vwap > 0:
        return vwap
    if prev_close is not None and np.isfinite(prev_close) and prev_close > 0:
        return float(prev_close)
    close = float(entry_row.get("close", np.nan))
    if np.isfinite(close) and close > 0:
        return close
    return float("nan")


def _same_bar_exit(
    side: str, fill_px: float, tp_px: float, sl_px: float, hi: float, lo: float
) -> tuple[float | None, str]:
    if side == "BUY":
        hit_sl = lo <= sl_px
        hit_tp = hi >= tp_px
    else:
        hit_sl = hi >= sl_px
        hit_tp = lo <= tp_px

    if hit_sl and hit_tp:
        return sl_px, "SL_SAME_BAR"
    if hit_sl:
        return sl_px, "SL_SAME_BAR"
    if hit_tp:
        return tp_px, "TP_SAME_BAR"
    return None, ""


def simulate_actual_like_candidate(
    row: pd.Series,
    day: dt.date,
    *,
    nominal: float,
    bias_bp: float,
    max_hold_min: float,
    cost_bp: float,
    counters: Dict[str, int],
) -> Dict[str, object] | None:
    code = str(row.get("code") or "").strip()
    session = str(row.get("session") or "").strip()
    signal_mode = str(row.get("signal_mode") or "j-only")
    atr_n = float(row.get("ATR_n", 3.0) or 3.0)
    tpk = float(row.get("TPk", 1.0) or 1.0)
    slk = float(row.get("SLk", 2.0) or 2.0)
    j_th = float(row.get("J_th", 0.8) or 0.8)
    budget_factor = float(row.get("BudgetFactor_row", 1.0) or 1.0)
    allowed_side_nky = str(row.get("NKY_AllowedSide") or "BOTH")
    allowed_side_topix = str(row.get("TOPIX_AllowedSide") or "BOTH")
    trend_policy = str(row.get("trend_allowed_policy") or "")
    trend_driver = str(row.get("trend_driver") or "NKY")
    trend_window = str(row.get("trend_window") or "window")
    trend_bp_th = float(row.get("trend_bp_th", 15.0) or 15.0)
    gapban_pct = float(row.get("GapBanPct", 3.0) or 3.0)
    no_trade_min = float(row.get("NoTradeMin", 5.0) or 5.0)
    candidate_id = str(row.get("candidate_id") or "")

    intraday = base.load_intraday(code, day)
    if intraday.empty:
        counters["missing_intraday"] = counters.get("missing_intraday", 0) + 1
        return None

    intraday = base.compute_features(intraday, atr_n)
    start_t, end_t = base._session_window(session)
    times = intraday.index.time
    market_open = dt.datetime.combine(day, dt.time(9, 0))
    no_trade_cutoff = market_open + dt.timedelta(minutes=float(no_trade_min))
    session_start = dt.datetime.combine(day, start_t)
    min_entry_t = max(session_start, no_trade_cutoff).time()
    mask_window = (times >= min_entry_t) & (times <= end_t)
    if not mask_window.any():
        counters["no_session_rows"] = counters.get("no_session_rows", 0) + 1
        return None

    df = intraday.loc[mask_window].copy()
    if df.empty:
        counters["no_session_rows"] = counters.get("no_session_rows", 0) + 1
        return None

    prev_close = base._prev_trading_close(code, day)
    if prev_close is not None and prev_close > 0:
        day_open = float(intraday["open"].iloc[0])
        gap_bp_day = (day_open - prev_close) / prev_close * 10000.0
        if abs(gap_bp_day) > float(gapban_pct) * 100.0:
            counters["skip_gapban"] = counters.get("skip_gapban", 0) + 1
            return None

    J = (df["close"] - df["vwap"]) / df["atr"]
    j_th_series = pd.Series(float(j_th), index=df.index, dtype=float)

    bias_adj = float(row.get("BiasSlope_row", 0.1) or 0.1) * (float(bias_bp) / 100.0)
    corr_adj = float(row.get("CorrSlope_row", 0.05) or 0.05) * _driver_corr(row) * (
        float(bias_bp) / 100.0
    )
    gap_abs_pct_series = pd.Series(0.0, index=df.index, dtype=float)
    if prev_close is not None and prev_close > 0:
        gap_bp_series = (df["vwap"] - float(prev_close)) / float(prev_close) * 10000.0
        gap_abs_pct_series = gap_bp_series.abs() / 100.0
    gap_adj_series = float(row.get("GapSlope_row", 0.2) or 0.2) * gap_abs_pct_series
    j_th_series = (j_th_series + bias_adj + gap_adj_series + corr_adj).clip(lower=0.0)

    if signal_mode.lower().startswith("j-cross"):
        abs_j = J.abs()
        prev = abs_j.shift(1).fillna(0.0)
        prev_th = j_th_series.shift(1).fillna(float(j_th_series.iloc[0]))
        sig = (abs_j >= j_th_series) & (prev < prev_th)
    else:
        sig = J.abs() >= j_th_series

    sig_idx = sig[sig].index
    if len(sig_idx) == 0:
        counters["no_signal"] = counters.get("no_signal", 0) + 1
        return None

    signal_ts = sig_idx[0]
    signal_row = intraday.loc[signal_ts]
    side = "BUY" if float(J.loc[signal_ts]) < 0 else "SELL"

    if not (_allow(allowed_side_nky, side) and _allow(allowed_side_topix, side)):
        counters["skip_allowed_side"] = counters.get("skip_allowed_side", 0) + 1
        return None

    policy = (trend_policy or "").strip().upper() or "ALIGNED_ONLY"
    if policy == "ALIGNED_ONLY":
        trend_side = base._trend_direction(
            trend_driver,
            trend_window,
            window_minutes=15,
            bp_threshold=float(trend_bp_th),
            day=day,
            asof=signal_ts,
        )
        if trend_side in {"BUY", "SELL"} and trend_side != side:
            counters["skip_trend_mismatch"] = counters.get("skip_trend_mismatch", 0) + 1
            return None

    adj_jth = float(j_th_series.loc[signal_ts])
    base_price = _select_base_price(signal_row, prev_close)
    limit_px = _actual_like_limit(base_price, side, adj_jth)
    if not (np.isfinite(limit_px) and limit_px > 0):
        counters["bad_limit"] = counters.get("bad_limit", 0) + 1
        return None

    atr_val = float(signal_row.get("atr", np.nan))
    if not (np.isfinite(atr_val) and atr_val > 0):
        counters["bad_atr"] = counters.get("bad_atr", 0) + 1
        return None

    after_signal = intraday.loc[signal_ts:]
    fill_ts = None
    fill_px = None
    fill_bar_index = None
    for idx, (ts, bar) in enumerate(after_signal.iloc[1:].iterrows(), start=1):
        lo = float(bar["low"])
        hi = float(bar["high"])
        if side == "BUY" and lo <= limit_px:
            fill_ts = ts
            fill_px = float(limit_px)
            fill_bar_index = idx
            break
        if side == "SELL" and hi >= limit_px:
            fill_ts = ts
            fill_px = float(limit_px)
            fill_bar_index = idx
            break
    if fill_ts is None or fill_px is None or fill_bar_index is None:
        counters["unfilled"] = counters.get("unfilled", 0) + 1
        return None

    if side == "BUY":
        tp_px = fill_px + tpk * atr_val
        sl_px = fill_px - slk * atr_val
    else:
        tp_px = fill_px - tpk * atr_val
        sl_px = fill_px + slk * atr_val

    fill_bar = after_signal.iloc[fill_bar_index]
    same_bar_exit_px, same_bar_reason = _same_bar_exit(
        side, fill_px, tp_px, sl_px, float(fill_bar["high"]), float(fill_bar["low"])
    )

    exit_px = None
    exit_ts = None
    exit_reason = ""
    bars_held = 0
    deadline = fill_ts + pd.Timedelta(minutes=float(max_hold_min))

    if same_bar_exit_px is not None:
        exit_px = same_bar_exit_px
        exit_ts = fill_ts
        exit_reason = same_bar_reason
    else:
        after_fill = intraday.loc[fill_ts:]
        for i, (ts, bar) in enumerate(after_fill.iloc[1:].iterrows(), start=1):
            hi = float(bar["high"])
            lo = float(bar["low"])
            bars_held = i
            if side == "BUY":
                hit_sl = lo <= sl_px
                hit_tp = hi >= tp_px
            else:
                hit_sl = hi >= sl_px
                hit_tp = lo <= tp_px

            if hit_sl and hit_tp:
                exit_px = sl_px
                exit_ts = ts
                exit_reason = "SL_SAME_BAR"
                break
            if hit_sl:
                exit_px = sl_px
                exit_ts = ts
                exit_reason = "SL"
                break
            if hit_tp:
                exit_px = tp_px
                exit_ts = ts
                exit_reason = "TP"
                break
            if ts >= deadline:
                exit_px = float(bar["close"])
                exit_ts = ts
                exit_reason = "TIMEOUT"
                break

    if exit_px is None or exit_ts is None:
        last_ts = intraday.index[-1]
        exit_ts = last_ts
        exit_px = float(intraday.loc[last_ts, "close"])
        exit_reason = "EOD"

    if side == "BUY":
        pnl_bp = (exit_px - fill_px) / fill_px * 10000.0
        tp_dist_bp = (tp_px - fill_px) / fill_px * 10000.0
        sl_dist_bp = (fill_px - sl_px) / fill_px * 10000.0
    else:
        pnl_bp = (fill_px - exit_px) / fill_px * 10000.0
        tp_dist_bp = (fill_px - tp_px) / fill_px * 10000.0
        sl_dist_bp = (sl_px - fill_px) / fill_px * 10000.0
    pnl_bp -= float(cost_bp)
    pnl_yen = float(nominal) * max(float(budget_factor), 0.0) * pnl_bp / 10000.0

    return {
        "date": day.strftime("%Y-%m-%d"),
        "code": code,
        "session": session,
        "signal_mode": signal_mode,
        "side": side,
        "signal_ts": signal_ts.isoformat(),
        "fill_ts": fill_ts.isoformat(),
        "exit_ts": exit_ts.isoformat(),
        "entry_px": fill_px,
        "exit_px": exit_px,
        "entry_limit_px": limit_px,
        "adj_jth": adj_jth,
        "bars": bars_held,
        "exit_reason": exit_reason,
        "tp_px": tp_px,
        "sl_px": sl_px,
        "tp_dist_bp": tp_dist_bp,
        "sl_dist_bp": sl_dist_bp,
        "pnl_bp": pnl_bp,
        "pnl_yen": pnl_yen,
        "budget_factor": budget_factor,
        "forward_exp_boot_mean": float(row.get("forward_exp_boot_mean", 0.0) or 0.0),
        "live_demo_class": str(row.get("live_demo_class") or "LIVE_BASE"),
        "candidate_id": candidate_id,
    }


def _class_rank(text: str) -> int:
    cls = (text or "").strip().upper()
    if cls == "LIVE_STRONG":
        return 2
    if cls == "LIVE_BASE":
        return 1
    return 0


def simulate_actual_like_day(
    raw_filtered: pd.DataFrame,
    day: dt.date,
    *,
    nominal: float,
    bias_bp: float,
    max_hold_min: float,
    cost_bp: float,
    daily_entry_cap: int,
) -> tuple[pd.DataFrame, Dict[str, object]]:
    if raw_filtered.empty:
        return (
            pd.DataFrame(),
            {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0},
        )

    df = _normalize_actual_like(raw_filtered)
    counters: Dict[str, int] = {}
    candidates: List[Dict[str, object]] = []
    for _, row in df.iterrows():
        sim = simulate_actual_like_candidate(
            row,
            day,
            nominal=nominal,
            bias_bp=bias_bp,
            max_hold_min=max_hold_min,
            cost_bp=cost_bp,
            counters=counters,
        )
        if sim:
            candidates.append(sim)

    if not candidates:
        summary = {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0}
        summary.update({f"diag_{k}": int(v) for k, v in counters.items()})
        return pd.DataFrame(), summary

    tdf_all = pd.DataFrame(candidates)
    tdf_all["fill_ts"] = pd.to_datetime(tdf_all["fill_ts"], errors="coerce")
    tdf_all["exit_ts"] = pd.to_datetime(tdf_all["exit_ts"], errors="coerce")
    tdf_all = tdf_all.dropna(subset=["fill_ts", "exit_ts"]).copy()
    tdf_all["_cls_rank"] = tdf_all["live_demo_class"].map(_class_rank)
    tdf_all["_budget"] = pd.to_numeric(tdf_all["budget_factor"], errors="coerce").fillna(1.0)
    tdf_all["_exp"] = pd.to_numeric(tdf_all["forward_exp_boot_mean"], errors="coerce").fillna(0.0)

    selected: List[Dict[str, object]] = []
    last_exit_by_code: Dict[str, pd.Timestamp] = {}
    used_entries = 0

    for fill_ts, bucket in tdf_all.groupby("fill_ts", sort=True):
        if daily_entry_cap > 0 and used_entries >= daily_entry_cap:
            break
        bucket = bucket.sort_values(["_cls_rank", "_budget", "_exp"], ascending=[False, False, False])
        for _, trade in bucket.iterrows():
            code = str(trade["code"])
            last_exit = last_exit_by_code.get(code)
            if last_exit is not None and trade["fill_ts"] < last_exit:
                continue
            selected.append({k: trade[k] for k in tdf_all.columns if not str(k).startswith("_")})
            last_exit_by_code[code] = trade["exit_ts"]
            used_entries += 1
            if daily_entry_cap > 0 and used_entries >= daily_entry_cap:
                break

    if not selected:
        summary = {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0}
        summary.update({f"diag_{k}": int(v) for k, v in counters.items()})
        return pd.DataFrame(), summary

    tdf = pd.DataFrame(selected)
    summary = {
        "date": day.strftime("%Y-%m-%d"),
        "trades": int(len(tdf)),
        "pnl_yen": float(tdf["pnl_yen"].sum()),
        "pnl_bp_mean": float(tdf["pnl_bp"].mean()),
    }
    summary.update({f"diag_{k}": int(v) for k, v in counters.items()})
    return tdf, summary


def simulate_baseline_day(
    raw_filtered: pd.DataFrame,
    day: dt.date,
    *,
    nominal: float,
) -> tuple[pd.DataFrame, Dict[str, object]]:
    if raw_filtered.empty:
        return (
            pd.DataFrame(),
            {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0},
        )

    with tempfile.TemporaryDirectory(prefix="asagake_reval_") as td:
        tmp_path = Path(td) / f"candidates_{day:%Y%m%d}.csv"
        raw_filtered.to_csv(tmp_path, index=False)
        trades_df, summary = base.simulate_day(
            tmp_path,
            day,
            nominal,
            cooldown_minutes=5,
            max_trades_per_ticker=2,
            stop_after_loss=False,
            dtw=None,
        )
    return trades_df.copy(), dict(summary)


def _summarize_trades(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(
            columns=["scenario", "trades", "pnl_yen_sum", "pnl_bp_mean", "win_rate", "pf"]
        )
    rows = []
    for scenario, sub in df.groupby("scenario", dropna=False):
        pnl = pd.to_numeric(sub["pnl_yen"], errors="coerce").fillna(0.0)
        pnl_bp = pd.to_numeric(sub["pnl_bp"], errors="coerce")
        rows.append(
            {
                "scenario": str(scenario),
                "trades": int(len(sub)),
                "pnl_yen_sum": float(pnl.sum()),
                "pnl_bp_mean": float(pnl_bp.mean()) if len(sub) else np.nan,
                "win_rate": float((pnl > 0).mean()) if len(sub) else np.nan,
                "pf": float(_pf(pnl)),
            }
        )
    return pd.DataFrame(rows).sort_values("pnl_yen_sum")


def _summarize_daily(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["scenario", "date", "trades", "pnl_yen", "pnl_bp_mean"])
    out = (
        df.groupby(["scenario", "date"], dropna=False)
        .agg(
            trades=("scenario", "size"),
            pnl_yen=("pnl_yen", "sum"),
            pnl_bp_mean=("pnl_bp", "mean"),
        )
        .reset_index()
        .sort_values(["scenario", "date"])
    )
    return out


def _summarize_class(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "live_demo_class" not in df.columns:
        return pd.DataFrame(columns=["scenario", "live_demo_class", "trades", "pnl_yen", "pnl_bp_mean"])
    out = (
        df.groupby(["scenario", "live_demo_class"], dropna=False)
        .agg(
            trades=("scenario", "size"),
            pnl_yen=("pnl_yen", "sum"),
            pnl_bp_mean=("pnl_bp", "mean"),
        )
        .reset_index()
        .sort_values(["scenario", "pnl_yen"])
    )
    return out


def _to_md(df: pd.DataFrame, rows: int = 20) -> str:
    if df.empty:
        return "_none_"
    head = df.head(rows)
    try:
        return head.to_markdown(index=False)
    except Exception:
        cols = [str(c) for c in head.columns]
        lines = [
            "| " + " | ".join(cols) + " |",
            "| " + " | ".join(["---"] * len(cols)) + " |",
        ]
        for _, row in head.iterrows():
            vals = [str(row[c]) for c in head.columns]
            lines.append("| " + " | ".join(vals) + " |")
        return "\n".join(lines)


def build_report(
    *,
    label: str,
    cand_path: Path,
    start_date: dt.date,
    end_date: dt.date,
    trades: pd.DataFrame,
    summary: pd.DataFrame,
    daily: pd.DataFrame,
    by_class: pd.DataFrame,
) -> str:
    lines: List[str] = []
    lines.append(f"# ASAGAKE Revalidation Report ({label})")
    lines.append("")
    lines.append(f"- candidates: `{cand_path}`")
    lines.append(f"- window: `{start_date}` to `{end_date}`")
    lines.append("")

    if summary.empty:
        lines.append("No trades were generated.")
        return "\n".join(lines)

    lines.append("## Scenario Summary")
    lines.append("")
    lines.append(_to_md(summary, rows=20))
    lines.append("")

    baseline = summary[summary["scenario"] == "baseline_current_all"]
    actual_all = summary[summary["scenario"] == "actual_like_all"]
    actual_strong = summary[summary["scenario"] == "actual_like_live_strong"]
    actual_base = summary[summary["scenario"] == "actual_like_live_base_only"]
    if not baseline.empty and not actual_all.empty:
        base_pnl = float(baseline["pnl_yen_sum"].iloc[0])
        actual_pnl = float(actual_all["pnl_yen_sum"].iloc[0])
        lines.append("## Readout")
        lines.append("")
        lines.append(f"- baseline_current_all pnl: `{base_pnl:,.0f} yen`")
        lines.append(f"- actual_like_all pnl: `{actual_pnl:,.0f} yen`")
        lines.append(f"- delta(actual_like - baseline): `{actual_pnl - base_pnl:,.0f} yen`")
        if not actual_strong.empty:
            lines.append(
                f"- actual_like_live_strong pnl: `{float(actual_strong['pnl_yen_sum'].iloc[0]):,.0f} yen`"
            )
        if not actual_base.empty:
            lines.append(
                f"- actual_like_live_base_only pnl: `{float(actual_base['pnl_yen_sum'].iloc[0]):,.0f} yen`"
            )
        lines.append("")

    lines.append("## Daily Summary")
    lines.append("")
    lines.append(_to_md(daily, rows=80))
    lines.append("")

    if not by_class.empty:
        lines.append("## Class Breakdown")
        lines.append("")
        lines.append(_to_md(by_class, rows=40))
        lines.append("")

    actual_trades = trades[trades["scenario"].astype(str).str.startswith("actual_like_")].copy()
    if not actual_trades.empty:
        worst = actual_trades.sort_values("pnl_yen").head(15)[
            ["scenario", "date", "code", "session", "signal_mode", "side", "exit_reason", "pnl_yen", "pnl_bp"]
        ]
        best = actual_trades.sort_values("pnl_yen", ascending=False).head(15)[
            ["scenario", "date", "code", "session", "signal_mode", "side", "exit_reason", "pnl_yen", "pnl_bp"]
        ]
        lines.append("## Worst Actual-like Trades")
        lines.append("")
        lines.append(_to_md(worst, rows=15))
        lines.append("")
        lines.append("## Best Actual-like Trades")
        lines.append("")
        lines.append(_to_md(best, rows=15))
        lines.append("")

    return "\n".join(lines)


def main() -> None:
    args = parse_args()
    start_date = _parse_date(args.start_date)
    end_date = _parse_date(args.end_date)
    label = args.label.strip() or f"{args.candidates.stem}_{start_date:%Y%m%d}_{end_date:%Y%m%d}"

    args.analysis_dir.mkdir(parents=True, exist_ok=True)
    args.reports_dir.mkdir(parents=True, exist_ok=True)

    base.DATA_ROOT = Path(args.data_root)
    base.INTRADAY_CACHE.clear()
    base.PREV_CLOSE_CACHE.clear()

    raw_candidates = base.load_candidates(args.candidates)
    dates = _business_dates(start_date, end_date)
    if not dates:
        raise SystemExit("No business days in the selected range")

    trade_frames: List[pd.DataFrame] = []
    daily_rows: List[Dict[str, object]] = []

    for day in dates:
        for scenario in SCENARIOS:
            filtered = _apply_class_filter(
                raw_candidates, scenario.get("include"), scenario.get("exclude")
            )
            if scenario["engine"] == "baseline":
                trades_df, summary = simulate_baseline_day(
                    filtered,
                    day,
                    nominal=float(args.nominal),
                )
            else:
                trades_df, summary = simulate_actual_like_day(
                    filtered,
                    day,
                    nominal=float(args.nominal),
                    bias_bp=float(args.bias_bp),
                    max_hold_min=float(args.max_hold_min),
                    cost_bp=float(args.cost_bp),
                    daily_entry_cap=int(args.daily_entry_cap),
                )

            summary["scenario"] = scenario["name"]
            daily_rows.append(summary)
            if not trades_df.empty:
                tdf = trades_df.copy()
                tdf["scenario"] = scenario["name"]
                trade_frames.append(tdf)

    trades = pd.concat(trade_frames, ignore_index=True) if trade_frames else pd.DataFrame()
    daily = pd.DataFrame(daily_rows).sort_values(["scenario", "date"])
    summary = _summarize_trades(trades)
    by_class = _summarize_class(trades)

    trade_out = args.analysis_dir / f"revalidation_trades_{label}.csv"
    daily_out = args.analysis_dir / f"revalidation_daily_{label}.csv"
    summary_out = args.analysis_dir / f"revalidation_summary_{label}.csv"
    class_out = args.analysis_dir / f"revalidation_class_{label}.csv"
    report_out = args.reports_dir / f"revalidation_report_{label}.md"

    if trades.empty:
        pd.DataFrame(columns=["scenario"]).to_csv(trade_out, index=False)
    else:
        trades.to_csv(trade_out, index=False)
    daily.to_csv(daily_out, index=False)
    summary.to_csv(summary_out, index=False)
    by_class.to_csv(class_out, index=False)

    report_text = build_report(
        label=label,
        cand_path=args.candidates,
        start_date=start_date,
        end_date=end_date,
        trades=trades,
        summary=summary,
        daily=daily,
        by_class=by_class,
    )
    report_out.write_text(report_text, encoding="utf-8")

    print(f"trade_out={trade_out}")
    print(f"daily_out={daily_out}")
    print(f"summary_out={summary_out}")
    print(f"class_out={class_out}")
    print(f"report_out={report_out}")


if __name__ == "__main__":
    main()
