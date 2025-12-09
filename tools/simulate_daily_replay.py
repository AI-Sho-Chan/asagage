#!/usr/bin/env python3
"""
Simulate realized PnL for a trading day using 1m data and candidates snapshot.

Usage example:
  python tools/simulate_daily_replay.py --date 20251117

This reads:
  - output/excel/candidates_for_<date>.csv  (per-ticker/session parameters)
  - data/raw/yahoo_1m/<code>/<YYYY-MM-DD>.parquet  (1m OHLCV)

and writes:
  - analysis/daily_trades_<date>.csv
  - analysis/daily_realized_pnl.csv (append)
"""

from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd

DATA_ROOT = Path("data/raw/yahoo_1m")
OUT_DIR = Path("analysis")

COST_BP = 8.0


SESSION_WINDOWS: Dict[str, Tuple[dt.time, dt.time]] = {
    "AM15": (dt.time(9, 0), dt.time(9, 15)),
    "AM0930": (dt.time(9, 0), dt.time(9, 30)),
    "AM0945": (dt.time(9, 0), dt.time(9, 45)),
    "AM1000": (dt.time(9, 0), dt.time(10, 0)),
    "AM1015": (dt.time(9, 0), dt.time(10, 15)),
    "AM1030": (dt.time(9, 0), dt.time(10, 30)),
    "MID1030": (dt.time(10, 30), dt.time(11, 0)),
    "PM1230": (dt.time(12, 30), dt.time(13, 0)),
    "PM1": (dt.time(12, 30), dt.time(13, 30)),
}


def _session_window(label: str) -> Tuple[dt.time, dt.time]:
    return SESSION_WINDOWS.get(label, (dt.time(9, 0), dt.time(15, 0)))


def load_candidates(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise SystemExit(f"Candidates file not found: {path}")
    df = pd.read_csv(path)
    if df.empty:
        raise SystemExit(f"Candidates file is empty: {path}")
    return df


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cols = {c.lower(): c for c in df.columns}

    def c(name: str) -> str | None:
        return cols.get(name.lower())

    out = pd.DataFrame()
    code_col = c("Ticker") or c("code") or c("ticker")
    if not code_col:
        raise SystemExit("Ticker/code column not found in candidates CSV")
    out["code"] = df[code_col].astype(str).str.strip()

    out["session"] = df.get(c("session") or "session", "")
    out["signal_mode"] = df.get(c("SignalMode") or c("signal_mode") or "SignalMode", "j-only")

    def num_col(name: str, default: float) -> pd.Series:
        col = c(name)
        if col and col in df.columns:
            return pd.to_numeric(df[col], errors="coerce").fillna(default)
        return pd.Series(default, index=df.index)

    out["ATR_n"] = num_col("ATR_n", 3.0)
    out["TPk"] = num_col("TPk", 1.0)
    out["SLk"] = num_col("SLk", 2.0)
    out["J_th"] = num_col("J_th", 0.8)
    out["BudgetFactor_row"] = num_col("BudgetFactor_row", 1.0).clip(lower=0.0)
    # helper to get text column with default
    def text_col(name: str, default: str) -> pd.Series:
        col = cols.get(name.lower())
        if col and col in df.columns:
            return df[col].fillna(default)
        return pd.Series(default, index=df.index)

    out["live_demo_class"] = text_col("live_demo_class", "LIVE_BASE")
    out["NKY_AllowedSide"] = text_col("nky_allowedside", "BOTH")
    out["TOPIX_AllowedSide"] = text_col("topix_allowedside", "BOTH")
    out["trend_allowed_policy"] = text_col("trend_allowed_policy", "")

    return out


def load_intraday(code: str, day: dt.date) -> pd.DataFrame:
    path = DATA_ROOT / code / f"{day:%Y-%m-%d}.parquet"
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_parquet(path)
    df = df.sort_index()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [str(c[0]).lower() for c in df.columns]
    else:
        df.columns = [str(c).lower() for c in df.columns]
    needed = {"open", "high", "low", "close", "volume"}
    if not needed.issubset(set(df.columns)):
        return pd.DataFrame()
    return df


def compute_features(df: pd.DataFrame, atr_n: float) -> pd.DataFrame:
    df = df.copy()
    df["amt"] = df["close"] * df["volume"]
    df["cum_amt"] = df["amt"].cumsum()
    df["cum_vol"] = df["volume"].cumsum().replace(0, np.nan)
    df["vwap"] = df["cum_amt"] / df["cum_vol"]
    # approximate ATR using intraday TR only
    high = df["high"]
    low = df["low"]
    close = df["close"]
    prev_close = close.shift(1)
    tr = pd.concat(
        [
            (high - low).abs(),
            (high - prev_close).abs(),
            (low - prev_close).abs(),
        ],
        axis=1,
    ).max(axis=1)
    df["atr"] = tr.ewm(alpha=2.0 / (atr_n + 1.0), adjust=False).mean().replace(0, np.nan)
    return df


def simulate_trade_for_candidate(
    code: str,
    session: str,
    signal_mode: str,
    atr_n: float,
    tpk: float,
    slk: float,
    j_th: float,
    day: dt.date,
    nominal: float,
    budget_factor: float,
    allowed_side_nky: str,
    allowed_side_topix: str,
    trend_policy: str,
) -> Dict[str, object] | None:
    intraday = load_intraday(code, day)
    if intraday.empty:
        return None
    intraday = compute_features(intraday, atr_n)
    start_t, end_t = _session_window(session)
    times = intraday.index.time
    mask_window = (times >= start_t) & (times <= end_t)
    if not mask_window.any():
        return None

    df = intraday.loc[mask_window].copy()
    if df.empty:
        return None

    J = (df["close"] - df["vwap"]) / df["atr"]
    if signal_mode.lower().startswith("j-cross"):
        absJ = J.abs()
        prev = absJ.shift(1).fillna(0.0)
        sig = (absJ >= j_th) & (prev < j_th)
    else:
        sig = J.abs() >= j_th

    # first signal only
    sig_idx = sig[sig].index
    if len(sig_idx) == 0:
        return None

    entry_ts = sig_idx[0]
    entry_row = intraday.loc[entry_ts]
    side = "BUY" if float(J.loc[entry_ts]) < 0 else "SELL"

    # AllowedSide フィルタ: BUY しか許容しない/SELLしか許容しない場合は弾く
    def _allow(side_val: str, side_signal: str) -> bool:
        s = (side_val or "").strip().upper()
        if s in ("", "BOTH"):
            return True
        if s == "BUY" and side_signal == "BUY":
            return True
        if s == "SELL" and side_signal == "SELL":
            return True
        return False

    # 方針B: policyが空でも ALIGNED_ONLY 相当で常に方向チェックを入れる
    if not (_allow(allowed_side_nky, side) and _allow(allowed_side_topix, side)):
        return None
    px = float(entry_row["close"])
    atr_val = float(entry_row["atr"])
    if not (np.isfinite(atr_val) and atr_val > 0):
        return None

    if side == "BUY":
        tp = px + tpk * atr_val
        sl = px - slk * atr_val
    else:
        tp = px - tpk * atr_val
        sl = px + slk * atr_val

    # simulate bar-by-bar after entry
    after = intraday.loc[entry_ts:]
    bars = 0
    exit_px = None
    exit_ts = None
    for ts, row in after.iloc[1:].iterrows():
        hi = float(row["high"])
        lo = float(row["low"])
        bars += 1
        if side == "BUY":
            if lo <= sl:
                exit_px = sl
                exit_ts = ts
                break
            if hi >= tp:
                exit_px = tp
                exit_ts = ts
                break
        else:
            if hi >= sl:
                exit_px = sl
                exit_ts = ts
                break
            if lo <= tp:
                exit_px = tp
                exit_ts = ts
                break

    if exit_px is None:
        # close at last bar of the day
        last_ts = intraday.index[-1]
        exit_ts = last_ts
        exit_px = float(intraday.loc[last_ts, "close"])

    if side == "BUY":
        pnl_bp = (exit_px - px) / px * 10000.0
    else:
        pnl_bp = (px - exit_px) / px * 10000.0
    pnl_bp -= COST_BP
    nominal_eff = nominal * max(budget_factor, 0.0)
    pnl_yen = nominal_eff * pnl_bp / 10000.0

    return {
        "date": day.strftime("%Y-%m-%d"),
        "code": code,
        "session": session,
        "signal_mode": signal_mode,
        "side": side,
        "entry_ts": entry_ts.isoformat(),
        "exit_ts": exit_ts.isoformat() if exit_ts is not None else "",
        "entry_px": px,
        "exit_px": exit_px,
        "bars": bars,
        "pnl_bp": pnl_bp,
        "pnl_yen": pnl_yen,
    }


def simulate_day(cand_path: Path, day: dt.date, nominal: float) -> Tuple[pd.DataFrame, Dict[str, float]]:
    df_raw = load_candidates(cand_path)
    df = normalize_columns(df_raw)
    trades: List[Dict[str, object]] = []

    for _, row in df.iterrows():
        code = str(row["code"])
        session = str(row.get("session") or "")
        mode = str(row.get("signal_mode") or "j-only")
        atr_n = float(row.get("ATR_n") or 3.0)
        tpk = float(row.get("TPk") or 1.0)
        slk = float(row.get("SLk") or 2.0)
        j_th = float(row.get("J_th") or 0.8)
        sim = simulate_trade_for_candidate(
            code,
            session,
            mode,
            atr_n,
            tpk,
            slk,
            j_th,
            day,
            nominal,
            float(row.get("BudgetFactor_row", 1.0) or 1.0),
            str(row.get("NKY_AllowedSide") or "BOTH"),
            str(row.get("TOPIX_AllowedSide") or "BOTH"),
            str(row.get("trend_allowed_policy") or ""),
        )
        if sim:
            sim["live_demo_class"] = str(row.get("live_demo_class") or "LIVE_BASE")
            trades.append(sim)

    if not trades:
        return pd.DataFrame(), {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0}

    tdf = pd.DataFrame(trades)
    pnl_yen = float(tdf["pnl_yen"].sum())
    pnl_bp_mean = float(tdf["pnl_bp"].mean())
    summary = {
        "date": day.strftime("%Y-%m-%d"),
        "trades": int(len(tdf)),
        "pnl_yen": pnl_yen,
        "pnl_bp_mean": pnl_bp_mean,
    }
    # Live/Demo 別集計
    if "live_demo_class" in tdf.columns:
        for cls in ("LIVE_STRONG", "LIVE_BASE", "DEMO_ONLY"):
            sub = tdf[tdf["live_demo_class"] == cls]
            summary[f"{cls}_trades"] = int(len(sub))
            summary[f"{cls}_pnl_yen"] = float(sub["pnl_yen"].sum()) if not sub.empty else 0.0
            summary[f"{cls}_pnl_bp_mean"] = float(sub["pnl_bp"].mean()) if not sub.empty else 0.0
    return tdf, summary


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Simulate realized PnL using 1m data and candidates snapshot.")
    ap.add_argument("--date", required=True, help="Trading date in YYYYMMDD")
    ap.add_argument(
        "--candidates",
        type=Path,
        help=(
            "Candidates CSV path. "
            "If omitted, tries output/excel/candidates_for_<date>.csv, "
            "then candidates_<date>.csv, then candidates_nextday.csv (fallback)."
        ),
    )
    ap.add_argument(
        "--nominal",
        type=float,
        default=10_000_000.0,
        help="Nominal per position in yen (default 10,000,000)",
    )
    return ap.parse_args()


def main() -> None:
    args = parse_args()
    try:
        trade_date = dt.datetime.strptime(args.date, "%Y%m%d").date()
    except ValueError as exc:
        raise SystemExit(f"invalid --date {args.date}") from exc

    if args.candidates is not None:
        cand_path = args.candidates
    else:
        # Prefer the dedicated snapshot for this date; fall back to
        # candidates_<date>.csv, and finally candidates_nextday.csv so that
        # DailyReplay keeps動作 even when nightly snapshot generation fails.
        base = Path("output/excel")
        primary = base / f"candidates_for_{args.date}.csv"
        alt_date = base / f"candidates_{args.date}.csv"
        alt_nextday = base / "candidates_nextday.csv"
        if primary.exists():
            cand_path = primary
        elif alt_date.exists():
            cand_path = alt_date
        elif alt_nextday.exists():
            cand_path = alt_nextday
        else:
            raise SystemExit(
                "No candidates CSV found; looked for "
                f"{primary}, {alt_date}, {alt_nextday}"
            )

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    trades_df, summary = simulate_day(cand_path, trade_date, args.nominal)

    trades_out = OUT_DIR / f"daily_trades_{args.date}.csv"
    if not trades_df.empty:
        trades_df.to_csv(trades_out, index=False, encoding="utf-8-sig")
        print(f"written {trades_out} ({len(trades_df)} trades)")
    else:
        print(f"no trades generated for {args.date}")

    summary_csv = OUT_DIR / "daily_realized_pnl.csv"
    row = pd.DataFrame([summary])
    if summary_csv.exists():
        row.to_csv(summary_csv, mode="a", header=False, index=False, encoding="utf-8-sig")
    else:
        row.to_csv(summary_csv, index=False, encoding="utf-8-sig")
    print(f"summary appended to {summary_csv}")

    summary_json = OUT_DIR / f"daily_replay_{args.date}.json"
    with open(summary_json, "w", encoding="utf-8") as f:
        import json
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"summary written to {summary_json}")


if __name__ == "__main__":
    main()
