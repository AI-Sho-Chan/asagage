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

# Replay constraints (comparison baseline)
# - One position per ticker at a time
# - Re-entry allowed only after exit + cooldown
# - Cooldown: 5 minutes
# - Max trades per ticker per day: 2
COOLDOWN_MINUTES = 5
MAX_TRADES_PER_TICKER = 2


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

INTRADAY_CACHE: Dict[Tuple[str, str], pd.DataFrame] = {}
PREV_CLOSE_CACHE: Dict[Tuple[str, str], float | None] = {}


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
    out["GapBanPct"] = num_col("GapBanPct", 3.0).clip(lower=0.0)
    out["NoTradeMin"] = num_col("NoTradeMin", 5.0).clip(lower=0.0)
    out["trend_bp_th"] = num_col("trend_bp_th", 15.0).clip(lower=0.0)
    out["forward_exp_boot_mean"] = num_col("forward_exp_boot_mean", 0.0)
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
    out["trend_driver"] = text_col("trend_driver", "NKY")
    out["trend_window"] = text_col("trend_window", "window")

    return out


def load_intraday(code: str, day: dt.date) -> pd.DataFrame:
    key = (code, day.isoformat())
    cached = INTRADAY_CACHE.get(key)
    if cached is not None:
        return cached

    path = DATA_ROOT / code / f"{day:%Y-%m-%d}.parquet"
    if not path.exists():
        INTRADAY_CACHE[key] = pd.DataFrame()
        return INTRADAY_CACHE[key]

    df = pd.read_parquet(path)
    df = df.sort_index()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [str(c[0]).lower() for c in df.columns]
    else:
        df.columns = [str(c).lower() for c in df.columns]
    needed = {"open", "high", "low", "close", "volume"}
    if not needed.issubset(set(df.columns)):
        INTRADAY_CACHE[key] = pd.DataFrame()
        return INTRADAY_CACHE[key]

    INTRADAY_CACHE[key] = df
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


def _prev_trading_close(code: str, day: dt.date) -> float | None:
    key = (code, day.isoformat())
    if key in PREV_CLOSE_CACHE:
        return PREV_CLOSE_CACHE[key]

    root = DATA_ROOT / code
    if not root.exists():
        PREV_CLOSE_CACHE[key] = None
        return None

    prev_path: Path | None = None
    for path in sorted(root.glob("????-??-??.parquet"), reverse=True):
        try:
            d = dt.datetime.strptime(path.stem, "%Y-%m-%d").date()
        except ValueError:
            continue
        if d < day:
            prev_path = path
            break

    if prev_path is None:
        PREV_CLOSE_CACHE[key] = None
        return None

    try:
        prev = pd.read_parquet(prev_path)
        if isinstance(prev.columns, pd.MultiIndex):
            prev.columns = [str(c[0]).lower() for c in prev.columns]
        else:
            prev.columns = [str(c).lower() for c in prev.columns]
        if prev.empty or "close" not in prev.columns:
            PREV_CLOSE_CACHE[key] = None
            return None
        close = float(prev["close"].iloc[-1])
        PREV_CLOSE_CACHE[key] = close if np.isfinite(close) and close > 0 else None
        return PREV_CLOSE_CACHE[key]
    except Exception:
        PREV_CLOSE_CACHE[key] = None
        return None


def _trend_proxy_code(driver: str) -> str | None:
    d = (driver or "").strip().upper()
    if d in {"NKY", "N225", "NIKKEI"}:
        return "1570.T"
    if d in {"TOPIX", "TOPX"}:
        return "1306.T"
    return None


def _trend_direction(
    driver: str,
    window_mode: str,
    window_minutes: int,
    bp_threshold: float,
    day: dt.date,
    asof: pd.Timestamp,
) -> str:
    proxy = _trend_proxy_code(driver)
    if proxy is None:
        return "BOTH"

    df = load_intraday(proxy, day)
    if df.empty:
        return "BOTH"

    close = df["close"]
    if close.empty:
        return "BOTH"

    sliced = close.loc[:asof]
    if sliced.empty:
        return "BOTH"
    close_asof = float(sliced.iloc[-1])
    if not (np.isfinite(close_asof) and close_asof > 0):
        return "BOTH"

    if (window_mode or "").strip().lower() == "day":
        base_close = float(close.iloc[0])
    else:
        start = asof - pd.Timedelta(minutes=window_minutes)
        w = close.loc[start:asof]
        base_close = float(w.iloc[0]) if not w.empty else float(close.iloc[0])

    if not (np.isfinite(base_close) and base_close > 0):
        return "BOTH"

    bp = (close_asof - base_close) / base_close * 10000.0
    if abs(bp) < float(bp_threshold):
        return "BOTH"
    return "BUY" if bp > 0 else "SELL"


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
    trend_driver: str,
    trend_window: str,
    trend_bp_th: float,
    gapban_pct: float,
    no_trade_min: float,
    counters: Dict[str, int],
) -> Dict[str, object] | None:
    intraday = load_intraday(code, day)
    if intraday.empty:
        counters["missing_intraday"] = counters.get("missing_intraday", 0) + 1
        return None
    intraday = compute_features(intraday, atr_n)
    start_t, end_t = _session_window(session)
    times = intraday.index.time
    min_entry_t = (dt.datetime.combine(day, start_t) + dt.timedelta(minutes=float(no_trade_min))).time()
    mask_window = (times >= min_entry_t) & (times <= end_t)
    if not mask_window.any():
        counters["no_session_rows"] = counters.get("no_session_rows", 0) + 1
        return None

    df = intraday.loc[mask_window].copy()
    if df.empty:
        counters["no_session_rows"] = counters.get("no_session_rows", 0) + 1
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
        counters["no_signal"] = counters.get("no_signal", 0) + 1
        return None

    entry_ts = sig_idx[0]
    entry_row = intraday.loc[entry_ts]
    side = "BUY" if float(J.loc[entry_ts]) < 0 else "SELL"

    prev_close = _prev_trading_close(code, day)
    if prev_close is not None and prev_close > 0:
        day_open = float(intraday["open"].iloc[0])
        gap_bp = (day_open - prev_close) / prev_close * 10000.0
        if abs(gap_bp) > float(gapban_pct) * 100.0:
            counters["skip_gapban"] = counters.get("skip_gapban", 0) + 1
            return None

    # AllowedSide filter: allow only BUY/SELL direction if specified (default BOTH).
    def _allow(side_val: str, side_signal: str) -> bool:
        s = (side_val or "").strip().upper()
        if s in ("", "BOTH"):
            return True
        if s == "BUY" and side_signal == "BUY":
            return True
        if s == "SELL" and side_signal == "SELL":
            return True
        return False

    # Policy B: treat empty policy as ALIGNED_ONLY (=always apply direction checks).
    if not (_allow(allowed_side_nky, side) and _allow(allowed_side_topix, side)):
        counters["skip_allowed_side"] = counters.get("skip_allowed_side", 0) + 1
        return None

    policy = (trend_policy or "").strip().upper() or "ALIGNED_ONLY"
    if policy == "ALIGNED_ONLY":
        trend_side = _trend_direction(
            trend_driver,
            trend_window,
            window_minutes=15,
            bp_threshold=float(trend_bp_th),
            day=day,
            asof=entry_ts,
        )
        if trend_side in {"BUY", "SELL"} and trend_side != side:
            counters["skip_trend_mismatch"] = counters.get("skip_trend_mismatch", 0) + 1
            return None

    px = float(entry_row["close"])
    atr_val = float(entry_row["atr"])
    if not (np.isfinite(atr_val) and atr_val > 0):
        counters["bad_atr"] = counters.get("bad_atr", 0) + 1
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
    candidate_trades: List[Dict[str, object]] = []
    counters: Dict[str, int] = {}

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
            str(row.get("trend_driver") or "NKY"),
            str(row.get("trend_window") or "window"),
            float(row.get("trend_bp_th") or 15.0),
            float(row.get("GapBanPct") or 3.0),
            float(row.get("NoTradeMin") or 5.0),
            counters,
        )
        if sim:
            sim["live_demo_class"] = str(row.get("live_demo_class") or "LIVE_BASE")
            sim["budget_factor"] = float(row.get("BudgetFactor_row", 1.0) or 1.0)
            sim["forward_exp_boot_mean"] = float(row.get("forward_exp_boot_mean", 0.0) or 0.0)
            candidate_trades.append(sim)

    if not candidate_trades:
        return pd.DataFrame(), {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0}

    def _cls_rank(cls: str) -> int:
        c = (cls or "").strip().upper()
        if c == "LIVE_STRONG":
            return 2
        if c == "LIVE_BASE":
            return 1
        return 0

    # Enforce one-position-per-ticker, cooldown, and max trades per ticker/day.
    # Each plan produces at most 1 trade ("first signal only"), but multiple plans
    # can trigger at the same time for the same ticker. We keep only the best one.
    tdf_all = pd.DataFrame(candidate_trades)
    tdf_all["entry_ts"] = pd.to_datetime(tdf_all["entry_ts"], errors="coerce")
    tdf_all["exit_ts"] = pd.to_datetime(tdf_all["exit_ts"], errors="coerce")
    tdf_all = tdf_all.dropna(subset=["entry_ts", "exit_ts"])

    cooldown = pd.Timedelta(minutes=COOLDOWN_MINUTES)
    selected_trades: List[Dict[str, object]] = []

    for code, sub in tdf_all.groupby("code", dropna=False):
        sub = sub.sort_values(["entry_ts", "exit_ts"])
        sub["_cls_rank"] = sub["live_demo_class"].map(_cls_rank)
        sub["_budget"] = pd.to_numeric(sub["budget_factor"], errors="coerce").fillna(1.0)
        sub["_exp"] = pd.to_numeric(sub["forward_exp_boot_mean"], errors="coerce").fillna(0.0)

        last_exit: pd.Timestamp | None = None
        trade_count = 0

        # Process in chronological order; for same entry_ts choose best by priority.
        for entry_ts, bucket in sub.groupby("entry_ts", sort=True):
            if trade_count >= MAX_TRADES_PER_TICKER:
                break

            bucket = bucket.sort_values(
                ["_cls_rank", "_budget", "_exp"],
                ascending=[False, False, False],
            )

            picked = None
            for _, r in bucket.iterrows():
                if last_exit is not None:
                    if r["entry_ts"] < last_exit:
                        continue
                    if r["entry_ts"] < last_exit + cooldown:
                        continue
                picked = r
                break

            if picked is None:
                continue

            selected_trades.append({k: picked[k] for k in tdf_all.columns if not str(k).startswith("_")})
            last_exit = picked["exit_ts"]
            trade_count += 1

    if not selected_trades:
        return pd.DataFrame(), {"date": day.strftime("%Y-%m-%d"), "trades": 0, "pnl_yen": 0.0, "pnl_bp_mean": 0.0}

    tdf = pd.DataFrame(selected_trades)
    pnl_yen = float(tdf["pnl_yen"].sum())
    pnl_bp_mean = float(tdf["pnl_bp"].mean())
    summary = {
        "date": day.strftime("%Y-%m-%d"),
        "trades": int(len(tdf)),
        "pnl_yen": pnl_yen,
        "pnl_bp_mean": pnl_bp_mean,
        "cooldown_minutes": COOLDOWN_MINUTES,
        "max_trades_per_ticker": MAX_TRADES_PER_TICKER,
    }
    summary.update({f"diag_{k}": int(v) for k, v in counters.items()})
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
        try:
            prev = pd.read_csv(summary_csv)
            if "date" in prev.columns:
                prev = prev[prev["date"].astype(str) != summary["date"]]
            combined = pd.concat([prev, row], ignore_index=True)
            combined.to_csv(summary_csv, index=False, encoding="utf-8-sig")
        except Exception:
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
