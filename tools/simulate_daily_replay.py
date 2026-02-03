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
import email.message
import hashlib
import json
import smtplib
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from yahooquery import Ticker

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if SRC.exists():
    sys.path.insert(0, str(SRC))

from asagake_io.csv_schemas import DT_V1, schema_columns
from asagake_io.csv_writer import DecisionTraceWriter, make_append_only_writer

DATA_ROOT = Path("data/raw/yahoo_1m")
OUT_DIR = Path("analysis")

COST_BP = 8.0

# Replay constraints (comparison baseline)
# - One position per ticker at a time
# - Re-entry allowed only after exit + cooldown
# - Cooldown: 5 minutes
# - Max trades per ticker per day: 2
DEFAULT_COOLDOWN_MINUTES = 5
DEFAULT_MAX_TRADES_PER_TICKER = 2

JST = dt.timezone(dt.timedelta(hours=9))


def _load_smtp_config(path: Path) -> dict | None:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def _send_mail_via_smtp(smtp_cfg: dict, recipient: str, subject: str, body: str) -> None:
    user = smtp_cfg.get("user")
    password = smtp_cfg.get("pass")
    host = smtp_cfg.get("host")
    port = int(smtp_cfg.get("port", 587))
    if not (user and password and host and recipient):
        raise RuntimeError("smtp.json must contain host/port/user/pass and recipient must be set")

    msg = email.message.EmailMessage()
    msg["From"] = user
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(user, password)
        smtp.send_message(msg)


def _iso_ts_jst(ts: dt.datetime | pd.Timestamp) -> str:
    if isinstance(ts, pd.Timestamp):
        ts = ts.to_pydatetime()
    if ts.tzinfo is None:
        ts = ts.replace(tzinfo=JST)
    return ts.isoformat()


def _detect_engine_version() -> str:
    try:
        r = subprocess.run(
            ["git", "rev-parse", "--short", "HEAD"],
            check=True,
            capture_output=True,
            text=True,
        )
        v = (r.stdout or "").strip()
        return v or "unknown"
    except Exception:
        return "unknown"


def _make_candidate_id(row: dict) -> str:
    ticker = str(row.get("code") or "").strip()
    session = str(row.get("session") or "").strip()
    payload = "|".join(
        [
            ticker,
            session,
            str(row.get("signal_mode") or ""),
            f"{float(row.get('J_th') or 0.0):.6g}",
            f"{float(row.get('TPk') or 0.0):.6g}",
            f"{float(row.get('SLk') or 0.0):.6g}",
            f"{float(row.get('ATR_n') or 0.0):.6g}",
            f"{float(row.get('BudgetFactor_row') or 1.0):.6g}",
            f"{float(row.get('NoTradeMin') or 0.0):.6g}",
            f"{float(row.get('GapBanPct') or 0.0):.6g}",
            str(row.get("trend_driver") or ""),
            str(row.get("trend_window") or ""),
            f"{float(row.get('trend_bp_th') or 0.0):.6g}",
            str(row.get("trend_allowed_policy") or ""),
        ]
    )
    h = hashlib.sha1(payload.encode("utf-8")).hexdigest()[:8]
    return f"C_{ticker}_{session}_{h}"


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


def build_mail_report(
    *,
    date_tag: str,
    cand_path: Path,
    trades_df: pd.DataFrame,
    summary: Dict[str, object],
    nominal: float,
) -> str:
    lines: List[str] = []
    lines.append(f"ASAGAKE DailyReplay（仮想売買） {date_tag}")
    lines.append("")
    lines.append(f"候補CSV: {cand_path.as_posix()}")
    lines.append("（注）このメールは LIVE_STRONG をメインに表示します。LIVE_BASE / DEMO_ONLY は参考枠です。")
    cooldown_minutes = int(summary.get("cooldown_minutes", DEFAULT_COOLDOWN_MINUTES) or DEFAULT_COOLDOWN_MINUTES)
    max_trades = int(summary.get("max_trades_per_ticker", DEFAULT_MAX_TRADES_PER_TICKER) or DEFAULT_MAX_TRADES_PER_TICKER)
    stop_after_loss = bool(summary.get("stop_after_loss", False))
    max_trades_label = "制限なし" if max_trades <= 0 else f"1日{max_trades}回まで"
    extra_rule = "（その銘柄で損失が出たら当日はそれ以上取引しない）" if stop_after_loss else ""
    lines.append(
        "前提: "
        f"予算={int(nominal):,}円/プラン、"
        "同時1ポジション（同一銘柄は同時に1つまで）、"
        "決済後のみ再エントリー、"
        f"クールダウン{cooldown_minutes}分、"
        f"同一銘柄は{max_trades_label}{extra_rule}"
    )
    lines.append("")

    trades = int(summary.get("trades", 0) or 0)
    pnl_yen = float(summary.get("pnl_yen", 0.0) or 0.0)
    pnl_bp_mean = summary.get("pnl_bp_mean")

    focus_class = "LIVE_STRONG"
    if "live_demo_class" in trades_df.columns:
        focus_df = trades_df[trades_df["live_demo_class"].astype(str) == focus_class].copy()
    else:
        focus_df = pd.DataFrame()

    focus_trades = int(len(focus_df))
    focus_pnl_yen = float(focus_df["pnl_yen"].sum()) if focus_trades else 0.0
    focus_pnl_bp_mean: float | None = None
    if focus_trades and "pnl_bp" in focus_df.columns:
        s = pd.to_numeric(focus_df["pnl_bp"], errors="coerce")
        if not s.isna().all():
            focus_pnl_bp_mean = float(s.mean())

    lines.append(f"結果サマリ（メイン: {focus_class}）:")
    lines.append(f"  {focus_class}: 取引 {focus_trades} 回 / 合計損益 {focus_pnl_yen:,.0f} 円")
    if focus_pnl_bp_mean is not None:
        lines.append(f"  （平均損益 {focus_pnl_bp_mean:.1f} bp）")

    lines.append(f"参考（全体合計）: 取引 {trades} 回 / 合計損益 {pnl_yen:,.0f} 円")
    if pnl_bp_mean is not None:
        try:
            lines.append(f"（平均損益 {float(pnl_bp_mean):.1f} bp）")
        except (TypeError, ValueError):
            pass

    skip_trend = summary.get("diag_skip_trend_mismatch")
    no_signal = summary.get("diag_no_signal")
    if skip_trend is not None or no_signal is not None:
        parts = []
        if skip_trend is not None:
            parts.append(f"方向不一致で見送り {skip_trend} 件")
        if no_signal is not None:
            parts.append(f"シグナル無し {no_signal} 件")
        if parts:
            lines.append("見送り: " + " / ".join(parts))

    lines.append("")
    if trades_df.empty:
        lines.append("この日は、条件に合う取引がありませんでした。")
        lines.append("")
        lines.append("補足: これは「取引終了後に、Yahooの1分足で再現した仮想売買」です。Excelの場中ログと一致しないことがあります。")
        return "\n".join(lines)

    # class breakdown
    lines.append("クラス別（参考: 強/標準/デモ）:")
    if all(k in summary for k in ("LIVE_STRONG_trades", "LIVE_BASE_trades", "DEMO_ONLY_trades")):
        lines.append(
            f"  LIVE_STRONG: {int(summary.get('LIVE_STRONG_trades', 0) or 0)}回 / {float(summary.get('LIVE_STRONG_pnl_yen', 0.0) or 0.0):,.0f}円"
        )
        lines.append(
            f"  LIVE_BASE:   {int(summary.get('LIVE_BASE_trades', 0) or 0)}回 / {float(summary.get('LIVE_BASE_pnl_yen', 0.0) or 0.0):,.0f}円"
        )
        lines.append(
            f"  DEMO_ONLY:   {int(summary.get('DEMO_ONLY_trades', 0) or 0)}回 / {float(summary.get('DEMO_ONLY_pnl_yen', 0.0) or 0.0):,.0f}円"
        )
    else:
        gb = trades_df.groupby("live_demo_class")["pnl_yen"].sum().sort_index()
        for k, v in gb.items():
            lines.append(f"  {k}: {v:,.0f}円")

    # exit reason breakdown
    rank_df = focus_df if not focus_df.empty else trades_df
    rank_label = focus_class if not focus_df.empty else "全体"
    if "exit_reason" in rank_df.columns:
        reason_map = {
            "TP": "利確（TP）",
            "SL": "損切り（SL）",
            "EOD": "引け決済（EOD）",
            "SL_SAME_BAR": "損切り（SL, 同じ1分でTPも到達の可能性）",
            "TP_SAME_BAR": "利確（TP, 同じ1分でSLも到達の可能性）",
        }
        vc = rank_df["exit_reason"].astype(str).fillna("").replace("", "UNKNOWN").value_counts()
        if not vc.empty:
            lines.append("")
            lines.append(f"決済理由（内訳: {rank_label}）:")
            for k, v in vc.items():
                label = reason_map.get(k, k)
                lines.append(f"  {label}: {int(v)}件")

    # top losses / wins
    cols = [
        "code",
        "session",
        "signal_mode",
        "side",
        "pnl_yen",
        "pnl_bp",
        "bars",
        "exit_reason",
        "live_demo_class",
        "budget_factor",
    ]
    for c in cols:
        if c not in rank_df.columns:
            rank_df[c] = ""

    losses = rank_df.sort_values("pnl_yen", ascending=True).head(5)
    wins = rank_df.sort_values("pnl_yen", ascending=False).head(5)

    lines.append("")
    lines.append(f"負けが大きかった順（上位5件: {rank_label}）:")
    for _, r in losses.iterrows():
        exit_reason = str(r.get("exit_reason") or "")
        bars = int(float(r.get("bars") or 0))
        lines.append(
            f"  {r['code']} {r['session']} {r['signal_mode']} {r['side']}: {float(r['pnl_yen']):,.0f}円 ({float(r['pnl_bp']):.1f}bp) {exit_reason} {bars}分 [{r['live_demo_class']}] x{float(r['budget_factor']) if str(r['budget_factor']) else ''}"
        )

    lines.append("")
    lines.append(f"勝ちが大きかった順（上位5件: {rank_label}）:")
    for _, r in wins.iterrows():
        exit_reason = str(r.get("exit_reason") or "")
        bars = int(float(r.get("bars") or 0))
        lines.append(
            f"  {r['code']} {r['session']} {r['signal_mode']} {r['side']}: {float(r['pnl_yen']):,.0f}円 ({float(r['pnl_bp']):.1f}bp) {exit_reason} {bars}分 [{r['live_demo_class']}] x{float(r['budget_factor']) if str(r['budget_factor']) else ''}"
        )

    lines.append("")
    lines.append("かんたんな解説:")
    min_loss = float(rank_df["pnl_yen"].min())
    pnl_ref = focus_pnl_yen if focus_trades else pnl_yen
    if min_loss < 0 and abs(min_loss) > abs(pnl_ref) * 0.6:
        lines.append("  1つの大きな負けが、1日の結果をほぼ決めています（大負けを減らすと安定します）。")
    else:
        lines.append("  複数の勝ち/負けの合計で結果が決まっています（大負けの有無を確認してください）。")
    lines.append("  ※これは「取引終了後に、Yahooの1分足データで再現した仮想売買」です。Excelの場中ログと一致しないことがあります。")

    return "\n".join(lines)

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
        # Fetch on-demand (so DailyReplay can run even when minute_cache is not pre-filled).
        try:
            hist = Ticker([code], asynchronous=False).history(
                start=str(day), end=str(day + dt.timedelta(days=1)), interval="1m"
            )
        except Exception:
            hist = pd.DataFrame()

        fetched = pd.DataFrame()
        if isinstance(hist, pd.DataFrame) and not hist.empty:
            if isinstance(hist.index, pd.MultiIndex):
                try:
                    sub = hist.xs(code, level=0, drop_level=True).reset_index()
                except Exception:
                    sub = pd.DataFrame()
            else:
                sub = hist.reset_index()

            if not sub.empty:
                time_column = None
                for candidate in ("date", "datetime", "ts"):
                    if candidate in sub.columns:
                        time_column = candidate
                        break
                if time_column is not None:
                    sub = sub.rename(columns={time_column: "ts"})
                    ts = pd.to_datetime(sub["ts"], errors="coerce", utc=True)
                    sub["ts"] = ts.dt.tz_convert("Asia/Tokyo")
                    sub = sub[sub["ts"].dt.date == day]
                    if not sub.empty:
                        fetched = sub.set_index("ts")

        if fetched.empty:
            INTRADAY_CACHE[key] = pd.DataFrame()
            return INTRADAY_CACHE[key]

        # Normalize columns and (best-effort) persist for future runs.
        if isinstance(fetched.columns, pd.MultiIndex):
            fetched.columns = [str(c[0]).lower() for c in fetched.columns]
        else:
            fetched.columns = [str(c).lower() for c in fetched.columns]

        keep = [c for c in ("open", "high", "low", "close", "volume") if c in fetched.columns]
        fetched = fetched[keep].sort_index()

        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            fetched.to_parquet(path)
        except Exception:
            pass

        INTRADAY_CACHE[key] = fetched
        return fetched

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


def _minutes_from_open(entry_ts: pd.Timestamp, market_open: dt.datetime) -> int:
    """Minutes from market open (09:00) for diagnostics/logging.

    Some intraday sources can yield tz-aware timestamps; for this coarse metric
    we treat the datetime as local clock time and drop tzinfo to avoid
    naive/aware subtraction errors.
    """

    entry_dt = entry_ts.to_pydatetime()
    if getattr(entry_dt, "tzinfo", None) is not None:
        entry_dt = entry_dt.replace(tzinfo=None)
    return int((entry_dt - market_open).total_seconds() // 60)


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
    *,
    dtw: DecisionTraceWriter | None = None,
    candidate_id: str = "",
) -> Dict[str, object] | None:
    intraday = load_intraday(code, day)
    if intraday.empty:
        counters["missing_intraday"] = counters.get("missing_intraday", 0) + 1
        return None
    intraday = compute_features(intraday, atr_n)
    start_t, end_t = _session_window(session)
    times = intraday.index.time
    # NoTradeMin means "ignore the first N minutes after the market open (09:00)".
    # For sessions that start later than 09:00, do not push the window later.
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

    decision_id = f"D_{day.strftime('%Y%m%d')}_{entry_ts.strftime('%H%M%S')}_{candidate_id[-8:] or '00000000'}"
    client_order_id = f"O_{day.strftime('%Y%m%d')}_{entry_ts.strftime('%H%M%S')}_{candidate_id[-8:] or '00000000'}"

    def _dt_emit(event_type: str, **fields: object) -> None:
        if dtw is None:
            return
        base = {
            "event_ts": _iso_ts_jst(entry_ts),
            "event_type": event_type,
            "ticker": code,
            "session": session,
            "candidate_id": candidate_id,
            "decision_id": decision_id,
        }
        base.update(fields)
        dtw.append_event(base)

    def _emit_snapshot() -> None:
        last = float(entry_row.get("close")) if "close" in entry_row else ""
        vwap_val = float(entry_row.get("vwap")) if "vwap" in entry_row else ""
        _dt_emit(
            "MARKET_SNAPSHOT",
            snap_ts=_iso_ts_jst(entry_ts),
            last=last,
            bid="",
            ask="",
            vwap=vwap_val,
            prev_close=_prev_trading_close(code, day) or "",
        )

    def _emit_features() -> None:
        atr_val = float(entry_row.get("atr")) if "atr" in entry_row else ""
        j_raw = float(J.loc[entry_ts]) if entry_ts in J.index else ""
        _dt_emit(
            "FEATURES_COMPUTED",
            atr_n=int(atr_n),
            atr=atr_val,
            j_raw=j_raw,
            j_bias_adj=0.0,
            j_gap_adj=0.0,
            j=j_raw,
            j_th=float(j_th),
            signal_mode=signal_mode,
            j_cross_state=("NONE" if not signal_mode.lower().startswith("j-cross") else "NONE"),
        )

    _emit_snapshot()
    _emit_features()

    prev_close = _prev_trading_close(code, day)
    if prev_close is not None and prev_close > 0:
        day_open = float(intraday["open"].iloc[0])
        gap_bp = (day_open - prev_close) / prev_close * 10000.0
        if abs(gap_bp) > float(gapban_pct) * 100.0:
            counters["skip_gapban"] = counters.get("skip_gapban", 0) + 1
            _dt_emit(
                "FILTER_EVAL",
                allowed=0,
                deny_reasons="GAP_BAN",
                no_trade_min=int(no_trade_min),
                minutes_from_open=_minutes_from_open(entry_ts, market_open),
                gap_pct=float(gap_bp / 100.0),
                gap_ban_pct=float(gapban_pct),
                trend_driver=trend_driver,
                trend_window=trend_window,
                trend_bp_th=float(trend_bp_th),
                trend_allowed_policy=str(trend_policy or ""),
                trend_aligned="",
            )
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
        _dt_emit(
            "FILTER_EVAL",
            allowed=0,
            deny_reasons="ALLOWED_SIDE",
            no_trade_min=int(no_trade_min),
            minutes_from_open=_minutes_from_open(entry_ts, market_open),
            gap_pct="",
            gap_ban_pct=float(gapban_pct),
            trend_driver=trend_driver,
            trend_window=trend_window,
            trend_bp_th=float(trend_bp_th),
            trend_allowed_policy=str(trend_policy or ""),
            trend_aligned="",
        )
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
            _dt_emit(
                "FILTER_EVAL",
                allowed=0,
                deny_reasons="TREND_MISMATCH",
                no_trade_min=int(no_trade_min),
                minutes_from_open=_minutes_from_open(entry_ts, market_open),
                gap_pct="",
                gap_ban_pct=float(gapban_pct),
                trend_driver=trend_driver,
                trend_window=trend_window,
                trend_bp_th=float(trend_bp_th),
                trend_allowed_policy=policy,
                trend_aligned=0,
            )
            return None

    _dt_emit(
        "FILTER_EVAL",
        allowed=1,
        deny_reasons="",
        no_trade_min=int(no_trade_min),
        minutes_from_open=_minutes_from_open(entry_ts, market_open),
        gap_pct="",
        gap_ban_pct=float(gapban_pct),
        trend_driver=trend_driver,
        trend_window=trend_window,
        trend_bp_th=float(trend_bp_th),
        trend_allowed_policy=policy,
        trend_aligned=1,
    )

    px = float(entry_row["close"])
    atr_val = float(entry_row["atr"])
    if not (np.isfinite(atr_val) and atr_val > 0):
        counters["bad_atr"] = counters.get("bad_atr", 0) + 1
        _dt_emit("ERROR", notes="bad_atr")
        return None

    if side == "BUY":
        tp = px + tpk * atr_val
        sl = px - slk * atr_val
    else:
        tp = px - tpk * atr_val
        sl = px + slk * atr_val

    _dt_emit(
        "DECISION",
        signal=("LONG" if side == "BUY" else "SHORT"),
        action="PLACE",
        entry_style="LIMIT_AHEAD",
        limit_price=float(px),
        qty="",
        tp_price=float(tp),
        sl_price=float(sl),
        trail_type="NONE",
        trail_value="",
        tmax_sec="",
        client_order_id=client_order_id,
        order_status="NEW",
    )

    # simulate bar-by-bar after entry
    after = intraday.loc[entry_ts:]
    bars = 0
    exit_px = None
    exit_ts = None
    exit_reason = ""
    for ts, row in after.iloc[1:].iterrows():
        hi = float(row["high"])
        lo = float(row["low"])
        bars += 1
        if side == "BUY":
            hit_sl = lo <= sl
            hit_tp = hi >= tp
            if hit_sl and hit_tp:
                # 1分足では「どちらが先に触れたか」が分からないため、保守的に損切り扱い
                exit_px = sl
                exit_ts = ts
                exit_reason = "SL_SAME_BAR"
                break
            if hit_sl:
                exit_px = sl
                exit_ts = ts
                exit_reason = "SL"
                break
            if hit_tp:
                exit_px = tp
                exit_ts = ts
                exit_reason = "TP"
                break
        else:
            hit_sl = hi >= sl
            hit_tp = lo <= tp
            if hit_sl and hit_tp:
                exit_px = sl
                exit_ts = ts
                exit_reason = "SL_SAME_BAR"
                break
            if hit_sl:
                exit_px = sl
                exit_ts = ts
                exit_reason = "SL"
                break
            if hit_tp:
                exit_px = tp
                exit_ts = ts
                exit_reason = "TP"
                break

    if exit_px is None:
        # close at last bar of the day
        last_ts = intraday.index[-1]
        exit_ts = last_ts
        exit_px = float(intraday.loc[last_ts, "close"])
        exit_reason = "EOD"

    if side == "BUY":
        pnl_bp = (exit_px - px) / px * 10000.0
    else:
        pnl_bp = (px - exit_px) / px * 10000.0
    pnl_bp -= COST_BP
    nominal_eff = nominal * max(budget_factor, 0.0)
    pnl_yen = nominal_eff * pnl_bp / 10000.0

    if side == "BUY":
        tp_dist_bp = (tp - px) / px * 10000.0
        sl_dist_bp = (px - sl) / px * 10000.0
    else:
        tp_dist_bp = (px - tp) / px * 10000.0
        sl_dist_bp = (sl - px) / px * 10000.0

    if dtw is not None and exit_ts is not None:
        dtw.append_event(
            {
                "event_ts": _iso_ts_jst(exit_ts),
                "event_type": "EXIT",
                "ticker": code,
                "session": session,
                "candidate_id": candidate_id,
                "decision_id": decision_id,
                "client_order_id": client_order_id,
                "order_status": "CLOSED",
                "pnl_realized_yen": float(pnl_yen),
                "notes": str(exit_reason),
            }
        )

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
        "exit_reason": exit_reason,
        "tp_px": tp,
        "sl_px": sl,
        "tp_dist_bp": tp_dist_bp,
        "sl_dist_bp": sl_dist_bp,
        "pnl_bp": pnl_bp,
        "pnl_yen": pnl_yen,
    }


def simulate_day(
    cand_path: Path,
    day: dt.date,
    nominal: float,
    *,
    cooldown_minutes: int = DEFAULT_COOLDOWN_MINUTES,
    max_trades_per_ticker: int = DEFAULT_MAX_TRADES_PER_TICKER,
    stop_after_loss: bool = False,
    dtw: DecisionTraceWriter | None = None,
) -> Tuple[pd.DataFrame, Dict[str, object]]:
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
        cand_id = _make_candidate_id(
            {
                "code": code,
                "session": session,
                "signal_mode": mode,
                "J_th": j_th,
                "TPk": tpk,
                "SLk": slk,
                "ATR_n": atr_n,
                "BudgetFactor_row": float(row.get("BudgetFactor_row", 1.0) or 1.0),
                "NoTradeMin": float(row.get("NoTradeMin") or 5.0),
                "GapBanPct": float(row.get("GapBanPct") or 3.0),
                "trend_driver": str(row.get("trend_driver") or "NKY"),
                "trend_window": str(row.get("trend_window") or "window"),
                "trend_bp_th": float(row.get("trend_bp_th") or 15.0),
                "trend_allowed_policy": str(row.get("trend_allowed_policy") or ""),
            }
        )
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
            dtw=dtw,
            candidate_id=cand_id,
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

    cooldown_minutes_i = int(cooldown_minutes)
    if cooldown_minutes_i < 0:
        cooldown_minutes_i = 0
    cooldown = pd.Timedelta(minutes=cooldown_minutes_i)
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
            if max_trades_per_ticker > 0 and trade_count >= max_trades_per_ticker:
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
            if stop_after_loss:
                try:
                    pnl_yen_val = float(picked.get("pnl_yen", 0.0) or 0.0)
                    if pnl_yen_val < 0:
                        break
                except (TypeError, ValueError):
                    pass

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
        "cooldown_minutes": cooldown_minutes_i,
        "max_trades_per_ticker": int(max_trades_per_ticker),
        "stop_after_loss": bool(stop_after_loss),
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
        "--label",
        default="",
        help=(
            "Optional label suffix for outputs (e.g. M0/M3). "
            "When set, writes per-label outputs to avoid overwriting production files."
        ),
    )
    ap.add_argument(
        "--nominal",
        type=float,
        default=10_000_000.0,
        help="Nominal per position in yen (default 10,000,000)",
    )
    ap.add_argument(
        "--cooldown-minutes",
        type=int,
        default=DEFAULT_COOLDOWN_MINUTES,
        help=f"Cooldown minutes after exit before re-entry (default {DEFAULT_COOLDOWN_MINUTES})",
    )
    ap.add_argument(
        "--max-trades-per-ticker",
        type=int,
        default=DEFAULT_MAX_TRADES_PER_TICKER,
        help=(
            "Max trades per ticker per day. "
            f"Use 0 to disable the limit (default {DEFAULT_MAX_TRADES_PER_TICKER})."
        ),
    )
    ap.add_argument(
        "--stop-after-loss",
        action="store_true",
        help="If set, stop trading the same ticker for the rest of the day after the first losing trade.",
    )
    ap.add_argument(
        "--decision-trace",
        action="store_true",
        help="If set, append DecisionTrace DT.v1 rows to analysis/decision_trace_<date>.csv",
    )
    ap.add_argument(
        "--run-id",
        default="",
        help="Optional run_id for DecisionTrace (default: auto-generated).",
    )
    ap.add_argument(
        "--engine-version",
        default="",
        help="Optional engine_version for DecisionTrace (default: git short sha if available).",
    )
    ap.add_argument(
        "--email",
        action="store_true",
        help="Send the report via SMTP using state/smtp.json.",
    )
    ap.add_argument(
        "--force-email",
        action="store_true",
        help=(
            "Send email even if the daily sent-flag exists. "
            "By default, once-per-day is enforced."
        ),
    )
    ap.add_argument(
        "--recipient",
        default="",
        help="Recipient email address (required when --email).",
    )
    ap.add_argument(
        "--smtp",
        default="state/smtp.json",
        help="SMTP config JSON path (default: state/smtp.json).",
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
        # DailyReplay keeps working even when nightly snapshot generation fails.
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

    dtw: DecisionTraceWriter | None = None
    if bool(getattr(args, "decision_trace", False)):
        safe_label0 = "".join(ch for ch in str(args.label or "") if ch.isalnum() or ch in ("_", "-"))
        run_id = str(getattr(args, "run_id", "") or "").strip()
        if not run_id:
            now = dt.datetime.now(JST).strftime("%H%M%S")
            tag = safe_label0 or "DEFAULT"
            run_id = f"{args.date}_REPLAY_{tag}_{now}"
        engine_version = str(getattr(args, "engine_version", "") or "").strip() or _detect_engine_version()
        dt_path = OUT_DIR / f"decision_trace_{args.date}.csv"
        dt_cols = schema_columns(DT_V1)
        dt_writer = make_append_only_writer(dt_path, schema_version=DT_V1, columns=dt_cols)
        dtw = DecisionTraceWriter(
            writer=dt_writer,
            run_id=run_id,
            env="REPLAY",
            engine="PY",
            engine_version=engine_version,
            trade_date=trade_date.strftime("%Y-%m-%d"),
            source="YAHOO_1M",
        )
    trades_df, summary = simulate_day(
        cand_path,
        trade_date,
        args.nominal,
        cooldown_minutes=int(args.cooldown_minutes),
        max_trades_per_ticker=int(args.max_trades_per_ticker),
        stop_after_loss=bool(args.stop_after_loss),
        dtw=dtw,
    )

    label = str(args.label or "").strip()
    safe_label = "".join(ch for ch in label if ch.isalnum() or ch in ("_", "-"))
    suffix = f"_{safe_label}" if safe_label else ""

    trades_out = OUT_DIR / f"daily_trades_{args.date}{suffix}.csv"
    if not trades_df.empty:
        trades_df.to_csv(trades_out, index=False, encoding="utf-8-sig")
        print(f"written {trades_out} ({len(trades_df)} trades)")
    else:
        print(f"no trades generated for {args.date}")

    summary_csv = OUT_DIR / f"daily_realized_pnl{suffix}.csv"
    if safe_label:
        summary = dict(summary)
        summary["label"] = safe_label
    row = pd.DataFrame([summary])
    if summary_csv.exists():
        try:
            prev = pd.read_csv(summary_csv)
            if "date" in prev.columns:
                mask = prev["date"].astype(str) != summary["date"]
                if "label" in prev.columns and safe_label:
                    mask &= prev["label"].astype(str) != safe_label
                prev = prev[mask]
            combined = pd.concat([prev, row], ignore_index=True)
            combined.to_csv(summary_csv, index=False, encoding="utf-8-sig")
        except Exception:
            row.to_csv(summary_csv, mode="a", header=False, index=False, encoding="utf-8-sig")
    else:
        row.to_csv(summary_csv, index=False, encoding="utf-8-sig")
    print(f"summary appended to {summary_csv}")

    summary_json = OUT_DIR / f"daily_replay_{args.date}{suffix}.json"
    with open(summary_json, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    print(f"summary written to {summary_json}")

    report_txt = OUT_DIR / f"daily_replay_{args.date}{suffix}_mail.txt"
    report = build_mail_report(
        date_tag=args.date,
        cand_path=cand_path,
        trades_df=trades_df,
        summary=summary,
        nominal=args.nominal,
    )
    report_txt.write_text(report, encoding="utf-8-sig")
    print(f"mail report written to {report_txt}")

    if bool(getattr(args, "email", False)):
        recipient = str(getattr(args, "recipient", "") or "").strip()
        smtp_cfg_path = Path(str(getattr(args, "smtp", "state/smtp.json")))
        sent_flag_path = ROOT / "logs" / f"daily_replay_sent_{args.date}.flag"
        if sent_flag_path.exists() and not bool(getattr(args, "force_email", False)):
            print(f"[warn] already sent today; skip email (flag={sent_flag_path})")
            return
        smtp_cfg = _load_smtp_config(smtp_cfg_path)
        if not smtp_cfg:
            print(f"[warn] smtp config not found at {smtp_cfg_path}; email failed")
            raise SystemExit(2)
        elif not recipient:
            print("[warn] --recipient is required when --email; email failed")
            raise SystemExit(2)
        else:
            try:
                subject = f"ASAGAKE DailyReplay {args.date}{suffix}"
                _send_mail_via_smtp(smtp_cfg, recipient, subject, report)
                print(f"Mail sent to {recipient}")
                sent_flag_path.parent.mkdir(parents=True, exist_ok=True)
                sent_flag_path.write_text(
                    f"sent_at={dt.datetime.now(tz=JST).isoformat()} recipient={recipient}\n",
                    encoding="utf-8",
                )
            except Exception as exc:
                print(f"[warn] email send failed: {exc!r}")
                raise SystemExit(2)


if __name__ == "__main__":
    main()
