import argparse
import datetime as dt
import itertools
import json
import math
import os
import concurrent.futures as cf
import subprocess
import time
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple, Optional

import numpy as np
import pandas as pd
from yahooquery import Ticker


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------


def parse_hhmm(value: str) -> dt.time:
    return dt.datetime.strptime(value, "%H:%M").time()


class RunLogger:
    """Simple stdout + file logger."""

    def __init__(self, path: Path):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.fp = self.path.open("a", encoding="utf-8")

    def log(self, msg: str) -> None:
        ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        print(line)
        self.fp.write(line + "\n")
        self.fp.flush()

    def close(self) -> None:
        try:
            self.fp.close()
        except Exception:
            pass


class StatusTracker:
    """Writes heartbeat JSON for progress monitoring."""

    def __init__(self, path: Path, throttle_sec: float = 2.0):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.throttle = throttle_sec
        self.last_write = 0.0
        self._latest: Dict = {}

    def update(self, *, force: bool = False, **payload) -> None:
        self._latest = payload
        now = time.time()
        if not force and now - self.last_write < self.throttle:
            return
        data = dict(payload)
        data["timestamp"] = dt.datetime.now().isoformat()
        tmp = self.path.with_suffix(".tmp")
        with tmp.open("w", encoding="utf-8") as fp:
            json.dump(data, fp, ensure_ascii=False, indent=2)
        tmp.replace(self.path)
        self.last_write = now

    def flush(self) -> None:
        if self._latest:
            self.update(force=True, **self._latest)


# ---------------------------------------------------------------------------
# Data access + filtering
# ---------------------------------------------------------------------------


def load_tickers_from_excel(path: Path) -> List[str]:
    df = pd.read_excel(path, sheet_name="Ticker", usecols="A", header=0)
    return df["Code"].dropna().astype(str).tolist()


def fetch_1m_chunks(
    tickers: Sequence[str],
    *,
    lookback_days: int,
    chunk_days: int,
    logger: RunLogger,
) -> pd.DataFrame:
    if lookback_days <= 0:
        raise ValueError("lookback_days must be positive")
    if chunk_days <= 0:
        raise ValueError("chunk_days must be positive")

    logger.log(f"Fetching Yahoo Finance 1m bars for {len(tickers)} tickers (lookback={lookback_days}, chunk={chunk_days})")
    all_frames: List[pd.DataFrame] = []
    today = dt.datetime.now().date()
    start_base = today - dt.timedelta(days=lookback_days)

    for offset in range(0, lookback_days, chunk_days):
        chunk_start = start_base + dt.timedelta(days=offset)
        chunk_end = start_base + dt.timedelta(days=offset + chunk_days)
        chunk_end = min(chunk_end, today + dt.timedelta(days=1))
        ticker = Ticker(tickers, asynchronous=True)
        try:
            df = ticker.history(start=str(chunk_start), end=str(chunk_end), interval="1m")
        except Exception as exc:
            logger.log(f"Chunk fetch error {chunk_start} - {chunk_end}: {exc}")
            continue
        if df is None or df.empty:
            continue
        df = df.reset_index()
        if "symbol" in df.columns:
            df = df.rename(columns={"symbol": "code"})
        if "date" in df.columns and "ts" not in df.columns:
            df = df.rename(columns={"date": "ts"})
        df["ts"] = pd.to_datetime(df["ts"])
        all_frames.append(df)

    if not all_frames:
        raise RuntimeError("1m data could not be fetched")

    merged = pd.concat(all_frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["ts", "code"])
    merged["date"] = merged["ts"].dt.date
    merged["amt"] = merged["close"] * merged["volume"]
    merged["cumAmt"] = merged.groupby(["code", "date"])["amt"].cumsum()
    merged["cumVol"] = merged.groupby(["code", "date"])["volume"].cumsum().replace(0, np.nan)
    merged["vwap"] = merged["cumAmt"] / merged["cumVol"]
    return merged[["ts", "code", "open", "high", "low", "close", "volume", "vwap", "date", "amt"]]


def apply_liquidity_and_mark_entry(
    df: pd.DataFrame,
    *,
    session_start: dt.time,
    session_end: dt.time,
    liquidity_quantile: float,
    logger: RunLogger,
) -> pd.DataFrame:
    """Keep full-day bars for codes/dates passing liquidity, add is_entry flag.

    - is_entry True if ts within [session_start, session_end]
    - Liquidity computed as daily amt_sum per (date, code) over full day
    - Return all bars (for EOD exit), restricted to selected codes/dates
    """
    if df.empty:
        return df

    df = df.copy()
    df["is_entry"] = df["ts"].dt.time.between(session_start, session_end)
    n_entry = int(df["is_entry"].sum())
    logger.log(f"Entry window marks set: {n_entry} rows flagged out of {len(df)}")

    # attach daily gap in bp per (code,date): (open_first - prev_close) / prev_close * 10000
    df = df.sort_values(["code", "ts"]).reset_index(drop=True)
    daily = (
        df.groupby(["code", "date"]).agg(open_first=("open", "first"), close_last=("close", "last")).reset_index()
    )
    daily = daily.sort_values(["code", "date"]).reset_index(drop=True)
    daily["prev_close"] = daily.groupby("code")["close_last"].shift(1)
    daily["gap_bp"] = ((daily["open_first"] - daily["prev_close"]) / daily["prev_close"]) * 10000.0
    df = df.merge(daily[["code", "date", "gap_bp"]], on=["code", "date"], how="left")

    if liquidity_quantile is None or liquidity_quantile <= 0:
        return df

    liquidity_quantile = min(liquidity_quantile, 0.999)
    by_code_date = df.groupby(["date", "code"])["amt"].sum().reset_index(name="amt_sum")
    thresh = by_code_date.groupby("date")["amt_sum"].quantile(liquidity_quantile)
    by_code_date = by_code_date.merge(
        thresh.rename("threshold"), left_on="date", right_index=True, how="left"
    )
    keep = by_code_date.loc[by_code_date["amt_sum"] >= by_code_date["threshold"], ["date", "code"]]
    kept = df.merge(keep, on=["date", "code"], how="inner")
    logger.log(
        f"Liquidity filter q={liquidity_quantile:.2f} kept {kept['code'].nunique()} tickers and {len(kept)} rows"
    )
    return kept

def build_slices(
    days: Sequence[dt.date],
    train_days: int,
    forward_days: int,
    max_slices: int,
) -> List[Tuple[List[dt.date], List[dt.date]]]:
    slices: List[Tuple[List[dt.date], List[dt.date]]] = []
    if train_days <= 0 or forward_days <= 0:
        return slices
    step = forward_days
    idx = 0
    while idx + train_days + forward_days <= len(days):
        train = list(days[idx : idx + train_days])
        forward = list(days[idx + train_days : idx + train_days + forward_days])
        if not train or not forward:
            break
        slices.append((train, forward))
        if len(slices) >= max_slices:
            break
        idx += step
    return slices


# ---------------------------------------------------------------------------
# Strategy evaluation helpers
# ---------------------------------------------------------------------------


def atr_ema(series: pd.Series, n: int) -> pd.Series:
    return series.ewm(alpha=1 / n, adjust=False).mean()


def make_tr(df: pd.DataFrame) -> pd.Series:
    return pd.concat(
        [
            (df["high"] - df["low"]).abs(),
            (df["high"] - df["close"].shift()).abs(),
            (df["low"] - df["close"].shift()).abs(),
        ],
        axis=1,
    ).max(axis=1)


def build_signal_mask(
    J: pd.Series,
    dJ: pd.Series,
    vE: pd.Series,
    params: Dict[str, float],
    signal_mode: str,
) -> pd.Series:
    j_th = float(params.get("J_th", 0.0))
    dJ_th = float(params.get("dJ_th", 0.0))
    vth = float(params.get("vEMA_th", 0.0))

    j_cond = J.abs() >= j_th

    if signal_mode == "j-only":
        return j_cond

    if signal_mode == "j-cross":
        prev = J.abs().shift(1).fillna(float("inf"))
        cross = prev < j_th
        return j_cond & cross

    if signal_mode != "full":
        raise ValueError(f"Unknown signal_mode={signal_mode}")

    if dJ_th > 0:
        dj_cond = dJ.abs() >= dJ_th
    else:
        dj_cond = pd.Series(True, index=J.index)

    if vth > 0:
        v_cond = vE.abs() >= vth
    else:
        v_cond = pd.Series(True, index=J.index)

    sign_cond = (np.sign(vE) != np.sign(dJ)) | (dJ == 0) | (vE == 0)

    return j_cond & dj_cond & v_cond & sign_cond


BP_EPS = 1e-6


def eval_one(
    df: pd.DataFrame,
    params: Dict[str, float],
    signal_mode: str,
    cost_bp: float,
    abs_guard_bp: float,
    dir_guard_bp: float,
) -> Dict[str, List[float]]:
    if df.empty:
        return {"pnl_bp": []}

    df = df.sort_values("ts").copy().reset_index(drop=True)
    tr = make_tr(df)
    atr = atr_ema(tr, int(params["ATR_n"])).replace(0, np.nan)
    J = (df["close"] - df["vwap"]) / atr
    dJ = J.diff()
    vE = dJ.ewm(alpha=0.3, adjust=False).mean()
    sig = build_signal_mask(J, dJ, vE, params, signal_mode)
    if "is_entry" in df.columns:
        sig = sig & df["is_entry"]
    sig = sig.fillna(False).astype(bool)

    max_bars = int(params.get("TMAX", 0) or 0)
    use_cap = max_bars > 0
    trade_bp: List[float] = []
    grouped_indices = {date: sorted(idxs) for date, idxs in df.groupby("date").indices.items()}

    for date in sorted(grouped_indices.keys()):
        day_positions = grouped_indices[date]
        if not day_positions:
            continue
        day_sig = sig.iloc[day_positions]
        for offset, idx in enumerate(day_positions):
            if not day_sig.iloc[offset]:
                continue
            a = atr.iloc[idx]
            if pd.isna(a) or a == 0:
                continue
            px = df.iloc[idx]["close"]
            if not np.isfinite(px) or px <= 0:
                continue

            side = "BUY" if J.iloc[idx] < 0 else "SELL"
            # Gap guard: absolute or directional skip using df['gap_bp'] (per day constant)
            g = float(df.loc[idx, "gap_bp"]) if "gap_bp" in df.columns else 0.0
            if abs_guard_bp and abs(g) >= abs_guard_bp:
                continue
            if dir_guard_bp and abs(g) >= dir_guard_bp:
                if (g > 0 and side == "SELL") or (g < 0 and side == "BUY"):
                    continue
            tp = px + float(params["TPk"]) * a if side == "BUY" else px - float(params["TPk"]) * a
            sl = px - float(params["SLk"]) * a if side == "BUY" else px + float(params["SLk"]) * a

            future_positions = day_positions[offset + 1 :]
            if not future_positions:
                continue
            if use_cap:
                future_positions = future_positions[:max_bars]
                if not future_positions:
                    continue

            exit_price = None
            for fut_idx in future_positions:
                row = df.iloc[fut_idx]
                high = row["high"]
                low = row["low"]
                if side == "BUY":
                    reached_tp = high >= tp
                    reached_sl = low <= sl
                else:
                    reached_tp = low <= tp
                    reached_sl = high >= sl

                if reached_tp and reached_sl:
                    exit_price = sl
                    break
                if reached_tp:
                    exit_price = tp
                    break
                if reached_sl:
                    exit_price = sl
                    break

            if exit_price is None:
                exit_price = df.iloc[future_positions[-1]]["close"]

            pnl_price = (exit_price - px) if side == "BUY" else (px - exit_price)
            pnl_bp = (pnl_price / px) * 10000.0
            pnl_bp -= cost_bp
            trade_bp.append(float(pnl_bp))

    return {"pnl_bp": trade_bp}


def metrics_from_bp(bp_list: Sequence[float]) -> Dict[str, float]:
    trades = len(bp_list)
    if trades == 0:
        return {
            "wins": 0,
            "losses": 0,
            "flats": 0,
            "trades": 0,
            "winrate": 0.0,
            "pf": 0.0,
            "pf_eff": 0.0,
            "exp_bp": 0.0,
        }

    wins = sum(1 for bp in bp_list if bp > BP_EPS)
    losses = sum(1 for bp in bp_list if bp < -BP_EPS)
    flats = trades - wins - losses

    pos_sum = sum(bp for bp in bp_list if bp > BP_EPS)
    neg_sum = sum(bp for bp in bp_list if bp < -BP_EPS)
    winrate = wins / trades if trades else 0.0

    if neg_sum < -BP_EPS:
        pf = pos_sum / abs(neg_sum) if pos_sum > BP_EPS else 0.0
    else:
        pf = float("inf") if pos_sum > BP_EPS else 0.0

    avg_win = pos_sum / wins if wins else 0.0
    avg_loss = abs(neg_sum) / losses if losses else 0.0
    if avg_loss == 0.0:
        pf_eff = float("inf") if avg_win > 0 else 0.0
    elif winrate >= 1.0:
        pf_eff = float("inf")
    else:
        pf_eff = (winrate * avg_win) / ((1 - winrate) * avg_loss)

    exp_bp = sum(bp_list) / trades
    pf_value = float(pf) if math.isfinite(pf) else 999.0
    pf_eff_value = float(pf_eff) if math.isfinite(pf_eff) else 999.0

    return {
        "wins": wins,
        "losses": losses,
        "flats": flats,
        "trades": trades,
        "winrate": winrate,
        "pf": pf_value,
        "pf_eff": pf_eff_value,
        "exp_bp": exp_bp,
    }


def aggregate_segments(
    segments: Sequence[Dict[str, List[float]]],
    min_slice_trades: int,
) -> Dict[str, float]:
    slice_pf_effs: List[float] = []
    slice_pass = 0
    all_bp: List[float] = []
    slice_exp_bp: List[float] = []

    for seg in segments:
        seg_bp = seg.get("pnl_bp", [])
        all_bp.extend(seg_bp)
        seg_metrics = metrics_from_bp(seg_bp)
        slice_pf_effs.append(seg_metrics["pf_eff"])
        slice_exp_bp.append(seg_metrics["exp_bp"])
        if seg_metrics["trades"] >= min_slice_trades and seg_metrics["pf_eff"] > 1.0:
            slice_pass += 1

    metrics = metrics_from_bp(all_bp)
    metrics.update(
        {
            "slices_total": len(segments),
            "slices_pass": slice_pass,
            "slice_pf_eff": slice_pf_effs,
            "slice_exp_bp": slice_exp_bp,
        }
    )
    return metrics


def evaluate_params_for_code(
    code_df: pd.DataFrame,
    slices: Sequence[Tuple[Sequence[dt.date], Sequence[dt.date]]],
    params: Dict[str, float],
    cost_bp: float,
    abs_guard_bp: float,
    dir_guard_bp: float,
    min_train_trades: int,
    min_forward_trades: int,
    signal_mode: str,
) -> Tuple[Dict[str, float], Dict[str, float]]:
    train_segments: List[Dict[str, List[float]]] = []
    forward_segments: List[Dict[str, List[float]]] = []
    for train_dates, forward_dates in slices:
        train_df = code_df[code_df["date"].isin(train_dates)]
        forward_df = code_df[code_df["date"].isin(forward_dates)]
        train_segments.append(eval_one(train_df, params, signal_mode, cost_bp, abs_guard_bp, dir_guard_bp))
        forward_segments.append(eval_one(forward_df, params, signal_mode, cost_bp, abs_guard_bp, dir_guard_bp))

    train_metrics = aggregate_segments(train_segments, min_train_trades)
    forward_metrics = aggregate_segments(forward_segments, min_forward_trades)
    return train_metrics, forward_metrics


def build_param_grid(mode: str, signal_mode: str) -> Dict[str, List[float]]:
    if mode == "coarse":
        grid = {
            "ATR_n": [1, 3, 5],
            "TPk": [1.2, 1.5],
            "SLk": [1.0, 1.5, 2.0],
            "J_th": [0.4, 0.6, 0.8, 1.0],
            "dJ_th": [0.01, 0.03, 0.05],
            "vEMA_th": [0.01, 0.03, 0.05],
            "TMAX": [0],
        }
    else:
        grid = {
            "ATR_n": [1, 2, 3, 5, 8],
            "TPk": [1.0, 1.2, 1.5, 2.0, 2.5],
            "SLk": [0.8, 1.0, 1.2, 1.5, 2.0],
            "J_th": [0.3, 0.4, 0.6, 0.8, 1.0, 1.2],
            "dJ_th": [0.0, 0.01, 0.03, 0.05],
            "vEMA_th": [0.0, 0.01, 0.03, 0.05],
            "TMAX": [0],
        }

    if signal_mode != "full":
        grid["dJ_th"] = [0.0]
        grid["vEMA_th"] = [0.0]

    return grid


def wilson_ci(k: int, n: int, z: float = 1.96) -> Tuple[float, float]:
    if n <= 0:
        return 0.0, 0.0
    p = (k / n) if n else 0.0
    denom = 1 + (z**2) / n
    center = (p + (z**2) / (2 * n)) / denom
    half = (z * math.sqrt((p * (1 - p) / n) + (z**2) / (4 * (n**2)))) / denom
    lo = max(0.0, center - half)
    hi = min(1.0, center + half)
    return lo, hi


def bootstrap_mean_ci(values: Sequence[float], n_boot: int = 300, alpha: float = 0.05) -> Tuple[float, float, float]:
    vals = [float(v) for v in values if v is not None and math.isfinite(float(v))]
    if len(vals) == 0:
        return 0.0, 0.0, 0.0
    rng = np.random.default_rng(12345)
    boots = []
    for _ in range(n_boot):
        smp = rng.choice(vals, size=len(vals), replace=True)
        boots.append(float(np.mean(smp)))
    boots.sort()
    mean = float(np.mean(boots))
    lo = boots[int((alpha / 2) * (len(boots) - 1))]
    hi = boots[int((1 - alpha / 2) * (len(boots) - 1))]
    return mean, lo, hi


def iter_param_dicts(param_spec: Dict[str, Sequence[float]]) -> Iterable[Dict[str, float]]:
    keys = list(param_spec.keys())
    for combo in itertools.product(*(param_spec[k] for k in keys)):
        yield dict(zip(keys, combo))


def score_record(
    rec: Dict[str, float],
    min_forward_trades: int,
    forward_pf_min: float,
) -> Tuple:
    fw_ok = (
        rec["forward_trades"] >= min_forward_trades
        and rec["forward_pf_eff"] >= forward_pf_min
    )
    return (
        1 if fw_ok else 0,
        rec["forward_pf_eff"] if fw_ok else 0.0,
        rec["forward_winrate"] if fw_ok else 0.0,
        rec["forward_trades"],
        rec["train_pf_eff"],
        rec["train_winrate"],
        rec["train_trades"],
    )


def safe_to_float(value: float) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        if math.isnan(value):
            return 0.0
        return float(value)
    return float(value)


def write_csv(path: Path, df: pd.DataFrame) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    df.to_csv(tmp, index=False, encoding="utf-8-sig")
    tmp.replace(path)


def _eval_code_task(payload) -> Tuple[str, pd.DataFrame, Dict[str, float]]:
    (
        code,
        code_df,
        slices,
        param_grid,
        cost_bp,
        abs_guard_bp,
        dir_guard_bp,
        min_train_trades,
        min_forward_trades,
        signal_mode,
        mode,
        session_label,
        bootstrap_n,
        forward_pf_min,
    ) = payload
    code_records: List[Dict[str, float]] = []
    for param in param_grid:
        train_metrics, forward_metrics = evaluate_params_for_code(
            code_df,
            slices,
            param,
            cost_bp,
            abs_guard_bp,
            dir_guard_bp,
            min_train_trades,
            min_forward_trades,
            signal_mode,
        )
        fw_wins = int(forward_metrics["wins"])
        fw_trades = int(forward_metrics["trades"])
        wlo, whi = wilson_ci(fw_wins, fw_trades)
        bmean, blo, bhi = bootstrap_mean_ci(
            forward_metrics.get("slice_exp_bp", []), n_boot=int(bootstrap_n)
        )
        record = {
            "mode": mode,
            "signal_mode": signal_mode,
            "session": session_label,
            "code": code,
            "ATR_n": param["ATR_n"],
            "TPk": param["TPk"],
            "SLk": param["SLk"],
            "J_th": param["J_th"],
            "dJ_th": param["dJ_th"],
            "vEMA_th": param["vEMA_th"],
            "TMAX": param["TMAX"],
            "train_wins": train_metrics["wins"],
            "train_losses": train_metrics["losses"],
            "train_flats": train_metrics["flats"],
            "train_trades": train_metrics["trades"],
            "train_winrate": train_metrics["winrate"],
            "train_pf": train_metrics["pf"],
            "train_pf_eff": train_metrics["pf_eff"],
            "train_exp_bp": train_metrics["exp_bp"],
            "train_slices_pass": train_metrics["slices_pass"],
            "train_slices_total": train_metrics["slices_total"],
            "forward_wins": forward_metrics["wins"],
            "forward_losses": forward_metrics["losses"],
            "forward_flats": forward_metrics["flats"],
            "forward_trades": forward_metrics["trades"],
            "forward_winrate": forward_metrics["winrate"],
            "forward_pf": forward_metrics["pf"],
            "forward_pf_eff": forward_metrics["pf_eff"],
            "forward_exp_bp": forward_metrics["exp_bp"],
            "forward_slices_pass": forward_metrics["slices_pass"],
            "forward_slices_total": forward_metrics["slices_total"],
            "forward_win_ci_low": wlo,
            "forward_win_ci_high": whi,
            "forward_exp_boot_mean": bmean,
            "forward_exp_boot_low": blo,
            "forward_exp_boot_high": bhi,
        }
        code_records.append(record)
    df_code = pd.DataFrame(code_records)
    if df_code.empty:
        best_record: Dict[str, float] = {
            "code": code,
            "train_pf_eff": 0.0,
            "forward_pf_eff": 0.0,
            "train_trades": 0,
            "forward_trades": 0,
        }
    else:
        best_record = max(
            df_code.to_dict("records"),
            key=lambda rec: score_record(rec, min_forward_trades, forward_pf_min),
        )
    return code, df_code, best_record


# ---------------------------------------------------------------------------
# Git helper (optional)
# ---------------------------------------------------------------------------


def git_push_results(repo_dir: Path, logger: RunLogger) -> None:
    proc = subprocess.run(
        ["git", "rev-parse", "--abbrev-ref", "HEAD"],
        cwd=repo_dir,
        capture_output=True,
        text=True,
    )
    if proc.returncode != 0:
        err = proc.stderr or proc.stdout
        logger.log(f"git rev-parse failed: {err}")
        return
    branch = proc.stdout.strip()
    cmd = (
        "git add output/bt30/bt30_*/*.csv output/bt30/bt30_*/*.json && "
        "git commit -m \"Add bt30 results\" && "
        f"git push origin {branch}"
    )
    p2 = subprocess.run(cmd, cwd=repo_dir, shell=True, capture_output=True, text=True)
    if p2.returncode != 0:
        err = p2.stderr or p2.stdout
        logger.log(f"Push failed branch={branch}: {err}")
    else:
        logger.log("Git push succeeded")


# ---------------------------------------------------------------------------
# Main execution
# ---------------------------------------------------------------------------


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="Path to SHINSOKU workbook")
    ap.add_argument("--outdir", required=True, help="Output directory for run artifacts")
    ap.add_argument("--mode", choices=["coarse", "refine"], default="coarse")
    ap.add_argument("--slipbp", type=float, default=4.0)
    ap.add_argument("--feebp", type=float, default=4.0)
    ap.add_argument("--lookback", type=int, default=60)
    ap.add_argument("--chunk-days", type=int, default=5)
    ap.add_argument("--train-days", type=int, default=16)
    ap.add_argument("--forward-days", type=int, default=5)
    ap.add_argument("--forward-slices", type=int, default=5)
    ap.add_argument("--session-start", default="09:00")
    ap.add_argument("--session-end", default="09:15")
    ap.add_argument("--liquidity-quantile", type=float, default=0.5)
    ap.add_argument("--signal-mode", choices=["full", "j-only", "j-cross"], default="full")
    ap.add_argument("--min-train-trades", type=int, default=30)
    ap.add_argument("--min-forward-trades", type=int, default=10)
    ap.add_argument("--train-pf-min", type=float, default=0.8)
    ap.add_argument("--forward-pf-min", type=float, default=1.3)
    ap.add_argument("--train-exp-min", type=float, default=0.0)
    ap.add_argument("--forward-exp-min", type=float, default=0.0)
    ap.add_argument("--forward-pass-ratio", type=float, default=0.5)
    ap.add_argument("--gap-guard-abs-bp", type=float, default=80.0, help="Skip entire day if abs gap >= bp")
    ap.add_argument("--gap-guard-dir-bp", type=float, default=40.0, help="Skip inverse-direction trades if abs gap >= bp")
    ap.add_argument("--codes-file", help="Optional CSV with column 'code' to limit evaluation universe")
    ap.add_argument("--top-n", type=int, default=30, help="Number of codes to pass from coarse to refine stage")
    ap.add_argument("--status-file", help="Optional custom status.json path")
    ap.add_argument("--log-file", help="Optional custom log path")
    ap.add_argument("--stop-file", help="Path to stop-file; if exists run halts gracefully")
    ap.add_argument("--push", action="store_true", help="If set, git add/commit/push results when complete")
    ap.add_argument("--candidate-dir", default="output/excel", help="Directory for Excel candidate CSV")
    ap.add_argument("--excel-summary", action="store_true", help="Write Excel workbook summary (summary_*.xlsx)")
    ap.add_argument("--grid-sample-size", type=int, default=200, help="Rows to include when sampling grid into Excel summary")
    ap.add_argument("--bootstrap-n", type=int, default=300, help="Bootstrap resamples for CI")
    ap.add_argument("--jobs", type=int, default=0, help="Parallel workers across codes (0=auto)")
    args = ap.parse_args()

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    status_path = Path(args.status_file) if args.status_file else outdir / "status.json"
    log_path = Path(args.log_file) if args.log_file else Path("logs") / f"opt30_run_{dt.datetime.now():%Y%m%d_%H%M%S}.log"
    stop_path = Path(args.stop_file) if args.stop_file else outdir / "STOP"

    logger = RunLogger(log_path)
    tracker = StatusTracker(status_path)
    start_ts = time.time()
    logger.log("Run started")

    tickers = load_tickers_from_excel(Path(args.excel))
    if args.codes_file:
        code_df = pd.read_csv(args.codes_file)
        allow = set(code_df["code"].astype(str).tolist())
        tickers = [t for t in tickers if t in allow]
        logger.log(f"Codes limited via {args.codes_file}: {len(tickers)} tickers")
    if not tickers:
        raise RuntimeError("No tickers available for evaluation")

    data = fetch_1m_chunks(
        tickers,
        lookback_days=args.lookback,
        chunk_days=args.chunk_days,
        logger=logger,
    )

    days = sorted(data["date"].unique())
    if len(days) < args.train_days + args.forward_days:
        raise RuntimeError(
            f"Insufficient distinct days ({len(days)}) for train={args.train_days} + forward={args.forward_days}"
        )

    session_start = parse_hhmm(args.session_start)
    session_end = parse_hhmm(args.session_end)
    filtered = apply_liquidity_and_mark_entry(
        data,
        session_start=session_start,
        session_end=session_end,
        liquidity_quantile=args.liquidity_quantile,
        logger=logger,
    )
    if filtered.empty:
        raise RuntimeError("All rows dropped by filters; adjust session or liquidity settings")

    slices = build_slices(
        sorted(filtered["date"].unique()),
        train_days=args.train_days,
        forward_days=args.forward_days,
        max_slices=args.forward_slices,
    )
    if not slices:
        raise RuntimeError("Unable to construct train/forward slices with current day counts")
    logger.log(f"Using {len(slices)} rolling slices (train={args.train_days}, forward={args.forward_days})")

    param_spec = build_param_grid(args.mode, args.signal_mode)
    param_grid = list(iter_param_dicts(param_spec))
    logger.log(f"Parameter grid size: {len(param_grid)} (mode={args.mode}, signal={args.signal_mode})")

    cost_bp = float(args.slipbp + args.feebp)
    summary_records: List[Dict[str, float]] = []
    grid_columns = None
    summary_partial_path = outdir / "_SUMMARY_TRAIN.partial.csv"
    top_partial_path = outdir / "_TOP_FORWARD.partial.csv"
    grid_path = outdir / "_GRID_FULL.csv"

    if grid_path.exists():
        grid_path.unlink()
    if summary_partial_path.exists():
        summary_partial_path.unlink()
    if top_partial_path.exists():
        top_partial_path.unlink()

    grouped = filtered.groupby("code")
    codes = list(grouped.groups.keys())
    total_tasks = len(codes) * len(param_grid)
    combo_done = 0
    session_label = "AM10" if session_end <= parse_hhmm("09:10") else "AM15"

    max_workers = args.jobs if args.jobs and args.jobs > 0 else max(1, (os.cpu_count() or 2) - 1)
    payloads = []
    for code in codes:
        payloads.append((
            code,
            grouped.get_group(code),
            slices,
            param_grid,
            cost_bp,
            float(getattr(args, "gap_guard_abs_bp", 0.0)),
            float(getattr(args, "gap_guard_dir_bp", 0.0)),
            args.min_train_trades,
            args.min_forward_trades,
            args.signal_mode,
            args.mode,
            session_label,
            int(getattr(args, "bootstrap_n", 300)),
            args.forward_pf_min,
        ))

    idx = 0
    with cf.ThreadPoolExecutor(max_workers=max_workers) as ex:
        futures = [ex.submit(_eval_code_task, p) for p in payloads]
        for fut in cf.as_completed(futures):
            code, df_code, best_record = fut.result()
            idx += 1
            combo_done += len(param_grid)
            if not df_code.empty:
                write_mode = "w" if not grid_path.exists() else "a"
                df_code.to_csv(
                    grid_path,
                    mode=write_mode,
                    header=write_mode == "w",
                    index=False,
                    encoding="utf-8-sig",
                )
            summary_records.append(best_record)
            summary_df = pd.DataFrame(summary_records)
            write_csv(summary_partial_path, summary_df)
            top_df = summary_df[
                (summary_df["train_pf_eff"] >= args.train_pf_min)
                & (summary_df["train_trades"] >= args.min_train_trades)
                & (summary_df["forward_pf_eff"] >= args.forward_pf_min)
                & (summary_df["forward_trades"] >= args.min_forward_trades)
                & (summary_df["train_exp_bp"] >= args.train_exp_min)
                & (summary_df["forward_exp_bp"] >= args.forward_exp_min)
            ].sort_values(
                ["forward_pf_eff", "forward_trades", "train_pf_eff"],
                ascending=[False, False, False],
            )
            write_csv(top_partial_path, top_df)
            progress = combo_done / total_tasks if total_tasks else 0.0
            elapsed = time.time() - start_ts
            eta = (elapsed / progress - elapsed) if progress > 0 else None
            tracker.update(
                phase=f"{args.mode}-eval",
                code=code,
                combo_done=combo_done,
                combo_total=total_tasks,
                ticker_index=idx,
                ticker_total=len(codes),
                elapsed_seconds=elapsed,
                eta_seconds=eta if eta is not None else None,
            )
            logger.log(
                f"[{idx}/{len(codes)}] {code}: best train_pf_eff={best_record.get('train_pf_eff', 0):.3f}, "
                f"forward_pf_eff={best_record.get('forward_pf_eff', 0):.3f}, trades={best_record.get('train_trades', 0)}/{best_record.get('forward_trades', 0)}"
            )

    tracker.update(
        phase="completed",
        elapsed_seconds=time.time() - start_ts,
        combo_done=combo_done,
        combo_total=total_tasks,
    )
    tracker.flush()

    if not summary_records:
        logger.log("No summary records produced")
        logger.close()
        return

    summary_df = pd.DataFrame(summary_records)
    summary_cols = [
        "code",
        "mode",
        "signal_mode",
        "ATR_n",
        "TPk",
        "SLk",
        "J_th",
        "dJ_th",
        "vEMA_th",
        "TMAX",
        "train_trades",
        "train_winrate",
        "train_pf",
        "train_pf_eff",
        "train_exp_bp",
        "train_slices_pass",
        "train_slices_total",
    ]
    forward_cols = [
        "code",
        "mode",
        "signal_mode",
        "forward_trades",
        "forward_winrate",
        "forward_pf",
        "forward_pf_eff",
        "forward_exp_bp",
        "forward_slices_pass",
        "forward_slices_total",
    ]
    write_csv(outdir / "_SUMMARY_TRAIN.csv", summary_df[summary_cols])
    write_csv(outdir / "_SUMMARY_FORWARD.csv", summary_df[forward_cols])
    write_csv(outdir / "_COMPARE.csv", summary_df)

    top_df = summary_df[
        (summary_df["train_pf_eff"] >= args.train_pf_min)
        & (summary_df["train_trades"] >= args.min_train_trades)
        & (summary_df["forward_pf_eff"] >= args.forward_pf_min)
        & (summary_df["forward_trades"] >= args.min_forward_trades)
    ].sort_values(
        ["forward_pf_eff", "forward_trades", "train_pf_eff"],
        ascending=[False, False, False],
    )
    write_csv(outdir / "_TOP_CANDIDATES.csv", top_df)

    if args.mode == "refine":
        required_pass = summary_df["forward_slices_total"].apply(
            lambda s: max(1, math.ceil(args.forward_pass_ratio * s))
        )
        final_candidates = summary_df[
            (summary_df["train_pf_eff"] >= args.train_pf_min)
            & (summary_df["train_trades"] >= args.min_train_trades)
            & (summary_df["forward_pf_eff"] >= args.forward_pf_min)
            & (summary_df["forward_trades"] >= args.min_forward_trades)
            & (summary_df["forward_slices_pass"] >= required_pass)
            & (summary_df["train_exp_bp"] >= args.train_exp_min)
            & (summary_df["forward_exp_bp"] >= args.forward_exp_min)
        ].copy()
        final_candidates = final_candidates.sort_values(
            ["forward_pf_eff", "forward_trades", "train_pf_eff"],
            ascending=[False, False, False],
        )
        candidate_dir = Path(args.candidate_dir)
        candidate_dir.mkdir(parents=True, exist_ok=True)
        candidate_path = candidate_dir / f"candidates_{dt.datetime.now():%Y%m%d}.csv"
        export_cols = [
            "code",
            "ATR_n",
            "TPk",
            "SLk",
            "J_th",
            "dJ_th",
            "vEMA_th",
            "TMAX",
            "session",
            "train_winrate",
            "train_pf_eff",
            "train_trades",
            "forward_winrate",
            "forward_pf_eff",
            "forward_trades",
            "forward_slices_pass",
            "forward_slices_total",
            "forward_win_ci_low",
            "forward_win_ci_high",
            "forward_exp_boot_mean",
            "forward_exp_boot_low",
            "forward_exp_boot_high",
        ]
        final_out = final_candidates[export_cols].rename(columns={"code": "Ticker"})
        final_out.insert(0, "Selected", 1)
        write_csv(candidate_path, final_out)
        logger.log(f"Refine candidates exported to {candidate_path}")

    if args.excel_summary:
        excel_path = outdir / f"summary_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
        try:
            writer = pd.ExcelWriter(excel_path, engine="openpyxl")
            with writer:
                top_forward = summary_df.sort_values(
                    ["forward_pf_eff", "forward_trades", "train_pf_eff"],
                    ascending=[False, False, False],
                ).head(100)
                top_train = summary_df.sort_values(
                    ["train_pf_eff", "train_trades", "forward_pf_eff"],
                    ascending=[False, False, False],
                ).head(100)
                compare_sample = summary_df.head(min(len(summary_df), 500))
                top_forward.to_excel(writer, sheet_name="TopForward", index=False)
                top_train.to_excel(writer, sheet_name="TopTrain", index=False)
                compare_sample.to_excel(writer, sheet_name="Compare", index=False)

                if outdir.joinpath("_GRID_FULL.csv").exists():
                    grid_df = pd.read_csv(outdir / "_GRID_FULL.csv")
                    grid_sample = grid_df.sort_values(
                        ["forward_pf_eff", "forward_trades", "train_pf_eff"],
                        ascending=[False, False, False],
                    ).head(args.grid_sample_size)
                    grid_sample.to_excel(writer, sheet_name="GridSamples", index=False)

                if args.mode == "refine":
                    final_sheet = final_candidates if "final_candidates" in locals() else pd.DataFrame()
                    final_sheet.to_excel(writer, sheet_name="FinalCandidates", index=False)
                else:
                    top_df.to_excel(writer, sheet_name="TopCandidates", index=False)
            logger.log(f"Excel summary written to {excel_path}")
        except Exception as exc:
            logger.log(f"Failed to create Excel summary ({excel_path}): {exc}")

    elapsed_total = time.time() - start_ts
    logger.log(f"Run completed in {elapsed_total/60:.1f} minutes")

    if args.push:
        logger.log("Attempting git push...")
        git_push_results(Path(__file__).resolve().parent.parent, logger)

    logger.close()


if __name__ == "__main__":
    main()
