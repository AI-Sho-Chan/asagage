import argparse
import datetime as dt
import json
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd


BP_EPS = 1e-6
ABS_GUARD_BP = 80.0
DIR_GUARD_BP = 40.0
VWAP_EPS = 1e-9

WINDOWS: List[Tuple[str, str]] = [
    ("09:00", "09:30"),
    ("09:00", "10:00"),
    ("09:00", "11:30"),
    ("09:00", "15:30"),
]

ALPHA_GRID = [0.2, 0.4, 0.6]
GAMMA_GRID = [0.05, 0.10, 0.20]
STANDBY_MAX_BARS = [5, 10, 20]

GAP_BUCKETS: List[Tuple[str, float, float]] = [
    ("<50bp", 0.0, 50.0),
    ("50-80bp", 50.0, 80.0),
    ("80-120bp", 80.0, 120.0),
    (">=120bp", 120.0, float("inf")),
]

GAP_RULE_TABLE = [
    {
        "label": "<50bp",
        "abs_min": 0.0,
        "abs_max": 50.0,
        "settings": {
            "j-only": {
                "enabled": True,
                "j_th_add": 0.0,
                "tp_add": 0.0,
                "sl_add": 0.0,
                "skip_opposite": False,
            },
            "j-cross": {
                "enabled": True,
                "j_th_add": 0.0,
                "tp_add": 0.0,
                "sl_add": 0.0,
                "skip_opposite": False,
            },
        },
    },
    {
        "label": "50-80bp",
        "abs_min": 50.0,
        "abs_max": 80.0,
        "settings": {
            "j-only": {
                "enabled": True,
                "j_th_add": 0.0,
                "tp_add": 0.0,
                "sl_add": 0.0,
                "skip_opposite": False,
            },
            "j-cross": {
                "enabled": True,
                "j_th_add": 0.0,
                "tp_add": 0.0,
                "sl_add": 0.0,
                "skip_opposite": False,
            },
        },
    },
    {
        "label": "80-120bp",
        "abs_min": 80.0,
        "abs_max": 120.0,
        "settings": {
            "j-only": {
                "enabled": True,
                "j_th_add": 0.2,
                "tp_add": 0.0,
                "sl_add": 0.1,
                "skip_opposite": True,
            },
            "j-cross": {
                "enabled": True,
                "j_th_add": 0.2,
                "tp_add": 0.0,
                "sl_add": 0.1,
                "skip_opposite": True,
            },
        },
    },
    {
        "label": ">=120bp",
        "abs_min": 120.0,
        "abs_max": float("inf"),
        "settings": {
            "j-only": {
                "enabled": False,
                "j_th_add": 0.3,
                "tp_add": -0.2,
                "sl_add": 0.2,
                "skip_opposite": True,
            },
            "j-cross": {
                "enabled": True,
                "j_th_add": 0.3,
                "tp_add": -0.2,
                "sl_add": 0.2,
                "skip_opposite": True,
            },
        },
    },
]

GAP_RULE_INDEX = {entry["label"]: entry for entry in GAP_RULE_TABLE}


@dataclass
class TradeResult:
    ticker: str
    signal_mode: str
    window: str
    method: str
    params_id: str
    alpha: Optional[float]
    gamma: Optional[float]
    standby_max: Optional[int]
    trades: int
    wins: int
    losses: int
    flats: int
    winrate: float
    pf: float
    pf_eff: float
    exp_bp: float
    total_bp: float
    avg_bars: float
    pnl_list: List[float]


def parse_time(label: str) -> dt.time:
    return dt.datetime.strptime(label, "%H:%M").time()


def load_candidates(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Candidate file not found: {path}")
    df = pd.read_csv(path)
    if "Ticker" not in df.columns:
        raise ValueError("Candidate CSV missing 'Ticker' column")
    return df.copy()


def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.columns, pd.MultiIndex):
        level0 = df.columns.get_level_values(0).tolist()
        df.columns = [c.lower() for c in level0]
    else:
        df.columns = [c.lower() for c in df.columns]
    return df


def load_intraday_data(ticker: str, root: Path) -> pd.DataFrame:
    folder = root / ticker
    if not folder.exists():
        raise FileNotFoundError(f"Minute data directory not found for {ticker}: {folder}")

    frames: List[pd.DataFrame] = []
    for pq in sorted(folder.glob("*.parquet")):
        df = pd.read_parquet(pq)
        df = flatten_columns(df)
        df = df.rename(
            columns={
                "adj close": "adj_close",
                "close": "close",
                "open": "open",
                "high": "high",
                "low": "low",
                "volume": "volume",
            }
        )
        df = df[["open", "high", "low", "close", "volume"]]
        frames.append(df)

    if not frames:
        raise RuntimeError(f"No parquet files found for {ticker} under {folder}")

    data = pd.concat(frames).sort_index()
    data = data[~data.index.duplicated(keep="last")]
    data["ticker"] = ticker
    return data


def prepare_dataset(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["date"] = df.index.date
    df["time"] = df.index.time
    df["amt"] = df["close"] * df["volume"]
    df["cum_amt"] = df.groupby("date")["amt"].cumsum()
    df["cum_vol"] = df.groupby("date")["volume"].cumsum()
    df["vwap"] = df["cum_amt"] / df["cum_vol"].replace(0, np.nan)
    df["prev_close"] = df["close"].shift(1)
    daily = (
        df.groupby("date")
        .agg(open_first=("open", "first"), close_last=("close", "last"))
        .reset_index()
        .sort_values("date")
    )
    daily["prev_close"] = daily["close_last"].shift(1)
    daily["gap_bp"] = (
        (daily["open_first"] - daily["prev_close"]) / daily["prev_close"]
    ) * 10000.0
    df = df.merge(daily[["date", "gap_bp"]], on="date", how="left")
    tr = pd.concat(
        [
            (df["high"] - df["low"]).abs(),
            (df["high"] - df["close"].shift()).abs(),
            (df["low"] - df["close"].shift()).abs(),
        ],
        axis=1,
    ).max(axis=1)
    df["tr"] = tr
    return df


def atr_ema(series: pd.Series, n: int) -> pd.Series:
    return series.ewm(alpha=1 / n, adjust=False).mean()


def build_gap_settings(gap_bp: float, signal_mode: str) -> Dict[str, float]:
    gap_abs = abs(gap_bp) if math.isfinite(gap_bp) else float("inf")
    for entry in GAP_RULE_TABLE:
        if entry["abs_min"] <= gap_abs < entry["abs_max"]:
            settings = entry["settings"].get(signal_mode, {})
            return {
                "bucket": entry["label"],
                "enabled": bool(settings.get("enabled", True)),
                "j_th_add": float(settings.get("j_th_add", 0.0)),
                "tp_add": float(settings.get("tp_add", 0.0)),
                "sl_add": float(settings.get("sl_add", 0.0)),
                "skip_opposite": bool(settings.get("skip_opposite", False)),
            }
    default = GAP_RULE_TABLE[0]["settings"].get(signal_mode, {})
    return {
        "bucket": GAP_RULE_TABLE[0]["label"],
        "enabled": bool(default.get("enabled", True)),
        "j_th_add": float(default.get("j_th_add", 0.0)),
        "tp_add": float(default.get("tp_add", 0.0)),
        "sl_add": float(default.get("sl_add", 0.0)),
        "skip_opposite": bool(default.get("skip_opposite", False)),
    }


def metrics_from_pnl(pnl_list: Sequence[float]) -> Dict[str, float]:
    trades = len(pnl_list)
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
            "total_bp": 0.0,
        }

    wins = sum(1 for bp in pnl_list if bp > BP_EPS)
    losses = sum(1 for bp in pnl_list if bp < -BP_EPS)
    flats = trades - wins - losses
    pos_sum = sum(bp for bp in pnl_list if bp > BP_EPS)
    neg_sum = sum(bp for bp in pnl_list if bp < -BP_EPS)
    total = sum(pnl_list)

    if neg_sum < -BP_EPS:
        pf = pos_sum / abs(neg_sum) if pos_sum > BP_EPS else 0.0
    else:
        pf = float("inf") if pos_sum > BP_EPS else 0.0

    winrate = wins / trades if trades else 0.0
    avg_win = pos_sum / wins if wins else 0.0
    avg_loss = abs(neg_sum) / losses if losses else 0.0
    if avg_loss == 0.0:
        pf_eff = float("inf") if avg_win > 0 else 0.0
    elif winrate >= 1.0:
        pf_eff = float("inf")
    else:
        pf_eff = (winrate * avg_win) / ((1 - winrate) * avg_loss)

    return {
        "wins": wins,
        "losses": losses,
        "flats": flats,
        "trades": trades,
        "winrate": winrate,
        "pf": float(pf) if math.isfinite(pf) else 999.0,
        "pf_eff": float(pf_eff) if math.isfinite(pf_eff) else 999.0,
        "exp_bp": total / trades,
        "total_bp": total,
    }


def simulate_trade(
    day_df: pd.DataFrame,
    day_positions: List[pd.Timestamp],
    index_pos: int,
    atr: pd.Series,
    J: pd.Series,
    params: Dict[str, float],
    gap_value: float,
    gap_rule: Dict[str, float],
    cost_bp: float,
) -> Optional[Dict[str, float]]:
    idx = day_positions[index_pos]
    a = atr.loc[idx]
    if not math.isfinite(a) or a <= 0:
        return None

    price = day_df.loc[idx, "close"]
    if not math.isfinite(price) or price <= 0:
        return None

    if not gap_rule["enabled"]:
        return None

    if ABS_GUARD_BP and math.isfinite(gap_value) and abs(gap_value) >= ABS_GUARD_BP:
        return None

    side = "BUY" if J.loc[idx] < 0 else "SELL"

    if DIR_GUARD_BP and math.isfinite(gap_value) and abs(gap_value) >= DIR_GUARD_BP:
        if (gap_value > 0 and side == "SELL") or (gap_value < 0 and side == "BUY"):
            return None

    if gap_rule.get("skip_opposite", False):
        if (gap_value > 0 and side == "SELL") or (gap_value < 0 and side == "BUY"):
            return None

    eff_tp = max(0.2, float(params["TPk"]) + float(gap_rule.get("tp_add", 0.0)))
    eff_sl = max(0.2, float(params["SLk"]) + float(gap_rule.get("sl_add", 0.0)))

    tp = price + eff_tp * a if side == "BUY" else price - eff_tp * a
    sl = price - eff_sl * a if side == "BUY" else price + eff_sl * a

    future_positions = day_positions[index_pos + 1 :]
    if not future_positions:
        return None

    tmax = int(params.get("TMAX", 0) or 0)
    if tmax > 0:
        future_positions = future_positions[:tmax]
        if not future_positions:
            return None

    exit_price = None
    exit_pos = None
    for j, fut_idx in enumerate(future_positions, start=1):
        row = day_df.loc[fut_idx]
        high = row["high"]
        low = row["low"]
        if side == "BUY":
            hit_tp = high >= tp - VWAP_EPS
            hit_sl = low <= sl + VWAP_EPS
        else:
            hit_tp = low <= tp + VWAP_EPS
            hit_sl = high >= sl - VWAP_EPS

        if hit_tp and hit_sl:
            exit_price = sl
            exit_pos = j
            break
        if hit_tp:
            exit_price = tp
            exit_pos = j
            break
        if hit_sl:
            exit_price = sl
            exit_pos = j
            break

    if exit_price is None:
        last_idx = future_positions[-1]
        exit_price = day_df.loc[last_idx, "close"]
        exit_pos = len(future_positions)

    pnl_price = (exit_price - price) if side == "BUY" else (price - exit_price)
    pnl_bp = (pnl_price / price) * 10000.0
    pnl_bp -= cost_bp

    return {
        "pnl_bp": pnl_bp,
        "bars": exit_pos,
    }


def baseline_entries(
    day_df: pd.DataFrame,
    day_positions: List[pd.Timestamp],
    atr: pd.Series,
    J: pd.Series,
    params: Dict[str, float],
    gap_value: float,
    gap_rule: Dict[str, float],
    signal_mode: str,
    window_start: dt.time,
    window_end: dt.time,
) -> List[int]:
    entries: List[int] = []
    eff_j_th = float(params["J_th"]) + float(gap_rule.get("j_th_add", 0.0))
    for i, idx in enumerate(day_positions):
        ts_time = day_df.loc[idx, "time"]
        if ts_time < window_start or ts_time > window_end:
            continue

        abs_j = abs(J.loc[idx])
        if abs_j < eff_j_th - 1e-12:
            continue

        if signal_mode == "j-only":
            entries.append(i)
            continue

        if signal_mode == "j-cross":
            prev_abs = abs(J.loc[day_positions[i - 1]]) if i > 0 else float("inf")
            if prev_abs < eff_j_th - 1e-12:
                entries.append(i)
            continue

        raise ValueError(f"Unsupported signal_mode: {signal_mode}")

    return entries


def new_method_entries(
    day_df: pd.DataFrame,
    day_positions: List[pd.Timestamp],
    atr: pd.Series,
    J: pd.Series,
    params: Dict[str, float],
    gap_value: float,
    gap_rule: Dict[str, float],
    signal_mode: str,
    window_start: dt.time,
    window_end: dt.time,
    alpha: float,
    gamma: float,
    standby_max: int,
) -> List[int]:
    entries: List[int] = []
    eff_j_th = float(params["J_th"]) + float(gap_rule.get("j_th_add", 0.0))

    state: Optional[Dict[str, float]] = None

    for i, idx in enumerate(day_positions):
        ts_time = day_df.loc[idx, "time"]
        abs_j = abs(J.loc[idx])

        if state is not None:
            # reset when outside window
            if ts_time > window_end or ts_time < window_start:
                state = None
                continue

            direction = state["direction"]
            if np.sign(J.loc[idx]) != direction:
                state = None
                continue

            prev_abs = state["prev_abs"]
            speed = abs_j - prev_abs
            max_speed = state["max_speed"]

            if speed > 0:
                max_speed = max(max_speed, speed)
            peak_abs = max(state["peak_abs"], abs_j)

            slowdown = False
            if max_speed <= BP_EPS:
                slowdown = speed <= 0
            else:
                slowdown = speed <= max_speed * alpha + BP_EPS

            reversal = (peak_abs - abs_j) >= gamma - BP_EPS

            state["prev_abs"] = abs_j
            state["max_speed"] = max_speed
            state["peak_abs"] = peak_abs
            state["bars_waited"] += 1

            if reversal and slowdown and abs_j >= eff_j_th - 1e-12:
                entries.append(i)
                state = None
                continue

            if state["bars_waited"] >= standby_max:
                state = None
                continue

        # no standby: check for fresh trigger within window
        if ts_time < window_start or ts_time > window_end:
            continue

        if abs_j < eff_j_th - 1e-12:
            continue

        if signal_mode == "j-cross":
            prev_abs = abs(J.loc[day_positions[i - 1]]) if i > 0 else float("inf")
            if prev_abs >= eff_j_th - 1e-12:
                continue

        direction = np.sign(J.loc[idx])
        if direction == 0:
            continue

        state = {
            "direction": direction,
            "start_idx": i,
            "prev_abs": abs_j,
            "peak_abs": abs_j,
            "max_speed": 0.0,
            "bars_waited": 0,
        }

    return entries


def evaluate_strategy(
    ticker: str,
    signal_mode: str,
    params: Dict[str, float],
    df: pd.DataFrame,
    window_label: str,
    window_start: dt.time,
    window_end: dt.time,
    method: str,
    cost_bp: float,
    alpha: Optional[float] = None,
    gamma: Optional[float] = None,
    standby_max: Optional[int] = None,
) -> TradeResult:
    atr = atr_ema(df["tr"], int(params["ATR_n"])).replace([np.nan, 0.0], np.nan)
    J = (df["close"] - df["vwap"]) / atr

    pnl_records: List[float] = []
    bars_list: List[int] = []

    for date, day_df in df.groupby("date"):
        if day_df["gap_bp"].isna().all():
            continue

        gap_value = float(day_df["gap_bp"].iloc[0]) if not day_df["gap_bp"].isna().all() else 0.0
        gap_rule = build_gap_settings(gap_value, signal_mode)

        day_positions = list(day_df.index)
        if not day_positions:
            continue

        if method == "baseline":
            entry_indices = baseline_entries(
                day_df,
                day_positions,
                atr,
                J,
                params,
                gap_value,
                gap_rule,
                signal_mode,
                window_start,
                window_end,
            )
        else:
            entry_indices = new_method_entries(
                day_df,
                day_positions,
                atr,
                J,
                params,
                gap_value,
                gap_rule,
                signal_mode,
                window_start,
                window_end,
                alpha=alpha or 0.0,
                gamma=gamma or 0.0,
                standby_max=standby_max or 0,
            )

        if not entry_indices:
            continue

        for idx_pos in entry_indices:
            trade = simulate_trade(
                day_df,
                day_positions,
                idx_pos,
                atr,
                J,
                params,
                gap_value,
                gap_rule,
                cost_bp=cost_bp,
            )
            if trade is None:
                continue
            pnl_records.append(float(trade["pnl_bp"]))
            bars_list.append(int(trade["bars"]))

    metrics = metrics_from_pnl(pnl_records)
    avg_bars = (sum(bars_list) / len(bars_list)) if bars_list else 0.0

    params_id = f"{ticker}_{signal_mode}_{window_label}"
    if method != "baseline":
        params_id += f"_a{alpha}_g{gamma}_w{standby_max}"

    return TradeResult(
        ticker=ticker,
        signal_mode=signal_mode,
        window=window_label,
        method=method,
        params_id=params_id,
        alpha=alpha,
        gamma=gamma,
        standby_max=standby_max,
        trades=metrics["trades"],
        wins=metrics["wins"],
        losses=metrics["losses"],
        flats=metrics["flats"],
        winrate=metrics["winrate"],
        pf=metrics["pf"],
        pf_eff=metrics["pf_eff"],
        exp_bp=metrics["exp_bp"],
        total_bp=metrics["total_bp"],
        avg_bars=avg_bars,
        pnl_list=pnl_records,
    )


def bootstrap_diff(
    baseline: Sequence[float],
    challenger: Sequence[float],
    n_boot: int = 1000,
    alpha: float = 0.05,
) -> Tuple[float, float, float]:
    base = np.array(baseline, dtype=float)
    chal = np.array(challenger, dtype=float)

    if len(base) == 0 or len(chal) == 0:
        return 0.0, 0.0, 0.0

    rng = np.random.default_rng(1234)
    diffs = []
    for _ in range(n_boot):
        sample_base = rng.choice(base, size=len(base), replace=True)
        sample_chal = rng.choice(chal, size=len(chal), replace=True)
        diffs.append(float(sample_chal.mean() - sample_base.mean()))
    diffs.sort()
    mean = float(np.mean(diffs))
    lo = diffs[int((alpha / 2) * (len(diffs) - 1))]
    hi = diffs[int((1 - alpha / 2) * (len(diffs) - 1))]
    return mean, lo, hi


def result_to_dict(result: TradeResult) -> Dict[str, object]:
    return {
        "ticker": result.ticker,
        "signal_mode": result.signal_mode,
        "window": result.window,
        "method": result.method,
        "alpha": result.alpha,
        "gamma": result.gamma,
        "standby_max": result.standby_max,
        "trades": result.trades,
        "wins": result.wins,
        "losses": result.losses,
        "flats": result.flats,
        "winrate": result.winrate,
        "pf": result.pf,
        "pf_eff": result.pf_eff,
        "exp_bp": result.exp_bp,
        "total_bp": result.total_bp,
        "avg_bars": result.avg_bars,
        "pnl_json": json.dumps(result.pnl_list),
    }


def run_analysis(
    candidates_path: Path,
    data_root: Path,
    out_dir: Path,
    cost_bp: float,
) -> None:
    candidates = load_candidates(candidates_path)
    tickers = sorted(candidates["Ticker"].unique())

    out_dir.mkdir(parents=True, exist_ok=True)

    baseline_rows: List[Dict[str, object]] = []
    new_rows: List[Dict[str, object]] = []
    summary_rows: List[Dict[str, object]] = []

    prepared_cache: Dict[str, pd.DataFrame] = {}

    for ticker in tickers:
        raw_df = load_intraday_data(ticker, data_root)
        prepared_cache[ticker] = prepare_dataset(raw_df)

    for _, row in candidates.iterrows():
        ticker = str(row["Ticker"])
        signal_mode = row.get("SignalMode", "j-only")
        params = {
            "ATR_n": float(row.get("ATR_n", 5)),
            "TPk": float(row.get("TPk", 1.5)),
            "SLk": float(row.get("SLk", 1.0)),
            "J_th": float(row.get("J_th", 0.8)),
            "dJ_th": float(row.get("dJ_th", 0.0)),
            "vEMA_th": float(row.get("vEMA_th", 0.0)),
            "TMAX": float(row.get("TMAX", 0.0)),
        }

        df = prepared_cache[ticker]

        for window_start, window_end in WINDOWS:
            start_time = parse_time(window_start)
            end_time = parse_time(window_end)
            window_label = f"{window_start}-{window_end}"

            baseline_result = evaluate_strategy(
                ticker=ticker,
                signal_mode=signal_mode,
                params=params,
                df=df,
                window_label=window_label,
                window_start=start_time,
                window_end=end_time,
                method="baseline",
                cost_bp=cost_bp,
            )
            baseline_rows.append(result_to_dict(baseline_result))

            for alpha in ALPHA_GRID:
                for gamma in GAMMA_GRID:
                    for standby_max in STANDBY_MAX_BARS:
                        new_result = evaluate_strategy(
                            ticker=ticker,
                            signal_mode=signal_mode,
                            params=params,
                            df=df,
                            window_label=window_label,
                            window_start=start_time,
                            window_end=end_time,
                            method="new",
                            cost_bp=cost_bp,
                            alpha=alpha,
                            gamma=gamma,
                            standby_max=standby_max,
                        )
                        new_rows.append(result_to_dict(new_result))

            # summary vs best new
            window_new = [
                r
                for r in new_rows
                if r["ticker"] == ticker
                and r["signal_mode"] == signal_mode
                and r["window"] == window_label
            ]
            if not window_new:
                continue
            best_new = max(window_new, key=lambda rec: rec["exp_bp"])
            diff_mean, diff_lo, diff_hi = bootstrap_diff(
                json.loads(baseline_rows[-1]["pnl_json"]),
                json.loads(best_new["pnl_json"]),
                n_boot=1000,
                alpha=0.05,
            )
            summary_rows.append(
                {
                    "ticker": ticker,
                    "signal_mode": signal_mode,
                    "window": window_label,
                    "baseline_trades": baseline_rows[-1]["trades"],
                    "baseline_exp_bp": baseline_rows[-1]["exp_bp"],
                    "baseline_pf_eff": baseline_rows[-1]["pf_eff"],
                    "baseline_avg_bars": baseline_rows[-1]["avg_bars"],
                    "new_trades": best_new["trades"],
                    "new_exp_bp": best_new["exp_bp"],
                    "new_pf_eff": best_new["pf_eff"],
                    "new_avg_bars": best_new["avg_bars"],
                    "alpha": best_new["alpha"],
                    "gamma": best_new["gamma"],
                    "standby_max": best_new["standby_max"],
                    "exp_diff_mean": diff_mean,
                    "exp_diff_lo": diff_lo,
                    "exp_diff_hi": diff_hi,
                }
            )

    baseline_df = pd.DataFrame(baseline_rows)
    new_df = pd.DataFrame(new_rows)
    summary_df = pd.DataFrame(summary_rows)

    baseline_df.to_csv(out_dir / "baseline_results.csv", index=False, encoding="utf-8-sig")
    new_df.to_csv(out_dir / "new_method_grid.csv", index=False, encoding="utf-8-sig")
    summary_df.to_csv(out_dir / "comparison_summary.csv", index=False, encoding="utf-8-sig")


def main() -> None:
    parser = argparse.ArgumentParser(description="Compare baseline vs new entry timing logic.")
    parser.add_argument(
        "--candidates",
        type=Path,
        default=Path("output/excel/candidates_nextday.csv"),
        help="Path to candidates CSV from nightly batch.",
    )
    parser.add_argument(
        "--data-root",
        type=Path,
        default=Path("data/raw/yahoo_1m"),
        help="Root directory containing per-ticker 1m parquet files.",
    )
    parser.add_argument(
        "--outdir",
        type=Path,
        default=Path("output/research/new_entry_vs_baseline"),
        help="Directory to store comparison outputs.",
    )
    parser.add_argument(
        "--cost-bp",
        type=float,
        default=8.0,
        help="Round-trip transaction cost in basis points.",
    )
    args = parser.parse_args()

    timestamped_outdir = args.outdir / dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    run_analysis(
        candidates_path=args.candidates,
        data_root=args.data_root,
        out_dir=timestamped_outdir,
        cost_bp=args.cost_bp,
    )
    print(f"Results written to {timestamped_outdir}")


if __name__ == "__main__":
    main()
