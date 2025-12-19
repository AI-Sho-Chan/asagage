import argparse
import datetime as dt
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import numpy as np
import pandas as pd
from yahooquery import Ticker
import yfinance as yf

STATUS_PATH = Path("logs/nightly_status.txt")
STATUS_DATA: Dict[str, str] = {}
DEFAULT_WORKBOOK_PATH = Path("C:/AI/asagake/ASAGAKE.xlsm")
START_ROW = 6
FORMULA_COLS = (8, 9, 10, 11, 14, 15, 16, 17, 18, 19, 20)


def write_status(**updates: object) -> None:
    """Persist key/value progress indicators so Excel can poll them."""
    now = dt.datetime.now()
    STATUS_DATA.setdefault("started", now.isoformat())
    for key, value in updates.items():
        STATUS_DATA[key] = str(value)
    STATUS_DATA["updated"] = now.isoformat()

    started = STATUS_DATA.get("started")
    if started:
        try:
            start_ts = dt.datetime.fromisoformat(started)
            STATUS_DATA["elapsed_seconds"] = str(int((now - start_ts).total_seconds()))
        except ValueError:
            pass

    state_value = updates.get("state")
    if isinstance(state_value, str) and state_value in {"success", "error"}:
        STATUS_DATA["completed"] = now.isoformat()

    STATUS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with STATUS_PATH.open("w", encoding="utf-8") as fh:
        for key, value in STATUS_DATA.items():
            fh.write(f"{key}={value}\n")


def run(cmd: List[str], cwd: Path, env: Dict[str, str] | None = None) -> None:
    """Run a subprocess and bubble up failures with context."""
    print("[run]", " ".join(cmd))
    proc = subprocess.run(cmd, cwd=cwd, env=env)
    if proc.returncode != 0:
        write_status(state="error", step="subprocess", message="Command failed", returncode=proc.returncode)
        raise SystemExit(proc.returncode)


def load_codes_from_excel(excel_path: Path, sheet: str) -> List[str]:
    try:
        df = pd.read_excel(excel_path, sheet_name=sheet, usecols="A", header=0)
        return df.iloc[:, 0].dropna().astype(str).tolist()
    except Exception:
        return []


def ensure_numeric_mean(df: pd.DataFrame, column: str, decimals: int = 4) -> str:
    if column not in df.columns:
        return ""
    series = pd.to_numeric(df[column], errors="coerce")
    series = series.replace([np.inf, -np.inf], np.nan).dropna()
    if series.empty:
        return ""
    return f"{series.mean():.{decimals}f}"


def aggregate_candidates(
    frames: List[pd.DataFrame],
    out_path: Path,
    *,
    min_forward_ci: float = 0.65,
    min_forward_winrate: float = 0.0,
    min_index_corr: float = 0.2,
    min_vwap_revert: float = 0.0,
    corr_map: Dict[str, Dict[str, float]] | None = None,
    vwap_stats: Dict[str, Dict[str, float]] | None = None,
    run_type: str = "weekday",
) -> Dict[str, str]:
    """Combine plan outputs, enforce quality filters, and keep all qualifying combos per ticker.

    Filters (hard):
      - J_th >= 0.8
      - forward_win_ci_low >= min_forward_ci (default 0.65)
      - forward_pf_eff >= 1.30
      - forward_trades >= 5
      - CorrTOPIX >= min_index_corr (default 0.2) if列がある
      - VWAP_revert_prob >= min_vwap_revert (データがある場合)

    Scoring:
      score = forward_pf_eff * (forward_winrate ** 1.2) * log1p(forward_trades) / (1 + MaxDD/1000)
      さらに Corr / VWAP 回帰があれば軽く加点
    """
    summary: Dict[str, str] = {
        "total_candidates": "0",
        "unique_tickers": "0",
        "avg_forward_winrate": "",
        "avg_forward_pf": "",
        "avg_expected_return": "",
        "avg_forward_trades": "",
    }
    if not frames:
        return summary

    df = pd.concat(frames, ignore_index=True)
    # Normalize/ensure columns
    cols = {c.lower(): c for c in df.columns}
    if "maxdd" not in cols and "max_dd" in cols:
        df = df.rename(columns={cols["max_dd"]: "MaxDD"})
        cols = {c.lower(): c for c in df.columns}
    if "MaxDD" not in df.columns:
        df["MaxDD"] = 0.0

    def c(name: str) -> str:
        return cols.get(name.lower(), name)

    def num(col: str) -> pd.Series:
        return pd.to_numeric(df.get(col, 0), errors="coerce").fillna(0)

    # 補助指標をマージ（あれば）
    if corr_map and "Ticker" in df.columns:
        upper_series = df["Ticker"].astype(str).str.upper()
        df["CorrNKY"] = upper_series.map(lambda t: corr_map.get(t, {}).get("CorrNKY", np.nan))
        df["CorrTOPIX"] = upper_series.map(lambda t: corr_map.get(t, {}).get("CorrTOPIX", np.nan))
    if vwap_stats and "Ticker" in df.columns:
        upper_series = df["Ticker"].astype(str).str.upper()
        df["VWAP_revert_prob"] = upper_series.map(lambda t: vwap_stats.get(t, {}).get("prob", np.nan))
        df["VWAP_revert_bars"] = upper_series.map(lambda t: vwap_stats.get(t, {}).get("avg_bars", np.nan))

    # Hard filters
    if c("J_th") in df.columns:
        df = df[num(c("J_th")) >= 0.8]
    if c("forward_win_ci_low") in df.columns:
        df = df[num(c("forward_win_ci_low")) >= float(min_forward_ci)]
    if c("forward_pf_eff") in df.columns:
        df = df[num(c("forward_pf_eff")) >= 1.30]
    if c("forward_trades") in df.columns:
        df = df[num(c("forward_trades")) >= 5]
    if min_forward_winrate > 0 and c("forward_winrate") in df.columns:
        df = df[num(c("forward_winrate")) >= float(min_forward_winrate)]
    if min_index_corr > 0 and "CorrTOPIX" in df.columns:
        df = df[pd.to_numeric(df["CorrTOPIX"], errors="coerce").fillna(-1) >= float(min_index_corr)]
    if min_vwap_revert > 0 and "VWAP_revert_prob" in df.columns:
        df = df[pd.to_numeric(df["VWAP_revert_prob"], errors="coerce").fillna(0) >= float(min_vwap_revert)]

    # Score
    pf = num(c("forward_pf_eff"))
    win = num(c("forward_winrate"))
    trades = num(c("forward_trades"))
    dd = num("MaxDD")
    score = pf * (win ** 1.2) * (trades.add(1).apply(np.log1p)) / (1.0 + dd / 1000.0)
    # 軽い加点: Corr と VWAP 回帰
    if "CorrTOPIX" in df.columns:
        score = score * (1.0 + pd.to_numeric(df["CorrTOPIX"], errors="coerce").fillna(0).clip(lower=-1, upper=1) * 0.05)
    if "VWAP_revert_prob" in df.columns:
        score = score * (1.0 + pd.to_numeric(df["VWAP_revert_prob"], errors="coerce").fillna(0).clip(0, 1) * 0.1)
    df["_score"] = score

    # Sort by ticker / score but keep all combos that passed the filters
    if "Ticker" in df.columns:
        df = df.sort_values(["Ticker", "_score"], ascending=[True, False])
    else:
        df = df.sort_values(["_score"], ascending=[False])

    weekly_base = Path("output/excel/weekly_candidates_latest.csv")
    if weekly_base.exists():
        try:
            weekly_df = pd.read_csv(weekly_base)
            if "Ticker" in weekly_df.columns:
                before = len(df)
                df = weekly_df.merge(df, on="Ticker", how="left", suffixes=("_weekly", ""))
                for col in list(df.columns):
                    if col.endswith("_weekly"):
                        base_col = col[:-7]
                        if base_col in df.columns:
                            df[base_col] = df[base_col].where(df[base_col].notna(), df[col])
                            df.drop(columns=[col], inplace=True)
                df = df[df["Ticker"].notna()]
                summary["message"] = f"weekly base {len(df)}/{before}"
        except Exception as exc:  # pragma: no cover
            summary["message"] = f"weekly merge failed: {exc}"

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if "_score" in df.columns:
        df = df.drop(columns=["_score"])
    # Defensive: we've seen rare runs where `run_type` was missing at runtime,
    # causing the final aggregation step to crash and skip exports/sync.
    batch_kind = locals().get("run_type") or "weekday"
    df["BatchKind"] = str(batch_kind)
    df.to_csv(out_path, index=False, encoding="utf-8-sig")

    latest = Path("output/excel") / "weekly_candidates_latest.csv"
    try:
        df.to_csv(latest, index=False, encoding="utf-8-sig")
        summary["weekly_synced"] = "1"
    except Exception:
        summary["weekly_synced"] = "0"

    cols = {c.lower(): c for c in df.columns}

    summary["total_candidates"] = str(int(len(df)))
    summary["unique_tickers"] = str(int(df["Ticker"].nunique())) if "Ticker" in df.columns else summary["total_candidates"]
    summary["avg_forward_winrate"] = ensure_numeric_mean(df, c("forward_winrate"))
    summary["avg_forward_pf"] = ensure_numeric_mean(df, c("forward_pf_eff"))
    summary["avg_expected_return"] = ensure_numeric_mean(df, c("forward_exp_boot_mean"))
    summary["avg_forward_trades"] = ensure_numeric_mean(df, c("forward_trades"), decimals=2)
    return summary


def chunked(seq: List[str], size: int) -> Iterable[List[str]]:
    for idx in range(0, len(seq), size):
        yield seq[idx : idx + size]


def download_daily_returns(symbols: List[str], period: str = "3mo") -> Dict[str, pd.Series]:
    returns: Dict[str, pd.Series] = {}
    if not symbols:
        return returns
    for batch in chunked(symbols, 30):
        data = yf.download(
            batch,
            period=period,
            interval="1d",
            auto_adjust=False,
            progress=False,
            group_by="ticker",
        )
        if data is None or data.empty:
            continue
        if isinstance(data.columns, pd.MultiIndex):
            for sym in batch:
                if (sym, "Close") not in data.columns:
                    continue
                series = (
                    data[(sym, "Close")]
                    .pct_change()
                    .replace([np.inf, -np.inf], np.nan)
                    .dropna()
                )
                if not series.empty:
                    series.index = pd.to_datetime(series.index).tz_localize(None)
                    returns[sym] = series
        else:
            sym = batch[0]
            series = (
                data["Close"].pct_change().replace([np.inf, -np.inf], np.nan).dropna()
            )
            if not series.empty:
                series.index = pd.to_datetime(series.index).tz_localize(None)
                returns[sym] = series
    return returns


def compute_corr_map(tickers: List[str], period: str = "6mo") -> Dict[str, Dict[str, float]]:
    uniq = [t.strip().upper() for t in tickers if isinstance(t, str) and t.strip()]
    uniq = sorted(dict.fromkeys(uniq))
    if not uniq:
        return {}
    nky_symbol = "^N225"
    topix_candidates = ["^TOPX", "^TPX", "1306.T"]
    symbols = uniq + [nky_symbol] + topix_candidates
    returns = download_daily_returns(symbols, period=period)
    corr_map: Dict[str, Dict[str, float]] = {}

    topix_symbol = None
    for candidate in topix_candidates:
        if candidate in returns and not returns[candidate].empty:
            topix_symbol = candidate
            break

    def corr_pair(series_a: pd.Series, series_b: pd.Series) -> float | None:
        if series_a is None or series_b is None:
            return None
        frame = pd.concat([series_a, series_b], axis=1, join="inner").dropna()
        if len(frame) < 10:
            return None
        value = frame.iloc[:, 0].corr(frame.iloc[:, 1])
        if pd.isna(value):
            return None
        return float(value)

    for sym in uniq:
        series = returns.get(sym)
        if series is None:
            continue
        corr_map[sym] = {
            "CorrNKY": corr_pair(series, returns.get(nky_symbol)) or 0.0,
            "CorrTOPIX": corr_pair(series, returns.get(topix_symbol)) or 0.0,
        }
    return corr_map


def compute_vwap_revert_stats(
    tickers: List[str],
    lookback_days: int = 60,
    j_threshold: float = 1.0,
    max_bars: int = 30,
    revert_epsilon: float = 0.25,
) -> Dict[str, Dict[str, float]]:
    """Compute simple VWAP回帰指標: 乖離>=J_th のイベントが max_bars 本以内に VWAP±ε*ATR へ戻る確率。"""
    root = Path("data/raw/yahoo_1m")
    uniq = [t.strip().upper() for t in tickers if isinstance(t, str) and t.strip()]
    uniq = sorted(dict.fromkeys(uniq))
    if not uniq:
        return {}

    stats: Dict[str, Dict[str, float]] = {}
    for code in uniq:
        directory = root / code
        if not directory.exists():
            continue
        files = sorted(directory.glob("*.parquet"))
        if not files:
            continue
        files = files[-lookback_days:]

        events = 0
        successes = 0
        bars_to_revert: List[int] = []

        for fp in files:
            try:
                df = pd.read_parquet(fp)
            except Exception:
                continue
            if df.empty:
                continue
            # normalize
            if isinstance(df.index, pd.DatetimeIndex):
                df = df.sort_index()
            elif "ts" in df.columns:
                df["ts"] = pd.to_datetime(df["ts"])
                df = df.sort_values("ts")
                df = df.set_index("ts")
            # 必要カラムがなければスキップ
            for col in ("close", "high", "low", "volume"):
                if col not in df.columns:
                    break
            else:
                close = pd.to_numeric(df["close"], errors="coerce")
                high = pd.to_numeric(df["high"], errors="coerce")
                low = pd.to_numeric(df["low"], errors="coerce")
                vol = pd.to_numeric(df["volume"], errors="coerce")
                if close.isna().all() or vol.isna().all():
                    continue

                cum_vol = vol.cumsum()
                vwap = (close * vol).cumsum() / cum_vol.replace(0, np.nan)

                prev_close = close.shift(1)
                tr = pd.concat(
                    [
                        high - low,
                        (high - prev_close).abs(),
                        (low - prev_close).abs(),
                    ],
                    axis=1,
                ).max(axis=1)
                atr = tr.rolling(14, min_periods=5).mean()
                dev = (close - vwap) / atr.replace(0, np.nan)

                dev_abs = dev.abs().to_numpy()
                n = len(dev_abs)
                for idx in range(n):
                    val = dev_abs[idx]
                    if not np.isfinite(val) or val < j_threshold:
                        continue
                    events += 1
                    window = dev_abs[idx : min(idx + max_bars, n)]
                    hits = np.where(window < revert_epsilon)[0]
                    if hits.size > 0:
                        successes += 1
                        bars_to_revert.append(int(hits[0]))

        if events > 0:
            prob = successes / events
            avg_bars = float(np.mean(bars_to_revert)) if bars_to_revert else np.nan
            stats[code] = {"prob": prob, "avg_bars": avg_bars}

    return stats


def enrich_dashboard_columns(csv_path: Path, coeff_path: Path) -> None:
    try:
        df = pd.read_csv(csv_path)
    except Exception:
        return
    if df.empty:
        return

    coeff_df = None
    if coeff_path.exists():
        try:
            coeff_df = pd.read_csv(coeff_path)
        except Exception:
            coeff_df = None
    if coeff_df is not None and not coeff_df.empty:
        coeff_df = coeff_df.rename(
            columns={
                "bias_slope": "BiasSlope_row",
                "gap_slope": "GapSlope_row",
                "corr_slope": "CorrSlope_row",
            }
        )
        keep_cols = ["Ticker", "BiasSlope_row", "GapSlope_row", "CorrSlope_row"]
        coeff_df = coeff_df[[c for c in keep_cols if c in coeff_df.columns]]
        if "Ticker" in coeff_df.columns:
            df = df.merge(coeff_df, on="Ticker", how="left")

    for col, default in (
        ("BiasSlope_row", 0.1),
        ("GapSlope_row", 0.2),
        ("CorrSlope_row", 0.05),
    ):
        if col not in df.columns:
            df[col] = default
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    if "TP_per_J_row" not in df.columns:
        df["TP_per_J_row"] = df.get("TPk")
    df["TP_per_J_row"] = pd.to_numeric(df["TP_per_J_row"], errors="coerce")
    if "TPk" in df.columns:
        df["TP_per_J_row"] = df["TP_per_J_row"].fillna(pd.to_numeric(df["TPk"], errors="coerce"))
    df["TP_per_J_row"] = df["TP_per_J_row"].fillna(0.15)

    if "SL_per_J_row" not in df.columns:
        df["SL_per_J_row"] = df.get("SLk")
    df["SL_per_J_row"] = pd.to_numeric(df["SL_per_J_row"], errors="coerce")
    if "SLk" in df.columns:
        df["SL_per_J_row"] = df["SL_per_J_row"].fillna(pd.to_numeric(df["SLk"], errors="coerce"))
    df["SL_per_J_row"] = df["SL_per_J_row"].fillna(0.1)

    if "Trail_per_J_row" not in df.columns:
        df["Trail_per_J_row"] = df["SL_per_J_row"]
    df["Trail_per_J_row"] = pd.to_numeric(df["Trail_per_J_row"], errors="coerce").fillna(df["SL_per_J_row"])
    df["Trail_per_J_row"] = df["Trail_per_J_row"].fillna(0.1)

    for eff_col in ("TP_per_J_eff", "SL_per_J_eff", "Trail_per_J_eff", "VolatilityTag"):
        if eff_col not in df.columns:
            df[eff_col] = ""

    ticker_list = df.get("Ticker")
    if ticker_list is not None:
        corr_map = compute_corr_map(ticker_list.dropna().astype(str).tolist())
        upper_series = ticker_list.astype(str).str.upper()
        df["CorrNKY"] = upper_series.map(lambda t: corr_map.get(t, {}).get("CorrNKY", np.nan))
        df["CorrTOPIX"] = upper_series.map(lambda t: corr_map.get(t, {}).get("CorrTOPIX", np.nan))
    else:
        df["CorrNKY"] = np.nan
        df["CorrTOPIX"] = np.nan

    for col in ("CorrNKY", "CorrTOPIX"):
        df[col] = pd.to_numeric(df.get(col), errors="coerce").fillna(0.0)

    df.to_csv(csv_path, index=False, encoding="utf-8-sig")


def ensure_dashboard_formulas(repo_root: Path, workbook_path: Path) -> None:
    """Restore formulas via COM and verify key RSS columns.

    Note: We prefer restore_dashboard_formulas.py as it reapplies the canonical
    formulas (I窶天) including SignalStatus/Kind, and re-protects the sheet.
    """
    try:
        run([sys.executable, "scripts/restore_dashboard_formulas.py"], cwd=repo_root)
    except SystemExit:
        write_status(
            state="running",
            step="repair_formulas",
            message="restore_dashboard_formulas.py failed",
        )
        return

    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        write_status(
            state="running",
            step="repair_formulas",
            message=f"Skipped formula verification (COM unavailable: {exc})",
        )
        return

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(workbook_path))
        try:
            ws = wb.Worksheets("NewDashboard")
        except Exception:
            write_status(
                state="running",
                step="repair_formulas",
                message="NewDashboard sheet not found",
            )
            return

        missing: List[str] = []
        col_names = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        for col in FORMULA_COLS:
            try:
                cell = ws.Cells(START_ROW, col)
                has_formula = bool(cell.HasFormula)
            except Exception:
                has_formula = False
            if not has_formula:
                label = col_names[col - 1] if col <= len(col_names) else f"col{col}"
                missing.append(label)

        if missing:
            write_status(
                state="running",
                step="repair_formulas",
                message="Formulas missing in " + ",".join(missing),
            )
        else:
            write_status(
                state="running",
                step="repair_formulas",
                message="Dashboard formulas verified (H6-T6)",
            )
    finally:
        try:
            wb.Close(SaveChanges=True)
        except Exception:
            pass
        excel.Quit()


def apply_trend_preferences(csv_path: Path, pref_path: Path, bp_threshold: float) -> None:
    if not csv_path.exists():
        return
    try:
        df = pd.read_csv(csv_path)
    except Exception:
        return
    if df.empty:
        return

    if "Ticker" in df.columns:
        ticker_key = df["Ticker"].astype(str).str.strip().str.upper()
    elif "code" in df.columns:
        ticker_key = df["code"].astype(str).str.strip().str.upper()
    else:
        ticker_key = pd.Series([""] * len(df))

    if "trend_driver" in df.columns:
        df["trend_driver"] = df["trend_driver"].fillna("NKY")
    else:
        df["trend_driver"] = "NKY"
    if "trend_window" in df.columns:
        df["trend_window"] = df["trend_window"].fillna("window")
    else:
        df["trend_window"] = "window"
    df["trend_bp_th"] = float(bp_threshold)
    if "trend_allowed_policy" in df.columns:
        df["trend_allowed_policy"] = df["trend_allowed_policy"].fillna("ALIGNED_ONLY")
    else:
        df["trend_allowed_policy"] = "ALIGNED_ONLY"

    if pref_path.exists():
        try:
            pref = pd.read_csv(pref_path)
        except Exception:
            pref = pd.DataFrame()
        if not pref.empty and {"code", "driver", "trend_type"}.issubset(pref.columns):
            pref = pref.copy()
            pref["key"] = pref["code"].astype(str).str.strip().str.upper()
            pref = pref.dropna(subset=["key"])
            pref = pref.set_index("key")
            if "driver" in pref.columns:
                df["trend_driver"] = ticker_key.map(pref["driver"]).fillna(df["trend_driver"])
            if "trend_type" in pref.columns:
                df["trend_window"] = ticker_key.map(pref["trend_type"]).fillna(df["trend_window"])

    try:
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    except Exception:
        pass


def format_plan_counts(plan_counts: Dict[str, int]) -> str:
    return ",".join(f"{key}:{value}" for key, value in plan_counts.items())


def build_yahoo_universe(
    base_codes: List[str],
    metric: str,
    universe_size: int,
    output_dir: Path,
    *,
    target_date: Optional[dt.date] = None,
) -> Optional[Path]:
    if not base_codes:
        return None

    anchor = target_date or dt.date.today()
    ticker = Ticker(base_codes, asynchronous=True)
    end_date = anchor + dt.timedelta(days=1)
    start_date = anchor - dt.timedelta(days=2)
    hist = ticker.history(start=str(start_date), end=str(end_date), interval="1m")
    if not isinstance(hist, pd.DataFrame) or hist.empty:
        return None

    df = hist.reset_index()
    if "symbol" in df.columns:
        df = df.rename(columns={"symbol": "code"})
    if "date" in df.columns and "ts" not in df.columns:
        df = df.rename(columns={"date": "ts"})
    df["ts"] = pd.to_datetime(df["ts"])
    df["date"] = df["ts"].dt.date
    df["amt"] = df["close"] * df["volume"]
    latest_day = df["date"].max()
    last_slice = df[df["date"] == latest_day]
    if metric == "amt":
        metric_df = last_slice.groupby("code")["amt"].sum().reset_index(name="score")
    else:
        metric_df = last_slice.groupby("code")["volume"].sum().reset_index(name="score")

    topn = metric_df.sort_values("score", ascending=False).head(int(universe_size))
    output_dir.mkdir(parents=True, exist_ok=True)
    out_file = output_dir / f"universe_{metric}_top_{universe_size}_{latest_day}.csv"
    topn[["code"]].to_csv(out_file, index=False)
    return out_file


def _main_impl() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", default="ASAGAKE.xlsm")
    ap.add_argument("--base-out", default="output/bt30")
    ap.add_argument("--lookback", type=int, default=60)
    ap.add_argument("--chunk-days", type=int, default=5)
    ap.add_argument("--train-days", type=int, default=12)
    ap.add_argument("--forward-days", type=int, default=4)
    ap.add_argument("--min-train-trades", type=int, default=15)
    ap.add_argument("--min-forward-trades", type=int, default=3)
    ap.add_argument("--forward-pf-min", type=float, default=1.3)
    ap.add_argument(
        "--min-forward-ci",
        type=float,
        default=0.65,
        help="Minimum forward winrate CI lower bound to accept in aggregation (default 0.65)",
    )
    ap.add_argument("--gap-guard-abs-bp", type=float, default=80.0)
    ap.add_argument("--gap-guard-dir-bp", type=float, default=40.0)
    ap.add_argument("--min-forward-winrate", type=float, default=0.0, help="Optional min forward winrate filter in aggregation (e.g., 0.60)")
    ap.add_argument("--headless", action="store_true", help="Skip Excel/COM work (do not open ASAGAKE.xlsm)")
    ap.add_argument("--slipbp", type=float, default=4.0)
    ap.add_argument("--feebp", type=float, default=4.0)
    ap.add_argument("--liquidity-quantile", type=float, default=0.5)
    ap.add_argument("--jobs", type=int, default=0, help="Worker processes for bt_opt30_forward (0=auto)")
    ap.add_argument("--excel-summary", action="store_true")
    ap.add_argument("--universe-mode", choices=["excel", "yahoo-top"], default="yahoo-top")
    ap.add_argument("--universe-size", type=int, default=300)
    ap.add_argument("--universe-source", default="data/universe/topvol_*.csv")
    ap.add_argument("--excel-ticker-sheet", default="Ticker")
    ap.add_argument("--universe-metric", choices=["amt", "vol"], default="amt")
    ap.add_argument("--run-type", choices=["auto", "weekday", "weekend"], default="auto")
    ap.add_argument("--target-date", help="Override trading date tag (YYYYMMDD). Defaults to today.")
    ap.add_argument(
        "--minute-cache-history",
        type=int,
        default=5,
        help="Trading days of Yahoo minute data to stage locally before runs (default 5)",
    )
    ap.add_argument(
        "--minute-cache-limit",
        type=int,
        default=1000,
        help="Maximum number of tickers to refresh when updating minute cache (default 1000)",
    )
    ap.add_argument(
        "--disable-minute-cache",
        action="store_true",
        help="Skip automatic minute cache refresh for Yahoo minute data",
    )
    ap.add_argument(
        "--minute-backfill-days",
        type=int,
        default=1,
        help="How many oldest trading days to extend per run (default 1, set 0 to disable)",
    )
    ap.add_argument("--enable-asha", action="store_true", help="Pass --enable-asha to coarse runs")
    ap.add_argument("--enable-bayes", action="store_true", help="Pass --enable-bayes to refine runs")
    ap.add_argument("--bayes-trials", type=int, default=40, help="Bayesian trials per refine run when enabled")
    ap.add_argument("--bayes-timeout", type=int, default=0, help="Bayesian timeout seconds (0=disabled)")
    ap.add_argument("--coeff-history-days", type=int, default=60, help="Trading days used for dashboard coefficient regression")
    ap.add_argument("--disable-dashboard-coeffs", action="store_true", help="Skip coefficient refresh/merge")
    ap.add_argument("--trend-pref", default="analysis/trend_ticker_preference.csv", help="Ticker→trend driver mapping CSV")
    ap.add_argument("--trend-bp-th", type=float, default=15.0, help="bp threshold for trend alignment policies")
    ap.add_argument("--mask-ineffective", action="store_true", help="Enable ineffective-band masking during coarse runs")
    ap.add_argument("--mask-window", type=int, default=20, help="Mask history window (runs)")
    ap.add_argument("--mask-threshold", type=float, default=1.05, help="Forward pf_eff threshold for mask retention")
    ap.add_argument("--cache-refresh-weekend", action="store_true", help="Force --cache-refresh on weekend runs")
    ap.add_argument("--plan-profile", choices=["auto", "weekend", "weekday"], default="auto")
    ap.add_argument("--analysis-ledger", action="store_true", help="(deprecated) no-op placeholder")
    ap.add_argument(
        "--refine-quick-grid",
        action="store_true",
        help="Add --quick-grid and --optimize-io to refine runs for faster turnaround.",
    )
    ap.add_argument(
        "--enable-market-features",
        action="store_true",
        help="Attach --enable-market-features to coarse/refine runs for hypothesis capture",
    )
    ap.add_argument(
        "--enable-rd-windows",
        action="store_true",
        help="Append optional mid/pm R&D windows",
    )
    ap.add_argument(
        "--rd-only",
        action="store_true",
        help="Run only the R&D windows (skip standard plans)",
    )
    ap.add_argument(
        "--plan-focus",
        help="Comma-separated plan tags to run (e.g., AM15_j-only)",
        default="",
    )
    ap.add_argument("--min-index-corr", type=float, default=0.2, help="Minimum CorrTOPIX to keep a candidate (default 0.2)")
    ap.add_argument(
        "--min-vwap-revert",
        type=float,
        default=0.6,
        help="Minimum VWAP revert probability to keep a candidate (default 0.6; ignored if data missing)",
    )
    ap.add_argument(
        "--vwap-lookback-days",
        type=int,
        default=60,
        help="Trading days for VWAP revert statistics (default 60)",
    )
    ap.add_argument(
        "--vwap-j-threshold",
        type=float,
        default=1.0,
        help="|J| threshold for VWAP revert detection (default 1.0 ATR)",
    )
    ap.add_argument(
        "--vwap-max-bars",
        type=int,
        default=30,
        help="Bars to wait for VWAP revert after signal (default 30)",
    )
    ap.add_argument(
        "--reopt-degraded-only",
        action="store_true",
        help="Weekday mode: only re-optimize tickers whose PF/CI fell below thresholds",
    )
    ap.add_argument(
        "--reopt-pf-th",
        type=float,
        default=1.2,
        help="PF threshold for degraded detection in weekday mode (default 1.2)",
    )
    ap.add_argument(
        "--reopt-ci-th",
        type=float,
        default=0.6,
        help="Win CI low threshold for degraded detection in weekday mode (default 0.6)",
    )
    args = ap.parse_args()

    repo_root = Path(__file__).resolve().parent.parent

    excel_path = Path(args.excel)
    if not excel_path.is_absolute():
        excel_path = (repo_root / excel_path).resolve()
    else:
        excel_path = excel_path.resolve()

    if args.target_date:
        try:
            target_date = dt.datetime.strptime(str(args.target_date), "%Y%m%d").date()
        except ValueError as exc:
            raise SystemExit(f"invalid --target-date: {args.target_date}") from exc
    else:
        target_date = dt.datetime.now().date()

    run_type = args.run_type
    if run_type == "auto":
        run_type = "weekend" if target_date.weekday() >= 5 else "weekday"
    is_weekend = run_type == "weekend"

    base = Path(args.base_out)
    date_tag = target_date.strftime("%Y%m%d")
    night_root = base / f"NIGHTLY_{date_tag}"
    night_root.mkdir(parents=True, exist_ok=True)

    # Plan entries: (label, session_start, session_end, signal_mode)
    plan_profile = args.plan_profile
    if plan_profile == "auto":
        plan_profile = "weekend" if run_type == "weekend" else "weekday"

    def _base_windows(profile: str) -> List[Tuple[str, str, str]]:
        if profile == "weekend":
            return [
                ("AM15", "09:00", "09:15"),
                ("AM0930", "09:00", "09:30"),
                ("AM0945", "09:00", "09:45"),
                ("AM1015", "09:00", "10:15"),
                ("AM1030", "09:00", "10:30"),
            ]
        return [
            ("AM15", "09:00", "09:15"),
            ("AM0930", "09:00", "09:30"),
            ("AM0945", "09:00", "09:45"),
            ("AM1000", "09:00", "10:00"),
            ("AM1015", "09:00", "10:15"),
            ("AM1030", "09:00", "10:30"),
            ("PM1", "12:30", "13:30"),
        ]

    plans: List[Tuple[str, str, str, str]] = []
    for base in _base_windows(plan_profile):
        label, start, end = base
        plans.append((label, start, end, "j-only"))
        plans.append((label, start, end, "j-cross"))

    rd_windows: List[Tuple[str, str, str, str]] = [
        ("MID1030", "10:30", "11:00", "j-cross"),
        ("PM1230", "12:30", "13:00", "j-cross"),
    ]
    if getattr(args, "rd_only", False):
        plans = rd_windows
    elif getattr(args, "enable_rd_windows", False):
        plans.extend(rd_windows)
    plan_focus = {p.strip() for p in args.plan_focus.split(",") if p.strip()}
    if plan_focus:
        plans = [tpl for tpl in plans if f"{tpl[0]}_{tpl[3]}" in plan_focus]
        if not plans:
            raise SystemExit(f"plan_focus yielded no plans: {plan_focus}")
    total_plans = len(plans)
    plan_counts: Dict[str, int] = {f"{label}_{sig}": 0 for label, _, __, sig in plans}
    plan_order: List[str] = []

    bt_env = os.environ.copy()
    bt_env["BT30_PARAM_PROFILE"] = plan_profile

    git_commit = ""
    try:
        result = subprocess.run(
            ["git", "rev-parse", "HEAD"],
            cwd=repo_root,
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            git_commit = result.stdout.strip()
    except Exception:
        git_commit = ""

    write_status(
        state="running",
        step="initializing",
        message="Starting nightly batch",
        base_out=str(base),
        date_tag=date_tag,
        target_date=date_tag,
        total_plans=total_plans,
        run_type=run_type,
        git_commit=git_commit or "unknown",
    )

    codes_file_for_runs: Optional[Path] = None
    base_codes: List[str] = []
    universe_source_note = ""
    universe_diag: List[str] = []
    if args.universe_mode == "yahoo-top":
        write_status(
            state="running",
            step="universe",
            message="Building Yahoo-top universe",
            total_plans=total_plans,
            completed_plans=0,
        )

        src_path = Path(args.universe_source)
        universe_diag.append(f"args.universe_source={args.universe_source}")
        universe_diag.append(f"initial_src={src_path}")
        if "*" in src_path.name or "?" in src_path.name:
            parent = src_path.parent if src_path.parent != Path("") else Path(".")
            matches = sorted(parent.glob(src_path.name))
            if matches:
                # Prefer non-TEST files so that ad-hoc topvol_TEST_* CSVs
                # do not accidentally become the canonical universe source.
                non_test = [m for m in matches if "TEST" not in m.name.upper()]
                if non_test:
                    matches = sorted(non_test)
                src_path = matches[-1]
        elif src_path.is_dir():
            matches = sorted(src_path.glob("topvol_*.csv"))
            if matches:
                non_test = [m for m in matches if "TEST" not in m.name.upper()]
                if non_test:
                    matches = sorted(non_test)
                src_path = matches[-1]
        universe_diag.append(f"resolved_src={src_path} exists={src_path.exists()}")
        if src_path.exists():
            universe_source_note = str(src_path)
            try:
                import pandas as _pd  # local import to avoid any shadowing issues
                dfu = _pd.read_csv(src_path)
                if "code" in dfu.columns:
                    base_codes = dfu["code"].dropna().astype(str).tolist()
                universe_diag.append(
                    f"src_path_ok rows={len(dfu)} cols={list(dfu.columns)} base={len(base_codes)}"
                )
                write_status(
                    state="running",
                    step="universe",
                    message="Loaded universe from universe-source CSV",
                    universe_base=len(base_codes),
                    universe_source=universe_source_note,
                )
            except Exception as exc:
                base_codes = []
                universe_diag.append(f"src_path_error={exc!r}")
                write_status(
                    state="running",
                    step="universe",
                    message=f"Failed to read universe-source CSV: {exc}",
                    universe_base=0,
                    universe_source=universe_source_note or args.universe_source,
                )

        # Fallback: in case base_codes is still empty (e.g. unexpected path handling),
        # try reading args.universe_source as a CSV relative to repo_root.
        if not base_codes:
            alt_path = Path(args.universe_source)
            if not alt_path.is_absolute():
                alt_path = (repo_root / alt_path).resolve()
            if alt_path.exists():
                try:
                    dfu_alt = pd.read_csv(alt_path)
                    if "code" in dfu_alt.columns:
                        base_codes = dfu_alt["code"].dropna().astype(str).tolist()
                        universe_source_note = str(alt_path)
                        write_status(
                            state="running",
                            step="universe",
                            message="Recovered universe from args.universe-source CSV",
                            universe_base=len(base_codes),
                            universe_source=universe_source_note,
                        )
                except Exception as exc:
                    write_status(
                        state="running",
                        step="universe",
                        message=f"Fallback universe-source CSV read failed: {exc}",
                        universe_base=0,
                        universe_source=universe_source_note or str(alt_path),
                    )

        # Fallback: in case base_codes is still empty (e.g. unexpected path handling),
        # try reading args.universe_source as a CSV relative to repo_root.
        if not base_codes:
            alt_path = Path(args.universe_source)
            if not alt_path.is_absolute():
                alt_path = (repo_root / alt_path).resolve()
            universe_diag.append(f"alt_path={alt_path} exists={alt_path.exists()}")
            if alt_path.exists():
                try:
                    import pandas as _pd  # local import to avoid any shadowing issues
                    dfu_alt = _pd.read_csv(alt_path)
                    if "code" in dfu_alt.columns:
                        base_codes = dfu_alt["code"].dropna().astype(str).tolist()
                        universe_source_note = str(alt_path)
                        universe_diag.append(
                            f"alt_path_ok rows={len(dfu_alt)} cols={list(dfu_alt.columns)} base={len(base_codes)}"
                        )
                        write_status(
                            state="running",
                            step="universe",
                            message="Recovered universe from args.universe-source CSV",
                            universe_base=len(base_codes),
                            universe_source=universe_source_note,
                        )
                except Exception as exc:
                    universe_diag.append(f"alt_path_error={exc!r}")
                    write_status(
                        state="running",
                        step="universe",
                        message=f"Fallback universe-source CSV read failed: {exc}",
                        universe_base=0,
                        universe_source=universe_source_note or str(alt_path),
                    )

        if not base_codes:
            base_codes = load_codes_from_excel(Path(args.excel), args.excel_ticker_sheet)
            if base_codes:
                universe_diag.append(f"excel_sheet_base={len(base_codes)}")
                write_status(
                    state="running",
                    step="universe",
                    message="Using Excel ticker sheet as Yahoo base",
                    universe_base=len(base_codes),
                )
        if not base_codes:
            fallback_csv = Path("output/excel/candidates_nextday.csv")
            if fallback_csv.exists():
                try:
                    df_fallback = pd.read_csv(fallback_csv)
                    if "Ticker" in df_fallback.columns:
                        base_codes = df_fallback["Ticker"].dropna().astype(str).tolist()
                        universe_diag.append(f"fallback_candidates_base={len(base_codes)}")
                        write_status(
                            state="running",
                            step="universe",
                            message="Using candidates_nextday.csv as Yahoo base",
                            universe_base=len(base_codes),
                        )
                except Exception:
                    base_codes = []

        if not base_codes:
            write_status(
                state="error",
                step="universe",
                message="No universe codes available for yahoo-top; aborting",
                universe_base=0,
                universe_source=universe_source_note or "none",
                universe_diag="|".join(universe_diag),
            )
            raise SystemExit("No universe codes available for yahoo-top; aborting")
        try:
            codes_file_for_runs = build_yahoo_universe(
                base_codes=base_codes,
                metric=args.universe_metric,
                universe_size=args.universe_size,
                output_dir=night_root,
                target_date=target_date,
            )
        except Exception as exc:
            codes_file_for_runs = None
            write_status(
                state="error",
                step="universe",
                message=f"Yahoo universe build failed: {exc}",
                universe_base=len(base_codes),
                universe_source=universe_source_note or args.universe_source,
            )
            raise SystemExit(f"Yahoo universe build failed: {exc}")

        if not codes_file_for_runs or not codes_file_for_runs.exists():
            write_status(
                state="error",
                step="universe",
                message="Yahoo universe build produced no file; aborting",
                universe_base=len(base_codes),
                universe_source=universe_source_note or args.universe_source,
            )
            raise SystemExit("Yahoo universe build produced no file; aborting")

        # Optional VWAP回帰フィルタ: 週末かつ Yahoo-universe の場合のみ適用
        if is_weekend and codes_file_for_runs and base_codes:
            try:
                df_u = pd.read_csv(codes_file_for_runs)
                if "code" not in df_u.columns:
                    write_status(
                        state="error",
                        step="universe",
                        message="Yahoo universe CSV missing 'code' column; aborting",
                        universe_file=str(codes_file_for_runs),
                        universe_base=len(base_codes),
                        universe_source=universe_source_note
                        or str(codes_file_for_runs),
                    )
                    raise SystemExit(
                        "Yahoo universe CSV missing 'code' column; aborting"
                    )
                if not len(df_u):
                    write_status(
                        state="error",
                        step="universe",
                        message="Yahoo universe CSV has no rows; aborting",
                        universe_file=str(codes_file_for_runs),
                        universe_base=len(base_codes),
                        universe_source=universe_source_note
                        or str(codes_file_for_runs),
                    )
                    raise SystemExit(
                        "Yahoo universe CSV has no rows; aborting"
                    )

                codes_u = (
                    df_u["code"]
                    .dropna()
                    .astype(str)
                    .str.strip()
                    .str.upper()
                    .tolist()
                )
                n_u = len(codes_u)
                if not n_u:
                    write_status(
                        state="error",
                        step="universe",
                        message="Yahoo universe empty after code normalization; aborting",
                        universe_file=str(codes_file_for_runs),
                        universe_base=len(base_codes),
                        universe_source=universe_source_note
                        or str(codes_file_for_runs),
                    )
                    raise SystemExit(
                        "Yahoo universe empty after code normalization; aborting"
                    )

                # 1) VWAP回帰統計を Top ユニバースに対して計算
                vwap_stats = compute_vwap_revert_stats(
                    codes_u,
                    lookback_days=int(getattr(args, "vwap_lookback_days", 60)),
                    j_threshold=float(getattr(args, "vwap_j_threshold", 1.0)),
                    max_bars=int(getattr(args, "vwap_max_bars", 30)),
                )
                # 2) ランクから簡易的な「流動性quantile」を算出 (上位ほど1.0に近い)
                quantiles = {}
                if n_u == 1:
                    quantiles[codes_u[0]] = 1.0
                else:
                    for idx, code in enumerate(codes_u):
                        quantiles[code] = 1.0 - (idx / float(n_u - 1))

                strong_th = 0.70
                weak_th = 0.50
                filtered: List[str] = []
                for code in codes_u:
                    info = vwap_stats.get(code)
                    prob = None
                    if isinstance(info, dict):
                        prob = info.get("prob")
                    q = quantiles.get(code, 0.0)
                    cond = False
                    if prob is None:
                        # データ不足銘柄: 従来どおりの流動性閾値 (0.3) で判断
                        cond = q >= 0.3
                    elif prob >= strong_th:
                        # VWAP回帰が強い銘柄は流動性をやや緩める
                        cond = q >= 0.2
                    elif prob >= weak_th:
                        # 中間層は従来どおり
                        cond = q >= 0.3
                    else:
                        # prob < 0.5: 基本弾くが、quantile>=0.5 の超高流動銘柄だけ残す
                        cond = q >= 0.5
                    if cond:
                        filtered.append(code)

                min_keep = max(10, n_u // 10)
                if not filtered:
                    write_status(
                        state="error",
                        step="universe",
                        message="VWAP filter removed all tickers; aborting",
                        universe_file=str(codes_file_for_runs),
                        universe_base=len(base_codes),
                        universe_source=universe_source_note
                        or str(codes_file_for_runs),
                    )
                    raise SystemExit(
                        "VWAP filter removed all tickers; aborting"
                    )
                if len(filtered) < min_keep:
                    write_status(
                        state="error",
                        step="universe",
                        message=(
                            f"VWAP filter survivors too few ({len(filtered)}/{n_u}); aborting"
                        ),
                        universe_file=str(codes_file_for_runs),
                        universe_base=len(base_codes),
                        universe_source=universe_source_note
                        or str(codes_file_for_runs),
                    )
                    raise SystemExit(
                        "VWAP filter survivors too few; aborting"
                    )

                tmp = night_root / "universe_vwap_filtered.csv"
                pd.DataFrame({"code": filtered}).to_csv(tmp, index=False)
                codes_file_for_runs = tmp
                write_status(
                    state="running",
                    step="universe",
                    message="Yahoo universe ready (VWAP filter applied)",
                    universe_file=str(codes_file_for_runs),
                    universe_base=len(base_codes),
                    universe_source=universe_source_note
                    or str(codes_file_for_runs),
                )
            except Exception as exc:  # pragma: no cover
                write_status(
                    state="error",
                    step="universe",
                    message=f"VWAP filter error; aborting: {exc}",
                    universe_file=str(codes_file_for_runs),
                    universe_base=len(base_codes),
                    universe_source=universe_source_note
                    or str(codes_file_for_runs),
                )
                raise SystemExit(f"VWAP filter error; aborting: {exc}")
        else:
            if codes_file_for_runs:
                write_status(
                    state="running",
                    step="universe",
                    message="Yahoo universe ready",
                    universe_file=str(codes_file_for_runs),
                    universe_base=len(base_codes),
                    universe_source=universe_source_note
                    or str(codes_file_for_runs),
                )
            else:
                write_status(
                    state="error",
                    step="universe",
                    message="Yahoo data unavailable for yahoo-top universe; aborting",
                    universe_base=len(base_codes),
                    universe_source=universe_source_note or "yahoo-fallback",
                )
                raise SystemExit(
                    "Yahoo data unavailable for yahoo-top universe; aborting"
                )
    else:
        write_status(
            state="running",
            step="universe",
            message="Using Excel universe",
            total_plans=total_plans,
            completed_plans=0,
        )
        base_codes = load_codes_from_excel(Path(args.excel), args.excel_ticker_sheet)

    # Weekday: optional差分再計算（PF/CIが悪化した銘柄のみ）
    degraded_set: set[str] = set()
    if run_type == "weekday" and args.reopt_degraded_only:
        prev_path = Path("output/excel/candidates_nextday.csv")
        if prev_path.exists():
            try:
                df_prev = pd.read_csv(prev_path)
                if "Ticker" in df_prev.columns:
                    pf_th = float(args.reopt_pf_th)
                    ci_th = float(args.reopt_ci_th)
                    tickers = df_prev["Ticker"].astype(str).str.upper()
                    pf = pd.to_numeric(df_prev.get("forward_pf_eff", 0), errors="coerce").fillna(0)
                    ci = pd.to_numeric(df_prev.get("forward_win_ci_low", 0), errors="coerce").fillna(0)
                    mask = (pf < pf_th) | (ci < ci_th)
                    degraded_set = set(tickers[mask].tolist())
            except Exception:
                degraded_set = set()
        if degraded_set:
            base_codes = [c for c in base_codes if c.upper() in degraded_set]
        else:
            # 全銘柄を再最適化
            pass

        if not base_codes:
            write_status(
                state="success",
                step="early_exit",
                message="No degraded tickers; skipping weekday reopt",
                total_plans=0,
                completed_plans=0,
            )
            return

    if not args.disable_minute_cache:
        minute_codes = set(base_codes)
        if codes_file_for_runs and codes_file_for_runs.exists():
            try:
                df_codes = pd.read_csv(codes_file_for_runs)
                if "code" in df_codes.columns:
                    minute_codes.update(
                        df_codes["code"].dropna().astype(str).str.strip().tolist()
                    )
            except Exception:
                pass
        if minute_codes:
            tmp_path: Optional[Path] = None
            with tempfile.NamedTemporaryFile(
                "w", suffix=".csv", delete=False, encoding="utf-8", newline=""
            ) as tmp:
                tmp.write("code\n")
                for code in sorted(minute_codes):
                    tmp.write(f"{code}\n")
                tmp_path = Path(tmp.name)
            try:
                update_cmd = [
                    sys.executable,
                    "tools/update_minute_cache.py",
                    "--codes-file",
                    str(tmp_path),
                    "--universe-glob",
                    "data/universe/topvol_*.csv",
                    "--universe-limit",
                    str(args.minute_cache_limit),
                    "--history-days",
                    str(args.minute_cache_history),
                    "--backfill-days",
                    str(args.minute_backfill_days),
                ]
                run(update_cmd, cwd=repo_root)
            finally:
                if tmp_path is not None:
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass
        write_status(
            state="running",
            step="minute_cache",
            message="Minute cache refresh",
            universe_effective=len(minute_codes),
            universe_base=len(base_codes),
            universe_source=universe_source_note or args.universe_source,
        )
    candidate_frames: List[pd.DataFrame] = []
    candidate_files: List[Path] = []
    completed_plans = 0

    # Load strategy rules (for session-specific overrides like AM1000 SELL)
    rules_kv: Dict[str, str] = {}
    try:
        rules_path = repo_root / "state/strategy_rules.ini"
        if rules_path.exists():
            for line in rules_path.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                rules_kv[k.strip()] = v.strip()
    except Exception:
        pass

    for plan_idx, (sess_label, sess_start, sess_end, sig) in enumerate(plans, start=1):
        tag = f"{sess_label}_{sig}"
        plan_order.append(tag)
        write_status(
            state="running",
            step=f"{tag} coarse {plan_idx}/{total_plans}",
            message="Running coarse scan",
            total_plans=total_plans,
            completed_plans=completed_plans,
            current_plan=tag,
            plans=",".join(plan_order),
            plan_counts=format_plan_counts(plan_counts),
        )

        out_coarse = night_root / f"RUN_coarse_{tag}"
        out_refine = night_root / f"RUN_refine_{tag}"
        cand_dir = Path("output/excel") / f"NIGHTLY_{date_tag}" / tag
        cand_dir.mkdir(parents=True, exist_ok=True)

        # Session-specific directional guard tuning (AM1000 SELL auto)
        dir_bp = float(args.gap_guard_dir_bp)
        if sess_label.upper().startswith("AM1000") and sig == "j-cross":
            if rules_kv.get("am1000_sell_enabled", "0") == "1":
                dir_bp = min(dir_bp, 15.0)

        coarse_cmd = [
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel",
            args.excel,
            "--outdir",
            str(out_coarse),
            "--mode",
            "coarse",
            "--signal-mode",
            sig,
            "--session-label",
            sess_label,
            "--session-start",
            sess_start,
            "--session-end",
            sess_end,
            "--lookback",
            str(args.lookback),
            "--chunk-days",
            str(args.chunk_days),
            "--train-days",
            str(args.train_days),
            "--forward-days",
            str(args.forward_days),
            "--min-train-trades",
            str(args.min_train_trades),
            "--min-forward-trades",
            str(args.min_forward_trades),
            "--forward-pf-min",
            str(args.forward_pf_min),
            "--gap-guard-abs-bp",
            str(args.gap_guard_abs_bp),
            "--gap-guard-dir-bp",
            str(dir_bp),
            "--slipbp",
            str(args.slipbp),
            "--feebp",
            str(args.feebp),
            "--liquidity-quantile",
            str(args.liquidity_quantile),
            "--jobs",
            str(args.jobs),
            "--use-local-raw",
            "--run-type",
            run_type,
        ]
        if is_weekend:
            coarse_cmd.append("--no-cache")
        if args.enable_asha:
            coarse_cmd.append("--enable-asha")
        if args.mask_ineffective:
            coarse_cmd.extend(
                [
                    "--mask-ineffective",
                    "--mask-window",
                    str(args.mask_window),
                    "--mask-threshold",
                    str(args.mask_threshold),
                ]
            )
        if is_weekend and args.cache_refresh_weekend:
            coarse_cmd.append("--cache-refresh")
        if codes_file_for_runs:
            coarse_cmd.extend(["--codes-file", str(codes_file_for_runs)])
        if args.excel_summary:
            coarse_cmd.append("--excel-summary")
        run(coarse_cmd, cwd=repo_root, env=bt_env)

        codes_file = out_coarse / "_TOP_CANDIDATES.csv"
        if not codes_file.exists() or codes_file.stat().st_size == 0:
            plan_counts[tag] = 0
            completed_plans += 1
            write_status(
                state="running",
                step=f"{tag} skipped {plan_idx}/{total_plans}",
                message="No coarse candidates produced",
                total_plans=total_plans,
                completed_plans=completed_plans,
                current_plan=tag,
                plans=",".join(plan_order),
                plan_counts=format_plan_counts(plan_counts),
            )
            continue
        else:
            try:
                with codes_file.open("r", encoding="utf-8") as fh:
                    has_data = False
                    for idx, _ in enumerate(fh, start=1):
                        if idx > 1:
                            has_data = True
                            break
            except Exception:
                has_data = False
            if not has_data:
                plan_counts[tag] = 0
                completed_plans += 1
                write_status(
                    state="running",
                    step=f"{tag} skipped {plan_idx}/{total_plans}",
                    message="Coarse candidates file empty",
                    total_plans=total_plans,
                    completed_plans=completed_plans,
                    current_plan=tag,
                    plans=",".join(plan_order),
                    plan_counts=format_plan_counts(plan_counts),
                )
                continue

        write_status(
            state="running",
            step=f"{tag} refine {plan_idx}/{total_plans}",
            message="Running refine scan",
            total_plans=total_plans,
            completed_plans=completed_plans,
            current_plan=tag,
            plans=",".join(plan_order),
            plan_counts=format_plan_counts(plan_counts),
        )

        # Apply same dir_bp in refine phase
        refine_cmd = [
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel",
            args.excel,
            "--outdir",
            str(out_refine),
            "--mode",
            "refine",
            "--signal-mode",
            sig,
            "--session-label",
            sess_label,
            "--session-start",
            sess_start,
            "--session-end",
            sess_end,
            "--lookback",
            str(args.lookback),
            "--chunk-days",
            str(args.chunk_days),
            "--train-days",
            str(args.train_days),
            "--forward-days",
            str(args.forward_days),
            "--min-train-trades",
            str(args.min_train_trades),
            "--min-forward-trades",
            str(args.min_forward_trades),
            "--forward-pf-min",
            str(args.forward_pf_min),
            "--gap-guard-abs-bp",
            str(args.gap_guard_abs_bp),
            "--gap-guard-dir-bp",
            str(dir_bp),
            "--slipbp",
            str(args.slipbp),
            "--feebp",
            str(args.feebp),
            "--liquidity-quantile",
            str(args.liquidity_quantile),
            "--jobs",
            str(args.jobs),
            "--codes-file",
            str(codes_file),
            "--candidate-dir",
            str(cand_dir),
            "--use-local-raw",
            "--run-type",
            run_type,
        ]
        if is_weekend:
            refine_cmd.append("--no-cache")
        if args.enable_bayes:
            refine_cmd.append("--enable-bayes")
            refine_cmd.extend(["--bayes-trials", str(args.bayes_trials)])
            if args.bayes_timeout > 0:
                refine_cmd.extend(["--bayes-timeout", str(args.bayes_timeout)])
        if is_weekend and args.cache_refresh_weekend:
            refine_cmd.append("--cache-refresh")
        if args.excel_summary:
            refine_cmd.append("--excel-summary")
        if args.refine_quick_grid:
            refine_cmd.extend(["--quick-grid", "--optimize-io"])
        run(refine_cmd, cwd=repo_root, env=bt_env)

        candidates_found = 0
        candidate_path = cand_dir / f"candidates_{date_tag}.csv"
        if not candidate_path.exists():
            # Backfill runs: bt_opt30_forward may emit candidates_{today}.csv when re-running older target dates.
            # Fall back to the newest candidates_*.csv in the plan directory.
            candidates = sorted(cand_dir.glob("candidates_*.csv"), key=lambda p: p.stat().st_mtime)
            if candidates:
                candidate_path = candidates[-1]
        if candidate_path.exists():
            try:
                df = pd.read_csv(candidate_path)
                candidates_found = len(df.index)
                df["plan_tag"] = tag
                candidate_frames.append(df)
                candidate_files.append(candidate_path)
            except Exception as exc:
                write_status(
                    state="running",
                    step=f"{tag} completed {plan_idx}/{total_plans}",
                    message=f"Failed to read candidate CSV: {exc}",
                )
        plan_counts[tag] = candidates_found
        completed_plans += 1
        write_status(
            state="running",
            step=f"{tag} completed {plan_idx}/{total_plans}",
            message=f"{candidates_found} candidates collected",
            total_plans=total_plans,
            completed_plans=completed_plans,
            current_plan=tag,
            plans=",".join(plan_order),
            plan_counts=format_plan_counts(plan_counts),
        )

    write_status(
        state="running",
        step="aggregating",
        message=f"Combining {len(candidate_frames)} plan outputs",
        total_plans=total_plans,
        completed_plans=completed_plans,
        plans=",".join(plan_order),
        plan_counts=format_plan_counts(plan_counts),
    )

    ticker_union: set[str] = set()
    for frame in candidate_frames:
        if "Ticker" in frame.columns:
            ticker_union.update(frame["Ticker"].dropna().astype(str).str.upper().tolist())
    corr_map = compute_corr_map(sorted(ticker_union)) if ticker_union else {}
    vwap_stats = (
        compute_vwap_revert_stats(
            sorted(ticker_union),
            lookback_days=int(getattr(args, "vwap_lookback_days", 60)),
            j_threshold=float(getattr(args, "vwap_j_threshold", 1.0)),
            max_bars=int(getattr(args, "vwap_max_bars", 30)),
        )
        if ticker_union
        else {}
    )

    out_all = Path("output/excel") / "candidates_nextday.csv"
    summary = aggregate_candidates(
        candidate_frames,
        out_all,
        min_forward_ci=float(getattr(args, "min_forward_ci", 0.65)),
        min_forward_winrate=float(getattr(args, "min_forward_winrate", 0.0)),
        min_index_corr=float(getattr(args, "min_index_corr", 0.2)),
        min_vwap_revert=float(getattr(args, "min_vwap_revert", 0.0)),
        corr_map=corr_map,
        vwap_stats=vwap_stats,
        run_type=run_type,
    )
    summary.update(
        {
            "plans": ",".join(plan_order),
            "plan_counts": format_plan_counts(plan_counts),
            "candidate_files": str(len(candidate_files)),
            "candidates_path": str(out_all.resolve()),
            "corr_map": str(len(corr_map)),
            "vwap_stats": str(len(vwap_stats)),
        }
    )

    # Also keep a dated snapshot of candidates for DailyReplay
    try:
        date_tag = date_tag  # type: ignore[name-defined]
        snapshot_name = f"candidates_for_{date_tag}.csv"
        snapshot_path = Path("output/excel") / snapshot_name
        shutil.copy2(out_all, snapshot_path)
        summary["candidates_snapshot"] = str(snapshot_path.resolve())
    except Exception:
        pass

    # Fallback aggregator when no candidates from in-process aggregation
    try:
        total_cand = int(summary.get("total_candidates", "0"))
    except Exception:
        total_cand = 0
    if total_cand <= 0 or not out_all.exists():
        try:
            run([sys.executable, "tools/aggregate_candidates_today.py", "--output", str(out_all)], cwd=repo_root)
            # best-effort stats refresh
            try:
                df_out = pd.read_csv(out_all)
                summary.update(
                    {
                        "total_candidates": str(int(len(df_out))),
                        "unique_tickers": str(int(df_out["Ticker"].nunique())) if "Ticker" in df_out.columns else str(int(len(df_out))),
                    }
                )
            except Exception:
                pass
        except SystemExit:
            pass

    coeff_latest = repo_root / "output/excel/dashboard_coeffs_latest.csv"
    if not args.disable_dashboard_coeffs:
        try:
            run(
                [
                    sys.executable,
                    "tools/compute_dashboard_coeffs.py",
                    "--codes-file",
                    str(out_all),
                    "--history-days",
                    str(args.coeff_history_days),
                ],
                cwd=repo_root,
            )
        except SystemExit:
            pass
    enrich_dashboard_columns(out_all, coeff_latest)
    apply_trend_preferences(out_all, (repo_root / args.trend_pref).resolve(), args.trend_bp_th)

    if not args.headless:
        ensure_dashboard_formulas(repo_root, excel_path)
        try:
            run(
                [
                    sys.executable,
                    "scripts/run_macros_on_copy.py",
                    "--excel",
                    str(excel_path),
                ],
                cwd=repo_root,
            )
        except SystemExit:
            pass

    write_status(
        state="success",
        step="completed",
        message="Nightly batch completed",
        **summary,
    )

    # Post-run analytics: generate param stats + auto-update masks + update optuna priors
    reports_root = Path("reports/param_stats") / night_root.name
    try:
        run([sys.executable, "tools/analyze_param_stats.py", "--root", str(night_root)], cwd=repo_root)
    except SystemExit:
        pass
    by_j_path = reports_root / "by_J_th.csv"
    if by_j_path.exists():
        try:
            cmd = [
                sys.executable,
                "tools/update_ineffective_bands.py",
                "--run-root",
                str(reports_root),
            ]
            if is_weekend:
                cmd.append("--allow-unmask")
            run(cmd, cwd=repo_root)
        except SystemExit:
            pass
    # Update priors: weekend replaces weekend set; weekday appends supplemental seeds
    try:
        run(
            [
                sys.executable,
                "tools/update_optuna_priors.py",
                "--run-root",
                str(night_root),
                "--source",
                "weekend" if is_weekend else "weekday",
                "--top-k",
                str(48 if is_weekend else 12),
            ],
            cwd=repo_root,
        )
    except SystemExit:
        pass

    # Slippage diagnostics (next-minute adverse move and intrabar ranges)
    try:
        run(
            [
                sys.executable,
                "tools/analyze_slippage.py",
                "--run-root",
                str(night_root),
                "--output",
                str(night_root / "slippage_detail.csv"),
                "--plan-output",
                str(night_root / "slippage_plan_summary.csv"),
                "--recommend-output",
                str(repo_root / "output/excel/slippage_overrides.csv"),
            ],
            cwd=repo_root,
        )
    except SystemExit:
        pass

    # Walk-forward summary artifacts for downstream reporting
    try:
        run(
            [
                sys.executable,
                "tools/walk_forward_report.py",
                "--run-root",
                str(night_root),
                "--output",
                str(night_root / "walk_forward_detail.csv"),
                "--plan-output",
                str(night_root / "walk_forward_plan_summary.csv"),
            ],
            cwd=repo_root,
        )
    except SystemExit:
        pass


def main() -> None:
    try:
        _main_impl()
    except Exception as exc:
        write_status(state="error", step="error", message=str(exc))
        raise


if __name__ == "__main__":
    main()
