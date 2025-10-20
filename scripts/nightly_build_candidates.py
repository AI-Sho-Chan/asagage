import argparse
import datetime as dt
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
from yahooquery import Ticker

STATUS_PATH = Path("logs/nightly_status.txt")
STATUS_DATA: Dict[str, str] = {}


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


def run(cmd: List[str], cwd: Path) -> None:
    """Run a subprocess and bubble up failures with context."""
    print("[run]", " ".join(cmd))
    proc = subprocess.run(cmd, cwd=cwd)
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


def aggregate_candidates(frames: List[pd.DataFrame], out_path: Path) -> Dict[str, str]:
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

    combined = pd.concat(frames, ignore_index=True)

    sort_cols: List[str] = []
    ascending: List[bool] = []
    if "forward_pf_eff" in combined.columns:
        sort_cols.append("forward_pf_eff")
        ascending.append(False)
    if "forward_trades" in combined.columns:
        sort_cols.append("forward_trades")
        ascending.append(False)
    if sort_cols:
        combined = combined.sort_values(sort_cols, ascending=ascending)

    if "Ticker" in combined.columns:
        combined = combined.drop_duplicates(subset=["Ticker"], keep="first")

    out_path.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(out_path, index=False, encoding="utf-8-sig")

    summary["total_candidates"] = str(int(len(combined)))
    if "Ticker" in combined.columns:
        summary["unique_tickers"] = str(int(combined["Ticker"].nunique()))
    else:
        summary["unique_tickers"] = summary["total_candidates"]

    summary["avg_forward_winrate"] = ensure_numeric_mean(combined, "forward_winrate")
    summary["avg_forward_pf"] = ensure_numeric_mean(combined, "forward_pf_eff")
    summary["avg_expected_return"] = ensure_numeric_mean(combined, "forward_exp_boot_mean")
    summary["avg_forward_trades"] = ensure_numeric_mean(combined, "forward_trades", decimals=2)

    return summary


def format_plan_counts(plan_counts: Dict[str, int]) -> str:
    return ",".join(f"{key}:{value}" for key, value in plan_counts.items())


def build_yahoo_universe(
    base_codes: List[str],
    metric: str,
    universe_size: int,
    output_dir: Path,
) -> Optional[Path]:
    if not base_codes:
        return None

    ticker = Ticker(base_codes, asynchronous=True)
    end_date = dt.date.today() + dt.timedelta(days=1)
    start_date = dt.date.today() - dt.timedelta(days=2)
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
    ap.add_argument("--excel", default="SHINSOKU.xlsm")
    ap.add_argument("--base-out", default="output/bt30")
    ap.add_argument("--lookback", type=int, default=60)
    ap.add_argument("--chunk-days", type=int, default=5)
    ap.add_argument("--train-days", type=int, default=12)
    ap.add_argument("--forward-days", type=int, default=4)
    ap.add_argument("--min-forward-trades", type=int, default=10)
    ap.add_argument("--forward-pf-min", type=float, default=1.3)
    ap.add_argument("--gap-guard-abs-bp", type=float, default=80.0)
    ap.add_argument("--gap-guard-dir-bp", type=float, default=40.0)
    ap.add_argument("--slipbp", type=float, default=4.0)
    ap.add_argument("--feebp", type=float, default=4.0)
    ap.add_argument("--liquidity-quantile", type=float, default=0.5)
    ap.add_argument("--excel-summary", action="store_true")
    ap.add_argument("--universe-mode", choices=["excel", "yahoo-top"], default="excel")
    ap.add_argument("--universe-size", type=int, default=300)
    ap.add_argument("--universe-source", default="data/universe_tse_prime.csv")
    ap.add_argument("--excel-ticker-sheet", default="Ticker")
    ap.add_argument("--universe-metric", choices=["amt", "vol"], default="amt")
    args = ap.parse_args()

    base = Path(args.base_out)
    date_tag = dt.datetime.now().strftime("%Y%m%d")
    night_root = base / f"NIGHTLY_{date_tag}"
    night_root.mkdir(parents=True, exist_ok=True)

    plans: List[Tuple[str, str, str]] = [
        ("AM10", "09:10", "j-only"),
        ("AM10", "09:10", "j-cross"),
        ("AM15", "09:15", "j-only"),
        ("AM15", "09:15", "j-cross"),
    ]
    total_plans = len(plans)
    plan_counts: Dict[str, int] = {f"{label}_{sig}": 0 for label, _, sig in plans}
    plan_order: List[str] = []

    write_status(
        state="running",
        step="initializing",
        message="Starting nightly batch",
        base_out=str(base),
        date_tag=date_tag,
        total_plans=total_plans,
    )

    codes_file_for_runs: Optional[Path] = None
    if args.universe_mode == "yahoo-top":
        write_status(
            state="running",
            step="universe",
            message="Building Yahoo-top universe",
            total_plans=total_plans,
            completed_plans=0,
        )

        base_codes: List[str] = []
        src_path = Path(args.universe_source)
        if src_path.exists():
            try:
                dfu = pd.read_csv(src_path)
                if "code" in dfu.columns:
                    base_codes = dfu["code"].dropna().astype(str).tolist()
            except Exception:
                base_codes = []

        if not base_codes:
            base_codes = load_codes_from_excel(Path(args.excel), args.excel_ticker_sheet)
            if base_codes:
                write_status(
                    state="running",
                    step="universe",
                    message="Using Excel ticker sheet as Yahoo base",
                    universe_base=len(base_codes),
                )

        if not base_codes:
            write_status(
                state="running",
                step="universe",
                message="No universe codes available; falling back to Excel-only mode",
                universe_base=0,
            )
        else:
            try:
                codes_file_for_runs = build_yahoo_universe(
                    base_codes=base_codes,
                    metric=args.universe_metric,
                    universe_size=args.universe_size,
                    output_dir=night_root,
                )
            except Exception as exc:
                codes_file_for_runs = None
                write_status(
                    state="running",
                    step="universe",
                    message=f"Yahoo universe build failed: {exc}",
                    universe_base=len(base_codes),
                )

            if codes_file_for_runs:
                write_status(
                    state="running",
                    step="universe",
                    message="Yahoo universe ready",
                    universe_file=str(codes_file_for_runs),
                )
            else:
                write_status(
                    state="running",
                    step="universe",
                    message="Yahoo data unavailable; using Excel universe",
                )
    else:
        write_status(
            state="running",
            step="universe",
            message="Using Excel universe",
            total_plans=total_plans,
            completed_plans=0,
        )

    repo_root = Path(__file__).resolve().parent.parent
    candidate_frames: List[pd.DataFrame] = []
    candidate_files: List[Path] = []
    completed_plans = 0

    for plan_idx, (sess_label, sess_end, sig) in enumerate(plans, start=1):
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

        run(
            [
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
                "--session-start",
                "09:00",
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
                "--min-forward-trades",
                str(args.min_forward_trades),
                "--forward-pf-min",
                str(args.forward_pf_min),
                "--gap-guard-abs-bp",
                str(args.gap_guard_abs_bp),
                "--gap-guard-dir-bp",
                str(args.gap_guard_dir_bp),
                "--slipbp",
                str(args.slipbp),
                "--feebp",
                str(args.feebp),
                "--liquidity-quantile",
                str(args.liquidity_quantile),
            ]
            + (["--codes-file", str(codes_file_for_runs)] if codes_file_for_runs else [])
            + (["--excel-summary"] if args.excel_summary else []),
            cwd=repo_root,
        )

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

        run(
            [
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
                "--session-start",
                "09:00",
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
                "--min-forward-trades",
                str(args.min_forward_trades),
                "--forward-pf-min",
                str(args.forward_pf_min),
                "--gap-guard-abs-bp",
                str(args.gap_guard_abs_bp),
                "--gap-guard-dir-bp",
                str(args.gap_guard_dir_bp),
                "--slipbp",
                str(args.slipbp),
                "--feebp",
                str(args.feebp),
                "--liquidity-quantile",
                str(args.liquidity_quantile),
                "--codes-file",
                str(codes_file),
                "--candidate-dir",
                str(cand_dir),
            ]
            + (["--excel-summary"] if args.excel_summary else []),
            cwd=repo_root,
        )

        candidates_found = 0
        candidate_path = next(cand_dir.glob(f"candidates_{date_tag}.csv"), None)
        if candidate_path and candidate_path.exists():
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

    out_all = Path("output/excel") / "candidates_nextday.csv"
    summary = aggregate_candidates(candidate_frames, out_all)
    summary.update(
        {
            "plans": ",".join(plan_order),
            "plan_counts": format_plan_counts(plan_counts),
            "candidate_files": str(len(candidate_files)),
            "candidates_path": str(out_all.resolve()),
        }
    )

    write_status(
        state="success",
        step="completed",
        message="Nightly batch completed",
        **summary,
    )


def main() -> None:
    try:
        _main_impl()
    except Exception as exc:
        write_status(state="error", step="error", message=str(exc))
        raise


if __name__ == "__main__":
    main()
