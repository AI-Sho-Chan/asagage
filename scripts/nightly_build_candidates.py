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
WORKBOOK_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
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
    """Combine plan outputs, enforce quality filters, and pick one combo per ticker.

    Filters (hard):
      - J_th >= 0.8
      - forward_win_ci_low >= 0.70
      - forward_pf_eff >= 1.30
      - forward_trades >= 5

    One-per-ticker selection:
      score = forward_pf_eff * (forward_winrate ** 1.2) * log1p(forward_trades) / (1 + MaxDD/1000)
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

    # Hard filters
    if c("J_th") in df.columns:
        df = df[num(c("J_th")) >= 0.8]
    if c("forward_win_ci_low") in df.columns:
        df = df[num(c("forward_win_ci_low")) >= 0.70]
    if c("forward_pf_eff") in df.columns:
        df = df[num(c("forward_pf_eff")) >= 1.30]
    if c("forward_trades") in df.columns:
        df = df[num(c("forward_trades")) >= 5]

    # Score
    pf = num(c("forward_pf_eff"))
    win = num(c("forward_winrate"))
    trades = num(c("forward_trades"))
    dd = num("MaxDD")
    # log1p from numpy via pandas
    score = pf * (win ** 1.2) * (trades.add(1).apply(np.log1p)) / (1.0 + dd / 1000.0)
    df["_score"] = score

    # One per ticker (best score)
    if "Ticker" in df.columns:
        df = df.sort_values(["Ticker", "_score"], ascending=[True, False])
        df = df.drop_duplicates(subset=["Ticker"], keep="first")
    else:
        df = df.sort_values(["_score"], ascending=[False])

    out_path.parent.mkdir(parents=True, exist_ok=True)
    if "_score" in df.columns:
        df = df.drop(columns=["_score"])
    df.to_csv(out_path, index=False, encoding="utf-8-sig")

    summary["total_candidates"] = str(int(len(df)))
    summary["unique_tickers"] = str(int(df["Ticker"].nunique())) if "Ticker" in df.columns else summary["total_candidates"]
    summary["avg_forward_winrate"] = ensure_numeric_mean(df, c("forward_winrate"))
    summary["avg_forward_pf"] = ensure_numeric_mean(df, c("forward_pf_eff"))
    summary["avg_expected_return"] = ensure_numeric_mean(df, c("forward_exp_boot_mean"))
    summary["avg_forward_trades"] = ensure_numeric_mean(df, c("forward_trades"), decimals=2)
    return summary


def ensure_dashboard_formulas(repo_root: Path) -> None:
    """Restore formulas via COM and verify key RSS columns.

    Note: We prefer restore_dashboard_formulas.py as it reapplies the canonical
    formulas (I–V) including SignalStatus/Kind, and re-protects the sheet.
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
        wb = excel.Workbooks.Open(str(WORKBOOK_PATH))
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
    ap.add_argument("--min-train-trades", type=int, default=15)
    ap.add_argument("--min-forward-trades", type=int, default=3)
    ap.add_argument("--forward-pf-min", type=float, default=1.3)
    ap.add_argument("--gap-guard-abs-bp", type=float, default=80.0)
    ap.add_argument("--gap-guard-dir-bp", type=float, default=40.0)
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
    ap.add_argument("--enable-asha", action="store_true", help="Pass --enable-asha to coarse runs")
    ap.add_argument("--enable-bayes", action="store_true", help="Pass --enable-bayes to refine runs")
    ap.add_argument("--bayes-trials", type=int, default=40, help="Bayesian trials per refine run when enabled")
    ap.add_argument("--bayes-timeout", type=int, default=0, help="Bayesian timeout seconds (0=disabled)")
    ap.add_argument("--mask-ineffective", action="store_true", help="Enable ineffective-band masking during coarse runs")
    ap.add_argument("--mask-window", type=int, default=20, help="Mask history window (runs)")
    ap.add_argument("--mask-threshold", type=float, default=1.05, help="Forward pf_eff threshold for mask retention")
    ap.add_argument("--cache-refresh-weekend", action="store_true", help="Force --cache-refresh on weekend runs")
    ap.add_argument("--analysis-ledger", action="store_true", help="Pass --analysis-ledger to refine runs")
    ap.add_argument(
        "--enable-market-features",
        action="store_true",
        help="Attach --enable-market-features to coarse/refine runs for hypothesis capture",
    )
    ap.add_argument(
        "--enable-rd-windows",
        action="store_true",
        help="Append optional mid/pm R&D windows (weekend only)",
    )
    args = ap.parse_args()

    run_type = args.run_type
    if run_type == "auto":
        run_type = "weekend" if dt.datetime.now().weekday() >= 5 else "weekday"
    is_weekend = run_type == "weekend"

    base = Path(args.base_out)
    date_tag = dt.datetime.now().strftime("%Y%m%d")
    night_root = base / f"NIGHTLY_{date_tag}"
    night_root.mkdir(parents=True, exist_ok=True)

    # Plan entries: (label, session_start, session_end, signal_mode)
    plans: List[Tuple[str, str, str, str]] = [
        ("AM0930", "09:00", "09:30", "j-only"),
        ("AM0930", "09:00", "09:30", "j-cross"),
        ("AM0945", "09:00", "09:45", "j-only"),
        ("AM0945", "09:00", "09:45", "j-cross"),
        ("AM1000", "09:00", "10:00", "j-only"),
        ("AM1000", "09:00", "10:00", "j-cross"),
        ("AM1015", "09:00", "10:15", "j-only"),
        ("AM1015", "09:00", "10:15", "j-cross"),
        ("AM1030", "09:00", "10:30", "j-only"),
        ("AM1030", "09:00", "10:30", "j-cross"),
    ]

    # Optional R&D windows for weekend runs (mid/pm time slices)
    if is_weekend and getattr(args, "enable_rd_windows", False):
        plans.extend(
            [
                ("MID1030", "10:30", "11:00", "j-cross"),
                ("PM1230", "12:30", "13:00", "j-cross"),
            ]
        )
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
        run_type=run_type,
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
        if "*" in src_path.name or "?" in src_path.name:
            parent = src_path.parent if src_path.parent != Path("") else Path(".")
            matches = sorted(parent.glob(src_path.name))
            if matches:
                src_path = matches[-1]
        elif src_path.is_dir():
            matches = sorted(src_path.glob("topvol_*.csv"))
            if matches:
                src_path = matches[-1]
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
            str(args.gap_guard_dir_bp),
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
        if args.enable_asha:
            coarse_cmd.append("--enable-asha")
        if args.mask_ineffective and not is_weekend:
            coarse_cmd.extend(
                [
                    "--mask-ineffective",
                    "--mask-window",
                    str(args.mask_window),
                    "--mask-threshold",
                    str(args.mask_threshold),
                    "--mask-keep-j-min",
                    "1.35",
                ]
            )
        if is_weekend and args.cache_refresh_weekend:
            coarse_cmd.append("--cache-refresh")
        if args.enable_market_features:
            coarse_cmd.extend(["--enable-market-features", "--market-adjust-j", "--market-j-delta-up", "0.10", "--market-j-delta-down", "0.10"])
            # 動的TP/SLの実験は粗段階でも軽くオン（効果検証用）
            coarse_cmd.extend(["--dynamic-risk-j", "--tp-per-j", "0.15", "--sl-per-j", "0.10"])
        if codes_file_for_runs:
            coarse_cmd.extend(["--codes-file", str(codes_file_for_runs)])
        if args.excel_summary:
            coarse_cmd.append("--excel-summary")
        run(coarse_cmd, cwd=repo_root)

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
            str(args.gap_guard_dir_bp),
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
        if args.enable_bayes:
            refine_cmd.append("--enable-bayes")
            refine_cmd.extend(["--bayes-trials", str(args.bayes_trials)])
            if args.bayes_timeout > 0:
                refine_cmd.extend(["--bayes-timeout", str(args.bayes_timeout)])
        if is_weekend and args.cache_refresh_weekend:
            refine_cmd.append("--cache-refresh")
        if args.excel_summary:
            refine_cmd.append("--excel-summary")
        if args.enable_market_features:
            refine_cmd.extend(["--enable-market-features", "--market-adjust-j", "--market-j-delta-up", "0.10", "--market-j-delta-down", "0.10"])
            refine_cmd.extend(["--dynamic-risk-j", "--tp-per-j", "0.15", "--sl-per-j", "0.10"])
        if args.analysis_ledger:
            refine_cmd.append("--analysis-ledger")
        run(refine_cmd, cwd=repo_root)

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

    ensure_dashboard_formulas(repo_root)

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


def main() -> None:
    try:
        _main_impl()
    except Exception as exc:
        write_status(state="error", step="error", message=str(exc))
        raise


if __name__ == "__main__":
    main()
