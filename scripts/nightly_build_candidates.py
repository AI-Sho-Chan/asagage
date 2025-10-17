import argparse
import datetime as dt
import subprocess
from pathlib import Path
import sys
import pandas as pd
import numpy as np
from yahooquery import Ticker


def run(cmd, cwd):
    print("[run]", " ".join(cmd))
    proc = subprocess.run(cmd, cwd=cwd)
    if proc.returncode != 0:
        raise SystemExit(proc.returncode)


def main():
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
    ap.add_argument("--universe-mode", choices=["excel", "yahoo-top"], default="excel",
                    help="excel: use Excel Ticker sheet. yahoo-top: rank by previous day notional and pick top N")
    ap.add_argument("--universe-size", type=int, default=300)
    ap.add_argument("--universe-source", default="data/universe_tse_prime.csv",
                    help="Optional CSV with 'code' column for base universe when yahoo-top is used. If missing, falls back to Excel Ticker sheet.")
    ap.add_argument("--excel-ticker-sheet", default="Ticker",
                    help="Sheet name in Excel holding codes (column header 'Code')")
    ap.add_argument("--universe-metric", choices=["amt", "vol"], default="amt",
                    help="Ranking metric for yahoo-top: amt (notional) or vol (volume)")
    args = ap.parse_args()

    base = Path(args.base_out)
    date_tag = dt.datetime.now().strftime("%Y%m%d")
    night_root = base / f"NIGHTLY_{date_tag}"
    night_root.mkdir(parents=True, exist_ok=True)

    plans = [
        ("AM10", "09:10", "j-only"),
        ("AM10", "09:10", "j-cross"),
        ("AM15", "09:15", "j-only"),
        ("AM15", "09:15", "j-cross"),
    ]

    candidate_csvs = []
    # Build universe if yahoo-top mode
    codes_file_for_runs: Path | None = None
    if args.universe_mode == "yahoo-top":
        base_codes = []
        src_path = Path(args.universe_source)
        if src_path.exists():
            try:
                dfu = pd.read_csv(src_path)
                base_codes = dfu['code'].astype(str).tolist()
            except Exception:
                base_codes = []
        if not base_codes:
            # fallback to Excel sheet 'Ticker' column 'Code'
            try:
                import pandas as pd  # noqa
                xl = pd.read_excel(args.excel, sheet_name=args.excel_ticker_sheet, usecols='A', header=0)
                base_codes = xl['Code'].dropna().astype(str).tolist()
            except Exception:
                base_codes = []
        if not base_codes:
            print("No base universe available for yahoo-top. Falling back to excel mode.")
        else:
            # fetch last 2 days 1m bars, compute per-code aggregate on latest day
            print(f"Ranking universe by previous day {args.universe_metric} from Yahoo (base={len(base_codes)} codes)...")
            ticker = Ticker(base_codes, asynchronous=True)
            end = dt.date.today() + dt.timedelta(days=1)
            start = dt.date.today() - dt.timedelta(days=2)
            hist = ticker.history(start=str(start), end=str(end), interval="1m")
            if isinstance(hist, pd.DataFrame) and not hist.empty:
                df = hist.reset_index()
                if 'symbol' in df.columns:
                    df = df.rename(columns={'symbol': 'code'})
                if 'date' in df.columns and 'ts' not in df.columns:
                    df = df.rename(columns={'date': 'ts'})
                df['ts'] = pd.to_datetime(df['ts'])
                df['date'] = df['ts'].dt.date
                df['amt'] = df['close'] * df['volume']
                last_day = df['date'].max()
                dlast = df[df['date'] == last_day]
                if args.universe_metric == 'amt':
                    metric = dlast.groupby('code')['amt'].sum().reset_index(name='score')
                else:
                    metric = dlast.groupby('code')['volume'].sum().reset_index(name='score')
                topn = metric.sort_values('score', ascending=False).head(int(args.universe_size))
                uni_dir = night_root
                uni_dir.mkdir(parents=True, exist_ok=True)
                codes_file_for_runs = uni_dir / f"universe_{args.universe_metric}_top_{args.universe_size}_{last_day}.csv"
                topn[['code']].to_csv(codes_file_for_runs, index=False)
                print("Universe ranked and saved:", codes_file_for_runs)
            else:
                print("Yahoo history fetch returned empty; cannot build yahoo-top universe.")
    for sess_label, sess_end, sig in plans:
        tag = f"{sess_label}_{sig}"
        out_coarse = night_root / f"RUN_coarse_{tag}"
        out_refine = night_root / f"RUN_refine_{tag}"
        cand_dir = Path("output/excel") / f"NIGHTLY_{date_tag}" / tag
        cand_dir.mkdir(parents=True, exist_ok=True)

        run([
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel", args.excel,
            "--outdir", str(out_coarse),
            "--mode", "coarse",
            "--signal-mode", sig,
            "--session-start", "09:00",
            "--session-end", sess_end,
            "--lookback", str(args.lookback),
            "--chunk-days", str(args.chunk_days),
            "--train-days", str(args.train_days),
            "--forward-days", str(args.forward_days),
            "--min-forward-trades", str(args.min_forward_trades),
            "--forward-pf-min", str(args.forward_pf_min),
            "--gap-guard-abs-bp", str(args.gap_guard_abs_bp),
            "--gap-guard-dir-bp", str(args.gap_guard_dir_bp),
            "--slipbp", str(args.slipbp),
            "--feebp", str(args.feebp),
            "--liquidity-quantile", str(args.liquidity_quantile),
        ] + (["--codes-file", str(codes_file_for_runs)] if codes_file_for_runs else [])
          + (["--excel-summary"] if args.excel_summary else []),
        cwd=str(Path(__file__).resolve().parent.parent))

        codes_file = out_coarse / "_TOP_CANDIDATES.csv"
        # run refine only if coarse produced something
        if not codes_file.exists() or codes_file.stat().st_size == 0:
            continue
        run([
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel", args.excel,
            "--outdir", str(out_refine),
            "--mode", "refine",
            "--signal-mode", sig,
            "--session-start", "09:00",
            "--session-end", sess_end,
            "--lookback", str(args.lookback),
            "--chunk-days", str(args.chunk_days),
            "--train-days", str(args.train_days),
            "--forward-days", str(args.forward_days),
            "--min-forward-trades", str(args.min_forward_trades),
            "--forward-pf-min", str(args.forward_pf_min),
            "--gap-guard-abs-bp", str(args.gap_guard_abs_bp),
            "--gap-guard-dir-bp", str(args.gap_guard_dir_bp),
            "--slipbp", str(args.slipbp),
            "--feebp", str(args.feebp),
            "--liquidity-quantile", str(args.liquidity_quantile),
            "--codes-file", str(codes_file),
            "--candidate-dir", str(cand_dir),
        ] + (["--excel-summary"] if args.excel_summary else []), cwd=str(Path(__file__).resolve().parent.parent))

        # pick candidate file if exists
        cand = next(cand_dir.glob(f"candidates_{date_tag}.csv"), None)
        if cand and cand.exists():
            candidate_csvs.append(cand)

    # aggregate to candidates_nextday.csv
    out_all = Path("output/excel") / "candidates_nextday.csv"
    if candidate_csvs:
        frames = []
        for p in candidate_csvs:
            df = pd.read_csv(p)
            frames.append(df)
        combined = pd.concat(frames, ignore_index=True)
        # dedup by Ticker choosing best forward_pf_eff then forward_trades
        combined = combined.sort_values(["forward_pf_eff", "forward_trades"], ascending=[False, False])
        combined = combined.drop_duplicates(subset=["Ticker"], keep="first")
        combined.to_csv(out_all, index=False, encoding="utf-8-sig")
        print("wrote", out_all, "rows", len(combined))
    else:
        print("no candidate CSVs found; nothing written")


if __name__ == "__main__":
    main()
