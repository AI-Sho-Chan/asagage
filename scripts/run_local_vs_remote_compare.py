import argparse
import subprocess
import sys
from pathlib import Path
from typing import List, Tuple

import pandas as pd


PLANS: List[Tuple[str, str, str]] = [
    ("AM0930", "09:30", "j-only"),
    ("AM0930", "09:30", "j-cross"),
    ("AM0945", "09:45", "j-only"),
    ("AM0945", "09:45", "j-cross"),
    ("AM1000", "10:00", "j-only"),
    ("AM1000", "10:00", "j-cross"),
    ("AM1015", "10:15", "j-only"),
    ("AM1015", "10:15", "j-cross"),
    ("AM1030", "10:30", "j-only"),
    ("AM1030", "10:30", "j-cross"),
]


def run_cmd(cmd: List[str]) -> None:
    print("[run]", " ".join(cmd), flush=True)
    proc = subprocess.run(cmd)
    if proc.returncode != 0:
        print("Command failed:", cmd, file=sys.stderr)
        sys.exit(proc.returncode)


def collect_summary(base: Path) -> pd.DataFrame:
    rows = []
    for d in base.glob("RUN_refine_*_*"):
        f = d / "_SUMMARY_FORWARD.csv"
        if not f.exists():
            continue
        try:
            df = pd.read_csv(f)
        except Exception:
            continue
        if df.empty:
            continue
        # expect columns
        need = {"code", "forward_exp_bp", "forward_trades", "session", "signal_mode"}
        if not need.issubset(set(df.columns)):
            # fallback: infer from path
            parts = d.name.split("_")
            sess = parts[2] if len(parts) >= 4 else ""
            mode = parts[3] if len(parts) >= 4 else ""
            df["session"] = sess
            df["signal_mode"] = mode
        rows.append(df)
    if not rows:
        return pd.DataFrame()
    return pd.concat(rows, ignore_index=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidates", default="output/excel/candidates_nextday.csv")
    ap.add_argument("--lookback", type=int, default=25)
    ap.add_argument("--chunk-days", type=int, default=5)
    ap.add_argument("--jobs", type=int, default=0)
    ap.add_argument("--out-local", default="output/bt30_compare_local")
    ap.add_argument("--out-remote", default="output/bt30_compare_remote")
    ap.add_argument("--train-days", type=int, default=12)
    ap.add_argument("--forward-days", type=int, default=5)
    args = ap.parse_args()

    cand = Path(args.candidates)
    df = pd.read_csv(cand)
    codes = sorted(df["Ticker"].dropna().astype(str).unique().tolist())
    codes_file = Path(args.out_local) / "codes_uniq.csv"
    codes_file.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame({"code": codes}).to_csv(codes_file, index=False)

    # run both local and remote refine for all plans
    for sess_label, sess_end, sig in PLANS:
        tag = f"{sess_label}_{sig}"
        # local
        out_local = Path(args.out_local) / f"RUN_refine_{tag}"
        out_local.mkdir(parents=True, exist_ok=True)
        run_cmd([
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel", "C:/AI/asagake/SHINSOKU.xlsm",
            "--outdir", str(out_local),
            "--mode", "refine",
            "--signal-mode", sig,
            "--session-start", "09:00",
            "--session-end", sess_end,
            "--lookback", str(args.lookback),
            "--chunk-days", str(args.chunk_days),
            "--train-days", str(args.train_days),
            "--forward-days", str(args.forward_days),
            "--jobs", str(args.jobs),
            "--codes-file", str(codes_file),
            "--use-local-raw",
        ])
        # remote
        out_remote = Path(args.out_remote) / f"RUN_refine_{tag}"
        out_remote.mkdir(parents=True, exist_ok=True)
        run_cmd([
            sys.executable,
            "scripts/bt_opt30_forward.py",
            "--excel", "C:/AI/asagake/SHINSOKU.xlsm",
            "--outdir", str(out_remote),
            "--mode", "refine",
            "--signal-mode", sig,
            "--session-start", "09:00",
            "--session-end", sess_end,
            "--lookback", str(args.lookback),
            "--chunk-days", str(args.chunk_days),
            "--train-days", str(args.train_days),
            "--forward-days", str(args.forward_days),
            "--jobs", str(args.jobs),
            "--codes-file", str(codes_file),
        ])

    # collect and compare
    df_local = collect_summary(Path(args.out_local))
    df_local["source"] = "local"
    df_remote = collect_summary(Path(args.out_remote))
    df_remote["source"] = "remote"
    both = pd.concat([df_local, df_remote], ignore_index=True)
    both.to_csv(Path(args.out_local).parent / "_COMPARE_RAW.csv", index=False, encoding="utf-8-sig")

    # pivot and compute deltas per (code, session, signal_mode)
    key = ["code", "session", "signal_mode"]
    met = both.groupby(key + ["source"]).agg(
        forward_exp_bp=("forward_exp_bp", "mean"),
        forward_trades=("forward_trades", "sum"),
    ).reset_index()
    loc = met[met["source"] == "local"].drop(columns=["source"]).rename(
        columns={"forward_exp_bp": "exp_bp_local", "forward_trades": "trades_local"}
    )
    rem = met[met["source"] == "remote"].drop(columns=["source"]).rename(
        columns={"forward_exp_bp": "exp_bp_remote", "forward_trades": "trades_remote"}
    )
    cmp_df = pd.merge(loc, rem, on=key, how="outer")
    cmp_df["delta_exp_bp"] = cmp_df["exp_bp_local"] - cmp_df["exp_bp_remote"]
    cmp_df["delta_trades"] = cmp_df["trades_local"] - cmp_df["trades_remote"]
    cmp_df.to_csv(Path(args.out_local).parent / "_COMPARE_SUMMARY.csv", index=False, encoding="utf-8-sig")
    print("Comparison written to:", str(Path(args.out_local).parent / "_COMPARE_SUMMARY.csv"))


if __name__ == "__main__":
    main()
