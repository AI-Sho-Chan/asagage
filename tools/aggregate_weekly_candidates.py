import argparse
import json
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd


def load_plan_candidates(root: Path, date_tag: str) -> List[pd.DataFrame]:
    frames: List[pd.DataFrame] = []
    for plan_dir in sorted((root / f"NIGHTLY_{date_tag}").glob("*")):
        path = plan_dir / f"candidates_{date_tag}.csv"
        if path.exists() and path.stat().st_size > 0:
            try:
                df = pd.read_csv(path)
                df["plan_tag"] = plan_dir.name
                frames.append(df)
            except Exception:
                continue
    return frames


def select_weekly_candidates(df: pd.DataFrame, target_top: int) -> pd.DataFrame:
    cols = {c.lower(): c for c in df.columns}

    def c(name: str) -> str:
        return cols.get(name.lower(), name)

    def num(col: str) -> pd.Series:
        if col in df.columns:
            return pd.to_numeric(df[col], errors="coerce").fillna(0)
        return pd.Series(0.0, index=df.index)

    # Normalise ticker列
    if "Ticker" not in df.columns and c("code") in df.columns:
        df = df.rename(columns={c("code"): "Ticker"})
        cols = {c.lower(): c for c in df.columns}

    # forward_exp_bp フォールバック
    if "forward_exp_bp" not in df.columns:
        fallback = c("forward_exp_boot_mean")
        if fallback in df.columns:
            df["forward_exp_bp"] = pd.to_numeric(df[fallback], errors="coerce").fillna(0)
            cols = {c.lower(): c for c in df.columns}

    # Hard filters for "勝ちやすさ"
    mask = (
        (num(c("forward_winrate")) >= 0.70)
        & (num(c("forward_pf_eff")) >= 1.30)
        & (num(c("forward_exp_bp")) > 0.0)
        & (num(c("forward_trades")) >= 5)
    )
    pool = df.loc[mask].copy()
    if pool.empty:
        return pool

    # Score: PF重視 + 勝率 + トレード数、DDはペナルティ
    pf = num(c("forward_pf_eff"))
    win = num(c("forward_winrate"))
    trd = num(c("forward_trades"))
    dd = num("MaxDD") if "MaxDD" in df.columns else pd.Series(0.0, index=df.index)
    score = pf * (win ** 1.3) * np.log1p(trd) / (1.0 + dd / 1000.0)
    pool["_score"] = score

    # 銘柄毎に1件
    key = "Ticker" if "Ticker" in pool.columns else c("code")
    pool.sort_values([key, "_score"], ascending=[True, False], inplace=True)
    pool = pool.drop_duplicates(subset=[key], keep="first")

    # TopK（上限）
    pool.sort_values(["_score"], ascending=False, inplace=True)
    return pool.head(int(target_top)).drop(columns=["_score"], errors="ignore")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", required=True)
    ap.add_argument("--root", default="output/excel")
    ap.add_argument("--target-top", type=int, default=50)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()

    root = Path(args.root)
    frames = load_plan_candidates(root, args.date)
    if not frames:
        print(json.dumps({"written": None, "rows": 0, "message": "no plan candidates"}))
        return

    df = pd.concat(frames, ignore_index=True)
    out = args.output or (root / f"weekly_candidates_{args.date}.csv")
    out.parent.mkdir(parents=True, exist_ok=True)
    selected = select_weekly_candidates(df, args.target_top)
    if selected is None or selected.empty:
        selected = pd.DataFrame()
    selected.to_csv(out, index=False, encoding="utf-8-sig")
    print(json.dumps({"written": str(out), "rows": int(len(selected))}))


if __name__ == "__main__":
    main()
