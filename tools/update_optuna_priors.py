from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


def load_grids(run_root: Path) -> pd.DataFrame:
    files = list(run_root.rglob("_GRID_FULL.csv"))
    frames: List[pd.DataFrame] = []
    for f in files:
        try:
            df = pd.read_csv(f)
            frames.append(df)
        except Exception:
            pass
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def ensure_numeric(df: pd.DataFrame, cols: List[str]) -> None:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")


def aggregate_candidates(df: pd.DataFrame, trades_min: int, win_min: float, exp_min: float, top_k: int) -> List[Dict[str, Any]]:
    keep = ["ATR_n", "TPk", "SLk", "J_th", "forward_trades", "forward_winrate", "forward_exp_boot_mean", "forward_pf_eff"]
    miss = [c for c in keep if c not in df.columns]
    if miss:
        return []
    sub = df[keep].copy()
    ensure_numeric(sub, ["ATR_n", "TPk", "SLk", "J_th", "forward_trades", "forward_winrate", "forward_exp_boot_mean", "forward_pf_eff"])
    sub = sub.dropna()
    grp = (
        sub.groupby(["ATR_n", "TPk", "SLk", "J_th"], as_index=False)
        .agg(
            trades=("forward_trades", "sum"),
            win=("forward_winrate", "mean"),
            exp=("forward_exp_boot_mean", "mean"),
            pf=("forward_pf_eff", "mean"),
        )
    )
    # filters
    grp = grp[(grp["trades"] >= trades_min) & (grp["exp"] >= exp_min) & (grp["win"] >= win_min)]
    if grp.empty:
        return []
    # score: prioritize expected return, then win, then pf
    grp["score"] = grp["exp"] * 1.0 + grp["win"] * 1.0 + (grp["pf"].clip(upper=300.0) / 300.0) * 0.5
    grp = grp.sort_values(["score", "exp", "win"], ascending=False).head(top_k)
    out: List[Dict[str, Any]] = []
    for _, r in grp.iterrows():
        out.append(
            {
                "ATR_n": int(r["ATR_n"]),
                "TPk": float(r["TPk"]),
                "SLk": float(r["SLk"]),
                "J_th": float(r["J_th"]),
                "trades": int(r["trades"]),
                "win": float(r["win"]),
                "exp": float(r["exp"]),
                "pf": float(r["pf"]),
            }
        )
    return out


def load_state(path: Path) -> Dict[str, Any]:
    if path.exists():
        try:
            with path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {"version": 1, "epoch": "", "ttl_days": 7, "priors": []}


def store_state(path: Path, state: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2, sort_keys=True)


def merge_priors(state: Dict[str, Any], new_priors: List[Dict[str, Any]], source: str, epoch: str, replace_source: bool) -> Dict[str, Any]:
    now = dt.datetime.now().isoformat()
    cur = list(state.get("priors", []))
    if replace_source:
        cur = [p for p in cur if p.get("source") != source]
    # add new with metadata
    for p in new_priors:
        p2 = dict(p)
        p2.update({"source": source, "ts": now, "epoch": epoch, "weight": p.get("exp", 0.0)})
        cur.append(p2)
    state["priors"] = cur
    if source == "weekend":
        state["epoch"] = epoch
    return state


def main() -> None:
    ap = argparse.ArgumentParser(description="Update Optuna priors from NIGHTLY run")
    ap.add_argument("--run-root", required=True, help="Path to output/bt30/NIGHTLY_YYYYMMDD")
    ap.add_argument("--state", default="state/optuna_priors.json")
    ap.add_argument("--source", choices=["weekend", "weekday"], required=True)
    ap.add_argument("--trades-min", type=int, default=500)
    ap.add_argument("--win-min", type=float, default=0.52)
    ap.add_argument("--exp-min", type=float, default=0.0)
    ap.add_argument("--top-k", type=int, default=24)
    args = ap.parse_args()

    run_root = Path(args.run_root)
    df = load_grids(run_root)
    if df.empty:
        raise SystemExit("no grids found")

    priors = aggregate_candidates(df, args.trades_min, args.win_min, args.exp_min, args.top_k)
    if not priors:
        print("no priors selected")
        return

    epoch = run_root.name.replace("NIGHTLY_", "")
    state = load_state(Path(args.state))
    state = merge_priors(state, priors, source=args.source, epoch=epoch, replace_source=(args.source == "weekend"))
    store_state(Path(args.state), state)
    print(f"stored {len(priors)} priors (source={args.source}, epoch={epoch})")


if __name__ == "__main__":
    main()

