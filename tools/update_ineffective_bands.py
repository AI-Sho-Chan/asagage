from __future__ import annotations

import argparse
import datetime as dt
import json
import math
from pathlib import Path
from typing import Any, Dict, List, Tuple

import pandas as pd


def load_candidates(path: Path, trade_threshold: int, j_max: float, exp_threshold: float) -> List[float]:
    df = pd.read_csv(path)
    df = df[pd.to_numeric(df.get("trades"), errors="coerce") >= trade_threshold]
    df = df[pd.to_numeric(df.get("J_th"), errors="coerce") <= j_max]
    df = df[pd.to_numeric(df.get("exp_mean"), errors="coerce") <= exp_threshold]
    j_vals = sorted({round(float(v), 4) for v in df["J_th"].tolist()})
    return j_vals


def ensure_entry(j_map: Dict[str, Dict[str, Any]], j_value: float) -> Dict[str, Any]:
    key = f"{j_value:.4f}"
    entry = j_map.get(key)
    if entry is None:
        entry = {
            "history": [],
            "avg_pf": None,
            "masked": False,
        }
        j_map[key] = entry
    return entry


def load_state(state_path: Path) -> Dict[str, Any]:
    if state_path.exists():
        try:
            with state_path.open("r", encoding="utf-8") as fh:
                data = json.load(fh)
                if isinstance(data, dict):
                    return data
        except Exception:
            pass
    return {"version": 1, "plans": {}}


def store_state(state_path: Path, state: Dict[str, Any]) -> None:
    state_path.parent.mkdir(parents=True, exist_ok=True)
    with state_path.open("w", encoding="utf-8") as fh:
        json.dump(state, fh, ensure_ascii=False, indent=2, sort_keys=True)


def mask_j_values(state: Dict[str, Any], j_values: List[float], note: str, ts: str) -> None:
    plans = state.get("plans", {})
    for plan_name, plan_data in plans.items():
        j_map = plan_data.setdefault("J_th", {})
        for j in j_values:
            entry = ensure_entry(j_map, j)
            if not entry.get("masked"):
                entry.setdefault("history", []).append(
                    {
                        "ts": ts,
                        "forward_pf_eff": -1.0,
                        "auto_mask": True,
                        "source": note,
                    }
                )
            entry["masked"] = True


def _recent_positive_pf(history: List[Dict[str, Any]], window: int) -> List[float]:
    vals: List[float] = []
    for rec in reversed(history or []):
        val = rec.get("forward_pf_eff")
        try:
            num = float(val)
        except (TypeError, ValueError):
            continue
        if not math.isfinite(num) or num <= 0:
            continue
        vals.append(num)
        if len(vals) >= window:
            break
    return vals


def apply_unmask(
    state: Dict[str, Any],
    window: int,
    threshold: float,
    min_count: int,
    ts: str,
) -> List[Tuple[str, str]]:
    """Return list of (plan, J_key) that were unmasked."""
    unmasked: List[Tuple[str, str]] = []
    plans = state.get("plans", {})
    for plan_name, plan_data in plans.items():
        j_map: Dict[str, Any] = plan_data.get("J_th", {})
        for j_key, entry in j_map.items():
            if not entry.get("masked"):
                continue
            history = entry.get("history", [])
            recent = _recent_positive_pf(history, window)
            if len(recent) < min_count:
                continue
            avg_pf = sum(recent) / len(recent)
            count_ge = sum(1 for v in recent if v >= threshold)
            if avg_pf >= threshold and count_ge >= min_count:
                entry["masked"] = False
                entry.setdefault("history", []).append(
                    {
                        "ts": ts,
                        "forward_pf_eff": avg_pf,
                        "auto_unmask": True,
                        "window": window,
                        "threshold": threshold,
                    }
                )
                unmasked.append((plan_name, j_key))
    return unmasked


def resolve_by_j(path_run: Path) -> Path:
    return path_run / "by_J_th.csv"


def main() -> None:
    ap = argparse.ArgumentParser(description="Auto-update ineffective J bands based on report")
    ap.add_argument("--by-j", help="Path to by_J_th.csv produced by analyze_param_stats")
    ap.add_argument("--run-root", help="reports/param_stats/NIGHTLY_xxx directory")
    ap.add_argument("--state", default="state/ineffective_bands.json", help="Path to state JSON")
    ap.add_argument("--trade-threshold", type=int, default=5000)
    ap.add_argument("--j-max", type=float, default=0.8)
    ap.add_argument("--exp-threshold", type=float, default=0.0, help="Upper bound of exp_mean to mask")
    ap.add_argument("--allow-unmask", action="store_true", help="If set, evaluate unmask candidates")
    ap.add_argument("--unmask-window", type=int, default=5)
    ap.add_argument("--unmask-threshold", type=float, default=1.1)
    ap.add_argument("--unmask-min-count", type=int, default=3)
    args = ap.parse_args()

    if not args.by_j and not args.run_root:
        raise SystemExit("Either --by-j or --run-root must be specified")

    if args.by_j:
        by_j_path = Path(args.by_j)
    else:
        by_j_path = resolve_by_j(Path(args.run_root))

    if not by_j_path.exists():
        raise SystemExit(f"by_J_th file not found: {by_j_path}")

    state_path = Path(args.state)
    state = load_state(state_path)
    ts = dt.datetime.now().isoformat()

    j_candidates = load_candidates(by_j_path, args.trade_threshold, args.j_max, args.exp_threshold)
    if j_candidates:
        mask_j_values(state, j_candidates, note="auto_mask_j_le_threshold", ts=ts)
        print("Updated ineffective_bands with", j_candidates)
    else:
        print("No J values qualified for masking")

    unmasked = []
    if args.allow_unmask:
        unmasked = apply_unmask(
            state,
            window=max(1, args.unmask_window),
            threshold=args.unmask_threshold,
            min_count=max(1, args.unmask_min_count),
            ts=ts,
        )
        if unmasked:
            print("Unmasked bands", unmasked)

    if j_candidates or unmasked or args.allow_unmask:
        store_state(state_path, state)


if __name__ == "__main__":
    main()
