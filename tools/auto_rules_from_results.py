#!/usr/bin/env python3
from __future__ import annotations

import argparse
from collections import OrderedDict
from pathlib import Path
from typing import Iterable

import pandas as pd


def load_session_mode_summary(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()


def _column_lookup(df: pd.DataFrame, candidates: Iterable[str]) -> str | None:
    cols = {c.lower(): c for c in df.columns}
    for name in candidates:
        col = cols.get(name.lower())
        if col:
            return col
    return None


def decide_am1000_sell_enable(df: pd.DataFrame) -> bool:
    """Return True when AM1000×j-cross×SELL shows strong, stable stats."""
    if df.empty:
        return False

    session_col = _column_lookup(df, ("session",))
    mode_col = _column_lookup(df, ("signal_mode",))
    side_col = _column_lookup(df, ("side",))
    if not session_col or not mode_col or not side_col:
        return False

    sub = df.copy()
    for col in (session_col, mode_col, side_col):
        sub[col] = sub[col].astype(str)

    mask = (
        sub[session_col].str.contains("AM1000", case=False, na=False)
        & sub[mode_col].str.contains("j-cross", case=False, na=False)
        & (sub[side_col].str.upper() == "SELL")
    )
    sub = sub[mask]
    if sub.empty:
        return False

    pf_col = _column_lookup(sub, ("forward_pf_eff", "forward_pf", "pf"))
    win_col = _column_lookup(sub, ("forward_winrate", "win_rate"))
    bp_col = _column_lookup(sub, ("expected_bp", "mean"))
    trades_col = _column_lookup(sub, ("forward_trades", "count"))
    if not win_col or not bp_col or not trades_col:
        return False

    numeric_cols = [win_col, bp_col, trades_col]
    if pf_col:
        numeric_cols.append(pf_col)
    sub[numeric_cols] = sub[numeric_cols].apply(pd.to_numeric, errors="coerce")
    sub = sub.dropna(subset=[win_col, bp_col, trades_col])
    if pf_col:
        sub = sub.dropna(subset=[pf_col])
    if sub.empty:
        return False

    sort_cols = [c for c in (pf_col, win_col, bp_col) if c]
    row = sub.sort_values(sort_cols, ascending=False).iloc[0]

    pf = float(row[pf_col]) if pf_col else None
    win = float(row[win_col])
    bp = float(row[bp_col])
    trades = float(row[trades_col])
    pf_ok = True if pf is None else pf >= 1.50
    return pf_ok and win >= 0.65 and bp >= 8.0 and trades >= 8


def update_rules(rules_path: Path, *, enable_am1000_sell: bool) -> None:
    content: "OrderedDict[str, str]" = OrderedDict()
    if rules_path.exists():
        for line in rules_path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            content[key.strip()] = value.strip()

    content["am1000_sell_enabled"] = "1" if enable_am1000_sell else "0"
    content.setdefault("jcross_sell_require_nky_down", "1")
    content.setdefault("jcross_sell_min_gap_bp", "20")

    rules_path.write_text(
        "\n".join(f"{key}={value}" for key, value in content.items()) + "\n",
        encoding="utf-8",
    )


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--summary", default="analysis/session_mode_summary.csv")
    parser.add_argument("--rules", default="state/strategy_rules.ini")
    args = parser.parse_args()

    df = load_session_mode_summary(Path(args.summary))
    flag = decide_am1000_sell_enable(df)
    update_rules(Path(args.rules), enable_am1000_sell=flag)
    print({"am1000_sell_enabled": int(flag)})


if __name__ == "__main__":
    main()
