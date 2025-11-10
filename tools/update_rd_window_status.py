#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path
from typing import Iterable

import pandas as pd


def load_summary(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(path)
    except Exception:
        return pd.DataFrame()


def resolve_column(df: pd.DataFrame, names: Iterable[str]) -> str | None:
    cols = {c.lower(): c for c in df.columns}
    for name in names:
        col = cols.get(name.lower())
        if col:
            return col
    return None


def build_status(df: pd.DataFrame, sessions: list[str]) -> pd.DataFrame:
    if df.empty:
        timestamp = dt.datetime.now(dt.timezone.utc).isoformat()
        return pd.DataFrame(
            {
                "session": sessions,
                "signal_mode": ["" for _ in sessions],
                "side": ["" for _ in sessions],
                "count": [0 for _ in sessions],
                "win_rate": [0.0 for _ in sessions],
                "mean_bp": [0.0 for _ in sessions],
                "last_updated": [timestamp for _ in sessions],
            }
        )

    session_col = resolve_column(df, ("session",))
    mode_col = resolve_column(df, ("signal_mode",))
    side_col = resolve_column(df, ("side",))
    count_col = resolve_column(df, ("forward_trades", "count"))
    win_col = resolve_column(df, ("forward_winrate", "win_rate"))
    mean_col = resolve_column(df, ("expected_bp", "mean"))

    if not session_col or not mode_col or not side_col:
        return pd.DataFrame()

    out_rows = []
    timestamp = dt.datetime.now(dt.timezone.utc).isoformat()
    for sess in sessions:
        mask = df[session_col].astype(str).str.contains(sess, case=False, na=False)
        sub = df[mask]
        if sub.empty:
            out_rows.append(
                {
                    "session": sess,
                    "signal_mode": "",
                    "side": "",
                    "count": 0,
                    "win_rate": 0.0,
                    "mean_bp": 0.0,
                    "last_updated": timestamp,
                }
            )
            continue

        sub = sub.copy()
        for col in (count_col, win_col, mean_col):
            if col:
                sub[col] = pd.to_numeric(sub[col], errors="coerce")
        sub = sub.dropna(subset=[c for c in (count_col, win_col, mean_col) if c])
        if sub.empty:
            out_rows.append(
                {
                    "session": sess,
                    "signal_mode": "",
                    "side": "",
                    "count": 0,
                    "win_rate": 0.0,
                    "mean_bp": 0.0,
                    "last_updated": timestamp,
                }
            )
            continue

        best = sub.sort_values([win_col, mean_col], ascending=False).iloc[0]
        out_rows.append(
            {
                "session": best[session_col],
                "signal_mode": best[mode_col],
                "side": best[side_col],
                "count": int(best[count_col]) if count_col else 0,
                "win_rate": float(best[win_col]) if win_col else 0.0,
                "mean_bp": float(best[mean_col]) if mean_col else 0.0,
                "last_updated": timestamp,
            }
        )
    return pd.DataFrame(out_rows)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--summary", default="analysis/session_mode_summary.csv")
    parser.add_argument("--output", default="analysis/rd_window_status.csv")
    parser.add_argument("--sessions", default="MID1030,PM1230")
    args = parser.parse_args()

    sessions = [s.strip() for s in args.sessions.split(",") if s.strip()]
    df = load_summary(Path(args.summary))
    status = build_status(df, sessions)
    status.to_csv(Path(args.output), index=False)


if __name__ == "__main__":
    main()
