#!/usr/bin/env python3
"""
候補銘柄ごとに「1分足データが何日分そろっているか」をざっくり点検するスクリプト。

用途（分かりやすく）:
- 週末バッチが「過去データ不足」で止まりやすいかどうかを早めに検知する
- 「今の lookback（日数）を伸ばして良いか」を判断する材料にする

使い方:
  python tools/report_history_adequacy.py --candidates output/excel/candidates_nextday.csv --lookback 90

出力:
- 画面にサマリ表示
- `analysis/history_adequacy_YYYYMMDD.json` を保存
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd


def estimate_history_days(codes: List[str], root: Path) -> Dict[str, int]:
    days_map: Dict[str, int] = {}
    for code in codes:
        directory = root / code
        if not directory.exists():
            days_map[code] = 0
            continue
        cnt = 0
        for fp in directory.glob("*.parquet"):
            try:
                dt.datetime.strptime(fp.stem, "%Y-%m-%d")
            except ValueError:
                continue
            cnt += 1
        days_map[code] = cnt
    return days_map


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidates", required=True, help="candidates_nextday.csv または NIGHTLY_xxx candidates ファイル")
    ap.add_argument("--minute-root", default="data/raw/yahoo_1m", help="1分足parquetの保存先ルート")
    ap.add_argument("--lookback", type=int, default=90, help="現在の lookback（日数）の目安")
    args = ap.parse_args()

    cand_path = Path(args.candidates)
    if not cand_path.exists():
        raise SystemExit(f"candidates file not found: {cand_path}")

    df = pd.read_csv(cand_path)
    if df.empty or "Ticker" not in df.columns:
        raise SystemExit("candidates file is empty or lacks Ticker column")

    tickers = df["Ticker"].dropna().astype(str).str.upper().tolist()
    uniq = sorted(dict.fromkeys(tickers))

    root = Path(args.minute_root)
    hist_days = estimate_history_days(uniq, root)
    df["TickerUpper"] = df["Ticker"].astype(str).str.upper()
    df["HistoryDays"] = df["TickerUpper"].map(hist_days).fillna(0).astype(int)

    train_trades = pd.to_numeric(df.get("train_trades", 0), errors="coerce").fillna(0)
    forward_trades = pd.to_numeric(df.get("forward_trades", 0), errors="coerce").fillna(0)
    df["train_trades_num"] = train_trades
    df["forward_trades_num"] = forward_trades

    summary: Dict[str, object] = {}

    hvals = df["HistoryDays"].to_numpy()
    summary["history_min"] = int(hvals.min()) if hvals.size else 0
    summary["history_median"] = float(np.median(hvals)) if hvals.size else 0.0
    summary["history_max"] = int(hvals.max()) if hvals.size else 0

    lb = int(args.lookback)
    summary["lookback_days"] = lb
    summary["pct_lt_half_lb"] = float((df["HistoryDays"] < lb * 0.5).mean()) if hvals.size else 0.0
    summary["pct_lt_lb"] = float((df["HistoryDays"] < lb).mean()) if hvals.size else 0.0
    summary["pct_ge_2x_lb"] = float((df["HistoryDays"] >= lb * 2).mean()) if hvals.size else 0.0

    tvals = df["train_trades_num"].to_numpy()
    fvals = df["forward_trades_num"].to_numpy()
    summary["train_trades_median"] = float(np.median(tvals)) if tvals.size else 0.0
    summary["forward_trades_median"] = float(np.median(fvals)) if fvals.size else 0.0

    messages: List[str] = []
    if float(summary["pct_lt_lb"]) >= 0.5:
        messages.append(
            "候補の半分以上が lookback（日数）より短い履歴しか持っていません。"
            "今より lookback を伸ばす前に、まず minute_cache の深掘り頻度を増やすことを検討してください。"
        )
    elif float(summary["pct_ge_2x_lb"]) >= 0.5:
        messages.append(
            "候補の半分以上が lookback（日数）の2倍以上の履歴を持っています。"
            "さらに長い期間での検証（lookback を伸ばす）も検討できます。"
        )
    else:
        messages.append(
            "現状の lookback（日数）に対して極端に不足している銘柄は多くなさそうです。"
            "ただし週末バッチが止まる場合は「その銘柄の履歴不足」が原因のことが多いので、"
            "Top200常連などの1分足保存を継続するのがおすすめです。"
        )
    summary["messages"] = messages

    out_dir = Path("analysis")
    out_dir.mkdir(parents=True, exist_ok=True)
    tag = dt.date.today().strftime("%Y%m%d")
    out_path = out_dir / f"history_adequacy_{tag}.json"
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"ASAGAKE history adequacy report ({dt.datetime.now():%Y-%m-%d %H:%M:%S})")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"[history_adequacy] written: {out_path}")


if __name__ == "__main__":
    main()

