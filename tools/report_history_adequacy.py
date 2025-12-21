#!/usr/bin/env python3
"""
ASAGAKE: 1分足データの「履歴が足りているか」簡易チェック

候補CSVに含まれる銘柄（Ticker）について、ローカル保存している Yahoo 1分足の
「営業日ファイル数」を数えます（例: data/raw/yahoo_1m/7203.T/2025-12-12.parquet）。

目的:
- 週末バッチの lookback（例: 90営業日）を伸ばしても大丈夫か
- minute_cache（ローカル保存）が十分に育っているか
を、数字でざっくり把握する。

使い方:
  python tools/report_history_adequacy.py --candidates output/excel/candidates_nextday.csv --lookback 90

出力:
- 標準出力に集計結果(JSON)を表示
- analysis/history_adequacy_YYYYMMDD.json を生成
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd


def estimate_history_days(tickers: List[str], root: Path) -> Dict[str, int]:
    days_map: Dict[str, int] = {}
    for ticker in tickers:
        directory = root / ticker
        if not directory.exists():
            days_map[ticker] = 0
            continue
        count = 0
        for parquet_path in directory.glob("*.parquet"):
            try:
                dt.datetime.strptime(parquet_path.stem, "%Y-%m-%d")
            except ValueError:
                continue
            count += 1
        days_map[ticker] = count
    return days_map


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--candidates", required=True, help="候補CSV（例: output/excel/candidates_nextday.csv）")
    parser.add_argument(
        "--minute-root",
        default="data/raw/yahoo_1m",
        help="Yahoo 1分足の保存先（ティッカー/日付.parquet）",
    )
    parser.add_argument("--lookback", type=int, default=90, help="「これくらい欲しい」履歴の目安（営業日）")
    args = parser.parse_args()

    candidates_path = Path(args.candidates)
    if not candidates_path.exists():
        raise SystemExit(f"candidates file not found: {candidates_path}")

    df = pd.read_csv(candidates_path)
    if df.empty or "Ticker" not in df.columns:
        raise SystemExit("candidates file is empty or lacks Ticker column")

    tickers = df["Ticker"].dropna().astype(str).str.upper().tolist()
    uniq_tickers = sorted(dict.fromkeys(tickers))

    root = Path(args.minute_root)
    hist_days = estimate_history_days(uniq_tickers, root)

    df["TickerUpper"] = df["Ticker"].astype(str).str.upper()
    df["HistoryDays"] = df["TickerUpper"].map(hist_days).fillna(0).astype(int)

    train_trades = pd.to_numeric(df.get("train_trades", 0), errors="coerce").fillna(0)
    forward_trades = pd.to_numeric(df.get("forward_trades", 0), errors="coerce").fillna(0)

    history_vals = df["HistoryDays"].to_numpy()
    train_vals = train_trades.to_numpy()
    forward_vals = forward_trades.to_numpy()

    lookback_days = int(args.lookback)

    summary: Dict[str, object] = {
        "history_min": int(history_vals.min()) if history_vals.size else 0,
        "history_median": float(np.median(history_vals)) if history_vals.size else 0.0,
        "history_max": int(history_vals.max()) if history_vals.size else 0,
        "lookback_days": lookback_days,
        "pct_lt_half_lb": float((df["HistoryDays"] < lookback_days * 0.5).mean()) if history_vals.size else 0.0,
        "pct_lt_lb": float((df["HistoryDays"] < lookback_days).mean()) if history_vals.size else 0.0,
        "pct_ge_2x_lb": float((df["HistoryDays"] >= lookback_days * 2).mean()) if history_vals.size else 0.0,
        "train_trades_median": float(np.median(train_vals)) if train_vals.size else 0.0,
        "forward_trades_median": float(np.median(forward_vals)) if forward_vals.size else 0.0,
    }

    messages: List[str] = []
    pct_lt_lb = float(summary["pct_lt_lb"])
    pct_lt_half = float(summary["pct_lt_half_lb"])
    pct_ge_2x = float(summary["pct_ge_2x_lb"])

    if pct_lt_lb >= 0.5:
        messages.append(
            "候補の半分以上が lookback 日数より短い履歴しか持っていません。"
            "今より lookback を伸ばす前に、minute_cache の深掘り（過去分の取得）頻度を増やすことを検討してください。"
        )
    if pct_lt_half >= 0.2:
        messages.append(
            "候補の2割以上が lookback の半分未満しか履歴がありません。"
            "一部銘柄はテスト期間が短くなり、結果がブレやすくなる可能性があります。"
        )
    if pct_ge_2x >= 0.5:
        messages.append(
            "候補の半分以上が lookback の2倍以上の履歴を持っています。"
            "lookback を伸ばす（例: 90→120）余地がありますが、計算時間とのバランスを見て段階的に行うのがおすすめです。"
        )
    if not messages:
        messages.append(
            "lookback に対して致命的な不足は見えません。週末バッチを継続しつつ、minute_cache を少しずつ育てていく運用でOKです。"
        )

    summary["messages"] = messages

    out_dir = Path("analysis")
    out_dir.mkdir(parents=True, exist_ok=True)
    tag = dt.date.today().strftime("%Y%m%d")
    out_path = out_dir / f"history_adequacy_{tag}.json"
    out_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    # Human-friendly mail body (Japanese, UTF-8).
    mail_lines: List[str] = []
    mail_lines.append(f"ASAGAKE 履歴の十分性チェック（JST） {dt.datetime.now():%Y-%m-%d %H:%M:%S}")
    mail_lines.append("")
    mail_lines.append("【対象】")
    mail_lines.append(f"- 候補CSV: {candidates_path.as_posix()}")
    mail_lines.append(f"- 保存先（minute_cache）: {root.as_posix()}")
    mail_lines.append("")
    mail_lines.append("【結論（短く）】")
    for m in messages:
        mail_lines.append(f"- {m}")
    mail_lines.append("")
    mail_lines.append("【数字】")
    mail_lines.append(f"- lookback（目安）: {lookback_days} 日")
    mail_lines.append(
        f"- 保存済み1分足（営業日数）: 最小 {summary['history_min']} / 中央 {summary['history_median']} / 最大 {summary['history_max']} 日"
    )
    mail_lines.append(f"- lookback 未満: {summary['pct_lt_lb']:.3f}（= {summary['pct_lt_lb']*100:.1f}%）")
    mail_lines.append(f"- lookback の半分未満: {summary['pct_lt_half_lb']:.3f}（= {summary['pct_lt_half_lb']*100:.1f}%）")
    mail_lines.append(f"- lookback の2倍以上: {summary['pct_ge_2x_lb']:.3f}（= {summary['pct_ge_2x_lb']*100:.1f}%）")
    mail_lines.append("")
    mail_lines.append("【補足】")
    mail_lines.append("- このチェックは「銘柄ごとの保存日数」を見るだけです（売買結果の良し悪しではありません）。")
    mail_lines.append("- lookback を伸ばすほどテストは安定しやすい一方、計算時間は増えます。")
    mail_lines.append("")
    mail_lines.append("（添付のJSONは生データです。必要なときだけ参照してください）")

    mail_path = out_dir / f"history_adequacy_{tag}_mail.txt"
    mail_path.write_text("\n".join(mail_lines) + "\n", encoding="utf-8")

    now = dt.datetime.now()
    print(f"ASAGAKE history adequacy report ({now:%Y-%m-%d %H:%M:%S} JST)")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"[history_adequacy] written: {out_path}")
    print(f"[history_adequacy] mail: {mail_path}")


if __name__ == "__main__":
    main()
