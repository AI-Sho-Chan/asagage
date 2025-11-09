#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path
import configparser
import pandas as pd


def load_session_mode_summary(path: Path) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame()
    try:
        df = pd.read_csv(path)
        return df
    except Exception:
        return pd.DataFrame()


def decide_am1000_sell_enable(df: pd.DataFrame) -> bool:
    if df.empty:
        return False
    # 期待する列例: session, signal_mode, side, forward_pf, forward_winrate, forward_trades
    cols = {c.lower(): c for c in df.columns}
    def lc(name: str) -> str:
        key = name.lower()
        return cols.get(key, name)

    if lc("session") not in df.columns:
        return False

    sub = df.copy()
    # 正規化
    for c in ("session", "signal_mode", "side"):
        if lc(c) in sub.columns:
            sub[lc(c)] = sub[lc(c)].astype(str)

    # AM1000 × j-cross × SELL のみ
    mask = (
        (sub[lc("session")].str.contains("AM1000", case=False, na=False)) &
        (lc("signal_mode") in sub.columns and sub[lc("signal_mode")].str.contains("j-cross", case=False, na=False)) &
        (lc("side") in sub.columns and sub[lc("side")].str.upper() == "SELL")
    )
    sub = sub[mask]
    if sub.empty:
        return False

    # 閾値: PF>=1.50, 勝率>=0.65, 取引数>=8（直近集計）
    pf_col = lc("forward_pf") if lc("forward_pf") in sub.columns else lc("forward_pf_eff")
    win_col = lc("forward_winrate")
    n_col = lc("forward_trades")
    if pf_col not in sub.columns or win_col not in sub.columns or n_col not in sub.columns:
        return False

    row = sub.sort_values(pf_col, ascending=False).iloc[0]
    try:
        pf = float(row[pf_col])
        win = float(row[win_col])
        n = float(row[n_col])
    except Exception:
        return False
    return (pf >= 1.50) and (win >= 0.65) and (n >= 8)


def update_rules(rules_path: Path, *, enable_am1000_sell: bool) -> None:
    rules = configparser.ConfigParser()
    rules.optionxform = str
    # INI(=key=value) をセクション無しで扱う
    content = {}
    if rules_path.exists():
        text = rules_path.read_text(encoding="utf-8")
        for line in text.splitlines():
            if not line.strip() or line.strip().startswith("#"):
                continue
            if "=" in line:
                k, v = line.split("=", 1)
                content[k.strip()] = v.strip()
    # 既存値を保ちつつ追記/更新
    content["am1000_sell_enabled"] = "1" if enable_am1000_sell else "0"
    # 既定のSELLゲートは厳しめ継続
    if "jcross_sell_require_nky_down" not in content:
        content["jcross_sell_require_nky_down"] = "1"
    if "jcross_sell_min_gap_bp" not in content:
        content["jcross_sell_min_gap_bp"] = "20"
    # 保存
    lines = [f"{k}={v}" for k, v in content.items()]
    rules_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--summary", default="analysis/session_mode_summary.csv")
    ap.add_argument("--rules", default="state/strategy_rules.ini")
    args = ap.parse_args()

    df = load_session_mode_summary(Path(args.summary))
    flag = decide_am1000_sell_enable(df)
    update_rules(Path(args.rules), enable_am1000_sell=flag)
    print({"am1000_sell_enabled": int(flag)})


if __name__ == "__main__":
    main()

