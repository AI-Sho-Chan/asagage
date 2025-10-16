# -*- coding: utf-8 -*-
import os, glob, pandas as pd

BASE = r"C:\AI\asagake"
LIVE_DIR = os.path.join(BASE, "data", "live")
SIG_DIR  = os.path.join(BASE, "data", "signals", pd.Timestamp.now(tz="Asia/Tokyo").strftime("%Y-%m-%d"))

def latest_snapshot():
    pats = glob.glob(os.path.join(LIVE_DIR, "rss_snapshot_*.csv"))
    return max(pats, key=os.path.getmtime) if pats else None

def main():
    snap = latest_snapshot()
    if not snap:
        print("NO SNAPSHOT"); return
    top = os.path.join(SIG_DIR, "top_candidates.csv")
    if not os.path.exists(top):
        print("NO TOP_CANDIDATES:", top); return

    s = pd.read_csv(snap)   # ts,code,px,vwap
    t = pd.read_csv(top)

    if "Ticker" not in t.columns:
        print("TOP has no Ticker"); return

    s["Ticker"] = s["code"].astype(str) + ".T"
    u = t.merge(s[["Ticker","px","vwap","ts"]], on="Ticker", how="left")

    m = u["px"].notna() & (u["px"] > 0) & u["vwap"].notna() & (u["vwap"] > 0)

    u.loc[m, "Close"] = u.loc[m, "px"]
    u.loc[m, "VWAPd"] = u.loc[m, "vwap"]
    if "ATR5" in u.columns:
        atr = u["ATR5"].replace({0: pd.NA})
        u.loc[m, "J"] = (u.loc[m, "Close"] - u.loc[m, "VWAPd"]) / atr
        u.loc[m, "side"] = u.loc[m, "J"].apply(lambda x: "BUY" if pd.notna(x) and x < 0 else ("SELL" if pd.notna(x) else None))
    # ts は取得できた行だけ更新
    u.loc[m, "ts"] = pd.Timestamp.now(tz="Asia/Tokyo").strftime("%Y-%m-%d %H:%M:%S")

    u.drop(columns=["px","vwap"], errors="ignore", inplace=True)
    u.to_csv(top, index=False, encoding="utf-8-sig")
    print("UPDATED:", top, "rows:", len(u))

if __name__ == "__main__":
    main()
