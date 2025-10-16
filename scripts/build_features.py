# -*- coding: utf-8 -*-
import os, glob, numpy as np, pandas as pd

BASE=r"C:\AI\asagake"
RAW =os.path.join(BASE,r"data\raw\yahoo_1m")
RUN =pd.Timestamp.now().strftime("RUN_%Y%m%d_%H%M")
OUT =os.path.join(BASE,r"data\proc\features_1m", RUN)
os.makedirs(OUT, exist_ok=True)

def flatten_cols(df):
    if isinstance(df.columns, pd.MultiIndex):
        df=df.copy()
        df.columns=["_".join([str(x) for x in t if x is not None]) for t in df.columns]
    return df

def normalize_ohlcv(df):
    df=flatten_cols(df)
    key={c:"".join(str(c).lower().split()) for c in df.columns}
    def pick(*names):
        names=[n.lower() for n in names]
        for c,k in key.items():
            if any(n in k for n in names): return c
        for c in df.columns:
            if any(str(c).lower().startswith(n) for n in names): return c
        return None
    m={}
    m["Open"]  = pick("open")
    m["High"]  = pick("high")
    m["Low"]   = pick("low")
    m["Close"] = pick("adjclose","close")
    m["Volume"]= pick("volume","vol")
    miss=[k for k,v in m.items() if v is None]
    if miss: raise AssertionError(f"cannot map OHLCV: missing={miss}, cols={list(df.columns)[:10]}")
    out=df[[m["Open"],m["High"],m["Low"],m["Close"],m["Volume"]]].copy()
    out.columns=["Open","High","Low","Close","Volume"]
    return out

def ensure_dt_index(df):
    if not isinstance(df.index, pd.DatetimeIndex):
        for c in ["Datetime","Date","Time","ts","index"]:
            if c in df.columns:
                df=df.set_index(pd.to_datetime(df[c], utc=True)).drop(columns=[c]); break
        if not isinstance(df.index, pd.DatetimeIndex):
            df.index=pd.to_datetime(df.index, utc=True)
    if df.index.tz is None: df.index=df.index.tz_localize("UTC")
    return df.index.tz_convert("Asia/Tokyo")

def vwap_intraday(df):
    g=df.index.date
    pv=(df["Close"]*df["Volume"]).groupby(g).cumsum()
    vv=df["Volume"].groupby(g).cumsum().replace(0,np.nan)
    return pv/vv

def atr_n(df,n=5):
    pc=df["Close"].shift(1)
    tr=pd.concat([(df["High"]-df["Low"]).abs(),
                  (df["High"]-pc).abs(),
                  (df["Low"]-pc).abs()],axis=1).max(axis=1)
    return tr.rolling(n,min_periods=n).mean()

def process_ticker(tkr):
    files=glob.glob(os.path.join(RAW,tkr,"*.parquet"))
    if not files: return 0
    parts=[]
    for fn in files:
        try:
            df=pd.read_parquet(fn).sort_index()
            df=normalize_ohlcv(df)
            _=ensure_dt_index(df)  # 参照でtz修正
            df.index=_

            df["VWAPd"]=vwap_intraday(df)
            df["ATR5"]=atr_n(df,5)
            df["J"]=(df["Close"]-df["VWAPd"])/df["ATR5"]
            df["dJ"]=df["J"].diff(); df["vEMA"]=df["dJ"].ewm(span=5,adjust=False).mean(); df["d2J"]=df["dJ"].diff()

            hhmm=df.index.strftime("%H%M").astype(int)
            df["Session"]=np.where((hhmm>=900)&(hhmm<1130),"AM",
                             np.where((hhmm>=1230)&(hhmm<1500),"PM","OFF"))

            keep=["Open","High","Low","Close","Volume","VWAPd","ATR5","J","dJ","d2J","vEMA","Session"]
            df=df[keep].dropna(subset=["VWAPd","ATR5","J"])
            if df.empty: continue
            df["Ticker"]=tkr; parts.append(df)
        except Exception:
            continue
    if not parts: return 0
    out=pd.concat(parts).sort_index()
    out.to_parquet(os.path.join(OUT,f"{tkr}.parquet"))
    return len(out)

def main():
    tks=[d for d in os.listdir(RAW) if os.path.isdir(os.path.join(RAW,d))]
    rows=0; done=0
    for t in tks:
        n=process_ticker(t)
        if n>0: done+=1; rows+=n; print(t,n)
    print("RUN:", RUN, "files:",done, "rows:",rows)
    pats=glob.glob(os.path.join(OUT,"*.parquet"))
    if pats:
        whole=pd.concat([pd.read_parquet(p) for p in pats]).sort_index()
        whole.to_parquet(os.path.join(OUT,"_ALL.parquet"))
if __name__=="__main__": main()
