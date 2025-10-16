# -*- coding: utf-8 -*-
"""
学習用データ生成（ts列を必ず出力）
"""
import os, glob, numpy as np, pandas as pd

BASE    = r"C:\AI\asagake"
RAWROOT = os.path.join(BASE, "data", "raw")
OUTROOT = os.path.join(BASE, "data", "ml")

INTERVALS = ["1m","5m","60m","1d"]
H=30; LAMBDA=0.7; THETA_J=0.6

def flatten_cols(df):
    if isinstance(df.columns, pd.MultiIndex):
        df=df.copy(); df.columns=["_".join([str(x) for x in t if x is not None]) for t in df.columns]
    return df

def norm_ohlcv(df):
    df=flatten_cols(df)
    key={c:"".join(str(c).lower().split()) for c in df.columns}
    def pick(*names):
        names=[n.lower() for n in names]
        for c,k in key.items():
            if any(n in k for n in names): return c
        for c in df.columns:
            if any(str(c).lower().startswith(n) for n in names): return c
        return None
    m={"Open":pick("open"),"High":pick("high"),"Low":pick("low"),
       "Close":pick("adjclose","close"),"Volume":pick("volume","vol")}
    miss=[k for k,v in m.items() if v is None]
    if miss: raise AssertionError(f"cannot map OHLCV: {miss}")
    out=df[[m["Open"],m["High"],m["Low"],m["Close"],m["Volume"]]].copy()
    out.columns=["Open","High","Low","Close","Volume"]
    return out

def ensure_jst(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.index, pd.MultiIndex):
        df=df.reset_index()
    if not isinstance(df.index, pd.DatetimeIndex):
        for c in ["Datetime","Date","Time","ts","index"]:
            if c in df.columns:
                df=df.set_index(pd.to_datetime(df[c], utc=True, errors="coerce")).drop(columns=[c]); break
        if not isinstance(df.index, pd.DatetimeIndex):
            df.index=pd.to_datetime(df.index, utc=True, errors="coerce")
    if df.index.tz is None: df.index=df.index.tz_localize("UTC")
    df=df.tz_convert("Asia/Tokyo")
    return df[~df.index.isna()]

def vwap_intraday(df):
    g=df.index.date
    pv=(df["Close"]*df["Volume"]).groupby(g).cumsum()
    vv=df["Volume"].groupby(g).cumsum().replace(0, np.nan)
    return pv/vv

def atr_n(df,n=5):
    pc=df["Close"].shift(1)
    tr=pd.concat([(df["High"]-df["Low"]).abs(),
                  (df["High"]-pc).abs(),
                  (df["Low"]-pc).abs()],axis=1).max(axis=1)
    return tr.rolling(n,min_periods=n).mean()

def make_features(df):
    df["VWAPd"]=vwap_intraday(df)
    df["ATR5"]=atr_n(df,5)
    df["J"]=(df["Close"]-df["VWAPd"])/df["ATR5"]
    df["dJ"]=df["J"].diff(); df["vEMA"]=df["dJ"].ewm(span=5,adjust=False).mean(); df["d2J"]=df["dJ"].diff()
    rng=(df["High"]-df["Low"]).replace(0,np.nan)
    df["IBS"]=(df["Close"]-df["Low"])/rng
    df["SMA20"]=df["Close"].rolling(20).mean(); df["STD20"]=df["Close"].rolling(20).std()
    df["Z20"]=(df["Close"]-df["SMA20"])/df["STD20"]
    df["ROC5"]=df["Close"].pct_change(5); df["Turnover"]=df["Close"]*df["Volume"]
    return df

def label_events(df):
    ent=df[(df["J"].abs()>=THETA_J)].copy()
    rows=[]
    for i,row in ent.iterrows():
        J0=row["J"]; px0=row["Close"]
        fwd=df.loc[i:].iloc[1:H+1]
        y=0
        for _,r in fwd.iterrows():
            if abs(r["J"])<=LAMBDA*abs(J0): y=1; break
        ret=(fwd["Close"].iloc[-1]-px0)*(+1 if J0<0 else -1) if len(fwd)>0 else np.nan
        rows.append([i,J0,px0,ret,y])
    if not rows:
        return pd.DataFrame(columns=["ts","J0","entry","ret","y"]).set_index("ts")
    return pd.DataFrame(rows,columns=["ts","J0","entry","ret","y"]).set_index("ts")

def run_interval(interval, out_dir):
    root=os.path.join(RAWROOT,f"yahoo_{interval}")
    if not os.path.isdir(root): return 0
    os.makedirs(out_dir, exist_ok=True)
    saved=0
    for tkr in os.listdir(root):
        tdir=os.path.join(root,tkr)
        if not os.path.isdir(tdir): continue
        parts=[]
        for f in os.listdir(tdir):
            if not f.endswith(".parquet"): continue
            try:
                df=pd.read_parquet(os.path.join(tdir,f))
                df=norm_ohlcv(df); df=ensure_jst(df)
                df=make_features(df).dropna(subset=["VWAPd","ATR5","J"])
                if df.empty: continue
                df["Ticker"]=tkr; parts.append(df)
            except Exception:
                continue
        if not parts: continue
        X=pd.concat(parts).sort_index()
        Y=label_events(X)
        if Y.empty: continue
        Z=X.join(Y, how="inner")              # ts=index
        Z=Z.reset_index().rename(columns={"index":"ts"})  # ← 常に ts で出力
        Z.to_parquet(os.path.join(out_dir,f"{tkr}.parquet"), index=False)
        saved+=1
        print(interval, tkr, len(Z))
    return saved

def main():
    run=pd.Timestamp.now().strftime("RUN_%Y%m%d_%H%M")
    out=os.path.join(OUTROOT, run); os.makedirs(out, exist_ok=True)
    total=0
    for itv in INTERVALS:
        cnt=run_interval(itv, os.path.join(out,f"ds_{itv}")); total+=cnt
    print("OUT:", out, "files:", total)

if __name__=="__main__": main()
