# -*- coding: utf-8 -*-
import os, glob, json, numpy as np, pandas as pd, joblib

BASE=r"C:\AI\asagake"
FEAT_ROOT=os.path.join(BASE,"data","proc","features_1m")
MODEL=os.path.join(BASE,"models","reversion.pkl")
OUTP=os.path.join(BASE,"models","opt_thresholds.json")

FEATS=["J","dJ","vEMA","d2J","ATR5","IBS","Z20","ROC5","Turnover"]
BASE_NEED={"Open","High","Low","Close","Volume","Session","Ticker"}
THR_GRID=[0.6,0.7,0.8]; ABSJ_GRID=[0.8,1.0,1.2]; VALUE_GRID=[3e7,5e7,1e8]
LAM_GRID=[0.7,0.8,0.9]; KATR_GRID=[1.5,2.0]; TMAX_GRID=[45,60,90]
N_DAYS=10; COST_PER_TRADE=0.0

def latest_run_dir(root):
    runs=[p for p in glob.glob(os.path.join(root,"RUN_*")) if os.path.isdir(p)]
    if not runs: raise FileNotFoundError("no RUN_* under features_1m")
    return max(runs,key=os.path.getmtime)

def ensure_jst(df):
    if isinstance(df.index,pd.MultiIndex): df=df.reset_index()
    if not isinstance(df.index,pd.DatetimeIndex):
        for c in ["Datetime","Date","Time","ts","index"]:
            if c in df.columns:
                df=df.set_index(pd.to_datetime(df[c], utc=True, errors="coerce")).drop(columns=[c]); break
        if not isinstance(df.index,pd.DatetimeIndex):
            df.index=pd.to_datetime(df.index, utc=True, errors="coerce")
    if df.index.tz is None: df.index=df.index.tz_localize("UTC")
    return df.tz_convert("Asia/Tokyo")

def vwap_intraday(df):
    g=df.index.date
    pv=(df["Close"]*df["Volume"]).groupby(g).cumsum()
    vv=df["Volume"].groupby(g).cumsum().replace(0,np.nan)
    return pv/vv

def atr_n(df,n=5):
    pc=df["Close"].shift(1)
    tr=pd.concat([(df["High"]-df["Low"]).abs(),
                  (df["High"]-pc).abs(),
                  (df["Low"] -pc).abs()],axis=1).max(axis=1)
    return tr.rolling(n,min_periods=n).mean()

def compute_missing_columns(df):
    df=ensure_jst(df)
    if "ATR5" not in df.columns or df["ATR5"].isna().all(): df["ATR5"]=atr_n(df,5)
    if "VWAPd" not in df.columns or df["VWAPd"].isna().all(): df["VWAPd"]=vwap_intraday(df)
    if "J" not in df.columns or df["J"].isna().all(): df["J"]=(df["Close"]-df["VWAPd"])/df["ATR5"]
    for c in ("dJ","vEMA","d2J","IBS","SMA20","STD20","Z20","ROC5","Turnover"):
        if c not in df.columns:
            if c=="dJ": df[c]=df["J"].diff()
            elif c=="vEMA": df[c]=df["dJ"].ewm(span=5,adjust=False).mean()
            elif c=="d2J": df[c]=df["dJ"].diff()
            elif c=="IBS": df[c]=(df["Close"]-df["Low"])/(df["High"]-df["Low"]).replace(0,np.nan)
            elif c=="SMA20": df[c]=df["Close"].rolling(20).mean()
            elif c=="STD20": df[c]=df["Close"].rolling(20).std()
            elif c=="Z20": df[c]=(df["Close"]-df["SMA20"])/df["STD20"]
            elif c=="ROC5": df[c]=df["Close"].pct_change(5)
            elif c=="Turnover": df[c]=df["Close"]*df["Volume"]
    return df

def simulate(df, THR, ABSJ, VALUE, LAM, KATR, TMAX):
    df=df.copy()
    mask=(df["p"]>=THR)&(np.abs(df["J"])>=ABSJ)&((df["Close"]*df["Volume"])>=VALUE)
    entries=df[mask].copy()
    if entries.empty: return 0.0,0
    # ts列を作ってループ（DatetimeIndex）
    entries=entries.reset_index().rename(columns={"index":"ts"})
    pnl=0.0; n=0
    # 同一（日, Ticker）で最初の1件だけ
    entries["d"]=entries["ts"].dt.date
    entries=entries.sort_values("ts").drop_duplicates(["d","Ticker"])
    for _,row in entries.iterrows():
        ts=row["ts"]; j0=row["J"]; px0=row["Close"]; atr=row["ATR5"]; side=1 if j0<0 else -1
        fwd=df.loc[ts:].iloc[1:int(TMAX)+1]
        if fwd.empty: continue
        hit=fwd[np.abs(fwd["J"])<=LAM*abs(j0)]
        if not hit.empty:
            px=hit.iloc[0]["Close"]; pnl+=(px-px0)*side - COST_PER_TRADE; n+=1; continue
        stop=px0 - side*KATR*atr; closed=False
        for _,r in fwd.iterrows():
            if side==1:
                stop=max(stop, r["Close"]-KATR*r["ATR5"])
                if r["Close"]<=stop: pnl+=(stop-px0)*side - COST_PER_TRADE; n+=1; closed=True; break
            else:
                stop=min(stop, r["Close"]+KATR*r["ATR5"])
                if r["Close"]>=stop: pnl+=(stop-px0)*side - COST_PER_TRADE; n+=1; closed=True; break
        if not closed:
            px=fwd.iloc[-1]["Close"]; pnl+=(px-px0)*side - COST_PER_TRADE; n+=1
    return pnl, n

def main():
    run=latest_run_dir(FEAT_ROOT)
    paths=[p for p in glob.glob(os.path.join(run,"*.parquet")) if os.path.basename(p)!="_ALL.parquet"]
    cutoff=pd.Timestamp.now(tz="Asia/Tokyo")-pd.Timedelta(days=N_DAYS)
    frames=[]
    for p in paths:
        df=pd.read_parquet(p)
        try: df=compute_missing_columns(df)
        except: continue
        df=df[(df.index>=cutoff) & (df["Session"].isin(["AM","PM"]))].dropna(subset=["ATR5","J"])
        if df.empty: continue
        frames.append(df)
    X=pd.concat(frames).sort_index()
    mdl=joblib.load(MODEL)
    # 特色不足を補完
    for c in FEATS:
        if c not in X.columns:
            X=compute_missing_columns(X); break
    X=X.dropna(subset=FEATS); X["p"]=mdl.predict(X[FEATS])

    best=None; best_obj=-1e18
    for THR in THR_GRID:
        for ABSJ in ABSJ_GRID:
            for VALUE in VALUE_GRID:
                for LAM in LAM_GRID:
                    for KATR in KATR_GRID:
                        for TMAX in TMAX_GRID:
                            pnl, n = simulate(X, THR, ABSJ, VALUE, LAM, KATR, TMAX)
                            if n==0: continue
                            if pnl>best_obj:
                                best_obj=pnl
                                best={"THR":THR,"ABSJ":ABSJ,"VALUE":VALUE,"LAM":LAM,"KATR":KATR,"TMAX":int(TMAX),
                                      "pnl":pnl,"trades":n,"days":N_DAYS}
    if best is None:
        best={"THR":0.8,"ABSJ":1.0,"VALUE":5e7,"LAM":0.8,"KATR":2.0,"TMAX":60,"pnl":0,"trades":0,"days":N_DAYS}
    os.makedirs(os.path.dirname(OUTP),exist_ok=True)
    json.dump(best, open(OUTP,"w"), ensure_ascii=False, indent=2)
    print("SAVED:", OUTP, best)

if __name__=="__main__": main()

