# force_build_one.py (robust)
import os, glob, pandas as pd, numpy as np

BASE=r"C:\AI\asagake"
RAW =os.path.join(BASE,r"data\raw\yahoo_1m","7203.T")
RUN =pd.Timestamp.now().strftime("RUN_%Y%m%d_%H%M")+"_one"
OUT =os.path.join(BASE,r"data\proc\features_1m",RUN); os.makedirs(OUT,exist_ok=True)

def flatten_cols(df):
    if isinstance(df.columns, pd.MultiIndex):
        # 例: ('Open','7203.T') / ('7203.T','Open') などを平坦化
        new=[]
        for tpl in df.columns:
            vals=[str(x) for x in tpl if x is not None]
            new.append("_".join(vals))
        df=df.copy(); df.columns=new
    return df

def normalize_ohlcv(df):
    df=flatten_cols(df)
    cols={c:c for c in df.columns}
    # 正規化キー（小文字・空白除去）
    key={c:"".join(str(c).lower().split()) for c in df.columns}
    def pick(*names):
        names=[n.lower() for n in names]
        for c,k in key.items():
            if k in names: return c
        return None
    m={}
    m["Open"]  = pick("open","open_7203.t","7203.t_open")
    m["High"]  = pick("high","high_7203.t","7203.t_high")
    m["Low"]   = pick("low","low_7203.t","7203.t_low")
    m["Close"] = pick("close","adjclose","close_7203.t","7203.t_close")
    m["Volume"]= pick("volume","vol","volume_7203.t","7203.t_volume")
    # 失敗時は列名の先頭一致で救済
    for std, cand in list(m.items()):
        if cand is None:
            for c in df.columns:
                if str(c).lower().startswith(std.lower()): m[std]=c; break
    missing=[k for k,v in m.items() if v is None]
    if missing: raise AssertionError(f"cannot map OHLCV: missing={missing}, cols={list(df.columns)[:10]}")
    out=df[[m["Open"],m["High"],m["Low"],m["Close"],m["Volume"]]].copy()
    out.columns=["Open","High","Low","Close","Volume"]
    return out

def ensure_dt_index(df):
    if not isinstance(df.index, pd.DatetimeIndex):
        # 見込みの日時列名を探す
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

files=glob.glob(os.path.join(RAW,"*.parquet"))
assert files, "NO RAW files for 7203.T"
df=pd.concat([pd.read_parquet(f) for f in files]).sort_index()
df = normalize_ohlcv(df)
df.index = ensure_dt_index(df)

df["VWAPd"]=vwap_intraday(df)
df["ATR5"]=atr_n(df,5)
df["J"]=(df["Close"]-df["VWAPd"])/df["ATR5"]
df["dJ"]=df["J"].diff(); df["vEMA"]=df["dJ"].ewm(span=5,adjust=False).mean(); df["d2J"]=df["dJ"].diff()

hhmm=df.index.strftime("%H%M").astype(int)
df["Session"]=np.where((hhmm>=900)&(hhmm<1130),"AM",
                np.where((hhmm>=1230)&(hhmm<1500),"PM","OFF"))

keep=["Open","High","Low","Close","Volume","VWAPd","ATR5","J","dJ","d2J","vEMA","Session"]
df=df[keep].dropna(subset=["VWAPd","ATR5","J"])
print("built rows:",len(df))
df.to_parquet(os.path.join(OUT,"7203.T.parquet"))
print("OUT:",OUT)
