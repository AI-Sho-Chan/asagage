# -*- coding: utf-8 -*-
"""
infer_reversion.py  (opt_thresholdsの自動反映 + 堅牢補完)
"""
import os, glob, json
import numpy as np
import pandas as pd
import joblib

BASE = r"C:\AI\asagake"
MODEL = os.path.join(BASE, "models", "reversion.pkl")
OPT_PATH = os.path.join(BASE, "models", "opt_thresholds.json")
FEAT_ROOT = os.path.join(BASE, "data", "proc", "features_1m")
OUT_ROOT  = os.path.join(BASE, "data", "signals")
LOG_DIR   = os.path.join(BASE, "logs", "ml"); os.makedirs(LOG_DIR, exist_ok=True)

# 既定（opt_thresholds.json があれば上書き）
THR=0.80; LAM=0.80; K_ATR=2.0; MIN_ABS_J=1.0; MIN_VALUE=5e7

FEATS = ["J","dJ","vEMA","d2J","ATR5","IBS","Z20","ROC5","Turnover"]
BASE_NEED = {"Open","High","Low","Close","Volume","Session","Ticker"}

def latest_features_dir(root: str) -> str:
    runs = [p for p in glob.glob(os.path.join(root, "RUN_*")) if os.path.isdir(p)]
    if runs: return max(runs, key=os.path.getmtime)
    subs = [p for p in glob.glob(os.path.join(root, "*")) if os.path.isdir(p)]
    if subs: return max(subs, key=os.path.getmtime)
    return None

def ensure_jst(df: pd.DataFrame) -> pd.DataFrame:
    if isinstance(df.index, pd.MultiIndex): df = df.reset_index()
    if not isinstance(df.index, pd.DatetimeIndex):
        for c in ["Datetime","Date","Time","ts","index"]:
            if c in df.columns:
                df = df.set_index(pd.to_datetime(df[c], utc=True, errors="coerce")).drop(columns=[c]); break
        if not isinstance(df.index, pd.DatetimeIndex):
            df.index = pd.to_datetime(df.index, utc=True, errors="coerce")
    if df.index.tz is None: df.index = df.index.tz_localize("UTC")
    return df.tz_convert("Asia/Tokyo")

def vwap_intraday(df: pd.DataFrame) -> pd.Series:
    g = df.index.date
    pv = (df["Close"]*df["Volume"]).groupby(g).cumsum()
    vv = df["Volume"].groupby(g).cumsum().replace(0, np.nan)
    return pv/vv

def atr_n(df: pd.DataFrame, n=5) -> pd.Series:
    pc = df["Close"].shift(1)
    tr = pd.concat([(df["High"]-df["Low"]).abs(),
                    (df["High"]-pc).abs(),
                    (df["Low"] -pc).abs()], axis=1).max(axis=1)
    return tr.rolling(n, min_periods=n).mean()

def compute_missing_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = ensure_jst(df)
    if not BASE_NEED.issubset(df.columns):
        missing = list(BASE_NEED - set(df.columns))
        raise ValueError(f"missing base cols: {missing}")
    if "ATR5" not in df.columns or df["ATR5"].isna().all(): df["ATR5"] = atr_n(df, 5)
    if "VWAPd" not in df.columns or df["VWAPd"].isna().all(): df["VWAPd"] = vwap_intraday(df)
    if "J" not in df.columns or df["J"].isna().all(): df["J"] = (df["Close"] - df["VWAPd"]) / df["ATR5"]
    if "dJ" not in df.columns:   df["dJ"]  = df["J"].diff()
    if "vEMA" not in df.columns: df["vEMA"]= df["dJ"].ewm(span=5, adjust=False).mean()
    if "d2J" not in df.columns:  df["d2J"] = df["dJ"].diff()
    if "IBS" not in df.columns:
        rng = (df["High"] - df["Low"]).replace(0, np.nan)
        df["IBS"] = (df["Close"] - df["Low"]) / rng
    if "SMA20" not in df.columns: df["SMA20"] = df["Close"].rolling(20).mean()
    if "STD20" not in df.columns: df["STD20"] = df["Close"].rolling(20).std()
    if "Z20" not in df.columns:   df["Z20"]   = (df["Close"] - df["SMA20"]) / df["STD20"]
    if "ROC5" not in df.columns:  df["ROC5"]  = df["Close"].pct_change(5)
    if "Turnover" not in df.columns: df["Turnover"] = df["Close"] * df["Volume"]
    return df

def load_features_dir(d: str) -> pd.DataFrame:
    files = [p for p in glob.glob(os.path.join(d, "**", "*.parquet"), recursive=True)
             if os.path.basename(p) != "_ALL.parquet"]
    dbg = {"run_dir": d, "files_total": len(files), "accepted": [], "rejected": []}
    frames=[]
    for p in files:
        try:
            df = pd.read_parquet(p)
            df = compute_missing_columns(df)
            need = BASE_NEED | set(FEATS)
            missing = list(need - set(df.columns))
            if missing:
                dbg["rejected"].append({"file": p, "reason": f"missing {missing}"}); continue
            df = df.dropna(subset=["ATR5","J"])
            frames.append(df); dbg["accepted"].append({"file": p, "rows": int(len(df))})
        except Exception as e:
            dbg["rejected"].append({"file": p, "reason": f"error: {e}"}); continue
    with open(os.path.join(LOG_DIR, "infer_debug_LAST.json"), "w", encoding="utf-8") as w:
        json.dump(dbg, w, ensure_ascii=False, indent=2)
    if not frames:
        raise RuntimeError(f"RUN内に有効な特徴量が0件: {d}\n詳細: {os.path.join(LOG_DIR,'infer_debug_LAST.json')}")
    return pd.concat(frames).sort_index()

def main():
    # 最適しきいの読込
    global THR, LAM, K_ATR, MIN_ABS_J, MIN_VALUE
    if os.path.exists(OPT_PATH):
        try:
            opt=json.load(open(OPT_PATH,"r"))
            THR=opt.get("THR",THR); LAM=opt.get("LAM",LAM); K_ATR=opt.get("KATR",K_ATR)
            MIN_ABS_J=opt.get("ABSJ",MIN_ABS_J); MIN_VALUE=opt.get("VALUE",MIN_VALUE)
        except: pass

    if not os.path.exists(MODEL): raise FileNotFoundError(MODEL)
    mdl = joblib.load(MODEL)

    run = latest_features_dir(FEAT_ROOT)
    if not run: raise FileNotFoundError(FEAT_ROOT)
    feat = load_features_dir(run)

    df = feat[feat["Session"].isin(["AM","PM"])].copy()
    df = df.dropna(subset=FEATS)
    df["p"]    = mdl.predict(df[FEATS])
    df["side"] = np.where(df["J"] < 0, "BUY", "SELL")

    # TP/SL
    dJ_tp = (1.0 - LAM) * np.abs(df["J"]) * df["ATR5"]
    df["TP_price"] = np.where(df["side"]=="BUY",  df["Close"] + dJ_tp, df["Close"] - dJ_tp)
    dSL = K_ATR * df["ATR5"]
    df["SL_price"] = np.where(df["side"]=="BUY",  df["Close"] - dSL,  df["Close"] + dSL)

    # 強化フィルタ
    df["Value"] = df["Close"] * df["Volume"]
    picks = df[(df["p"] >= THR) & (np.abs(df["J"]) >= MIN_ABS_J) & (df["Value"] >= MIN_VALUE)].copy()
    # 最新1本/銘柄
    picks = (picks.reset_index().rename(columns={"index":"ts"})
                    .sort_values(["Ticker","ts"], ascending=[True, False])
                    .drop_duplicates(["Ticker"]))

    # 出力
    day = pd.Timestamp.now(tz="Asia/Tokyo").strftime("%Y-%m-%d")
    out_dir = os.path.join(OUT_ROOT, day); os.makedirs(out_dir, exist_ok=True)
    all_csv = os.path.join(out_dir, "signals_1m.csv")
    top_csv = os.path.join(out_dir, "top_candidates.csv")
    df.reset_index().rename(columns={"index":"ts"}).to_csv(all_csv, index=False, encoding="utf-8-sig")
    picks.sort_values(["p","ts"], ascending=[False, True]).to_csv(top_csv, index=False, encoding="utf-8-sig")

    print("RUN_FEAT:", run)
    print("OUT_ALL :", all_csv, " rows=", len(df))
    print("OUT_TOP :", top_csv, " rows=", len(picks))
    print(f"(auto) THR={THR} |J|>={MIN_ABS_J} Value>={MIN_VALUE:.0f}  LAM={LAM} K_ATR={K_ATR}")

if __name__ == "__main__":
    main()
