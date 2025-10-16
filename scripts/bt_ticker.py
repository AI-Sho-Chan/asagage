import pandas as pd, numpy as np
import yfinance as yf
from pathlib import Path
import datetime as dt
import argparse

# ---------- 0) Excel(Ticker) → list ----------
def load_ticker_from_excel(path: Path) -> list:
    df = pd.read_excel(path, sheet_name="Ticker", usecols="A", header=0)
    # A1=Code, A2〜がティッカー
    ticks = df["Code"].dropna().astype(str).tolist()
    return ticks

# ---------- 1) データ取得（直近7日×1分足） ----------
def fetch_recent_1m(tickers: list) -> pd.DataFrame:
    raw = yf.download(
        tickers, period="7d", interval="1m",
        group_by="ticker", threads=True, auto_adjust=False
    )
    out = []

    # ① MultiIndex/単一列の両対応（yfinanceはMultiIndex返却が一般的）
    if isinstance(raw.columns, pd.MultiIndex):  # 例: columns = [(ticker, 'Open'), ...]
        top = set(raw.columns.get_level_values(0))
        for t in tickers:
            if t not in top: 
                continue
            d = raw[t].dropna().copy()
            d["code"] = t
            d["ts"]   = pd.to_datetime(d.index).tz_localize(None)
            out.append(d.reset_index(drop=True))
    else:
        # 単一銘柄・単一列ケース
        d = raw.dropna().copy()
        # yfinanceの戻りが小文字open/closeのことがあるので安全に正規化
        colmap = {c: c.title() for c in ["open","high","low","close","volume"] if c in d.columns}
        if colmap: d = d.rename(columns=colmap)
        d["code"] = tickers[0] if tickers else "UNKNOWN"
        d["ts"]   = pd.to_datetime(d.index).tz_localize(None)
        out.append(d.reset_index(drop=True))

    if not out:
        raise RuntimeError("1m データが取得できる銘柄なし")

    all_df = pd.concat(out, ignore_index=True, sort=False)

    # ② ts 列は rename で壊さない（OHLCV だけタイトル化）
    colmap = {c: c.title() for c in ["open","high","low","close","volume"] if c in all_df.columns}
    if colmap: 
        all_df = all_df.rename(columns=colmap)

    # ③ intraday VWAP 等を計算
    all_df["date"]   = pd.to_datetime(all_df["ts"]).dt.date
    all_df["amt"]    = all_df["Close"] * all_df["Volume"]
    all_df["cumAmt"] = all_df.groupby(["code","date"])["amt"].cumsum()
    all_df["cumVol"] = all_df.groupby(["code","date"])["Volume"].cumsum().replace(0, np.nan)
    all_df["vwap"]   = all_df["cumAmt"] / all_df["cumVol"]

    return all_df[["ts","date","code","Open","High","Low","Close","Volume","vwap"]]

# ---------- 2) 学習/検証に分割（5日/2日） ----------
def split_train_test(df_all: pd.DataFrame, n_train_days: int = 5):
    dates = sorted(df_all["date"].unique())
    if len(dates) <= n_train_days:
        raise RuntimeError("日数が足りないので分割できない")
    train_dates = dates[:n_train_days]
    test_dates  = dates[n_train_days:]
    return (df_all[df_all["date"].isin(train_dates)].copy(),
            df_all[df_all["date"].isin(test_dates)].copy())

# ---------- 3) 指標計算 & 1銘柄評価（逆張りルール） ----------
def atr_ema(s: pd.Series, n: int) -> pd.Series:
    # TRのEMA（Wilder近似）。ATR 定義参考:contentReference[oaicite:3]{index=3}
    return s.ewm(alpha=1/n, adjust=False).mean()

def make_tr(g: pd.DataFrame) -> pd.Series:
    return pd.concat([
        (g["High"]-g["Low"]).abs(),
        (g["High"]-g["Close"].shift()).abs(),
        (g["Low"] -g["Close"].shift()).abs()
    ], axis=1).max(axis=1)

def eval_one(g: pd.DataFrame, n_atr:int, k_tp:float, k_sl:float,
             j_th:float, dj_th:float, vema_th:float, tmax:int=60) -> dict:
    g = g.sort_values("ts").copy()
    tr  = make_tr(g)
    atr = atr_ema(tr, n_atr).replace(0, np.nan)
    J   = (g["Close"] - g["vwap"]) / atr
    dJ  = J.diff()
    vE  = dJ.ewm(alpha=0.3, adjust=False).mean()
    sig = (J.abs()>=j_th) & (dJ.abs()>=dj_th) & (vE.abs()>=vema_th) & (np.sign(vE)!=np.sign(dJ))

    idx = g.index.to_list()
    wins=losses=flats=trades=0
    for i in g.index[sig]:
        a = atr.loc[i]
        if pd.isna(a): 
            continue
        trades += 1
        side = "BUY" if J.loc[i] < 0 else "SELL"
        px   = g.loc[i,"Close"]
        tp   = px + k_tp*a if side=="BUY" else px - k_tp*a
        sl   = px - k_sl*a if side=="BUY" else px + k_sl*a
        j    = idx.index(i); j2 = min(j+tmax, len(idx)-1)
        slc  = g.iloc[j+1:j2+1]
        hit_tp = (slc["High"]>=tp).any() if side=="BUY" else (slc["Low"]<=tp).any()
        hit_sl = (slc["Low"]<=sl).any() if side=="BUY" else (slc["High"]>=sl).any()
        if hit_sl and not hit_tp: losses += 1
        elif hit_tp and not hit_sl: wins += 1
        elif hit_tp and hit_sl: losses += 1
        else: flats += 1

    wr = wins/(wins+losses+flats) if (wins+losses+flats)>0 else 0.0
    pf = (wins*k_tp)/(losses*k_sl) if losses>0 else (999.0 if wins>0 else 0.0)
    return dict(winrate=wr, pf=pf, trades=trades)

# ---------- 4) グリッド探索（銘柄別の最適パラメタ） ----------
def add_pf_eff(df: pd.DataFrame, slip_bp=4.0, fee_bp=4.0) -> pd.DataFrame:
    c = (slip_bp + fee_bp) / 10000.0
    d = df.copy()
    d["PF_eff"] = (d["winrate"]*(d["TPk"]-c)) / (((1.0-d["winrate"])*(d["SLk"]+c)).replace(0,np.nan))
    return d

def grid_eval(df: pd.DataFrame):
    outs=[]
    for code, g in df.groupby("code"):
        best=None
        for n in [1,2,3,5]:
            for tp in [0.8,1.0,1.2]:
                for sl in [1.5,2.0]:
                    for jth in [0.6,0.8,1.0]:
                        for djth in [0.02,0.05]:
                            for vth in [0.02,0.05]:
                                m = eval_one(g, n, tp, sl, jth, djth, vth, tmax=60)
                                rec = dict(code=code, ATR_n=n, TPk=tp, SLk=sl,
                                           J_th=jth, dJ_th=djth, vEMA_th=vth, **m)
                                outs.append(rec)
                                if (best is None) or ((m["pf"], m["winrate"])>(best["pf"], best["winrate"])):
                                    best = rec
    grid = pd.DataFrame(outs)
    best = (grid.sort_values(["code","pf","winrate"], ascending=[True,False,False])
                 .drop_duplicates("code","first"))
    return best, grid

# ---------- 5) 出力フォルダ（ユニーク） ----------
def unique_output_folder(base: Path, label: str) -> Path:
    out_dir = base / f"{label}_{dt.date.today():%Y%m%d}"
    if out_dir.exists():
        k=1
        while True:
            alt = Path(f"{out_dir}_{k}")
            if not alt.exists():
                out_dir = alt; break
            k+=1
    out_dir.mkdir(parents=True, exist_ok=False)
    return out_dir

# ---------- 6) メイン ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help="C:\\AI\\asagake\\SHINSOKU.xlsm")
    ap.add_argument("--base-out", default=r"C:\AI\asagake\data\bt_temp")
    ap.add_argument("--train-days", type=int, default=5)
    args = ap.parse_args()

    ticks = load_ticker_from_excel(Path(args.excel))
    print("Ticker list count:", len(ticks))
    df_all = fetch_recent_1m(ticks)
    print("Fetched 1m rows:", len(df_all))

    df_train, df_test = split_train_test(df_all, n_train_days=args.train_days)
    print("Train rows:", len(df_train), "Test rows:", len(df_test))

    best_train, grid_tr = grid_eval(df_train)
    best_train = add_pf_eff(best_train)

    out_folder = unique_output_folder(Path(args.base_out), "SHINSOKU")
    best_train.to_csv(out_folder / "_SUMMARY_TRAIN.csv", index=False)
    grid_tr.to_csv(out_folder / "_GRID_TRAIN.csv", index=False)

    # フォワード
    outs=[]
    for _, r in best_train.iterrows():
        g = df_test[df_test["code"]==r["code"]]
        if g.empty: 
            continue
        m = eval_one(g, int(r["ATR_n"]), float(r["TPk"]), float(r["SLk"]),
                     float(r["J_th"]), float(r["dJ_th"]), float(r["vEMA_th"]), tmax=60)
        m.update(dict(code=r["code"], ATR_n=r["ATR_n"], TPk=r["TPk"], SLk=r["SLk"],
                      J_th=r["J_th"], dJ_th=r["dJ_th"], vEMA_th=r["vEMA_th"]))
        outs.append(m)
    fwd = pd.DataFrame(outs)
    if not fwd.empty: fwd = add_pf_eff(fwd)
    fwd.to_csv(out_folder / "_SUMMARY_TEST.csv", index=False)

    comp = best_train.merge(fwd, on="code", how="left", suffixes=("_TRAIN","_TEST"))
    comp.to_csv(out_folder / "_COMPARE.csv", index=False)
    print("Output folder:", out_folder)

if __name__=="__main__":
    main()
