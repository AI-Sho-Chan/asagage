# bt_ticker.py — Excel(Ticker)→30日×1分(yahooquery)→21日学習+7日FWD→ユニーク出力
import pandas as pd, numpy as np
from pathlib import Path
import datetime as dt
import argparse

# yahooquery を使用（pip install yahooquery）
from yahooquery import Ticker

# ---------- Excel(Ticker) → list ----------
def load_ticker_from_excel(path: Path) -> list:
    df = pd.read_excel(path, sheet_name="Ticker", usecols="A", header=0)
    ticks = df["Code"].dropna().astype(str).tolist()
    return ticks

# ---------- 30日×1分の取得（yahooquery） ----------
def fetch_30d_1m(tickers: list) -> pd.DataFrame:
    # 1mは7日/reqだが、period='1mo'指定で30日を内部的に取得（制限は7日/reqだが連結で30日まで可）
    # 公式ガイド参照: 1分は7日/リクエストだが30日利用可能:contentReference[oaicite:2]{index=2}
    tq = Ticker(' '.join(tickers), asynchronous=True)
    df = tq.history(period='1mo', interval='1m')  # → MultiIndex(index=[symbol,date]) で返ることが多い:contentReference[oaicite:3]{index=3}
    if df.empty:
        raise RuntimeError("1m×30d データが取得できませんでした")
    if isinstance(df.index, pd.MultiIndex):
        df = df.reset_index()  # ['symbol','date',...]
        df = df.rename(columns={'symbol':'code','date':'ts'})
    else:
        df = df.reset_index().rename(columns={'symbol':'code','date':'ts'})
    # 列名正規化（OHLCV）
    colmap = {c: c.title() for c in ['open','high','low','close','volume'] if c in df.columns}
    df = df.rename(columns=colmap)
    # タイムゾーン除去 & 派生列
    df['ts'] = pd.to_datetime(df['ts']).dt.tz_localize(None)
    df['date'] = df['ts'].dt.date
    # intraday vwap
    df['amt']  = df['Close']*df['Volume']
    df['cumAmt'] = df.groupby(['code','date'])['amt'].cumsum()
    df['cumVol'] = df.groupby(['code','date'])['Volume'].cumsum().replace(0, np.nan)
    df['vwap']   = df['cumAmt']/df['cumVol']
    return df[['ts','date','code','Open','High','Low','Close','Volume','vwap']]

# ---------- 学習/検証に分割（21日/7日） ----------
def split_train_test(df_all: pd.DataFrame, n_train_days: int = 21):
    days = sorted(df_all['date'].unique())
    if len(days) <= n_train_days:
        raise RuntimeError("30日未満の取得。学習/検証に分割できません")
    train, test = days[:n_train_days], days[n_train_days:]
    return (df_all[df_all['date'].isin(train)].copy(),
            df_all[df_all['date'].isin(test)].copy())

# ---------- 逆張り評価ロジック ----------
def atr_ema(s: pd.Series, n: int) -> pd.Series:
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

def add_pf_eff(df: pd.DataFrame, slip_bp:float=4.0, fee_bp:float=4.0) -> pd.DataFrame:
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
                 .drop_duplicates(subset=["code"], keep="first"))  # ←キーワード引数で修正
    return best, grid

# ---------- 出力フォルダ（ユニーク化） ----------
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

# ---------- メイン ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--excel", required=True, help=r"C:\AI\asagake\SHINSOKU.xlsm")
    ap.add_argument("--base-out", default=r"C:\AI\asagake\data\bt_temp")
    ap.add_argument("--train-days", type=int, default=21)  # 21日学習+残り(~7日)FWD
    args = ap.parse_args()

    ticks = load_ticker_from_excel(Path(args.excel))
    print("Ticker list count:", len(ticks))

    df_all = fetch_30d_1m(ticks)
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
