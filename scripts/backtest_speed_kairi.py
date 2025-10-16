# backtest_speed_kairi.py  — grid search: ATR_n in {1,2,3,5}, TPk in {1,2,3}, SLk in {1,2}
import os, glob
import numpy as np
import pandas as pd

RUN_DIR = r"C:\AI\asagake\data\proc\features_1m"

def load_latest_run():
    runs = sorted(glob.glob(os.path.join(RUN_DIR, "RUN_*")), reverse=True)
    if not runs:
        raise SystemExit("no RUN_* found under " + RUN_DIR)
    return runs[0]

# ---- column normalization ----
def norm_columns(df: pd.DataFrame) -> pd.DataFrame:
    low = {c.lower(): c for c in df.columns}

    def pick(cands, req=True, name="col"):
        for k in cands:
            if k in low:
                return low[k]
        if req:
            raise KeyError(f"missing column for {name}. got {list(df.columns)[:20]}")
        return None

    c_code = pick(["code","ticker","symbol"], True, "code")
    c_close= pick(["close","px_close","last"], True, "close")
    c_high = pick(["high","px_high"], True, "high")
    c_low  = pick(["low","px_low"], True, "low")
    c_vwap = pick(["vwap","vwapd"], True, "vwap")
    c_ts   = pick(["ts","datetime","time"], False, "ts")

    ren = {c_code:"code", c_close:"close", c_high:"high", c_low:"low", c_vwap:"vwap"}
    if c_ts: ren[c_ts] = "ts"
    out = df.rename(columns=ren).copy()

    if "ts" not in out.columns:
        out["ts"] = out.groupby("code").cumcount()

    return out[["ts","code","close","high","low","vwap"]].copy()

# ---- ATR(EMA) ----
def atr_ema(s: pd.Series, n: int) -> pd.Series:
    """s: Series of TrueRange, returns EMA(n)"""
    return s.ewm(alpha=1/n, adjust=False).mean()

def eval_strategy(df: pd.DataFrame, n_atr: int, k_tp: float, k_sl: float, tmax: int = 60) -> pd.DataFrame:
    df = norm_columns(df)

    outs = []
    for code, g in df.groupby("code"):
        g = g.sort_values("ts").copy()

        # --- True Range (Series) ---
        tr = pd.concat([
            (g["high"] - g["low"]).abs(),
            (g["high"] - g["close"].shift()).abs(),
            (g["low"]  - g["close"].shift()).abs()
        ], axis=1).max(axis=1)

        atr = atr_ema(tr, n_atr).replace(0, np.nan)
        J   = (g["close"] - g["vwap"]) / atr
        dJ  = J.diff()
        vEMA= dJ.ewm(alpha=0.3, adjust=False).mean()

        # simple signal (same as sheet draft): |J|>=0.8 & sign flip
        sig = (J.abs() >= 0.8) & (np.sign(vEMA) != np.sign(dJ))

        wins = losses = flats = 0
        idx = g.index.to_list()
        for i in g.index[sig]:
            a = atr.loc[i]
            if pd.isna(a): 
                continue
            side = "BUY" if J.loc[i] < 0 else "SELL"
            px   = g.loc[i, "close"]
            tp   = px + k_tp*a if side == "BUY" else px - k_tp*a
            sl   = px - k_sl*a if side == "BUY" else px + k_sl*a

            j  = idx.index(i)
            j2 = min(j + tmax, len(idx)-1)
            slc = g.iloc[j+1:j2+1]

            hit_tp = (slc["high"] >= tp).any() if side == "BUY" else (slc["low"]  <= tp).any()
            hit_sl = (slc["low"]  <= sl).any() if side == "BUY" else (slc["high"] >= sl).any()

            if hit_tp and not hit_sl:
                wins += 1
            elif hit_sl and not hit_tp:
                losses += 1
            else:
                flats += 1

        cnt = wins + losses + flats
        winrate = wins / cnt if cnt else 0.0
        pf = (wins * k_tp) / (losses * k_sl) if losses else (999.0 if wins > 0 else 0.0)
        outs.append(dict(code=code, winrate=winrate, pf=pf, cnt=cnt))

    return pd.DataFrame(outs)

def main():
    run = load_latest_run()
    files = glob.glob(os.path.join(run, "*.parquet"))
    if not files:
        raise SystemExit("no parquet in " + run)
    df = pd.concat([pd.read_parquet(p) for p in files], ignore_index=True)

    grid = []
    for n in [1, 2, 3, 5]:
        for tp in [1, 2, 3]:
            for sl in [1, 2]:
                m = eval_strategy(df, n, tp, sl)
                m["ATR_n"] = n
                m["TPk"]   = tp
                m["SLk"]   = sl
                grid.append(m)

    res  = pd.concat(grid, ignore_index=True)
    best = res.sort_values(["code","winrate","pf"], ascending=[True, False, False]).drop_duplicates("code", keep="first")
    out  = best[["code","winrate","pf","ATR_n","TPk","SLk"]]
    out_path = os.path.join(run, "_SUMMARY.csv")
    out.to_csv(out_path, index=False)
    print("written:", out_path)

if __name__ == "__main__":
    main()
