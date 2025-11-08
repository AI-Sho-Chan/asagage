import datetime as dt
from pathlib import Path
from typing import Sequence, Dict, List

import numpy as np
import pandas as pd
from yahooquery import Ticker

# Config
LOOKBACK_DAYS = 20
SESSION_START = dt.time(9, 0)
SESSION_END = dt.time(9, 15)
COST_BP = 8.0  # slip + fee
PARAM = dict(ATR_n=3, TPk=1.2, SLk=2.0, J_th=0.6)  # example band from prior discussion

AMT_GRID = Path('output/bt30/RUN_coarse_AM15_20251017_114040/_GRID_FULL.csv')
OUT = Path('out/gap_analysis_AM15_amt.csv')

def unique_codes_from_grid(path: Path) -> List[str]:
    df = pd.read_csv(path)
    return df['code'].astype(str).dropna().unique().tolist()

def fetch_1m(tickers: Sequence[str], start: dt.date, end: dt.date) -> pd.DataFrame:
    tq = Ticker(list(tickers), asynchronous=True)
    df = tq.history(start=str(start), end=str(end), interval='1m')
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    df = df.reset_index()
    if 'symbol' in df.columns:
        df = df.rename(columns={'symbol': 'code'})
    if 'date' in df.columns and 'ts' not in df.columns:
        df = df.rename(columns={'date': 'ts'})
    df['ts'] = pd.to_datetime(df['ts'])
    df['date'] = df['ts'].dt.date
    df['amt'] = df['close'] * df['volume']
    # intraday VWAP using cumulative sums
    df = df.sort_values(['code', 'ts']).reset_index(drop=True)
    df['cumAmt'] = df.groupby(['code', 'date'])['amt'].cumsum()
    df['cumVol'] = df.groupby(['code', 'date'])['volume'].cumsum().replace(0, np.nan)
    df['vwap'] = df['cumAmt'] / df['cumVol']
    return df[['ts','code','open','high','low','close','volume','vwap','date']]

def true_range(df: pd.DataFrame) -> pd.Series:
    prev_close = df['close'].shift(1)
    tr = pd.concat([
        (df['high'] - df['low']).abs(),
        (df['high'] - prev_close).abs(),
        (df['low'] - prev_close).abs(),
    ], axis=1).max(axis=1)
    return tr

def atr_ema(tr: pd.Series, n: int) -> pd.Series:
    return tr.ewm(alpha=2/(n+1), adjust=False).mean()

def simulate_day(day_df: pd.DataFrame, params: Dict) -> List[Dict]:
    # mark session window
    m = day_df["ts"].dt.time.between(SESSION_START, SESSION_END)
    enter_idx = list(day_df.index[m])
    if not enter_idx:
        return []
    # ATR is precomputed in column "atr"
    J = (day_df["close"] - day_df["vwap"]) / day_df["atr"]
    sig = J.abs() >= float(params["J_th"])  # j-only
    results = []
    last_idx = int(day_df.index[-1])
    for idx in enter_idx:
        if not sig.loc[idx]:
            continue
        a = day_df.loc[idx, "atr"]
        if not np.isfinite(a) or a == 0:
            continue
        px = float(day_df.loc[idx, "close"])
        side = "BUY" if J.loc[idx] < 0 else "SELL"
        tp = px + float(params["TPk"]) * a if side == "BUY" else px - float(params["TPk"]) * a
        sl = px - float(params["SLk"]) * a if side == "BUY" else px + float(params["SLk"]) * a
        exit_px = None
        bars = 0
        for j in range(idx + 1, last_idx + 1):
            hi = float(day_df.loc[j, "high"])
            lo = float(day_df.loc[j, "low"])
            bars += 1
            if side == "BUY":
                if lo <= sl:
                    exit_px = sl
                    break
                if hi >= tp:
                    exit_px = tp
                    break
            else:
                if hi >= sl:
                    exit_px = sl
                    break
                if lo <= tp:
                    exit_px = tp
                    break
        if exit_px is None:
            exit_px = float(day_df.loc[last_idx, "close"])
        pnl_bp = ((exit_px - px) / px) * 10000.0 if side == "BUY" else ((px - exit_px) / px) * 10000.0
        pnl_bp -= COST_BP
        results.append(dict(pnl_bp=pnl_bp, bars=bars, idx=idx))
    return results

def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    codes = unique_codes_from_grid(AMT_GRID)[:300]
    end = dt.date.today() + dt.timedelta(days=1)
    start = end - dt.timedelta(days=LOOKBACK_DAYS)
    raw = fetch_1m(codes, start, end)
    if raw.empty:
        print('no data'); return
    # daily gap (open_first - prev_close)/prev_close*bp
    daily = raw.groupby(['code','date']).agg(open_first=('open','first'), close_last=('close','last')).reset_index()
    daily = daily.sort_values(['code','date']).reset_index(drop=True)
    daily['prev_close'] = daily.groupby('code')['close_last'].shift(1)
    daily['gap_bp'] = ((daily['open_first'] - daily['prev_close'])/daily['prev_close'])*10000.0
    df = raw.merge(daily[['code','date','gap_bp']], on=['code','date'], how='left')
    # Precompute ATR per code across all days
    df = df.sort_values(['code','ts']).reset_index(drop=True)
    def _atr_per_code(sub: pd.DataFrame) -> pd.DataFrame:
        pc = sub['close'].shift(1)
        tr = pd.concat([(sub['high']-sub['low']).abs(), (sub['high']-pc).abs(), (sub['low']-pc).abs()], axis=1).max(axis=1)
        atr = tr.ewm(alpha=2/(PARAM['ATR_n']+1), adjust=False).mean()
        sub = sub.copy()
        sub['atr'] = atr.replace(0, np.nan)
        return sub
    df = df.groupby('code', group_keys=False).apply(_atr_per_code)

    # simulate per (code,date)
    records = []
    for (code, date), day_df in df.groupby(['code','date']):
        day_df = day_df.sort_values('ts').reset_index(drop=True)
        if pd.isna(day_df['gap_bp'].iloc[0]):
            continue
        trades = simulate_day(day_df, PARAM)
        for t in trades:
            records.append({
                'code': code,
                'date': date,
                'gap_bp': float(day_df['gap_bp'].iloc[0]),
                'pnl_bp': t['pnl_bp'],
                'bars': t['bars'],
            })
    if not records:
        print('no trades generated'); return
    rd = pd.DataFrame(records)
    rd['gap_abs'] = rd['gap_bp'].abs()
    bins = [-1e9, 20, 50, 80, 120, 1e9]
    labels = ['<20bp','20-50','50-80','80-120','>=120']
    rd['gap_bucket'] = pd.cut(rd['gap_abs'], bins=bins, labels=labels)
    summ = rd.groupby('gap_bucket').agg(
        trades=('pnl_bp','count'),
        winrate=('pnl_bp', lambda s: (s>0).mean()),
        mean_pnl_bp=('pnl_bp','mean'),
        median_pnl_bp=('pnl_bp','median'),
        mean_bars=('bars','mean'),
    ).reset_index()
    summ.to_csv(OUT, index=False, encoding='utf-8-sig')
    print('written', OUT)

if __name__ == '__main__':
    main()


