import pandas as pd
from pathlib import Path
from collections import defaultdict, Counter

DATA = Path('analysis/all_trades_snapshot.csv')
trades = pd.read_csv(DATA, parse_dates=['date','ts'])
max_date = trades['date'].max()
start_date = max_date - pd.Timedelta(days=28)
recent = trades[trades['date'] >= start_date].copy()
recent['session_rank'] = recent['session'].map({
    'AM15':0,'AM0930':1,'AM0945':2,'AM1000':3,'AM1015':4,'AM1030':5,
    'MID1030':6,'PM1230':7,'PM1':8
}).fillna(99)
# choose top codes by trade count
code_counts = recent['code'].value_counts().head(5)
candidate_codes = code_counts.index.tolist()
print('codes', candidate_codes)

def summarize(df):
    total = len(df)
    if total == 0:
        return {'trades':0,'win_rate':None,'avg_bp':None,'sum_bp':0}
    wins = (df['pnl_bp'] > 0).sum()
    return {
        'trades': total,
        'win_rate': wins/total,
        'avg_bp': df['pnl_bp'].mean(),
        'sum_bp': df['pnl_bp'].sum()
    }

# Case A: best single strategy per code
case_a = {}
for code in candidate_codes:
    df_code = recent[recent['code']==code]
    if df_code.empty:
        case_a[code]={'trades':0,'win_rate':None,'avg_bp':None,'sum_bp':0,'combo':None}
        continue
    grouped = df_code.groupby(['session','signal_mode']).agg({'pnl_bp':'mean','code':'count'}).reset_index()
    best = grouped.sort_values('pnl_bp', ascending=False).iloc[0]
    combo_mask = (df_code['session']==best['session']) & (df_code['signal_mode']==best['signal_mode'])
    df_best = df_code[combo_mask]
    case_a[code] = summarize(df_best)
    case_a[code]['combo'] = f"{best['session']} x {best['signal_mode']}"

print('\nCase A (best combo only)')
for code, stats in case_a.items():
    print(code, stats)

# Case B: sequential multi-strategy with BAN rule
case_b = {}
for code in candidate_codes:
    df_code = recent[recent['code']==code].copy()
    if df_code.empty:
        case_b[code]={'trades':0,'win_rate':None,'avg_bp':None,'sum_bp':0}
        continue
    df_code = df_code.sort_values(['date','session_rank','ts'])
    records = []
    for date, df_day in df_code.groupby('date'):
        banned = False
        used_sessions = set()
        for _, row in df_day.sort_values(['session_rank','ts']).iterrows():
            sess = row['session']
            rank = row['session_rank']
            if rank == 99:
                continue
            if sess in used_sessions:
                continue
            if banned:
                continue
            records.append(row)
            used_sessions.add(sess)
            if row['pnl_bp'] <= 0:
                banned = True
    df_records = pd.DataFrame(records)
    case_b[code] = summarize(df_records)

print('\nCase B (sequential multi-strategy)')
for code, stats in case_b.items():
    print(code, stats)

# Aggregate comparison
import json
result = {'case_a':case_a,'case_b':case_b,'codes':candidate_codes,'start_date':str(start_date.date()),'end_date':str(max_date.date())}
Path('analysis/multi_strategy_simulation.json').write_text(json.dumps(result, indent=2))
print('\nSaved analysis/multi_strategy_simulation.json')
