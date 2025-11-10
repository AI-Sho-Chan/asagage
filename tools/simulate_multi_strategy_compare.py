import pandas as pd
from pathlib import Path

DATA = Path('analysis/all_trades_snapshot.csv')
trades = pd.read_csv(DATA, parse_dates=['date','ts'])
max_date = trades['date'].max()
start_date = max_date - pd.Timedelta(days=28)
recent = trades[trades['date'] >= start_date].copy()
recent['session_rank'] = recent['session'].map({
    'AM15':0,'AM0930':1,'AM0945':2,'AM1000':3,'AM1015':4,'AM1030':5,
    'MID1030':6,'PM1230':7,'PM1':8
}).fillna(99)

codes = recent['code'].value_counts().head(5).index.tolist()
print('codes', codes)

def summarize(df):
    total = len(df)
    if total == 0:
        return {'trades':0,'win_rate':0.0,'avg_bp':0.0,'sum_bp':0.0}
    wins = (df['pnl_bp'] > 0).sum()
    return {
        'trades': int(total),
        'win_rate': wins/total,
        'avg_bp': df['pnl_bp'].mean(),
        'sum_bp': df['pnl_bp'].sum()
    }

# Case ALL: allow全戦略同時
case_all = {code: summarize(recent[recent['code']==code]) for code in codes}

# Case BEST (1戦略のみ)
case_best = {}
for code in codes:
    df_code = recent[recent['code']==code]
    grouped = df_code.groupby(['session','signal_mode']).agg({'pnl_bp':'mean','code':'count'}).reset_index()
    best = grouped.sort_values('pnl_bp', ascending=False).iloc[0]
    mask = (df_code['session']==best['session']) & (df_code['signal_mode']==best['signal_mode'])
    case_best[code] = summarize(df_code[mask])
    case_best[code]['combo'] = f"{best['session']} x {best['signal_mode']}"

# Case QUEUE+BAN (時間順1本、負けたら停止)
case_queue_ban = {}
case_queue = {}
for code in codes:
    df_code = recent[recent['code']==code].sort_values(['date','session_rank','ts'])
    records_ban = []
    records_no_ban = []
    for date, day_df in df_code.groupby('date'):
        used_sessions = set()
        banned = False
        for _, row in day_df.sort_values(['session_rank','ts']).iterrows():
            if row['session_rank']==99 or row['session'] in used_sessions:
                continue
            # queue (no ban)
            records_no_ban.append(row)
            if not banned:
                records_ban.append(row)
            used_sessions.add(row['session'])
            if row['pnl_bp'] <= 0:
                banned = True
    case_queue_ban[code] = summarize(pd.DataFrame(records_ban))
    case_queue[code] = summarize(pd.DataFrame(records_no_ban))

import json
result = {
    'codes': codes,
    'start_date': str(start_date.date()),
    'end_date': str(max_date.date()),
    'case_all': case_all,
    'case_best': case_best,
    'case_queue': case_queue,
    'case_queue_ban': case_queue_ban,
}
Path('analysis/multi_strategy_compare.json').write_text(json.dumps(result, indent=2))
for label, data in [('ALL',case_all),('BEST',case_best),('QUEUE',case_queue),('QUEUE+BAN',case_queue_ban)]:
    print('\n',label)
    for code, stats in data.items():
        print(code, stats)
