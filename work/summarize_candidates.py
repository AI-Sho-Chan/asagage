import pandas as pd
from pathlib import Path
cand = Path(r"output/excel/candidates_nextday.csv")
df = pd.read_csv(cand)
print('rows', len(df))
print(df[['Ticker','SignalMode','session','forward_pf_eff','forward_winrate','forward_trades']].head(10).to_string(index=False))
print('\nby session counts:')
print(df.groupby(['session','SignalMode']).size())
