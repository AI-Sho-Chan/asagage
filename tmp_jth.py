import pandas as pd
from pathlib import Path
cols = ['run','mode','signal_mode','J_th','forward_pf_eff','forward_winrate','forward_trades','train_pf_eff','train_winrate','train_trades']

def load_grid(root_dir):
    root = Path(root_dir)
    frames = []
    for path in root.glob('RUN_coarse_*_*/_GRID_FULL.csv'):
        try:
            df = pd.read_csv(path)
        except FileNotFoundError:
            continue
        df['run'] = path.parent.name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True)
    return df[cols]

def summarize(df):
    if df.empty:
        return pd.DataFrame()
    agg = df.groupby('J_th').agg(
        combos=('J_th','count'),
        forward_pf_mean=('forward_pf_eff','mean'),
        forward_win_mean=('forward_winrate','mean'),
        forward_trades_mean=('forward_trades','mean'),
        train_pf_mean=('train_pf_eff','mean'),
        train_win_mean=('train_winrate','mean'),
        train_trades_mean=('train_trades','mean')
    ).reset_index().sort_values('J_th')
    return agg

base = load_grid('output/bt30/NIGHTLY_20251028')
agg_base = summarize(base)
agg_base.to_csv('jth_baseline_stats.csv', index=False)
print('Baseline stats saved to jth_baseline_stats.csv')
print(agg_base.to_string(index=False, float_format=lambda x: f"{x:0.3f}"))

new = load_grid('output/bt30_test_newgrid/NIGHTLY_20251028')
agg_new = summarize(new)
if not agg_new.empty:
    agg_new.to_csv('jth_newgrid_stats.csv', index=False)
    print('\nNew-grid stats saved to jth_newgrid_stats.csv')
    print(agg_new.to_string(index=False, float_format=lambda x: f"{x:0.3f}"))
else:
    print('\nNew-grid stats not available')
