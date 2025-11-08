import pandas as pd
from pathlib import Path
root = Path(r"output/bt30/NIGHTLY_20251104_FAST/RUN_coarse_AM0930_j-only")
full = pd.read_csv(root/"_TOP_CANDIDATES._TOP20_codes.csv")
print("top codes:", list(full['code']))
refine_dir = Path(r"output/bt30/NIGHTLY_20251104_FAST/RUN_refine_AM0930_j-only")
summary_path = refine_dir/"_SUMMARY_TRAIN.partial.csv"
if summary_path.exists():
    df = pd.read_csv(summary_path)
    done = df['code'].unique().tolist()
    print("completed codes:", done)
    remaining = [c for c in full['code'] if c not in done]
    print("remaining codes:", remaining)
else:
    print("no summary partial")
