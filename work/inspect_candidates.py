import pandas as pd
from pathlib import Path
path = Path(r"output/excel/NIGHTLY_20251104/AM0930_j-cross/candidates_20251104.csv")
if not path.exists():
    print('missing file')
else:
    print('size', path.stat().st_size)
    df = pd.read_csv(path)
    print('rows', len(df))
    print(df.head())
