import glob, pandas as pd
p=sorted(glob.glob(r"C:\AI\asagake\data\raw\yahoo_1m\7203.T\*.parquet"))[0]
df=pd.read_parquet(p)
print("FILE:",p)
print("INDEX:",type(df.index), "tz=",getattr(df.index,"tz",None))
print("COLUMNS:", list(df.columns))
import pandas as pd; print("IsMultiIndex:", isinstance(df.columns, pd.MultiIndex))
