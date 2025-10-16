# fix_ml_ts.py : MLデータの col "index" → "ts" に一括改名
import os, glob, pandas as pd
ROOT=r"C:\AI\asagake\data\ml"
run=max([p for p in glob.glob(os.path.join(ROOT,"RUN_*")) if os.path.isdir(p)], key=os.path.getmtime)
paths=glob.glob(os.path.join(run,"ds_*","*.parquet"))
fixed=0; skipped=0
for p in paths:
    try:
        df=pd.read_parquet(p)
        if "ts" in df.columns: 
            skipped+=1; continue
        if "index" in df.columns:
            df=df.rename(columns={"index":"ts"})
            df.to_parquet(p, index=False); fixed+=1
        else:
            skipped+=1
    except Exception as e:
        print("ERR", p, e)
print("RUN:", run, "fixed:", fixed, "skipped:", skipped, "total:", len(paths))
