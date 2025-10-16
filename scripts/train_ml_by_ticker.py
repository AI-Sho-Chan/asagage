# -*- coding: utf-8 -*-
"""
train_ml_by_ticker.py
- 最新RUNの ds_* を銘柄ごとに読み、イベント十分な銘柄のみ LightGBM を個別学習
- 出力: models\by_ticker\<Ticker>.pkl と index: models\by_ticker\index.json
"""
import os, glob, json, numpy as np, pandas as pd, lightgbm as lgb, joblib

BASE=r"C:\AI\asagake"; ML=os.path.join(BASE,"data","ml"); MD=os.path.join(BASE,"models","by_ticker")
os.makedirs(MD,exist_ok=True)
FEATS=["J","dJ","vEMA","d2J","ATR5","IBS","Z20","ROC5","Turnover"]; TARGET="y"

def latest_run(root): 
    runs=[p for p in glob.glob(os.path.join(root,"RUN_*")) if os.path.isdir(p)]
    return max(runs,key=os.path.getmtime)

def load_all(run):
    paths=glob.glob(os.path.join(run,"ds_*","*.parquet"))
    frames=[]
    for p in paths:
        try:
            df=pd.read_parquet(p); 
            if set(["Ticker","ts",TARGET]).issubset(df.columns):
                frames.append(df)
        except: pass
    X=pd.concat(frames,ignore_index=True).dropna(subset=FEATS+[TARGET])
    return X

def main():
    run=latest_run(ML)
    X=load_all(run)
    idx={}
    for tkr,df in X.groupby("Ticker"):
        if (df[TARGET].sum()<100) or (len(df)<5000): 
            continue
        d=lgb.Dataset(df[FEATS], label=df[TARGET].astype(int))
        params={"objective":"binary","metric":"binary_logloss","learning_rate":0.05,
                "num_leaves":64,"min_data_in_leaf":200,"feature_fraction":0.8,"bagging_fraction":0.8,"bagging_freq":5}
        mdl=lgb.train(params,d,num_boost_round=600)
        path=os.path.join(MD,f"{tkr}.pkl"); joblib.dump(mdl,path)
        idx[tkr]=path
        print("saved",tkr,len(df))
    json.dump(idx, open(os.path.join(MD,"index.json"),"w"), indent=2)
    print("INDEX:", os.path.join(MD,"index.json"))
if __name__=="__main__": main()
