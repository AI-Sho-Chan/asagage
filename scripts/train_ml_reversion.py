# -*- coding: utf-8 -*-
"""
train_ml_reversion.py  (LightGBM v4 対応版)
- 入力: C:\AI\asagake\data\ml\RUN_*\ds_*/<Ticker>.parquet（make_ml_dataset.py の出力）
- モデル: LightGBM（binary） + Optuna（時系列CV; purge付）
- 指標: PR-AUC（average_precision_score）＋簡易 profit_proxy
- 出力:
  C:\AI\asagake\models\reversion.pkl
  C:\AI\asagake\models\best_params.json
  C:\AI\asagake\models\train_summary.json
  C:\AI\asagake\logs\ml\load_debug_LAST.json   ← 読込の合否ログ
"""

import os, glob, json, numpy as np, pandas as pd
from sklearn.metrics import average_precision_score
import lightgbm as lgb
import optuna, joblib

BASE      = r"C:\AI\asagake"
MLROOT    = os.path.join(BASE, "data", "ml")
MODEL_DIR = os.path.join(BASE, "models"); os.makedirs(MODEL_DIR, exist_ok=True)
LOG_DIR   = os.path.join(BASE, "logs", "ml"); os.makedirs(LOG_DIR, exist_ok=True)

FEATS    = ["J","dJ","vEMA","d2J","ATR5","IBS","Z20","ROC5","Turnover"]
TARGET   = "y"
RET_COL  = "ret"
NEED_COL = set(["ts","Ticker","J0", TARGET, RET_COL] + FEATS)

# --------- loader ----------
def latest_run_dir(root: str) -> str:
    runs = [p for p in glob.glob(os.path.join(root, "RUN_*")) if os.path.isdir(p)]
    if not runs: raise FileNotFoundError("No RUN_* under data\\ml")
    return max(runs, key=os.path.getmtime)

def scan_paths(run_dir: str) -> list[str]:
    paths = []
    for itv in ["ds_1m","ds_5m","ds_60m","ds_1d"]:
        d = os.path.join(run_dir, itv)
        if os.path.isdir(d):
            paths += glob.glob(os.path.join(d, "*.parquet"))
    return paths

def load_dataset(run_dir: str, max_files_per_interval: int = 999999) -> pd.DataFrame:
    paths = scan_paths(run_dir)
    dbg = {"run_dir": run_dir, "files_total": len(paths), "accepted": [], "rejected": []}
    frames = []
    for p in paths[:max_files_per_interval]:
        try:
            df = pd.read_parquet(p)
            cols = set(df.columns)
            # ts が無ければ index を ts に昇格（後方互換）
            if "ts" not in cols and "index" in cols:
                df = df.rename(columns={"index":"ts"}); cols=set(df.columns)
            miss = list(NEED_COL - cols)
            if miss:
                dbg["rejected"].append({"file": p, "reason": f"missing {miss}", "cols": list(df.columns)[:20]})
                continue
            use = df[list(NEED_COL)].dropna(subset=list(set(FEATS)|{TARGET,"ts"}))
            if use.empty:
                dbg["rejected"].append({"file": p, "reason": "empty after dropna"})
                continue
            frames.append(use); dbg["accepted"].append({"file": p, "rows": int(len(use))})
        except Exception as e:
            dbg["rejected"].append({"file": p, "reason": f"read_error: {e}"})
    with open(os.path.join(LOG_DIR, "load_debug_LAST.json"), "w", encoding="utf-8") as w:
        json.dump(dbg, w, ensure_ascii=False, indent=2)
    if not frames:
        print(f"[DEBUG] No valid training files. Details: {os.path.join(LOG_DIR,'load_debug_LAST.json')}")
        raise RuntimeError("No valid training files found.")
    X = pd.concat(frames, ignore_index=True).sort_values("ts").reset_index(drop=True)
    return X

# --------- CV & metrics ----------
def ts_folds(df: pd.DataFrame, k: int = 5, purge_frac: float = 0.02):
    n = len(df); idx = np.arange(n); folds = []
    for i in range(k):
        lo = int(n * (i / k)); hi = int(n * ((i+1) / k))
        p = int(np.ceil(n * purge_frac))
        tr_idx = idx[:max(lo - p, 0)]
        vl_idx = idx[lo:hi]
        if len(vl_idx) < 1000 or len(tr_idx) < 2000:  # 小分割はスキップ
            continue
        folds.append((df.iloc[tr_idx], df.iloc[vl_idx]))
    return folds

def pr_auc(vl: pd.DataFrame, p: np.ndarray) -> float:
    return float(average_precision_score(vl[TARGET].astype(int), p))

def profit_proxy(vl: pd.DataFrame, p: np.ndarray, thr: float = 0.6) -> float:
    mask = p >= thr
    return 0.0 if not mask.any() else float(vl.loc[mask, RET_COL].mean())

# --------- Optuna ----------
def run_optuna(df: pd.DataFrame, n_trials: int = 50, k: int = 5):
    FOLDS = ts_folds(df, k=k, purge_frac=0.02)

    def objective(trial: optuna.Trial):
        params = {
            "objective": "binary", "metric": "binary_logloss",
            "boosting_type": "gbdt", "verbosity": -1,
            "learning_rate": trial.suggest_float("learning_rate", 0.01, 0.2, log=True),
            "num_leaves": trial.suggest_int("num_leaves", 31, 256, log=True),
            "min_data_in_leaf": trial.suggest_int("min_data_in_leaf", 100, 5000, log=True),
            "feature_fraction": trial.suggest_float("feature_fraction", 0.6, 1.0),
            "bagging_fraction": trial.suggest_float("bagging_fraction", 0.6, 1.0),
            "bagging_freq": trial.suggest_int("bagging_freq", 1, 10),
            "lambda_l2": trial.suggest_float("lambda_l2", 1e-3, 10.0, log=True),
        }
        nround = trial.suggest_int("num_boost_round", 300, 1500, log=True)

        scores, profs = [], []
        for tr, vl in FOLDS:
            dtr = lgb.Dataset(tr[FEATS], label=tr[TARGET].astype(int))
            dvl = lgb.Dataset(vl[FEATS], label=vl[TARGET].astype(int))
            # LightGBM v4: verbose_eval は callbacks で制御
            mdl = lgb.train(
                params, dtr, num_boost_round=nround, valid_sets=[dvl],
                callbacks=[lgb.log_evaluation(period=0)]  # サイレント
            )
            p = mdl.predict(vl[FEATS])
            scores.append(pr_auc(vl, p))
            profs.append(profit_proxy(vl, p, thr=0.6))
        return 0.0 if not scores else float(np.mean(scores) + 0.05*np.mean(profs))

    study = optuna.create_study(direction="maximize")
    study.optimize(objective, n_trials=n_trials)
    return study.best_params, study.best_value

# --------- Full fit ----------
def fit_full(df: pd.DataFrame, params: dict) -> lgb.Booster:
    params = dict(params)
    params.update({"objective":"binary","metric":"binary_logloss","boosting_type":"gbdt","verbosity":-1})
    nround = int(params.pop("num_boost_round", 800))
    d = lgb.Dataset(df[FEATS], label=df[TARGET].astype(int))
    mdl = lgb.train(params, d, num_boost_round=nround, callbacks=[lgb.log_evaluation(period=0)])
    return mdl

# --------- main ----------
def main():
    run_dir = latest_run_dir(MLROOT)
    print("RUN:", run_dir)
    df = load_dataset(run_dir)
    print("DATA:", len(df), "rows from", df["Ticker"].nunique(), "tickers")

    best_params, best_score = run_optuna(df, n_trials=50, k=5)
    print("BEST_SCORE:", best_score)
    print("BEST_PARAMS:", best_params)

    mdl = fit_full(df, best_params)
    model_path = os.path.join(MODEL_DIR, "reversion.pkl")
    joblib.dump(mdl, model_path)

    p = mdl.predict(df[FEATS])
    overall_prauc = pr_auc(df, p)
    thr = 0.6
    prof = profit_proxy(df, p, thr=thr)

    json.dump(best_params, open(os.path.join(MODEL_DIR, "best_params.json"), "w"))
    json.dump({
        "run_dir": run_dir,
        "rows": int(len(df)),
        "tickers": int(df["Ticker"].nunique()),
        "best_score": float(best_score),
        "overall_pr_auc": float(overall_prauc),
        "profit_proxy_thr": float(thr),
        "profit_proxy_val": float(prof),
        "features": FEATS,
    }, open(os.path.join(MODEL_DIR, "train_summary.json"), "w"), ensure_ascii=False, indent=2)

    print("MODEL_SAVED:", model_path)
    print("SUMMARY:", os.path.join(MODEL_DIR, "train_summary.json"))

if __name__ == "__main__":
    main()
