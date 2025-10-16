# -*- coding: utf-8 -*-
"""
Yahoo Backfill robust (v2):
- UNIVERSE_EXTRA!A列 → {4桁}.T ユニバース
- 無効/分足なしを自動スキップ（fast_info + 直近7日1mテスト）
- 既存ファイルは読み飛ばし（再開可）
- 週末は最初から除外（1m/5m/60m/1d共通）
- 例外はすべて Exception で受けて status に文字列化（yfinance内部例外に依存しない）
- 並列取得、詳細ログ JSON を保存
"""
import os, time, json
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import yfinance as yf

# ====== 設定 ======
BASE = r"C:\AI\asagake"
XLSM = os.path.join(BASE, "excel", "ASAGAKE_template_30_safe.xlsm")
RAW  = os.path.join(BASE, "data", "raw")
LOGD = os.path.join(BASE, "logs", "yahoo_backfill")
os.makedirs(RAW, exist_ok=True)
os.makedirs(LOGD, exist_ok=True)

UNIVERSE_SHEET = "UNIVERSE_EXTRA"
UNIVERSE_COL   = 0
UNIVERSE_N     = 200
BAD = {"8729.T", "0285.T"}  # 取得不可をここに追記

PLAN = [
    ("1m",  30,     7),
    ("5m",  60,    30),
    ("60m", 730,   60),
    ("1d",  365*20, 365),
]

MAX_WORKERS = 6
RETRY_EACH  = 2

# ====== ユーティリティ ======
def load_universe():
    df = pd.read_excel(XLSM, sheet_name=UNIVERSE_SHEET, usecols=[UNIVERSE_COL], engine="openpyxl")
    codes = df.iloc[:, 0].astype(str).str.extract(r"(\d{4})")[0].dropna().head(UNIVERSE_N)
    tks = [f"{c}.T" for c in codes]
    return [t for t in tks if t not in BAD]

def ensure_jst_index(df: pd.DataFrame) -> pd.DataFrame:
    if not isinstance(df.index, pd.DatetimeIndex):
        df.index = pd.to_datetime(df.index, utc=True, errors="coerce")
    if df.index.tz is None:
        df.index = df.index.tz_localize("UTC")
    return df.tz_convert("Asia/Tokyo")

def chunks(days, span):
    end = pd.Timestamp.now(tz="Asia/Tokyo").floor("D") + pd.Timedelta(days=1)
    start = end - pd.Timedelta(days=days)
    cur = start
    while cur < end:
        nxt = min(cur + pd.Timedelta(days=span), end)
        yield cur.tz_localize(None), nxt.tz_localize(None)
        cur = nxt

def has_intraday_data(tkr: str) -> bool:
    # timezone チェック
    try:
        fi = yf.Ticker(tkr).fast_info
        if not getattr(fi, "timezone", None):
            return False
    except Exception:
        return False
    # 直近7日 1m の軽テスト
    try:
        end = pd.Timestamp.now(tz="Asia/Tokyo").floor("D") + pd.Timedelta(days=1)
        start = end - pd.Timedelta(days=7)
        df = safe_download(tkr, start, end, "1m")
        return not df.empty
    except Exception:
        return False

def safe_download(tkr, s, e, interval):
    """yfinance.download を安全に。失敗時は空DataFrameを返す。"""
    try:
        df = yf.download(tkr, start=s.tz_localize(None), end=e.tz_localize(None),
                         interval=interval, progress=False, prepost=False, auto_adjust=False)
        if df is None:
            return pd.DataFrame()
        # yfinance の内部 concat エラー対策：None/空はここで潰す
        if hasattr(df, "empty") and df.empty:
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()

# ====== 本体 ======
def fetch_one_interval(tkr, interval, days, span):
    root = os.path.join(RAW, f"yahoo_{interval}", tkr)
    os.makedirs(root, exist_ok=True)
    rows = 0
    files_made = 0

    for s, e in chunks(days, span):
        # このチャンクの“未保存”日だけ対象。週末は除外。
        day_list = pd.date_range(s, e - pd.Timedelta(minutes=1), freq="D", tz="Asia/Tokyo").date
        to_get = []
        for d in day_list:
            if pd.Timestamp(d).weekday() >= 5:  # 週末スキップ
                continue
            if not os.path.exists(os.path.join(root, f"{d}.parquet")):
                to_get.append(d)
        if not to_get:
            continue

        got_any = False
        for _ in range(RETRY_EACH):
            df = safe_download(tkr, s, e, interval)
            if not df.empty:
                df = df[~df.index.duplicated(keep="last")]
                df = ensure_jst_index(df)
                df["date"] = df.index.date
                for d, g in df.groupby("date"):
                    if d not in to_get:
                        continue
                    fn = os.path.join(root, f"{d}.parquet")
                    g.drop(columns=["date"], errors="ignore").to_parquet(fn)
                    rows += len(g); files_made += 1; got_any = True
                break
            time.sleep(0.4)
        time.sleep(0.2)

        # 取れないチャンクは静かにスキップ（休日連続/欠損など）
        if not got_any:
            continue

    return rows, files_made

def fetch_all_for_ticker(tkr: str):
    out = {"ticker": tkr, "total_rows": 0, "made_files": 0, "status": "OK"}
    try:
        if not has_intraday_data(tkr):
            out["status"] = "SKIP_NO_INTRADAY"; return out

        for interval, days, span in PLAN:
            r, m = fetch_one_interval(tkr, interval, days, span)
            out["total_rows"] += r; out["made_files"] += m

        return out

    except KeyboardInterrupt:
        raise  # 上位で停止

    except Exception as e:
        out["status"] = f"ERROR:{type(e).__name__}: {e}"
        return out

def main():
    tks = load_universe()

    # まず fast_info で timezone があるものだけに絞る（軽量化）
    quick = []
    for t in tks:
        try:
            fi = yf.Ticker(t).fast_info
            if getattr(fi, "timezone", None):
                quick.append(t)
        except Exception:
            continue

    results = []
    try:
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
            fut = {ex.submit(fetch_all_for_ticker, t): t for t in quick}
            for f in as_completed(fut):
                res = f.result()
                results.append(res)
                print(res)
    except KeyboardInterrupt:
        print("KeyboardInterrupt: saving partial log and exiting...")

    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    logp = os.path.join(LOGD, f"backfill_{ts}.json")
    with open(logp, "w", encoding="utf-8") as w:
        json.dump(results, w, ensure_ascii=False, indent=2)
    print("LOG:", logp)

    df = pd.DataFrame(results)
    if not df.empty:
        ok  = (df["status"] == "OK").sum()
        skp = df["status"].str.startswith("SKIP").sum()
        err = df["status"].str.startswith("ERROR").sum()
        rows = int(df["total_rows"].sum())
        print(f"SUMMARY: OK {ok}  SKIP {skp}  ERR {err}  rows {rows}")
    else:
        print("SUMMARY: no results")

if __name__ == "__main__":
    main()


