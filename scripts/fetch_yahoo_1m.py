# -*- coding: utf-8 -*-
r"""
C:\AI\asagake\scripts\fetch_yahoo_1m.py
UNIVERSE_EXTRA!A列（上から200件）の証券コードを {4桁}.T に変換し、
Yahoo Financeの1分足を直近30日分（必ず<=7日ずつの5分割）で取得して Parquet 保存します。

保存先:
  C:\AI\asagake\data\raw\yahoo_1m\<TICKER>\YYYY-MM-DD.parquet
ログ:
  C:\AI\asagake\logs\yahoo_fetch\fetch_YYYYMMDD_HHMMSS.csv
"""

import os, sys, time, re, logging, warnings
import pandas as pd
import yfinance as yf

# ===== ログ/警告の抑制（yfinanceの "Failed download" を極力黙らせる） =====
warnings.filterwarnings("ignore")
logging.getLogger("yfinance").setLevel(logging.CRITICAL)
logging.getLogger("urllib3").setLevel(logging.CRITICAL)

# ===== 設定 =====
BASE_DIR   = r"C:\AI\asagake"
XLSM_PATH  = os.path.join(BASE_DIR, "excel", "ASAGAKE_template_30_safe.xlsm")  # 実在パスに合わせて
SHEET_NAME = "UNIVERSE_EXTRA"
COL_CODE   = 0        # A列
UNIVERSE_N = 200
INTERVAL   = "1m"     # 1分足
DAYS       = 30       # 直近30日
CHUNK_D    = 7        # 7日チャンク
RETRIES    = 2        # 軽い再試行回数
SLEEP_S    = 0.4      # リクエスト間スリープ

RAW_DIR = os.path.join(BASE_DIR, r"data\raw\yahoo_1m")
LOG_DIR = os.path.join(BASE_DIR, r"logs\yahoo_fetch")
os.makedirs(RAW_DIR, exist_ok=True)
os.makedirs(LOG_DIR, exist_ok=True)

# ===== ユーティリティ =====
def _to_jst_index(df: pd.DataFrame) -> pd.DataFrame:
    """UTC/不明TZ → JST。既にTZ付きでも安全に変換。"""
    if df.empty:
        return df
    try:
        if df.index.tz is None:
            df.index = df.index.tz_localize("UTC").tz_convert("Asia/Tokyo")
        else:
            df.index = df.index.tz_convert("Asia/Tokyo")
    except Exception:
        try:
            df.index = df.index.tz_convert("Asia/Tokyo")
        except Exception:
            df.index = df.index.tz_localize("UTC").tz_convert("Asia/Tokyo")
    return df

def load_universe_from_excel(path: str) -> list[str]:
    """UNIVERSE_EXTRA!A列→4桁コード抽出→.T付与（上位200件）。"""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel not found: {path}")
    df = pd.read_excel(path, sheet_name=SHEET_NAME, engine="openpyxl", usecols=[COL_CODE])
    codes = df.iloc[:, 0].astype(str).str.extract(r"(\d{4})")[0].dropna().tolist()
    codes = codes[:UNIVERSE_N]
    return [f"{c}.T" for c in codes]

def has_minute_data(ticker: str) -> bool:
    """
    分足が取れるか事前チェック（7日×1mを軽く叩いて空なら除外）。
    fast_info.timezone は生きていても1mが無い銘柄があるため二重チェック。
    """
    try:
        # 直近7日で軽く試す（naiveで渡す）
        end = pd.Timestamp.utcnow().floor("D") + pd.Timedelta(days=1)
        start = end - pd.Timedelta(days=7)
        df = yf.download(
            ticker, start=start, end=end, interval="1m",
            prepost=False, progress=False, auto_adjust=False
        )
        return not df.empty
    except Exception:
        return False

def fetch_1m_30d(ticker: str, retries: int = RETRIES, sleep: float = SLEEP_S) -> pd.DataFrame:
    """
    1分足を直近30日分、7日以下のチャンクに厳密分割して連結取得。
    yfinanceのstart/endはnaive（tz無し）で渡し、取得後にJSTへ統一。
    """
    if not has_minute_data(ticker):
        return pd.DataFrame()

    end_jst = pd.Timestamp.now(tz="Asia/Tokyo").floor("D") + pd.Timedelta(days=1)

    # 30→23→16→9→2→0（6点）の区切りで5チャンク（すべて<=7日）
    cut_days = (DAYS, DAYS-CHUNK_D, DAYS-2*CHUNK_D, DAYS-3*CHUNK_D, 2, 0)
    cut_points = [end_jst - pd.Timedelta(days=d) for d in cut_days]  # tz-aware(JST)

    frames = []
    for s, e in zip(cut_points[:-1], cut_points[1:]):
        s_naive = s.tz_localize(None)
        e_naive = e.tz_localize(None)
        if (e_naive - s_naive) > pd.Timedelta(days=7):
            e_naive = s_naive + pd.Timedelta(days=7) - pd.Timedelta(minutes=1)

        ok = False
        for _ in range(retries + 1):
            try:
                df = yf.download(
                    ticker, start=s_naive, end=e_naive, interval=INTERVAL,
                    prepost=False, progress=False, auto_adjust=False
                )
            except Exception:
                df = pd.DataFrame()
            if not df.empty:
                df = df[~df.index.duplicated(keep="last")]
                df = _to_jst_index(df)
                frames.append(df)
                ok = True
                break
            time.sleep(sleep)  # リトライ間隔
        time.sleep(sleep)      # チャンク間隔

    if not frames:
        return pd.DataFrame()
    out = pd.concat(frames).sort_index()
    return out[~out.index.duplicated(keep="last")]

def save_daily_parquet(df: pd.DataFrame, ticker: str) -> int:
    """1日1ファイルで保存。"""
    if df.empty:
        return 0
    df["date"] = df.index.date
    ddir = os.path.join(RAW_DIR, ticker)
    os.makedirs(ddir, exist_ok=True)
    count = 0
    for d, g in df.groupby("date"):
        fn = os.path.join(ddir, f"{d}.parquet")
        g.drop(columns=["date"], errors="ignore").to_parquet(fn, index=True)
        count += len(g)
    return count

# ===== メイン =====
def main():
    tickers = load_universe_from_excel(XLSM_PATH)
    log_rows = []
    total = len(tickers)
    for i, t in enumerate(tickers, 1):
        try:
            df = fetch_1m_30d(t)
            if df.empty:
                log_rows.append((t, 0, "SKIP_INVALID_OR_NO_1M"))
                print(f"[{i}/{total}] {t}: 0 rows (skip)")
                continue
            n = save_daily_parquet(df, t)
            log_rows.append((t, n, "OK"))
            print(f"[{i}/{total}] {t}: {n} rows")
        except Exception as e:
            log_rows.append((t, None, f"ERROR:{e}"))
            print(f"[{i}/{total}] {t}: ERROR {e}", file=sys.stderr)

    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    pd.DataFrame(log_rows, columns=["ticker", "rows", "status"]).to_csv(
        os.path.join(LOG_DIR, f"fetch_{ts}.csv"),
        index=False, encoding="utf-8-sig"
    )

if __name__ == "__main__":
    main()

