# analyze_rankings.py — 8ランキングのBT/FWDを集計し、許可銘柄を抽出
import argparse, re
from pathlib import Path
import pandas as pd
import numpy as np

RANK_RE = re.compile(r'^(NEAGARI|NESAGARI|YORI_NEAGARI|YORI_NESAGARI|DEKIDAKA|DEKIDAKA_KYUZO|DEKIDAKA_KAIRI|TICK)_(\d{8})$', re.I)

def _read_csv_safe(p: Path) -> pd.DataFrame:
    if not p.exists():
        return pd.DataFrame()
    try:
        return pd.read_csv(p)
    except Exception:
        # 文字化け対策（UTF-8 BOM / cp932）
        for enc in ("utf-8-sig","cp932","utf-8"):
            try:
                return pd.read_csv(p, encoding=enc)
            except Exception:
                pass
        return pd.read_csv(p, errors="ignore")

def _norm_col(d: pd.DataFrame, cands, default=None):
    for c in cands:
        if c in d.columns: return c
    # 大文字小文字の揺れ吸収
    low = {c.lower():c for c in d.columns}
    for c in cands:
        lc = c.lower()
        if lc in low: return low[lc]
    return default

def load_one_dir(d: Path):
    comp = _read_csv_safe(d / "_COMPARE.csv")
    back = _read_csv_safe(d / "_SUMMARY_BACK.csv")
    fwd  = _read_csv_safe(d / "_SUMMARY_FWD.csv")
    for df in (comp, back, fwd):
        if df.empty: continue
        df["ranking"] = d.name.split("_")[0].upper()
        df["stamp"]   = d.name.split("_")[-1]
    return comp, back, fwd

def summarize_per_ranking(all_comp: pd.DataFrame) -> pd.DataFrame:
    # 列名を捕捉
    c_code = _norm_col(all_comp, ["code","ticker","symbol"], "code")
    cols = dict(
        wr_b = _norm_col(all_comp, ["winrate_BACK"], "winrate_BACK"),
        wr_f = _norm_col(all_comp, ["winrate_FWD"], "winrate_FWD"),
        pf_b = _norm_col(all_comp, ["PF_eff_BACK","pf_BACK"], "PF_eff_BACK"),
        pf_f = _norm_col(all_comp, ["PF_eff_FWD","pf_FWD"], "PF_eff_FWD"),
        tr_b = _norm_col(all_comp, ["trades_BACK"], "trades_BACK"),
        tr_f = _norm_col(all_comp, ["trades_FWD"], "trades_FWD"),
    )
    g = all_comp.groupby("ranking", as_index=False).agg({
        c_code:"nunique",
        cols["wr_b"]:"median", cols["wr_f"]:"median",
        cols["pf_b"]:"median", cols["pf_f"]:"median",
        cols["tr_b"]:"sum",   cols["tr_f"]:"sum"
    })
    g = g.rename(columns={
        c_code:"n_codes",
        cols["wr_b"]:"WR_BACK_med", cols["wr_f"]:"WR_FWD_med",
        cols["pf_b"]:"PF_BACK_med", cols["pf_f"]:"PF_FWD_med",
        cols["tr_b"]:"TR_BACK_sum", cols["tr_f"]:"TR_FWD_sum",
    }).sort_values(["PF_FWD_med","WR_FWD_med","TR_FWD_sum"], ascending=[False,False,False])
    return g

def pick_allow(all_comp: pd.DataFrame,
               min_tr=8, min_wr=0.50, min_pf=1.15,
               relax=False) -> pd.DataFrame:
    c_code = _norm_col(all_comp, ["code","ticker","symbol"], "code")
    wr_f = _norm_col(all_comp, ["winrate_FWD"], "winrate_FWD")
    pf_f = _norm_col(all_comp, ["PF_eff_FWD","pf_FWD"], "PF_eff_FWD")
    tr_f = _norm_col(all_comp, ["trades_FWD"], "trades_FWD")
    # パラメタ列
    atr = _norm_col(all_comp, ["ATR_n_FWD","ATR_n_BACK","ATR_n"], "ATR_n_BACK")
    tpk = _norm_col(all_comp, ["TPk_FWD","TPk_BACK","TPk"], "TPk_BACK")
    slk = _norm_col(all_comp, ["SLk_FWD","SLk_BACK","SLk"], "SLk_BACK")
    jth = _norm_col(all_comp, ["J_th_FWD","J_th_BACK","J_th"], "J_th_BACK")
    dj  = _norm_col(all_comp, ["dJ_th_FWD","dJ_th_BACK","dJ_th"], "dJ_th_BACK")
    vth = _norm_col(all_comp, ["vEMA_th_FWD","vEMA_th_BACK","vEMA_th"], "vEMA_th_BACK")

    df = all_comp.copy()
    df["ok"] = (df[tr_f] >= min_tr) & (df[wr_f] >= min_wr) & (df[pf_f] >= min_pf)
    out = df.loc[df["ok"], [c_code, "ranking", wr_f, pf_f, tr_f, atr, tpk, slk, jth, dj, vth]].copy()
    out = out.rename(columns={
        c_code:"code", wr_f:"winrate", pf_f:"PF_eff", tr_f:"trades",
        atr:"ATR_n", tpk:"TPk", slk:"SLk", jth:"J_th", dj:"dJ_th", vth:"vEMA_th"
    })
    if not out.empty:
        # 重複コードは FWD PF優先
        out = (out.sort_values(["code","PF_eff","winrate","trades"], ascending=[True,False,False,False])
                  .drop_duplicates("code", keep="first"))
        return out

    if relax:
        # 少なくとも候補を返す（緩和）
        df["_score"] = df[wr_f].fillna(0)*0.6 + (df[pf_f].fillna(0)/2.0)*0.4
        out = (df.sort_values(["_score", tr_f], ascending=[False,False])
                 .drop_duplicates(c_code,"first")
                 .head(30))
        return out.rename(columns={c_code:"code"})
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--bt-root", default=r"C:\AI\asagake\data\bt")
    ap.add_argument("--relax", action="store_true", help="0件回避の緩和抽出も行う")
    args = ap.parse_args()

    root = Path(args.bt_root)
    dirs = [d for d in root.iterdir() if d.is_dir() and RANK_RE.match(d.name)]
    if not dirs:
        raise SystemExit(f"no labeled folders under {root}")

    comps, backs, fwds = [], [], []
    for d in dirs:
        comp, back, fwd = load_one_dir(d)
        if not comp.empty: comps.append(comp)
        if not back.empty: backs.append(back)
        if not fwd.empty:  fwds.append(fwd)

    all_comp = pd.concat(comps, ignore_index=True) if comps else pd.DataFrame()
    if all_comp.empty:
        raise SystemExit("no _COMPARE.csv found")

    # 保存先
    out_dir = root / "ALL_RANKINGS"
    out_dir.mkdir(exist_ok=True)

    # 1) ランキング別サマリ
    rank_sum = summarize_per_ranking(all_comp)
    rank_sum.to_csv(out_dir / "RANKING_SUMMARY.csv", index=False)

    # 2) 全明細
    all_comp.to_csv(out_dir / "ALL_COMPARE.csv", index=False)

    # 3) 許可銘柄（厳格）
    allow = pick_allow(all_comp, min_tr=8, min_wr=0.50, min_pf=1.15, relax=False)
    allow.to_csv(out_dir / "ALLOW_STRICT.csv", index=False)

    # 4) 許可銘柄（緩和）
    allow2 = pick_allow(all_comp, min_tr=6, min_wr=0.48, min_pf=1.10, relax=True)
    allow2.to_csv(out_dir / "ALLOW_RELAX.csv", index=False)

    # 5) 簡易インサイト（テキスト）
    with open(out_dir / "INSIGHTS.txt", "w", encoding="utf-8") as f:
        f.write("== RANKING SUMMARY (median, sums) ==\n")
        f.write(rank_sum.to_string(index=False))
        f.write("\n\n== NOTES ==\n")
        f.write("・実運用は PF_FWD_med と TR_FWD_sum の高いランキングを優先。\n")
        f.write("・ALLOW_STRICT が空なら ALLOW_RELAX で暫定運用し、日々再評価。\n")
        f.write("・銘柄は trades_FWD が少なすぎる高PFを避け、安定性重視で選定。\n")

if __name__ == "__main__":
    main()
