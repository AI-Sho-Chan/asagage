import argparse
import datetime as dt
import math
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ----------------------------- Config defaults -----------------------------
TOTAL_CAP = 100_000_000  # yen, user-specified
MIN_NOMINAL = 2_000_000  # yen per trade (default reasonable minimum)
MAX_NOMINAL = 25_000_000  # yen per trade cap (avoid single-name concentration)
FEE_BP = 4.0  # round-trip fees in bp (as per pipeline defaults)
BASE_SLIP_BP = 4.0  # base slippage bp at reference depth
SLIP_LIMIT_BP = 12.0  # target cap for dynamic sizing (tunable)
DEPTH_FACTOR = 3.0  # scales median per-minute notional to depth capacity (tunable)
SLIP_ALPHA = 1.0  # slippage grows roughly linearly with nominal/depth
FORWARD_DAYS = 5  # bt_opt30_forward default


SESSION_MINUTES: Dict[str, int] = {
    "AM0930": 30,
    "AM0945": 45,
    "AM1000": 60,
    "AM1015": 75,
    "AM1030": 90,
}


@dataclass
class Candidate:
    code: str
    session: str
    mode: str
    winrate: float
    pf_eff: float
    exp_bp: float
    exp_boot_low: float
    trades_fwd: int
    avg_bars: float


def list_refine_top(nightly_root: Path) -> List[Path]:
    paths: List[Path] = []
    for d in nightly_root.iterdir():
        if d.is_dir() and d.name.startswith("RUN_refine_"):
            f = d / "_TOP_CANDIDATES.csv"
            if f.exists():
                paths.append(f)
    return paths


def load_candidates(nightly_root: Path) -> List[Candidate]:
    cands: List[Candidate] = []
    for path in list_refine_top(nightly_root):
        try:
            df = pd.read_csv(path)
        except Exception:
            continue
        if df.empty:
            continue
        session = str(df.get("session", "").iloc[0]) if "session" in df.columns else ""
        mode = str(df.get("signal_mode", "").iloc[0]) if "signal_mode" in df.columns else ""

        need = {
            "code",
            "forward_winrate",
            "forward_pf_eff",
            "forward_exp_bp",
            "forward_exp_boot_low",
            "forward_trades",
            "forward_avg_bars",
        }
        if not need.issubset(set(df.columns)):
            # minimal fallback: try boot mean low missing
            df["forward_exp_boot_low"] = df.get("forward_exp_boot_low", df.get("forward_exp_boot_mean", df["forward_exp_bp"]))
            if not {"code", "forward_winrate", "forward_pf_eff", "forward_exp_bp", "forward_trades"}.issubset(
                set(df.columns)
            ):
                continue
            if "forward_avg_bars" not in df.columns:
                df["forward_avg_bars"] = 10.0

        for _, r in df.iterrows():
            cands.append(
                Candidate(
                    code=str(r.get("code")),
                    session=session,
                    mode=mode,
                    winrate=float(r.get("forward_winrate", 0.0) or 0.0),
                    pf_eff=float(r.get("forward_pf_eff", 0.0) or 0.0),
                    exp_bp=float(r.get("forward_exp_bp", 0.0) or 0.0),
                    exp_boot_low=float(r.get("forward_exp_boot_low", 0.0) or 0.0),
                    trades_fwd=int(r.get("forward_trades", 0) or 0),
                    avg_bars=float(r.get("forward_avg_bars", 10.0) or 10.0),
                )
            )
    return cands


def per_trade_win_loss_bp(p: float, R: float, E: float) -> Tuple[float, float]:
    """Solve avg win (W) and loss (L) from winrate p, PF_eff R, expected E (bp).

    R = (p*W)/((1-p)*L), E = p*W - (1-p)*L
    => L = E / ((1-p)*(R-1)), W = (R*(1-p)/p)*L
    Guard against degenerate R, p values.
    """
    q = 1 - p
    R_eff = max(R, 1.001)
    p_eff = min(max(p, 0.001), 0.999)
    q_eff = 1 - p_eff
    try:
        L = E / (q_eff * (R_eff - 1.0))
    except ZeroDivisionError:
        L = abs(E) if E != 0 else 1.0
    if L <= 0:
        L = abs(E) if E != 0 else 1.0
    W = (R_eff * q_eff / p_eff) * L
    return float(W), float(L)


def per_trade_var_bp(p: float, W: float, L: float, E: float) -> float:
    # outcomes: +W with prob p, -L with prob (1-p)
    return float(p * (W - E) ** 2 + (1 - p) * (-L - E) ** 2)


def minute_depth_capacity(code: str, session: str, raw_root: Path) -> float:
    """Estimate per-minute absorbable notional (yen) via median amt in session.
    capacity = DEPTH_FACTOR * median(minute_amt)
    """
    d = raw_root / f"{code}"
    if not d.exists():
        return 5_000_000.0  # fallback conservative
    end_map = {
        "AM0930": dt.time(9, 30),
        "AM0945": dt.time(9, 45),
        "AM1000": dt.time(10, 0),
        "AM1015": dt.time(10, 15),
        "AM1030": dt.time(10, 30),
    }
    end_t = end_map.get(session, dt.time(10, 30))
    am_start = dt.time(9, 0)
    am_end = end_t
    amts: List[float] = []
    for pq in d.glob("*.parquet"):
        try:
            df = pd.read_parquet(pq)
        except Exception:
            continue
        # normalize columns
        if isinstance(df.columns, pd.MultiIndex):
            level0 = [c.lower() for c in df.columns.get_level_values(0)]
            df.columns = level0
        else:
            df.columns = [c.lower() for c in df.columns]
        if not {"close", "volume"}.issubset(set(df.columns)):
            continue
        # timezone-aware index
        try:
            ts = pd.DatetimeIndex(df.index)
        except Exception:
            continue
        loc = (ts.time >= am_start) & (ts.time <= am_end)
        if not np.any(loc):
            continue
        sub = df.loc[loc]
        amt = (sub["close"] * sub["volume"]).astype(float)
        if not amt.empty:
            amts.extend(amt.tolist())
    if not amts:
        return 5_000_000.0
    med = float(np.median(amts))
    return max(1_000_000.0, DEPTH_FACTOR * med)


def slippage_bp(nominal: float, depth_cap: float) -> float:
    ratio = nominal / max(depth_cap, 1.0)
    if ratio <= 1.0:
        return BASE_SLIP_BP * ratio ** SLIP_ALPHA
    return min(50.0, BASE_SLIP_BP * ratio ** SLIP_ALPHA)


def score_candidate(c: Candidate) -> float:
    # normalize features to 0..1
    wr = max(0.0, min(1.0, (c.winrate - 0.5) / (0.8 - 0.5)))  # 0.5->0, 0.8->1
    pf = max(0.0, min(1.0, (c.pf_eff - 1.0) / (3.0 - 1.0)))   # 1.0->0, 3.0->1
    ci = max(0.0, min(1.0, (c.exp_boot_low - 0.0) / (5.0 - 0.0)))  # 0bp->0, 5bp->1
    # crude risk proxy from variance
    W, L = per_trade_win_loss_bp(max(0.0, min(1.0, c.winrate)), max(c.pf_eff, 1.001), c.exp_bp)
    var_bp = per_trade_var_bp(c.winrate, W, L, c.exp_bp)
    rp = 1.0 / (1.0 + var_bp / 1000.0)
    s = 0.35 * wr + 0.35 * pf + 0.2 * ci + 0.1 * rp
    return float(max(0.0, min(1.0, s)))


def simulate_portfolio(nightly_root: Path, raw_root: Path, total_cap: float = TOTAL_CAP,
                       dynamic: bool = False) -> Tuple[pd.DataFrame, Dict[str, float]]:
    cands = load_candidates(nightly_root)
    if not cands:
        raise SystemExit("No candidates found under " + str(nightly_root))

    rows = []
    # first pass: decide nominal preferences and occupancy fractions
    allocs: List[Dict] = []
    for c in cands:
        minutes = SESSION_MINUTES.get(c.session, 90)
        trades_per_day = (c.trades_fwd / FORWARD_DAYS) if c.trades_fwd else 0.0
        avg_bars = c.avg_bars if c.avg_bars and c.avg_bars > 0 else 10.0
        occ = min(1.0, (trades_per_day * avg_bars) / max(1.0, minutes))
        s = score_candidate(c)
        if dynamic:
            nominal_pref = MIN_NOMINAL + s * (MAX_NOMINAL - MIN_NOMINAL)
        else:
            nominal_pref = (MIN_NOMINAL + MAX_NOMINAL) / 2.0  # flat baseline

        depth_cap = minute_depth_capacity(c.code, c.session, raw_root)
        # adjust nominal to respect slippage cap in dynamic mode
        if dynamic:
            nom = float(nominal_pref)
            for _ in range(5):
                slip = slippage_bp(nom, depth_cap)
                if slip <= SLIP_LIMIT_BP:
                    break
                nom *= 0.8
            nominal = max(MIN_NOMINAL, min(MAX_NOMINAL, nom))
        else:
            nominal = max(MIN_NOMINAL, min(MAX_NOMINAL, nominal_pref))

        allocs.append(
            dict(code=c.code, session=c.session, mode=c.mode, score=s, nominal=nominal,
                 occ=occ, p=c.winrate, pf=c.pf_eff, E=c.exp_bp, trades_per_day=trades_per_day,
                 depth_cap=depth_cap, avg_bars=avg_bars)
        )

    # enforce portfolio cap using occupancy-weighted sum
    total_occ_nom = sum(a["nominal"] * a["occ"] for a in allocs)
    scale = 1.0
    if total_occ_nom > total_cap:
        scale = total_cap / total_occ_nom
    for a in allocs:
        a["nominal"] *= scale

    # compute day-level mean/variance & stats
    day_mean_yen = 0.0
    day_var_yen2 = 0.0
    total_trades_day = 0.0
    for a in allocs:
        p = max(0.0, min(1.0, a["p"]))
        R = max(a["pf"], 1.001)
        E = a["E"] - (FEE_BP + BASE_SLIP_BP)  # include fees + base slip
        W, L = per_trade_win_loss_bp(p, R, E)
        var_bp = per_trade_var_bp(p, W, L, E)
        N = a["trades_per_day"]
        notional = a["nominal"]
        # adjust slippage by depth
        slip_extra = slippage_bp(notional, a["depth_cap"]) - BASE_SLIP_BP
        slip_extra = max(0.0, slip_extra)
        E_effective = E - slip_extra

        mean_trade_yen = (E_effective / 10000.0) * notional
        var_trade_yen2 = ((math.sqrt(var_bp) / 10000.0) * notional) ** 2
        day_mean_yen += N * mean_trade_yen
        day_var_yen2 += N * var_trade_yen2
        total_trades_day += N

        rows.append(
            dict(
                code=a["code"],
                session=a["session"],
                mode=a["mode"],
                nominal=notional,
                occ=a["occ"],
                trades_per_day=N,
                mean_trade_yen=mean_trade_yen,
                var_trade_yen2=var_trade_yen2,
                depth_cap=a["depth_cap"],
                score=a["score"],
            )
        )

    sigma_yen = math.sqrt(max(0.0, day_var_yen2))
    var95 = 1.65 * sigma_yen
    summary = dict(
        day_mean_yen=day_mean_yen,
        day_sigma_yen=sigma_yen,
        day_VaR95_yen=var95,
        total_trades_day=total_trades_day,
        occ_weighted_notional=sum(a["nominal"] * a["occ"] for a in allocs),
        scale_applied=scale,
    )
    return pd.DataFrame(rows), summary


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--nightly-root", default="output/bt30/NIGHTLY_20251024", help="NIGHTLY_YYYYMMDD dir")
    ap.add_argument("--raw-1m-root", default="data/raw/yahoo_1m", help="Root of 1m parquet per ticker")
    ap.add_argument("--outdir", default="output/research/portfolio_sim", help="Output directory")
    args = ap.parse_args()

    nightly = Path(args.nightly_root)
    raw_root = Path(args.raw_1m_root)
    outdir = Path(args.outdir) / dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    outdir.mkdir(parents=True, exist_ok=True)

    base_df, base_sum = simulate_portfolio(nightly, raw_root, TOTAL_CAP, dynamic=False)
    dyn_df, dyn_sum = simulate_portfolio(nightly, raw_root, TOTAL_CAP, dynamic=True)

    base_df.to_csv(outdir / "baseline_allocations.csv", index=False, encoding="utf-8-sig")
    dyn_df.to_csv(outdir / "dynamic_allocations.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame([base_sum]).to_csv(outdir / "baseline_summary.csv", index=False)
    pd.DataFrame([dyn_sum]).to_csv(outdir / "dynamic_summary.csv", index=False)

    comp = {
        "baseline_day_mean_yen": base_sum["day_mean_yen"],
        "dynamic_day_mean_yen": dyn_sum["day_mean_yen"],
        "delta_day_mean_yen": dyn_sum["day_mean_yen"] - base_sum["day_mean_yen"],
        "baseline_VaR95_yen": base_sum["day_VaR95_yen"],
        "dynamic_VaR95_yen": dyn_sum["day_VaR95_yen"],
        "delta_VaR95_yen": dyn_sum["day_VaR95_yen"] - base_sum["day_VaR95_yen"],
        "baseline_occ_notional": base_sum["occ_weighted_notional"],
        "dynamic_occ_notional": dyn_sum["occ_weighted_notional"],
    }
    pd.DataFrame([comp]).to_csv(outdir / "comparison_summary.csv", index=False)
    print("Wrote:", outdir)


if __name__ == "__main__":
    main()
