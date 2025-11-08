from pathlib import Path
import pandas as pd

BASE = Path("output/bt30")

COMBOS = {
    "excel_AM10_j-only": BASE / "NIGHTLY_20251018" / "RUN_coarse_AM10_j-only" / "_SUMMARY_FORWARD.csv",
    "excel_AM10_j-cross": BASE / "NIGHTLY_20251018" / "RUN_coarse_AM10_j-cross" / "_SUMMARY_FORWARD.csv",
    "excel_AM15_j-only": BASE / "NIGHTLY_20251018" / "RUN_coarse_AM15_j-only" / "_SUMMARY_FORWARD.csv",
    "excel_AM15_j-cross": BASE / "NIGHTLY_20251018" / "RUN_coarse_AM15_j-cross" / "_SUMMARY_FORWARD.csv",
    "amt_AM10_j-only": BASE / "RUN_coarse_AM10_20251017_112911" / "_SUMMARY_FORWARD.csv",
    "amt_AM10_j-cross": BASE / "RUN_coarse_AM10_jcross_20251017_120341" / "_SUMMARY_FORWARD.csv",
    "amt_AM15_j-only": BASE / "RUN_coarse_AM15_20251017_114040" / "_SUMMARY_FORWARD.csv",
    "amt_AM15_j-cross": BASE / "RUN_coarse_AM15_jcross_20251017_122030" / "_SUMMARY_FORWARD.csv",
    "vol_AM10_j-only": BASE / "RUN_coarse_AM10_vol_jonly" / "_SUMMARY_FORWARD.csv",
    "vol_AM10_j-cross": BASE / "RUN_coarse_AM10_vol_jcross" / "_SUMMARY_FORWARD.csv",
    "vol_AM15_j-only": BASE / "RUN_coarse_AM15_vol_jonly" / "_SUMMARY_FORWARD.csv",
    "vol_AM15_j-cross": BASE / "RUN_coarse_AM15_vol_jcross" / "_SUMMARY_FORWARD.csv",
}

records = []

for label, path in COMBOS.items():
    if not path.exists():
        continue
    df = pd.read_csv(path)
    df = df.fillna(0)
    mask = (df["forward_trades"] >= 5) & (df["forward_pf_eff"] >= 1.2)
    filtered = df.loc[mask].copy()
    total = len(df)
    passes = len(filtered)
    if passes > 0:
        mean_pf = filtered["forward_pf_eff"].mean()
        median_pf = filtered["forward_pf_eff"].median()
        mean_trades = filtered["forward_trades"].mean()
        mean_bars = filtered.get("forward_avg_bars", pd.Series([0])).mean()
    else:
        mean_pf = float("nan")
        median_pf = float("nan")
        mean_trades = float("nan")
        mean_bars = float("nan")
    records.append(
        {
            "combo": label,
            "total_codes": total,
            "pass_count": passes,
            "pass_ratio": passes / total if total else 0.0,
            "mean_forward_pf_eff": mean_pf,
            "median_forward_pf_eff": median_pf,
            "mean_forward_trades": mean_trades,
            "mean_forward_avg_bars": mean_bars,
        }
    )

out_dir = BASE / "NIGHTLY_20251018"
out_dir.mkdir(parents=True, exist_ok=True)
out_path = out_dir / "universe_comparison_summary.csv"
pd.DataFrame(records).to_csv(out_path, index=False)
print("written", out_path)
