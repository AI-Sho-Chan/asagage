import argparse
import glob
import json
import os
import re
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

SRC = (Path(__file__).resolve().parents[1] / "src").resolve()
if SRC.exists():
    sys.path.insert(0, str(SRC))

import numpy as np
import pandas as pd

from asagake_core.candidates import append_candidate_metadata, make_candidate_metadata_defaults

ROOT = Path("output/excel")
LOG_DIR = Path("logs")
# Guardrail for production stability:
# When aggregation would shrink the dashboard input to a tiny set, keep the previous
# "last good" set instead. This avoids accidental wipes when upstream outputs are
# missing/partial (e.g., task failures, transient file locks).
FALLBACK_MIN_ROWS_DEFAULT = 10


@dataclass
class Thresholds:
    min_j: float
    min_win_ci: float
    min_pf: float
    min_trades: int
    win_power: float
    dd_scale: float
    default_gapban_pct: float
    default_no_trade_min: int


@dataclass
class AggregateSummary:
    rows_in: int
    rows_filtered_out: int
    unique_tickers: int
    avg_forward_pf: Optional[float]
    avg_forward_winrate: Optional[float]
    avg_expected_bp: Optional[float]

    def to_json(self) -> Dict[str, object]:
        return {
            "rows_in": self.rows_in,
            "rows_filtered_out": self.rows_filtered_out,
            "unique_tickers": self.unique_tickers,
            "avg_forward_pf": self.avg_forward_pf,
            "avg_forward_winrate": self.avg_forward_winrate,
            "avg_expected_bp": self.avg_expected_bp,
        }


def collect_candidate_paths(date_tag: Optional[str] = None) -> List[Path]:
    """Return candidate CSV paths sorted oldest→newest so later files win.

    When date_tag is provided (YYYYMMDD), only reads `output/excel/NIGHTLY_YYYYMMDD/*/candidates_*.csv`.
    """
    # When aggregating a specific NIGHTLY folder, only keep the newest
    # `candidates_*.csv` per plan directory. This avoids mixing rows from
    # restarts/re-runs that leave multiple candidate snapshots under the same plan.
    if date_tag:
        root = ROOT / f"NIGHTLY_{date_tag}"
        if not root.exists():
            return []

        picked: List[Path] = []
        for plan_dir in sorted(p for p in root.iterdir() if p.is_dir()):
            candidates = sorted(plan_dir.glob("candidates_*.csv"))
            if not candidates:
                continue

            def sort_key(path: Path) -> Tuple[str, float, str]:
                tag = _date_from_name(path) or ""
                try:
                    mtime = path.stat().st_mtime
                except OSError:
                    mtime = 0.0
                return (tag, mtime, path.name)

            best = max(candidates, key=sort_key)
            name_upper = best.name.upper()
            if name_upper.startswith("CANDIDATES_NEXTDAY"):
                continue
            if name_upper.startswith("WEEKLY_CANDIDATES_"):
                continue
            picked.append(best)

        return sorted(picked)

    patterns: List[Path] = [
        ROOT / "NIGHTLY_*" / "*" / "candidates_*.csv",
        ROOT / "candidates_*.csv",
    ]

    paths: List[Path] = []
    for pat in patterns:
        for path_str in glob.glob(str(pat)):
            p = Path(path_str)
            name_upper = p.name.upper()
            # Never feed the aggregated outputs back into the aggregation.
            if name_upper.startswith("CANDIDATES_NEXTDAY"):
                continue
            if name_upper.startswith("WEEKLY_CANDIDATES_"):
                continue
            paths.append(p)

    return sorted(paths)


def resolve_latest_date_tag() -> str:
    """Pick the newest NIGHTLY_YYYYMMDD folder date for default aggregation."""
    candidates: List[str] = []
    for p in ROOT.glob("NIGHTLY_*"):
        m = re.match(r"^NIGHTLY_(\d{8})$", p.name)
        if m:
            candidates.append(m.group(1))
    return max(candidates) if candidates else ""


def _date_from_name(path: Path) -> Optional[str]:
    m = re.search(r"(\d{8})", path.name)
    return m.group(1) if m else None


def _resolve_latest_root_candidate_file() -> Optional[Path]:
    """Pick the newest single-file candidate set in output/excel/.

    Used as a fallback when NIGHTLY folders exist but contain no candidate rows.
    Preference order:
      1) candidates_YYYYMMDD_M3.csv
      2) candidates_for_YYYYMMDD.csv
      3) candidates_YYYYMMDD_M0.csv
      4) candidates_YYYYMMDD.csv (rare)

    NOTE: We explicitly avoid candidates_nextday*.csv and backups.
    """
    patterns: List[str] = [
        "candidates_*_M3.csv",
        "candidates_for_*.csv",
        "candidates_*_M0.csv",
        "candidates_*.csv",
    ]

    best: Optional[Path] = None
    best_date: str = ""

    for pat in patterns:
        for p in sorted(ROOT.glob(pat)):
            name_upper = p.name.upper()
            if name_upper.startswith("CANDIDATES_NEXTDAY"):
                continue
            if "BACKUP" in name_upper:
                continue
            if name_upper.startswith("WEEKLY_CANDIDATES_"):
                continue
            date = _date_from_name(p) or ""
            if date and date > best_date:
                best = p
                best_date = date
    return best


def _resolve_latest_b_candidate_file() -> Optional[Path]:
    """Pick the newest B-type candidates_nextday CSV in output/excel/.

    These files are excluded from normal aggregation (to avoid feeding the
    aggregated output back into the input), but they are useful as a safe
    fallback when nightly aggregation yields too few rows.

    Preference order:
      1) candidates_nextday_B_coarse3.csv (fixed name)
      2) candidates_nextday_B_from_vm_YYYYMMDD.csv (dated snapshots)
      3) candidates_nextday_B_*.csv (other variants)
    """

    coarse3 = ROOT / "candidates_nextday_B_coarse3.csv"
    if coarse3.is_file():
        return coarse3

    best: Optional[Path] = None
    best_date: str = ""

    for p in sorted(ROOT.glob("candidates_nextday_B_from_vm_*.csv")):
        if not p.is_file():
            continue
        date = _date_from_name(p) or ""
        if date and date > best_date:
            best = p
            best_date = date

    if best is not None:
        return best

    # Last resort: other B variants. Prefer date when present; otherwise newest mtime.
    best_mtime: float = 0.0
    for p in sorted(ROOT.glob("candidates_nextday_B_*.csv")):
        if not p.is_file():
            continue
        date = _date_from_name(p) or ""
        try:
            mtime = p.stat().st_mtime
        except OSError:
            mtime = 0.0

        if date:
            if date > best_date:
                best = p
                best_date = date
                best_mtime = mtime
        else:
            if not best_date and mtime >= best_mtime:
                best = p
                best_mtime = mtime

    return best


def _read_nonempty_csvs(paths: Iterable[Path]) -> Tuple[List[pd.DataFrame], List[Path]]:
    frames: List[pd.DataFrame] = []
    used: List[Path] = []
    for path in paths:
        try:
            df = pd.read_csv(path)
        except Exception:
            continue
        if df.empty:
            continue
        frames.append(df)
        used.append(path)
    return frames, used


def _column_lookup(df: pd.DataFrame) -> Dict[str, str]:
    return {c.lower(): c for c in df.columns}


def _col(cols: Dict[str, str], name: str) -> Optional[str]:
    return cols.get(name.lower())


def _num(df: pd.DataFrame, column: Optional[str]) -> pd.Series:
    if column is None or column not in df.columns:
        return pd.Series(0.0, index=df.index)
    return pd.to_numeric(df[column], errors="coerce").fillna(0.0)


def _avg(df: pd.DataFrame, column: Optional[str]) -> Optional[float]:
    if not column or column not in df.columns:
        return None
    series = pd.to_numeric(df[column], errors="coerce").dropna()
    if series.empty:
        return None
    return float(series.mean())


def _ensure_live_demo_fields(df: pd.DataFrame) -> pd.DataFrame:
    cols = _column_lookup(df)
    pf_col = _col(cols, "forward_pf_eff")
    win_col = _col(cols, "forward_winrate")
    trades_col = _col(cols, "forward_trades")

    if not (pf_col and win_col and trades_col):
        df["BudgetFactor_row"] = 1.0
        df["live_demo_class"] = "LIVE_BASE"
        return df

    pf = _num(df, pf_col).clip(lower=1.0, upper=5.0)
    win = _num(df, win_col)
    trades = _num(df, trades_col)

    live_strong = (trades >= 30) & (win >= 0.7) & (pf >= 1.8)
    live_base = (trades >= 15) & (win >= 0.6) & (pf >= 1.3) & ~live_strong

    df["BudgetFactor_row"] = np.where(live_strong, 2.0, np.where(live_base, 1.0, 0.5))
    df["live_demo_class"] = np.where(
        live_strong, "LIVE_STRONG", np.where(live_base, "LIVE_BASE", "DEMO_ONLY")
    )
    return df


def _ensure_allowed_side_fields(df: pd.DataFrame) -> pd.DataFrame:
    cols = _column_lookup(df)
    nky_col = cols.get("nky_allowedside") or cols.get("nky_allowed_side")
    topix_col = cols.get("topix_allowedside") or cols.get("topix_allowed_side")

    df["NKY_AllowedSide"] = df[nky_col].fillna("BOTH") if nky_col else "BOTH"
    df["TOPIX_AllowedSide"] = df[topix_col].fillna("BOTH") if topix_col else "BOTH"
    return df


def _ensure_gapban_notrade_fields(df: pd.DataFrame, thresholds: Thresholds) -> pd.DataFrame:
    cols = _column_lookup(df)
    gap_col = (
        cols.get("gapbanpct")
        or cols.get("gapbanpc")
        or cols.get("gap_ban_pct")
        or cols.get("gap_ban_pc")
    )
    notrade_col = cols.get("notrademin") or cols.get("no_trade_min") or cols.get("notrade_min")

    df["GapBanPct"] = _num(df, gap_col) if gap_col else float(thresholds.default_gapban_pct)
    df["NoTradeMin"] = _num(df, notrade_col) if notrade_col else float(thresholds.default_no_trade_min)
    return df


def _ensure_tp_sl_trail_fields(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure per-row TP/SL/Trail multipliers exist for the dashboard.

    Some candidate sources (e.g. *_M0/M3 condensed exports) may not include the
    newer `*_per_J_row` columns. In that case, we fall back to TPk/SLk which
    historically match the dashboard's per-row multipliers.
    """

    cols = _column_lookup(df)
    tpk_col = _col(cols, "TPk")
    slk_col = _col(cols, "SLk")

    tp_row_col = _col(cols, "TP_per_J_row")
    sl_row_col = _col(cols, "SL_per_J_row")
    trail_row_col = _col(cols, "Trail_per_J_row")

    # Build fallback series (may be all-zeros when missing).
    tp_fallback = _num(df, tpk_col)
    sl_fallback = _num(df, slk_col)
    trail_fallback = _num(df, slk_col)

    def fill_or_fallback(col: Optional[str], fallback: pd.Series, out_name: str) -> None:
        if col and col in df.columns:
            current = pd.to_numeric(df[col], errors="coerce")
            missing = current.isna() | (current <= 0)
            if missing.any():
                current = current.fillna(0.0)
                current.loc[missing] = fallback.loc[missing]
            df[out_name] = current
        else:
            df[out_name] = fallback

    fill_or_fallback(tp_row_col, tp_fallback, "TP_per_J_row")
    fill_or_fallback(sl_row_col, sl_fallback, "SL_per_J_row")
    fill_or_fallback(trail_row_col, trail_fallback, "Trail_per_J_row")
    return df


def aggregate_frames(frames: Iterable[pd.DataFrame], thresholds: Thresholds) -> Optional[pd.DataFrame]:
    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        return None

    df = pd.concat(frames, ignore_index=True)
    cols = _column_lookup(df)

    if "maxdd" not in cols and "max_dd" in cols:
        df = df.rename(columns={cols["max_dd"]: "MaxDD"})
        cols = _column_lookup(df)
    if "MaxDD" not in df.columns:
        df["MaxDD"] = 0.0

    j_col = _col(cols, "J_th")
    if j_col:
        df = df[_num(df, j_col) >= thresholds.min_j]

    win_ci_col = _col(cols, "forward_win_ci_low")
    if win_ci_col:
        df = df[_num(df, win_ci_col) >= thresholds.min_win_ci]

    pf_col = _col(cols, "forward_pf_eff")
    if pf_col:
        df = df[_num(df, pf_col) >= thresholds.min_pf]

    trades_col = _col(cols, "forward_trades")
    if trades_col:
        df = df[_num(df, trades_col) >= thresholds.min_trades]

    if df.empty:
        return df

    cols = _column_lookup(df)
    pf = _num(df, _col(cols, "forward_pf_eff"))
    win = _num(df, _col(cols, "forward_winrate"))
    trades = _num(df, _col(cols, "forward_trades"))
    dd = _num(df, "MaxDD")
    score = pf * np.power(win, thresholds.win_power) * np.log1p(trades) / (1.0 + (dd / thresholds.dd_scale))
    df["_score"] = score

    df = _ensure_live_demo_fields(df)
    df = _ensure_allowed_side_fields(df)
    df = _ensure_gapban_notrade_fields(df, thresholds)
    df = _ensure_tp_sl_trail_fields(df)

    cols = _column_lookup(df)
    ticker_col = _col(cols, "ticker")
    if ticker_col:
        df = df.sort_values([ticker_col, "_score"], ascending=[True, False])
    else:
        df = df.sort_values("_score", ascending=False)

    df.drop(columns=["_score"], inplace=True, errors="ignore")
    return df


def build_summary(df: pd.DataFrame, rows_before: int) -> AggregateSummary:
    cols = _column_lookup(df)
    ticker_col = _col(cols, "ticker")
    return AggregateSummary(
        rows_in=rows_before,
        rows_filtered_out=rows_before - len(df),
        unique_tickers=len(df[ticker_col].unique()) if ticker_col and ticker_col in df.columns else len(df),
        avg_forward_pf=_avg(df, _col(cols, "forward_pf_eff")),
        avg_forward_winrate=_avg(df, _col(cols, "forward_winrate")),
        avg_expected_bp=_avg(df, _col(cols, "forward_exp_boot_mean")),
    )


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Aggregate candidates into candidates_nextday.csv with quality filters."
    )
    ap.add_argument("--output", type=Path, default=ROOT / "candidates_nextday.csv")
    ap.add_argument(
        "--date-tag",
        type=str,
        default="",
        help="Optional YYYYMMDD to aggregate only that NIGHTLY folder (e.g. 20251212).",
    )
    ap.add_argument("--min-j", type=float, default=0.8)
    ap.add_argument("--min-win-ci", type=float, default=0.70)
    ap.add_argument("--min-pf", type=float, default=1.30)
    ap.add_argument("--min-trades", type=int, default=5)
    ap.add_argument("--win-power", type=float, default=1.2)
    ap.add_argument("--dd-scale", type=float, default=1000.0)
    ap.add_argument("--default-gapban-pct", type=float, default=3.0)
    ap.add_argument("--default-no-trade-min", type=int, default=5)
    ap.add_argument(
        "--fallback-min-rows",
        type=int,
        default=FALLBACK_MIN_ROWS_DEFAULT,
        help=(
            "When the aggregated candidates_nextday.csv would contain too few rows, "
            "prefer a B-type candidates file (e.g. candidates_nextday_B_coarse3.csv) "
            "if it exists and contains at least this many rows."
        ),
    )
    ap.add_argument(
        "--allow-empty-overwrite",
        action="store_true",
        help=(
            "Allow overwriting the output CSV with an empty file when no candidates are available. "
            "By default we keep the previous non-empty candidates_nextday.csv to avoid wiping the dashboard input."
        ),
    )
    ap.add_argument(
        "--diag-json",
        type=Path,
        default=None,
        help="Optional path to write a JSON diagnostic record (default: logs/aggregate_candidates_*.json).",
    )
    return ap.parse_args()


def _timestamp() -> str:
    return time.strftime("%Y%m%d_%H%M%S", time.localtime())


def _git_short_sha() -> Optional[str]:
    try:
        import subprocess

        out = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=Path(__file__).resolve().parents[1],
            stderr=subprocess.DEVNULL,
        )
        return out.decode("utf-8", errors="ignore").strip() or None
    except Exception:
        return None


def _read_nonempty_row_count(path: Path) -> int:
    if not path.exists() or not path.is_file():
        return 0
    try:
        df = pd.read_csv(path)
    except Exception:
        return 0
    return int(len(df))


def _atomic_write_csv(df: pd.DataFrame, out: Path) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    tmp = out.with_name(f".{out.name}.{os.getpid()}.{_timestamp()}.tmp")
    df.to_csv(tmp, index=False, encoding="utf-8-sig", lineterminator="\r\n")
    os.replace(tmp, out)


def _backup_existing(out: Path, keep: int = 10) -> Optional[Path]:
    if not out.exists() or not out.is_file():
        return None

    base = out.name
    backup = out.with_name(f"{base}.backup_{_timestamp()}")
    try:
        backup.write_bytes(out.read_bytes())
    except OSError:
        return None

    backups = sorted(out.parent.glob(f"{base}.backup_*"), key=lambda p: p.name, reverse=True)
    for old in backups[keep:]:
        try:
            old.unlink()
        except OSError:
            pass
    return backup


def _restore_from_backups(out: Path, min_rows: int) -> Optional[Path]:
    """Restore candidates_nextday.csv from a newer snapshot that has enough rows.

    We support both the current backup naming (`candidates_nextday.csv.backup_...`)
    and the legacy naming (`candidates_nextday_backup_...csv`).
    """
    if min_rows <= 0:
        return None
    if not out.exists() or not out.is_file():
        return None

    last_good = out.with_name(f"{out.stem}_last_good{out.suffix}")
    candidates: List[Path] = []
    if last_good.is_file():
        candidates.append(last_good)

    # Current scheme: "candidates_nextday.csv.backup_YYYYMMDD_HHMMSS"
    candidates.extend(out.parent.glob(f"{out.name}.backup_*"))

    # Legacy scheme: "candidates_nextday_backup_YYYYMMDD_HHMMSS.csv"
    candidates.extend(out.parent.glob(f"{out.stem}_backup_*{out.suffix}"))

    # Sort newest-first using (date tag if present, mtime, name).
    def sort_key(path: Path) -> Tuple[str, float, str]:
        date = _date_from_name(path) or ""
        try:
            mtime = path.stat().st_mtime
        except OSError:
            mtime = 0.0
        return (date, mtime, path.name)

    backups = sorted({p.resolve() for p in candidates if p.is_file()}, key=sort_key, reverse=True)
    for backup in backups:
        if _read_nonempty_row_count(backup) < min_rows:
            continue
        try:
            tmp = out.with_name(f".{out.name}.restore.{os.getpid()}.{_timestamp()}.tmp")
            tmp.write_bytes(backup.read_bytes())
            os.replace(tmp, out)
            return backup
        except OSError:
            continue
    return None


def _write_last_good_snapshot(out: Path) -> Optional[Path]:
    """Keep a stable copy that we can restore from when sources go missing."""
    if not out.exists() or not out.is_file():
        return None
    last_good = out.with_name(f"{out.stem}_last_good{out.suffix}")
    try:
        tmp = last_good.with_name(f".{last_good.name}.{os.getpid()}.{_timestamp()}.tmp")
        tmp.write_bytes(out.read_bytes())
        os.replace(tmp, last_good)
        return last_good
    except OSError:
        return None


def _write_diag(path: Path, payload: Dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def main() -> None:
    args = parse_args()
    thresholds = Thresholds(
        min_j=args.min_j,
        min_win_ci=args.min_win_ci,
        min_pf=args.min_pf,
        min_trades=args.min_trades,
        win_power=args.win_power,
        dd_scale=args.dd_scale,
        default_gapban_pct=args.default_gapban_pct,
        default_no_trade_min=args.default_no_trade_min,
    )

    requested_date_tag = args.date_tag.strip()
    nightly_date_tag = resolve_latest_date_tag()
    latest_root = _resolve_latest_root_candidate_file()
    latest_root_date = _date_from_name(latest_root) if latest_root else None
    latest_b = _resolve_latest_b_candidate_file()
    fallback_min_rows = max(int(getattr(args, "fallback_min_rows", FALLBACK_MIN_ROWS_DEFAULT)), 1)
    latest_b_rows = _read_nonempty_row_count(latest_b) if latest_b else 0

    # Default behaviour: pick the newest run we have. When nightly folders are stale
    # (e.g. nightly is disabled but weekend candidates_for_YYYYMMDD.csv is new),
    # we prefer the newest single-file candidate export to avoid clobbering
    # candidates_nextday.csv with an old NIGHTLY folder.
    prefer_root = False
    if not requested_date_tag and latest_root_date:
        if not nightly_date_tag or latest_root_date > nightly_date_tag:
            prefer_root = True

    date_tag = requested_date_tag or (latest_root_date or nightly_date_tag)

    source_label = "all"
    frames: List[pd.DataFrame] = []
    used_paths: List[Path] = []
    nightly_tag_to_use = ""
    if requested_date_tag:
        nightly_tag_to_use = requested_date_tag
    elif not prefer_root and nightly_date_tag:
        nightly_tag_to_use = nightly_date_tag

    if nightly_tag_to_use:
        nightly_paths = collect_candidate_paths(nightly_tag_to_use)
        frames, used_paths = _read_nonempty_csvs(nightly_paths)
        source_label = f"NIGHTLY_{nightly_tag_to_use}"

    if not frames:
        fallback = latest_root
        if fallback is not None:
            frames, used_paths = _read_nonempty_csvs([fallback])
            source_label = fallback.name

    if not frames and latest_b is not None and latest_b_rows >= fallback_min_rows:
        frames, used_paths = _read_nonempty_csvs([latest_b])
        if frames:
            source_label = latest_b.name

    out = args.output

    diag: Dict[str, object] = {
        "ts": _timestamp(),
        "output": str(out),
        "date_tag": date_tag,
        "selected_source": source_label,
        "prefer_root": prefer_root,
        "requested_date_tag": requested_date_tag or None,
        "latest_nightly_date_tag": nightly_date_tag or None,
        "latest_root_candidate": str(latest_root) if latest_root else None,
        "latest_b_candidate": str(latest_b) if latest_b else None,
        "latest_b_rows": latest_b_rows,
        "fallback_min_rows": fallback_min_rows,
        "thresholds": {
            "min_j": thresholds.min_j,
            "min_win_ci": thresholds.min_win_ci,
            "min_pf": thresholds.min_pf,
            "min_trades": thresholds.min_trades,
            "win_power": thresholds.win_power,
            "dd_scale": thresholds.dd_scale,
            "default_gapban_pct": thresholds.default_gapban_pct,
            "default_no_trade_min": thresholds.default_no_trade_min,
        },
        "source": source_label,
        "source_files": [str(p) for p in used_paths],
    }

    diag_path = args.diag_json or (LOG_DIR / f"aggregate_candidates_{_timestamp()}.json")

    if not frames:
        if args.allow_empty_overwrite:
            existing_rows = _read_nonempty_row_count(out)
            diag["result"] = {
                "reason": "no_source_candidates",
                "existing_rows": existing_rows,
                "action": "overwrite_empty",
            }
            _write_diag(diag_path, diag)
            _backup_existing(out)
            _atomic_write_csv(pd.DataFrame(), out)
            print(json.dumps({"written": str(out), "rows": 0, "kept_previous": False}, ensure_ascii=False))
        else:
            existing_rows = _read_nonempty_row_count(out)
            diag["result"] = {
                "reason": "no_source_candidates",
                "existing_rows": existing_rows,
                "action": "keep_previous",
            }

            restored_from: Optional[Path] = None
            if existing_rows < fallback_min_rows:
                restored_from = _restore_from_backups(out, fallback_min_rows)
                if restored_from is not None:
                    existing_rows = _read_nonempty_row_count(out)
                    diag["result"]["action"] = "restore_backup"
                    diag["result"]["restored_from"] = str(restored_from)

            _write_diag(diag_path, diag)
            print(
                json.dumps(
                    {
                        "written": str(out),
                        "rows": existing_rows,
                        "kept_previous": restored_from is None,
                        "restored_from": str(restored_from) if restored_from is not None else None,
                        "message": (
                            "No candidates found; restored candidates_nextday.csv from backup"
                            if restored_from is not None
                            else "No candidates found; keeping previous candidates_nextday.csv"
                        ),
                    },
                    ensure_ascii=False,
                )
            )
        return

    combined = aggregate_frames(frames, thresholds)

    if combined is None or combined.empty:
        # If NIGHTLY candidates exist but yield no usable rows, fall back to the newest
        # single-file candidate set (typically weekend-derived) so Import Candidates
        # has something reasonable to work with.
        if source_label.startswith("NIGHTLY_"):
            fallback = _resolve_latest_root_candidate_file()
            if fallback is not None:
                alt_frames, alt_paths = _read_nonempty_csvs([fallback])
                alt_combined = aggregate_frames(alt_frames, thresholds) if alt_frames else None
                if alt_combined is not None and not alt_combined.empty:
                    frames = alt_frames
                    used_paths = alt_paths
                    combined = alt_combined
                    source_label = fallback.name

    # If the aggregation results in too few rows, prefer the latest B-type candidate set.
    if (
        combined is not None
        and not combined.empty
        and len(combined) < fallback_min_rows
        and latest_b is not None
        and latest_b_rows >= fallback_min_rows
    ):
        try:
            b_df = pd.read_csv(latest_b)
        except Exception:
            b_df = pd.DataFrame()
        if not b_df.empty and len(b_df) >= fallback_min_rows:
            frames = [b_df]
            used_paths = [latest_b]
            combined = b_df
            source_label = latest_b.name
            diag["fallback_b_used"] = True
            diag["fallback_b_reason"] = "aggregated_rows_below_threshold"

    # Guardrail: if we would shrink candidates_nextday.csv to an unusually small set,
    # keep the previous file and/or restore from a last-good snapshot instead of overwriting.
    #
    # Why: when combined candidates briefly drop to 0-1 rows (e.g. transient filter/parse issue),
    # Excel's "Import Candidates" appears to load only 1 ticker/plan (or none). Worse, once the
    # output file shrinks below the threshold, subsequent runs may keep overwriting the tiny file,
    # because the "keep previous" condition no longer applies. This block prevents that.
    if combined is not None and not combined.empty and len(combined) < fallback_min_rows:
        existing_rows = _read_nonempty_row_count(out)
        restored_from: Optional[Path] = None

        if existing_rows < fallback_min_rows:
            restored_from = _restore_from_backups(out, fallback_min_rows)
            if restored_from is not None:
                existing_rows = _read_nonempty_row_count(out)
                diag["result"] = {
                    "reason": "aggregated_rows_below_threshold",
                    "aggregated_rows": int(len(combined)),
                    "existing_rows": existing_rows,
                    "action": "restore_backup",
                    "restored_from": str(restored_from),
                }
                diag["source_after_fallback"] = source_label
                diag["source_files_after_fallback"] = [str(p) for p in used_paths]
                _write_diag(diag_path, diag)
                print(
                    json.dumps(
                        {
                            "written": str(out),
                            "rows": existing_rows,
                            "kept_previous": False,
                            "restored_from": str(restored_from),
                            "message": (
                                f"Aggregated rows below threshold ({len(combined)}<{fallback_min_rows}); restored candidates_nextday.csv from backup"
                            ),
                            "source": source_label,
                            "source_files": [str(p) for p in used_paths],
                        },
                        ensure_ascii=False,
                    )
                )
                return

            # If we couldn't restore a safe snapshot, do NOT overwrite the output with a tiny set.
            # Keeping the current file (even if small) is safer than writing an unreliable partial
            # result that can confuse Excel Import and make recovery harder.
            diag["result"] = {
                "reason": "aggregated_rows_below_threshold",
                "aggregated_rows": int(len(combined)),
                "existing_rows": existing_rows,
                "action": "keep_previous_no_backup",
                "restored_from": None,
            }
            diag["warning"] = (
                "restore_from_backups_failed; keeping existing candidates_nextday.csv to avoid overwriting with a tiny set"
            )
            diag["source_after_fallback"] = source_label
            diag["source_files_after_fallback"] = [str(p) for p in used_paths]
            _write_diag(diag_path, diag)
            print(
                json.dumps(
                    {
                        "written": str(out),
                        "rows": existing_rows,
                        "kept_previous": True,
                        "message": (
                            f"Aggregated rows below threshold ({len(combined)}<{fallback_min_rows}); "
                            "no last_good/backup snapshot could be restored, so keeping existing candidates_nextday.csv"
                        ),
                        "source": source_label,
                        "source_files": [str(p) for p in used_paths],
                    },
                    ensure_ascii=False,
                )
            )
            return

        if existing_rows >= fallback_min_rows:
            diag["result"] = {
                "reason": "aggregated_rows_below_threshold",
                "aggregated_rows": int(len(combined)),
                "existing_rows": existing_rows,
                "action": "keep_previous",
            }
            diag["source_after_fallback"] = source_label
            diag["source_files_after_fallback"] = [str(p) for p in used_paths]
            _write_diag(diag_path, diag)
            print(
                json.dumps(
                    {
                        "written": str(out),
                        "rows": existing_rows,
                        "kept_previous": True,
                        "message": f"Aggregated rows below threshold ({len(combined)}<{fallback_min_rows}); keeping previous candidates_nextday.csv",
                        "source": source_label,
                        "source_files": [str(p) for p in used_paths],
                    },
                    ensure_ascii=False,
                )
            )
            return

    if combined is None or combined.empty:
        existing_rows = _read_nonempty_row_count(out)
        diag["result"] = {
            "reason": "filters_removed_all_rows",
            "existing_rows": existing_rows,
            "action": "overwrite_empty" if args.allow_empty_overwrite else "keep_previous",
        }
        diag["source_after_fallback"] = source_label
        diag["source_files_after_fallback"] = [str(p) for p in used_paths]
        _write_diag(diag_path, diag)

        if not args.allow_empty_overwrite:
            print(
                json.dumps(
                    {
                        "written": str(out),
                        "rows": existing_rows,
                        "kept_previous": True,
                        "source": source_label,
                        "source_files": [str(p) for p in used_paths],
                        "message": "All candidates filtered out; keeping previous candidates_nextday.csv",
                    },
                    ensure_ascii=False,
                )
            )
            return

        _backup_existing(out)
        _atomic_write_csv(pd.DataFrame(), out)
        print(
            json.dumps(
                {
                    "written": str(out),
                    "rows": 0,
                    "kept_previous": False,
                    "source": source_label,
                    "source_files": [str(p) for p in used_paths],
                },
                ensure_ascii=False,
            )
        )
        return

    summary = build_summary(combined, sum(len(frame) for frame in frames))
    defaults = make_candidate_metadata_defaults(date_tag=date_tag, git_short_sha=_git_short_sha())
    combined = append_candidate_metadata(combined, defaults=defaults)
    _backup_existing(out)
    _atomic_write_csv(combined, out)
    if len(combined) >= fallback_min_rows:
        _write_last_good_snapshot(out)

    payload = {"written": str(out), "rows": int(len(combined))}
    payload.update(summary.to_json())
    payload["source"] = source_label
    payload["source_files"] = [str(p) for p in used_paths]
    payload["kept_previous"] = False

    diag["result"] = {
        "reason": "ok",
        "written_rows": int(len(combined)),
        "rows_in": summary.rows_in,
        "rows_filtered_out": summary.rows_filtered_out,
    }
    diag["source_after_fallback"] = source_label
    diag["source_files_after_fallback"] = [str(p) for p in used_paths]
    _write_diag(diag_path, diag)
    print(json.dumps(payload, ensure_ascii=False))


if __name__ == "__main__":
    main()
