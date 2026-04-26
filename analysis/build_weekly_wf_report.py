import argparse
import datetime as dt
import json
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd


ANALYSIS_DIR = Path("analysis")
EXCEL_OUT_DIR = Path("output/excel")
SMTP_CFG_PATH = Path("state/smtp.json")
EMAIL_BLOCK_MARKER = Path("state/suspend_scheduled_jobs.txt")


def _to_date(tag: str) -> dt.date:
    return dt.datetime.strptime(tag, "%Y%m%d").date()


def _to_numeric(series: Iterable[float]) -> pd.Series:
    return pd.to_numeric(pd.Series(list(series)), errors="coerce")


def _list_daily_trade_dates() -> List[str]:
    dates: List[str] = []
    for path in ANALYSIS_DIR.glob("daily_trades_*.csv"):
        try:
            # daily_trades_YYYYMMDD.csv
            tag = path.stem.split("_")[2]
            _ = _to_date(tag)
            dates.append(tag)
        except Exception:
            continue
    return sorted(set(dates))


def _pick_week_dates(week_ending: dt.date, max_days: int = 5) -> List[str]:
    all_tags = _list_daily_trade_dates()
    filtered = [t for t in all_tags if _to_date(t) <= week_ending]
    return filtered[-max_days:]


def _load_candidates_for_date(date_tag: str) -> Optional[pd.DataFrame]:
    snap_path = EXCEL_OUT_DIR / f"candidates_for_{date_tag}.csv"
    if not snap_path.exists():
        return None
    try:
        df = pd.read_csv(snap_path)
    except Exception:
        return None
    if df.empty:
        return None
    # session を plan_tag から復元（例: AM0930_j-only → AM0930）
    if "plan_tag" in df.columns and "session" not in df.columns:
        df["session"] = df["plan_tag"].astype(str).str.split("_").str[0]
    df["date_tag"] = date_tag
    return df


def _load_trades_for_date(date_tag: str) -> Optional[pd.DataFrame]:
    path = ANALYSIS_DIR / f"daily_trades_{date_tag}.csv"
    if not path.exists():
        return None
    try:
        df = pd.read_csv(path)
    except Exception:
        return None
    if df.empty:
        return None
    return df


def _load_expected_pnl() -> Optional[pd.DataFrame]:
    path = ANALYSIS_DIR / "expected_pnl_daily.csv"
    if not path.exists():
        return None
    try:
        df = pd.read_csv(path)
    except Exception:
        return None
    if df.empty or "date" not in df.columns:
        return None
    return df


def summarize_day(
    date_tag: str,
    cand: Optional[pd.DataFrame],
    trades: Optional[pd.DataFrame],
    expected_daily: Optional[pd.DataFrame],
) -> Tuple[str, List[str]]:
    """Return (date_tag, lines[]) summary for a single day."""
    lines: List[str] = []
    lines.append(f"{date_tag}:")

    # Candidates summary
    if cand is None:
        lines.append("  candidates: (none)")
    else:
        n_rows = len(cand)
        n_tickers = cand["Ticker"].astype(str).nunique() if "Ticker" in cand.columns else n_rows
        pf = _to_numeric(cand.get("forward_pf_eff", []))
        win = _to_numeric(cand.get("forward_winrate", []))
        trades_c = _to_numeric(cand.get("forward_trades", []))
        lines.append(
            f"  candidates: {n_rows} rows, {n_tickers} tickers "
            f"(pf_mean={pf.mean():.2f} win_mean={win.mean():.3f} trades_mean={trades_c.mean():.1f})"
        )

        # strong combos
        if not pf.empty and not win.empty and not trades_c.empty:
            mask = (pf >= 1.3) & (win >= 0.6) & (trades_c >= 10)
            strong = cand.loc[mask].copy()
            n_strong = len(strong)
            if n_strong > 0:
                sample = ", ".join(
                    f"{row.get('Ticker','?')}@{row.get('plan_tag','?')}"
                    for _, row in strong.head(5).iterrows()
                )
                lines.append(f"    strong combos: {n_strong} (examples: {sample})")
            else:
                lines.append("    strong combos: 0")

    # Trades summary
    if trades is None:
        lines.append("  trades: (none)")
    else:
        pnl = _to_numeric(trades.get("pnl_bp", []))
        side = trades.get("side", pd.Series(["?"] * len(trades)))
        n_trades = len(trades)
        n_wins = int((pnl > 0).sum())
        winrate = n_wins / n_trades if n_trades else 0.0
        lines.append(
            f"  trades: {n_trades} fills (winrate={winrate:.3f}, pnl_bp_sum={pnl.sum():.1f}, pnl_bp_mean={pnl.mean():.1f})"
        )

    # Expected daily PnL
    if expected_daily is not None:
        row = expected_daily.loc[expected_daily["date"] == date_tag]
        if not row.empty:
            exp = float(row["expected_bp"].iloc[0]) if "expected_bp" in row.columns else None
            if exp is not None:
                lines.append(f"  expected_pnl: {exp:.1f} bp (from expected_pnl_daily.csv)")

    # Simple joined view: candidate vs trades per ticker/session
    if cand is not None and trades is not None and "Ticker" in cand.columns and "code" in trades.columns:
        cand_key = cand.copy()
        cand_key["Ticker"] = cand_key["Ticker"].astype(str)
        cand_key["session"] = cand_key.get("session", cand_key.get("Session", ""))
        trades_key = trades.copy()
        trades_key["code"] = trades_key["code"].astype(str)
        merged = trades_key.merge(
            cand_key,
            left_on=["code", "session"],
            right_on=["Ticker", "session"],
            how="left",
            suffixes=("_trade", "_wf"),
        )
        if not merged.empty:
            pnl = _to_numeric(merged.get("pnl_bp", []))
            lines.append(
                f"  joined (trade∩WF): {len(merged)} rows, pnl_bp_sum={pnl.sum():.1f}, pnl_bp_mean={pnl.mean():.1f}"
            )

    return date_tag, lines


def build_weekly_report(week_ending: str) -> str:
    """Build a human-readable weekly report string for the given week-ending date_tag."""
    week_date = _to_date(week_ending)
    dates = _pick_week_dates(week_date)
    if not dates:
        return f"ASAGAKE weekly WF report (week ending {week_ending})\nNo daily_trades_*.csv found.\n"

    expected_daily = _load_expected_pnl()

    lines: List[str] = []
    lines.append(f"ASAGAKE weekly WF report (week ending {week_ending})")
    lines.append("")
    lines.append("Summary by day:")

    for tag in dates:
        cand = _load_candidates_for_date(tag)
        trades = _load_trades_for_date(tag)
        _, day_lines = summarize_day(tag, cand, trades, expected_daily)
        lines.extend(day_lines)
        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def _send_email(subject: str, body: str, recipient: str) -> None:
    if EMAIL_BLOCK_MARKER.exists():
        print(f"[info] weekly WF report email disabled by marker: {EMAIL_BLOCK_MARKER}")
        return
    if not SMTP_CFG_PATH.exists():
        print(f"[warn] smtp config not found at {SMTP_CFG_PATH}; skip email")
        return
    try:
        cfg = json.loads(SMTP_CFG_PATH.read_text(encoding="utf-8"))
        host = cfg.get("host")
        port = int(cfg.get("port", 587))
        user = cfg.get("user")
        password = cfg.get("pass")
    except Exception as exc:
        print(f"[warn] failed to read smtp config: {exc}; skip email")
        return
    if not host or not user or not password:
        print("[warn] smtp config incomplete; skip email")
        return

    from email.message import EmailMessage
    import smtplib

    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    try:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)
        print(f"[info] weekly WF report email sent to {recipient}")
    except Exception as exc:
        print(f"[warn] failed to send weekly WF report email: {exc}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Build a weekly WF report by combining candidates_for_YYYYMMDD.csv, "
            "daily_trades_YYYYMMDD.csv, and expected_pnl_daily.csv."
        )
    )
    parser.add_argument(
        "--week-ending",
        type=str,
        help="Week-ending date_tag (YYYYMMDD). If omitted, use the latest daily_trades_*.csv date.",
    )
    parser.add_argument(
        "--email",
        action="store_true",
        help="Send the report via SMTP using state/smtp.json.",
    )
    parser.add_argument(
        "--recipient",
        type=str,
        default="shouichi.ikeda@gmail.com",
        help="Email recipient when --email is set.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.week_ending:
        week_ending = args.week_ending
    else:
        dates = _list_daily_trade_dates()
        if not dates:
            print("No daily_trades_*.csv found; nothing to report.")
            return
        week_ending = dates[-1]

    report = build_weekly_report(week_ending)
    print(report)

    if args.email:
        subject = f"ASAGAKE weekly WF report (week ending {week_ending})"
        _send_email(subject, report, args.recipient)


if __name__ == "__main__":
    main()

