import argparse
import datetime as dt
import shutil
import time
from pathlib import Path
from typing import Optional

from board_snapshot import snapshot_board, snapshot_dashboard_j


def parse_time(value: str) -> dt.time:
    return dt.datetime.strptime(value, "%H:%M").time()


def wait_until(target: dt.datetime) -> None:
    while True:
        now = dt.datetime.now()
        delta = (target - now).total_seconds()
        if delta <= 0:
            break
        time.sleep(min(delta, 30))


def run(
    board_path: Path,
    out_root: Path,
    start_t: dt.time,
    end_t: dt.time,
    interval_sec: int,
    dashboard_path: Optional[Path],
    dash_out_root: Optional[Path],
    retain_days: int,
) -> None:
    today = dt.date.today()
    start_dt = dt.datetime.combine(today, start_t)
    end_dt = dt.datetime.combine(today, end_t)

    now = dt.datetime.now()
    if now < start_dt:
        wait_until(start_dt)
        now = dt.datetime.now()

    while now <= end_dt:
        day_dir = out_root / today.strftime("%Y%m%d")
        day_dir.mkdir(parents=True, exist_ok=True)
        try:
            snapshot_board(board_path, day_dir)
        except Exception as exc:  # pragma: no cover
            err_path = day_dir / "errors.log"
            with err_path.open("a", encoding="utf-8") as fh:
                fh.write(f"{dt.datetime.now().isoformat()} board:{exc}\n")
        if dashboard_path and dash_out_root:
            try:
                dash_dir = dash_out_root / today.strftime("%Y%m%d")
                snapshot_dashboard_j(dashboard_path, dash_dir)
            except Exception as exc:  # pragma: no cover
                err_path = day_dir / "errors.log"
                with err_path.open("a", encoding="utf-8") as fh:
                    fh.write(f"{dt.datetime.now().isoformat()} dashboard:{exc}\n")

        if retain_days > 0:
            purge_old_dirs(out_root, retain_days)
            if dash_out_root:
                purge_old_dirs(dash_out_root, retain_days)

        now = dt.datetime.now()
        next_tick = (now + dt.timedelta(seconds=interval_sec))
        # Align to next whole interval boundary
        next_tick = next_tick.replace(second=0, microsecond=0)
        if next_tick <= now:
            next_tick = now + dt.timedelta(seconds=interval_sec)
        if next_tick > end_dt:
            break
        wait_until(next_tick)
        now = dt.datetime.now()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--board", default=r"excel/BoardLogger.xlsx")
    ap.add_argument("--outdir", default=r"output/board_logs")
    ap.add_argument("--dashboard", default=r"C:/AI/asagake/ASAGAKE.xlsm")
    ap.add_argument("--dash-outdir", default=r"output/j_logs")
    ap.add_argument("--start", default="09:00")
    ap.add_argument("--end", default="15:30")
    ap.add_argument("--interval", type=int, default=60)
    ap.add_argument("--retain-days", type=int, default=30, help="Number of days to retain logs")
    args = ap.parse_args()

    board_path = Path(args.board)
    out_root = Path(args.outdir)
    start_t = parse_time(args.start)
    end_t = parse_time(args.end)
    dash_path = Path(args.dashboard) if args.dashboard else None
    dash_out = Path(args.dash_outdir) if args.dash_outdir else None

    run(
        board_path,
        out_root,
        start_t,
        end_t,
        max(5, args.interval),
        dash_path,
        dash_out,
        max(0, args.retain_days),
    )


def purge_old_dirs(root: Path, retain_days: int) -> None:
    if retain_days <= 0:
        return
    cutoff = dt.date.today() - dt.timedelta(days=retain_days)
    if not root.exists():
        return
    for child in root.iterdir():
        if not child.is_dir():
            continue
        try:
            day = dt.datetime.strptime(child.name, "%Y%m%d").date()
        except ValueError:
            continue
        if day < cutoff:
            shutil.rmtree(child, ignore_errors=True)


if __name__ == "__main__":
    main()
