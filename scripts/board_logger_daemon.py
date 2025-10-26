import argparse
import datetime as dt
import time
from pathlib import Path

from board_snapshot import snapshot_board


def parse_time(value: str) -> dt.time:
    return dt.datetime.strptime(value, "%H:%M").time()


def wait_until(target: dt.datetime) -> None:
    while True:
        now = dt.datetime.now()
        delta = (target - now).total_seconds()
        if delta <= 0:
            break
        time.sleep(min(delta, 30))


def run(board_path: Path, out_root: Path, start_t: dt.time, end_t: dt.time, interval_sec: int) -> None:
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
                fh.write(f"{dt.datetime.now().isoformat()} {exc}\n")

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
    ap.add_argument("--start", default="09:00")
    ap.add_argument("--end", default="15:30")
    ap.add_argument("--interval", type=int, default=60)
    args = ap.parse_args()

    board_path = Path(args.board)
    out_root = Path(args.outdir)
    start_t = parse_time(args.start)
    end_t = parse_time(args.end)

    run(board_path, out_root, start_t, end_t, max(10, args.interval))


if __name__ == "__main__":
    main()

