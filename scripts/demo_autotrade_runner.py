from __future__ import annotations

import argparse
import datetime as dt
import sys
import time
from pathlib import Path
from typing import Any

WB_PATH = Path("C:/AI/asagake/SHINSOKU.xlsm")
LOG_PATH = Path("C:/AI/asagake/logs/demo_autotrade.log")


def jst_now() -> dt.datetime:
    try:
        from zoneinfo import ZoneInfo  # type: ignore

        return dt.datetime.now(ZoneInfo("Asia/Tokyo"))
    except Exception:
        return dt.datetime.now()


def log_line(message: str) -> None:
    LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_PATH.open("a", encoding="utf-8") as fh:
        fh.write(f"{jst_now():%Y-%m-%d %H:%M:%S} {message}\n")


def sleep_until(target: dt.datetime) -> None:
    while True:
        now = jst_now()
        remaining = (target - now).total_seconds()
        if remaining <= 0:
            break
        time.sleep(min(remaining, 5))


def parse_excel_time(raw: Any, fallback: str) -> dt.time:
    if raw in (None, ""):
        raw = fallback

    if isinstance(raw, dt.datetime):
        return raw.time()
    if isinstance(raw, dt.time):
        return raw
    if isinstance(raw, (int, float)):
        frac = float(raw) % 1.0
        total_seconds = int(round(frac * 24 * 3600))
        hours = (total_seconds // 3600) % 24
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return dt.time(hours, minutes, seconds)

    text = str(raw).strip()
    if ":" in text:
        parts = text.split(":")
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2]) if len(parts) > 2 else 0
        return dt.time(hours, minutes, seconds)
    if len(text) == 4 and text.isdigit():
        return dt.time(int(text[:2]), int(text[2:]), 0)
    raise ValueError(f"Cannot parse session time from {text!r}")


def run_session(mode: str, start_when: str | None) -> int:
    try:
        import win32com.client  # type: ignore
    except Exception as e:  # pragma: no cover
        log_line(f"pywin32_error {e}")
        return 2

    if start_when == "09:00":
        now = jst_now()
        tgt = now.replace(hour=9, minute=0, second=0, microsecond=0)
        if tgt <= now:
            tgt = tgt + dt.timedelta(days=1)
        log_line(f"waiting_until {tgt:%Y-%m-%d %H:%M:%S}")
        sleep_until(tgt)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.AutomationSecurity = 1
    except Exception:
        pass

    wb = excel.Workbooks.Open(str(WB_PATH))
    try:
        ws = wb.Worksheets("NewDashboard")
        live_val = 0 if mode == "demo" else 1
        ws.Range("B9").Value = live_val
        ws.Range("B10").Value = "MS2Bridge.Place"

        log_line(f"session_start mode={mode} start_when={start_when}")
        macros = [
            "AutoTrader.ResetDashboardHeaders",
            "AutoTrader.ButtonLoadCandidates",
            "AutoTrader.ButtonPushCandidates",
            "AutoTrader.InstallRealtimeFormulas",
        ]
        for macro in macros:
            try:
                excel.Run(macro)
                log_line(f"macro_ok {macro}")
            except Exception as exc:
                log_line(f"macro_err {macro}: {exc}")
                raise

        start_t = parse_excel_time(ws.Range("B4").Value, "09:00")
        end_t = parse_excel_time(ws.Range("B5").Value, "09:15")
        now = jst_now()
        today_start = now.replace(hour=start_t.hour, minute=start_t.minute, second=start_t.second, microsecond=0)
        today_end = now.replace(hour=end_t.hour, minute=end_t.minute, second=end_t.second, microsecond=0)
        if now < today_start:
            log_line(f"waiting_for_session_start {today_start:%H:%M:%S}")
            sleep_until(today_start)

        log_line("auto_start")
        excel.Run("AutoTrader.ButtonStartAuto")
        while jst_now() < today_end:
            time.sleep(1.0)
        log_line("auto_stop")
        excel.Run("AutoTrader.ButtonStopAuto")

        wb.Save()
        log_line("session_done")
        return 0
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--mode", choices=["demo", "live"], default="demo")
    ap.add_argument("--start-when", choices=["now", "09:00"], default="now")
    args = ap.parse_args()
    return run_session(args.mode, None if args.start_when == "now" else "09:00")


if __name__ == "__main__":
    sys.exit(main())
