#!/usr/bin/env python3
"""
Resend a history adequacy report email in Japanese, based on an existing JSON file.

Why:
- The scheduled history adequacy task historically sent only a raw JSON block, which is hard to read.
- This tool renders a human-friendly Japanese explanation and sends it via SMTP (state/smtp.json).

Usage:
  python tools/send_history_adequacy_mail.py --json analysis/history_adequacy_20251215.json --recipient you@example.com
"""

from __future__ import annotations

import argparse
import json
import smtplib
from email.message import EmailMessage
from pathlib import Path


def load_smtp_config(path: Path) -> dict:
    if not path.exists():
        raise SystemExit(f"smtp config not found: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise SystemExit(f"failed to read smtp config: {exc!r}") from exc


def render_body(payload: dict, label: str) -> str:
    lookback = payload.get("lookback_days")
    history_min = payload.get("history_min")
    history_median = payload.get("history_median")
    history_max = payload.get("history_max")

    pct_lt_lb = payload.get("pct_lt_lb", 0.0)
    pct_lt_half = payload.get("pct_lt_half_lb", 0.0)
    pct_ge_2x = payload.get("pct_ge_2x_lb", 0.0)

    train_med = payload.get("train_trades_median")
    forward_med = payload.get("forward_trades_median")

    messages = payload.get("messages") or []
    if not isinstance(messages, list):
        messages = [str(messages)]

    lines: list[str] = []
    lines.append(f"ASAGAKE 履歴の十分性チェック（{label}）")
    lines.append("")
    lines.append("【これは何？】")
    lines.append("候補銘柄ごとに、ローカル保存している Yahoo 1分足が「何営業日ぶんあるか」を数えたレポートです。")
    lines.append("（売買成績の良し悪しではなく、データの揃い具合のチェックです）")
    lines.append("")
    lines.append("【結論（短く）】")
    if messages:
        for m in messages:
            lines.append(f"- {m}")
    else:
        lines.append("- 特記事項なし（大きな不足は見えません）。")
    lines.append("")
    lines.append("【数字】")
    lines.append(f"- lookback（目安）: {lookback} 日")
    lines.append(f"- 保存済み1分足（営業日数）: 最小 {history_min} / 中央 {history_median} / 最大 {history_max} 日")
    lines.append(f"- lookback 未満: {pct_lt_lb*100:.1f}%")
    lines.append(f"- lookback の半分未満: {pct_lt_half*100:.1f}%")
    lines.append(f"- lookback の2倍以上: {pct_ge_2x*100:.1f}%")
    lines.append("")
    lines.append("【WF（テスト）量の目安】")
    lines.append(f"- train_trades（中央値）: {train_med}")
    lines.append(f"- forward_trades（中央値）: {forward_med}")
    lines.append("")
    lines.append("【次にやると良いこと（例）】")
    lines.append("- lookback を伸ばす前に、minute_cache（ローカル保存）を増やす頻度を上げる")
    lines.append("- Top200常連の銘柄を優先して1分足を保存し、週末バッチのネット取得を減らす")
    lines.append("")
    lines.append("※ 添付のJSONは生データです。必要なときだけ参照してください。")
    return "\n".join(lines) + "\n"


def send_mail(smtp_cfg: dict, recipient: str, subject: str, body: str, attachment: Path) -> None:
    user = smtp_cfg.get("user")
    password = smtp_cfg.get("pass")
    host = smtp_cfg.get("host")
    port = int(smtp_cfg.get("port", 587))
    if not (user and password and host and port):
        raise SystemExit("smtp.json must contain host/port/user/pass")

    msg = EmailMessage()
    msg["From"] = user
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    msg.add_attachment(
        attachment.read_bytes(),
        maintype="application",
        subtype="json",
        filename=attachment.name,
    )

    with smtplib.SMTP(host, port, timeout=30) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True, help="analysis/history_adequacy_YYYYMMDD.json")
    ap.add_argument("--recipient", required=True)
    ap.add_argument("--smtp", default="state/smtp.json")
    args = ap.parse_args()

    json_path = Path(args.json)
    if not json_path.exists():
        raise SystemExit(f"json not found: {json_path}")

    try:
        payload = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception as exc:
        raise SystemExit(f"failed to read json: {exc!r}") from exc

    # Label derived from filename when possible.
    label = json_path.stem.replace("history_adequacy_", "")
    body = render_body(payload, label=label)

    smtp_cfg = load_smtp_config(Path(args.smtp))
    subject = f"ASAGAKE 履歴の十分性チェック（{label}）"
    send_mail(smtp_cfg, args.recipient, subject, body, json_path)
    print(f"[info] resent to {args.recipient}: {json_path}")


if __name__ == "__main__":
    main()

