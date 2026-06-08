#!/usr/bin/env python3
"""
캐리 주별 매출 보고 이메일 발송
- 지난 월요일~일요일 매출을 숙소별, 일자별로 요약
- 기존 정산 금액 계산 기준을 재사용
- Gmail SMTP로 지정 수신자에게 발송
"""

from __future__ import annotations

import os
import smtplib
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from html import escape
from pathlib import Path

import psycopg2

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import generate_invoices as gi  # noqa: E402

KST = timezone(timedelta(hours=9))
DEFAULT_RECIPIENT = "kham0126@gmail.com"
WEEKLY_REPORT_TO = os.environ.get("WEEKLY_REPORT_TO", DEFAULT_RECIPIENT).strip() or DEFAULT_RECIPIENT


def previous_week_window(today: date) -> tuple[date, date]:
    """오늘 기준 마지막으로 완료된 월~일 기간을 반환한다."""
    this_monday = today - timedelta(days=today.weekday())
    start = this_monday - timedelta(days=7)
    return start, this_monday


def format_period(start: date, end_exclusive: date) -> str:
    end = end_exclusive - timedelta(days=1)
    return f"{start.month:02d}/{start.day:02d}~{end.month:02d}/{end.day:02d}"


def get_weekly_rows(start: date, end_exclusive: date) -> list[tuple]:
    """주간 세탁 기록을 조회한다."""
    conn = psycopg2.connect(gi.SUPABASE_URI)
    cur = conn.cursor()
    cur.execute(
        """
        SELECT record_date, location, blanket, mat, pillow_cover, towel, body_towel, pillow_fill, cotton_blanket
        FROM laundry_records
        WHERE record_date >= %s
          AND record_date < %s
        ORDER BY record_date, location
        """,
        (start, end_exclusive),
    )
    rows = cur.fetchall()
    cur.close()
    conn.close()
    return rows


def build_weekly_summary(rows: list[tuple], start: date, end_exclusive: date) -> dict:
    """주간 매출을 숙소별, 일자별로 집계한다."""
    locations: dict[str, int] = defaultdict(int)
    daily: dict[date, int] = defaultdict(int)
    total = 0
    record_count = 0

    for row in rows:
        record_date, location, *_ = row
        amount = gi.calculate_record_amount(row)
        if amount <= 0:
            continue
        locations[location] += amount
        daily[record_date] += amount
        total += amount
        record_count += 1

    return {
        "start": start,
        "end_exclusive": end_exclusive,
        "total": total,
        "record_count": record_count,
        "locations": dict(sorted(locations.items())),
        "daily": dict(sorted(daily.items())),
    }


def _bar_cell(amount: int, max_amount: int, color: str) -> str:
    width = int(amount / max_amount * 100) if max_amount else 0
    return (
        f'<div style="background:#e5e7eb;border-radius:4px;height:16px;width:100%;">'
        f'<div style="background:{color};border-radius:4px;height:16px;width:{max(width, 2) if amount else 0}%;"></div>'
        f"</div>"
    )


def build_weekly_email_html(summary: dict) -> str:
    """Gmail에서 바로 읽히는 주간 매출 보고 HTML을 만든다."""
    period = format_period(summary["start"], summary["end_exclusive"])
    location_items = list(summary["locations"].items())
    daily_items = list(summary["daily"].items())
    max_location = max((amount for _, amount in location_items), default=0)
    max_daily = max((amount for _, amount in daily_items), default=0)

    location_rows = "".join(
        "<tr>"
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;text-align:left;">{escape(location)}</td>'
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;width:45%;">{_bar_cell(amount, max_location, "#f97316")}</td>'
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;text-align:right;font-weight:bold;white-space:nowrap;">{gi.format_won(amount)}</td>'
        "</tr>"
        for location, amount in location_items
    )
    daily_rows = "".join(
        "<tr>"
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;text-align:left;white-space:nowrap;">{day.month:02d}/{day.day:02d}</td>'
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;width:55%;">{_bar_cell(amount, max_daily, "#2563eb")}</td>'
        f'<td style="padding:7px 10px;border-bottom:1px solid #e5e7eb;text-align:right;font-weight:bold;white-space:nowrap;">{gi.format_won(amount)}</td>'
        "</tr>"
        for day, amount in daily_items
    )

    if not location_rows:
        location_rows = '<tr><td colspan="3" style="padding:12px 10px;color:#64748b;">해당 기간 매출 기록이 없습니다.</td></tr>'
    if not daily_rows:
        daily_rows = '<tr><td colspan="3" style="padding:12px 10px;color:#64748b;">해당 기간 매출 기록이 없습니다.</td></tr>'

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:'Apple SD Gothic Neo',Arial,sans-serif;max-width:720px;margin:0 auto;padding:20px;color:#111827;">
  <h2 style="margin:0 0 6px;font-size:21px;color:#111827;">[캐리] 주별 매출 보고</h2>
  <p style="margin:0 0 18px;font-size:13px;color:#64748b;">보고 기간 {period} · 매출 기록 {summary['record_count']}건</p>
  <div style="background:#fff7ed;border:1px solid #fed7aa;border-radius:8px;padding:16px;margin-bottom:18px;">
    <p style="margin:0 0 4px;color:#9a3412;font-size:13px;font-weight:bold;">주간 총 매출</p>
    <p style="margin:0;font-size:30px;line-height:1.2;font-weight:bold;color:#111827;">{gi.format_won(summary['total'])}</p>
  </div>
  <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:8px;padding:16px;margin-bottom:18px;">
    <h3 style="margin:0 0 12px;font-size:15px;color:#c2410c;">숙소별 매출</h3>
    <table style="border-collapse:collapse;width:100%;font-size:13px;">{location_rows}</table>
  </div>
  <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:8px;padding:16px;">
    <h3 style="margin:0 0 12px;font-size:15px;color:#1d4ed8;">일자별 매출</h3>
    <table style="border-collapse:collapse;width:100%;font-size:13px;">{daily_rows}</table>
  </div>
</body></html>"""


def send_weekly_report_email(subject: str, html: str, recipient: str = WEEKLY_REPORT_TO) -> bool:
    """주간 매출 보고 이메일을 발송한다."""
    if not gi.EMAIL_PASSWORD:
        print("EMAIL_PASSWORD 미설정. 주별 매출 보고 이메일 건너뜀.")
        return False

    msg = MIMEMultipart("alternative")
    msg["From"] = gi.EMAIL_FROM
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.attach(MIMEText(html, "html", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(gi.EMAIL_FROM, gi.EMAIL_PASSWORD)
            refused = server.sendmail(gi.EMAIL_FROM, [recipient], msg.as_string())
        if refused:
            print(f"주별 매출 보고 이메일 일부 수신 거부: {refused}")
        else:
            print(f"주별 매출 보고 이메일 발송 완료 → {recipient}")
        return True
    except Exception as e:
        print(f"주별 매출 보고 이메일 발송 실패: {e}")
        return False


def parse_window(argv: list[str]) -> tuple[date, date]:
    """명령행 또는 환경변수로 보고 기간을 정한다."""
    env_start = os.environ.get("REPORT_START_DATE", "").strip()
    env_end = os.environ.get("REPORT_END_DATE", "").strip()
    if len(argv) >= 3:
        return date.fromisoformat(argv[1]), date.fromisoformat(argv[2])
    if env_start and env_end:
        return date.fromisoformat(env_start), date.fromisoformat(env_end)
    return previous_week_window(datetime.now(KST).date())


def main() -> None:
    start, end_exclusive = parse_window(sys.argv)
    print(f"[캐리 주별 매출 보고] {format_period(start, end_exclusive)}")

    rows = get_weekly_rows(start, end_exclusive)
    summary = build_weekly_summary(rows, start, end_exclusive)
    html = build_weekly_email_html(summary)
    subject = f"[캐리] {format_period(start, end_exclusive)} 주별 매출 보고"

    print(f"총 매출: {gi.format_won(summary['total'])} / 기록 {summary['record_count']}건 / 수신자 {WEEKLY_REPORT_TO}")

    if os.environ.get("DRY_RUN", "false").lower() == "true":
        print("[DRY_RUN] 실제 이메일 발송 안 함")
        return

    send_weekly_report_email(subject, html, WEEKLY_REPORT_TO)


if __name__ == "__main__":
    main()
