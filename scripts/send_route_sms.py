"""
캐리 세탁물 수거 동선 문자 자동 발송
- 고정 스케줄 기반 (월요일/목요일)
- 월요일: 둘째·넷째주에 청량리 추가
- 목요일: 장한평 포함
- Solapi API로 기사님께 LMS 발송
"""

import os
import hashlib
import hmac
import time
import json
import uuid
import urllib.request
from datetime import datetime, date, timedelta, timezone
from typing import Optional
import pytz

KST = pytz.timezone("Asia/Seoul")

# ──────────────────────────────────────────────
# 숙소 정보 (단축 주소 기준)
# ──────────────────────────────────────────────
LOCATIONS: dict[str, dict] = {
    "장충단로 225": {
        "region": "장충동",
        "name": "메종드브릭",
        "access": "1층 키오스크 옆",
        "parking": None,
    },
    "고산자로 508-3": {
        "region": "제기동",
        "name": "스테이브리즈",
        "access": "0000* / 문열고 바로 왼쪽",
        "parking": None,
    },
    "능동로 165-1": {
        "region": "건대",
        "name": "화양프라하임",
        "access": "#1236 / 창고 2848* / 엘베 내려서 회색문",
        "parking": "주차가능",
    },
    "회기로 189": {
        "region": "회기",
        "name": "오를리",
        "access": "2층 바닥",
        "parking": None,
    },
    "장한로26나길 21": {
        "region": "장한평",
        "name": "프라하임장안2",
        "access": "#1236 / 210호 1482* / 2층 210호",
        "parking": None,
    },
    "왕산로 200, 1004호": {
        "region": "청량리",
        "name": "롯데캐슬 SKY-L65",
        "access": "1004호 문앞 / 주차 지하2층 하역장 진입",
        "parking": None,
    },
    "가락로28길 3-10": {
        "region": "송파",
        "name": "스테이브리즈",
        "access": "1234* / 건물 왼쪽 LOUNGE 세탁실",
        "parking": None,
    },
    "신림동1길 19-5": {
        "region": "신림",
        "name": "스테이모먼트",
        "access": "1210# / 공동주방 지나 STAFF ONLY 문 안",
        "parking": None,
    },
    "연희로4길 25-7": {
        "region": "연남",
        "name": "에코리빙",
        "access": "7777* / 반지하 라운지자동문 앞",
        "parking": None,
    },
}

# 기본 동선 (월·목 공통)
_BASE = [
    "장충단로 225",
    "고산자로 508-3",
    "능동로 165-1",
    "회기로 189",
    "가락로28길 3-10",
    "신림동1길 19-5",
    "연희로4길 25-7",
]


# ──────────────────────────────────────────────
# 동선 계산
# ──────────────────────────────────────────────
def get_route(today: date) -> list[str]:
    """오늘 요일·주차에 따라 방문 순서대로 location key 반환"""
    weekday = today.weekday()

    if weekday == 3:  # 목요일 — 장한평 포함 (회기 다음)
        return _BASE[:4] + ["장한로26나길 21"] + _BASE[4:]

    if weekday == 0:  # 월요일
        week_num = (today.day - 1) // 7 + 1  # 이번 달 몇 번째 월요일
        if week_num % 2 == 0:  # 둘째·넷째 → 청량리 포함
            return _BASE[:4] + ["왕산로 200, 1004호"] + _BASE[4:]
        return list(_BASE)

    return []  # 월·목 외 발송 안 함


def _next_thursday(today: date) -> date:
    days = (3 - today.weekday()) % 7
    return today + timedelta(days=days if days else 7)


def _next_conditional_monday(today: date) -> tuple[date, str]:
    """다음 2번째 또는 4번째 월요일과 '둘째주'/'넷째주' 레이블 반환"""
    days = (0 - today.weekday()) % 7
    candidate = today + timedelta(days=days if days else 7)
    while True:
        week_num = (candidate.day - 1) // 7 + 1
        if week_num % 2 == 0:
            label = "둘째주" if week_num == 2 else "넷째주"
            return candidate, label
        candidate += timedelta(days=7)


def get_next_notes(today: date, route: list[str]) -> list[str]:
    """오늘 동선에 없는 조건부 숙소 다음 일정 안내 문구 생성"""
    notes = []
    weekday = today.weekday()

    if "왕산로 200, 1004호" not in route:
        next_mon, label = _next_conditional_monday(today)
        notes.append(f"청량리는 {label} 월요일({next_mon.month}/{next_mon.day})")

    if "장한로26나길 21" not in route:
        next_thu = _next_thursday(today)
        notes.append(f"장한평은 목요일({next_thu.month}/{next_thu.day})")

    return notes


# ──────────────────────────────────────────────
# 메시지 생성
# ──────────────────────────────────────────────
CIRCLED_NUMS = "①②③④⑤⑥⑦⑧⑨⑩"
WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


def build_message(today: date, route: list[str]) -> tuple[str, str]:
    """LMS subject, text 반환"""
    weekday = WEEKDAY_KO[today.weekday()]
    subject = f"{today.month}/{today.day}({weekday}) 동선"

    lines = []
    for i, loc_key in enumerate(route):
        info = LOCATIONS[loc_key]
        num = CIRCLED_NUMS[i]
        lines.append(f"{num} {info['region']} | {info['name']}")
        lines.append("")
        lines.append(loc_key)
        lines.append(info["access"])
        if info.get("parking"):
            lines.append(info["parking"])
        if i < len(route) - 1:
            lines.append("↓")

    notes = get_next_notes(today, route)
    if notes:
        lines.append("")
        lines.extend(notes)

    lines.extend(["", "안전 운전하시고 감사합니다!"])
    return subject, "\n".join(lines)


# ──────────────────────────────────────────────
# Solapi SMS 발송
# ──────────────────────────────────────────────
def _solapi_signature(api_key: str, api_secret: str) -> tuple:
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    salt = str(uuid.uuid4()).replace("-", "")[:20]
    msg = date_str + salt
    signature = hmac.new(api_secret.encode(), msg.encode(), "sha256").hexdigest()
    return date_str, salt, signature


def _send_single(api_key: str, api_secret: str, sender: str, to: str, text: str, msg_type: str = "LMS", subject: str = "") -> None:
    date_str, salt, signature = _solapi_signature(api_key, api_secret)
    auth_header = f"HMAC-SHA256 apiKey={api_key}, date={date_str}, salt={salt}, signature={signature}"

    msg: dict = {"to": to, "from": sender, "text": text, "type": msg_type}
    if subject:
        msg["subject"] = subject

    payload = json.dumps({"message": msg}, ensure_ascii=False).encode("utf-8")

    req = urllib.request.Request(
        "https://api.solapi.com/messages/v4/send",
        data=payload,
        headers={
            "Authorization": auth_header,
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if result.get("errorCode"):
        raise RuntimeError(f"Solapi 발송 실패: {result}")


def send_sms(message: tuple[str, str], stop_count: int) -> None:
    api_key = os.environ["SOLAPI_API_KEY"]
    api_secret = os.environ["SOLAPI_API_SECRET"]
    sender = os.environ["SOLAPI_SENDER"]
    recipient = os.environ["RECIPIENT_PHONE"]
    owner_phone = os.environ.get("OWNER_PHONE", "")

    subject, body = message
    _send_single(api_key, api_secret, sender, recipient, body, "LMS", subject)
    print(f"[OK] 기사님 SMS 발송 완료 → {recipient}")

    if owner_phone:
        notify_text = f"[캐리] 동선 문자 발송 완료 ({stop_count}개 스톱)"
        _send_single(api_key, api_secret, sender, owner_phone, notify_text, "SMS")
        print(f"[OK] 오너 확인 알림 → {owner_phone}")


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main() -> None:
    test_date = os.environ.get("TEST_DATE")
    today = date.fromisoformat(test_date) if test_date else datetime.now(KST).date()

    print(f"[캐리 동선 발송] {today} ({WEEKDAY_KO[today.weekday()]})")

    route = get_route(today)
    if not route:
        print(f"[SKIP] {WEEKDAY_KO[today.weekday()]}요일은 발송 대상 아님")
        return

    subject, body = build_message(today, route)
    print("── 발송 메시지 ──")
    print(f"[제목] {subject}")
    print(body)
    print("─────────────────")

    dry_run = os.environ.get("DRY_RUN", "false").lower() == "true"
    if dry_run:
        print("[DRY_RUN] 실제 발송 안 함")
        return

    send_sms((subject, body), len(route))


if __name__ == "__main__":
    main()
