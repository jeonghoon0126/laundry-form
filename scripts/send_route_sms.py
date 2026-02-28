"""
캐리 세탁물 수거 동선 문자 자동 발송
- Supabase laundry_records에서 오늘 수거 위치 조회
- 미리 정의된 순서로 동선 구성
- Solapi API로 기사님께 문자 발송
"""

import os
import sys
import hashlib
import hmac
import time
import json
import uuid
import urllib.request
import urllib.parse
from datetime import datetime, date, timedelta, timezone
from typing import Optional
import pytz

KST = pytz.timezone("Asia/Seoul")

# ──────────────────────────────────────────────
# 숙소 정보 (주소 → 표시 정보 매핑)
# ──────────────────────────────────────────────
LOCATIONS = {
    "동대문구 고산자로 508-3": {
        "order": 1,
        "region": "제기동",
        "name": "스테이브리즈",
        "access": "0000* / 문 열고 바로 왼쪽",
        "parking": None,
    },
    "광진구 능동로 165-1": {
        "order": 2,
        "region": "건대",
        "name": "화양프라하임",
        "access": "#1236 / 창고 2848* / 엘베 내려서 회색문",
        "parking": "주차 가능",
    },
    "동대문구 회기로 189": {
        "order": 3,
        "region": "회기",
        "name": "오를리",
        "access": "2층 바닥",
        "parking": None,
    },
    "동대문구 장한로26나길 21": {
        "order": 4,
        "region": "장한평",
        "name": "프라하임장안2",
        "access": "#1236 / 210호 1482* / 2층 210호",
        "parking": None,
    },
    "동대문구 왕산로 200": {
        "order": 4,  # 장한평과 동일 슬롯 (동시에 안 나옴)
        "region": "청량리",
        "name": "롯데캐슬 SKY-L65",
        "access": "1004호 문앞 / 주차 지하2층 하역장 진입",
        "parking": None,
    },
    "송파구 가락로28길 3-10": {
        "order": 5,
        "region": "송파",
        "name": "스테이브리즈",
        "access": "1234* / 건물 왼쪽 LOUNGE 세탁실",
        "parking": None,
    },
    "관악구 신림동1길 19-5": {
        "order": 6,
        "region": "신림",
        "name": "스테이모먼트",
        "access": "1210# / 공동주방 지나 STAFF ONLY 문 안",
        "parking": None,
    },
    "서대문구 연희로4길 25-7": {
        "order": 7,
        "region": "연남",
        "name": "에코리빙",
        "access": "7777* / 반지하 라운지 자동문 앞",
        "parking": None,
    },
}

# laundry_records.location 값 → LOCATIONS 키 정규화 매핑
# (시트 입력값이 다를 수 있어서 보정)
ADDRESS_NORMALIZE = {
    "고산자로 508-3": "동대문구 고산자로 508-3",
    "동대문구 고산자로 508-3": "동대문구 고산자로 508-3",
    "능동로 165-1": "광진구 능동로 165-1",
    "광진구 능동로 165-1": "광진구 능동로 165-1",
    "회기로 189": "동대문구 회기로 189",
    "동대문구 회기로 189": "동대문구 회기로 189",
    "장한로26나길 21": "동대문구 장한로26나길 21",
    "동대문구 장한로26나길 21": "동대문구 장한로26나길 21",
    "왕산로 200": "동대문구 왕산로 200",
    "동대문구 왕산로 200": "동대문구 왕산로 200",
    "가락로28길 3-10": "송파구 가락로28길 3-10",
    "송파구 가락로28길 3-10": "송파구 가락로28길 3-10",
    "신림동1길 19-5": "관악구 신림동1길 19-5",
    "관악구 신림동1길 19-5": "관악구 신림동1길 19-5",
    "연희로4길 25-7": "서대문구 연희로4길 25-7",
    "서대문구 연희로4길 25-7": "서대문구 연희로4길 25-7",
}


def normalize_address(raw: str) -> Optional[str]:
    """laundry_records.location → LOCATIONS 키로 정규화"""
    raw = raw.strip()
    if raw in ADDRESS_NORMALIZE:
        return ADDRESS_NORMALIZE[raw]
    # 부분 매칭 시도
    for key, normalized in ADDRESS_NORMALIZE.items():
        if key in raw or raw in key:
            return normalized
    return None


# ──────────────────────────────────────────────
# Supabase: 오늘 수거 위치 조회
# ──────────────────────────────────────────────
def fetch_today_locations(today: date) -> list[str]:
    """오늘 날짜의 laundry_records에서 수거 위치 목록 반환"""
    url = os.environ["SUPABASE_URL"].rstrip("/")
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]
    date_str = today.isoformat()

    req_url = (
        f"{url}/rest/v1/laundry_records"
        f"?select=location"
        f"&record_date=eq.{date_str}"
    )
    req = urllib.request.Request(
        req_url,
        headers={
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req) as resp:
        records = json.loads(resp.read())

    locations = []
    seen = set()
    for r in records:
        loc = normalize_address(r["location"])
        if loc and loc not in seen:
            seen.add(loc)
            locations.append(loc)
    return locations


def fetch_next_schedule(location_key: str, after: date) -> Optional[date]:
    """특정 위치의 다음 수거 예정일 조회 (미래 레코드)"""
    url = os.environ["SUPABASE_URL"].rstrip("/")
    key = os.environ["SUPABASE_SERVICE_ROLE_KEY"]

    # LOCATIONS 역매핑으로 원본 주소 패턴 찾기
    info = LOCATIONS.get(location_key, {})
    region = info.get("region", "")

    req_url = (
        f"{url}/rest/v1/laundry_records"
        f"?select=record_date,location"
        f"&record_date=gt.{after.isoformat()}"
        f"&order=record_date.asc"
        f"&limit=30"
    )
    req = urllib.request.Request(
        req_url,
        headers={
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req) as resp:
        records = json.loads(resp.read())

    for r in records:
        normalized = normalize_address(r["location"])
        if normalized == location_key:
            return date.fromisoformat(r["record_date"])
    return None


# ──────────────────────────────────────────────
# 동선 메시지 생성
# ──────────────────────────────────────────────
CIRCLED_NUMS = "①②③④⑤⑥⑦⑧⑨⑩"
WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]


def build_message(today: date, location_keys: list[str]) -> str:
    """기사님용 동선 문자 메시지 생성"""
    weekday = WEEKDAY_KO[today.weekday()]
    date_label = f"{today.month}/{today.day}({weekday})"

    lines = [f"[Web발신]", f"{date_label} 동선", ""]

    # 순서 정렬
    sorted_locs = sorted(location_keys, key=lambda k: LOCATIONS[k]["order"])

    for i, loc_key in enumerate(sorted_locs):
        info = LOCATIONS[loc_key]
        num = CIRCLED_NUMS[i]
        lines.append(f"{num} {info['region']} | {info['name']}")
        lines.append(loc_key)  # 전체 주소
        lines.append(info["access"])
        if info.get("parking"):
            lines.append(info["parking"])
        if i < len(sorted_locs) - 1:
            lines.append("↓")

    # 오늘 빠진 위치의 다음 일정 안내
    all_keys = set(LOCATIONS.keys())
    missing_keys = all_keys - set(location_keys)
    next_notes = []
    for key in missing_keys:
        next_date = fetch_next_schedule(key, today)
        if next_date:
            info = LOCATIONS[key]
            next_notes.append(
                f"{info['region']}({info['name']})은 "
                f"{next_date.month}/{next_date.day}"
            )

    if next_notes:
        lines.append("")
        lines.extend(next_notes)

    lines.extend(["", "안전 운전하시고 감사합니다!"])
    return "\n".join(lines)


# ──────────────────────────────────────────────
# Solapi SMS 발송
# ──────────────────────────────────────────────
def _solapi_signature(api_key: str, api_secret: str) -> tuple:
    """Solapi HMAC-SHA256 서명 생성 (ISO 8601 date + random salt)"""
    date_str = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    salt = str(uuid.uuid4()).replace("-", "")[:20]
    msg = date_str + salt
    signature = hmac.new(api_secret.encode(), msg.encode(), "sha256").hexdigest()
    return date_str, salt, signature


def _send_single(api_key: str, api_secret: str, sender: str, to: str, text: str, msg_type: str = "LMS") -> None:
    """Solapi API로 단건 발송"""
    now_ms, salt, signature = _solapi_signature(api_key, api_secret)
    auth_header = f"HMAC-SHA256 apiKey={api_key}, date={now_ms}, salt={salt}, signature={signature}"

    payload = json.dumps({
        "message": {
            "to": to,
            "from": sender,
            "text": text,
            "type": msg_type,
        }
    }).encode()

    req = urllib.request.Request(
        "https://api.solapi.com/messages/v4/send",
        data=payload,
        headers={
            "Authorization": auth_header,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req) as resp:
        result = json.loads(resp.read())

    if result.get("errorCode"):
        raise RuntimeError(f"Solapi 발송 실패: {result}")


def send_sms(message: str, stop_count: int) -> None:
    """기사님께 동선 LMS 발송 + 오너에게 확인 SMS 발송"""
    api_key = os.environ["SOLAPI_API_KEY"]
    api_secret = os.environ["SOLAPI_API_SECRET"]
    sender = os.environ["SOLAPI_SENDER"]
    recipient = os.environ["RECIPIENT_PHONE"]
    owner_phone = os.environ.get("OWNER_PHONE", "")

    # 기사님 동선 발송
    _send_single(api_key, api_secret, sender, recipient, message, "LMS")
    print(f"[OK] 기사님 SMS 발송 완료 → {recipient}")

    # 오너 확인 알림 (OWNER_PHONE 설정 시)
    if owner_phone:
        notify_text = f"[캐리] 동선 문자 발송 완료 ({stop_count}개 스톱)"
        _send_single(api_key, api_secret, sender, owner_phone, notify_text, "SMS")
        print(f"[OK] 오너 확인 알림 → {owner_phone}")


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main() -> None:
    test_date = os.environ.get("TEST_DATE")
    if test_date:
        today = date.fromisoformat(test_date)
    else:
        today = datetime.now(KST).date()

    print(f"[캐리 동선 발송] {today} ({WEEKDAY_KO[today.weekday()]})")

    location_keys = fetch_today_locations(today)
    if not location_keys:
        print(f"[SKIP] {today} 수거 레코드 없음 — 발송 안 함")
        sys.exit(0)

    known = [k for k in location_keys if k in LOCATIONS]
    unknown = [k for k in location_keys if k not in LOCATIONS]

    if unknown:
        print(f"[WARN] 매핑 없는 주소: {unknown}")

    if not known:
        print("[SKIP] 인식된 주소 없음")
        sys.exit(0)

    message = build_message(today, known)
    print("── 발송 메시지 ──")
    print(message)
    print("─────────────────")

    dry_run = os.environ.get("DRY_RUN", "false").lower() == "true"
    if dry_run:
        print("[DRY_RUN] 실제 발송 안 함")
        return

    send_sms(message, len(known))


if __name__ == "__main__":
    main()
