# CARRY ROUTE 왕산로 200 추가

## Goal
- 2026-04-09 동선 문자에 `왕산로 200, 1004호`를 즉시 포함한다.
- 2026-04-13부터는 2주 간격 월요일마다 같은 주소가 자동 포함되게 맞춘다.

## Working Contract
- In scope: 오늘 목요일 동선에 `왕산로 200, 1004호` 1회 추가.
- In scope: 2026-04-13 시작 2주 주기 월요일 반복 로직 반영.
- In scope: 기사님 문자 하단 다음 일정 안내 문구를 새 규칙 기준으로 정리.
- In scope: 관련 운영 문서 한 곳 업데이트.
- Out of scope: 숙소 정보 변경, 사업자 정보 변경, 다른 주소 동선 재배치.
- Done means: 2026-04-09, 2026-04-13, 2026-04-27 테스트에서 `왕산로 200, 1004호`가 포함되고, 2026-04-20 테스트에서는 빠진다.
- How to verify: `TEST_DATE`로 `scripts/send_route_sms.py`를 실행해 발송 메시지 비교 확인.
- Main risks: 월요일 주기를 월차 기준이 아니라 14일 기준으로 바꾸기 때문에 다음 일정 안내 문구도 함께 바뀌어야 한다.

## Recent Changes
- `scripts/send_route_sms.py`에서 `왕산로 200, 1004호`를 2026-04-09 1회 + 2026-04-13 시작 14일 간격 월요일 규칙으로 분리했다.
- 목요일 기본 동선은 유지하되, 2026-04-09에는 청량리와 장한평이 함께 들어가도록 합쳤다.
- 기사님 안내 문구에서 기존 `둘째주/넷째주` 표현을 제거하고 실제 다음 일정 날짜로 바꿨다.
- `CLAUDE.md`에 운영 기준을 추가했다.
- 커밋 `8b80378`로 `main` 브랜치 푸시를 완료했다.
- 사용자 정정 반영: 반복 시작점은 `2026-04-13`이 아니라 `2026-04-06(금주 월요일)` 기준으로 다시 맞춘다.
- 커밋 `046e82f`로 정정분을 `main`에 추가 푸시했다.

## Verification
- `env DRY_RUN=true TEST_DATE=2026-04-09 python3 scripts/send_route_sms.py` → 오늘 목요일 메시지에 `왕산로 200, 1004호`와 `장한로26나길 21`이 모두 포함됨 확인.
- `env DRY_RUN=true TEST_DATE=2026-04-13 python3 scripts/send_route_sms.py` → 다음 주 월요일 메시지에서 `왕산로 200, 1004호` 제외, 다음 일정 `4/20(월)` 안내 확인.
- `env DRY_RUN=true TEST_DATE=2026-04-20 python3 scripts/send_route_sms.py` → 다다음주 월요일 메시지에 `왕산로 200, 1004호` 포함 확인.
- `env DRY_RUN=true TEST_DATE=2026-04-27 python3 scripts/send_route_sms.py` → 그다음 월요일 메시지에서 `왕산로 200, 1004호` 제외, 다음 일정 `5/4(월)` 안내 확인.
- `env DRY_RUN=true TEST_DATE=2026-05-04 python3 scripts/send_route_sms.py` → 다음 회차 월요일 메시지에 `왕산로 200, 1004호` 포함 확인.
- `python3 -m py_compile scripts/send_route_sms.py` → 문법 통과.

## Independent Review
- Contract met: 예. 오늘 1회 추가, 다음 주 월요일 시작 2주 주기, 안내 문구 정리, 운영 문서 반영까지 범위 내에서 끝냈다.
- Contract met (정정 반영): 예. 반복 시작점을 `4/13`에서 `4/6` 기준으로 바로잡아 실제 다음 방문일이 `4/20`이 되게 수정했다.
- Out-of-scope preserved: 예. 주소 정보, 다른 숙소 순서, 사업자 정보는 건드리지 않았다.
- Verification evidence present: 예. 오늘/다음 주기/중간 주차/다음 달 회차까지 날짜별 메시지를 직접 찍어 확인했다.
- Remaining risk: 별도 수동 발송은 하지 않았으므로 2026-04-09 10:00 KST GitHub Actions 스케줄이 정상 실행돼야 실제 문자가 발송된다.

## Done Gate
- Success criteria: 충족. 오늘 목요일 1회 추가와 `4/6` 기준 격주 월요일 규칙이 모두 테스트 메시지에서 확인됐다.
- Verification evidence: 날짜별 DRY_RUN 5건 + 문법 검사 1건 기록 완료.
- Remaining risks: 스케줄러 자체 장애가 있으면 문자 발송은 별도 확인이 필요하다. 코드와 원격 반영 상태는 완료됐다.
- Next action: 종료. 2026-04-09 10:00 KST 배치가 `4/6` 기준 격주 규칙으로 실행된다.
