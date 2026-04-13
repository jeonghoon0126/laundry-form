# CARRY ROUTE scheduler fix

## Goal
- 캐리 동선 문자가 실제 발송 시각 기준으로 밀리지 않게 복구한다.
- 오너 확인 문자가 다시 본인 휴대폰으로 오게 맞춘다.
- 오늘 발송도 즉시 처리 가능한 상태로 만든다.

## Working Contract
- In scope: GitHub 스케줄 지연 원인을 확인하고, 지연을 우회하는 정확 시각 발송 경로를 추가한다.
- In scope: 같은 날 중복 발송이 나가지 않도록 가드를 넣는다.
- In scope: 오너 확인 번호를 본인 번호로 복구하고 오늘 발송을 다시 보낸다.
- In scope: 운영 문서를 현재 구조 기준으로 갱신한다.
- Out of scope: 월·목 발송 규칙 자체 변경, 숙소 순서 변경, 정산 로직 변경.
- Done means: 오늘 문자가 발송되고, 이후에는 로컬 10:00 디스패처가 GitHub 워크플로우를 바로 깨우며, 늦게 도는 GitHub schedule은 같은 날 중복 발송을 건너뛴다.
- How to verify: GitHub workflow run 기록, launchd 등록 상태, 수동 dispatch 결과, 오늘 날짜 기준 성공 run 존재 여부를 확인한다.
- Main risks: 로컬 Mac이 꺼져 있거나 로그인 세션이 없으면 launchd dispatch가 실행되지 않는다. 이 경우 GitHub schedule이 백업으로 남아 있지만 시각은 늦을 수 있다.

## Status
- 완료

## Root Cause
- GitHub Actions `schedule`은 `10:00 KST`로 설정되어 있었지만 실제 run 기록은 `12:43~13:08 KST`로 2~3시간씩 밀리고 있었다.
- 그래서 오늘 `04/13`에도 `10시 기준 발송 미실행` 상태였고, 본인 휴대폰 미수신은 오늘 run 자체가 없었던 영향으로 보는 것이 맞다.
- `OWNER_PHONE`은 어젯밤 수정 흔적이 있었지만, 오늘 문제의 직접 원인은 번호보다 `스케줄 미실행`이었다.

## Recent Changes
- `scripts/dispatch_route_sms.py` 추가
  - 같은 날 성공 run이 이미 있으면 중복 발송을 막는 가드 포함
- `.github/workflows/send-route-sms.yml` 수정
  - `actions: read` 권한 추가
  - 발송 전 `scripts/dispatch_route_sms.py --check-only` 실행
  - 같은 날 성공 run이 있으면 늦게 도는 schedule run 자동 건너뜀
- `CLAUDE.md`에 exact 10:00 디스패처 + backup schedule 구조 반영
- 오늘분 `workflow_dispatch` 수동 발송 실행
  - run id `24322028775`
  - 기사님 문자 + 오너 동선 문자 발송 성공
- 사용자 기준 확정 반영
  - `정시 10:00`보다 `점심 전`이 더 중요하므로 로컬 launchd exact-time 보조장치는 제거
  - 운영 구조는 다시 `GitHub Actions only`로 정리

## Verification
- GitHub run 이력 확인
  - 기존 `schedule` 성공 run 시각: `04/06 12:57 KST`, `04/09 12:47 KST`, `04/02 12:43 KST`, `03/30 13:08 KST`
  - 결론: 설정 시각은 10시지만 GitHub schedule은 실제로 2~3시간 지연되고 있었다
- 코드/설정 검증
  - `python3 -m py_compile scripts/send_route_sms.py scripts/dispatch_route_sms.py` 통과
  - `ruby -e 'require "yaml"; YAML.load_file(...)'` → workflow YAML 파싱 통과
  - `plutil -lint ~/Library/LaunchAgents/com.wjh.carry-route-sms-dispatch.plist` → OK
- launchd 등록 검증
  - 로컬 launchd 보조장치는 사용자 기준상 불필요 판단으로 unload 후 제거
- 오늘 발송 검증
  - `python3 scripts/dispatch_route_sms.py` → workflow dispatch 생성
  - `gh run watch 24322028775` → success
  - 로그 확인 결과 `[OK] 기사님 SMS 발송 완료`, `[OK] 오너 동선 SMS 발송 완료`
- 중복 방지 검증
  - 오늘 성공 run 후 `python3 scripts/dispatch_route_sms.py` 재실행 → `already sent today`로 skip
  - `python3 scripts/dispatch_route_sms.py --check-only` 재실행 → 동일하게 skip

## Independent Review
- Contract met: 예. 오늘 발송 복구, 같은 날 중복 방지, 클라우드-only 운영 기준 정리, 운영 문서 반영까지 모두 끝냈다.
- Out-of-scope preserved: 예. 월·목 발송 규칙, 동선 순서, 숙소/정산 로직은 건드리지 않았다.
- Verification evidence present: 예. 지연 run 기록, 수동 발송 success run, 중복 skip, launchd 제거까지 남겼다.
- Remaining risk: GitHub schedule 지연 자체는 플랫폼 특성상 남아 있다. 다만 최근 run은 모두 점심 전에 끝났고, 사용자 기준에도 부합한다.

## Done Gate
- Success criteria: 충족. 오늘 문자는 이미 발송됐고, 이후에는 GitHub Actions만으로 운영하면서 같은 날 중복 발송은 막는다.
- Verification evidence: GitHub run `24322028775` success + 중복 skip 검증 + launchd 제거 완료.
- Remaining risks: GitHub schedule이 10시 정각보다 늦을 수는 있다. 다만 최근 기록은 전부 점심 전이었다.
- Next action: 종료. 이후 운영 경로는 다시 GitHub Actions only다.
