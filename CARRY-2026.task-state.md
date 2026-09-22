# 왕산로 8월 이후 동선 유지

## 작업 계약

- 현재 상태: 2026/07/01부터 왕산로를 월·목 동선에 계속 포함하도록 구현했다.
- 목표 상태: 2026/08/01 이후에도 7월과 동일하게 왕산로가 매주 월·목 동선에 포함된다.
- 이번 범위: 왕산로 포함 날짜 계산, 관련 회귀 테스트, 왕산로 운영 기준 문구만 수정한다.
- 범위 밖: 이태원·정산·입력 화면·문자 발송·Supabase·GitHub workflow 수동 실행.
- 완료 기준: 07/30, 08/03, 08/06, 08/10, 08/17, 08/20에 왕산로가 포함되고, 06/01에는 포함되며 06/08에는 빠진다.
- 배포 기준: 로컬 테스트와 dry-run 메시지 검증 후 변경 커밋을 원격 기본 브랜치에 반영하고, 반영된 소스 버전을 확인한다.
- 안전 기준: 실제 문자 발송과 외부 데이터 변경은 하지 않는다.

## 지침 적용 확인

- 읽은 문서: `/Users/wjh/AGENTS.md`, `/Users/wjh/.ai-context/CODEX.md`, `/Users/wjh/.ai-context/rules/session.md`, `/Users/wjh/.ai-context/rules/pm-human-communication.md`, `/Users/wjh/.ai-context/rules/operating-change.md`, `/Users/wjh/.ai-context/rules/development.md`, `/Users/wjh/.ai-context/rules/release-qa.md`, 저장소 `AGENTS.md`, `docs/operations.md`, `docs/project-context-legacy.md`, 왕산로·8월 작업 상태 문서.
- 이번 적용: 운영 변경은 dry-run과 날짜 재현으로 검증하고, 기존 사용자 변경을 보존하며, 왕산로 관련 파일만 커밋한다.
- 다시 읽을 조건: 범위가 왕산로 외 운영 규칙으로 넓어지거나 실제 발송·Supabase 변경을 요청받는 경우.

## 검증 기록

- 테스트 기준을 먼저 8월 월·목 포함으로 변경했다.
- `pytest -q tests/test_send_route_sms.py`: 실행 파일 미설치로 실행 불가.
- `python3 -m unittest tests.test_send_route_sms -v`: 저장소 테스트 패키지 구조상 모듈 경로 import 오류로 실행 불가. `unittest discover`로 재실행한다.
- `python3 -m unittest discover -s tests -p 'test_*.py' -v`: 21개 통과.
- `python3 -m py_compile scripts/send_route_sms.py scripts/generate_invoices.py scripts/dispatch_route_sms.py`: 통과.
- `git diff --check`: 통과.
- `DRY_RUN=true TEST_DATE=2026-06-01/06-08/07-30/08-03/08-06/08-10/08-17/08-20`: 06/01·07/30·08/03 이후 월·목에는 왕산로 포함, 06/08에는 제외, 08월에는 다음 일정 안내 미표시를 확인했다.
- `main` 기준 배포 worktree에서 `python3 -m unittest discover -s tests -p 'test_*.py' -v`: 35개 통과.
- 배포 worktree에서 `py_compile` 3개 스크립트와 `git diff --check`: 통과.
- 배포 worktree `DRY_RUN=true TEST_DATE=2026-06-01/06-08/07-30/08-03/08-06/08-10/08-17/08-20`: 06/01·07/30·08/03 이후 월·목 포함, 06/08 제외, 08월 다음 일정 미표시 확인.

## 독립 검토 및 배포 기록

- 현재 변경과 `main` 기준 배포 패치 모두 독립 Codex 리뷰 `GO`를 받았다.
- 커밋 범위: `docs/operations.md`, `docs/project-context-legacy.md`, `scripts/send_route_sms.py`, `tests/test_send_route_sms.py` 4개 파일.
- 배포 PR: 개인 레포 `jeonghoon0126/laundry-form` PR #6, `main`에 squash merge 완료.
- 배포된 원격 `main`: `eb2ec0a`; 원격 소스에서 `WANGSANRO_DUAL_ROUTE_START`와 7월 이후 월·목 조건을 확인했다.
- PR 브랜치의 최신화 과정에서 기존 `main` 변경을 merge했지만, 최종 PR diff는 왕산로 관련 4개 파일만 남겼다.
- 실제 문자 발송, GitHub Actions workflow dispatch, Supabase·Google Sheets 변경은 실행하지 않았다.

## 진행 상태

- 상태: 완료 (왕산로 유지 + 이태원 07/30 동선 시작 정정)
- 완료 시각: 2026-07-29, PR #6·#7 squash merge 및 원격 `main` 소스 확인 후.
- 다음 행동: 없음. 다음 월·목 예약 실행이 `main`의 `dca58b2` 소스를 사용한다.
- 남은 운영 위험: push 시 workflow가 실행되는 구조가 아니라 월·목 schedule만 실행되므로, 배포 후 실제 배치 run과 실제 문자 발송은 아직 관찰하지 않았다.

## 운영 기준 정정: 이태원 동선 시작일

- 사용자 정정: 이태원은 이번 주 목요일인 2026-07-30(목)부터 동선에 포함한다.
- 변경 범위: 동선 계산 상수와 관련 운영 문서·경계 테스트만 변경했다.
- 유지 기준: 입력 화면 `ITAEWON_START_DATE`와 정산 `ITAEWON_SETTLEMENT_START_DATE`는 모두 2026-08-01로 유지했다.
- 검증: 07/23 DRY_RUN은 이태원 미포함, 07/30과 08/03 DRY_RUN은 이태원을 포함했고, 07/30 순서는 강남 → 송파 → 건대 → 왕산로 → 회기 → 제기동 → 장충동 → 이태원 → 연남이다.
- 테스트: `python3 -m unittest discover -s tests -p 'test_*.py'` 36개 통과, 관련 스크립트 `py_compile` 통과, diff 공백 검사 통과.
- 배포: 개인 레포 `jeonghoon0126/laundry-form` PR #7을 squash merge했고, 원격 `main`은 `dca58b2`에서 2026-07-30 이태원 동선 기준을 확인했다.
- 미실행: 실제 SMS 발송, GitHub Actions 수동 dispatch, Supabase·Google Sheets 변경은 하지 않았다.
