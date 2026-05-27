# AGENTS.md - laundry-form

이 파일은 `laundry-form`에서만 적용되는 작업 시작용 지침이다. 상위 `/Users/wjh/AGENTS.md`와 `.ai-context` 규칙을 약화하거나 대체하지 않는다.

## 역할

- 캐리 세탁 정산과 동선 문자 운영을 다루는 개인 계정 프로젝트다.
- 화면은 `index.html`, 운영 스크립트는 `scripts/` 아래 Python 파일을 기준으로 본다.
- 숙소, 단가, 사업자 매핑, 발송 방식의 상세 기준은 `docs/operations.md`와 `docs/project-context-legacy.md`를 참고한다.

## 시작 전 확인

- 루트 지침, 활성 Linear description, 대응 `.task-state.md`를 먼저 읽는다.
- 문자 발송, GitHub Actions, Supabase, 정산 금액을 다루면 API/운영 영향 작업으로 분류한다.
- 날짜는 `MM/DD` 형식으로 정리하되, 스크립트 입력값은 기존 형식을 따른다.

## 금지

- 사용자 명시 없이 실제 문자 발송, GitHub workflow dispatch, Supabase 데이터 변경, 정산 결과 발행을 하지 않는다.
- 테스트 없이 단가, 사업자 매핑, 숙소 동선을 바꾸지 않는다.
- 운영 비밀값이나 전화번호를 새 문서에 풀어 쓰지 않는다.

## 검증

- 문자/정산 변경은 dry-run 또는 `--check-only`로 먼저 확인한다.
- 날짜 의존 로직은 `TEST_DATE`로 재현 가능한 케이스를 남긴다.
- 실제 발송이 필요한 경우 사람 실행 절차와 중복 발송 방지 기준을 먼저 남긴다.
