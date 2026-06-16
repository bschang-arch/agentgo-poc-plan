# Gemini 작업 컨텍스트 — 수행계획서 자동 작성 에이전트

> 이 문서는 Gemini가 본 프로젝트의 현재 상태를 빠르게 이해하고 개발을 이어가도록
> 핵심만 정리한 브리핑이다. 상세는 `PRD.md` / `template-spec.md` / `KICKOFF.md` 참조.

## 한 줄 요약
회사 표준 워드 템플릿(.docx)의 **서식을 보존한 채** 입력 JSON의 가변 값만 주입해
PoC 수행계획서를 자동 생성하는 CLI 도구. + Gemini 연계로 입력 초안 생성·개발 보조.

## 절대 원칙 (위반 금지)
1. **개인정보·고객사 정보를 외부(Gemini 등)로 원문 전송 금지.** 사용자가 통제하는
   입력(brief)만 전송한다. 기준 템플릿(.docx)·생성본·실제 inputs.json은 전송하지 않는다.
2. **항상 템플릿에서 새로 생성**한다(이전 고객사 값 잔존 방지). 생성본을 다시 열어
   일부만 패치하는 방식은 금지.
3. **비밀키 하드코딩 금지** — `.env`/`*.env`(gitignore)에서 환경변수로 로드.
4. 새 외부 의존성은 **이유 + 대안**을 한 줄로 밝힌다.
5. 코드 주석은 '무엇'보다 **'왜'**. 응답·문서는 한국어.

## 핵심 파일
| 파일 | 역할 |
|---|---|
| `generate_plan.py` | 생성기 본체. 입력 JSON → 템플릿 치환 → .docx. 변경이력/diff, 고객사명 안전장치 포함 |
| `draft_inputs.py` | 자유 서술 → 입력 JSON 초안 (Gemini, JSON 출력 모드) |
| `call_gemini.py` | Gemini 호출 + `.env` 키 로드. `load_api_key()`를 다른 스크립트가 재사용 |
| `gemini_dev.py` | 컨텍스트+소스+질문을 Gemini에 보내는 개발 보조 CLI (민감파일 전송 차단) |
| `template-spec.md` | 입력값 ↔ 문서 섹션/셀 매핑 사양 (0장 입력 정의 A/B/C) |
| `inputs.example.json` | 입력 예시 (가명). 실제 값은 `inputs.json`(gitignore) |

## 입력 스키마
- **필수(A) 8개**: `client_name`, `start_date`(YYYY-MM-DD), `total_weeks`(int),
  `poc_purpose`, `verify_tasks`[], `integration_targets`{sso, legacy_system, m365},
  `solution_config`[], `client_pm`(연락처 마스킹).
- **선택(B)**: `kpi_targets`, `infra_os`, `use_nas`, `meeting_cycle`, `author`, `approver`,
  `_client_name_variants`[], `_template_revision_authors`[].

## 생성기 동작 핵심 (generate_plan.py)
- **고객사명 치환**: `_client_name_variants`의 모든 표기를 센티넬 경유로 `client_name`으로
  교체. **치환 0건이면 저장 중단**(이전 고객사명 잔존=정보유출 위험). 가명 테스트는
  `--allow-no-client-name`. 변형별 매칭 건수를 출력해 오타를 잡는다.
- **변경 이력/재생성**: 재실행 시 직전 입력과의 diff를 출력하고
  `<output>.history.json`(out/ 아래, gitignore)에 누적. `--show-history`로 조회, `--no-history`로 생략.
- **일정**: 'N주차' 라벨을 `start_date` 기준 금요일로 재계산. 단계 '수'는 고정.
- **개정이력**: v1.0 1행만 남기고 이전 작성자/승인자(실명) 제거.
- **검토 필요 항목**: 본문 오염 방지 위해 콘솔로만 출력.

## 진행 상태 (2026-06-16 기준)
- ✅ MVP: 입력 → 가변값 주입 → 표준 .docx 출력
- ✅ 고객사명 0건 안전장치
- ✅ Gemini 키 로드 + 호출(`call_gemini.py`)
- ✅ 자유 서술 → 입력 초안(`draft_inputs.py`)
- ✅ 수정·재생성 + 변경 이력
- ⏳ 다음: 빠진 값 체크 강화(선택 입력 B 되묻기), 출력 포맷 확장(HWP/PDF)

## 환경/도구 주의 (Windows)
- Python 3.14, `python-docx` 1.2.0, `google-genai`, `python-dotenv`.
- 콘솔이 cp949라 한글이 깨질 수 있음 → `PYTHONIOENCODING=utf-8`, 스크립트는 stdout을 UTF-8로 재설정.
- 커밋 메시지에 큰따옴표가 있으면 PowerShell이 인자를 쪼갬 → `git commit -F <파일>` 사용.
- GitHub: `bschang-arch/agentgo-poc-plan` (private). 흐름: add → commit(-F) → push.
