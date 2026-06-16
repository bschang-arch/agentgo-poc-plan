# 온보딩 가이드 — 수행 계획서 자동 작성 에이전트

> 처음 합류하는 개발자가 이 프로젝트를 이해하고 곧바로 이어서 작업할 수 있도록
> 정리한 안내문서다. 더 깊은 내용은 `PRD.md`(요구사항), `template-spec.md`(매핑 사양),
> `KICKOFF.md`(착수 지시), `GEMINI.md`(AI 작업 컨텍스트)를 참고한다.

---

## 1. 프로젝트 개요와 해결하는 문제

회사 표준 워드 템플릿(`.docx`)으로 매번 손수 수행 계획서를 쓰는 일은 번거롭고,
이전 사업 문서를 복사해 고치다 보면 **이전 고객사명·날짜가 남아 유출되는 사고**가
나기 쉽다.

이 도구는 **필수 값 몇 개를 입력하면, 회사 표준 템플릿의 서식을 그대로 보존한 채
가변 값만 채워** 수행 계획서(PoC 계획서)를 자동 생성한다. 추가로 Gemini를 연계해
자유 서술로부터 입력 초안을 만들고, 개발 자체도 보조한다. 클로드 코드/일반 터미널
(Python CLI) 환경에서 동작한다.

핵심 난제: **기준 템플릿이 빈 양식이 아니라 특정 고객 건으로 채워진 사례 문서**다.
따라서 "고정 골격 / 사업별 가변 값"을 구분해, 가변 값만 안전하게 교체해야 한다.

---

## 2. 아키텍처와 파일별 역할

| 파일 | 역할 |
|---|---|
| `generate_plan.py` | **생성기 본체.** 입력 JSON → 템플릿 치환 → `.docx` 출력. 고객사명 안전장치, 일정 재계산, 변경 이력, PDF/HWP 변환 옵션 포함 |
| `convert.py` | (선택) 생성된 `.docx` → PDF(Word COM)/HWP(한컴 COM) 변환. 서식 보존 |
| `register_hwp.py` | (선택) HWP 자동화용 한컴 보안 모듈 1회 등록 (pyhwpx의 DLL 사용) |
| `draft_inputs.py` | (선택) 자유 서술 → 입력 JSON 초안 생성 (Gemini, JSON 출력 모드) |
| `call_gemini.py` | Gemini API 호출 예제 + `.env` 키 로드. `load_api_key()`를 다른 스크립트가 재사용 |
| `gemini_dev.py` | (선택) Gemini 개발 보조 CLI. 컨텍스트+소스+질문 전송, 민감 파일 차단 |
| `GEMINI.md` | 프로젝트 컨텍스트 브리핑. `gemini_dev.py`와 Gemini CLI가 함께 읽음 |
| `template-spec.md` | 입력값 ↔ 문서 섹션/셀 매핑 사양 (0장: 입력 정의 A/B/C) |
| `PRD.md` / `KICKOFF.md` | 제품 요구사항 / 개발 착수 지시 |
| `inputs.example.json` | 입력 예시 (가명). 복사해 `inputs.json`으로 실제 값 채워 사용 |
| `.env.example` | 비밀키 파일 형식 예시 (실제 키 없음) |
| `requirements.txt` | 의존성 (python-docx, google-genai, python-dotenv) |

> 기준 템플릿 `.docx`는 **저장소에 포함되지 않는다**(민감정보 → `.gitignore`).
> 회사 표준 템플릿 파일을 직접 준비해 같은 폴더에 두고 `--template`으로 지정한다.

---

## 3. 데이터 흐름

**(A) 기본 생성 흐름**
```
inputs.json  ──►  generate_plan.py  ──►  out/수행계획서_생성본.docx
(필수 8 + 선택)      │  ├ 고객사명 치환 (변형 목록 → 센티넬 → client_name)
                    │  ├ 표지 작성일/버전, 개정이력 v1.0 1행화(이전 실명 제거)
                    │  ├ 검증 대상 솔루션 문단 재작성
                    │  ├ 고객 측 비상연락망 신설
                    │  └ 일정표 'N주차' → start_date 기준 날짜 재계산
                    └► out/<output>.history.json (변경 이력 누적)
```

**(B) Gemini 입력 초안 흐름**
```
자유 서술(brief) ──► draft_inputs.py ──► inputs.draft.json ──(사람 검토·보정)──► inputs.json
```

**(C) Gemini 개발 보조 흐름**
```
GEMINI.md + 소스파일 + 질문 ──► gemini_dev.py / Gemini CLI ──► 코드·설계 답
```

---

## 4. 핵심 설계 원칙과 제약 (반드시 지킬 것)

1. **개인정보·고객사 정보를 외부(Gemini 등)로 원문 전송 금지.** 사용자가 통제하는
   입력(brief)만 전송한다. 기준 템플릿(.docx)·생성본·실제 `inputs.json`은 전송하지 않는다.
   (`gemini_dev.py`는 `.docx`/`inputs.json`/`*.env`/이력 파일을 자동 차단한다.)
2. **항상 템플릿에서 새로 생성**한다(이전 고객사 값 잔존 방지). 생성본을 다시 열어
   일부만 패치하는 방식은 금지.
3. **고객사명 0건 안전장치**: 치환이 0건이면(입력 변형이 템플릿 표기와 불일치)
   템플릿의 이전 고객사명이 그대로 남아 유출되므로 **저장하지 않고 중단**한다.
   가명 예시·테스트만 `--allow-no-client-name`으로 우회한다.
4. **비밀키 하드코딩 금지** — `.env`/`*.env`(gitignore)에 `GEMINI_API_KEY=...`로 두고
   환경변수로 로드한다.
5. 새 외부 의존성은 **이유 + 대안**을 한 줄로 밝힌다. 코드 주석은 '무엇'보다 **'왜'**.

---

## 5. 환경 설정 / 설치

- **OS**: Windows 11. 프로젝트 경로는 OneDrive 동기화 폴더 안에 있다.
- **Python**: 3.14 (`python` / `py` 사용 가능). `python-docx` 1.2.0.
- **설치**:
  ```bash
  pip install -r requirements.txt
  ```
- **Gemini 키**: `.env.example`를 복사해 `.env`로 만들고 `GEMINI_API_KEY=...`를 채운다.
  `.env`/`*.env`는 `.gitignore`로 커밋 제외된다. (`GOOGLE_API_KEY`도 인식)
- **Node/Gemini CLI**(선택): Node v24+. `npm install -g @google/gemini-cli`.

---

## 6. 실행 방법

**문서 생성**
```bash
python generate_plan.py \
  --template "회사표준_수행계획서_템플릿.docx" \
  --input inputs.json \
  --output "out/수행계획서_생성본.docx"
```

**입력 초안 생성 (Gemini)** — 민감정보는 brief에 넣지 말 것
```bash
python draft_inputs.py --brief "고객사 PoC 6주, 7/6 시작, M365·기간계 연동, RAG 검증"
# → inputs.draft.json (모르는 값은 "[확인 필요]"로 표시) → 검토·보정 후 사용
```

**개발 보조 (Gemini Q&A)** — 민감 파일은 자동 차단
```bash
python gemini_dev.py --file generate_plan.py "PDF 출력 옵션 추가 방법은?"
```

**PDF / HWP 출력** — 서식 보존을 위해 설치된 오피스로 `.docx`를 변환
```bash
python generate_plan.py ... --pdf --hwp           # 생성과 동시에 변환
python convert.py "out/수행계획서_생성본.docx" --pdf   # 기존 docx만 변환
```
> PDF=MS Word(✅ 검증됨). HWP=한컴오피스 필요 + `pip install pyhwpx && python register_hwp.py`(보안 모듈 1회 등록).
> 단 한컴 2018은 COM `Open`이 외부 포맷(docx 등)을 못 열어 자동 변환이 안 됨 → 한컴 GUI에서 수동 "HWP로 저장" 권장.

**에이전트형 코딩 (Gemini CLI)** — `.env`(키)·`GEMINI.md`(컨텍스트) 자동 인식
```bash
gemini                 # 대화형 세션 (편집·실행까지 수행)
gemini -p "다음 작업 제안하고 구현해줘"   # 비대화형 1회
```
> 최초 실행 시 "신뢰된 폴더" 확인 → 이 `project` 폴더만 신뢰(상위 폴더 신뢰는 금지).
> 헤드리스 자동 실행은 `GEMINI_CLI_TRUST_WORKSPACE=true`.

---

## 7. 입력 스키마

**필수(A) 8개** — 하나라도 비면 생성 불가
| 키 | 형식 | 설명 |
|---|---|---|
| `client_name` | text | 고객사명 |
| `start_date` | `YYYY-MM-DD` | 착수일 |
| `total_weeks` | int | 총 수행 주차 |
| `poc_purpose` | text | PoC 목적·배경 |
| `verify_tasks` | list | 검증 과제 목록 |
| `integration_targets` | object | `{sso, legacy_system(bool), m365(bool)}` |
| `solution_config` | list | 솔루션 구성 (제품·User·LLM·vCPU) |
| `client_pm` | text | 고객 측 PM (연락처 **마스킹**: 직무/등급) |

**선택(B)**: `kpi_targets`, `infra_os`, `use_nas`, `meeting_cycle`, `author`, `approver`.
**메타**: `_client_name_variants`(템플릿에 박힌 고객사명의 **모든 표기 변형** — 한 표기만
넣으면 다른 표기가 누락됨), `_template_revision_authors`(개정이력에서 제거할 이전 실명).

상세는 `template-spec.md` 0장 참조.

---

## 8. 수정·재생성과 변경 이력

입력을 고쳐 다시 실행하면 **항상 템플릿에서 새로 생성**하며, 직전 입력과 달라진
부분만 **diff**로 보여준다. 변경 로그는 출력 문서별 `<output>.history.json`
(`out/` 아래 → 커밋 제외)에 누적된다.
```bash
# 재실행 시 자동 출력 예
- 변경 이력: 직전 대비 2건 변경
    · total_weeks: 6 → 8
    · meeting_cycle: 격주 → 주간

python generate_plan.py --output "out/...docx" --show-history   # 과거 이력 조회
python generate_plan.py ... --no-history                         # 이번 실행 기록 생략
```
- diff 대상은 docx가 아니라 **입력 JSON**(의미 명확 + 로컬에만 머물러 안전).
- 주석성 키(`_comment`, `*_note`)는 비교에서 제외된다.

---

## 9. 알려진 한계

- 치환 구간 내부의 부분 서식(굵게·색상)은 첫 런 서식으로 통일될 수 있음 → 출력본 육안 확인 권장.
- 일정은 단계 '수'는 고정하고 날짜만 `start_date` 기준으로 재계산 (단계 구성 변경은 범위 밖).
- KPI·회의주기 등 표기 변형이 큰 항목은 자동 치환하지 않고 콘솔에 보고만 함.
- Gemini 무료 등급은 **하루 요청 수 제한**(모델별, 예: 20/일)이 있어 에이전트형 CLI는
  쉽게 소진된다. 막히면 다음 날 리셋 또는 유료 등급으로 전환.
- HWP 변환: `register_hwp.py`로 보안 모듈을 등록하면 보안창 멈춤은 해결되나, **한컴 2018은
  COM `Open`이 .docx/.doc/.rtf 등 외부 포맷을 열지 못해**(SaveAs는 정상) 자동 변환이 안 된다.
  → 생성된 `.docx`/PDF를 한컴 GUI에서 열어 "다른 이름으로 저장 → HWP". (다른 한컴 빌드에선 COM 변환이 될 수 있음)

---

## 10. 다음 작업 로드맵

- ✅ MVP: 입력 → 가변값 주입 → 표준 `.docx` 출력
- ✅ 고객사명 0건 안전장치
- ✅ Gemini 연계: 키 로드/호출, 자유 서술 → 입력 초안, 개발 보조(gemini_dev / Gemini CLI)
- ✅ 수정·재생성 + 변경 이력
- ✅ 빠진 값 체크: 필수값 누락·미확정·형식오류 일괄 검증(차단), 선택값 미입력 알림 (`validate_inputs`)
- ✅ 출력 포맷 확장: `--pdf`(Word)/`--hwp`(한컴) 변환 (`convert.py`)
- ⏳ 다음(후보): HWP 보안 모듈 자동 등록, 일정 단계 구성 편집, 부분 서식 보존 개선

---

## 11. 개발 워크플로

- **GitHub**: `bschang-arch/agentgo-poc-plan` (private). 흐름: `git add` → `commit` → `push`.
- **커밋 메시지 주의**: PowerShell에서 메시지에 큰따옴표가 있으면 인자가 쪼개져 실패한다.
  한글·기호가 많은 메시지는 파일로 작성해 `git commit -F <파일>`로 커밋한다.
- **콘솔 인코딩**: Windows 콘솔이 cp949라 한글이 깨질 수 있다. 실행 전
  `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`, 필요 시 `PYTHONIOENCODING=utf-8`.
  스크립트들은 stdout을 UTF-8로 재설정한다.
- **PATH**: `git`/`gh`는 사용자 PATH에 등록되어 새 터미널에서 바로 쓸 수 있다.
- **커밋 금지 대상**(`.gitignore`): `*.docx`/`*.hwp*`, `out/`, `inputs.json`/`inputs.*.json`
  (단 `inputs.example.json`은 허용), `.env`/`*.env`(단 `.env.example` 허용), `.claude/`, `*.skill`.
  → 실제 고객사·개인정보·비밀키는 절대 커밋하지 말 것.
