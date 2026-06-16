# 수행 계획서 자동 작성 에이전트

필수 값 몇 개만 입력하면, **회사 표준 워드 템플릿(.docx)의 서식을 보존한 채** 가변 값만
채워 수행 계획서를 자동 생성하는 도구입니다. 클로드 코드(CLI) 환경에서 동작합니다.

## ⚠️ 민감정보 주의
- 기준 템플릿(.docx)·생성본·실제 입력 파일에는 **고객사 정보와 개인정보**가 들어갑니다.
  이 저장소에는 커밋되지 않으며(`.gitignore` 참조), **절대 커밋하지 마세요.**
- 실명·연락처는 마스킹하거나 직무·등급으로 대체해 입력하는 것을 권장합니다.
- 이 저장소는 **비공개(private)** 운영을 전제로 합니다.

## 구성
| 파일 | 설명 |
|---|---|
| `generate_plan.py` | 생성기 본체 (python-docx 기반) |
| `draft_inputs.py` | (선택) 자유 서술 → 입력 JSON 초안 생성 (Gemini 연계) |
| `call_gemini.py` | Gemini API 호출 예제 (키는 `.env`/`*.env`에서 로드) |
| `requirements.txt` | 의존성 |
| `inputs.example.json` | 입력 예시 (가명) — 복사해 실제 값으로 채워 사용 |
| `.env.example` | 비밀키 파일 형식 예시 (실제 키 없음) |
| `PRD.md` | 제품 요구사항 |
| `template-spec.md` | 입력값 ↔ 문서 섹션/셀 매핑 사양 |
| `KICKOFF.md` | 개발 착수 지시문 |

> 기준 템플릿 `.docx`는 저장소에 포함되지 않습니다. 회사 표준 템플릿 파일을 직접
> 준비해 같은 폴더에 두고 `--template` 으로 지정하세요.

## 설치
```bash
pip install -r requirements.txt
```

## (선택) AI로 입력 초안 만들기 — `draft_inputs.py`
사업 개요를 자유롭게 적으면 Gemini가 입력 스키마에 맞춰 `inputs.draft.json` **초안**을 만들어 줍니다.
손으로 JSON을 쓰는 마찰을 줄이는 용도이며, 결과는 반드시 검토·보정 후 사용합니다.
```bash
python draft_inputs.py --brief "고객사 PoC 6주, 7/6 시작, M365·기간계 연동, RAG 검증"
# → inputs.draft.json 생성 (모르는 값은 "[확인 필요]"로 표시)
```
> ⚠️ `--brief` 텍스트는 **외부 API(Gemini)로 전송**됩니다. 실명·연락처 등 민감정보는
> 넣지 말고 직무·등급으로 마스킹해 적으세요. 키는 `.env`/`*.env`에서 읽습니다([Gemini 연계](#gemini-연계) 참조).

## 사용법
1. 입력 JSON 준비 — `inputs.example.json`을 복사해 `inputs.json`을 만들고 실제 값을 채웁니다.
   (`draft_inputs.py`로 만든 `inputs.draft.json`을 보정해 써도 됩니다. 두 파일 모두 `.gitignore`로 커밋 제외.)
2. 실행:
```bash
python generate_plan.py \
  --template "회사표준_수행계획서_템플릿.docx" \
  --input inputs.json \
  --output "out/수행계획서_생성본.docx"
```
3. 실행 종료 시 콘솔에 출력되는 **[확인 필요]** 항목을 제출 전 검토하세요.

### 수정·재생성 (변경 이력)
입력값을 고쳐 다시 실행하면 **항상 템플릿에서 새로 생성**하며(이전 값 잔존 방지),
**직전 생성과 달라진 입력만 diff로** 보여줍니다. 변경 로그는 출력 문서별 이력 파일
(`<output>.history.json`, `out/` 아래라 커밋 제외)에 누적됩니다.
```bash
# inputs.json 일부 수정 후 동일 명령으로 재실행 → 변경점이 출력됨
#   - 변경 이력: 직전 대비 2건 변경
#       · total_weeks: 6 → 8
#       · meeting_cycle: 격주 → 주간

python generate_plan.py --output "out/수행계획서_생성본.docx" --show-history  # 과거 이력 조회
python generate_plan.py ... --no-history                                      # 이번 실행은 기록 안 함
```

> **안전장치 — 고객사명 치환 0건 시 중단**: 입력한 `_client_name_variants` 가
> 템플릿의 실제 표기와 하나도 일치하지 않으면(치환 0건) **저장하지 않고 중단**합니다.
> 치환이 안 되면 템플릿에 박힌 *이전 고객사명*이 새 문서에 그대로 남아 정보가
> 유출되기 때문입니다. 콘솔의 변형별 매칭 건수(예: `ABC:2, ABC사:0`)로 0건 변형(오타)을
> 확인하세요. 가명 예시(`inputs.example.json`)로 동작만 볼 때는 `--allow-no-client-name`
> 를 붙이면 경고만 남기고 진행합니다(실제 문서 생성에는 사용 금지).

## 입력 값 (필수 8개)
`client_name`, `start_date`, `total_weeks`, `poc_purpose`, `verify_tasks`,
`integration_targets`, `solution_config`, `client_pm`
— 상세는 `template-spec.md` 0장 참조.

`_client_name_variants` 에는 템플릿에 박힌 고객사명의 **모든 표기 변형**을 넣어야
모든 위치가 치환됩니다(한 표기만 넣으면 다른 표기가 누락됨).

## Gemini 연계
입력 초안 생성(`draft_inputs.py`)과 호출 예제(`call_gemini.py`)에 쓰입니다.
- **키 보관**: API 키는 코드에 하드코딩하지 않고 `.env`(또는 임의 `*.env`) 파일에
  `GEMINI_API_KEY=...` 형식으로 둡니다. `.env`/`*.env`는 `.gitignore`로 커밋 제외되며,
  형식 예시는 `.env.example` 참조. (`GOOGLE_API_KEY` 도 인식)
- **설치**: `pip install -r requirements.txt` (google-genai, python-dotenv 포함)
- **호출 테스트**: `python call_gemini.py "한 문장으로 자기소개 해줘"`
- **모델**: 기본 `gemini-2.5-flash`, `--model` 로 변경. 사용 가능 모델은
  [Google AI Studio](https://aistudio.google.com)에서 확인하세요.
- **프라이버시**: Gemini는 외부 API이므로 전송 텍스트에 실명·연락처·고객 원문을
  넣지 마세요. 본 도구는 사용자가 입력한 brief만 보내며, 템플릿·생성본은 보내지 않습니다.

## 현재 범위 (MVP)
- ✅ 필수 값 입력 → 가변 항목 자동 작성 → 표준 .docx 출력
- ✅ Gemini 연계: 자유 서술 → 입력 JSON 초안 (`draft_inputs.py`)
- ✅ 수정·재생성: 입력 변경점 diff + 출력 문서별 변경 이력 (`--show-history`)
- ⏳ 다음: 빠진 값 체크 강화(선택 입력 B 되묻기), 출력 포맷 확장(HWP/PDF)

## 알려진 한계
- 치환 구간 내부의 부분 서식(굵게·색상)은 첫 런 서식으로 통일될 수 있음 → 출력본 육안 확인 권장
- 일정은 단계 '수'는 고정하고 날짜만 `start_date` 기준으로 재계산 (단계 구성 변경은 범위 밖)
