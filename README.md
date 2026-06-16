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
| `requirements.txt` | 의존성 |
| `inputs.example.json` | 입력 예시 (가명) — 복사해 실제 값으로 채워 사용 |
| `PRD.md` | 제품 요구사항 |
| `template-spec.md` | 입력값 ↔ 문서 섹션/셀 매핑 사양 |
| `KICKOFF.md` | 개발 착수 지시문 |

> 기준 템플릿 `.docx`는 저장소에 포함되지 않습니다. 회사 표준 템플릿 파일을 직접
> 준비해 같은 폴더에 두고 `--template` 으로 지정하세요.

## 설치
```bash
pip install -r requirements.txt
```

## 사용법
1. `inputs.example.json`을 복사해 `inputs.json`을 만들고 실제 값을 채웁니다.
   (`inputs.json`은 `.gitignore`로 커밋에서 제외됩니다.)
2. 실행:
```bash
python generate_plan.py \
  --template "회사표준_수행계획서_템플릿.docx" \
  --input inputs.json \
  --output "out/수행계획서_생성본.docx"
```
3. 실행 종료 시 콘솔에 출력되는 **[확인 필요]** 항목을 제출 전 검토하세요.

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

## 현재 범위 (MVP)
- ✅ 필수 값 입력 → 가변 항목 자동 작성 → 표준 .docx 출력
- ⏳ 다음: 수정·재생성, 빠진 값 체크, 출력 포맷 확장(HWP/PDF)

## 알려진 한계
- 치환 구간 내부의 부분 서식(굵게·색상)은 첫 런 서식으로 통일될 수 있음 → 출력본 육안 확인 권장
- 일정은 단계 '수'는 고정하고 날짜만 `start_date` 기준으로 재계산 (단계 구성 변경은 범위 밖)
