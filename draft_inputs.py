"""자유 서술 → 수행계획서 입력 JSON 초안 생성 (Gemini 연계).

사용자가 사업 개요를 자유롭게 적으면, Gemini가 generate_plan.py 가 요구하는
입력 스키마(필수 8개 + 선택)에 맞춰 JSON '초안'을 만든다.

설계 메모(왜):
- 입력 JSON 손수 작성이 번거롭고 키·형식을 틀리기 쉬워, 자유 서술을 구조화해 마찰을 줄인다.
- 어디까지나 '초안'이다. 날짜·연락처·솔루션 스펙 등은 사람이 반드시 검토/보정한다.
  그래서 결과를 inputs.json 이 아니라 inputs.draft.json 에 쓴다(실사용 파일 덮어쓰기 방지).
- 프라이버시: brief 텍스트는 Gemini(외부 API)로 전송된다. 실명·연락처 등 민감정보는
  직접 넣지 말고 직무·등급으로 마스킹해 적을 것(README 원칙과 동일).
- 모델이 형식을 어기지 않도록 JSON 출력 모드(response_mime_type)를 강제하고,
  모르는 값은 "[확인 필요]"로 표시하게 해 검토 지점을 드러낸다.

사용:
    python draft_inputs.py --brief "AgentGo PoC 6주, 7/6 시작, M365·기간계 연동, RAG 검증"
    python draft_inputs.py --brief-file brief.txt --output inputs.draft.json
    echo "사업 개요..." | python draft_inputs.py        # stdin 도 가능
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# 키 로드 로직은 call_gemini 와 공유(중복 방지).
from call_gemini import load_api_key

# generate_plan 이 필수로 검사하는 A그룹 8개 키 — 초안 검증에 재사용.
REQUIRED_KEYS = ["client_name", "start_date", "total_weeks", "poc_purpose",
                 "verify_tasks", "integration_targets", "solution_config", "client_pm"]

SYSTEM_INSTRUCTION = """\
너는 IT PoC '수행 계획서'의 입력값을 구조화하는 도우미다.
사용자의 자유 서술에서 아래 JSON 스키마에 맞는 값을 추출/요약해 채운다.

규칙:
- 반드시 아래 키 구조의 JSON 하나만 출력한다(설명·마크다운 금지).
- 서술에 없는 값은 추측하지 말고 문자열은 "[확인 필요]", 불리언은 false 로 둔다.
- start_date 는 "YYYY-MM-DD" 형식. 연도가 없으면 서술 맥락의 가장 가능성 높은 연도를
  쓰되 불확실하면 "[확인 필요]".
- total_weeks 는 정수. 명시 없으면 6.
- client_pm 등 인물은 실명·연락처가 보여도 직무/등급+마스킹 형태로만 적는다
  (예: "고객사 PM [담당 차장 / 010-****-**** / name@example.com]").
- verify_tasks, solution_config 는 문자열 배열. 항목이 하나여도 배열로.
- _client_name_variants 에는 client_name 의 표기 변형(영문/약칭/'OO사' 등)을 추정해 모두 넣는다.
- _notes 에는 네가 가정했거나 사람이 꼭 확인해야 할 항목을 한국어로 짧게 나열한다.

JSON 스키마(이 키들만, 이 형태로):
{
  "client_name": "string",
  "start_date": "YYYY-MM-DD",
  "total_weeks": 6,
  "poc_purpose": "string(1~2문장)",
  "verify_tasks": ["string", ...],
  "integration_targets": {"sso": "string", "legacy_system": false, "m365": false},
  "solution_config": ["string", ...],
  "client_pm": "string",
  "kpi_targets": {"accuracy": "string", "usability": "string"},
  "infra_os": "string",
  "use_nas": true,
  "meeting_cycle": "string",
  "_client_name_variants": ["string", ...],
  "_notes": ["string", ...]
}
"""


def read_brief(args) -> str:
    if args.brief:
        return args.brief
    if args.brief_file:
        return Path(args.brief_file).read_text(encoding="utf-8")
    if not sys.stdin.isatty():
        data = sys.stdin.read().strip()
        if data:
            return data
    sys.exit("brief가 비어 있습니다. --brief, --brief-file, 또는 stdin 으로 사업 개요를 주세요.")


def generate_draft(api_key: str, brief: str, model: str) -> dict:
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        sys.exit("google-genai가 필요합니다. 먼저: pip install -r requirements.txt")

    client = genai.Client(api_key=api_key)
    try:
        resp = client.models.generate_content(
            model=model,
            contents=f"다음 사업 개요를 스키마에 맞춰 JSON으로 구조화해줘:\n\n{brief}",
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_INSTRUCTION,
                response_mime_type="application/json",
                temperature=0.2,
            ),
        )
    except Exception as e:  # 네트워크·인증·모델명 오류 등
        sys.exit(f"Gemini 호출 실패: {type(e).__name__}: {e}")

    raw = resp.text
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        sys.exit(f"모델이 유효한 JSON을 반환하지 않았습니다: {e}\n--- 원문 ---\n{raw}")


def validate(data: dict) -> list[str]:
    """초안의 누락·미확정 지점을 점검해 경고 목록을 만든다(차단은 하지 않음)."""
    warns: list[str] = []
    for k in REQUIRED_KEYS:
        if k not in data or data[k] in ("", None, []):
            warns.append(f"필수 키 '{k}' 가 비어 있음 — 직접 채워야 함")
    # 값이 "[확인 필요]" 또는 빈 문자열인 항목 수집
    def scan(prefix, obj):
        if isinstance(obj, dict):
            for kk, vv in obj.items():
                if kk == "_notes":  # 메모는 '확인 필요'라는 단어를 담으므로 스캔 제외(오탐 방지)
                    continue
                scan(f"{prefix}.{kk}" if prefix else kk, vv)
        elif isinstance(obj, list):
            for i, vv in enumerate(obj):
                scan(f"{prefix}[{i}]", vv)
        elif isinstance(obj, str) and "확인 필요" in obj:
            warns.append(f"'{prefix}' = [확인 필요] — 값 보정 필요")
    scan("", data)
    # start_date 형식 점검(generate_plan은 strptime으로 깐깐하게 받는다)
    sd = data.get("start_date", "")
    if sd and "확인 필요" not in sd:
        import re
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", sd):
            warns.append(f"start_date '{sd}' 형식이 YYYY-MM-DD 가 아님")
    return warns


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    ap = argparse.ArgumentParser(description="자유 서술 → 수행계획서 입력 JSON 초안 (Gemini)")
    ap.add_argument("--brief", help="사업 개요 자유 서술 텍스트")
    ap.add_argument("--brief-file", help="사업 개요가 담긴 텍스트 파일 경로")
    ap.add_argument("--output", default="inputs.draft.json",
                    help="초안 출력 경로 (기본: inputs.draft.json, .gitignore 대상)")
    ap.add_argument("--env-file", default=None, help="키가 든 env 파일 경로(기본: 자동 탐색)")
    ap.add_argument("--model", default="gemini-2.5-flash", help="사용할 모델")
    args = ap.parse_args()

    brief = read_brief(args)
    api_key = load_api_key(args.env_file)

    data = generate_draft(api_key, brief, args.model)
    warns = validate(data)

    out_path = Path(args.output)
    if out_path.exists():
        print(f"※ 기존 파일을 덮어씁니다: {out_path}", file=sys.stderr)
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"초안 생성: {out_path}")

    notes = data.get("_notes") or []
    if notes:
        print("\n[모델 메모 — 가정/확인 필요]")
        for n in notes:
            print(f"  - {n}")
    print("\n[검토 필요 — 사용 전 보정]")
    if warns:
        for w in warns:
            print(f"  - {w}")
    else:
        print("  - (자동 점검상 빈 값 없음 — 그래도 날짜·연락처·스펙은 사람이 확인)")
    print(f"\n다음: {out_path} 를 검토·보정 후 inputs.json 으로 복사해 generate_plan.py 에 사용하세요.")


if __name__ == "__main__":
    main()
