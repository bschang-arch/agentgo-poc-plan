"""Gemini 개발 보조 CLI.

프로젝트 컨텍스트(GEMINI_CONTEXT.md) + 선택한 소스 파일 + 질문을 Gemini에 보내
코드/설계 답을 받는다. 이 스크립트를 '직접' 터미널에서 돌리면 개발 질의가
Gemini로 처리되어(=Gemini 토큰 사용) Claude 세션 비용을 들이지 않는다.

설계 메모(왜):
- 프로젝트 원칙상 개인정보·고객사 원문은 외부로 보내지 않는다. 그래서 민감 파일
  (.env/*.env, *.docx, 실제 inputs.json, 이력 파일)은 --file 로 줘도 전송을 거부한다.
- 컨텍스트는 GEMINI_CONTEXT.md 한 장으로 고정해 매 호출 토큰을 절약한다.

사용:
    python gemini_dev.py "generate_plan.py에 PDF 출력 옵션을 어떻게 추가할까?"
    python gemini_dev.py --file generate_plan.py "rebuild_schedule_dates 리팩터링 제안"
    python gemini_dev.py --file draft_inputs.py --file call_gemini.py "중복 로직 정리안"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from call_gemini import load_api_key  # 키 로드 로직 공유

# 외부(Gemini)로 절대 보내면 안 되는 파일 패턴 — 프로젝트 프라이버시 원칙.
BLOCKED_GLOBS = ["*.docx", "*.doc", "*.hwp", "*.hwpx", ".env", "*.env",
                 "inputs.json", "inputs.*.json", "*.history.json"]
# 단, 가명 예시는 허용.
ALLOWED_EXCEPTIONS = {"inputs.example.json"}

SYSTEM_INSTRUCTION = """\
너는 이 프로젝트의 시니어 개발 보조자다. 아래 '프로젝트 컨텍스트'의 원칙과 구조를
반드시 지켜 답한다. 특히: 개인정보·고객사 원문을 다루지 말 것, 항상 템플릿에서 새로
생성하는 설계를 깨지 말 것, 비밀키 하드코딩 금지, 새 의존성은 이유+대안 명시.
답변은 한국어로, 구체적인 코드/디프와 근거를 함께 제시한다. 불확실하면 단정하지 말 것.
"""


def is_blocked(path: Path) -> bool:
    if path.name in ALLOWED_EXCEPTIONS:
        return False
    return any(path.match(g) for g in BLOCKED_GLOBS)


def build_file_context(files: list[str]) -> str:
    blocks = []
    for f in files:
        p = Path(f)
        if is_blocked(p):
            sys.exit(f"전송 거부(민감 파일): {f}\n"
                     "  개인정보·고객사 데이터 또는 비밀키가 들어갈 수 있어 외부로 보내지 않습니다.")
        if not p.exists():
            sys.exit(f"파일 없음: {f}")
        text = p.read_text(encoding="utf-8")
        blocks.append(f"### 파일: {f}\n```\n{text}\n```")
    return "\n\n".join(blocks)


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    ap = argparse.ArgumentParser(description="Gemini 개발 보조 CLI")
    ap.add_argument("prompt", help="개발 질문/요청")
    ap.add_argument("--file", action="append", default=[],
                    help="함께 보낼 소스 파일 (반복 가능). 민감 파일은 자동 차단")
    ap.add_argument("--context", default="GEMINI_CONTEXT.md",
                    help="프로젝트 컨텍스트 문서 (기본: GEMINI_CONTEXT.md)")
    ap.add_argument("--env-file", default=None, help="키가 든 env 파일(기본: 자동 탐색)")
    ap.add_argument("--model", default="gemini-2.5-flash", help="사용할 모델")
    args = ap.parse_args()

    context_path = Path(args.context)
    context = context_path.read_text(encoding="utf-8") if context_path.exists() else ""
    if not context:
        print(f"※ 컨텍스트 문서를 찾지 못함: {args.context} (컨텍스트 없이 진행)", file=sys.stderr)

    file_context = build_file_context(args.file)

    api_key = load_api_key(args.env_file)
    try:
        from google import genai
        from google.genai import types
    except ImportError:
        sys.exit("google-genai가 필요합니다. 먼저: pip install -r requirements.txt")

    parts = ["## 프로젝트 컨텍스트\n" + context]
    if file_context:
        parts.append("## 관련 소스\n" + file_context)
    parts.append("## 요청\n" + args.prompt)
    contents = "\n\n".join(parts)

    client = genai.Client(api_key=api_key)
    try:
        resp = client.models.generate_content(
            model=args.model,
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=SYSTEM_INSTRUCTION),
        )
    except Exception as e:
        sys.exit(f"Gemini 호출 실패: {type(e).__name__}: {e}")

    if args.file:
        print(f"[전송 파일: {', '.join(args.file)}]\n", file=sys.stderr)
    print(resp.text)


if __name__ == "__main__":
    main()
