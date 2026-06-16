"""Gemini API 호출 예제.

.env 파일(또는 *.env)에 저장한 API 키를 환경변수로 읽어 Gemini를 호출한다.
키는 코드에 하드코딩하지 않으며, 파일은 .gitignore로 커밋에서 제외된다.

준비:
    pip install -r requirements.txt
    # 키 파일 예: 'gemini api key.env' 안에  GEMINI_API_KEY=AIza... 한 줄

사용:
    python call_gemini.py "한 문장으로 자기소개 해줘"
    python call_gemini.py --env-file "gemini api key.env" --model gemini-2.5-flash "질문"
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
from pathlib import Path


def load_api_key(env_file: str | None) -> str:
    """env 파일을 로드한 뒤 환경변수에서 키를 꺼낸다.

    설계 메모(왜):
    - 키 파일명이 표준 '.env'가 아닐 수 있어(예: 'gemini api key.env') 명시 경로 →
      '.env' → 디렉터리 내 임의 '*.env' 순으로 탐색한다.
    - 이미 셸 환경에 키가 있으면 파일이 없어도 동작하도록 load_dotenv는
      기존 값을 덮어쓰지 않는다(override=False, 기본값).
    """
    try:
        from dotenv import load_dotenv
    except ImportError:
        sys.exit("python-dotenv가 필요합니다. 먼저: pip install -r requirements.txt")

    candidates: list[str] = []
    if env_file:
        candidates.append(env_file)
    if os.path.exists(".env"):
        candidates.append(".env")
    # 표준 이름이 없을 때만 디렉터리의 다른 *.env 를 후보로(예: 'gemini api key.env')
    candidates += sorted(p for p in glob.glob("*.env") if p not in candidates)

    loaded_from = None
    for path in candidates:
        if os.path.exists(path):
            load_dotenv(path)
            loaded_from = path
            break
    if loaded_from:
        print(f"[env] 로드: {loaded_from}", file=sys.stderr)

    # SDK는 GEMINI_API_KEY 또는 GOOGLE_API_KEY 를 쓴다 — 둘 다 허용.
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        sys.exit(
            "API 키를 찾지 못했습니다. .env(또는 *.env)에 다음 형식으로 넣으세요:\n"
            "  GEMINI_API_KEY=발급받은_키\n"
            f"(탐색한 파일: {candidates or '없음'})"
        )
    return key


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    ap = argparse.ArgumentParser(description="Gemini API 호출 예제")
    ap.add_argument("prompt", nargs="?", default="안녕하세요. 한 문장으로 자기소개 해주세요.",
                    help="모델에 보낼 프롬프트")
    ap.add_argument("--env-file", default=os.environ.get("GEMINI_ENV_FILE"),
                    help="키가 든 env 파일 경로 (기본: .env → *.env 자동 탐색)")
    # 모델명은 시점에 따라 추가/변경될 수 있다. 사용 가능 목록은 AI Studio/문서에서 확인.
    ap.add_argument("--model", default="gemini-2.5-flash",
                    help="사용할 모델 (기본: gemini-2.5-flash)")
    args = ap.parse_args()

    api_key = load_api_key(args.env_file)

    try:
        from google import genai
    except ImportError:
        sys.exit("google-genai가 필요합니다. 먼저: pip install -r requirements.txt")

    client = genai.Client(api_key=api_key)

    try:
        response = client.models.generate_content(
            model=args.model,
            contents=args.prompt,
        )
    except Exception as e:  # 네트워크·인증·잘못된 모델명 등 호출 실패를 사용자 메시지로
        sys.exit(f"Gemini 호출 실패: {type(e).__name__}: {e}")

    print(response.text)


if __name__ == "__main__":
    main()
