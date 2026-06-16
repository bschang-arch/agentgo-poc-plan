"""생성된 .docx 를 PDF 로 변환 (Windows 전용, MS Word COM 자동화).

설계 메모(왜):
- 이 프로젝트의 핵심은 회사 표준 템플릿 서식 보존이다. 따라서 PDF를 새로 렌더링하지
  않고(서식 깨짐), 설치된 MS Word로 .docx 를 '변환'한다(FileFormat 17 = wdFormatPDF).
- early-binding(EnsureDispatch)을 쓰는 이유: 순수 late-binding은 Word Document 의
  SaveAs 같은 메서드를 타입라이브러리 없이 해석하지 못한다.
- Word는 문서 로딩 중 잠깐 바쁘면 호출을 거부(RPC_E_CALL_REJECTED)하므로 재시도한다.
- COM은 절대경로를 요구하고, 실패 시 Word 프로세스가 남을 수 있어 finally에서 정리한다.

HWP는 이 환경(한컴 2018)에서 COM Open이 외부 포맷(docx 등) 가져오기를 지원하지 않아
자동 변환하지 않는다. 생성된 .docx 또는 PDF를 한컴에서 열어 "다른 이름으로 저장 → HWP".

standalone 사용:
    python convert.py "out/수행계획서_생성본.docx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _abs(p) -> str:
    return str(Path(p).resolve())


# Word 자동화는 문서 로딩 중 잠깐 바쁘면 호출을 거부한다(일시 오류). 재시도로 넘긴다.
_RETRYABLE_HRESULTS = (-2147418111,  # RPC_E_CALL_REJECTED (피호출자가 호출을 거부)
                       -2147417846)  # RPC_E_SERVERCALL_RETRYLATER


def _com_retry(func, tries: int = 10, delay: float = 0.5):
    import time
    import pywintypes
    last = None
    for _ in range(tries):
        try:
            return func()
        except pywintypes.com_error as e:
            if e.args and e.args[0] in _RETRYABLE_HRESULTS:
                last = e
                time.sleep(delay)
                continue
            raise
    raise last


def to_pdf(docx_path) -> Path:
    """MS Word COM으로 .docx → .pdf. 같은 이름의 .pdf 경로를 반환."""
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"원본 없음: {src}")
    out = src.with_suffix(".pdf")

    import pythoncom
    from win32com.client import gencache
    pythoncom.CoInitialize()
    word = None
    try:
        # early-binding(EnsureDispatch): 타입라이브러리로 Document.SaveAs 등을 제대로 해석한다.
        word = gencache.EnsureDispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = _com_retry(lambda: word.Documents.Open(_abs(src)))   # (FileName)
        _com_retry(lambda: doc.SaveAs(_abs(out), 17))              # 17 = wdFormatPDF
        doc.Close(0)                                               # 0 = 저장 안 함(이미 PDF 저장됨)
        return out
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    ap = argparse.ArgumentParser(description="생성된 .docx → PDF 변환 (Windows/MS Word COM)")
    ap.add_argument("docx", help="변환할 .docx 경로")
    args = ap.parse_args()

    try:
        print(f"PDF 변환 완료: {to_pdf(args.docx)}")
    except Exception as e:
        sys.exit(f"PDF 변환 실패: {type(e).__name__}: {e}")


if __name__ == "__main__":
    main()
