"""생성된 .docx 를 PDF/HWP 로 변환 (Windows 전용, 오피스 COM 자동화).

설계 메모(왜):
- 이 프로젝트의 핵심은 회사 표준 템플릿 서식 보존이다. 따라서 PDF/HWP를 새로
  렌더링하지 않고(서식 깨짐), 설치된 오피스 앱으로 .docx 를 '변환'한다.
    · PDF: MS Word COM (FileFormat 17 = wdFormatPDF)
    · HWP: 한컴오피스 COM (HWPFrame.HwpObject) — docx 가져오기 후 HWP 저장
- COM은 절대경로를 요구하고, 실패 시 앱 프로세스가 남을 수 있어 finally에서 정리한다.
- HWP는 보안 모듈 미등록 시 자동화가 막히거나 보안창이 뜰 수 있다 → 호출부가 실패를
  잡아 .docx 생성 자체는 성공으로 두도록, 예외를 그대로 올린다.

standalone 사용:
    python convert.py --pdf --hwp "out/수행계획서_생성본.docx"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _abs(p) -> str:
    return str(Path(p).resolve())


def to_pdf(docx_path) -> Path:
    """MS Word COM으로 .docx → .pdf. 같은 이름의 .pdf 경로를 반환."""
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"원본 없음: {src}")
    out = src.with_suffix(".pdf")

    import pythoncom
    import win32com.client as win32
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(_abs(src))           # (FileName)
        doc.SaveAs(_abs(out), 17)                       # (FileName, FileFormat=wdFormatPDF)
        doc.Close(0)                                    # 0 = 저장 안 함(이미 SaveAs로 PDF 저장)
        return out
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def to_hwp(docx_path) -> Path:
    """한컴오피스 COM으로 .docx → .hwp. 같은 이름의 .hwp 경로를 반환.

    한컴 자동화는 파일 접근 시 '보안 승인 창'이 떠 헤드리스에서 멈춘다. 이를 막으려면
    먼저 `python register_hwp.py` 로 보안 모듈을 레지스트리에 1회 등록해야 한다.
    그러면 RegisterModule 호출로 모듈이 로드되어 보안창 없이 변환된다."""
    src = Path(docx_path)
    if not src.exists():
        raise FileNotFoundError(f"원본 없음: {src}")
    out = src.with_suffix(".hwp")

    import pythoncom
    from win32com.client import dynamic
    pythoncom.CoInitialize()
    hwp = None
    try:
        # 순수 late-binding으로 생성: pywin32 gencache(makepy)의 순환 임포트 버그
        # (Python 3.14 호환성)를 피해 COM 메서드 호출이 정상 마샬링되게 한다.
        hwp = dynamic.Dispatch("HWPFrame.HwpObject")
        # 사전 등록된 보안 모듈을 로드해 보안 승인 창을 억제한다.
        # (먼저 `python register_hwp.py` 로 레지스트리에 모듈을 등록해야 함)
        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        opened = hwp.Open(_abs(src), "", "")
        if not opened:
            opened = hwp.Open(_abs(src), "MS Word", "")   # 포맷 자동 감지 실패 시 명시
        if not opened:
            raise RuntimeError("한컴오피스가 .docx 를 열지 못했습니다(가져오기 필터 확인 필요)")
        hwp.SaveAs(_abs(out), "HWP", "")
        return out
    finally:
        if hwp is not None:
            try:
                hwp.Clear(1)   # 1 = 저장 없이 닫기
            except Exception:
                pass
            try:
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass

    ap = argparse.ArgumentParser(description="생성된 .docx → PDF/HWP 변환 (Windows/오피스 COM)")
    ap.add_argument("docx", help="변환할 .docx 경로")
    ap.add_argument("--pdf", action="store_true", help="PDF로 변환 (MS Word)")
    ap.add_argument("--hwp", action="store_true", help="HWP로 변환 (한컴오피스)")
    args = ap.parse_args()

    if not (args.pdf or args.hwp):
        ap.error("--pdf 또는 --hwp 중 하나 이상을 지정하세요.")

    failed = False
    if args.pdf:
        try:
            print(f"PDF 변환 완료: {to_pdf(args.docx)}")
        except Exception as e:
            failed = True
            print(f"PDF 변환 실패: {type(e).__name__}: {e}", file=sys.stderr)
    if args.hwp:
        try:
            print(f"HWP 변환 완료: {to_hwp(args.docx)}")
        except Exception as e:
            failed = True
            print(f"HWP 변환 실패: {type(e).__name__}: {e}", file=sys.stderr)
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
