"""한컴오피스 HWP 자동화 보안 모듈 등록 (1회 설정, Windows 전용).

왜 필요한가: 한컴 자동화는 파일 접근 시 '보안 승인 창'을 띄워 헤드리스 변환이 멈춘다.
보안 모듈 DLL(FilePathCheckerModule.dll)을 레지스트리에 등록하면 이 창이 억제된다.
DLL은 pyhwpx 패키지에 번들된 것을 사용한다.

주의(보안 트레이드오프): 이 등록은 자동화 클라이언트의 파일 경로 접근 제한을 우회하도록
허용한다. HWP 자동화의 표준 방식이지만, 신뢰할 수 있는 환경에서만 사용한다.

사용:
    python register_hwp.py            # 보안 모듈 등록
    python register_hwp.py --check    # 현재 등록 상태만 확인
"""

from __future__ import annotations

import argparse
import os
import sys
import winreg

REG_PATH = r"Software\HNC\HwpAutomation\Modules"
MODULE_NAME = "FilePathCheckerModule"   # to_hwp 의 RegisterModule 두 번째 인자와 일치해야 함


def _dll_path() -> str:
    try:
        import pyhwpx
    except ImportError:
        sys.exit("pyhwpx 가 필요합니다(보안 모듈 DLL 제공): pip install pyhwpx")
    dll = os.path.join(os.path.dirname(pyhwpx.__file__), "FilePathCheckerModule.dll")
    if not os.path.exists(dll):
        sys.exit(f"보안 모듈 DLL을 찾지 못함: {dll}")
    return dll


def _pe_arch(path: str) -> str:
    with open(path, "rb") as f:
        data = f.read(0x200)
    pe = int.from_bytes(data[0x3C:0x40], "little")
    machine = int.from_bytes(data[pe + 4:pe + 6], "little")
    return {0x14C: "x86(32bit)", 0x8664: "x64(64bit)"}.get(machine, hex(machine))


def check() -> str | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH) as k:
            val, _ = winreg.QueryValueEx(k, MODULE_NAME)
            return val
    except FileNotFoundError:
        return None


def register() -> str:
    dll = _dll_path()
    # HWP는 32bit이므로 DLL도 32bit여야 로드된다(참고용 안내).
    print(f"DLL: {dll}\n아키텍처: {_pe_arch(dll)} (한컴 2018은 32bit)")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH) as k:
        winreg.SetValueEx(k, MODULE_NAME, 0, winreg.REG_SZ, dll)
    return dll


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass
    ap = argparse.ArgumentParser(description="HWP 자동화 보안 모듈 등록")
    ap.add_argument("--check", action="store_true", help="등록 상태만 확인")
    args = ap.parse_args()

    if args.check:
        cur = check()
        print(f"등록됨: {cur}" if cur else "등록 안 됨")
        return
    register()
    print(f"등록 완료: HKCU\\{REG_PATH} : {MODULE_NAME}")
    print("확인:", check())


if __name__ == "__main__":
    main()
