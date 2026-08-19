"""[1] 입력 계층 CLI 디스패처.

사용자가 직접 실행하지 않는다. SKILL.md 지시에 따라 Claude가 [1] 단계에서
실행하는 헬퍼. 정규화된 회의 텍스트를 표준 출력으로 내보낸다.

사용:
    python scripts/transcribe.py --source text    회의.txt
    python scripts/transcribe.py --source whisper 회의.wav
    python scripts/transcribe.py --source groq    회의.m4a
"""
import argparse
import os
import sys
from pathlib import Path

from adapters import get_adapter, available_sources


def _load_env_file(path: str = ".env") -> None:
    """작업 디렉토리의 .env(단순 KEY=VALUE 형식)를 읽어 os.environ에 주입.

    - 이미 설정된 실제 환경변수는 덮지 않는다(setdefault).
    - 빈 줄·`#` 주석·`=` 없는 줄은 무시. 값의 감싼 따옴표는 제거.
    - 파일이 없으면 아무것도 하지 않는다.
    비밀 키(GROQ_API_KEY 등)를 커밋 없이 로컬 파일로 관리하기 위함. `.env`는
    반드시 .gitignore에 두어야 한다.
    """
    p = Path(path)
    if not p.exists():
        return
    # utf-8-sig: Windows 메모장으로 저장한 .env의 BOM을 제거해 첫 줄 키가
    # '﻿GROQ_API_KEY'로 깨지지 않게 한다.
    for raw in p.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


def run(source_name: str, source: str, **opts) -> str:
    adapter = get_adapter(source_name)
    return adapter(source, **opts)


def main() -> None:
    _load_env_file()  # cwd의 .env에서 GROQ_API_KEY 등 로드
    parser = argparse.ArgumentParser(description="회의 입력 → 정규화된 회의 텍스트")
    parser.add_argument(
        "--source", default="text", choices=available_sources(),
        help="입력 어댑터 선택 (기본: text)",
    )
    parser.add_argument("input", help="텍스트 문자열, .txt 경로, 또는 오디오 파일 경로")
    args = parser.parse_args()
    # 플랫폼 콘솔 인코딩(예: Windows cp949)에 무관하게 UTF-8로 출력.
    # 리다이렉트(> output/meeting.txt) 결과가 하위 단계에서 UTF-8로 읽히도록 보장.
    sys.stdout.buffer.write(run(args.source, args.input).encode("utf-8"))


if __name__ == "__main__":
    main()
