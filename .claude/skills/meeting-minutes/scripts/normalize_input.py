"""입력 계층: 회의 원문 텍스트를 정규화 + 입력 파일 위치 탐색.

`resolve_input_path`는 모든 입력 어댑터(text/whisper/groq 및 앞으로 추가될 것)가
공통으로 호출하는 파일 탐색 헬퍼다. 작업 폴더에 없더라도 파일 이름만으로
공통 위치를 찾아준다. 어댑터는 이 헬퍼만 쓰면 되므로 탐색 로직이 한 곳에 모인다.
"""
import os
import sys
from pathlib import Path

# 파일 이름만 주어졌을 때 재귀 탐색을 건너뛸 잡음 디렉토리(속도·안전).
_SKIP_DIRS = {
    "node_modules", "__pycache__", ".venv", "venv", "env",
    ".pytest_cache", "AppData", "$Recycle.Bin",
    "System Volume Information", "Windows", "Program Files",
    "Program Files (x86)",
}
# 재귀 탐색 최대 깊이(루트 기준). 홈 등 큰 트리에서 과도한 탐색 방지.
_MAX_DEPTH = 6


def _search_roots() -> list:
    """파일 이름 탐색에 쓸 루트 목록(우선순위 순, 존재하는 것만, 중복 제거).

    MEETING_INPUT_DIRS(os.pathsep 구분)로 앞쪽에 루트를 추가할 수 있다.
    기본: 현재 작업 폴더 → 바탕화면 → 다운로드 → 문서 → 홈.
    """
    roots = []
    env = os.environ.get("MEETING_INPUT_DIRS")
    if env:
        roots += [Path(p) for p in env.split(os.pathsep) if p.strip()]
    home = Path.home()
    roots += [Path.cwd(), home / "Desktop", home / "Downloads",
              home / "Documents", home]
    seen, uniq = set(), []
    for r in roots:
        try:
            key = r.resolve()
        except OSError:
            key = r
        if key not in seen and r.is_dir():
            seen.add(key)
            uniq.append(r)
    return uniq


def _find_under(root: Path, name: str) -> Path:
    """root 아래를 재귀 탐색해 파일명이 일치하는 첫 파일 경로를 반환(없으면 None)."""
    root = Path(root)
    root_depth = len(root.parts)
    try:
        for dirpath, dirnames, filenames in os.walk(root):
            depth = len(Path(dirpath).parts) - root_depth
            if depth >= _MAX_DEPTH:
                dirnames[:] = []
            else:
                # 잡음·숨김 디렉토리 가지치기(탐색 속도 확보)
                dirnames[:] = [d for d in dirnames
                               if d not in _SKIP_DIRS and not d.startswith(".")]
            if name in filenames:
                return Path(dirpath) / name
    except (OSError, PermissionError):
        return None
    return None


def resolve_input_path(source: str, *, must_exist: bool = False):
    """입력 경로 문자열을 실제 파일 경로로 해석.

    1. source가 그대로 존재하는 파일이면 그 경로를 반환(cwd 기준 상대경로 포함).
    2. 아니고 절대경로면 탐색하지 않음(경로가 명시된 것이므로).
    3. 파일 이름/상대경로면 공통 루트들을 재귀 탐색해 첫 일치를 반환.
    4. 못 찾으면 must_exist=True면 FileNotFoundError, 아니면 None.

    작업 폴더 밖에서 찾았을 때는 어느 경로를 썼는지 stderr로 알린다(투명성).
    """
    candidate = Path(source)
    if candidate.exists() and candidate.is_file():
        return candidate
    if candidate.is_absolute():
        if must_exist:
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {source}")
        return None
    name = candidate.name
    for root in _search_roots():
        match = _find_under(root, name)
        if match is not None:
            print(f"[transcribe] 입력 파일을 찾았습니다: {match}", file=sys.stderr)
            return match
    if must_exist:
        locations = ", ".join(str(r) for r in _search_roots())
        raise FileNotFoundError(
            f"'{source}' 파일을 찾지 못했습니다. 검색한 위치: {locations}. "
            "정확한 경로를 지정하거나 MEETING_INPUT_DIRS 환경변수로 폴더를 추가하세요."
        )
    return None


def _looks_like_path(source: str) -> bool:
    """source를 파일 경로 후보로 볼지 판단.

    파일 경로에는 개행이 없고 길이도 OS 한도(약 260자) 안이다. 붙여넣은 회의
    원문(여러 줄이거나 매우 긴 문자열)까지 경로로 해석하면, 매번 공통 폴더를
    재귀 탐색(os.walk)하는 낭비가 생기고 원문 끝부분이 우연히 실제 파일명과
    겹치면 엉뚱한 파일을 읽는다. 이 특징으로 경로 후보만 걸러낸다.
    """
    s = source.strip()
    return bool(s) and "\n" not in s and "\r" not in s and len(s) <= 260


# 입력 파일로 볼 확장자. 이 확장자로 끝나는 한 줄 입력은 "파일을 주려 한 것"이므로
# 못 찾으면 조용히 원문으로 삼지 않고 실패해야 한다 — 오타 하나로 파일명 한 줄이
# 회의 원문이 되어 회의록 전체가 그 한 줄로 만들어지는 사고를 막는다.
# 붙여넣은 회의 원문이 우연히 이런 확장자로 끝날 일은 없으므로 오탐 위험은 없다.
_INPUT_SUFFIXES = {
    ".txt", ".text", ".md", ".log", ".csv", ".json", ".vtt", ".srt",
    ".rtf", ".docx", ".doc", ".hwp", ".hwpx", ".pdf",
    ".wav", ".mp3", ".m4a", ".mp4", ".flac", ".ogg", ".aac", ".wma", ".webm",
}


def _looks_like_file(source: str) -> bool:
    """경로 후보이면서 알려진 입력 확장자로 끝나면 True(= 반드시 존재해야 하는 입력)."""
    return (
        _looks_like_path(source)
        and Path(source.strip()).suffix.lower() in _INPUT_SUFFIXES
    )


def load_meeting_text(source: str) -> str:
    """텍스트 문자열 또는 .txt 파일 경로를 정규화된 회의 원문으로 변환.

    - source가 경로 후보이고 (작업 폴더 밖이라도) 파일로 해석되면 파일 내용을 읽음.
    - 파일 확장자(`.txt` 등)로 끝나면 **못 찾을 때 실패**한다 — 파일을 주려 한 입력이
      조용히 회의 원문으로 둔갑하지 않도록.
    - 확장자가 없는 문자열은 붙여넣은 원문으로 간주한다.
    """
    if _looks_like_file(source):
        resolved = resolve_input_path(source, must_exist=True)
    else:
        resolved = resolve_input_path(source, must_exist=False) if _looks_like_path(source) else None
    if resolved is not None:
        text = resolved.read_text(encoding="utf-8")
    else:
        text = source
    normalized = _normalize(text)
    if not normalized:
        raise ValueError("빈 입력입니다: 회의 원문 텍스트가 없습니다.")
    return normalized


def _normalize(text: str) -> str:
    """줄 끝 공백 제거, 연속 빈 줄을 하나로 축약, 앞뒤 공백 제거."""
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").split("\n")]
    result = []
    prev_blank = False
    for line in lines:
        blank = (line == "")
        if blank and prev_blank:
            continue
        result.append(line)
        prev_blank = blank
    return "\n".join(result).strip()


