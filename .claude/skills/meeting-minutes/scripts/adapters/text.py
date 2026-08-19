"""text 어댑터: 붙여넣은 문자열 또는 .txt 경로 → 정규화된 회의 텍스트."""
from normalize_input import load_meeting_text


def transcribe(source: str, **opts) -> str:
    """기존 입력 정규화에 위임. opts는 무시(인터페이스 통일용)."""
    return load_meeting_text(source)
