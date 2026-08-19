"""입력 어댑터 레지스트리.

새 입력 방식 추가: 어댑터 모듈 작성 후 REGISTRY에 한 줄 등록.
삭제: REGISTRY에서 한 줄 제거 + 모듈 파일 삭제. 하위 단계는 무영향.

STT 어댑터 모듈은 top-level에서 무거운 라이브러리를 import하지 않으므로
(faster_whisper/groq는 각 transcribe 함수 내부에서 lazy import),
이 패키지를 import해도 STT 미설치 상태에서 안전하다.
"""
from adapters import text, whisper_local, groq_cloud

REGISTRY = {
    "text": text.transcribe,
    "whisper": whisper_local.transcribe,
    "groq": groq_cloud.transcribe,
}


def available_sources() -> list:
    return list(REGISTRY)


def get_adapter(name: str):
    try:
        return REGISTRY[name]
    except KeyError:
        raise ValueError(
            f"알 수 없는 소스 '{name}'. 사용 가능: {', '.join(available_sources())}"
        )
