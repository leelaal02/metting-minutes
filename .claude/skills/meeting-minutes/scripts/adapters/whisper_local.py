"""로컬 STT 어댑터: faster-whisper로 오디오 → 정규화된 회의 텍스트.

무거운 라이브러리는 _load_model 안에서 lazy import한다.
"""
from normalize_input import _normalize, resolve_input_path


def _load_model(model_size: str = "medium", cpu_threads: int = 4):
    """faster-whisper 모델 로딩(lazy import). 미설치 시 설치 안내."""
    try:
        from faster_whisper import WhisperModel
    except ImportError as e:
        raise ImportError(
            "faster-whisper가 설치되지 않았습니다. "
            "pip install -r requirements-stt.txt 를 실행하세요."
        ) from e
    return WhisperModel(model_size, cpu_threads=cpu_threads)


def transcribe(source: str, **opts) -> str:
    """오디오 파일 경로(source)를 전사해 정규화된 텍스트로 반환.

    opts: model_size(기본 "medium"), cpu_threads(기본 4), language(기본 "ko").
    """
    model = _load_model(
        model_size=opts.get("model_size", "medium"),
        cpu_threads=opts.get("cpu_threads", 4),
    )
    audio_path = resolve_input_path(source, must_exist=True)
    segments, _info = model.transcribe(str(audio_path), language=opts.get("language", "ko"))
    text = "\n".join(seg.text.strip() for seg in segments)
    normalized = _normalize(text)
    if not normalized:
        raise ValueError("전사 결과가 비어 있습니다: 오디오에서 텍스트를 얻지 못했습니다.")
    return normalized
