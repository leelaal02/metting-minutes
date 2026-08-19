"""클라우드 STT 어댑터: Groq Whisper API(whisper-large-v3)로 오디오 → 텍스트.

groq 라이브러리는 _client 안에서 lazy import한다.

Groq 업로드 한도(약 25MB)를 초과하는 오디오는 imageio_ffmpeg가 번들한 ffmpeg로
16kHz 모노 저비트레이트 mp3로 자동 재인코딩해 전송한다(원본은 건드리지 않음).
"""
import os
import subprocess
import tempfile

from normalize_input import _normalize, resolve_input_path

# Groq 업로드 한도는 약 25MB. 안전 여유를 두고 이보다 크면 자동 재인코딩한다.
GROQ_MAX_BYTES = 24 * 1024 * 1024


def _ffmpeg_exe() -> str:
    """imageio_ffmpeg가 번들한 ffmpeg 실행 파일 경로. 없으면 설치 안내."""
    try:
        import imageio_ffmpeg
    except ImportError as e:
        raise ImportError(
            "오디오가 Groq 업로드 한도를 초과해 재인코딩이 필요하지만 "
            "imageio_ffmpeg가 설치되지 않았습니다. "
            "pip install -r requirements-stt.txt 를 실행하세요."
        ) from e
    return imageio_ffmpeg.get_ffmpeg_exe()


def _compress_for_upload(audio_path: str) -> str:
    """오디오를 16kHz 모노 저비트레이트 mp3로 재인코딩해 임시 파일 경로를 반환.

    STT는 16kHz 모노면 충분하므로 용량을 크게 줄여 업로드 한도 아래로 맞춘다.
    반환된 임시 파일은 호출자가 사용 후 삭제한다.
    """
    ffmpeg = _ffmpeg_exe()
    fd, out_path = tempfile.mkstemp(suffix=".mp3", prefix="groq_stt_")
    os.close(fd)
    cmd = [
        ffmpeg, "-y", "-i", str(audio_path),
        "-ac", "1", "-ar", "16000", "-b:a", "48k",
        out_path, "-loglevel", "error",
    ]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        if os.path.exists(out_path):
            os.remove(out_path)
        raise RuntimeError(
            f"오디오 재인코딩(ffmpeg) 실패: {proc.stderr.strip() or proc.returncode}"
        )
    return out_path


def _client():
    """Groq 클라이언트 생성. 키 미설정/라이브러리 미설치를 명확히 안내."""
    api_key = os.environ.get("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError(
            "GROQ_API_KEY 환경변수가 설정되지 않았습니다. "
            "Groq 클라우드 STT를 쓰려면 API 키를 설정하세요."
        )
    try:
        from groq import Groq
    except ImportError as e:
        raise ImportError(
            "groq가 설치되지 않았습니다. "
            "pip install -r requirements-stt.txt 를 실행하세요."
        ) from e
    return Groq(api_key=api_key)


def transcribe(source: str, **opts) -> str:
    """오디오 파일 경로(source)를 Groq로 전사해 정규화된 텍스트로 반환.

    opts: model(기본 "whisper-large-v3"), language(기본 "ko").
    """
    client = _client()
    audio_path = resolve_input_path(source, must_exist=True)

    # 한도 초과 시 재인코딩한 임시 파일로 대체하고, 전송 후 정리한다.
    tmp_path = None
    if os.path.getsize(audio_path) > GROQ_MAX_BYTES:
        tmp_path = _compress_for_upload(audio_path)
        audio_path = tmp_path
    try:
        with open(audio_path, "rb") as audio:
            result = client.audio.transcriptions.create(
                file=audio,
                model=opts.get("model", "whisper-large-v3"),
                language=opts.get("language", "ko"),
            )
    finally:
        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)
    normalized = _normalize(result.text)
    if not normalized:
        raise ValueError("전사 결과가 비어 있습니다: 오디오에서 텍스트를 얻지 못했습니다.")
    return normalized
