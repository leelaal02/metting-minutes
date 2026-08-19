import sys
from pathlib import Path

import pytest

from adapters.text import transcribe as text_transcribe
from adapters import whisper_local, groq_cloud

SAMPLE_TXT = Path(__file__).resolve().parent.parent / "examples" / "sample_meeting.txt"


def test_text_adapter_from_string():
    assert text_transcribe("첫 줄   \n\n\n\n둘째 줄  ") == "첫 줄\n\n둘째 줄"


def test_text_adapter_from_file():
    result = text_transcribe(str(SAMPLE_TXT))
    assert "STT 연동" in result
    assert "김수민" in result


class _FakeSegment:
    def __init__(self, text):
        self.text = text


class _FakeModel:
    def transcribe(self, path, **kwargs):
        # faster-whisper API: (segments, info) 반환. segments는 이터러블.
        return [_FakeSegment("첫 줄  "), _FakeSegment(" 둘째 줄")], None


def test_whisper_adapter_joins_and_normalizes(monkeypatch, tmp_path):
    audio = tmp_path / "dummy.wav"
    audio.write_bytes(b"fake")
    monkeypatch.setattr(whisper_local, "_load_model", lambda **kw: _FakeModel())
    result = whisper_local.transcribe(str(audio))
    # 세그먼트 텍스트를 이어 붙이고 정규화(줄 끝 공백 제거)
    assert result == "첫 줄\n둘째 줄"


def test_whisper_adapter_missing_library(monkeypatch):
    # faster_whisper import 실패를 강제 → 친절한 설치 안내
    monkeypatch.setitem(sys.modules, "faster_whisper", None)
    with pytest.raises(ImportError) as exc:
        whisper_local._load_model()
    assert "requirements-stt.txt" in str(exc.value)


class _FakeTranscription:
    text = "첫 줄  \n\n\n둘째 줄  "


class _FakeAudio:
    class transcriptions:
        @staticmethod
        def create(**kwargs):
            return _FakeTranscription()


class _FakeGroqClient:
    audio = _FakeAudio()


def test_groq_adapter_normalizes(monkeypatch, tmp_path):
    audio = tmp_path / "a.m4a"
    audio.write_bytes(b"fake")
    monkeypatch.setattr(groq_cloud, "_client", lambda: _FakeGroqClient())
    result = groq_cloud.transcribe(str(audio))
    assert result == "첫 줄\n\n둘째 줄"


def test_groq_adapter_compresses_when_oversized(monkeypatch, tmp_path):
    # 한도 초과 파일이면 _compress_for_upload로 재인코딩 후 전송한다.
    audio = tmp_path / "big.m4a"
    audio.write_bytes(b"fake-large-audio")
    monkeypatch.setattr(groq_cloud, "GROQ_MAX_BYTES", 2)  # 무조건 초과로 취급
    monkeypatch.setattr(groq_cloud, "_client", lambda: _FakeGroqClient())

    seen = {}

    def fake_compress(path):
        seen["path"] = path
        out = tmp_path / "small.mp3"
        out.write_bytes(b"tiny")
        return str(out)

    monkeypatch.setattr(groq_cloud, "_compress_for_upload", fake_compress)
    result = groq_cloud.transcribe(str(audio))
    assert str(seen["path"]) == str(audio)  # 원본 경로로 재인코딩 호출
    assert not (tmp_path / "small.mp3").exists()  # 전송 후 임시 파일 정리
    assert result == "첫 줄\n\n둘째 줄"


def test_groq_adapter_skips_compression_when_small(monkeypatch, tmp_path):
    # 한도 이하면 재인코딩하지 않고 원본을 그대로 전송한다.
    audio = tmp_path / "small.m4a"
    audio.write_bytes(b"fake")
    monkeypatch.setattr(groq_cloud, "_client", lambda: _FakeGroqClient())

    def boom(path):
        raise AssertionError("작은 파일은 재인코딩하면 안 됨")

    monkeypatch.setattr(groq_cloud, "_compress_for_upload", boom)
    assert groq_cloud.transcribe(str(audio)) == "첫 줄\n\n둘째 줄"


def test_groq_adapter_missing_api_key(monkeypatch):
    monkeypatch.delenv("GROQ_API_KEY", raising=False)
    # groq 라이브러리 존재 여부와 무관하게 키 미설정을 먼저 안내
    with pytest.raises(RuntimeError) as exc:
        groq_cloud._client()
    assert "GROQ_API_KEY" in str(exc.value)


def test_groq_adapter_missing_library(monkeypatch):
    monkeypatch.setenv("GROQ_API_KEY", "dummy")
    monkeypatch.setitem(sys.modules, "groq", None)
    with pytest.raises(ImportError) as exc:
        groq_cloud._client()
    assert "requirements-stt.txt" in str(exc.value)
