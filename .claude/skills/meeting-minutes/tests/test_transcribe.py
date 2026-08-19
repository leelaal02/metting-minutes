import os

import pytest

from transcribe import run, _load_env_file
from adapters import get_adapter, available_sources


def test_available_sources_includes_text():
    assert "text" in available_sources()


def test_get_adapter_returns_callable():
    adapter = get_adapter("text")
    assert callable(adapter)


def test_get_adapter_unknown_source_raises_with_list():
    with pytest.raises(ValueError) as exc:
        get_adapter("nope")
    # 오류 메시지에 사용 가능한 소스를 안내
    assert "text" in str(exc.value)


def test_run_text_source_returns_normalized_text():
    result = run("text", "첫 줄   \n\n\n\n둘째 줄  ")
    assert result == "첫 줄\n\n둘째 줄"


def test_run_default_source_is_text():
    # source_name 없이 호출하면 text로 처리
    result = run("text", "회의 시작\n논의 내용")
    assert result == "회의 시작\n논의 내용"


def test_load_env_file_injects_key(tmp_path, monkeypatch):
    env = tmp_path / ".env"
    env.write_text('GROQ_API_KEY="secret123"\n# 주석\nFOO=bar\n', encoding="utf-8")
    monkeypatch.delenv("GROQ_API_KEY", raising=False)
    monkeypatch.delenv("FOO", raising=False)
    _load_env_file(str(env))
    assert os.environ["GROQ_API_KEY"] == "secret123"  # 따옴표 제거됨
    assert os.environ["FOO"] == "bar"


def test_load_env_file_does_not_override_existing(monkeypatch, tmp_path):
    env = tmp_path / ".env"
    env.write_text("GROQ_API_KEY=fromfile\n", encoding="utf-8")
    monkeypatch.setenv("GROQ_API_KEY", "fromenv")
    _load_env_file(str(env))
    # 이미 설정된 실제 환경변수가 우선(setdefault)
    assert os.environ["GROQ_API_KEY"] == "fromenv"


def test_load_env_file_missing_is_noop(tmp_path):
    # 파일이 없어도 예외 없이 통과
    _load_env_file(str(tmp_path / "nope.env"))
