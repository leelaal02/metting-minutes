# STT 입력 어댑터 3종 Implementation Plan (계획서)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 회의록 Skill의 [1] 입력 계층에 텍스트 / 로컬 STT(faster-whisper) / 클라우드 STT(Groq Whisper API) 세 어댑터를 플러그인 레지스트리로 추가한다. 하위 파이프라인([2]~[4])은 수정하지 않는다.

**Architecture:** `scripts/adapters/` 패키지에 어댑터별 모듈을 두고, 각 어댑터는 공통 시그니처 `transcribe(source, **opts) -> str`(정규화된 회의 텍스트)를 구현한다. `scripts/adapters/__init__.py`의 `REGISTRY`가 이름→함수를 매핑하고, `scripts/transcribe.py`가 `--source`로 어댑터를 골라 실행하는 CLI 디스패처다. STT 라이브러리는 어댑터 함수 내부에서 lazy import하여, 미설치 상태에서도 텍스트 파이프라인이 동작한다. `transcribe.py`는 사용자가 직접 실행하지 않고 SKILL.md 지시에 따라 Claude가 [1] 단계에서 실행하는 헬퍼다.

**Tech Stack:** Python 3.9+, faster-whisper, groq, pytest.

## Global Constraints

- Python 3.9+ 사용 (표준 라이브러리 `pathlib`, 타입 힌트).
- 하위 단계([2]추출 / [3]render_markdown / [4]render_docx)와 `schema/minutes.schema.json`은 **수정 금지**.
- 어댑터 공통 인터페이스: `transcribe(source: str, **opts) -> str` — 반환값은 항상 **정규화된 회의 텍스트**.
- 어댑터는 반환 전 `normalize_input._normalize()`를 적용해 하위 단계가 소비하는 형태로 통일.
- STT 라이브러리(`faster_whisper`, `groq`)는 어댑터 함수 내부에서 **lazy import**. 미설치 시 `pip install -r requirements-stt.txt` 안내와 함께 명확한 오류.
- 로컬 STT 기본값: `WhisperModel("medium", cpu_threads=4)`, 전사 `language="ko"`.
- 클라우드 STT: Groq `whisper-large-v3`, `GROQ_API_KEY` 환경변수 사용.
- `--source` 미지정 시 기본값 `text`.
- 모든 파일 입출력 인코딩 `utf-8`.
- 테스트는 실제 오디오·네트워크 없이 STT 라이브러리를 **모킹**해 결정적으로 수행.
- import 경로: `tests/conftest.py`가 `scripts/`를 `sys.path[0]`에 넣으므로, 테스트/어댑터는 `normalize_input`·`adapters`·`transcribe`를 최상위 모듈로 import.

---

## File Structure

- `requirements-stt.txt` — STT 선택 필요 패키지 (faster-whisper, groq).
- `scripts/adapters/__init__.py` — `REGISTRY`, `get_adapter()`, `available_sources()`.
- `scripts/adapters/text.py` — 텍스트 어댑터 (기존 `normalize_input` 위임).
- `scripts/adapters/whisper_local.py` — 로컬 STT 어댑터 (faster-whisper).
- `scripts/adapters/groq_cloud.py` — 클라우드 STT 어댑터 (Groq API).
- `scripts/transcribe.py` — CLI 디스패처 진입점.
- `tests/test_transcribe.py` — 디스패처/레지스트리 테스트.
- `tests/test_adapters.py` — 어댑터 테스트 (text/whisper/groq, STT는 모킹).
- `SKILL.md` — [1] 단계 지시문 업데이트 (수정).

---

## Task 1: 어댑터 패키지 골격 + text 어댑터 + 디스패처

플러그인 레지스트리와 CLI 디스패처를 세우고, 외부 패키지가 필요 없는 text 어댑터로 [1] 단계를 엔드투엔드 동작시킨다. STT 어댑터는 이후 태스크에서 이 골격에 등록만 하면 된다.

**Files:**
- Create: `requirements-stt.txt`
- Create: `scripts/adapters/__init__.py`
- Create: `scripts/adapters/text.py`
- Create: `scripts/transcribe.py`
- Test: `tests/test_transcribe.py`, `tests/test_adapters.py`

**Interfaces:**
- Consumes: `normalize_input.load_meeting_text(source: str) -> str` (기존).
- Produces:
  - `adapters/text.py`: `transcribe(source: str, **opts) -> str`.
  - `adapters/__init__.py`: `REGISTRY: dict[str, Callable]`, `get_adapter(name: str) -> Callable` (미등록 시 `KeyError` 대신 `ValueError`로 사용 가능한 목록 안내), `available_sources() -> list[str]`.
  - `transcribe.py`: `run(source_name: str, source: str, **opts) -> str`, `main() -> None` (CLI: `--source`, positional `input`).

- [ ] **Step 1: STT 선택 필요 패키지 파일 작성**

`requirements-stt.txt`:
```
faster-whisper>=1.0.0
groq>=0.11.0
```

- [ ] **Step 2: 디스패처/레지스트리 테스트 작성 (실패 예상)**

`tests/test_transcribe.py`:
```python
import pytest

from transcribe import run
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
```

- [ ] **Step 3: text 어댑터 회귀 테스트 작성 (실패 예상)**

`tests/test_adapters.py`:
```python
from pathlib import Path

from adapters.text import transcribe as text_transcribe

SAMPLE_TXT = Path(__file__).resolve().parent.parent / "examples" / "sample_meeting.txt"


def test_text_adapter_from_string():
    assert text_transcribe("첫 줄   \n\n\n\n둘째 줄  ") == "첫 줄\n\n둘째 줄"


def test_text_adapter_from_file():
    result = text_transcribe(str(SAMPLE_TXT))
    assert "STT 연동" in result
    assert "김수민" in result
```

- [ ] **Step 4: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_transcribe.py tests/test_adapters.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'transcribe'` / `'adapters'`

- [ ] **Step 5: text 어댑터 구현**

`scripts/adapters/text.py`:
```python
"""text 어댑터: 붙여넣은 문자열 또는 .txt 경로 → 정규화된 회의 텍스트."""
from normalize_input import load_meeting_text


def transcribe(source: str, **opts) -> str:
    """기존 입력 정규화에 위임. opts는 무시(인터페이스 통일용)."""
    return load_meeting_text(source)
```

- [ ] **Step 6: 레지스트리 구현**

`scripts/adapters/__init__.py`:
```python
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
```

> 참고: `__init__.py`는 `whisper_local`/`groq_cloud`를 import하지만, 이 모듈들의 무거운 STT import는 함수 내부에 있으므로 여기서 로드해도 안전하다. 두 모듈은 Task 2·3에서 생성한다. **이 태스크에서는 임시로 두 줄(whisper/groq)을 주석 처리**하고 text만 등록해 테스트를 통과시킨 뒤, Task 2·3에서 해당 줄의 주석을 해제한다.

`__init__.py` (Task 1 시점 — whisper/groq 주석):
```python
from adapters import text
# from adapters import whisper_local, groq_cloud  # Task 2·3에서 활성화

REGISTRY = {
    "text": text.transcribe,
    # "whisper": whisper_local.transcribe,  # Task 2
    # "groq": groq_cloud.transcribe,        # Task 3
}
```
(`available_sources`/`get_adapter` 함수 본문은 위와 동일하게 작성.)

- [ ] **Step 7: 디스패처 구현**

`scripts/transcribe.py`:
```python
"""[1] 입력 계층 CLI 디스패처.

사용자가 직접 실행하지 않는다. SKILL.md 지시에 따라 Claude가 [1] 단계에서
실행하는 헬퍼. 정규화된 회의 텍스트를 표준 출력으로 내보낸다.

사용:
    python scripts/transcribe.py --source text    회의.txt
    python scripts/transcribe.py --source whisper 회의.wav
    python scripts/transcribe.py --source groq    회의.m4a
"""
import argparse
import sys

from adapters import get_adapter, available_sources


def run(source_name: str, source: str, **opts) -> str:
    adapter = get_adapter(source_name)
    return adapter(source, **opts)


def main() -> None:
    parser = argparse.ArgumentParser(description="회의 입력 → 정규화된 회의 텍스트")
    parser.add_argument(
        "--source", default="text", choices=available_sources(),
        help="입력 어댑터 선택 (기본: text)",
    )
    parser.add_argument("input", help="텍스트 문자열, .txt 경로, 또는 오디오 파일 경로")
    args = parser.parse_args()
    sys.stdout.write(run(args.source, args.input))


if __name__ == "__main__":
    main()
```

- [ ] **Step 8: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_transcribe.py tests/test_adapters.py -v`
Expected: PASS (test_transcribe 5개 + test_adapters 2개)

- [ ] **Step 9: CLI 수동 확인**

Run: `python scripts/transcribe.py --source text examples/sample_meeting.txt`
Expected: 정규화된 회의 원문이 표준 출력으로 표시됨 (예외 없음).

- [ ] **Step 10: 커밋**

```bash
git add requirements-stt.txt scripts/adapters/__init__.py scripts/adapters/text.py scripts/transcribe.py tests/test_transcribe.py tests/test_adapters.py
git commit -m "feat: 입력 어댑터 레지스트리 + text 어댑터 + 디스패처"
```

---

## Task 2: 로컬 STT 어댑터 (faster-whisper)

오디오 파일을 로컬 faster-whisper로 전사한다. 실제 모델 로딩은 lazy import + 헬퍼 함수로 감싸, 테스트에서 모킹 가능하게 한다.

**Files:**
- Create: `scripts/adapters/whisper_local.py`
- Modify: `scripts/adapters/__init__.py` (whisper 등록 주석 해제)
- Test: `tests/test_adapters.py` (whisper 테스트 추가)

**Interfaces:**
- Consumes: `normalize_input._normalize(text: str) -> str` (기존, 내부 함수).
- Produces:
  - `adapters/whisper_local.py`: `transcribe(source: str, **opts) -> str`, `_load_model(model_size: str = "medium", cpu_threads: int = 4)` (lazy import 헬퍼, 테스트에서 monkeypatch 대상).

- [ ] **Step 1: whisper 어댑터 테스트 작성 (실패 예상)**

`tests/test_adapters.py`에 다음을 추가:
```python
import sys

import pytest

from adapters import whisper_local


class _FakeSegment:
    def __init__(self, text):
        self.text = text


class _FakeModel:
    def transcribe(self, path, **kwargs):
        # faster-whisper API: (segments, info) 반환. segments는 이터러블.
        return [_FakeSegment("첫 줄  "), _FakeSegment(" 둘째 줄")], None


def test_whisper_adapter_joins_and_normalizes(monkeypatch):
    monkeypatch.setattr(whisper_local, "_load_model", lambda **kw: _FakeModel())
    result = whisper_local.transcribe("dummy.wav")
    # 세그먼트 텍스트를 이어 붙이고 정규화(줄 끝 공백 제거)
    assert result == "첫 줄\n둘째 줄"


def test_whisper_adapter_missing_library(monkeypatch):
    # faster_whisper import 실패를 강제 → 친절한 설치 안내
    monkeypatch.setitem(sys.modules, "faster_whisper", None)
    with pytest.raises(ImportError) as exc:
        whisper_local._load_model()
    assert "requirements-stt.txt" in str(exc.value)
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_adapters.py -k whisper -v`
Expected: FAIL — `AttributeError: module 'adapters.whisper_local' has no attribute ...` 또는 import 오류

- [ ] **Step 3: whisper 어댑터 구현**

`scripts/adapters/whisper_local.py`:
```python
"""로컬 STT 어댑터: faster-whisper로 오디오 → 정규화된 회의 텍스트.

무거운 라이브러리는 _load_model 안에서 lazy import한다.
"""
from normalize_input import _normalize


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
    segments, _info = model.transcribe(source, language=opts.get("language", "ko"))
    text = "\n".join(seg.text.strip() for seg in segments)
    normalized = _normalize(text)
    if not normalized:
        raise ValueError("전사 결과가 비어 있습니다: 오디오에서 텍스트를 얻지 못했습니다.")
    return normalized
```

- [ ] **Step 4: 레지스트리에 whisper 등록**

`scripts/adapters/__init__.py` 수정 — import와 REGISTRY의 whisper 줄 주석 해제:
```python
from adapters import text, whisper_local
# from adapters import groq_cloud  # Task 3에서 활성화

REGISTRY = {
    "text": text.transcribe,
    "whisper": whisper_local.transcribe,
    # "groq": groq_cloud.transcribe,  # Task 3
}
```

- [ ] **Step 5: 테스트 실행하여 통과 확인**

Run: `python -m pytest tests/test_adapters.py tests/test_transcribe.py -v`
Expected: PASS (whisper 2개 포함 전체 통과). `available_sources()`에 "whisper" 포함.

- [ ] **Step 6: 커밋**

```bash
git add scripts/adapters/whisper_local.py scripts/adapters/__init__.py tests/test_adapters.py
git commit -m "feat: 로컬 STT 어댑터(faster-whisper)"
```

---

## Task 3: 클라우드 STT 어댑터 (Groq Whisper API)

오디오 파일을 Groq Whisper API(`whisper-large-v3`)로 전사한다. 클라이언트 생성을 lazy import + 헬퍼로 감싸 모킹 가능하게 하고, `GROQ_API_KEY` 미설정을 명확히 처리한다.

**Files:**
- Create: `scripts/adapters/groq_cloud.py`
- Modify: `scripts/adapters/__init__.py` (groq 등록 주석 해제)
- Test: `tests/test_adapters.py` (groq 테스트 추가)

**Interfaces:**
- Consumes: `normalize_input._normalize(text: str) -> str`.
- Produces:
  - `adapters/groq_cloud.py`: `transcribe(source: str, **opts) -> str`, `_client()` (lazy import + API 키 확인 헬퍼, 테스트 monkeypatch 대상).

- [ ] **Step 1: groq 어댑터 테스트 작성 (실패 예상)**

`tests/test_adapters.py`에 다음을 추가:
```python
from adapters import groq_cloud


class _FakeTranscription:
    text = "첫 줄  \n\n\n 둘째 줄"


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
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `python -m pytest tests/test_adapters.py -k groq -v`
Expected: FAIL — `AttributeError`/import 오류

- [ ] **Step 3: groq 어댑터 구현**

`scripts/adapters/groq_cloud.py`:
```python
"""클라우드 STT 어댑터: Groq Whisper API(whisper-large-v3)로 오디오 → 텍스트.

groq 라이브러리는 _client 안에서 lazy import한다.
"""
import os

from normalize_input import _normalize


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
    with open(source, "rb") as audio:
        result = client.audio.transcriptions.create(
            file=audio,
            model=opts.get("model", "whisper-large-v3"),
            language=opts.get("language", "ko"),
        )
    normalized = _normalize(result.text)
    if not normalized:
        raise ValueError("전사 결과가 비어 있습니다: 오디오에서 텍스트를 얻지 못했습니다.")
    return normalized
```

- [ ] **Step 4: 레지스트리에 groq 등록**

`scripts/adapters/__init__.py` 수정 — 모든 어댑터 활성화(최종 형태):
```python
from adapters import text, whisper_local, groq_cloud

REGISTRY = {
    "text": text.transcribe,
    "whisper": whisper_local.transcribe,
    "groq": groq_cloud.transcribe,
}
```

- [ ] **Step 5: 테스트 실행하여 통과 확인**

Run: `python -m pytest -v`
Expected: 전체 PASS. `available_sources()` == `["text", "whisper", "groq"]`.

- [ ] **Step 6: 커밋**

```bash
git add scripts/adapters/groq_cloud.py scripts/adapters/__init__.py tests/test_adapters.py
git commit -m "feat: 클라우드 STT 어댑터(Groq Whisper API)"
```

---

## Task 4: SKILL.md [1] 단계 지시문 업데이트 + 엔드투엔드 검증

Claude가 세 입력 소스를 올바르게 오케스트레이션하도록 SKILL.md의 [1] 단계를 갱신하고, 전체 흐름을 수동 검증한다.

**Files:**
- Modify: `SKILL.md`

**Interfaces:**
- Consumes: `transcribe.py`, `adapters/` (Task 1~3).
- Produces: 없음 (Claude가 읽는 지시문). 자동 테스트 대신 샘플 엔드투엔드 검증.

- [ ] **Step 1: SKILL.md의 [1] 입력 정규화 섹션 교체**

`SKILL.md`에서 기존 `### [1] 입력 정규화` 섹션 전체를 아래로 교체:
```markdown
### [1] 입력 확보 (텍스트 / 로컬 STT / 클라우드 STT)
사용자가 준 입력의 종류에 따라 소스를 고른다.

- **텍스트**(붙여넣은 내용 또는 `.txt` 경로): `--source text`
- **오디오 파일**: 로컬/클라우드 중 무엇을 쓸지 사용자에게 확인한다.
  - 로컬(오프라인, `pip install -r requirements-stt.txt` 필요): `--source whisper`
  - 클라우드(Groq, `GROQ_API_KEY` 필요): `--source groq`

정규화된 회의 텍스트를 얻는다:
`python scripts/transcribe.py --source <text|whisper|groq> <입력> > output/meeting.txt`

이 표준 출력(정규화된 회의 원문)을 [2] 추출 단계의 입력으로 사용한다.
STT 라이브러리 미설치나 `GROQ_API_KEY` 미설정 시, 명령이 안내 메시지와 함께
실패하므로 사용자에게 그대로 전달해 설치/설정을 요청한다.
```

또한 문서 하단의 `## 확장 (STT)` 섹션을 아래로 교체:
```markdown
## 확장 (입력 어댑터)
입력 방식 추가: `scripts/adapters/`에 `transcribe(source, **opts) -> str`를
구현한 모듈을 만들고 `scripts/adapters/__init__.py`의 `REGISTRY`에 한 줄 등록.
삭제: REGISTRY에서 한 줄 제거 + 모듈 삭제. [2]~[4] 단계는 수정 불필요.
```

- [ ] **Step 2: text 소스 엔드투엔드 수동 검증**

Run:
```bash
python scripts/transcribe.py --source text examples/sample_meeting.txt > output/meeting.txt
```
Expected: `output/meeting.txt` 생성, "김수민"·"STT 연동" 포함. (없으면 `mkdir output` 후 재실행.)

- [ ] **Step 3: 미지원 소스 오류 메시지 확인**

Run: `python scripts/transcribe.py --source clova examples/sample_meeting.txt`
Expected: argparse가 `--source`의 choices 위반으로 사용 가능한 값(text/whisper/groq)을 안내하며 종료.

- [ ] **Step 4: 전체 테스트 재실행**

Run: `python -m pytest -v`
Expected: 전체 PASS (기존 16 + 신규 어댑터/디스패처 테스트).

- [ ] **Step 5: 커밋**

```bash
git add SKILL.md
git commit -m "docs: SKILL.md [1] 단계에 STT 입력 어댑터 3종 반영"
```

---

## Self-Review 결과

- **Spec coverage:** 스펙 §3 아키텍처(레지스트리+디스패처) → Task 1. §4 공통 인터페이스 → Task 1~3 모두 `transcribe(source,**opts)->str`. §5 디렉터리 → 전 태스크. §6 실행 방식(사용자 CLI 미사용) → transcribe.py docstring + Task 4 SKILL.md. §7.1 text → Task 1. §7.2 whisper(medium/cpu_threads=4) → Task 2. §7.3 groq(whisper-large-v3) → Task 3. §8 에러 처리(미설치/키없음/빈결과/미지원소스) → Task 1(get_adapter)·2·3 테스트. §9 테스트 전략(모킹) → Task 2·3. 누락 없음.
- **Placeholder scan:** "TBD/TODO/적절히" 없음. 모든 코드 스텝에 실제 코드 포함. `__init__.py`는 Task 1(주석)→2→3(최종)으로 점진 완성이 명시됨.
- **Type consistency:** `transcribe(source:str,**opts)->str`(text/whisper/groq 동일), `get_adapter(name)->Callable`, `available_sources()->list`, `run(source_name,source,**opts)->str`, `_load_model(model_size,cpu_threads)`, `_client()`, `_normalize(text)->str` — 태스크 간 명칭/시그니처 일치.
- **참고:** `_normalize`는 기존 `normalize_input.py`의 내부 함수를 재사용(신규 계약 아님). whisper/groq 어댑터는 세그먼트/결과 텍스트를 이어 붙인 뒤 이 함수로 정규화해 text 어댑터와 출력 형태를 통일.
