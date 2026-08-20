
# 회의록 자동 생성 Skill Implementation Plan (계획서)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 텍스트 회의 내용을 입력받아 구조화된 `minutes.json`을 거쳐 Markdown과 Word(.docx) 회의록을 생성하는 Claude Code Skill을 구현한다.

**Architecture:** 명시적 4단계 파이프라인 — [1] 입력 정규화 → [2] LLM 추출(`minutes.json`) → [3] Markdown 렌더 → [4] docx 렌더. `minutes.json`이 단일 진실 원천이며 Markdown·docx 렌더러가 이를 공유한다. [2] 추출만 `SKILL.md` 지시문에 따라 Claude가 수행하고, 나머지는 결정적 Python 헬퍼 스크립트다.

**Tech Stack:** Python 3, python-docx, jsonschema, pytest.

## Global Constraints

- Python 3.9+ 사용 (표준 라이브러리 `pathlib`, 타입 힌트).
- 외부 바이너리 의존 금지 — 순수 pip 패키지만 (`python-docx`, `jsonschema`, `pytest`).
- 원문에 없는 정보는 지어내지 않는다 — 누락 시 `null` 또는 빈 배열.
- 추출 결과 6종 항목: `title`(+`date`), `attendees`, `discussion`, `decisions`, `action_items`, `next_meeting`.
- `minutes.json`은 렌더링 전 반드시 스키마 검증을 통과해야 한다.
- 모든 파일 입출력 인코딩은 `utf-8`.
- 스크립트는 `scripts/` 디렉터리에 두고, 서로를 형제 모듈로 import (`python scripts/xxx.py` 실행 시 `scripts/`가 `sys.path[0]`이 됨).

---

## File Structure

- `requirements.txt` — 필요 패키지.
- `schema/minutes.schema.json` — 중간 계약(JSON Schema draft-07).
- `scripts/validate.py` — 스키마 검증 및 로드.
- `scripts/normalize_input.py` — [1] 입력 정규화 + STT 어댑터 자리.
- `scripts/render_markdown.py` — [3] JSON → Markdown.
- `scripts/render_docx.py` — [4] JSON → docx.
- `examples/sample_meeting.txt` — 샘플 입력.
- `examples/sample_minutes.json` — 기대 추출 결과(테스트 픽스처).
- `tests/conftest.py` — `scripts/`를 import 경로에 추가.
- `tests/test_schema.py`, `tests/test_normalize_input.py`, `tests/test_render_markdown.py`, `tests/test_render_docx.py` — 단위 테스트.
- `SKILL.md` — 스킬 진입점(4단계 오케스트레이션 지시문).

---

## Task 1: 스키마 + 검증 모듈 + 픽스처

**Files:**
- Create: `requirements.txt`
- Create: `schema/minutes.schema.json`
- Create: `examples/sample_minutes.json`
- Create: `scripts/validate.py`
- Create: `tests/conftest.py`
- Test: `tests/test_schema.py`

**Interfaces:**
- Consumes: (없음 — 첫 태스크)
- Produces:
  - `scripts/validate.py`: `validate_minutes(data: dict) -> None` (위반 시 `jsonschema.ValidationError`), `load_minutes(json_path: str) -> dict` (읽고 검증 후 dict 반환).
  - `examples/sample_minutes.json`: 이후 모든 렌더러 테스트의 공용 픽스처.

- [ ] **Step 1: 필요 패키지 파일 작성**

`requirements.txt`:
```
python-docx>=1.1.0
jsonschema>=4.0.0
pytest>=7.0.0
```

설치: `pip install -r requirements.txt`

- [ ] **Step 2: JSON 스키마 작성**

`schema/minutes.schema.json`:
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "Meeting Minutes",
  "type": "object",
  "required": ["title", "date", "attendees", "discussion", "decisions", "action_items", "next_meeting"],
  "additionalProperties": false,
  "properties": {
    "title": { "type": "string" },
    "date": { "type": ["string", "null"] },
    "attendees": { "type": "array", "items": { "type": "string" } },
    "discussion": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["topic", "points"],
        "additionalProperties": false,
        "properties": {
          "topic": { "type": "string" },
          "points": { "type": "array", "items": { "type": "string" } }
        }
      }
    },
    "decisions": { "type": "array", "items": { "type": "string" } },
    "action_items": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["task", "owner", "due"],
        "additionalProperties": false,
        "properties": {
          "task": { "type": "string" },
          "owner": { "type": ["string", "null"] },
          "due": { "type": ["string", "null"] }
        }
      }
    },
    "next_meeting": { "type": ["string", "null"] }
  }
}
```

- [ ] **Step 3: 샘플 픽스처 작성**

`examples/sample_minutes.json`:
```json
{
  "title": "2026 3분기 제품 로드맵 회의",
  "date": "2026-07-16 14:00",
  "attendees": ["김수민", "이정우", "박서연"],
  "discussion": [
    {
      "topic": "STT 연동 우선순위",
      "points": ["실시간 STT는 3분기 범위에서 제외", "배치 STT부터 도입하기로 의견 수렴"]
    },
    {
      "topic": "출력 포맷",
      "points": ["Markdown 미리보기 후 docx 확정 흐름 확인"]
    }
  ],
  "decisions": ["docx 변환은 python-docx로 진행", "STT는 입력 어댑터 자리만 우선 확보"],
  "action_items": [
    { "task": "python-docx 렌더러 PoC 작성", "owner": "이정우", "due": "2026-07-23" },
    { "task": "샘플 회의 원문 수집", "owner": "박서연", "due": null }
  ],
  "next_meeting": "2026-07-23 14:00"
}
```

- [ ] **Step 4: 검증 모듈 테스트 작성 (실패 예상)**

`tests/conftest.py`:
```python
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
```

`tests/test_schema.py`:
```python
import json
from pathlib import Path

import pytest
from jsonschema import ValidationError

from validate import validate_minutes, load_minutes

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"


def test_sample_fixture_is_valid():
    validate_minutes(json.loads(SAMPLE.read_text(encoding="utf-8")))


def test_load_minutes_returns_dict():
    data = load_minutes(str(SAMPLE))
    assert data["title"] == "2026 3분기 제품 로드맵 회의"
    assert len(data["action_items"]) == 2


def test_missing_required_field_rejected():
    bad = json.loads(SAMPLE.read_text(encoding="utf-8"))
    del bad["decisions"]
    with pytest.raises(ValidationError):
        validate_minutes(bad)


def test_wrong_type_rejected():
    bad = json.loads(SAMPLE.read_text(encoding="utf-8"))
    bad["attendees"] = "김수민"  # 배열이어야 함
    with pytest.raises(ValidationError):
        validate_minutes(bad)
```

- [ ] **Step 5: 테스트 실행하여 실패 확인**

Run: `pytest tests/test_schema.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'validate'`

- [ ] **Step 6: 검증 모듈 구현**

`scripts/validate.py`:
```python
"""minutes.json 스키마 검증 및 로드."""
import json
from pathlib import Path

import jsonschema

_SCHEMA_PATH = Path(__file__).resolve().parent.parent / "schema" / "minutes.schema.json"


def _load_schema() -> dict:
    with open(_SCHEMA_PATH, encoding="utf-8") as f:
        return json.load(f)


def validate_minutes(data: dict) -> None:
    """스키마 위반 시 jsonschema.ValidationError 발생."""
    jsonschema.validate(instance=data, schema=_load_schema())


def load_minutes(json_path: str) -> dict:
    """JSON 파일을 읽어 검증 후 dict 반환."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    validate_minutes(data)
    return data
```

- [ ] **Step 7: 테스트 실행하여 통과 확인**

Run: `pytest tests/test_schema.py -v`
Expected: PASS (4 passed)

- [ ] **Step 8: 커밋**

```bash
git add requirements.txt schema examples/sample_minutes.json scripts/validate.py tests/conftest.py tests/test_schema.py
git commit -m "feat: minutes.json 스키마 및 검증 모듈"
```

---

## Task 2: 입력 정규화 계층

**Files:**
- Create: `scripts/normalize_input.py`
- Create: `examples/sample_meeting.txt`
- Test: `tests/test_normalize_input.py`

**Interfaces:**
- Consumes: (없음)
- Produces: `scripts/normalize_input.py`: `load_meeting_text(source: str) -> str` — `source`가 존재하는 파일 경로면 파일을 읽고, 아니면 문자열 자체를 원문으로 간주. 정규화(줄 끝 공백 제거, 연속 빈 줄 1개로 축약, 앞뒤 공백 제거) 후 반환. 빈 결과면 `ValueError`.

- [ ] **Step 1: 샘플 입력 파일 작성**

`examples/sample_meeting.txt`:
```
2026 3분기 제품 로드맵 회의 (2026-07-16 14:00)
참석: 김수민, 이정우, 박서연

김수민: 이번 분기 STT 연동 어디까지 갈까요.
이정우: 실시간은 무리고, 배치 STT부터 하시죠.
박서연: 동의합니다.
김수민: 출력은 Markdown 먼저 보고 docx로 확정하는 걸로.

결정: docx는 python-docx로. STT는 입력 어댑터 자리만 확보.
이정우님이 다음 주까지 python-docx PoC 작성.
다음 회의는 2026-07-23 14시.
```

- [ ] **Step 2: 정규화 테스트 작성 (실패 예상)**

`tests/test_normalize_input.py`:
```python
from pathlib import Path

import pytest

from normalize_input import load_meeting_text

SAMPLE_TXT = Path(__file__).resolve().parent.parent / "examples" / "sample_meeting.txt"


def test_load_from_string():
    result = load_meeting_text("회의 시작\n논의 내용")
    assert result == "회의 시작\n논의 내용"


def test_load_from_file():
    result = load_meeting_text(str(SAMPLE_TXT))
    assert "STT 연동" in result
    assert "김수민" in result


def test_trailing_whitespace_and_blank_lines_collapsed():
    result = load_meeting_text("첫 줄   \n\n\n\n둘째 줄  ")
    assert result == "첫 줄\n\n둘째 줄"


def test_empty_input_raises():
    with pytest.raises(ValueError):
        load_meeting_text("   \n\n  ")
```

- [ ] **Step 3: 테스트 실행하여 실패 확인**

Run: `pytest tests/test_normalize_input.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'normalize_input'`

- [ ] **Step 4: 정규화 모듈 구현**

`scripts/normalize_input.py`:
```python
"""입력 계층: 회의 원문 텍스트를 정규화. STT 어댑터 자리 확보."""
from pathlib import Path


def load_meeting_text(source: str) -> str:
    """텍스트 문자열 또는 .txt 파일 경로를 정규화된 회의 원문으로 변환.

    - source가 존재하는 파일 경로면 파일 내용을 읽음.
    - 그렇지 않으면 source 자체를 원문 텍스트로 간주.
    """
    candidate = Path(source)
    if candidate.exists() and candidate.is_file():
        text = candidate.read_text(encoding="utf-8")
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


# --- STT 확장 자리 (이번 범위에서 구현하지 않음) ---
# def load_from_stt(audio_path: str) -> str:
#     """STT 추가 시 이 함수만 구현하면 하위 단계([2]~[4]) 무수정."""
#     raise NotImplementedError("STT 입력은 아직 지원하지 않습니다.")
```

- [ ] **Step 5: 테스트 실행하여 통과 확인**

Run: `pytest tests/test_normalize_input.py -v`
Expected: PASS (4 passed)

- [ ] **Step 6: 커밋**

```bash
git add scripts/normalize_input.py examples/sample_meeting.txt tests/test_normalize_input.py
git commit -m "feat: 입력 정규화 계층 및 STT 어댑터 자리"
```

---

## Task 3: Markdown 렌더러

**Files:**
- Create: `scripts/render_markdown.py`
- Test: `tests/test_render_markdown.py`

**Interfaces:**
- Consumes: `validate.load_minutes(json_path) -> dict`, `examples/sample_minutes.json`.
- Produces: `scripts/render_markdown.py`: `render_markdown(data: dict) -> str`. CLI: `python scripts/render_markdown.py <in.json> <out.md>`.

- [ ] **Step 1: Markdown 렌더 테스트 작성 (실패 예상)**

`tests/test_render_markdown.py`:
```python
import json
from pathlib import Path

from render_markdown import render_markdown

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"


def _sample():
    return json.loads(SAMPLE.read_text(encoding="utf-8"))


def test_title_and_sections_present():
    md = render_markdown(_sample())
    assert md.startswith("# 2026 3분기 제품 로드맵 회의")
    for heading in ["## 참석자", "## 논의 내용", "## 결정 사항",
                    "## Action Items", "## 다음 회의 일정"]:
        assert heading in md


def test_action_items_table():
    md = render_markdown(_sample())
    assert "| 할 일 | 담당자 | 기한 |" in md
    assert "| python-docx 렌더러 PoC 작성 | 이정우 | 2026-07-23 |" in md


def test_null_due_rendered_as_dash():
    md = render_markdown(_sample())
    assert "| 샘플 회의 원문 수집 | 박서연 | - |" in md


def test_empty_and_null_fields():
    data = _sample()
    data["attendees"] = []
    data["action_items"] = []
    data["next_meeting"] = None
    md = render_markdown(data)
    assert "- (없음)" in md          # 빈 참석자
    assert "(미정)" in md            # next_meeting None
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `pytest tests/test_render_markdown.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'render_markdown'`

- [ ] **Step 3: Markdown 렌더러 구현**

`scripts/render_markdown.py`:
```python
"""[3] Markdown 렌더러: minutes.json → Markdown 문자열."""
import sys

from validate import load_minutes


def render_markdown(data: dict) -> str:
    lines = [f"# {data['title']}", ""]
    if data.get("date"):
        lines.append(f"**일시:** {data['date']}")
        lines.append("")

    lines.append("## 참석자")
    if data["attendees"]:
        lines.extend(f"- {a}" for a in data["attendees"])
    else:
        lines.append("- (없음)")
    lines.append("")

    lines.append("## 논의 내용")
    if data["discussion"]:
        for item in data["discussion"]:
            lines.append(f"### {item['topic']}")
            lines.extend(f"- {p}" for p in item["points"])
            lines.append("")
    else:
        lines.append("(없음)")
        lines.append("")

    lines.append("## 결정 사항")
    if data["decisions"]:
        lines.extend(f"- {d}" for d in data["decisions"])
    else:
        lines.append("- (없음)")
    lines.append("")

    lines.append("## Action Items")
    if data["action_items"]:
        lines.append("| 할 일 | 담당자 | 기한 |")
        lines.append("| --- | --- | --- |")
        for a in data["action_items"]:
            owner = a["owner"] or "-"
            due = a["due"] or "-"
            lines.append(f"| {a['task']} | {owner} | {due} |")
    else:
        lines.append("(없음)")
    lines.append("")

    lines.append("## 다음 회의 일정")
    lines.append(data["next_meeting"] or "(미정)")
    lines.append("")

    return "\n".join(lines)


def main() -> None:
    json_path, out_path = sys.argv[1], sys.argv[2]
    data = load_minutes(json_path)
    md = render_markdown(data)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md)
    print(f"Markdown 생성 완료: {out_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `pytest tests/test_render_markdown.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: CLI 수동 확인**

Run: `python scripts/render_markdown.py examples/sample_minutes.json examples/sample_minutes.md`
Expected: `Markdown 생성 완료: examples/sample_minutes.md` 출력, 파일 생성됨.

- [ ] **Step 6: 커밋**

```bash
git add scripts/render_markdown.py tests/test_render_markdown.py
git commit -m "feat: Markdown 렌더러"
```

---

## Task 4: docx 렌더러

**Files:**
- Create: `scripts/render_docx.py`
- Test: `tests/test_render_docx.py`

**Interfaces:**
- Consumes: `validate.load_minutes(json_path) -> dict`, `examples/sample_minutes.json`, `python-docx`.
- Produces: `scripts/render_docx.py`: `render_docx(data: dict, out_path: str) -> None`. CLI: `python scripts/render_docx.py <in.json> <out.docx>`.

- [ ] **Step 1: docx 렌더 테스트 작성 (실패 예상)**

`tests/test_render_docx.py`:
```python
import json
from pathlib import Path

from docx import Document

from render_docx import render_docx

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"


def _sample():
    return json.loads(SAMPLE.read_text(encoding="utf-8"))


def test_creates_valid_docx(tmp_path):
    out = tmp_path / "out.docx"
    render_docx(_sample(), str(out))
    assert out.exists()
    doc = Document(str(out))  # 열리면 유효한 docx
    texts = [p.text for p in doc.paragraphs]
    assert "2026 3분기 제품 로드맵 회의" in texts


def test_headings_present(tmp_path):
    out = tmp_path / "out.docx"
    render_docx(_sample(), str(out))
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    for heading in ["참석자", "논의 내용", "결정 사항", "Action Items", "다음 회의 일정"]:
        assert heading in texts


def test_action_items_table_has_rows(tmp_path):
    out = tmp_path / "out.docx"
    render_docx(_sample(), str(out))
    doc = Document(str(out))
    assert len(doc.tables) == 1
    # 헤더 1행 + action_items 2행
    assert len(doc.tables[0].rows) == 3


def test_null_next_meeting(tmp_path):
    data = _sample()
    data["next_meeting"] = None
    out = tmp_path / "out.docx"
    render_docx(data, str(out))
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert "(미정)" in texts
```

- [ ] **Step 2: 테스트 실행하여 실패 확인**

Run: `pytest tests/test_render_docx.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'render_docx'`

- [ ] **Step 3: docx 렌더러 구현**

`scripts/render_docx.py`:
```python
"""[4] docx 렌더러: minutes.json → .docx (python-docx)."""
import sys

from docx import Document

from validate import load_minutes


def render_docx(data: dict, out_path: str) -> None:
    doc = Document()
    doc.add_heading(data["title"], level=0)
    if data.get("date"):
        doc.add_paragraph(f"일시: {data['date']}")

    doc.add_heading("참석자", level=1)
    if data["attendees"]:
        for a in data["attendees"]:
            doc.add_paragraph(a, style="List Bullet")
    else:
        doc.add_paragraph("(없음)")

    doc.add_heading("논의 내용", level=1)
    if data["discussion"]:
        for item in data["discussion"]:
            doc.add_heading(item["topic"], level=2)
            for p in item["points"]:
                doc.add_paragraph(p, style="List Bullet")
    else:
        doc.add_paragraph("(없음)")

    doc.add_heading("결정 사항", level=1)
    if data["decisions"]:
        for d in data["decisions"]:
            doc.add_paragraph(d, style="List Bullet")
    else:
        doc.add_paragraph("(없음)")

    doc.add_heading("Action Items", level=1)
    if data["action_items"]:
        table = doc.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        hdr = table.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = "할 일", "담당자", "기한"
        for a in data["action_items"]:
            row = table.add_row().cells
            row[0].text = a["task"]
            row[1].text = a["owner"] or "-"
            row[2].text = a["due"] or "-"
    else:
        doc.add_paragraph("(없음)")

    doc.add_heading("다음 회의 일정", level=1)
    doc.add_paragraph(data["next_meeting"] or "(미정)")

    doc.save(out_path)


def main() -> None:
    json_path, out_path = sys.argv[1], sys.argv[2]
    data = load_minutes(json_path)
    render_docx(data, out_path)
    print(f"docx 생성 완료: {out_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 테스트 실행하여 통과 확인**

Run: `pytest tests/test_render_docx.py -v`
Expected: PASS (4 passed)

- [ ] **Step 5: 전체 테스트 실행**

Run: `pytest -v`
Expected: 모든 테스트 PASS.

- [ ] **Step 6: 커밋**

```bash
git add scripts/render_docx.py tests/test_render_docx.py
git commit -m "feat: docx 렌더러"
```

---

## Task 5: SKILL.md 오케스트레이션 지시문

**Files:**
- Create: `SKILL.md`

**Interfaces:**
- Consumes: 위 모든 스크립트(`normalize_input.py`, `validate.py`, `render_markdown.py`, `render_docx.py`)와 스키마.
- Produces: 없음(Claude가 읽고 실행하는 지시문). 자동 테스트 대신 샘플로 엔드투엔드 수동 검증.

- [ ] **Step 1: SKILL.md 작성**

`SKILL.md`:
```markdown
---
name: meeting-minutes
description: Use when the user provides meeting notes, a transcript, or STT output and wants structured meeting minutes exported to Markdown and Word (.docx). Extracts 회의 주제/참석자/논의 내용/결정 사항/Action Items/다음 회의 일정.
---

# 회의록 자동 생성 Skill

텍스트 회의 내용을 구조화된 회의록(Markdown + Word .docx)으로 변환한다.
파이프라인은 4단계이며 각 단계의 계약은 `schema/minutes.schema.json`이다.

## 사전 준비

필요 패키지 설치(최초 1회): `pip install -r requirements.txt`

## 단계별 절차

### [1] 입력 정규화
사용자의 회의 원문(붙여넣은 텍스트 또는 .txt 경로)을 확보한다.
정규화가 필요하면:
`python scripts/normalize_input.py` 의 `load_meeting_text(source)`를 사용하거나,
텍스트를 직접 다음 단계로 넘긴다.

### [2] 추출 (이 단계는 네가 직접 수행)
회의 원문을 읽고 `schema/minutes.schema.json`을 **정확히** 따르는
`minutes.json`을 작성한다. 규칙:
- 6개 항목을 모두 채운다: title, date, attendees, discussion, decisions, action_items, next_meeting.
- 원문에 없는 정보는 **지어내지 않는다**. 없으면 `null`(date/next_meeting/owner/due) 또는 빈 배열.
- `discussion`은 주제별로 `{topic, points}`로 묶는다.
- 결과를 `output/minutes.json`으로 저장한다.

검증: `python scripts/validate.py` 는 별도 CLI가 없으므로,
`render_markdown.py`/`render_docx.py`가 `load_minutes()`로 로드하며 자동 검증한다.
검증 실패(ValidationError) 시 `minutes.json`을 스키마에 맞게 수정한다.

### [3] Markdown 생성 및 사용자 검토
`python scripts/render_markdown.py output/minutes.json output/minutes.md`
생성된 Markdown을 사용자에게 보여주고 검토를 요청한다.
사용자가 수정을 요청하면 `minutes.json`을 고치고 이 단계를 반복한다.

### [4] docx 생성 (사용자 승인 후)
`python scripts/render_docx.py output/minutes.json output/minutes.docx`
최종 `.docx` 경로를 사용자에게 안내한다.

## 확장 (STT)
STT 입력은 현재 미지원. 추가 시 `scripts/normalize_input.py`의
`load_from_stt(audio_path)`만 구현하면 [2]~[4]는 수정 불필요.
```

- [ ] **Step 2: 엔드투엔드 수동 검증**

Run:
```bash
python scripts/render_markdown.py examples/sample_minutes.json output/minutes.md
python scripts/render_docx.py examples/sample_minutes.json output/minutes.docx
```
Expected: `output/minutes.md`와 `output/minutes.docx`가 생성되고, Markdown에 5개 섹션과 Action Items 표가, docx가 정상적으로 열림.

- [ ] **Step 3: 전체 테스트 재실행**

Run: `pytest -v`
Expected: 전체 PASS.

- [ ] **Step 4: 커밋**

```bash
git add SKILL.md
git commit -m "feat: 회의록 생성 스킬 오케스트레이션 지시문(SKILL.md)"
```

---

## Self-Review 결과

- **Spec coverage:** 스펙 §3 파이프라인 4단계 → Task 2([1]), SKILL.md §[2]추출(Task 5), Task 3([3]), Task 4([4]). §4 스키마 → Task 1. §5 디렉터리 → 전 태스크. §6 STT 자리 → Task 2. §7 에러 처리(스키마 검증/빈 입력/null·빈 배열) → Task 1·2·3·4 테스트. §8 테스트 전략 → 각 태스크 TDD. 누락 없음.
- **Placeholder scan:** "TBD/TODO/적절히 처리" 없음. 모든 코드 스텝에 실제 코드 포함.
- **Type consistency:** `load_meeting_text(str)->str`, `validate_minutes(dict)->None`, `load_minutes(str)->dict`, `render_markdown(dict)->str`, `render_docx(dict, str)->None` — 태스크 간 명칭/시그니처 일치 확인.
- **참고:** `output/` 디렉터리는 실행 시 생성 필요. SKILL.md 실행 시 없으면 만들 것(또는 실행 전 `mkdir output`).
