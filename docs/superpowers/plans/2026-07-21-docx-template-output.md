# 템플릿 기반 docx 출력 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 사용자가 제공한 `.docx` 양식(스마트 토큰)에 회의록 내용을 채워 레이아웃을 보존한 docx로 출력하는 렌더러를 추가한다.

**Architecture:** 기존 파이프라인([1]입력→[2]추출→[3]Markdown→[4]docx)은 불변. [4] docx 단계에 `render_docx_template.py`라는 두 번째 렌더러 갈래를 추가한다. `minutes.json`(9개 항목 스키마)을 `docxtpl`(Jinja2) 컨텍스트로 변환해 템플릿의 표시자를 치환한다. 템플릿이 없으면 기존 `render_docx.py`를 그대로 쓴다.

**Tech Stack:** Python 3.9+, docxtpl(Jinja2 기반 docx 템플릿), python-docx(기존), jsonschema(기존), pytest.

## Global Constraints

- 작업 루트는 스킬 디렉토리 `C:\Users\user\Desktop\metting_form\.claude\skills\meeting-minutes\` 이다. 아래 모든 상대경로는 이 디렉토리 기준이다.
- 기존 파이프라인([1]~[3], `render_docx.py`, `schema/minutes.schema.json`, 추출 항목 9종)은 **수정하지 않는다.**
- 데이터 계약은 기존 `minutes.json` 스키마 그대로. `title`, `attendees`, `discussion`, `decisions`, `action_items`, `notes`는 항상 존재(빈 배열 가능), `date`/`purpose`/`next_meeting`/`action_items[].owner`/`action_items[].due`는 `null` 가능.
- 표시자 토큰 이름은 **영어(ASCII)** 를 쓴다. 한글(비ASCII) 토큰은 범위 밖.
- 테스트는 `tests/` 아래에 두며, `tests/conftest.py`가 `scripts/`를 `sys.path`에 넣으므로 테스트에서 `from render_docx_template import ...` 형태로 import 한다.
- 커밋 메시지는 한국어 요약 + 다음 트레일러로 끝낸다:
  `Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>` / `Claude-Session: https://claude.ai/code/session_01ENQ6TTxgxz5QFo3MLe6mKF`
- 테스트 실행 명령: `python -m pytest tests -q` (스킬 디렉토리에서).

---

### Task 1: `build_context` — minutes.json → 렌더 컨텍스트 (순수 함수)

`docxtpl` 없이도 단위 테스트 가능한 순수 변환 함수부터 만든다. null 처리(스칼라→`""`, owner/due→`"-"`)와 `attendees_joined` 생성이 이 함수의 책임이다.

**Files:**
- Create: `scripts/render_docx_template.py`
- Test: `tests/test_render_docx_template.py`

**Interfaces:**
- Consumes: `minutes.json` 스키마 dict (기존 `examples/sample_minutes.json` 형태).
- Produces: `build_context(data: dict) -> dict` — 키: `title, date, purpose, next_meeting`(str), `attendees`(list[str]), `attendees_joined`(str), `discussion`(list[{topic, points}]), `decisions`(list[str]), `action_items`(list[{task, owner, due}] — owner/due는 절대 null 아님), `notes`(list[str]).

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_render_docx_template.py` 생성:

```python
import json
from pathlib import Path

import pytest
from docx import Document

from render_docx_template import build_context

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"


def _sample():
    return json.loads(SAMPLE.read_text(encoding="utf-8"))


def test_build_context_passes_scalars_and_joins_attendees():
    ctx = build_context(_sample())
    assert ctx["title"] == "2026 3분기 제품 로드맵 회의"
    assert ctx["date"] == "2026-07-16 14:00"
    assert ctx["attendees_joined"] == "김수민, 이정우, 박서연"
    assert ctx["attendees"] == ["김수민", "이정우", "박서연"]


def test_build_context_null_scalars_become_empty_string():
    data = _sample()
    data["date"] = None
    data["purpose"] = None
    data["next_meeting"] = None
    ctx = build_context(data)
    assert ctx["date"] == ""
    assert ctx["purpose"] == ""
    assert ctx["next_meeting"] == ""


def test_build_context_null_owner_due_become_dash():
    ctx = build_context(_sample())
    # 샘플 2번째 action_item은 due=null
    tasks = {a["task"]: a for a in ctx["action_items"]}
    assert tasks["샘플 회의 원문 수집"]["due"] == "-"
    assert tasks["샘플 회의 원문 수집"]["owner"] == "박서연"


def test_build_context_empty_attendees_joined_is_empty():
    data = _sample()
    data["attendees"] = []
    ctx = build_context(data)
    assert ctx["attendees_joined"] == ""
    assert ctx["attendees"] == []
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_render_docx_template.py -q`
Expected: FAIL — `ImportError` 또는 `ModuleNotFoundError: No module named 'render_docx_template'`

- [ ] **Step 3: 최소 구현 작성**

`scripts/render_docx_template.py` 생성:

```python
"""[4-대체] 템플릿 렌더러: .docx 양식 + minutes.json → 표시자 치환 docx."""
import sys

from normalize_input import resolve_input_path
from validate import load_minutes


def build_context(data: dict) -> dict:
    """minutes.json(dict) → docxtpl 렌더 컨텍스트.

    null 스칼라는 빈 문자열로, action_items의 owner/due null은 "-"로 정규화한다.
    """
    return {
        "title": data["title"],
        "date": data.get("date") or "",
        "purpose": data.get("purpose") or "",
        "next_meeting": data.get("next_meeting") or "",
        "attendees": data["attendees"],
        "attendees_joined": ", ".join(data["attendees"]),
        "discussion": data["discussion"],
        "decisions": data["decisions"],
        "action_items": [
            {"task": a["task"], "owner": a["owner"] or "-", "due": a["due"] or "-"}
            for a in data["action_items"]
        ],
        "notes": data["notes"],
    }
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_render_docx_template.py -q`
Expected: PASS (4 passed)

- [ ] **Step 5: 커밋**

```bash
git add scripts/render_docx_template.py tests/test_render_docx_template.py
git commit -m "feat: build_context 추가 (minutes.json → docxtpl 컨텍스트)"
```

---

### Task 2: 템플릿 렌더 + CLI + 오류 처리

`docxtpl`로 실제 표시자 치환과 표 행 자동 반복을 수행하고, CLI 진입점과 오류 처리를 붙인다.

**Files:**
- Modify: `requirements.txt` (docxtpl 추가)
- Modify: `scripts/render_docx_template.py` (render_template + main 추가)
- Modify: `tests/test_render_docx_template.py` (렌더/표/오류 테스트 추가)

**Interfaces:**
- Consumes: Task 1의 `build_context`, 기존 `resolve_input_path`(파일 탐색·존재 검증), 기존 `load_minutes`(스키마 검증 로드).
- Produces:
  - `render_template(template_path: str, data: dict, out_path: str) -> None` — 템플릿을 렌더해 `out_path`에 저장. `.docx`가 아니면 `ValueError`, 파일 없으면 `FileNotFoundError`, docxtpl 미설치면 `ImportError`(설치 안내).
  - `main()` — `sys.argv[1:4] = template_path, json_path, out_path`.

- [ ] **Step 1: docxtpl 의존성 추가 및 설치**

`requirements.txt`를 다음으로 수정(마지막 줄 추가):

```
python-docx>=1.1.0
jsonschema>=4.0.0
pytest>=7.0.0
docxtpl>=0.16.0
```

설치: `pip install -r requirements.txt`
확인: `python -c "from docxtpl import DocxTemplate; print('ok')"` → `ok`

- [ ] **Step 2: 실패하는 테스트 작성**

`tests/test_render_docx_template.py`의 맨 위 import 아래에 템플릿 빌더 헬퍼와 테스트를 추가한다.

```python
def _make_template(path):
    """토큰 + {%tr%} 표를 담은 최소 템플릿 docx를 만든다."""
    doc = Document()
    doc.add_paragraph("{{ title }}")
    doc.add_paragraph("일시: {{ date }}")
    doc.add_paragraph("목적: {{ purpose }}")
    doc.add_paragraph("참석자: {{ attendees_joined }}")
    doc.add_paragraph("다음 회의: {{ next_meeting }}")
    # 실행 항목 표: {%tr for%} 행 / 데이터 행 / {%tr endfor%} 행 (3행 구조).
    # docxtpl는 {%tr ...%}가 든 행 전체를 삭제하고 for·endfor 사이 행을 반복한다.
    # (for와 endfor를 한 행에 함께 넣으면 파싱 실패하므로 반드시 분리)
    table = doc.add_table(rows=4, cols=3)
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "할 일", "담당자", "기한"
    table.rows[1].cells[0].text = "{%tr for a in action_items %}"
    data = table.rows[2].cells
    data[0].text = "{{ a.task }}"
    data[1].text = "{{ a.owner }}"
    data[2].text = "{{ a.due }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"
    doc.save(str(path))


def _rendered_paragraphs(out_path):
    return [p.text for p in Document(str(out_path)).paragraphs]


def test_simple_tokens_substituted(tmp_path):
    from render_docx_template import render_template
    tpl = tmp_path / "tpl.docx"
    out = tmp_path / "out.docx"
    _make_template(tpl)
    render_template(str(tpl), _sample(), str(out))
    texts = _rendered_paragraphs(out)
    assert "2026 3분기 제품 로드맵 회의" in texts
    assert "일시: 2026-07-16 14:00" in texts
    assert "참석자: 김수민, 이정우, 박서연" in texts


def test_action_items_table_rows_repeat(tmp_path):
    from render_docx_template import render_template
    tpl = tmp_path / "tpl.docx"
    out = tmp_path / "out.docx"
    _make_template(tpl)
    render_template(str(tpl), _sample(), str(out))
    table = Document(str(out)).tables[0]
    # 헤더 1행 + action_items 2행
    assert len(table.rows) == 3
    body_tasks = [table.rows[i].cells[0].text for i in (1, 2)]
    assert "python-docx 렌더러 PoC 작성" in body_tasks
    assert "샘플 회의 원문 수집" in body_tasks


def test_null_scalar_renders_blank(tmp_path):
    from render_docx_template import render_template
    data = _sample()
    data["next_meeting"] = None
    tpl = tmp_path / "tpl.docx"
    out = tmp_path / "out.docx"
    _make_template(tpl)
    render_template(str(tpl), data, str(out))
    assert "다음 회의: " in _rendered_paragraphs(out)


def test_null_due_renders_dash(tmp_path):
    from render_docx_template import render_template
    tpl = tmp_path / "tpl.docx"
    out = tmp_path / "out.docx"
    _make_template(tpl)
    render_template(str(tpl), _sample(), str(out))
    table = Document(str(out)).tables[0]
    due_cells = [table.rows[i].cells[2].text for i in (1, 2)]
    assert "-" in due_cells


def test_empty_action_items_leaves_header_only(tmp_path):
    from render_docx_template import render_template
    data = _sample()
    data["action_items"] = []
    tpl = tmp_path / "tpl.docx"
    out = tmp_path / "out.docx"
    _make_template(tpl)
    render_template(str(tpl), data, str(out))
    table = Document(str(out)).tables[0]
    assert len(table.rows) == 1  # 헤더만 남음


def test_non_docx_template_raises(tmp_path):
    from render_docx_template import render_template
    bad = tmp_path / "form.txt"
    bad.write_text("not a docx", encoding="utf-8")
    with pytest.raises(ValueError, match="docx"):
        render_template(str(bad), _sample(), str(tmp_path / "out.docx"))


def test_missing_template_raises(tmp_path):
    from render_docx_template import render_template
    missing = tmp_path / "nope.docx"  # 존재하지 않는 절대경로
    with pytest.raises(FileNotFoundError):
        render_template(str(missing), _sample(), str(tmp_path / "out.docx"))
```

- [ ] **Step 3: 테스트 실패 확인**

Run: `python -m pytest tests/test_render_docx_template.py -q`
Expected: FAIL — `render_template`가 아직 없어 `ImportError` / `cannot import name 'render_template'`

- [ ] **Step 4: render_template + main 구현**

`scripts/render_docx_template.py`에 아래를 추가한다(파일 하단, `build_context` 아래):

```python
def _load_docxtemplate():
    try:
        from docxtpl import DocxTemplate
    except ImportError as e:
        raise ImportError(
            "docxtpl가 설치되어 있지 않습니다. "
            "'pip install -r requirements.txt'로 설치하세요."
        ) from e
    return DocxTemplate


def render_template(template_path: str, data: dict, out_path: str) -> None:
    """.docx 템플릿의 표시자를 minutes 데이터로 치환해 out_path에 저장한다."""
    resolved = resolve_input_path(template_path, must_exist=True)
    if resolved.suffix.lower() != ".docx":
        raise ValueError(
            f"템플릿은 .docx여야 합니다: {resolved.name}. "
            "hwp/pdf 양식이면 한글/워드에서 .docx로 저장해 다시 주세요."
        )
    DocxTemplate = _load_docxtemplate()
    tpl = DocxTemplate(str(resolved))
    tpl.render(build_context(data))
    tpl.save(out_path)


def main() -> None:
    template_path, json_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    data = load_minutes(json_path)
    render_template(template_path, data, out_path)
    print(f"템플릿 docx 생성 완료: {out_path}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 5: 테스트 통과 확인**

Run: `python -m pytest tests/test_render_docx_template.py -q`
Expected: PASS (11 passed — Task 1의 4개 + Task 2의 7개)

- [ ] **Step 6: 커밋**

```bash
git add requirements.txt scripts/render_docx_template.py tests/test_render_docx_template.py
git commit -m "feat: docx 템플릿 렌더러 render_template + CLI + 오류 처리"
```

---

### Task 3: 예시 템플릿 생성기 + `templates/example-template.docx`

사용자가 복사해 쓸 예시 양식을 **코드로 생성**한다(git 재현 가능). 예시는 9개 항목 토큰을 모두 담고, 실행 항목은 `{%tr%}` 표, 목록형 항목은 `{%p%}` 문단 반복으로 구성한다.

**Files:**
- Create: `scripts/make_example_template.py`
- Create: `templates/example-template.docx` (생성 스크립트로 산출)
- Modify: `tests/test_render_docx_template.py` (예시 템플릿 렌더 가능 테스트 추가)

**Interfaces:**
- Consumes: python-docx, Task 2의 `render_template`.
- Produces: `build_example() -> docx.Document`(모든 토큰을 담은 예시 문서), `save_example(path: str) -> None`(파일로 저장). 스크립트 직접 실행 시 `templates/example-template.docx`에 저장.

- [ ] **Step 1: 실패하는 테스트 작성**

`tests/test_render_docx_template.py`에 추가:

```python
def test_example_template_renders_with_sample(tmp_path):
    from make_example_template import build_example
    from render_docx_template import render_template
    tpl = tmp_path / "example.docx"
    build_example().save(str(tpl))
    out = tmp_path / "out.docx"
    render_template(str(tpl), _sample(), str(out))
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    # 제목 치환 확인
    assert any("2026 3분기 제품 로드맵 회의" in t for t in texts)
    # 결정 사항 목록 반복 확인
    assert any("docx 변환은 python-docx로 진행" in t for t in texts)
    # 실행 항목 표에 데이터 행 존재
    table = doc.tables[0]
    assert len(table.rows) >= 3  # 헤더 + 2행 이상
```

- [ ] **Step 2: 테스트 실패 확인**

Run: `python -m pytest tests/test_render_docx_template.py::test_example_template_renders_with_sample -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'make_example_template'`

- [ ] **Step 3: 예시 생성기 구현**

`scripts/make_example_template.py` 생성:

```python
"""예시 템플릿 생성기: templates/example-template.docx를 코드로 재현한다.

바이너리 docx를 불투명하게 커밋하지 않기 위해, 예시 양식을 이 스크립트로
언제든 재생성한다. 사용자는 산출된 파일을 열어 자기 양식으로 편집하거나
토큰만 복사해 쓴다.
"""
from pathlib import Path

from docx import Document

OUT = Path(__file__).resolve().parent.parent / "templates" / "example-template.docx"


def build_example() -> Document:
    """9개 항목 토큰을 모두 담은 예시 템플릿 Document를 만든다."""
    doc = Document()
    # 표시자 위치는 자유롭게 편집 가능. 아래는 안내(브레이스/퍼센트 미포함).
    doc.add_paragraph(
        "아래 표시자 위치를 원하는 대로 편집하세요. "
        "실행 항목 표의 행은 항목 수만큼 자동으로 늘어납니다."
    )

    doc.add_heading("{{ title }}", level=0)
    doc.add_paragraph("일시: {{ date }}")
    doc.add_paragraph("참석자: {{ attendees_joined }}")

    doc.add_heading("회의 목적", level=1)
    doc.add_paragraph("{{ purpose }}")

    doc.add_heading("논의 내용", level=1)
    doc.add_paragraph("{%p for d in discussion %}")
    doc.add_paragraph("{{ d.topic }}")
    doc.add_paragraph("{%p for p in d.points %}")
    doc.add_paragraph("- {{ p }}")
    doc.add_paragraph("{%p endfor %}")
    doc.add_paragraph("{%p endfor %}")

    doc.add_heading("결정 사항", level=1)
    doc.add_paragraph("{%p for x in decisions %}")
    doc.add_paragraph("- {{ x }}")
    doc.add_paragraph("{%p endfor %}")

    doc.add_heading("실행 항목", level=1)
    # 3행 구조: {%tr for%} 행 / 데이터 행 / {%tr endfor%} 행.
    # docxtpl가 for·endfor 행을 삭제하고 사이 데이터 행을 항목 수만큼 반복한다.
    table = doc.add_table(rows=4, cols=3)
    table.style = "Table Grid"
    hdr = table.rows[0].cells
    hdr[0].text, hdr[1].text, hdr[2].text = "할 일", "담당자", "기한"
    table.rows[1].cells[0].text = "{%tr for a in action_items %}"
    data = table.rows[2].cells
    data[0].text = "{{ a.task }}"
    data[1].text = "{{ a.owner }}"
    data[2].text = "{{ a.due }}"
    table.rows[3].cells[0].text = "{%tr endfor %}"

    doc.add_heading("다음 회의", level=1)
    doc.add_paragraph("{{ next_meeting }}")

    doc.add_heading("기타·특이사항", level=1)
    doc.add_paragraph("{%p for n in notes %}")
    doc.add_paragraph("- {{ n }}")
    doc.add_paragraph("{%p endfor %}")
    return doc


def save_example(path: str) -> None:
    build_example().save(path)


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    save_example(str(OUT))
    print(f"예시 템플릿 생성 완료: {OUT}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: 테스트 통과 확인**

Run: `python -m pytest tests/test_render_docx_template.py::test_example_template_renders_with_sample -q`
Expected: PASS (1 passed)

- [ ] **Step 5: 실제 예시 파일 생성**

Run: `python scripts/make_example_template.py`
Expected: `예시 템플릿 생성 완료: ...templates\example-template.docx` 출력, `templates/example-template.docx` 파일 생성.

- [ ] **Step 6: 커밋**

```bash
git add scripts/make_example_template.py templates/example-template.docx tests/test_render_docx_template.py
git commit -m "feat: 예시 템플릿 생성기 + example-template.docx"
```

---

### Task 4: SKILL.md 오케스트레이션 연결 + 전체 회귀

Claude가 [4] 단계에서 템플릿 유무에 따라 렌더러를 고르도록 SKILL.md에 지시를 추가하고, 토큰 치트시트를 문서화한다. 전체 테스트로 회귀를 확인한다.

**Files:**
- Modify: `SKILL.md` ([4] 단계 + 토큰 치트시트)

**Interfaces:**
- Consumes: Task 2의 `render_docx_template.py` CLI, 기존 `render_docx.py` CLI.
- Produces: 코드 없음(문서). 산출물은 사용자 안내 문구.

- [ ] **Step 1: SKILL.md [4] 단계 교체**

`SKILL.md`의 `### [4] docx 생성 (사용자 승인 후)` 섹션을 아래로 교체한다:

````markdown
### [4] docx 생성 (사용자 승인 후)
사용자가 **회의록 양식(.docx 템플릿)**을 제공했는지에 따라 렌더러를 고른다.

- **양식 없음(기본):**
  `python scripts/render_docx.py output/minutes.json output/minutes.docx`
- **양식 있음(.docx 템플릿):**
  `python scripts/render_docx_template.py <template.docx> output/minutes.json output/minutes.docx`
  - 템플릿 파일은 이름만 줘도 공통 위치를 탐색한다(`resolve_input_path` 사용).
  - 사용자가 **hwp/pdf 양식**을 주면, 한글/워드에서 **.docx로 저장(다른 이름으로 저장 → Word)**해 달라고 요청한다. 이 렌더러는 .docx만 받는다.
  - 표시자 문법 오류·.docx 아님·파일 없음 시 명확한 오류가 나므로 그대로 사용자에게 전달한다.

최종 `.docx` 경로를 사용자에게 안내한다.

#### 양식 템플릿 표시자 (스마트 토큰)
사용자 양식에 아래 토큰을 넣으면 해당 자리에 회의록 내용이 채워진다.
복사해 쓸 예시는 `templates/example-template.docx` (재생성:
`python scripts/make_example_template.py`).

| 토큰 | 채워지는 값 |
|---|---|
| `{{ title }}` | 회의 제목 |
| `{{ date }}` | 회의 일시 (없으면 빈칸) |
| `{{ purpose }}` | 회의 목적 |
| `{{ next_meeting }}` | 다음 회의 (없으면 빈칸) |
| `{{ attendees_joined }}` | 참석자 한 줄 결합 "홍길동, 김철수" |
| `{% for a in attendees %}{{ a }}{% endfor %}` | 참석자 목록 반복 |
| `{% for d in discussion %}{{ d.topic }} … {% for p in d.points %}{{ p }}{% endfor %}{% endfor %}` | 논의 주제·포인트 반복 |
| `{% for x in decisions %}{{ x }}{% endfor %}` | 결정 사항 반복 |
| 실행 항목 표 — 아래 "표 행 반복" 참고 (`action_items` 사용) | 실행 항목 표 — 행 자동 반복 |
| `{% for n in notes %}{{ n }}{% endfor %}` | 기타·특이사항 반복 |

**표 행 반복(`{%tr%}`)은 3행 구조로 넣는다** — 한 행에 for와 endfor를 함께 넣으면
동작하지 않는다. 표에 다음 3개 행을 만들고, `{%tr%}`가 든 for·endfor 행은
렌더 시 삭제되며 그 사이 데이터 행이 항목 수만큼 반복된다:

| 할 일 | 담당자 | 기한 |
|---|---|---|
| `{%tr for a in action_items %}` | (빈칸) | (빈칸) |
| `{{ a.task }}` | `{{ a.owner }}` | `{{ a.due }}` |
| `{%tr endfor %}` | (빈칸) | (빈칸) |

문단 단위로 반복시키려면 `{% %}` 대신 `{%p %}`를 쓴다.
없는 값은 빈칸(스칼라)·행 미생성(목록)으로 처리되어 양식이 깔끔하게 유지된다.
복사해 쓸 예시는 `templates/example-template.docx`.
````

- [ ] **Step 2: 전체 테스트 회귀 확인**

Run: `python -m pytest tests -q`
Expected: PASS — 기존 테스트 전부 + 신규 `test_render_docx_template.py`(12개) 모두 통과, 실패 0.

- [ ] **Step 3: 커밋**

```bash
git add SKILL.md
git commit -m "docs: SKILL.md [4]에 템플릿 렌더러 분기 + 토큰 치트시트 추가"
```

---

## Self-Review (작성자 점검 결과)

**1. Spec coverage** — 스펙 각 절과 태스크 대응:
- §3 아키텍처(두 갈래 분기) → Task 2 `render_template`, Task 4 SKILL.md 분기.
- §4 데이터 계약(컨텍스트 매핑, null 규칙) → Task 1 `build_context` + 테스트.
- §4.2 표 행 반복(`{%tr%}`) → Task 2 표 테스트, Task 3 예시.
- §5 디렉터리(신규 파일들) → Task 1~3에서 전부 생성.
- §7 예시 템플릿 + 생성 스크립트 → Task 3.
- §8 에러 처리(미설치/파일없음/.docx아님/문법오류/스키마위반) → Task 2 오류 테스트 + `_load_docxtemplate`/`render_template`, 스키마위반은 기존 `load_minutes` 재사용.
- §9 테스트 전략(치환/표반복/빈값/컨텍스트빌더/오류/회귀) → Task 1~4 테스트로 전부 커버.
- §6 실행 방식(Claude가 분기 실행) → Task 4 SKILL.md.

**2. Placeholder scan** — TBD/TODO/"적절히 처리" 없음. 모든 코드 스텝에 실제 코드 포함.

**3. Type consistency** — `build_context`(Task 1)가 만드는 키(`action_items[].owner/due`가 항상 str)를 Task 2 표 렌더와 Task 3 예시가 동일하게 소비. `render_template(template_path, data, out_path)` 시그니처가 Task 2 정의·Task 3 테스트에서 일치. `build_example()` 반환(Document)이 Task 3 테스트에서 `.save()`로 사용 — 일치.
