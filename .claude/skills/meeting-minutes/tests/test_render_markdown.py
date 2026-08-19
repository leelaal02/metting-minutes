import json
from pathlib import Path

from render_markdown import render_markdown

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"


def _sample():
    return json.loads(SAMPLE.read_text(encoding="utf-8"))


def test_title_and_sections_present():
    md = render_markdown(_sample())
    assert md.startswith("# 2026 3분기 제품 로드맵 회의")
    for heading in ["## 참석자", "## 회의 목적", "## 논의 내용", "## 결정 사항",
                    "## 실행 항목 Action Items", "## 다음 회의 일정", "## 기타·특이사항"]:
        assert heading in md


def test_section_order():
    md = render_markdown(_sample())
    # 회의 목적은 참석자와 논의 내용 사이, 기타·특이사항은 맨 끝(다음 회의 뒤)
    assert md.index("## 참석자") < md.index("## 회의 목적") < md.index("## 논의 내용")
    assert md.index("## 다음 회의 일정") < md.index("## 기타·특이사항")


def test_open_issues_section_sits_between_decisions_and_actions():
    """미결 사항은 결정 사항 바로 뒤 — "정해진 것 / 안 정해진 것"이 붙어 읽혀야 한다."""
    md = render_markdown(_sample())
    assert "## 미결 사항" in md
    assert md.index("## 결정 사항") < md.index("## 미결 사항") < md.index("## 실행 항목")
    assert "- 실시간 STT 도입 시점 미정 (4분기 재논의)" in md


def test_empty_open_issues_renders_none_marker():
    """미결이 없는 회의도 항목 자체는 남긴다 — 빠뜨린 것과 없는 것을 구분."""
    data = _sample()
    data["open_issues"] = []
    md = render_markdown(data)
    assert "## 미결 사항\n- (없음)" in md


def test_purpose_and_notes_content():
    md = render_markdown(_sample())
    assert "3분기 제품 로드맵과 STT 연동 범위 확정" in md   # purpose
    assert "- STT 엔진 후보 벤치마크는 다음 스프린트에 진행" in md  # notes 항목


def test_action_items_table():
    md = render_markdown(_sample())
    assert "| 할 일 | 담당자 | 기한 |" in md
    assert "| python-docx 렌더러 PoC 작성 | 이정우 | 2026-07-23 |" in md


def test_null_due_rendered_as_dash():
    md = render_markdown(_sample())
    assert "| 샘플 회의 원문 수집 | 박서연 | - |" in md


def test_table_cell_escapes_pipe_and_newline():
    """할 일/담당자/기한에 든 `|`·개행이 표를 깨뜨리지 않도록 이스케이프된다."""
    data = _sample()
    data["action_items"] = [
        {"task": "A안 | B안 비교", "owner": "김수민\n(대행 이정우)", "due": "-"},
    ]
    md = render_markdown(data)
    # 파이프는 \| 로 이스케이프, 개행은 공백으로 → 데이터 행이 정확히 3열을 유지
    row = next(
        line for line in md.splitlines()
        if line.startswith("|") and "A안" in line
    )
    assert row == r"| A안 \| B안 비교 | 김수민 (대행 이정우) | - |"
    # 실제 셀(열) 개수 검증: 양끝 빈 항목 제외 3개, 이스케이프된 \| 는 열로 세지 않음
    import re
    cells = [c for c in re.split(r"(?<!\\)\|", row)[1:-1]]
    assert len(cells) == 3


def test_empty_and_null_fields():
    data = _sample()
    data["attendees"] = []
    data["action_items"] = []
    data["next_meeting"] = None
    data["purpose"] = None
    data["notes"] = []
    md = render_markdown(data)
    assert "- (없음)" in md          # 빈 참석자/빈 notes
    assert "(미정)" in md            # next_meeting None
    # 회의 목적이 None이면 (없음)으로 렌더
    assert "## 회의 목적\n(없음)" in md
