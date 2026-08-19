"""[4-대체] 템플릿 렌더러: .docx 양식 + minutes.json → 표시자 치환 docx."""
import sys

from normalize_input import resolve_input_path
from validate import load_minutes

_TODO_TEXT = "입력필요"
_TODO_COLOR = "FF0000"


def _todo():
    """미입력 표시용 "입력필요"(빨강·굵게) RichText."""
    from docxtpl import RichText

    rt = RichText()
    rt.add(_TODO_TEXT, color=_TODO_COLOR, bold=True)
    return rt


def _rt(value):
    """값이 있으면 평문 RichText, 없으면 "입력필요"(빨강·굵게).

    자동매핑의 `{{r ... }}` 슬롯은 항상 RichText여야 하므로(평문 str을 주면
    docxtpl가 빈칸으로 렌더) 값 유무와 무관하게 RichText로 감싼다.
    """
    from docxtpl import RichText

    if value not in (None, "", []):
        rt = RichText()
        rt.add(str(value))
        return rt
    return _todo()


def _topic_rt(index: int, topic: str):
    """"N. 주제"를 굵게 표시하는 소제목 RichText(넘버링+굵게)."""
    from docxtpl import RichText

    rt = RichText()
    rt.add(f"{index}. {topic}", bold=True)
    return rt


def _topic_plain_rt(topic: str):
    """번호 없이 주제만 굵게 — 양식이 이미 번호를 써서 기호가 겹칠 때 쓴다.

    이때 계층 기호(□ 등)는 `apply_form_mapping`이 토큰 앞에 평문으로 붙인다.
    """
    from docxtpl import RichText

    rt = RichText()
    rt.add(topic, bold=True)
    return rt


def build_context(data: dict) -> dict:
    """minutes.json(dict) → docxtpl 렌더 컨텍스트.

    두 계열의 키를 함께 담는다(superset):
    - 평문 키(title/date/…): 사용자 작성 토큰(`{{ date }}`) 경로. 하위호환 유지.
    - RichText 키(`todo`·`*_rt`·`discussion_rt`·`action_items_rt`): 자동매핑의
      `{{r ... }}` 경로. 빈 값은 "입력필요"(빨강·굵게), 소제목은 넘버링+굵게.
    null 스칼라는 빈 문자열로, action_items의 owner/due null은 "-"로 정규화한다.
    """
    attendees_joined = ", ".join(data["attendees"])
    return {
        # --- 평문 키: 사용자 {{ }} 토큰 경로 (무변경, 하위호환) ---
        "title": data["title"],
        "date": data.get("date") or "",
        "purpose": data.get("purpose") or "",
        "next_meeting": data.get("next_meeting") or "",
        "attendees": data["attendees"],
        "attendees_joined": attendees_joined,
        # 참석인원 수는 세지 않고 attendees 길이에서 얻는다 — 명단과 어긋날 수 없다.
        "attendee_count": len(data["attendees"]),
        "discussion": data["discussion"],
        "decisions": data["decisions"],
        "open_issues": data["open_issues"],
        "action_items": [
            {"task": a["task"], "owner": a["owner"] or "-", "due": a["due"] or "-"}
            for a in data["action_items"]
        ],
        "notes": data["notes"],
        # --- RichText 키: 자동매핑 {{r }} 경로 (신설) ---
        "todo": _todo(),
        "title_rt": _rt(data["title"]),
        "date_rt": _rt(data.get("date")),
        "purpose_rt": _rt(data.get("purpose")),
        "next_meeting_rt": _rt(data.get("next_meeting")),
        "attendees_rt": _rt(attendees_joined or None),
        # 원본 불변: 새 리스트를 만든다(data mutation 금지).
        "discussion_rt": [
            {
                "topic_rt": _topic_rt(i, d["topic"]),
                "topic_plain_rt": _topic_plain_rt(d["topic"]),
                "points": d["points"],
            }
            for i, d in enumerate(data["discussion"], 1)
        ],
        "action_items_rt": [
            {"task": a["task"], "owner_rt": _rt(a["owner"]), "due_rt": _rt(a["due"])}
            for a in data["action_items"]
        ],
    }


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
    from jinja2 import TemplateError
    tpl = DocxTemplate(str(resolved))
    try:
        # autoescape=True: 값에 든 XML 특수문자(&, <, >)를 이스케이프해 보존한다.
        # (기본값 False면 "R&D"·"<A>" 같은 값이 렌더 시 잘리거나 사라진다.)
        tpl.render(build_context(data), autoescape=True)
    except TemplateError as e:
        raise ValueError(
            f"템플릿의 표시자 문법에 오류가 있습니다: {e}. "
            "{{ }}·{% %}·{%tr %}·{%p %} 토큰을 치트시트와 비교해 확인하세요."
        ) from e
    # 양식 표의 행 나눔 금지를 풀어 긴 내용이 페이지를 자연스럽게 넘어가게 한다.
    from docx_postprocess import allow_rows_to_break, inherit_mark_fonts
    allow_rows_to_break(tpl.docx)
    # 렌더로 새로 생긴 런(RichText)에 양식 글꼴을 입힌다.
    inherit_mark_fonts(tpl.docx)
    tpl.save(out_path)


def main() -> None:
    template_path, json_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]
    data = load_minutes(json_path)
    render_template(template_path, data, out_path)
    print(f"템플릿 docx 생성 완료: {out_path}")


if __name__ == "__main__":
    main()
