"""output_naming.py 테스트: '회의록_<제목>_<생성일>' stem 생성 규칙."""
import json
import re
from datetime import date
from pathlib import Path

from output_naming import build_output_stem, disambiguate_stem

SAMPLE = Path(__file__).resolve().parent.parent / "examples" / "sample_minutes.json"
_TODAY = date(2026, 7, 22)


def _sample():
    return json.loads(SAMPLE.read_text(encoding="utf-8"))


def test_sample_stem_uses_generation_date():
    # 제목 앞 몇 단어만 공백 그대로 이어 간결하게(12자 이내 온전한 단어)
    stem = build_output_stem(_sample(), today=_TODAY)
    assert stem == "회의록_2026 3분기 제품_2026-07-22"


def test_spaces_preserved_not_hyphenated():
    # 공백을 하이픈으로 바꾸지 않고 그대로 둔다(짧으면 전체 유지)
    assert build_output_stem({"title": "주간 팀 회의"}, today=_TODAY) \
        == "회의록_주간 팀 회의_2026-07-22"


def test_long_multiword_title_keeps_leading_words():
    # 긴 제목은 앞부분 온전한 단어만 남긴다(하이픈 남발 없음)
    stem = build_output_stem(
        {"title": "주간보고 검토 및 공통 모듈 방향 협의 회의"}, today=_TODAY)
    assert stem == "회의록_주간보고 검토 및 공통_2026-07-22"


def test_ignores_meeting_date():
    # 회의 일시가 있어도 파일명 날짜는 생성일을 쓴다
    stem = build_output_stem({"title": "킥오프", "date": "2099-01-01 14:00"}, today=_TODAY)
    assert stem == "회의록_킥오프_2026-07-22"


def test_null_date_still_appends_generation_date():
    assert build_output_stem({"title": "킥오프", "date": None}, today=_TODAY) \
        == "회의록_킥오프_2026-07-22"


def test_illegal_chars_stripped():
    # 파일명 금지문자(/ : * ? " < > | \)는 제거된다
    stem = build_output_stem({"title": 'Q3/Q4 계획: "로드맵"?'}, today=_TODAY)
    for ch in '\\/:*?"<>|':
        assert ch not in stem
    assert stem.startswith("회의록_")
    assert stem.endswith("_2026-07-22")


def test_long_title_truncated():
    stem = build_output_stem({"title": "가" * 100}, today=_TODAY)
    title_part = stem[len("회의록_"):-len("_2026-07-22")]
    assert len(title_part) <= 12


def test_form_name_prepended_stripping_회의록():
    # 양식명을 맨 앞에, 중복되는 '회의록_'은 떼고 붙인다
    stem = build_output_stem(
        {"title": "주간 팀 회의"}, today=_TODAY, form_name="회의록_KISA.docx")
    assert stem == "KISA_회의록_주간 팀 회의_2026-07-22"
    stem2 = build_output_stem(
        {"title": "주간 팀 회의"}, today=_TODAY,
        form_name=r"c:\Users\u\Desktop\양식\회의록_누리미디어.docx")
    assert stem2 == "누리미디어_회의록_주간 팀 회의_2026-07-22"


def test_form_name_without_회의록_prefix_uses_full_stem():
    # 양식명이 '회의록'으로 시작하지 않으면 stem 전체를 접두어로
    stem = build_output_stem(
        {"title": "킥오프"}, today=_TODAY, form_name="team_form.docx")
    assert stem == "team_form_회의록_킥오프_2026-07-22"


def test_no_form_name_keeps_plain_stem():
    # form_name 미지정 시 접두어 없이 기존과 동일
    assert build_output_stem({"title": "킥오프"}, today=_TODAY) \
        == "회의록_킥오프_2026-07-22"


def test_default_today_is_used():
    # today 미지정 시 오늘 날짜가 붙는다(형식만 확인)
    stem = build_output_stem({"title": "테스트"})
    assert re.match(r"회의록_테스트_\d{4}-\d{2}-\d{2}$", stem)


# --- disambiguate_stem: 같은 날 같은 제목 덮어쓰기 방지 ---

def test_disambiguate_returns_stem_when_no_collision(tmp_path):
    assert disambiguate_stem("회의록_주간 회의_2026-07-22", ".docx", tmp_path) \
        == "회의록_주간 회의_2026-07-22"


def test_disambiguate_appends_suffix_on_collision(tmp_path):
    stem = "회의록_주간 회의_2026-07-22"
    (tmp_path / f"{stem}.docx").write_bytes(b"first")
    # 첫 충돌 → _2
    assert disambiguate_stem(stem, ".docx", tmp_path) == f"{stem}_2"
    (tmp_path / f"{stem}_2.docx").write_bytes(b"second")
    # _2도 있으면 _3
    assert disambiguate_stem(stem, ".docx", tmp_path) == f"{stem}_3"


def test_disambiguate_is_ext_specific(tmp_path):
    stem = "회의록_킥오프_2026-07-22"
    (tmp_path / f"{stem}.md").write_bytes(b"md")
    # .md만 있으면 .docx는 충돌 아님
    assert disambiguate_stem(stem, ".docx", tmp_path) == stem
