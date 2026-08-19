"""[3]·[4] 출력 파일명 생성: minutes.json → '회의록_<제목>_<생성일>' 안전 파일명.

산출물(md·docx) 이름을 회의 제목과 **생성한 날짜**로 지어 여러 회의록을 구분하기
쉽게 한다. 날짜는 회의 일시가 아니라 파일을 만든 날(오늘)을 쓴다.
파일명 금지문자를 제거하고 공백을 하이픈으로 바꾼 결정적 로직이라 스크립트에 둔다
(매핑·분석 같은 비결정 판단은 하지 않는다).
"""
import re
import sys
from datetime import date
from pathlib import Path

from validate import load_minutes

# Windows/유닉스 공통 파일명 금지문자.
_ILLEGAL = r'[\\/:*?"<>|]'
# 파일명이 간결하도록 제목은 앞부분 몇 단어만 쓴다(이 길이까지 온전한 단어 단위로).
_MAX_TITLE_LEN = 12


def _slug(text: str, max_len: int = _MAX_TITLE_LEN) -> str:
    """파일명에 안전한 짧은 조각으로 변환.

    금지문자를 제거하고, 제목 앞에서부터 **온전한 단어 단위로** 최대 길이까지만
    공백으로 이어 붙인다 — 단어 사이 공백은 그대로 두어(하이픈으로 바꾸지 않아
    이름에 하이픈이 남발되지 않음) 자연스럽고, 긴 제목도 앞부분 2~3단어로
    간결해진다. 첫 단어가 이미 최대 길이를 넘으면 그 단어를 잘라서 쓴다.
    """
    text = re.sub(_ILLEGAL, "", text or "").strip()
    words = text.split()
    if not words:
        return ""
    stem = ""
    for w in words:
        candidate = w if not stem else f"{stem} {w}"
        if len(candidate) > max_len:
            break
        stem = candidate
    if not stem:  # 첫 단어가 이미 max_len보다 길면 잘라서라도 사용
        stem = words[0][:max_len]
    return stem.strip()


def _form_prefix(form_name: str) -> str:
    """양식 파일명에서 구분용 접두어를 뽑는다.

    같은 회의를 여러 양식에 채울 때 파일명이 겹치지 않도록 양식의 이름을 앞에
    붙이기 위한 것. '회의록_KISA.docx' → 'KISA', '회의록_누리미디어.docx' →
    '누리미디어'처럼 앞에 중복되는 '회의록' 표기를 떼고 금지문자를 제거한다.
    뗄 것이 없으면(양식명이 '회의록'뿐) stem을 그대로 쓴다.
    """
    stem = Path(form_name).stem  # 확장자 제거
    trimmed = re.sub(r"^회의록[\s_-]*", "", stem).strip()
    stem = trimmed or stem       # '회의록'뿐이면 원래 stem 유지
    return re.sub(_ILLEGAL, "", stem).strip()


def build_output_stem(data: dict, today: date = None, form_name: str = None) -> str:
    """minutes 데이터 → '[<양식명>_]회의록_<제목>_<생성일 YYYY-MM-DD>' 파일명 stem.

    - 제목: 앞부분 몇 단어만 공백 그대로 이어 간결하게(금지문자 제거, 하이픈 변환 안 함).
    - 날짜: 회의 일시가 아니라 이 파일을 만든 날(today, 기본값은 오늘).
    - form_name: 양식(.docx)을 쓰면 그 파일명에서 뽑은 접두어를 **맨 앞**에 붙여
      같은 회의를 여러 양식에 채워도 파일명이 겹치지 않게 한다. 없으면 접두어 없음.
    """
    if today is None:
        today = date.today()
    title = _slug(data.get("title") or "회의록")
    stem = f"회의록_{title}" if title else "회의록"
    if form_name:
        prefix = _form_prefix(form_name)
        if prefix:
            stem = f"{prefix}_{stem}"
    return f"{stem}_{today.isoformat()}"


def disambiguate_stem(stem: str, ext: str, out_dir) -> str:
    """out_dir에 `stem+ext`가 이미 있으면 '_2','_3'…을 붙여 겹치지 않는 stem을 반환.

    같은 날 제목 앞 단어가 겹치는 회의(예: "주간 회의" 두 건)가 서로를 조용히
    덮어쓰지 않게 한다. 충돌이 없으면 stem을 그대로 돌려주므로 기존 동작과 동일.
    """
    out_dir = Path(out_dir)
    if not (out_dir / f"{stem}{ext}").exists():
        return stem
    i = 2
    while (out_dir / f"{stem}_{i}{ext}").exists():
        i += 1
    return f"{stem}_{i}"


def _default_out_dir(json_path: str) -> Path:
    """최종본이 놓일 디렉토리 추정. minutes.json은 output/.work/에 있으므로
    그 부모(output/)를 쓰고, .work가 아니면 json이 있는 디렉토리를 쓴다."""
    parent = Path(json_path).parent
    return parent.parent if parent.name == ".work" else parent


def main() -> None:
    data = load_minutes(sys.argv[1])
    ext = sys.argv[2] if len(sys.argv) > 2 else ".docx"
    if not ext.startswith("."):
        ext = "." + ext
    form = sys.argv[3] if len(sys.argv) > 3 else None
    out_dir = Path(sys.argv[4]) if len(sys.argv) > 4 else _default_out_dir(sys.argv[1])
    stem = build_output_stem(data, form_name=form)
    print(disambiguate_stem(stem, ext, out_dir) + ext)


if __name__ == "__main__":
    main()
