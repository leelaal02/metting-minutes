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


def test_missing_purpose_rejected():
    bad = json.loads(SAMPLE.read_text(encoding="utf-8"))
    del bad["purpose"]
    with pytest.raises(ValidationError):
        validate_minutes(bad)


def test_missing_notes_rejected():
    bad = json.loads(SAMPLE.read_text(encoding="utf-8"))
    del bad["notes"]
    with pytest.raises(ValidationError):
        validate_minutes(bad)


def test_notes_must_be_array():
    bad = json.loads(SAMPLE.read_text(encoding="utf-8"))
    bad["notes"] = "특이사항 없음"  # 배열이어야 함
    with pytest.raises(ValidationError):
        validate_minutes(bad)
