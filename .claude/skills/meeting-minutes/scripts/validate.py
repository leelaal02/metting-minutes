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
