# 템플릿 기반 docx 출력 — 설계 문서 (요구사항)

- 작성일: 2026-07-21
- 상태: 작성됨 (사용자 검토 대기)
- 선행 문서: [2026-07-16-meeting-minutes-skill-design.md](./2026-07-16-meeting-minutes-skill-design.md), [2026-07-20-stt-input-adapters-design.md](./2026-07-20-stt-input-adapters-design.md)

## 1. 목표

사용자가 **자신의 회의록 양식(.docx 템플릿)**을 제공하면, 추출된 회의록 내용을 그 양식의 **표시자(placeholder)** 자리에 채워 넣어 **레이아웃·서식을 100% 보존한 .docx**로 출력한다.

기존 파이프라인([1]입력 → [2]추출 → [3]Markdown → [4]docx)은 **한 줄도 수정하지 않는다.** [4] docx 출력 단계에만 "템플릿 렌더러"라는 두 번째 갈래를 추가한다.

## 2. 확정된 설계 결정

| 항목 | 결정 | 이유 |
|---|---|---|
| 입력 양식 포맷 | **.docx 템플릿만** 지원 | .docx일 때만 레이아웃 100% 보존이 안정적. hwp/pdf는 원리적으로 docx 재현이 어긋남 |
| hwp/pdf 처리 | **사용자가 직접 docx로 변환**해서 제공 | 자동 변환기(한글/LibreOffice) 의존을 피해 단순·안정 유지 |
| 표시자 방식 | **스마트 토큰 (docxtpl / Jinja2)** | 단순 치환(`{{title}}`)의 상위 호환. 표 행 자동 반복·목록 반복 지원 |
| 토큰 이름 | **영어(ASCII)** + 복붙용 예시 템플릿·한글 설명서 | Jinja2의 비ASCII 식별자 지원 불확실. 예시 제공으로 사용 편의 보완 |
| 기존 파이프라인 | **불변 유지** (Markdown 검토 관문 포함) | 사용자 지시: "원래 만든 건 그대로 둔다" |
| 렌더러 선택 | 템플릿 있으면 템플릿 렌더러, 없으면 **기존 `render_docx.py`** | 신규 파일 1개 추가, 기존 렌더러 무수정 |
| 데이터 계약 | 기존 `minutes.json` (9개 항목 스키마) **그대로** | 단일 진실 원천 불변. 추출 계층 재사용 |
| 실행 주체 | 사용자는 `/meeting-minutes` 호출, **Claude가 SKILL.md 지시로 실행** | 기존 Harness 패턴 유지 |

## 3. 아키텍처

### 3.1 전체 파이프라인 ([4] 출력 단계만 확장)

```
[1] 입력 → [2] 추출 → [3] Markdown 렌더·검토
                              ↓ minutes.json (단일 진실 원천, schema 검증)
                        [4] docx 렌더  ← 이번 작업: 두 갈래로 분기
                         ├─ 템플릿 있음 → render_docx_template.py (신규)
                         └─ 템플릿 없음 → render_docx.py (기존, 불변)
```

의존 방향은 `[1] → [2] → [3] → [4]` 단방향 유지. `minutes.json`은 두 렌더러가 각자 소비하는 형제 관계이며, docx는 Markdown 텍스트가 아니라 **JSON에서 직접** 생성된다(기존과 동일). Markdown 단계는 사람이 검토하는 관문으로 그대로 남는다.

### 3.2 템플릿 렌더러 내부

```
render_docx_template.py <template.docx> <minutes.json> <out.docx>
        │
        ▼
  load_minutes(json)          # 기존 validate.py 재사용 → schema 검증된 dict
        │
        ▼
  build_context(data)         # minutes.json → Jinja2 컨텍스트 (토큰 매핑)
        │
        ▼
  DocxTemplate(template).render(context)   # docxtpl: 표시자 치환 + 표 행 반복
        │
        ▼
  .save(out.docx)             # 레이아웃 100% 보존된 최종 docx
```

## 4. 데이터 계약 — 표시자 ↔ 컨텍스트

`minutes.json`의 9개 항목을 다음 컨텍스트로 매핑한다. 단순 치환용 편의 문자열과, 반복용 구조 리스트를 **둘 다** 제공한다.

### 4.1 단순 치환 (스칼라)

| 템플릿 토큰 | 채워지는 값 | 빈 값일 때 |
|---|---|---|
| `{{ title }}` | 회의 제목 | (항상 존재) |
| `{{ date }}` | 회의 일시 | `""` (빈칸) |
| `{{ purpose }}` | 회의 목적 | `""` |
| `{{ next_meeting }}` | 다음 회의 | `""` |
| `{{ attendees_joined }}` | `"홍길동, 김철수"` (쉼표 결합) | `""` |

### 4.2 반복 (리스트/표)

| 템플릿 구문 | 채워지는 값 |
|---|---|
| `{% for a in attendees %}{{ a }}{% endfor %}` | 참석자 각각 (문자열) |
| `{% for d in discussion %}{{ d.topic }} … {% for p in d.points %}{{ p }}{% endfor %}{% endfor %}` | 논의 주제별·포인트별 |
| `{% for x in decisions %}{{ x }}{% endfor %}` | 결정 사항 각각 |
| `{%tr for a in action_items %}{{ a.task }} · {{ a.owner }} · {{ a.due }}{%tr endfor %}` | **실행 항목 표 — 행 자동 반복** |
| `{% for n in notes %}{{ n }}{% endfor %}` | 기타·특이사항 각각 |

- `action_items`의 각 원소는 `{task, owner, due}` 3필드. `owner`/`due`가 `null`이면 컨텍스트에서 `"-"`로 치환해 넘긴다.
- `{%tr ... %}`는 docxtpl의 "표 행 반복" 전용 태그로, 표의 한 행에 넣으면 항목 개수만큼 행이 자동 증식한다.
- **빈 값 원칙**: 없는 스칼라는 빈 문자열, 없는 리스트는 반복 0회(행 미생성) → 양식이 깔끔하게 유지된다.

### 4.3 컨텍스트 빌더 (`build_context`)

```python
def build_context(data: dict) -> dict:
    """minutes.json(dict) → docxtpl 렌더 컨텍스트."""
    return {
        "title": data["title"],
        "date": data.get("date") or "",
        "purpose": data.get("purpose") or "",
        "next_meeting": data.get("next_meeting") or "",
        "attendees": data["attendees"],
        "attendees_joined": ", ".join(data["attendees"]),
        "discussion": data["discussion"],       # [{topic, points}]
        "decisions": data["decisions"],
        "action_items": [
            {"task": a["task"], "owner": a["owner"] or "-", "due": a["due"] or "-"}
            for a in data["action_items"]
        ],
        "notes": data["notes"],
    }
```

## 5. 디렉터리 구조

```
.claude/skills/meeting-minutes/
├─ SKILL.md                          # [4] 단계에 템플릿 옵션 안내 추가 (수정)
├─ requirements.txt                  # docxtpl 추가 (수정)
├─ schema/minutes.schema.json        # 데이터 계약 (불변)
├─ scripts/
│  ├─ render_docx.py                 # 기존 기본 렌더러 (불변)
│  ├─ render_docx_template.py        # ★신규: 템플릿 렌더러
│  ├─ make_example_template.py       # ★신규: 예시 템플릿(.docx)을 코드로 생성
│  ├─ validate.py                    # load_minutes 재사용 (불변)
│  └─ …                              # 나머지 (불변)
├─ templates/
│  └─ example-template.docx          # ★신규: 사용자가 복사해 쓰는 예시 양식
├─ tests/
│  └─ test_render_docx_template.py   # ★신규: 토큰 치환·표 반복·빈 값·오류
└─ examples/
   └─ sample_minutes.json            # 기존 샘플 (렌더 회귀용)
```

★ = 이번 작업 추가/수정. 표시 없는 항목은 불변.

## 6. 실행 방식 — 기존 Harness 패턴 유지

사용자 진입점은 `/meeting-minutes` 하나. Claude가 SKILL.md [4] 단계 지시대로 분기한다.

```bash
# 템플릿을 제공한 경우
python scripts/render_docx_template.py <template.docx> output/minutes.json output/minutes.docx

# 템플릿이 없는 경우 (기존 그대로)
python scripts/render_docx.py output/minutes.json output/minutes.docx
```

- SKILL.md [4] 지시문: "사용자가 양식 .docx를 제공했으면 템플릿 렌더러를, 아니면 기본 렌더러를 실행한다. 사용자가 hwp/pdf 양식을 주면 docx로 변환해 달라고 요청한다."
- 템플릿 파일 위치는 기존 `resolve_input_path`(파일 이름만으로 공통 위치 탐색)를 재사용해 해석한다.

## 7. 예시 템플릿 (`make_example_template.py` + `templates/example-template.docx`)

- 바이너리 .docx를 git에 불투명하게 넣지 않기 위해, **예시 템플릿을 코드로 생성**하는 스크립트를 둔다(재현 가능·버전관리 용이).
- 생성되는 예시에는 9개 항목 토큰이 모두 들어가며, 실행 항목은 `{%tr …%}`로 자동 반복되는 표로 구성한다. 사용자는 이 파일을 열어 자기 양식으로 편집하거나 토큰만 복사해 쓴다.
- 예시 상단에 "이 토큰들을 원하는 위치에 배치하세요"라는 한글 안내와 토큰 치트시트를 포함한다.

## 8. 에러 처리

| 상황 | 처리 |
|---|---|
| `docxtpl` 미설치 | `pip install -r requirements.txt` 안내 후 오류 |
| 템플릿 파일 없음 | `FileNotFoundError` (검색 위치 안내) |
| 템플릿이 .docx 아님 (hwp/pdf 등) | "한글/워드에서 .docx로 저장해 다시 주세요" 안내 후 오류 |
| 템플릿의 Jinja 문법 오류 | docxtpl 예외를 잡아 어느 토큰이 잘못됐는지 메시지로 전달 |
| `minutes.json` 스키마 위반 | 기존 `load_minutes`의 `ValidationError` 그대로 |
| 템플릿에 없는 토큰 | 무시 (해당 자리 비움) — 오류 아님 |

## 9. 테스트 전략 (`test_render_docx_template.py`)

실제 사용자 양식 없이, **테스트 안에서 python-docx로 임시 템플릿을 만들어** 결정적으로 검증한다.

- **단순 치환**: `{{ title }}`·`{{ date }}` 등이 값으로 바뀌는지, 렌더 후 텍스트를 읽어 확인.
- **표 행 반복**: `{%tr …%}` 표에 `action_items` 개수만큼 행이 생기는지.
- **빈 값**: `date`/`purpose`/`next_meeting`이 `null`이면 빈칸, `owner`/`due` null이면 `"-"`, 빈 리스트면 행 미생성.
- **컨텍스트 빌더**: `build_context`가 스키마 dict를 올바른 컨텍스트로 변환하는지(순수 함수 단위 테스트).
- **오류**: 템플릿 파일 없음, .docx 아님, 잘못된 토큰 시 명확한 예외.
- **회귀**: 기존 `render_docx.py`/`render_markdown.py` 테스트는 그대로 통과.

## 10. 범위 밖 (Non-Goals)

- hwp/pdf 템플릿의 **자동 변환** 및 직접 파싱 (사용자가 docx로 변환해 제공).
- 템플릿으로부터 **레이아웃을 새로 그리는 근사 재구성** (스마트 토큰으로 원본 보존만).
- 기존 파이프라인([1]~[3], 기본 docx 렌더러) 변경.
- 추출 항목 9종·스키마 변경.
- docx 외 출력 포맷(PDF 등) 추가.
- 한글(비ASCII) 토큰 이름 지원 — 필요 시 구현 단계에서 Jinja2 지원 여부 검증 후 별도 결정.
