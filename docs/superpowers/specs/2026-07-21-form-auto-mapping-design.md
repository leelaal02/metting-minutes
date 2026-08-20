# 양식 항목 자동 매핑 보강 — 설계 문서 (요구사항)

- 작성일: 2026-07-21
- 갱신일: 2026-07-22 — v1 양식 범위를 **표 + 문단 양식**으로 확장(사용자 승인). 문단 양식 지원 반영.
- 상태: 구현됨 (inspect_template.py·apply_form_mapping.py·output_naming.py·테스트·SKILL.md [4] 분기)
- 선행 문서: [2026-07-21-docx-template-output-design.md](./2026-07-21-docx-template-output-design.md)

## 1. 목표

사용자가 **토큰이 없는 빈 서식(.docx)**을 제공해도, 그 양식의 라벨을 자동 인식해 **고정 9항목을 알맞은 자리에 자동 매핑·삽입**하고 채워서 출력한다. **표 기반 양식(칸 채움)과 문단 기반 양식(본문 문단 채움)을 모두 지원**하며, 둘이 섞인 양식도 처리한다. (지금까지는 사람이 양식의 각 자리에 토큰을 수동 삽입해야 했다 — 예: KISA 표 양식, 공문형 문단 양식.)

**9항목 추출 규칙은 그대로 유지한다.** 양식은 여전히 추출을 바꾸지 못하며, 이 기능은 "9항목을 어느 칸에 넣을지"를 자동화하는 **[4] 출력 단계의 보강**이다.

## 2. 확정된 설계 결정

| 항목 | 결정 | 이유 |
|---|---|---|
| 매핑 주체 | **Claude가 의미 판단**(칸 라벨 → 9항목), 스크립트는 기계적 삽입 | 추출[2]과 동일한 Harness 방식. 어떤 양식이든 대응 |
| 비결정/결정 분리 | Claude 매핑은 스크립트 **밖**, 스크립트는 순수 기계적 | 스크립트 전체가 단위 테스트 가능, 매핑 오류가 코드에 안 스며듦 |
| 토큰 문법 | Claude는 문법을 안 짜고, 스크립트가 **검증된 토큰 블록**만 사용 | `{%p%}`/`{%tr%}` 문법 오류 원천 차단 |
| 불일치 처리 | **데이터 보존 우선(자동)** — 자리 없는 9항목은 가장 어울리는 칸에 몰아넣어 버리지 않음 | 사용자 선택. 정보 손실 방지 |
| 승인 게이트 | 없음(자동) + **비차단 매핑 요약 보고** | 자동 진행하되 오매핑을 사용자가 눈으로 잡을 안전망 |
| 양식 범위 | **표 + 문단 양식(v1)** | 공식 양식이 표(KISA)뿐 아니라 공문형 문단 구성도 흔해 둘 다 필요(2026-07-22 확장) |
| 중간 산출물 | 토큰 삽입본 `*_tokenized.docx` **보존** | 다음부터 매핑 없이 바로 재사용 가능(부가 이득) |
| 기존 파이프라인 | **불변** (9항목 스키마, [2]추출, render_markdown, render_docx, render_docx_template) | CLAUDE.md 하드 규칙 준수, 기존 렌더러 재사용 |

## 3. 아키텍처

### 3.1 [4] 출력 단계의 세 갈래

```
[4] docx 출력
 ├─ 양식 없음                       → render_docx.py (기존 기본 서식)
 ├─ 양식 있고 토큰 있음              → render_docx_template.py (기존)
 └─ 양식 있고 토큰 없음(표·문단·혼합) → ★자동 매핑 경로 (신규)
```

### 3.2 자동 매핑 경로 (4스텝)

```
① inspect_template.py <template.docx>
      → 양식 구조(표·칸 라벨·좌표·빈칸·병합, 토큰 유무) JSON 출력
② Claude가 구조 + 9항목을 보고 mapping.json 작성 (의미 판단)
③ apply_form_mapping.py <template.docx> <mapping.json> <out_tokenized.docx>
      → 매핑대로 양식 복사본에 검증된 토큰 블록 삽입
④ render_docx_template.py <out_tokenized.docx> <minutes.json> <final.docx>
      → (기존, 무수정) 토큰 채워 최종 docx
```

의존 방향 단방향. ③까지 끝나면 **기존 렌더러를 그대로 재사용**하므로 render_docx_template는 수정하지 않는다. 원본 양식은 절대 수정하지 않는다(복사본에만 삽입).

## 4. 데이터 계약

### 4.1 inspect_template.py 출력 (Claude 입력)

표 구조(`tables`)와 본문 문단(`paragraphs`)을 함께 JSON으로 출력한다. 병합 셀은 원점(origin) 1회만 출력한다.

```json
{
  "has_tokens": false,
  "tables": [
    {
      "index": 0, "rows": 11, "cols": 2,
      "cells": [
        {"row": 0, "col": 0, "text": "회 의 록", "is_empty": false, "merged": true},
        {"row": 1, "col": 0, "text": "제 목 :", "is_empty": false, "merged": true},
        {"row": 2, "col": 0, "text": "일 시 :", "is_empty": false, "merged": false},
        {"row": 2, "col": 1, "text": "참 가 자", "is_empty": false, "merged": false},
        {"row": 6, "col": 0, "text": "", "is_empty": true, "merged": true}
      ]
    }
  ],
  "paragraphs": [
    {"index": 0, "text": "회의 결과", "is_empty": false},
    {"index": 1, "text": "ㅇ (목적)", "is_empty": false},
    {"index": 2, "text": "", "is_empty": true}
  ]
}
```

- `has_tokens`: 문서 전체에 `{{`/`{%` 가 하나라도 있으면 true(있으면 자동 매핑 경로를 타지 않음). 탐지는 본문뿐 아니라 **머리말/꼬리말·셀 안 중첩표까지 재귀로 훑는다** — 한 곳이라도 놓치면 이미 토큰이 있는 양식이 자동 매핑으로 흘러가 토큰이 이중 삽입되기 때문.
- `is_empty`: 셀/문단 텍스트가 공백뿐이면 true(값이 들어갈 빈 자리 후보).
- `merged`: 가로/세로 병합 영역의 원점 셀이면 true. 병합 중복 셀은 출력에서 제외한다.
- `paragraphs[].index`는 `doc.paragraphs` 상의 위치이며 mapping.json의 `para` 주소와 일치한다. 표 안의 문단은 제외된다(표는 `tables`로 다룬다). 표 양식이면 `paragraphs`는 참고용이고 매핑은 `tables`를 쓰며, 문단 양식이면 그 반대다.

### 4.2 mapping.json (Claude 출력 → apply 입력)

표 칸 채움 `fills`와 본문 문단 채움 `paragraphs`를 가진다. 표 양식이면 `fills`만, 문단 양식이면 `paragraphs`만, 섞인 양식이면 둘 다 쓴다.

```json
{
  "table": 0,
  "fills": [
    {"row": 1, "col": 0, "mode": "inline", "fields": ["title"]},
    {"row": 2, "col": 0, "mode": "inline", "fields": ["date"]},
    {"row": 6, "col": 0, "mode": "block", "fields": ["discussion", "decisions"]}
  ],
  "paragraphs": [
    {"para": 1, "mode": "inline", "fields": ["purpose"]},
    {"para": 4, "mode": "block",  "fields": ["decisions", "action_items"]}
  ]
}
```

- `fills` 항목은 표 칸을 `{row, col, mode, fields}`로, `paragraphs` 항목은 본문 문단을 `{para, mode, fields}`로 지정한다. `para`는 inspect의 `paragraphs[].index`와 일치.
- `fields`: 고정 어휘만 허용 — `title, date, attendees, purpose, discussion, decisions, action_items, next_meeting, notes`. (알 수 없는 이름은 apply가 오류)
- `mode`:
  - `inline` — 대상 칸/문단의 기존 텍스트 끝에 토큰을 덧붙임(라벨 서식 유지). 라벨과 값이 같은 자리거나("제목: __", "ㅇ (목적)"), 값 전용 빈 자리일 때.
  - `block` — 목록형(discussion/decisions/action_items/notes)이나 여러 항목을 한 자리에 넣을 때. 대상 자리에 **라벨 텍스트가 있으면 지우지 않고**(inline과 동일한 라벨 보존) 블록 전체를 그 **바로 뒤 문단으로 순서대로 삽입**하고, 빈 자리면 첫 줄을 그 자리에 넣어 서식을 유지한다(뒤 문단 밀림은 스크립트가 대상 객체를 먼저 스냅샷해 처리).
- **채우는 순서(라벨 칸 먼저)**: Claude는 라벨(글자)이 있는 칸을 먼저 채운 뒤 빈 칸/넓은 영역을 채운다.
- **중복 금지(한 항목은 한 곳만)**: 이미 넣은 항목을 다른 칸/블록에 반복하지 않는다. 같은 라벨이 여러 칸에 있으면(예: "제목:"이 2곳) 같은 값을 복사하지 말고 각 칸에 서로 다른 알맞은 항목을 배치한다(예: title / purpose). 마땅한 별도 항목이 없으면 빈칸으로 둔다.
- **데이터 보존 규칙**: Claude는 비어 있지 않은 9항목이 모두 어느 fill/paragraph엔가 포함되도록 매핑한다. 양식에 전용 자리가 없는 항목(예: purpose)은 가장 어울리는 자리에 함께 넣는다. 양식에만 있고 매핑할 데이터가 없는 자리는 매핑을 만들지 않는다(빈칸 유지).

### 4.3 apply_form_mapping.py — 토큰 블록 라이브러리

각 field는 mode에 따라 아래 토큰 텍스트로 확장된다(검증된 블록만).

**inline 토큰**
| field | inline 토큰 |
|---|---|
| title | `{{ title }}` |
| date | `{{ date }}` |
| attendees | `{{ attendees_joined }}` |
| purpose | `{{ purpose }}` |
| next_meeting | `{{ next_meeting }}` |

**block 토큰(문단 라인 목록)** — block mode에서 각 field 앞에 한글 섹션 라벨을 붙인다(같은 칸에 여러 field가 있을 때 구분).
| field | 섹션 라벨 | 블록 라인 |
|---|---|---|
| purpose | `[회의 목적]` | `[회의 목적] {{ purpose }}` |
| next_meeting | `[다음 회의]` | `[다음 회의] {{ next_meeting }}` |
| attendees | `[참석자]` | `[참석자] {{ attendees_joined }}` |
| discussion | `[논의 내용]` | `{%p for d in discussion %}` / `[{{ d.topic }}]` / `{%p for p in d.points %}` / ` - {{ p }}` / `{%p endfor %}` / `{%p endfor %}` |
| decisions | `[결정 사항]` | `{%p for x in decisions %}` / ` - {{ x }}` / `{%p endfor %}` |
| action_items | `[실행 항목]` | `{%p for a in action_items %}` / ` - {{ a.task }} (담당: {{ a.owner }} / 기한: {{ a.due }})` / `{%p endfor %}` |
| notes | `[기타·특이사항]` | `{%p for n in notes %}` / ` - {{ n }}` / `{%p endfor %}` |

- inline mode: 각 field의 inline 토큰을 공백으로 이어 대상 칸 첫 문단 끝에 덧붙인다.
- block mode: fill의 `fields` 순서대로, field가 **하나면 섹션 라벨 생략**, **둘 이상이면 각 field에 섹션 라벨 줄**을 붙여 이어 넣는다. 대상 자리에 라벨 텍스트가 있으면 그대로 두고 블록 전체를 그 뒤에 새 문단으로 삽입하고, 빈 자리면 첫 라인을 `paragraphs[0]`에 넣고 나머지를 `add_paragraph`(문단 양식은 대상 뒤 삽입)로 추가한다 — inline과 같은 라벨 보존 규칙.
- title/date는 목록형이 아니므로 block에서도 inline 토큰 한 줄로 처리한다(섹션 라벨 규칙 동일).

## 5. 디렉터리 구조

```
.claude/skills/meeting-minutes/
├─ SKILL.md                          # [4]에 자동 매핑 경로 안내 추가 (수정)
├─ scripts/
│  ├─ inspect_template.py            # ★신규: 표(tables)+문단(paragraphs) 구조 JSON 덤프
│  ├─ apply_form_mapping.py          # ★신규: 매핑 → 토큰 삽입(표 fills + 문단 paragraphs, +블록 라이브러리)
│  ├─ output_naming.py               # ★신규: 최종 파일명 '회의록_<제목>_<생성일>'
│  ├─ render_docx_template.py        # (재사용) autoescape=True로 &,<,> 보존
│  ├─ render_docx.py / render_markdown.py / validate.py  # (불변)
│  └─ …
├─ schema/minutes.schema.json        # (불변)
└─ tests/
   ├─ test_inspect_template.py       # ★신규 (문단 덤프 포함)
   ├─ test_apply_form_mapping.py     # ★신규 (문단 매핑 포함)
   └─ test_output_naming.py          # ★신규
```

## 6. 실행 방식 — SKILL.md [4] 오케스트레이션

Claude가 [4]에서 분기한다.

1. 양식 없음 → `render_docx.py`.
2. 양식 있음 → `inspect_template.py <양식>`로 구조 확인.
   - `has_tokens == true` → 그대로 `render_docx_template.py <양식> minutes.json 최종.docx`.
   - `has_tokens == false`(표·문단·혼합) → 자동 매핑:
     - Claude가 구조 JSON(`tables`+`paragraphs`)을 읽고 `mapping.json` 작성(표는 `fills`, 문단은 `paragraphs`, 데이터 보존 규칙 준수).
     - `apply_form_mapping.py <양식> mapping.json output/양식_tokenized.docx`
     - `render_docx_template.py output/양식_tokenized.docx minutes.json 최종.docx`
     - **매핑 요약을 사용자에게 텍스트로 보고**(어느 칸/문단에 무엇을 넣었는지, 데이터 없는 빈자리, 자리 없어 몰아넣은 항목).
3. 최종 docx 파일명은 `output_naming.py`로 `회의록_<제목>_<생성일>.docx`를 구해 쓴다.

## 7. 에러 처리

| 상황 | 처리 |
|---|---|
| mapping.json의 알 수 없는 field 이름 | `apply`가 허용 어휘를 안내하며 오류 |
| mapping의 (row,col)이 표 범위 밖 | 명확한 IndexError 대체 메시지 |
| mapping의 `para`가 문단 범위 밖 | 명확한 IndexError 대체 메시지("문단 인덱스 …") |
| 양식에 표가 없음(문단전용) | inspect의 `paragraphs`로 문단 매핑(`paragraphs`) 진행 |
| 토큰이 이미 있는 양식에 자동 매핑 시도 | `has_tokens`로 걸러 기존 렌더러로 우회 |
| render 단계 문법 오류 | 기존 render_docx_template의 §8 래핑 메시지 사용 |

## 8. 테스트 전략

**결정적 부분(스크립트)만 테스트한다. Claude 매핑은 테스트 대상이 아니다.**

- `test_inspect_template.py`: python-docx로 표 양식을 만들어 → 좌표·`is_empty`·`merged` 원점 dedup·`has_tokens`(토큰 있는/없는 두 케이스)를 검증. 추가로 **머리말/꼬리말·중첩표에만 토큰이 있어도 탐지**하는지 검증.
- `test_apply_form_mapping.py`:
  - inline: 라벨 칸에 fill 적용 후 토큰이 라벨 뒤에 붙는지(라벨 텍스트 보존).
  - block 단일 field: 섹션 라벨 없이 블록만.
  - block 복수 field: 각 field에 섹션 라벨이 붙는지, 순서 유지.
  - **block 라벨 보존**: 라벨이 있는 칸/문단에 block을 적용해도 라벨이 지워지지 않고 블록이 그 뒤에 삽입되는지(표·문단 각각).
  - 알 수 없는 field/범위 밖 좌표 → 명확한 오류.
  - **엔드투엔드**: apply로 토큰 삽입 → `render_docx_template.render_template`로 sample minutes를 채워 → 결과 docx의 해당 칸에 값이 들어갔는지(빈 값은 빈칸, 목록은 반복) 확인.
- 블록 라이브러리: 각 field → 토큰 라인 목록 단위 테스트.

## 9. 범위 밖 (Non-Goals)

- 라벨↔값 자리(칸·문단) 인접성의 완벽 자동 판정 — Claude가 좌표/문단 index를 결정하고 스크립트는 실행만.
- 9항목 스키마·[2]추출·기존 렌더러 변경.
- 매핑 승인 게이트(자동 진행, 요약 보고만).
- 이미지/머리말/꼬리말 등 표 밖 요소 채우기.
