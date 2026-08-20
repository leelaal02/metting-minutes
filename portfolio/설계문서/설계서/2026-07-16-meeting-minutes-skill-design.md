# 회의록 자동 생성 Skill — 설계 문서 (요구사항)

- 작성일: 2026-07-16
- 상태: 승인됨 (구현 계획 작성 단계로 이행)

## 1. 목표

Harness Engineering을 적용한 **회의록 자동 생성 Claude Code Skill**. 텍스트(또는 향후 STT 결과)로 된 회의 내용을 입력받아 구조화된 회의록을 생성하고, Markdown을 거쳐 Word(.docx)로 출력한다.

## 2. 확정된 설계 결정

| 항목 | 결정 | 이유 |
|---|---|---|
| 산출물 형태 | Claude Code Skill (`SKILL.md`) + Python 헬퍼 스크립트 | 분석/추출은 LLM, 결정적 변환만 코드 |
| Harness 수준 | 명시적 4단계 분리, 단계 간 JSON 계약 | 유지보수·확장성. 단계 간 계약 문서화 |
| docx 변환 | python-docx (JSON 직접 소비) | pip만으로 동작, 외부 바이너리 불필요, 서식 제어 용이 |
| 입력 | 텍스트(붙여넣기/`.txt`)만. STT는 자리만 확보 | YAGNI. 확장 지점만 추상화 |
| 단일 진실 원천 | 추출 JSON(`minutes.json`) | Markdown·docx 렌더러가 동일 스키마 소비 |

## 3. 아키텍처 (4단계 파이프라인)

```
[1] 입력 계층        → 텍스트 정규화 (STT 어댑터 자리 확보)
        ↓ 정규화된 회의 원문 텍스트
[2] 추출 계층        → Claude가 SKILL.md 지시문에 따라 분석 → minutes.json
        ↓ minutes.json (단일 진실 원천)
[3] Markdown 렌더러  → minutes.md 생성 → 사용자에게 미리보기/검토
        ↓ 사용자 승인
[4] docx 렌더러      → minutes.json → minutes.docx (python-docx)
```

- **[2] 추출**: Claude(LLM)가 `SKILL.md` 지시문에 따라 수행 — 코드 아님.
- **[1] 입력 정규화 / [3] Markdown 렌더 / [4] docx 렌더**: 결정적이므로 Python 헬퍼 스크립트.
- Markdown이 먼저 생성되어 사용자 검토를 받고, 승인 후 같은 JSON에서 docx 생성 → "Markdown 우선" 규칙 충족.
- 각 단계는 독립 모듈이며, 단계 간 결합은 정규화 텍스트(str)와 `minutes.json`(스키마)이라는 명시적 계약으로만 이뤄진다.

## 4. 중간 계약 — `minutes.json` 스키마 (단일 진실 원천)

```json
{
  "title": "회의 주제 (문자열)",
  "date": "회의 일시 (문자열, 원문에서 추출·없으면 null)",
  "attendees": ["참석자1", "참석자2"],
  "discussion": [
    { "topic": "논의 주제", "points": ["논의 내용 요점", "..."] }
  ],
  "decisions": ["결정 사항1", "결정 사항2"],
  "action_items": [
    { "task": "할 일", "owner": "담당자 or null", "due": "기한 or null" }
  ],
  "next_meeting": "다음 회의 일정 (문자열 or null)"
}
```

필드 ↔ 요구 항목 매핑 (6종 모두 포함):

| 요구 항목 | JSON 필드 |
|---|---|
| 회의 주제 | `title` (+ `date`) |
| 참석자 | `attendees` |
| 논의 내용 | `discussion` |
| 결정 사항 | `decisions` |
| Action Items | `action_items` |
| 다음 회의 일정 | `next_meeting` |

- 스키마는 `schema/minutes.schema.json` 파일 하나로 두어, 추출 지시문·MD 렌더러·docx 렌더러 3자가 공유한다.
- 원문에 없는 정보는 **지어내지 않고** `null` 또는 빈 배열로 둔다. 이 규칙을 추출 지시문에 명시한다.

## 5. 디렉터리 구조

```
metting/
├─ CLAUDE.md
├─ SKILL.md                      # 스킬 진입점 (4단계 오케스트레이션 지시문)
├─ schema/minutes.schema.json    # 중간 계약
├─ scripts/
│   ├─ normalize_input.py        # [1] 입력 정규화 (+ STT 어댑터 자리)
│   ├─ render_markdown.py        # [3] JSON → md
│   └─ render_docx.py            # [4] JSON → docx (python-docx)
├─ examples/
│   ├─ sample_meeting.txt        # 샘플 입력
│   └─ sample_minutes.json       # 기대 추출 결과 (테스트 픽스처)
├─ tests/                        # 렌더러 단위 테스트
└─ requirements.txt              # python-docx, jsonschema, pytest
```

## 6. STT 확장 자리 (이번 범위에서 구현하지 않음)

`normalize_input.py` 안에서 입력 소스를 함수로 추상화한다.

```python
def load_meeting_text(source: str) -> str:
    """지금: .txt 파일 경로 또는 문자열을 정규화된 회의 원문으로 변환."""
    ...

# def load_from_stt(audio_path: str) -> str:
#     """자리만 확보. STT 추가 시 이 함수만 구현하면 하위 단계 무수정."""
#     raise NotImplementedError
```

STT 추가 시 이 함수 하나만 구현하면 [2]~[4] 단계는 수정할 필요가 없다 → 확장성 규칙 충족.

## 7. 에러 처리

- **스키마 검증**: 추출된 `minutes.json`을 `jsonschema`로 검증하여, 필수 항목 누락·타입 오류를 렌더링 이전에 차단한다.
- **입력 검증**: 빈 입력·존재하지 않는 파일 경로는 `normalize_input.py`에서 명확한 오류로 처리.
- **렌더러 견고성**: `null`/빈 배열 필드(예: `next_meeting: null`, 빈 `action_items`)를 렌더러가 깨지지 않고 처리(해당 섹션 생략 또는 "없음" 표기).

## 8. 테스트 전략

결정적 코드만 단위 테스트한다 (LLM 추출 자체는 테스트 대상 아님).

- `render_markdown.py`: `sample_minutes.json` → 기대 Markdown 구조 생성.
- `render_docx.py`: `sample_minutes.json` → 오류 없이 `.docx` 생성(파일 유효성/핵심 문단 존재 확인).
- `normalize_input.py`: 문자열/`.txt` 입력 정규화, 빈 입력 오류.
- 스키마 검증: 유효 JSON 통과, 필수 항목 누락 JSON 거부.
- 엣지 케이스: `next_meeting: null`, 빈 `attendees`/`action_items`/`decisions`.

## 9. 범위 밖 (Non-Goals)

- 오디오 파일에서의 실제 STT 변환 (자리만 확보).
- 화자 분리(diarization) 등 STT 후처리.
- Markdown/docx 외 출력 포맷.
- 회의록 저장소·검색·이력 관리.
