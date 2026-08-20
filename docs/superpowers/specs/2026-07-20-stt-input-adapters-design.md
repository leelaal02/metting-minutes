# STT 입력 어댑터 3종 — 설계 문서 (요구사항)

- 작성일: 2026-07-20
- 상태: 승인됨 (구현 계획 작성 단계로 이행)
- 선행 문서: [2026-07-16-meeting-minutes-skill-design.md](./2026-07-16-meeting-minutes-skill-design.md)

## 1. 목표

기존 회의록 자동 생성 Skill의 **[1] 입력 계층**에 세 가지 입력 어댑터(텍스트 / 로컬 STT / 클라우드 STT)를 플러그인 형태로 추가한다. 하위 파이프라인([2]추출 → [3]Markdown → [4]docx)은 **한 줄도 수정하지 않는다.**

사용자 방침: **세 어댑터를 일단 다 만들고, 나중에 필요 없는 것은 삭제한다.** 따라서 어댑터 추가·삭제가 각각 한 줄로 끝나는 레지스트리 구조가 핵심이다.

## 2. 확정된 설계 결정

| 항목 | 결정 | 이유 |
|---|---|---|
| 텍스트 입력 | 기존 `normalize_input.py` 재사용 | 이미 완성·검증됨 |
| 로컬 STT | **faster-whisper** (`WhisperModel("medium", cpu_threads=4)`) | pip 설치로 동작, 로컬 오프라인, 한국어 우수 (사용자 테스트 중 확정) |
| 클라우드 STT | **Groq Whisper API** (`whisper-large-v3`) | 초고속 Whisper API, `GROQ_API_KEY`만 필요 |
| 실행 주체 | 사용자는 `/meeting-minutes` 호출, **Claude가 SKILL.md 지시로 헬퍼 실행** | Harness 패턴. 사용자가 직접 `python`을 치지 않음 |
| 화자 분리 | **미지원 (순수 통문장 텍스트)** | 두 STT 모두 Whisper 기반이라 diarization 없음. 필요 시 별도 확장 |
| 어댑터 선택 | **명시적 소스 지정** `--source text\|whisper\|groq` | 추가·삭제가 한 줄, 어느 STT를 쓸지 명확 |
| 필요 패키지 | `requirements-stt.txt`로 **선택적 분리** | 텍스트만 쓰는 사용자는 무거운 STT 패키지 불필요 |
| 어댑터 로딩 | **lazy import** | 미설치 상태에서도 텍스트 파이프라인 정상 동작 |

## 3. 아키텍처

### 3.1 전체 파이프라인 (입력 계층만 확장)

```
[1] 입력 계층  ← 이번 작업 (STT 어댑터 3종 추가)
   text | whisper | groq  →  정규화된 회의 텍스트 (str)
        ↓
[2] 추출 계층  → Claude가 SKILL.md 지시로 minutes.json 생성
        ↓ minutes.json (단일 진실 원천, schema 검증)
[3] Markdown 렌더 (완성)  →  [4] docx 렌더 (완성)
```

의존 방향은 `[1] → [2] → [3] → [4]` 단방향이며 역방향 의존이 없다. 입력 방식 추가/삭제는 [1]에서만 발생한다.

### 3.2 입력 계층 내부 (디스패처 + 레지스트리)

```
python scripts/transcribe.py --source <name> <입력>
                │
                ▼
        transcribe.py (디스패처)
        REGISTRY[name] 선택
                │
   ┌────────────┼────────────┐
   ▼            ▼            ▼
 text        whisper       groq
 (.txt/      _local        _cloud
  붙여넣기)  (faster-      (Groq
             whisper)      Whisper API)
   └────────────┼────────────┘
                ▼
   transcribe(source, **opts) -> str   (공통 인터페이스)
                ▼
        _normalize() 정규화
                ▼
     정규화된 회의 텍스트 (str)  →  [2] 추출로
```

## 4. 데이터 계약 — 어댑터 공통 인터페이스

모든 어댑터는 동일한 시그니처를 구현한다.

```python
def transcribe(source: str, **opts) -> str:
    """오디오 경로 또는 텍스트 → 정규화된 회의 원문(str)."""
```

| 항목 | 내용 |
|---|---|
| 입력 | text: 문자열 또는 `.txt` 경로 · STT: 오디오 파일 경로 |
| 출력 | **정규화된 회의 텍스트(str)** — 세 어댑터 모두 동일 형태 |
| 정규화 | 어댑터가 반환 전 `normalize_input._normalize()`를 적용 (하위 단계가 소비하는 형태 통일) |
| 등록 | `adapters/__init__.py`의 `REGISTRY`에 `"name": 함수` 한 줄 |
| 제거 | 레지스트리 한 줄 삭제 + 어댑터 파일 삭제 → 하위 단계 영향 없음 |

`REGISTRY`는 어댑터 모듈을 lazy import하여 `{"text": ..., "whisper": ..., "groq": ...}` 형태로 이름→호출 함수를 매핑한다. STT 어댑터의 무거운 import는 실제 호출 시점까지 지연한다.

## 5. 디렉터리 구조

```
metting/
├─ SKILL.md                      # [1] 단계 지시문 업데이트 (세 소스 안내)
├─ requirements.txt              # 기존: python-docx, jsonschema, pytest
├─ requirements-stt.txt          # ★신규: faster-whisper, groq (선택 설치)
├─ schema/minutes.schema.json    # 데이터 계약 (불변)
├─ scripts/
│  ├─ transcribe.py              # ★신규: CLI 디스패처 + 레지스트리 진입
│  ├─ adapters/                  # ★신규: 입력 어댑터 플러그인
│  │  ├─ __init__.py             #        REGISTRY = {text, whisper, groq}
│  │  ├─ text.py                 #        기존 normalize_input 재사용
│  │  ├─ whisper_local.py        #        faster-whisper
│  │  └─ groq_cloud.py           #        Groq Whisper API
│  ├─ normalize_input.py         # 기존: _normalize 공용 함수로 유지
│  ├─ validate.py                # (불변)
│  ├─ render_markdown.py         # (불변)
│  └─ render_docx.py             # (불변)
├─ examples/
│  └─ sample_meeting.txt         # 기존 샘플 입력 (텍스트 어댑터 회귀용)
├─ tests/
│  ├─ test_transcribe.py         # ★신규: 디스패처/레지스트리
│  └─ test_adapters.py           # ★신규: 어댑터 (STT는 모킹)
└─ output/
```

★ = 이번 작업에서 추가/수정. 표시 없는 항목은 불변.

## 6. 실행 방식 — 사용자는 CLI를 직접 치지 않는다

사용자 관점의 진입점은 오직 **`/meeting-minutes` 스킬 호출**이다. `transcribe.py`는 사용자가 아니라 **Claude가 [1] 단계에서 실행하는 결정적 헬퍼**이며, 이는 기존 [3]`render_markdown.py`·[4]`render_docx.py`와 동일한 Harness 패턴이다.

```
사용자:  /meeting-minutes  (+ 오디오 또는 텍스트 제공)
   │
   ▼
Claude가 SKILL.md 지시대로 오케스트레이션:
   [1] transcribe.py 실행       → 정규화된 회의 텍스트
   [2] 텍스트 → minutes.json    (Claude가 직접 추출)
   [3] render_markdown.py 실행  → 미리보기·검토
   [4] render_docx.py 실행      → 최종 .docx
```

Claude가 [1]에서 실제로 실행하는 명령 형태 (내부 헬퍼):

```bash
python scripts/transcribe.py --source text    회의.txt   > output/meeting.txt
python scripts/transcribe.py --source whisper 회의.wav   > output/meeting.txt
python scripts/transcribe.py --source groq    회의.m4a   > output/meeting.txt
```

- 표준 출력으로 정규화된 회의 텍스트를 내보내며, 이를 [2] 추출 단계로 넘긴다.
- `--source` 미지정 시 기본값은 `text`.
- 소스 판별: SKILL.md의 [1] 단계 지시문이 "오디오 파일이면 로컬/클라우드 중 무엇을 쓸지 사용자에게 확인, 텍스트면 `text`"를 Claude에게 안내한다.

## 7. 어댑터별 구현 요지

### 7.1 text (`adapters/text.py`)
- 기존 `normalize_input.load_meeting_text(source)`를 그대로 위임.
- 문자열 또는 `.txt` 경로 → 정규화 텍스트.

### 7.2 whisper_local (`adapters/whisper_local.py`)
- `faster_whisper.WhisperModel`을 lazy import.
- 기본 설정: `WhisperModel("medium", cpu_threads=4)`, 전사 시 `language="ko"`. 모델 크기·스레드 수는 `opts`로 조정 가능.
- 세그먼트 텍스트를 이어 붙여 통문장 생성 후 `_normalize()` 적용.
- `faster-whisper` 미설치 시 설치 안내 메시지와 함께 명확한 오류.

### 7.3 groq_cloud (`adapters/groq_cloud.py`)
- `groq.Groq` 클라이언트를 lazy import.
- `GROQ_API_KEY` 환경변수 사용. 미설정 시 명확한 안내 오류.
- 오디오 파일을 `whisper-large-v3`로 전사(`language="ko"` 기본) → 텍스트 → `_normalize()` 적용.
- `groq` 미설치 시 설치 안내 메시지와 함께 명확한 오류.

## 8. 에러 처리

| 상황 | 처리 |
|---|---|
| STT 패키지 미설치 | `pip install -r requirements-stt.txt` 안내 후 오류 |
| `GROQ_API_KEY` 미설정 | 환경변수 설정 안내 후 오류 |
| 오디오 파일 없음 | `FileNotFoundError` |
| 변환 결과 빈 텍스트 | `ValueError` (기존 정규화 규칙과 동일) |
| 알 수 없는 `--source` | 사용 가능한 소스 목록 출력 후 비정상 종료 |

## 9. 테스트 전략

실제 오디오·네트워크 없이 결정적으로 검증한다.

- **디스패처/레지스트리** (`test_transcribe.py`): `--source` 라우팅이 올바른 어댑터를 호출하는지, 미지원 소스가 명확히 실패하는지, 기본값이 `text`인지.
- **text 어댑터** (`test_adapters.py`): 기존 정규화 동작 유지 회귀 테스트.
- **whisper/groq 어댑터** (`test_adapters.py`): STT 라이브러리를 **모킹**하여 (1) 전사 결과에 정규화가 적용되는지, (2) 미설치·키 없음 시 정확한 오류 메시지가 나오는지.
- **통합**: `--source text`는 실제로 끝까지 동작.

## 10. 범위 밖 (Non-Goals)

- 화자 분리(diarization) 및 화자 라벨 텍스트 생성.
- 실시간 스트리밍 STT.
- 오디오 포맷 변환·전처리(ffmpeg 등 명시적 파이프라인화).
- STT 결과 캐싱·이력 관리.
- 세 어댑터 외 추가 STT 제공자 (필요 시 동일 패턴으로 확장).
