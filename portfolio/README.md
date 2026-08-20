# 포트폴리오 — 회의록 자동 생성 파이프라인

녹음 파일이나 회의 메모를 넣으면 10개 항목으로 구조화되고, 사용자가 검토한 뒤
**자기 회사 회의록 양식(.docx)에 그대로 채워진 문서**가 나오는 파이프라인입니다.

| | |
|---|---|
| 기간 | 2026.07 – 08 (5주) · 개인 프로젝트 |
| 역할 | 설계 · 구현 · 테스트 전담 (1인) |
| 스택 | Python · python-docx · docxtpl · jsonschema · pytest · Whisper |
| 실적 | 회사 양식 **5종**에 회의록 **26건** 실제 생성 · 테스트 **186개** 통과 |

---

## 어디부터 보면 되나

| 순서 | 파일 | 소요 |
|---|---|---|
| 1 | **[`index.html`](./index.html)** — 브라우저로 여는 포트폴리오 (로그인·인터넷 불필요) | 5분 |
| 2 | [`samples/`](./samples) — 파이프라인이 실제로 만든 파일들 (아래 표) | 2분 |
| 3 | [`PORTFOLIO.md`](./PORTFOLIO.md) — 같은 내용의 Markdown 판 (GitHub·노션 붙여넣기용) | 5분 |
| 4 | [`설계문서/`](./설계문서) — 코드를 쓰기 전에 남긴 설계서 6건·구현 계획 3건 | 필요할 때 |

> `index.html`은 더블클릭하면 열립니다. `Ctrl + P` → "PDF로 저장"으로 제출용 PDF가 나옵니다.

---

## `samples/` — 실제 산출물

**직접 생성한 파일들입니다.** 설명을 위해 그린 그림이 아니라, 이 저장소의 스크립트를
그대로 돌려 나온 결과입니다. 번호가 파이프라인 단계 순서입니다.

| 파일 | 파이프라인 단계 | 무엇을 보여주나 |
|---|---|---|
| `01_입력_회의원문.txt` | [1] 입력 | 정규화된 회의 원문 (STT 결과와 같은 형태) |
| `02_추출_minutes.json` | [2] 추출 | LLM이 뽑은 **10개 항목** — 스키마로 검증되는 단일 진실 원천 |
| `03_회의록_minutes.md` | [3] Markdown | 사람이 읽고 고칠 미리보기 |
| `04_결과_기본서식.docx` | [4] 출력 | 양식이 없을 때의 기본 서식 회의록 |
| `05_사용한_빈양식.docx` | [4] 입력 | **표시자가 없는 빈 회사 양식** — 라벨만 있고 값 칸은 비어 있음 |
| `06_양식구조_inspect.json` | [4] ① | 그 양식을 기계적으로 뜯어본 구조 (`shaded`·`numbered`·`blocks`) |
| `07_매핑_mapping.json` | [4] ② | 어느 항목을 어느 칸에 넣을지 결정한 배치 — **LLM은 여기까지만 관여** |
| `08_결과_표양식_채움.docx` | [4] ③ | 그 빈 양식이 채워진 최종 결과 |

**`05` → `08`을 나란히 열어 보시면 이 프로젝트가 하는 일이 한눈에 들어옵니다.**

`08`에서 확인할 수 있는 것:

- 논의 내용이 `1. 주제` / `- 항목` 2계층으로 정리됨 (양식이 쓰는 기호와 겹치지 않게 자동 선택)
- 실행 항목에 `(담당: 이정우 / 기한: 2026-07-23)` 자동 부착
- 기한이 원문에 없는 항목은 지어내지 않고 빨간 **입력필요** 표시
- `다음 회의` 칸 하나에 두 항목이 들어가자 `[다음 회의]` `[기타·특이사항]` 섹션 라벨이 자동 부여

---

## 직접 돌려보기

```bash
pip install -r .claude/skills/meeting-minutes/requirements.txt

# 테스트 (186개)
python -m pytest -q

# 01 → 03 : 원문에서 Markdown 회의록까지
python .claude/skills/meeting-minutes/scripts/render_markdown.py \
       portfolio/samples/02_추출_minutes.json  /tmp/minutes.md

# 05 → 08 : 빈 양식을 채우는 3단계
python .claude/skills/meeting-minutes/scripts/inspect_template.py \
       portfolio/samples/05_사용한_빈양식.docx                      # ① 구조 덤프
python .claude/skills/meeting-minutes/scripts/apply_form_mapping.py \
       portfolio/samples/05_사용한_빈양식.docx \
       portfolio/samples/07_매핑_mapping.json  /tmp/tokenized.docx   # ② 표시자 삽입
python .claude/skills/meeting-minutes/scripts/render_docx_template.py \
       /tmp/tokenized.docx \
       portfolio/samples/02_추출_minutes.json  /tmp/최종.docx        # ③ 값 채움
```

실사용에서는 이 과정을 사용자가 직접 실행하지 않습니다 — `/meeting-minutes` 한 번이면
`SKILL.md`의 지시대로 순서가 자동으로 돌아갑니다.

---

## `설계문서/` — 코드보다 먼저 쓴 것

기능을 바로 짜지 않고 **설계서 → 구현 계획 → 구현 → 리뷰** 순서로 진행했고,
그 과정을 문서로 남겼습니다. 커밋 이력과 날짜가 맞물립니다.

| 폴더 | 내용 |
|---|---|
| `설계문서/설계서/` | 무엇을 왜 만드는지 · 확정한 설계 결정과 근거 **6건** |
| `설계문서/구현계획/` | 그 설계를 어떤 순서로 구현할지 **3건** |

특히 `2026-07-27-회의록-렌더링-고도화-*` 두 건은 **구현을 시작하기 전에 라이브러리를
직접 돌려 제약을 실측하고 확정한 기록**입니다 — 되돌리기 비싼 문제를 가장 싼 시점에
발견하려고 한 작업입니다.

---

## 프로젝트 본체 위치

포트폴리오 폴더에는 결과물과 문서만 담았습니다. 실행되는 코드는 저장소 안에 있습니다.

```
.claude/skills/meeting-minutes/
├── SKILL.md          오케스트레이션 지시문 (100줄로 유지)
├── schema/           단일 진실 원천 (minutes.schema.json)
├── scripts/          12개 모듈 + 입력 어댑터 3종 (2,191줄)
├── references/       조건부로 읽는 상세 규칙 3종
└── tests/            186개 (2,355줄)
```
