---
name: meeting-minutes
description: Use when the user wants structured meeting minutes (회의록) from meeting notes, a transcript, STT output, or an audio recording (녹취·녹음 파일) — exported to Markdown and Word (.docx). Also use when they want the minutes filled into their own .docx form/template (회의록 양식·템플릿 채우기). Extracts 10 fixed items: 제목/일시/참석자/회의 목적/논의 내용/결정 사항/미결 사항/실행 항목(Action Items)/다음 회의/기타·특이사항. Trigger this whenever someone turns meeting content, recordings, or transcripts into minutes — even if they phrase it as "문서로 정리해줘" and don't say "회의록" explicitly. Do NOT use for plain audio transcription without structuring, planning a meeting agenda in advance, scheduling meetings on a calendar, or translating an existing 회의록 — those are different tasks.
---

# 회의록 자동 생성 Skill

텍스트·STT·오디오(녹취)를 구조화된 회의록(Markdown → Word .docx)으로 변환한다. 사용자가 준
.docx 양식이 있으면 그 양식에 맞춰 채운다. 각 단계의 계약은 `schema/minutes.schema.json`이다.

## 참조 문서 — 해당 분기에 들어갈 때만 읽는다

- `references/form-mapping.md` — [4]에서 **토큰 없는 양식을 자동으로 채울 때 반드시.** `mapping.json` 작성 규칙.
- `references/template-tokens.md` — 사용자가 양식에 토큰을 직접 넣어 왔거나, 넣는 법을 안내할 때.
- `references/input-sources.md` — 입력 파일을 못 찾아 탐색 폴더를 넓히거나, 입력 어댑터를 추가·삭제할 때.

## 경로·산출물 규칙 (중요)

- `scripts/`·`schema/`·`references/`는 **이 SKILL.md가 있는 스킬 디렉토리** 기준 — 실행할 때
  절대경로로 바꾼다(예: `python .claude/skills/meeting-minutes/scripts/transcribe.py ...`).
- `output/`은 **사용자의 현재 작업 디렉토리** 기준. **최종 docx만 `output/`에, 중간 파일
  (`meeting.txt`·`minutes.json`·`minutes.md`·양식 구조 JSON·`mapping.json`·`*_tokenized.docx`)은
  전부 `output/.work/`에** 둔다 — 아래 명령이 이미 이 규칙대로 적혀 있다.
- 최종본 `output/회의록_<제목>_<생성일 YYYY-MM-DD>.docx` — 이름은 손으로 짓지 말고 [4]에서
  `output_naming.py`로 구한다(날짜는 회의 일시가 아니라 **파일을 만든 날**).
- `output/`은 통째로 `.gitignore` 대상이다. 중간 파일은 다음 수정 때 재사용하게 지우지 않는다.

## 사전 준비

`pip install -r requirements.txt`(최초 1회, 오디오 STT는 `requirements-stt.txt`도) · `mkdir -p output/.work`.

## [1] 입력 확보

`python scripts/transcribe.py --source <text|whisper|groq> <입력> > output/.work/meeting.txt`

- **텍스트**(붙여넣은 내용 또는 `.txt` 경로): `--source text`로 바로 실행.
- **오디오**: 기본은 클라우드 `--source groq`(작업 폴더 `.env`의 `GROQ_API_KEY=...`를 자동 로드,
  실제 환경변수가 있으면 그쪽 우선). 오프라인이 필요하면 `--source whisper`
  (`requirements-stt.txt` 필요). 오디오를 외부로 보내는 문제라 **어느 쪽을 쓸지 사용자에게 확인**한다.

이 표준 출력(정규화된 회의 원문)을 [2]의 입력으로 쓴다. 입력 파일은 **이름만** 줘도 공통 위치를
탐색해 찾는다. 실패하면 원인(라이브러리 미설치·키 미설정·파일 못 찾음)이 안내되므로 그대로
사용자에게 전달해 설치/설정을 요청한다.

## [2] 추출 (이 단계는 네가 직접 수행)

회의 원문을 읽고 스키마를 **정확히** 따르는 `minutes.json`을 `output/.work/`에 쓴다. 렌더러가
`load_minutes()`로 로드하며 스키마를 자동 검증하므로, ValidationError가 나면 맞게 고친다.

- 10개 항목을 모두 채운다: title, date, attendees, purpose, discussion, decisions, open_issues,
  action_items, next_meeting, notes. `discussion`은 주제별 `{topic, points}`로 묶고, `purpose`는
  이 회의를 왜 하는지 한두 문장, `notes`는 어느 항목에도 안 들어가는 기타·특이사항을 배열로.
- 원문에 없는 정보는 **지어내지 않는다** — 없으면 `null`(date/purpose/next_meeting/owner/due)
  또는 빈 배열(attendees/decisions/open_issues/action_items/notes).
- **`decisions`와 `open_issues`를 섞지 않는다** — 회의록에서 가장 먼저 찾는 구분이다. 결론이 난
  것만 `decisions`, "~검토·~미정·~이견·~필요"로 끝나는 것은 `open_issues`로 빼고 무엇이 걸려
  있는지(선택지·쟁점)를 함께 적는다.
  - 결정 ✓ "지역별 단가 예측 제외, 전국 단가로 진행"
  - 미결 ✓ "LLM 구축형 vs 상용 API 미정 (직접 구축 시 보안·버전 관리 부담)"
- **문체는 개조식(箇條式) — 서술식 금지.** 회의록은 한눈에 들어와야 하므로 모든 텍스트 값
  (`purpose`·`points`·`decisions`·`open_issues`·`action_items[].task`·`notes`)을 명사·체언 종결로
  압축한다. "~합니다/~한다/~이다" 종결과 "따라서·또한·그리고" 같은 접속 부사를 빼고 한 항목에 한
  사실만 담는다. 수치·근거·조건은 괄호로, 인과·순서는 `→`, 대비는 `vs`로. 부가 설명이 꼭 필요하면
  개조식+서술식 혼합까지만 허용.
  - **압축하되 정보를 빠뜨리거나 왜곡하지 않는다** — 문장을 줄이는 것이지 사실을 줄이는 게 아니다.
    원문의 수치·주체·조건은 그대로 살린다.
  - 서술식 ✗ "라마단이 지난 8월 18일에 끝났습니다. 따라서 중동 항로의 거래량과 실제 적재 비율이 다시 늘어날 것으로 보입니다."
  - 개조식 ✓ "라마단 종료 후 중동항로 물동량 회복 예상 (8/18)"

## [3] Markdown 생성 및 사용자 검토

`python scripts/render_markdown.py output/.work/minutes.json output/.work/minutes.md`

생성된 Markdown을 사용자에게 보여주고 검토를 요청한다. 수정 요청이 오면 `minutes.json`을 고치고 반복한다.

## [4] docx 생성 (사용자 승인 후)

**먼저 최종 파일명을 구한다:** `python scripts/output_naming.py output/.work/minutes.json .docx [양식.docx]`
→ 출력된 이름을 `output/<그 이름>`으로 쓴다. **양식을 쓸 때는 세 번째 인자로 양식 경로를 넘긴다**
— 양식명이 앞에 붙어(`누리미디어_회의록_…`) 같은 회의를 여러 양식에 채워도 서로 덮이지 않는다.
양식이 없으면 생략(`회의록_3분기 로드맵_2026-07-22.docx`).

- **양식 없음(기본):** `python scripts/render_docx.py output/.work/minutes.json output/<최종.docx>`
- **양식 있음:** `python scripts/inspect_template.py <template.docx>`로 구조 JSON을 얻어 분기한다
  (`has_tokens`, `blocks`[표·문단이 놓인 순서], `tables`[라벨·좌표·`is_empty`·`merged`·`shaded`],
  `paragraphs`[index·텍스트·`is_empty`·`numbered`]).
  - **토큰 있음(`has_tokens: true`):**
    `python scripts/render_docx_template.py <template.docx> output/.work/minutes.json output/<최종.docx>`
  - **토큰 없음(`has_tokens: false`) → 자동 매핑**(표·문단 양식 모두 지원):
    1. **`references/form-mapping.md`를 읽고** 10항목을 알맞은 칸/문단에 배치한 매핑을
       `output/.work/mapping.json`으로 작성한다.
    2. `python scripts/apply_form_mapping.py <template.docx> output/.work/mapping.json output/.work/<양식>_tokenized.docx`
    3. `python scripts/render_docx_template.py output/.work/<양식>_tokenized.docx output/.work/minutes.json output/<최종.docx>`
    4. **매핑 요약을 사용자에게 텍스트로 보고**한다 — 어느 칸/문단에 무엇을 넣었는지, 데이터 없이
       비워둔 곳, 전용 자리가 없어 다른 곳에 함께 넣은 항목.
  - 양식 파일도 **이름만** 줘도 탐색해 찾는다. **hwp/pdf 양식**은 한글/워드에서 **.docx로 저장
    (다른 이름으로 저장 → Word)**해 달라고 요청한다. 표시자 문법 오류·.docx 아님·파일 없음은
    명확한 오류가 나므로 그대로 사용자에게 전달한다.

최종 `.docx` 경로를 사용자에게 안내한다.
