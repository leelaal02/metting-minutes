# 양식 템플릿 표시자 (스마트 토큰)

사용자가 자기 양식에 토큰을 직접 넣어 왔을 때(`inspect_template.py`가 `has_tokens: true`)
그 문법을 확인하거나, 사용자에게 "양식에 이렇게 넣으시면 됩니다"라고 안내할 때 읽는다.
토큰 없는 양식을 자동으로 채우는 경로는 `references/form-mapping.md`를 본다
— 그 경로에서는 토큰을 직접 쓰지 않고 스크립트가 만들어 준다.

복사해 쓸 예시는 `templates/example-template.docx` (재생성: `python scripts/make_example_template.py`).

## 토큰 표

| 토큰 | 채워지는 값 |
|---|---|
| `{{ title }}` | 회의 제목 |
| `{{ date }}` | 회의 일시 (없으면 빈칸) |
| `{{ purpose }}` | 회의 목적 |
| `{{ next_meeting }}` | 다음 회의 (없으면 빈칸) |
| `{{ attendees_joined }}` | 참석자 한 줄 결합 "홍길동, 김철수" |
| `{{ attendee_count }}` | 참석 인원수(숫자) — "총 `{{ attendee_count }}`명 참석"처럼 문장에 끼워 쓴다. 명단 길이에서 나오므로 참석자와 어긋나지 않는다 |
| `{% for a in attendees %}{{ a }}{% endfor %}` | 참석자 목록 반복 |
| `{% for d in discussion %}{{ d.topic }} … {% for p in d.points %}{{ p }}{% endfor %}{% endfor %}` | 논의 주제·포인트 반복 |
| `{% for x in decisions %}{{ x }}{% endfor %}` | 결정 사항 반복 |
| `{% for x in open_issues %}{{ x }}{% endfor %}` | 미결 사항 반복 |
| 실행 항목 표 — 아래 "표 행 반복" 참고 (`action_items` 사용) | 실행 항목 표 — 행 자동 반복 |
| `{% for n in notes %}{{ n }}{% endfor %}` | 기타·특이사항 반복 |

## 표 행 반복

**표 행 반복(`{%tr%}`)은 3행 구조로 넣는다** — 한 행에 for와 endfor를 함께 넣으면
동작하지 않는다. 표에 다음 3개 행을 만들고, `{%tr%}`가 든 for·endfor 행은
렌더 시 삭제되며 그 사이 데이터 행이 항목 수만큼 반복된다:

| 할 일 | 담당자 | 기한 |
|---|---|---|
| `{%tr for a in action_items %}` | (빈칸) | (빈칸) |
| `{{ a.task }}` | `{{ a.owner }}` | `{{ a.due }}` |
| `{%tr endfor %}` | (빈칸) | (빈칸) |

문단 단위로 반복시키려면 `{% %}` 대신 `{%p %}`를 쓴다.

없는 값은 이렇게 처리되어 양식이 깔끔하게 유지된다:

- 스칼라(`date`·`purpose`·`next_meeting` 등)가 없으면 **빈칸**.
- 목록(`decisions`·`notes` 등)이 비면 반복 자체가 돌지 않아 **행·문단 미생성**.
- 단 실행 항목의 `{{ a.owner }}`·`{{ a.due }}`는 값이 없으면 빈칸이 아니라 **`-`**로 채워진다 — 표 칸이 통째로 비면 담당·기한을 아직 안 정한 것인지 표가 깨진 것인지 구분되지 않기 때문. 빨간 "입력필요"로 받고 싶으면 토큰을 직접 쓰는 대신 토큰 없는 양식으로 주고 자동 매핑(`references/form-mapping.md`)을 쓴다.
