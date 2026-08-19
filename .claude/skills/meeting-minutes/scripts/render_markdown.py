"""[3] Markdown 렌더러: minutes.json → Markdown 문자열."""
import sys

from validate import load_minutes

def _md_cell(value: str) -> str:
    """Markdown 표 셀용 이스케이프.

    셀 값에 `|`가 있으면 가짜 열이 생기고, 개행이 있으면 행이 끊겨 표가 깨진다.
    파이프는 `\\|`로 이스케이프하고 개행은 공백으로 바꿔 한 셀·한 행을 유지한다.
    """
    return (
        value.replace("|", "\\|")
        .replace("\r\n", " ")
        .replace("\n", " ")
        .replace("\r", " ")
    )


def render_markdown(data: dict) -> str:
    lines = [f"# {data['title']}", ""]
    if data.get("date"):
        lines.append(f"**일시:** {data['date']}")
        lines.append("")

    lines.append("## 참석자")
    if data["attendees"]:
        lines.extend(f"- {a}" for a in data["attendees"])
    else:
        lines.append("- (없음)")
    lines.append("")

    lines.append("## 회의 목적")
    lines.append(data["purpose"] or "(없음)")  # 
    lines.append("")

    lines.append("## 논의 내용")
    if data["discussion"]:
        for item in data["discussion"]:
            lines.append(f"### {item['topic']}")
            lines.extend(f"- {p}" for p in item["points"])
            lines.append("")
    else:
        lines.append("(없음)")
        lines.append("")

    lines.append("## 결정 사항")
    if data["decisions"]:
        lines.extend(f"- {d}" for d in data["decisions"])
    else:
        lines.append("- (없음)")
    lines.append("")

    lines.append("## 미결 사항")
    if data["open_issues"]:
        lines.extend(f"- {o}" for o in data["open_issues"])
    else:
        lines.append("- (없음)")
    lines.append("")

    lines.append("## 실행 항목 Action Items")
    if data["action_items"]:
        lines.append("| 할 일 | 담당자 | 기한 |")
        lines.append("| --- | --- | --- |")
        for a in data["action_items"]:
            task = _md_cell(a["task"])
            owner = _md_cell(a["owner"] or "-")
            due = _md_cell(a["due"] or "-")
            lines.append(f"| {task} | {owner} | {due} |")
    else:
        lines.append("(없음)")
    lines.append("")

    lines.append("## 다음 회의 일정")
    lines.append(data["next_meeting"] or "(미정)")
    lines.append("")

    lines.append("## 기타·특이사항")
    if data["notes"]:
        lines.extend(f"- {n}" for n in data["notes"])
    else:
        lines.append("- (없음)")
    lines.append("")

    return "\n".join(lines)


def main() -> None:
    json_path, out_path = sys.argv[1], sys.argv[2]
    data = load_minutes(json_path)
    md = render_markdown(data)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(md)
    print(f"Markdown 생성 완료: {out_path}")


if __name__ == "__main__":
    main()
