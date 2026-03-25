from __future__ import annotations

from typing import Any


SOURCE_LABELS = {
    "lecture_explicit": "講義で明示",
    "lecture_inferred": "講義内容から推論",
    "model_background": "一般知識補完",
}


def _stringify(value: Any, default: str = "") -> str:
    if value is None:
        return default
    return str(value)


def _as_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    if value is None:
        return []
    return [value]


def _format_time_range(value: Any) -> str | None:
    if isinstance(value, dict):
        start_time = value.get("start_time")
        end_time = value.get("end_time")
        if start_time and end_time:
            return f"{start_time} 〜 {end_time}"
        return None

    if isinstance(value, list):
        ranges = [_format_time_range(item) for item in value]
        ranges = [time_range for time_range in ranges if time_range]
        if ranges:
            return " / ".join(ranges)

    return None


def _format_source_label(source: Any) -> str:
    text = _stringify(source)
    return SOURCE_LABELS.get(text, text)


def _format_topic_heading(topic: dict[str, Any], index: int) -> str:
    topic_title = _stringify(topic.get("topic_title"), f"トピック {index}")
    time_range = _format_time_range(topic.get("topic_time_range"))
    if time_range:
        return f"### {topic_title} ({time_range})"
    return f"### {topic_title}"


def _append_block(lines: list[str], heading: str, values: list[str]) -> None:
    lines.append(heading)
    lines.append("")
    for value in values:
        lines.append(value)
    lines.append("")


def render_markdown_note(summary_data: dict[str, Any]) -> str:
    lines: list[str] = []
    title = _stringify(summary_data.get("title"), "タイトルなし")
    summary = _stringify(summary_data.get("summary"), "要約なし")

    lines.append(f"# {title}")
    lines.append("")
    lines.append("## 全体要約")
    lines.append("")
    lines.append(summary)
    lines.append("")
    lines.append("## トピック詳細")

    for index, topic in enumerate(_as_list(summary_data.get("topics")), start=1):
        lines.append("")
        lines.append(_format_topic_heading(topic, index))
        lines.append("")

        keywords = [f"- {_stringify(keyword)}" for keyword in _as_list(topic.get("topic_keywords"))]
        _append_block(lines, "キーワード", keywords)

        _append_block(lines, "要約", [_stringify(topic.get("topic_summary"), "要約なし")])

        points = [f"- {_stringify(point)}" for point in _as_list(topic.get("topic_points"))]
        _append_block(lines, "ポイント", points)

        knowledge_lines: list[str] = []
        for item in _as_list(topic.get("important_knowledge")):
            knowledge_lines.append(f"- {_stringify(item.get('statement'))}")
            knowledge_lines.append(f"  詳細: {_stringify(item.get('explanation'))}")
            knowledge_lines.append(f"  根拠: {_format_source_label(item.get('source'))}")

            importance = item.get("importance")
            if isinstance(importance, int) and importance > 0:
                knowledge_lines.append(f"  重要度: {'★' * importance}")

            evidence = _format_time_range(item.get("evidence"))
            if evidence:
                knowledge_lines.append(f"  タイムスタンプ: {evidence}")

        _append_block(lines, "重要な知見", knowledge_lines)

        technical_term_lines: list[str] = []
        for term in _as_list(topic.get("technical_term")):
            technical_term_lines.append(
                f"- {_stringify(term.get('word'))}: {_stringify(term.get('explanation'))}"
            )

            lecture_definition = _stringify(term.get("lecture_definition"))
            if lecture_definition:
                technical_term_lines.append(f"  講義での定義: {lecture_definition}")

            pitfall = _stringify(term.get("pitfall"))
            if pitfall:
                technical_term_lines.append(f"  誤解・混同ポイント: {pitfall}")

            evidence = _format_time_range(term.get("evidence"))
            if evidence:
                technical_term_lines.append(f"  タイムスタンプ: {evidence}")

        if technical_term_lines:
            _append_block(lines, "専門用語", technical_term_lines)

    assignments = _as_list(summary_data.get("assignments"))
    if assignments:
        lines.append("## 課題、提出物")
        lines.append("")

        for assignment in assignments:
            lines.append(
                f"- {_stringify(assignment.get('task_title'))}: {_stringify(assignment.get('task_detail'))}"
            )

            due = _stringify(
                assignment.get("due_datetime") or assignment.get("due_date")
            )
            if due:
                lines.append(f"  提出期限: {due}")

            evidence = _format_time_range(assignment.get("evidence"))
            if evidence:
                lines.append(f"  タイムスタンプ: {evidence}")

            lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def populate_word_document(doc: Any, summary_data: dict[str, Any]) -> None:
    doc.add_heading(_stringify(summary_data.get("title"), "タイトルなし"), 0)
    doc.add_heading("全体要約", level=1)
    doc.add_paragraph(_stringify(summary_data.get("summary"), "要約なし"))
    doc.add_heading("トピック詳細", level=1)

    for index, topic in enumerate(_as_list(summary_data.get("topics")), start=1):
        heading = _format_topic_heading(topic, index).removeprefix("### ")
        doc.add_heading(heading, level=2)

        doc.add_paragraph("キーワード")
        for keyword in _as_list(topic.get("topic_keywords")):
            doc.add_paragraph(_stringify(keyword), style="List Bullet")

        doc.add_paragraph("要約")
        doc.add_paragraph(_stringify(topic.get("topic_summary"), "要約なし"))

        doc.add_paragraph("ポイント")
        for point in _as_list(topic.get("topic_points")):
            doc.add_paragraph(_stringify(point), style="List Bullet")

        doc.add_paragraph("重要な知見")
        for item in _as_list(topic.get("important_knowledge")):
            doc.add_paragraph(_stringify(item.get("statement")), style="List Bullet")
            doc.add_paragraph(f"詳細: {_stringify(item.get('explanation'))}")
            doc.add_paragraph(f"根拠: {_format_source_label(item.get('source'))}")

            importance = item.get("importance")
            if isinstance(importance, int) and importance > 0:
                doc.add_paragraph(f"重要度: {'★' * importance}")

            evidence = _format_time_range(item.get("evidence"))
            if evidence:
                doc.add_paragraph(f"タイムスタンプ: {evidence}")

        technical_terms = _as_list(topic.get("technical_term"))
        if technical_terms:
            doc.add_paragraph("専門用語")
            for term in technical_terms:
                doc.add_paragraph(
                    f"{_stringify(term.get('word'))}: {_stringify(term.get('explanation'))}",
                    style="List Bullet",
                )

                lecture_definition = _stringify(term.get("lecture_definition"))
                if lecture_definition:
                    doc.add_paragraph(f"講義での定義: {lecture_definition}")

                pitfall = _stringify(term.get("pitfall"))
                if pitfall:
                    doc.add_paragraph(f"誤解・混同ポイント: {pitfall}")

                evidence = _format_time_range(term.get("evidence"))
                if evidence:
                    doc.add_paragraph(f"タイムスタンプ: {evidence}")

        doc.add_paragraph("")

    assignments = _as_list(summary_data.get("assignments"))
    if assignments:
        doc.add_heading("課題、提出物", level=1)
        for assignment in assignments:
            doc.add_paragraph(
                f"{_stringify(assignment.get('task_title'))}: {_stringify(assignment.get('task_detail'))}",
                style="List Bullet",
            )

            due = _stringify(
                assignment.get("due_datetime") or assignment.get("due_date")
            )
            if due:
                doc.add_paragraph(f"提出期限: {due}")

            evidence = _format_time_range(assignment.get("evidence"))
            if evidence:
                doc.add_paragraph(f"タイムスタンプ: {evidence}")
