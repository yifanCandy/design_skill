#!/usr/bin/env python3
"""Render a ツール仕様書 (.xlsx, the high-automation half of 詳細仕様) from a
content JSON file. Chapters 1/2/7 reuse fill_basic_design's ch1/ch2/ch5
renderers verbatim (structurally identical); ch3/ch4 are the same shape but
with fixed chapter titles; ch5/ch6 have screenshot placeholders; ch8/9 are
new. See templates/詳細仕様/tool_spec.schema.json for the content contract.

Usage:
    python fill_tool_spec.py --content content.json --output out.xlsx
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_chapter_heading, emit_item_group, emit_numbered_subsection,
    emit_placeholder_note, emit_table2, emit_table_n, load_house_style, prepare_chapter_sheets,
    render_revision_history, render_toc, save_workbook, set_cell, sync_page_header,
)
from fill_basic_design import render_ch1 as render_ch1_shared  # noqa: E402
from fill_basic_design import render_ch2 as render_ch2_shared  # noqa: E402
from fill_basic_design import render_ch5 as render_ch7_shared  # noqa: E402

CHAPTER_TITLES = ["本書の定義", "機能の概要", "ツール要件", "ツール機能", "操作イメージ", "画面定義", "機能詳細", "プリファレンス", "エラーメッセージ一覧"]


def render_ch3(ws, content, styles):
    title = "ツール要件"
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    categories = [
        ("3.1", "機能要件", content["functional"]), ("3.2", "性能要件", content["performance"]),
        ("3.3", "可用性要件", content["availability"]), ("3.4", "運用要件", content["operational"]),
        ("3.5", "セキュリティ要件", content["security"]), ("3.6", "保守性・拡張性要件", content["maintainability"]),
    ]
    for num, label, body in categories:
        row = emit_numbered_subsection(ws, row, num, label, body, styles)
        outline.append((num, label))
    return title, outline


def render_ch4(ws, content, styles):
    title = "ツール機能"
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    for i, item in enumerate(content["items"], start=1):
        num = f"4.{i}"
        row = emit_numbered_subsection(ws, row, num, item["title"], item["body"], styles)
        outline.append((num, item["title"]))
    return title, outline


def render_ch5(ws, content, styles):
    title = "操作イメージ"
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    for i, step in enumerate(content["steps"], start=1):
        num = f"5.{i}"
        row = emit_numbered_subsection(ws, row, num, step["title"], step["body"], styles)
        row = emit_placeholder_note(ws, row - 1, f"別途スクリーンショットを貼付: {step['screenshot_note']}", styles)
        row += 1
        outline.append((num, step["title"]))
    return title, outline


def render_ch6(ws, content, styles):
    title = "画面定義"
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    for i, screen in enumerate(content["screens"], start=1):
        num = f"6.{i}"
        set_cell(ws, row, "C", num, styles["subsection_num"])
        set_cell(ws, row, "D", screen["name"], styles["subsection_title"])
        row += 1
        row = emit_placeholder_note(ws, row, f"別途スクリーンショットを貼付: {screen['screenshot_note']}", styles)
        row += 1
        set_cell(ws, row, "D", "画面制御", styles["item_title"])
        row += 1
        row = emit_table2(ws, row, ["属性", "値"], screen["screen_control"], styles)
        set_cell(ws, row, "D", "項目定義", styles["item_title"])
        row += 1
        row = emit_table_n(ws, row, screen["items_headers"], screen["items_rows"], styles)
        outline.append((num, screen["name"]))
    return title, outline


def render_ch8(ws, content, styles):
    title = "プリファレンス"
    row = emit_chapter_heading(ws, 5, title, styles)
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    emit_table_n(ws, row, content["headers"], content["rows"], styles)
    return title, []


def render_ch9(ws, content, styles):
    title = "エラーメッセージ一覧"
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    for i, section in enumerate(content["sections"], start=1):
        num = f"9.{i}"
        set_cell(ws, row, "C", num, styles["subsection_num"])
        set_cell(ws, row, "D", section["title"], styles["subsection_title"])
        row += 1
        if section["has_messages"]:
            row = emit_table_n(ws, row, section["headers"], section["rows"], styles)
        else:
            set_cell(ws, row, "D", "※なし", styles["body_line"])
            row += 2
        outline.append((num, section["title"]))
    return title, outline


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    function_name = meta["function_name"]

    sheet_names = [f"{i+1}.{t}" for i, t in enumerate(CHAPTER_TITLES)]
    sheets = prepare_chapter_sheets(wb, sheet_names, include_toc=True)
    ws1, ws2, ws3, ws4, ws5, ws6, ws7, ws8, ws9 = sheets

    ch1_title, ch1_outline = render_ch1_shared(ws1, content["ch1_definition"], function_name, styles)
    ch2_title, ch2_outline = render_ch2_shared(ws2, content["ch2_overview"], styles)
    ch3_title, ch3_outline = render_ch3(ws3, content["ch3_requirements"], styles)
    ch4_title, ch4_outline = render_ch4(ws4, content["ch4_functions"], styles)
    ch5_title, ch5_outline = render_ch5(ws5, content["ch5_usage_image"], styles)
    ch6_title, ch6_outline = render_ch6(ws6, content["ch6_screen_definition"], styles)
    ch7_title, ch7_outline = render_ch7_shared(ws7, content["ch7_details"], styles, chapter_num="7")
    ch8_title, ch8_outline = render_ch8(ws8, content["ch8_preferences"], styles)
    ch9_title, ch9_outline = render_ch9(ws9, content["ch9_error_messages"], styles)

    chapters = [
        (ch1_title, ch1_outline), (ch2_title, ch2_outline), (ch3_title, ch3_outline),
        (ch4_title, ch4_outline), (ch5_title, ch5_outline), (ch6_title, ch6_outline),
        (ch7_title, ch7_outline), (ch8_title, ch8_outline), (ch9_title, ch9_outline),
    ]
    render_toc(wb["目次"], chapters, styles)
    render_revision_history(wb["改定履歴"], meta, styles)

    all_sheet_names = ["改定履歴", "目次"] + [ws.title for ws in sheets]
    sync_page_header(wb, meta, all_sheet_names, title_label="ツール仕様書")

    save_workbook(wb, output_path, cover_lines(meta, "ツール仕様書"))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    content = json.loads(args.content.read_text(encoding="utf-8"))
    render(content, args.output)
    print(f"wrote {args.output}")


if __name__ == "__main__":
    sys.exit(main())
