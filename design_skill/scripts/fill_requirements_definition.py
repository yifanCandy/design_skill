#!/usr/bin/env python3
"""Render a 要件定義書 (.xlsx) from a content JSON file.

Usage:
    python fill_requirements_definition.py --content content.json --output out.xlsx

Chapters 1/2/4/5/6 use a fixed generic skeleton (see
templates/要件定義/content.schema.json); chapter 3 (the requirements table)
and chapter 5's diagram are the free-form/manual parts. See
scripts/design_doc_common.py for the shared house-style primitives.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_chapter_heading, emit_item_group, emit_numbered_subsection,
    emit_placeholder_note, emit_table_n, load_house_style, prepare_chapter_sheets,
    render_revision_history, render_toc, save_workbook, set_cell, sync_page_header,
)


def render_ch1(ws, content, styles):
    row = emit_chapter_heading(ws, 5, "本書の定義", styles)
    outline = []
    row = emit_numbered_subsection(ws, row, "1.1", "参照資料", content["references"], styles)
    outline.append(("1.1", "参照資料"))
    row = emit_numbered_subsection(ws, row, "1.2", "本書の目的", content["purpose"], styles)
    outline.append(("1.2", "本書の目的"))
    row = emit_numbered_subsection(ws, row, "1.3", "現状の課題", content["current_problems"], styles)
    outline.append(("1.3", "現状の課題"))
    set_cell(ws, row, "C", "1.4", styles["subsection_num"])
    set_cell(ws, row, "D", "用語定義", styles["subsection_title"])
    row += 1
    row = emit_item_group(ws, row, content["glossary"], styles, bullet=False)
    outline.append(("1.4", "用語定義"))
    return "本書の定義", outline


def render_ch2(ws, content, styles):
    row = emit_chapter_heading(ws, 5, "要件概要", styles)
    outline = []
    row = emit_numbered_subsection(ws, row, "2.1", "背景・目的", content["background"], styles)
    outline.append(("2.1", "背景・目的"))
    scope = [x if x.strip().startswith("・") else f"・{x}" for x in content["scope"]]
    row = emit_numbered_subsection(ws, row, "2.2", "対象範囲", scope, styles)
    outline.append(("2.2", "対象範囲"))
    set_cell(ws, row, "C", "2.3", styles["subsection_num"])
    set_cell(ws, row, "D", "ステークホルダ", styles["subsection_title"])
    row += 1
    row = emit_item_group(ws, row, content["stakeholders"], styles, bullet=False)
    outline.append(("2.3", "ステークホルダ"))
    return "要件概要", outline


def render_ch3(ws, content, styles):
    title = content["chapter_title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    row = emit_table_n(ws, row, content["headers"], content["rows"], styles)
    return title, []


def render_ch4(ws, content, styles):
    title = content["chapter_title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    categories = [
        ("4.1", "機能要件", content["functional"]),
        ("4.2", "性能要件", content["performance"]),
        ("4.3", "可用性要件", content["availability"]),
        ("4.4", "運用要件", content["operational"]),
        ("4.5", "セキュリティ要件", content["security"]),
        ("4.6", "保守性・拡張性要件", content["maintainability"]),
    ]
    for num, label, body in categories:
        row = emit_numbered_subsection(ws, row, num, label, body, styles)
        outline.append((num, label))
    return title, outline


def render_ch5(ws, content, styles):
    row = emit_chapter_heading(ws, 5, "利用イメージ", styles)
    outline = []
    set_cell(ws, row, "C", "5.1", styles["subsection_num"])
    set_cell(ws, row, "D", "業務フロー・画面イメージ", styles["subsection_title"])
    row += 1
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row = emit_placeholder_note(ws, row, f"別途図を貼付: {content['diagram_placeholder']}", styles)
    row += 1
    outline.append(("5.1", "業務フロー・画面イメージ"))

    set_cell(ws, row, "C", "5.2", styles["subsection_num"])
    set_cell(ws, row, "D", "利用シーン", styles["subsection_title"])
    row += 1
    row = emit_item_group(ws, row, content["use_case_steps"], styles, bullet=False)
    outline.append(("5.2", "利用シーン"))
    return "利用イメージ", outline


def render_ch6(ws, content, styles):
    title = content["chapter_title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    for line in content["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    row = emit_table_n(ws, row, content["headers"], content["rows"], styles, cols=["D", "H", "L"])
    return title, []


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    ch3_title = content["ch3_requirements_list"]["chapter_title"]
    ch4_title = content["ch4_nonfunctional"]["chapter_title"]
    ch6_title = content["ch6_schedule"]["chapter_title"]

    sheet_names = ["1.本書の定義", "2.要件概要", f"3.{ch3_title}", f"4.{ch4_title}", "5.利用イメージ", f"6.{ch6_title}"]
    ws1, ws2, ws3, ws4, ws5, ws6 = prepare_chapter_sheets(wb, sheet_names, include_toc=True)

    ch1_title, ch1_outline = render_ch1(ws1, content["ch1_definition"], styles)
    ch2_title, ch2_outline = render_ch2(ws2, content["ch2_overview"], styles)
    ch3_title, ch3_outline = render_ch3(ws3, content["ch3_requirements_list"], styles)
    ch4_title, ch4_outline = render_ch4(ws4, content["ch4_nonfunctional"], styles)
    ch5_title, ch5_outline = render_ch5(ws5, content["ch5_usage_image"], styles)
    ch6_title, ch6_outline = render_ch6(ws6, content["ch6_schedule"], styles)

    chapters = [
        (ch1_title, ch1_outline), (ch2_title, ch2_outline), (ch3_title, ch3_outline),
        (ch4_title, ch4_outline), (ch5_title, ch5_outline), (ch6_title, ch6_outline),
    ]
    render_toc(wb["目次"], chapters, styles)
    render_revision_history(wb["改定履歴"], meta, styles)

    all_sheet_names = ["改定履歴", "目次"] + [ws.title for ws in (ws1, ws2, ws3, ws4, ws5, ws6)]
    sync_page_header(wb, meta, all_sheet_names, title_label="要件定義書")

    save_workbook(wb, output_path, cover_lines(meta, "要件定義書"))


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
