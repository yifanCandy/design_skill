#!/usr/bin/env python3
"""Render a 基本設計書 (.xlsx) from a content JSON file.

Usage:
    python fill_basic_design.py --content content.json --output out.xlsx

See scripts/design_doc_common.py for the shared house-style rendering
primitives (reused across all 設計書 categories).
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_chapter_heading, emit_item_group, emit_numbered_subsection, emit_table2,
    emit_table_n, ensure_bullet, load_house_style, prepare_chapter_sheets, render_revision_history,
    render_toc, save_workbook, set_cell, sync_page_header,
)

CHAPTER_SHEET_NAMES = [
    "1.本書の定義", "2.機能の概要", "3.ログ出力機構要件",
    "4.ログ出力機構の機能", "5.機能詳細", "6.ログ仕様",
]


# ---------------------------------------------------------------------------
# Per-sheet renderers. Each returns an "outline" entry used later for 目次:
# (chapter_title, [(sub_num, sub_title), ...])
# ---------------------------------------------------------------------------

def render_ch1(ws, content, function_name, styles):
    row = emit_chapter_heading(ws, 5, "本書の定義", styles)
    outline = []

    row = emit_numbered_subsection(ws, row, "1.1", "参照資料", content["references"], styles)
    outline.append(("1.1", "参照資料"))

    row = emit_numbered_subsection(ws, row, "1.2", "本書の目的", content["purpose"], styles)
    outline.append(("1.2", "本書の目的"))

    problem_title = f"{function_name}設計により解決する課題"
    row = emit_numbered_subsection(ws, row, "1.3", problem_title, content["problems"], styles)
    outline.append(("1.3", problem_title))

    usage_title = f"{function_name}の利用目的"
    set_cell(ws, row, "C", "1.4", styles["subsection_num"])
    set_cell(ws, row, "D", usage_title, styles["subsection_title"])
    row += 1
    row = emit_item_group(ws, row, content["usage_purposes"], styles, bullet=False)
    outline.append(("1.4", usage_title))

    return "本書の定義", outline


def render_ch2(ws, content, styles):
    row = emit_chapter_heading(ws, 5, "機能の概要", styles)
    outline = []

    row = emit_numbered_subsection(ws, row, "2.1", "機能概要", content["overview"], styles)
    outline.append(("2.1", "機能概要"))

    use_cases = [ensure_bullet(x) for x in content["use_cases"]]
    row = emit_numbered_subsection(ws, row, "2.2", "ユースケース", use_cases, styles)
    outline.append(("2.2", "ユースケース"))

    row = emit_numbered_subsection(ws, row, "2.3", "処理概要", content["process_overview"], styles)
    outline.append(("2.3", "処理概要"))

    return "機能の概要", outline


def render_ch3(ws, content, styles):
    title = content["chapter_title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    categories = [
        ("3.1", "機能要件", content["functional"]),
        ("3.2", "性能要件", content["performance"]),
        ("3.3", "可用性要件", content["availability"]),
        ("3.4", "運用要件", content["operational"]),
        ("3.5", "セキュリティ要件", content["security"]),
        ("3.6", "保守性・拡張性要件", content["maintainability"]),
    ]
    for num, label, body in categories:
        row = emit_numbered_subsection(ws, row, num, label, body, styles)
        outline.append((num, label))
    return title, outline


def render_ch4(ws, content, styles):
    title = content["chapter_title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []
    for i, item in enumerate(content["items"], start=1):
        num = f"4.{i}"
        row = emit_numbered_subsection(ws, row, num, item["title"], item["body"], styles)
        outline.append((num, item["title"]))
    return title, outline


def render_ch5(ws, content, styles, chapter_num="5"):
    row = emit_chapter_heading(ws, 5, "機能詳細", styles)
    outline = []
    for i, section in enumerate(content, start=1):
        num = f"{chapter_num}.{i}"
        title = section["title"]
        set_cell(ws, row, "C", num, styles["subsection_num"])
        set_cell(ws, row, "D", title, styles["subsection_title"])
        row += 1
        for line in section.get("intro", []):
            set_cell(ws, row, "D", line, styles["body_line"])
            row += 1
        if section["type"] == "list":
            row = emit_item_group(ws, row, section["items"], styles, bullet=True)
        else:  # table2col
            row = emit_table2(ws, row, section["col_headers"], section["rows"], styles)
        outline.append((num, title))
    return "機能詳細", outline


def render_ch6(ws, content, styles):
    title = content["title"]
    row = emit_chapter_heading(ws, 5, title, styles)
    outline = []

    set_cell(ws, row, "C", "6.1", styles["subsection_num"])
    set_cell(ws, row, "D", "出力形式", styles["subsection_title"])
    row += 1
    for line in content["format_section"]["intro"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row = emit_item_group(
        ws, row,
        [{"title": "出力形式", "body": [content["format_section"]["format_string"]]}],
        styles, bullet=True,
    )
    outline.append(("6.1", "出力形式"))

    set_cell(ws, row, "C", "6.2", styles["subsection_num"])
    set_cell(ws, row, "D", "項目説明", styles["subsection_title"])
    row += 1
    row = emit_table_n(ws, row, content["items_table"]["headers"], content["items_table"]["rows"], styles, cols=["D", "E", "K"])
    outline.append(("6.2", "項目説明"))

    set_cell(ws, row, "C", "6.3", styles["subsection_num"])
    set_cell(ws, row, "D", "フォーマット例", styles["subsection_title"])
    row += 1
    examples = [{"title": ex["label"], "body": [ex["text"]]} for ex in content["examples"]]
    row = emit_item_group(ws, row, examples, styles, bullet=True)
    outline.append(("6.3", "フォーマット例"))

    set_cell(ws, row, "C", "6.4", styles["subsection_num"])
    set_cell(ws, row, "D", "補足事項", styles["subsection_title"])
    row += 1
    for line in content["notes"]:
        set_cell(ws, row, "D", line, styles["body_line"])
        row += 1
    row += 1
    outline.append(("6.4", "補足事項"))

    return title, outline


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()

    meta = content["meta"]
    function_name = meta["function_name"]
    ch3_title = content["ch3_requirements"]["chapter_title"]
    ch4_title = content["ch4_functions"]["chapter_title"]
    ch6_title = content["ch6_spec"]["title"]

    sheet_names = [
        "1.本書の定義", "2.機能の概要", f"3.{ch3_title}",
        f"4.{ch4_title}", "5.機能詳細", f"6.{ch6_title}",
    ]
    ws1, ws2, ws3, ws4, ws5, ws6 = prepare_chapter_sheets(wb, sheet_names, include_toc=True)

    ch1_title, ch1_outline = render_ch1(ws1, content["ch1_definition"], function_name, styles)
    ch2_title, ch2_outline = render_ch2(ws2, content["ch2_overview"], styles)
    ch3_title, ch3_outline = render_ch3(ws3, content["ch3_requirements"], styles)
    ch4_title, ch4_outline = render_ch4(ws4, content["ch4_functions"], styles)
    ch5_title, ch5_outline = render_ch5(ws5, content["ch5_details"], styles)
    ch6_title, ch6_outline = render_ch6(ws6, content["ch6_spec"], styles)

    chapters = [
        (ch1_title, ch1_outline), (ch2_title, ch2_outline), (ch3_title, ch3_outline),
        (ch4_title, ch4_outline), (ch5_title, ch5_outline), (ch6_title, ch6_outline),
    ]
    render_toc(wb["目次"], chapters, styles)
    render_revision_history(wb["改定履歴"], meta, styles)

    all_sheet_names = ["改定履歴", "目次"] + [ws.title for ws in (ws1, ws2, ws3, ws4, ws5, ws6)]
    sync_page_header(wb, meta, all_sheet_names, title_label=f"{function_name}仕様書")

    save_workbook(wb, output_path, cover_lines(meta, "基本設計書"))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--content", required=True, type=Path, help="content JSON file matching content.schema.json")
    parser.add_argument("--output", required=True, type=Path, help="output .xlsx path")
    args = parser.parse_args()

    content = json.loads(args.content.read_text(encoding="utf-8"))
    render(content, args.output)
    print(f"wrote {args.output}")


if __name__ == "__main__":
    sys.exit(main())
