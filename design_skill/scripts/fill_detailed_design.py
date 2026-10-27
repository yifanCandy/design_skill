#!/usr/bin/env python3
"""Render a 詳細設計書本体 (.xlsx) from a content JSON file (the
"detailed_design" half of templates/詳細設計/content.schema.json — see
fill_if_definition.py for the other half, "if_definition").

Usage:
    python fill_detailed_design.py --content content.json --output out.xlsx

One sheet per sub-function, each a sequence of ①②③... step blocks (a step
title plus an ordered mix of text lines and tables) — no visual box border
around each step (a deliberate v1 simplification; spacing + numbering carry
the separation instead). A final fixed "3.メッセージID一覧" sheet closes the
workbook. No 目次 (per-screen docs are short). シーケンス図 is out of scope.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_section_mark, emit_table_n, load_house_style, prepare_chapter_sheets,
    render_revision_history, save_workbook, set_cell, sync_page_header,
)

MESSAGE_LIST_SHEET = "3.メッセージID一覧"


def render_function_sheet(ws, function, styles):
    row = 5
    for step in function["steps"]:
        set_cell(ws, row, "D", f"{step['marker']}{step['title']}", styles["item_title"])
        row += 1
        for part in step["parts"]:
            if part["type"] == "text":
                for line in part["lines"]:
                    set_cell(ws, row, "E", line, styles["item_body"])
                    row += 1
            else:  # table
                row = emit_table_n(ws, row, part["headers"], part["rows"], styles)
        row += 1  # spacer between steps


def render_message_list(ws, message_list, styles):
    row = emit_section_mark(ws, 5, "メッセージID一覧", styles)
    row = emit_table_n(ws, row, message_list["headers"], message_list["rows"], styles)


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    dd = content["detailed_design"]
    function_names = [f["name"] for f in dd["functions"]] + [MESSAGE_LIST_SHEET]
    sheets = prepare_chapter_sheets(wb, function_names, include_toc=False)

    *function_sheets, message_sheet = sheets
    for ws, function in zip(function_sheets, dd["functions"]):
        render_function_sheet(ws, function, styles)
    render_message_list(message_sheet, dd["message_list"], styles)

    render_revision_history(wb["改定履歴"], meta, styles)
    all_sheet_names = ["改定履歴"] + [ws.title for ws in sheets]
    sync_page_header(wb, meta, all_sheet_names, title_label="詳細設計書")

    save_workbook(wb, output_path, cover_lines(meta, "詳細設計書"))


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
