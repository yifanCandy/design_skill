#!/usr/bin/env python3
"""Render an IF定義書 (.xlsx) from a content JSON file.

Usage:
    python fill_if_definition.py --content content.json --output out.xlsx

One sheet per scenario, no 目次 (per-screen docs are short). See
templates/詳細設計/if_definition.schema.json for the content contract and
scripts/design_doc_common.py for the shared house-style primitives.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_section_mark, emit_sub_mark, emit_table_n, load_house_style,
    prepare_chapter_sheets, render_revision_history, save_workbook, set_cell, sync_page_header,
)


def render_scenario(ws, scenario, styles):
    row = emit_section_mark(ws, 5, "処理", styles)
    row = emit_sub_mark(ws, row, "処理名", [scenario["process_name"]], styles)

    row = emit_section_mark(ws, row, "リクエスト", styles)
    comm = scenario["communication"]
    row = emit_sub_mark(ws, row, "通信", [f"タイプ: {comm['type']}", f"URL: {comm['url']}"], styles)
    for tbl in scenario["request_tables"]:
        set_cell(ws, row, "D", f"○{tbl['label']}", styles["item_title"])
        row += 1
        row = emit_table_n(ws, row, tbl["headers"], tbl["rows"], styles)

    row = emit_section_mark(ws, row, "レスポンス", styles)
    row = emit_sub_mark(ws, row, "正常時", scenario["response_success"], styles)
    for tbl in scenario["response_tables"]:
        set_cell(ws, row, "D", f"○{tbl['label']}", styles["item_title"])
        row += 1
        row = emit_table_n(ws, row, tbl["headers"], tbl["rows"], styles)


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    scenarios = content["if_definition"]["scenarios"]
    scenario_names = [s["name"] for s in scenarios]
    sheets = prepare_chapter_sheets(wb, scenario_names, include_toc=False)

    for ws, scenario in zip(sheets, scenarios):
        render_scenario(ws, scenario, styles)

    render_revision_history(wb["改定履歴"], meta, styles)
    all_sheet_names = ["改定履歴"] + [ws.title for ws in sheets]
    sync_page_header(wb, meta, all_sheet_names, title_label="I/F定義書")

    save_workbook(wb, output_path, cover_lines(meta, "I/F定義書"))


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
