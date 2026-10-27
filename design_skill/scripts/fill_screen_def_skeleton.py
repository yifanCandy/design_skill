#!/usr/bin/env python3
"""Render a skeleton-only 画面定義書 (.xlsx): page header + one placeholder
sheet per named UI widget. LOW automation value by design — see
templates/詳細仕様/screen_def_skeleton.schema.json for why. The actual widget
behavior, config keys and screenshots are left for manual authoring.

Usage:
    python fill_screen_def_skeleton.py --content content.json --output out.xlsx
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_placeholder_note, load_house_style, prepare_chapter_sheets,
    render_revision_history, sanitize_sheet_title, save_workbook, sync_page_header,
)


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    widget_sheet_names = [sanitize_sheet_title(w) for w in content["widgets"]]
    sheets = prepare_chapter_sheets(wb, widget_sheet_names, include_toc=False)

    for ws, widget in zip(sheets, content["widgets"]):
        emit_placeholder_note(ws, 5, f"「{widget}」の仕様は別途手動で記入してください（v1では骨格のみ自動生成）", styles)

    render_revision_history(wb["改定履歴"], meta, styles)
    all_sheet_names = ["改定履歴"] + [ws.title for ws in sheets]
    sync_page_header(wb, meta, all_sheet_names, title_label="画面定義書")

    save_workbook(wb, output_path, cover_lines(meta, "画面定義書"))


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
