#!/usr/bin/env python3
"""Render a テスト結果報告書 (.xlsx) from the code-generation skill's test-case catalog.

Usage:
    python fill_test_report.py --content tests.json --output out.xlsx

The 結果 column is read from the JUnit XML the test runner actually wrote
(content.project.junit_xml, see scripts/test_results.py), not from the catalog.
generate.py adds `_function_names` ({F-xx: 機能名}) from the design master before rendering.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from design_doc_common import (  # noqa: E402
    cover_lines, emit_chapter_heading, emit_item_group, emit_numbered_subsection, emit_table_n,
    load_house_style, prepare_chapter_sheets, render_revision_history, render_toc, save_workbook,
    set_cell, sync_page_header,
)
import test_results  # noqa: E402


def subsection(ws, row, num, title, styles):
    set_cell(ws, row, "C", num, styles["subsection_num"])
    set_cell(ws, row, "D", title, styles["subsection_title"])
    return row + 1


def render_ch1(ws, content, rows, run, names, styles):
    row = emit_chapter_heading(ws, 5, "テスト概要", styles)
    project = content["project"]
    row = emit_numbered_subsection(ws, row, "1.1", "テスト対象", [
        f"システム: {content['meta']['system_name']}",
        f"ソースコード: {project['dir']}",
        f"起動コマンド: {project['start_command']}" + (f"（アクセス URL: {project['url']}）" if project.get("url") else ""),
        *[f"・{t}" for t in project["tech_stack"]],
    ], styles)
    row = emit_numbered_subsection(ws, row, "1.2", "テスト環境・実施", [
        f"テストフレームワーク: {project['test_framework']}",
        f"実行コマンド: {project['test_command']}",
        f"実施日時: {run['executed_at'] or '未実施'}" + (f"（所要 {run['seconds']} 秒）" if run["seconds"] else ""),
        f"結果ファイル: {project['junit_xml']}（テストランナーが出力した JUnit XML。本書の結果欄はこのファイルから転記）",
    ], styles)

    row = subsection(ws, row, "1.3", "結果サマリ", styles)
    by_f = test_results.summary_by_function(rows)
    table = [[f, names.get(f, ""), s["total"], s["OK"], s["NG"], s["未実施"]] for f, s in sorted(by_f.items())]
    total = {k: sum(s[k] for s in by_f.values()) for k in ("total", "OK", "NG", "未実施")}
    table.append(["合計", "", total["total"], total["OK"], total["NG"], total["未実施"]])
    row = emit_table_n(ws, row, ["機能ID", "機能名", "件数", "OK", "NG", "未実施"], table, styles,
                       cols=["D", "H", "V", "Y", "AB", "AE"])

    row = subsection(ws, row, "1.4", "未検証項目", styles)
    untested = content["untested"]
    if untested:
        row = emit_table_n(ws, row, ["ID", "理由"], [[u["id"], u["reason"]] for u in untested], styles, cols=["D", "H"])
    else:
        set_cell(ws, row, "D", "なし（主表の全 IF・メッセージをいずれかのテストケースで検証した）", styles["body_line"])
        row += 2
    return "テスト概要", [("1.1", "テスト対象"), ("1.2", "テスト環境・実施"), ("1.3", "結果サマリ"), ("1.4", "未検証項目")]


def render_ch2(ws, rows, styles):
    row = emit_chapter_heading(ws, 5, "テスト結果一覧", styles)
    table = [[str(i), r["id"], r["function"], r["category"], r["result"], r["viewpoint"]]
             for i, r in enumerate(rows, start=1)]
    emit_table_n(ws, row + 1, ["No", "テストID", "機能", "分類", "結果", "テスト観点"], table, styles,
                 cols=["C", "E", "J", "M", "P", "S"])
    return "テスト結果一覧", []


def render_ch3(ws, rows, styles):
    row = emit_chapter_heading(ws, 5, "テストケース詳細", styles)
    items = []
    for r in rows:
        body = [f"対象: {', '.join(r['targets'])}"]
        body += ["前提条件:"] + [f"　{p}" for p in r["precondition"] or ["なし"]]
        body += ["手順:"] + [f"　{i}. {s}" for i, s in enumerate(r["steps"], start=1)]
        body += ["期待結果:"] + [f"　・{e}" for e in r["expected"]]
        body.append(f"結果: {r['result']}" + (f"（{r['message']}）" if r["message"] else ""))
        items.append({"title": f"{r['id']}［{r['category']}］{r['viewpoint']}", "body": body})
    emit_item_group(ws, row + 1, items, styles)
    return "テストケース詳細", []


def render_ch4(ws, review, styles):
    row = emit_chapter_heading(ws, 5, "設計書レビュー結果", styles)
    set_cell(ws, row, "D", "実装前に詳細仕様書・詳細設計書をレビューした結果と、実装での扱いを示す。", styles["body_line"])
    items = [{"title": f"No.{i}［{r['severity']}］{r['document']}", "body": [f"指摘: {r['finding']}", f"対応: {r['action']}"]}
             for i, r in enumerate(review, start=1)]
    emit_item_group(ws, row + 2, items, styles)
    return "設計書レビュー結果", []


def render(content: dict, output_path: Path):
    wb, styles = load_house_style()
    meta = content["meta"]
    rows, run = test_results.results(content)
    names = content.get("_function_names", {})

    sheet_names = ["1.テスト概要", "2.テスト結果一覧", "3.テストケース詳細", "4.設計書レビュー結果"]
    ws1, ws2, ws3, ws4 = prepare_chapter_sheets(wb, sheet_names, include_toc=True)
    chapters = [
        render_ch1(ws1, content, rows, run, names, styles),
        render_ch2(ws2, rows, styles),
        render_ch3(ws3, rows, styles),
        render_ch4(ws4, content["review"], styles),
    ]
    render_toc(wb["目次"], chapters, styles)
    render_revision_history(wb["改定履歴"], meta, styles)

    all_sheet_names = ["改定履歴", "目次"] + [ws.title for ws in (ws1, ws2, ws3, ws4)]
    sync_page_header(wb, meta, all_sheet_names, title_label="テスト結果報告書")
    save_workbook(wb, output_path, cover_lines(meta, "テスト結果報告書"))


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
