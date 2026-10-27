#!/usr/bin/env python3
"""Validate a content JSON and render the .xlsx deliverable(s) for one category, in one step.

This is the single command every skill runs after drafting content.json:
  1. pops the top-level `_notes_for_review` (not part of any schema) and keeps it aside
  2. validates the rest against the category's schema (exit 1 + error list on failure)
  3. if master.json exists two levels up from the content (<project dir>/<システム名>/_design/<skill>/<ID>.json),
     checks that every ID the document uses is in the master and that it mentions the IDs it must
     (scripts/traceability.py; exit 1 + error list on failure), that the upstream documents exist —
     要件定義 → 基本設計 → 詳細仕様 → 詳細設計 → コード・テスト (exit 3 + the missing ones), and that the
     content carried over from the master and the upstream documents matches (scripts/consistency.py; exit 1);
     for test-report, also that the JUnit XML matches the test-case catalog and covers every
     function, IF and message (scripts/test_results.py; exit 1)
  4. resolves output path(s) — <project dir>/<システム名>/<category dir>/, file name from meta.function_name
     plus a _YYYYMMDDHHMMSS suffix, so every run writes new files and never overwrites an earlier one
  5. renders, prints the written paths and the review notes, and refreshes the system's
     トレーサビリティマトリクス.xlsx when a master exists

Usage:
    python generate.py <category> --content content.json [--out DIR_OR_XLSX]

Categories: requirements-definition, basic-design, detailed-spec, screen-def, detailed-design, test-report
Exit codes: 0 written, 1 schema/trace error, 3 upstream document missing
"""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

import jsonschema

SCRIPTS = Path(__file__).resolve().parent
TEMPLATES = SCRIPTS.parent / "templates"
sys.path.insert(0, str(SCRIPTS))

import consistency  # noqa: E402
import traceability as tracing  # noqa: E402

import fill_basic_design  # noqa: E402
import fill_detailed_design  # noqa: E402
import fill_if_definition  # noqa: E402
import fill_requirements_definition  # noqa: E402
import fill_screen_def_skeleton  # noqa: E402
import fill_test_report  # noqa: E402
import fill_tool_spec  # noqa: E402
import test_results  # noqa: E402

# category -> (schema, output dir under <project dir>/<システム名>/, [(file name suffix, render function), ...])
CATEGORIES = {
    "basic-design": (
        TEMPLATES / "基本設計" / "content.schema.json", "03基本設計",
        [("", fill_basic_design.render)],
    ),
    "requirements-definition": (
        TEMPLATES / "要件定義" / "content.schema.json", "02要件定義",
        [("", fill_requirements_definition.render)],
    ),
    "detailed-spec": (
        TEMPLATES / "詳細仕様" / "tool_spec.schema.json", "04詳細仕様",
        [("_ツール仕様書", fill_tool_spec.render)],
    ),
    "screen-def": (
        TEMPLATES / "詳細仕様" / "screen_def_skeleton.schema.json", "04詳細仕様",
        [("_画面定義書", fill_screen_def_skeleton.render)],
    ),
    "detailed-design": (
        TEMPLATES / "詳細設計" / "content.schema.json", "05詳細設計",
        [("_IF定義書", fill_if_definition.render), ("_詳細設計書", fill_detailed_design.render)],
    ),
    "test-report": (
        TEMPLATES / "テスト結果" / "test_report.schema.json", "06テスト結果",
        [("_テスト結果報告書", fill_test_report.render)],
    ),
}


def resolve_outputs(out, default_dir, function_name, renderers):
    """An explicit .xlsx path is used as is; otherwise <dir>/<function_name><suffix>_<timestamp>.xlsx
    (one timestamp for all files of the run)."""
    if out is not None and out.suffix.lower() == ".xlsx":
        if len(renderers) != 1:
            sys.exit(f"error: this category writes {len(renderers)} files; pass a directory to --out, not an .xlsx path")
        return [(out, renderers[0][1])]
    out_dir = out if out is not None else default_dir
    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
    return [(out_dir / f"{function_name}{suffix}_{stamp}.xlsx", render) for suffix, render in renderers]


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("category", choices=CATEGORIES)
    parser.add_argument("--content", required=True, type=Path)
    parser.add_argument("--out", type=Path, help="output directory, or an .xlsx path for single-file categories")
    args = parser.parse_args()

    schema_path, out_folder, renderers = CATEGORIES[args.category]
    content = json.loads(args.content.read_text(encoding="utf-8"))
    notes = content.pop("_notes_for_review", None)

    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    errors = sorted(jsonschema.Draft7Validator(schema).iter_errors(content), key=lambda e: list(e.absolute_path))
    if errors:
        print(f"VALIDATION FAILED against {schema_path.relative_to(SCRIPTS.parent)}:", file=sys.stderr)
        for err in errors:
            path = "/".join(str(p) for p in err.absolute_path) or "(root)"
            print(f"- {path}: {err.message}", file=sys.stderr)
        return 1

    master_path = args.content.resolve().parent.parent / "master.json"
    master = tracing.load_master(master_path) if master_path.exists() else None
    if master is not None:
        doc_dir, target = args.content.resolve().parent.name, args.content.stem
        trace_errors = tracing.check_master(master) or tracing.check_content(master, doc_dir, target, content)
        if trace_errors:
            print(f"TRACE FAILED against {master_path}:", file=sys.stderr)
            for err in trace_errors:
                print(f"- {err}", file=sys.stderr)
            return 1
        missing = tracing.missing_prereqs(master, master_path.parent, doc_dir, target)
        if missing:
            print("PREREQUISITE MISSING (要件定義 → 基本設計 → 詳細仕様 → 詳細設計 → コード・テスト の順に作成する):", file=sys.stderr)
            for label in missing:
                print(f"- {label}", file=sys.stderr)
            return 3
        content_errors = consistency.check(master, master_path.parent, doc_dir, target, content)
        if content_errors:
            print("CONSISTENCY FAILED (主表・上流の文書と内容が一致しない。この文書を上流に合わせて直す):", file=sys.stderr)
            for err in content_errors:
                print(f"- {err}", file=sys.stderr)
            return 1
        if args.category == "test-report":
            test_errors = test_results.check(master, content)
            if test_errors:
                print("TEST REPORT FAILED (テストケース台帳と JUnit XML の照合):", file=sys.stderr)
                for err in test_errors:
                    print(f"- {err}", file=sys.stderr)
                return 1
            content["_function_names"] = {f["id"]: f["name"] for f in master["functions"]}

    default_dir = tracing.doc_root() / content["meta"]["system_name"] / out_folder
    outputs = resolve_outputs(args.out, default_dir, content["meta"]["function_name"], renderers)
    for path, render in outputs:
        render(content, path)
        print(f"WROTE {path}")

    if args.category == "test-report":
        rows, run = test_results.results(content)
        counts = {k: sum(r["result"] == k for r in rows) for k in ("OK", "NG", "未実施")}
        print(f"TEST_RESULTS total={len(rows)} OK={counts['OK']} NG={counts['NG']} 未実施={counts['未実施']} executed_at={run['executed_at']}")
        for r in rows:
            if r["result"] != "OK":
                print(f"TEST_{'NG' if r['result'] == 'NG' else 'NOT_RUN'} {r['id']} {r['viewpoint']}: {r['message']}")

    if notes:
        print("NOTES_FOR_REVIEW:")
        for note in notes if isinstance(notes, list) else [notes]:
            print(f"- {note}")

    if master is not None:
        matrix = tracing.default_matrix_path(master)
        issues = tracing.write_matrix(master, master_path.parent, matrix)
        print(f"MATRIX {matrix}")
        pending = [f"{target}" for kind, target, _ in issues if kind == "未作成"]
        if pending:
            print(f"NOT_YET_CREATED: {', '.join(pending)}")
        for kind, target, msg in issues:
            if kind != "未作成":
                print(f"TRACE_ISSUE {target}: {msg}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
