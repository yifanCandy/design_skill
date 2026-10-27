#!/usr/bin/env python3
"""Validate a content JSON against a content.schema.json.

Usage: python validate_content.py content.json [--schema path/to/content.schema.json]
Defaults to templates/基本設計/content.schema.json when --schema is omitted.
Exits non-zero and prints the validation errors if the content doesn't match.
"""
import argparse
import json
import sys
from pathlib import Path

import jsonschema

DEFAULT_SCHEMA_PATH = Path(__file__).resolve().parent.parent / "templates" / "基本設計" / "content.schema.json"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("content", type=Path)
    parser.add_argument("--schema", type=Path, default=DEFAULT_SCHEMA_PATH)
    args = parser.parse_args()

    content = json.loads(args.content.read_text(encoding="utf-8"))
    schema = json.loads(args.schema.read_text(encoding="utf-8"))

    validator = jsonschema.Draft7Validator(schema)
    errors = sorted(validator.iter_errors(content), key=lambda e: list(e.absolute_path))
    if not errors:
        print("OK: content matches content.schema.json")
        return 0

    for err in errors:
        path = "/".join(str(p) for p in err.absolute_path) or "(root)"
        print(f"- {path}: {err.message}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    sys.exit(main())
