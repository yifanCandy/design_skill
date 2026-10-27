"""Actual test results for the テスト結果報告書: reads the JUnit XML the test runner wrote and
matches its test cases to the test-case catalog (_design/code-generation/tests.json) by ID.

The OK/NG/未実施 of every case comes only from the runner's JUnit XML, never from the catalog,
so the report records what was actually executed. check() is what generate.py runs before
rendering; results() is what the renderer and the traceability matrix display.
"""
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

TC_RE = re.compile(r"TC-\d{2}-\d{3}")
MIN_CASES_PER_FUNCTION = 3


def junit_path(content):
    return Path(content["project"]["dir"]) / content["project"]["junit_xml"]


def parse_junit(path):
    """{TC-ID: {"result": OK/NG/未実施, "message": str}} plus run info (executed_at, seconds, unmatched names).
    A TC-ID used by several test functions (e.g. parameterized) is NG if any of them failed."""
    root = ET.parse(path).getroot()
    by_id, unmatched = {}, []
    rank = {"OK": 0, "未実施": 1, "NG": 2}
    for tc in root.iter("testcase"):
        name = f"{tc.get('classname', '')} {tc.get('name', '')}".strip()
        failed = tc.find("failure") if tc.find("failure") is not None else tc.find("error")
        if failed is not None:
            result, message = "NG", (failed.get("message") or failed.text or "").strip().splitlines()[:1]
        elif tc.find("skipped") is not None:
            result, message = "未実施", ["skipped"]
        else:
            result, message = "OK", []
        ids = set(TC_RE.findall(name))
        if not ids:
            unmatched.append(name)
        for tc_id in ids:
            prev = by_id.get(tc_id)
            if prev is None or rank[result] > rank[prev["result"]]:
                by_id[tc_id] = {"result": result, "message": message[0] if message else ""}

    suites = [root] if root.tag == "testsuite" else list(root.iter("testsuite"))
    stamps = [s.get("timestamp") for s in [root, *suites] if s.get("timestamp")]
    executed_at = stamps[0] if stamps else datetime.fromtimestamp(Path(path).stat().st_mtime).isoformat(timespec="seconds")
    seconds = root.get("time") or (f"{sum(float(s.get('time') or 0) for s in suites):.1f}" if suites else "")
    return by_id, {"executed_at": executed_at, "seconds": seconds, "unmatched": unmatched}


def check(master, content):
    """Errors that block the report (empty list = OK to render)."""
    errors = []
    cases = content["cases"]
    ids = [c["id"] for c in cases]
    errors += [f"テストID {i} が重複している" for i in sorted({i for i in ids if ids.count(i) > 1})]
    for c in cases:
        if c["id"][3:5] != c["function"][2:]:
            errors.append(f"{c['id']}: ID の機能番号が function {c['function']} と一致しない（TC-<F番号>-<連番>）")

    for f in master["functions"]:
        mine = [c for c in cases if c["function"] == f["id"]]
        cats = {c["category"] for c in mine}
        if len(mine) < MIN_CASES_PER_FUNCTION or "正常系" not in cats or "異常系" not in cats:
            errors.append(f"{f['id']} {f['name']}: テストケースが不足（{len(mine)}件。"
                          f"{MIN_CASES_PER_FUNCTION}件以上、正常系と異常系を各1件以上）")

    targeted = {t for c in cases for t in c["targets"]}
    excused = {u["id"] for u in content["untested"]}
    for key in ("interfaces", "messages"):
        for item in master[key]:
            if item["id"] not in targeted and item["id"] not in excused:
                errors.append(f"{item['id']}: どのテストケースでも検証していない（cases[].targets に入れるか、untested に理由を書く）")

    path = junit_path(content)
    if not path.exists():
        return errors + [f"JUnit XML が無い: {path}（test_command を実行してから生成する）"]
    by_id, run = parse_junit(path)
    errors += [f"{i}: 台帳にあるがテストコードに無い（JUnit XML に結果が無い）" for i in ids if i not in by_id]
    errors += [f"{i}: テストコードにあるが台帳 cases に無い" for i in sorted(set(by_id) - set(ids))]
    errors += [f"テスト名に TC-ID が無い: {n}" for n in run["unmatched"]]
    return errors


def results(content):
    """(per-case results in catalog order, run info). Missing JUnit → every case 未実施."""
    path = junit_path(content)
    by_id, run = parse_junit(path) if path.exists() else ({}, {"executed_at": "", "seconds": "", "unmatched": []})
    rows = [{**c, **by_id.get(c["id"], {"result": "未実施", "message": "結果なし"})} for c in content["cases"]]
    return rows, run


def summary_by_function(rows):
    """{F-xx: {"total", "OK", "NG", "未実施"}}"""
    out = {}
    for r in rows:
        s = out.setdefault(r["function"], {"total": 0, "OK": 0, "NG": 0, "未実施": 0})
        s["total"] += 1
        s[r["result"]] += 1
    return out
