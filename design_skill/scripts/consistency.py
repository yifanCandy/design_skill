"""Content consistency between a document and its upstream documents (beyond ID existence).

traceability.check_content() only guarantees that the IDs a document uses exist in the master and that the
required ones are mentioned. This module compares the *content* that must be carried over unchanged:

    要件定義書        3章の要件一覧: 全 REQ が1行ずつ、要件名・対応機能が主表どおり
    基本設計書 F      本文の REQ はこの F の REQ だけ
    詳細仕様書 F      4章のサブ機能構成 = 同じ F の基本設計書4章
                      6章の画面 = この F を含む SCR（「SCR-xx <画面名>」）、項目表に 日本語名/必須/桁数 列
                      同じ画面の項目の 必須・桁数 が、作成済みの他の F の詳細仕様書と一致
                      9章のメッセージ = この F の MSG、区分・文言が主表どおり
    詳細設計書 SCR    IF定義書のシナリオ = この画面の IF、process_name・通信方式・URL が主表どおり
                      メッセージID一覧 = この画面の機能の MSG、「ツール仕様書のエラーID」が同じ MSG
                      入力チェック表（項目名/必須/桁数）の各項目が詳細仕様書6章と一致し、必須項目が全部ある

check() is what generate.py runs before rendering (exit 1, CONSISTENCY FAILED); traceability.collect_issues()
runs it for every existing document, so a downstream document that no longer matches a regenerated upstream
one shows up as 不整合 in the matrix.
"""
import json
import re
import unicodedata

REQ_HEADERS = ["要件ID", "要件名", "対応機能"]
SPEC_ITEM_HEADERS = ["日本語名", "必須", "桁数"]
SPEC_MSG_HEADERS = ["メッセージID", "区分", "メッセージ"]
DESIGN_MSG_HEADERS = ["メッセージID", "ツール仕様書のエラーID"]
DESIGN_CHECK_HEADERS = ["項目名", "必須", "桁数"]

YES = {"○", "◯", "〇", "●", "必須", "y", "yes", "true", "要"}
ID_IN = {
    "REQ": re.compile(r"REQ-\d{3}"), "F": re.compile(r"(?<!I)F-\d{2}"),
    "IF": re.compile(r"IF-\d{2}"), "MSG": re.compile(r"MSG-\d{3}"),
}


def norm(s):
    """Compare text ignoring width (ＡＢ１ vs AB1) and whitespace."""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(s or "")))


def required(s):
    return norm(s).lower() in YES


def length(s):
    return norm(s).replace("桁", "").replace("文字", "") or "-"


def col(headers, name):
    return next((i for i, h in enumerate(headers) if norm(h) == norm(name)), None)


def missing_headers(headers, names):
    return [n for n in names if col(headers, n) is None]


def cell(row, i):
    return row[i] if i is not None and i < len(row) else ""


def load(sys_dir, doc_dir, target):
    path = sys_dir / doc_dir / f"{target}.json"
    if not path.exists():
        return None
    content = json.loads(path.read_text(encoding="utf-8"))
    content.pop("_notes_for_review", None)
    return content


def set_diff(label, got, want):
    errors = []
    if want - got:
        errors.append(f"{label}に {', '.join(sorted(want - got))} が無い")
    if got - want:
        errors.append(f"{label}に対象外の {', '.join(sorted(got - want))} がある")
    return errors


# ---------------------------------------------------------------- 要件定義書

def check_requirements(master, content):
    t = content["ch3_requirements_list"]
    lost = missing_headers(t["headers"], REQ_HEADERS)
    if lost:
        return [f"3章の要件一覧に列 {lost} が無い（列は master_rules.md の規定どおり）"]
    i_id, i_name, i_func = (col(t["headers"], h) for h in REQ_HEADERS)
    errors, seen = [], []
    for row in t["rows"]:
        rid = norm(cell(row, i_id))
        req = next((r for r in master["requirements"] if r["id"] == rid), None)
        if req is None:
            errors.append(f"3章の要件ID「{cell(row, i_id)}」が主表の REQ ではない")
            continue
        seen.append(rid)
        if norm(cell(row, i_name)) != norm(req["name"]):
            errors.append(f"3章 {rid} の要件名「{cell(row, i_name)}」が主表「{req['name']}」と違う")
        funcs = {f["id"] for f in master["functions"] if rid in f["requirements"]}
        errors += set_diff(f"3章 {rid} の対応機能", set(ID_IN["F"].findall(cell(row, i_func))), funcs)
    errors += [f"3章の要件一覧で {r} が重複している" for r in sorted({r for r in seen if seen.count(r) > 1})]
    errors += set_diff("3章の要件一覧", set(seen), {r["id"] for r in master["requirements"]})
    return errors


# ---------------------------------------------------------------- 基本設計書

def check_basic(master, target, content):
    func = next(f for f in master["functions"] if f["id"] == target)
    used = set(ID_IN["REQ"].findall(json.dumps(content, ensure_ascii=False)))
    other = used - set(func["requirements"])
    return [f"{target} の REQ ではない {', '.join(sorted(other))} を書いている（この F の REQ は "
            f"{', '.join(func['requirements'])}）"] if other else []


# ---------------------------------------------------------------- 詳細仕様書

def spec_screen_items(spec, screen_id):
    """{日本語名: (必須, 桁数)} of one screen in a tool spec's 6章 (None if the screen or columns are absent)."""
    for scr in spec["ch6_screen_definition"]["screens"]:
        if norm(scr["name"]).startswith(screen_id):
            h = scr["items_headers"]
            i_name, i_req, i_len = (col(h, n) for n in SPEC_ITEM_HEADERS)
            if None in (i_name, i_req, i_len):
                return None
            return {norm(cell(r, i_name)): (required(cell(r, i_req)), length(cell(r, i_len)))
                    for r in scr["items_rows"] if norm(cell(r, i_name))}
    return None


def check_spec(master, sys_dir, target, content):
    errors = []
    basic = load(sys_dir, "basic-design", target)
    if basic is not None:
        want = [norm(i["title"]) for i in basic["ch4_functions"]["items"]]
        got = [norm(i["title"]) for i in content["ch4_functions"]["items"]]
        if got != want:
            errors.append("4章のサブ機能構成が基本設計書4章と違う（基本設計書: "
                          f"{[i['title'] for i in basic['ch4_functions']['items']]}）")

    screens = [s for s in master["screens"] if target in s["functions"]]
    names = {norm(f"{s['id']} {s['name']}") for s in screens}
    got = {norm(s["name"]) for s in content["ch6_screen_definition"]["screens"]}
    errors += [f"6章の画面名「{n}」が主表の「SCR-xx <画面名>」と一致しない、またはこの F の画面ではない" for n in sorted(got - names)]
    errors += [f"6章に画面「{n}」が無い" for n in sorted(names - got)]
    for scr in content["ch6_screen_definition"]["screens"]:
        lost = missing_headers(scr["items_headers"], SPEC_ITEM_HEADERS)
        if lost:
            errors.append(f"6章「{scr['name']}」の項目定義表に列 {lost} が無い")

    for s in screens:  # the same screen in other functions' specs must define the same items the same way
        mine = spec_screen_items(content, s["id"])
        for other in s["functions"]:
            theirs = spec_screen_items(load(sys_dir, "detailed-spec", other) or {"ch6_screen_definition": {"screens": []}}, s["id"]) if other != target else None
            for name in sorted(set(mine or {}) & set(theirs or {})):
                if mine[name] != theirs[name]:
                    errors.append(f"6章 {s['id']} の項目「{name}」の必須・桁数が詳細仕様書 {other} と違う"
                                  f"（{target}: {fmt(mine[name])} / {other}: {fmt(theirs[name])}）")

    msgs = {m["id"]: m for m in master["messages"]}
    found = set()
    for sec in content["ch9_error_messages"]["sections"]:
        if not sec["has_messages"]:
            continue
        h = sec.get("headers") or []
        lost = missing_headers(h, SPEC_MSG_HEADERS)
        if lost:
            errors.append(f"9章「{sec['title']}」の表に列 {lost} が無い")
            continue
        i_id, i_lv, i_tx = (col(h, n) for n in SPEC_MSG_HEADERS)
        for row in sec.get("rows") or []:
            mid = norm(cell(row, i_id))
            if mid not in msgs:
                errors.append(f"9章のメッセージID「{cell(row, i_id)}」が主表の MSG ではない")
                continue
            found.add(mid)
            if norm(cell(row, i_lv)) != norm(msgs[mid]["level"]):
                errors.append(f"9章 {mid} の区分「{cell(row, i_lv)}」が主表「{msgs[mid]['level']}」と違う")
            if norm(cell(row, i_tx)) != norm(msgs[mid]["text"]):
                errors.append(f"9章 {mid} の文言「{cell(row, i_tx)}」が主表「{msgs[mid]['text']}」と違う")
    errors += set_diff("9章のメッセージ一覧", found, {m["id"] for m in master["messages"] if target in m["functions"]})
    return errors


def fmt(item):
    return f"必須={'○' if item[0] else '-'}, 桁数={item[1]}"


# ---------------------------------------------------------------- 詳細設計書

def check_design(master, sys_dir, target, content):
    errors = []
    screen = next(s for s in master["screens"] if s["id"] == target)
    ifs = {i["id"]: i for i in master["interfaces"]}

    found = set()
    for sc in content["if_definition"]["scenarios"]:
        hit = ID_IN["IF"].findall(sc["process_name"])
        if not hit:
            errors.append(f"IF定義書シナリオ「{sc['name']}」の process_name に IF-xx が無い（「IF-xx <IF名>」）")
            continue
        i = ifs.get(hit[0])
        if i is None:
            continue  # unknown IDs are reported by check_content
        found.add(i["id"])
        if norm(sc["process_name"]) != norm(f"{i['id']} {i['name']}"):
            errors.append(f"シナリオ「{sc['name']}」の process_name が主表の「{i['id']} {i['name']}」と違う")
        if norm(sc["communication"]["type"]) != norm(i["method"]):
            errors.append(f"{i['id']} の通信方式「{sc['communication']['type']}」が主表「{i['method']}」と違う")
        if norm(sc["communication"]["url"]) != norm(i["path"]):
            errors.append(f"{i['id']} の URL「{sc['communication']['url']}」が主表「{i['path']}」と違う")
    errors += set_diff("IF定義書のシナリオ", found, {i["id"] for i in master["interfaces"] if i["screen"] == target})

    ml = content["detailed_design"]["message_list"]
    lost = missing_headers(ml["headers"], DESIGN_MSG_HEADERS)
    if lost:
        errors.append(f"メッセージID一覧に列 {lost} が無い")
    else:
        i_id, i_spec = (col(ml["headers"], n) for n in DESIGN_MSG_HEADERS)
        listed = set()
        for row in ml["rows"]:
            mid = norm(cell(row, i_id))
            listed.add(mid)
            if norm(cell(row, i_spec)) != mid:
                errors.append(f"メッセージID一覧 {mid} の「ツール仕様書のエラーID」が「{cell(row, i_spec)}」（同じ MSG にする）")
        want = {m["id"] for m in master["messages"] if set(m["functions"]) & set(screen["functions"])}
        errors += set_diff("メッセージID一覧", listed, want)

    spec_items = {}
    for f in screen["functions"]:
        items = spec_screen_items(load(sys_dir, "detailed-spec", f) or {"ch6_screen_definition": {"screens": []}}, target)
        for name, v in (items or {}).items():
            spec_items.setdefault(name, (v, f))
    checked = set()
    for fn in content["detailed_design"]["functions"]:
        for step in fn["steps"]:
            for part in step["parts"]:
                h = part.get("headers") or []
                if part["type"] != "table" or missing_headers(h, DESIGN_CHECK_HEADERS):
                    continue
                i_name, i_req, i_len = (col(h, n) for n in DESIGN_CHECK_HEADERS)
                for row in part.get("rows") or []:
                    name = norm(cell(row, i_name))
                    mine = (required(cell(row, i_req)), length(cell(row, i_len)))
                    if name not in spec_items:
                        errors.append(f"「{fn['name']} {step['marker']}」の項目「{cell(row, i_name)}」が詳細仕様書6章の {target} に無い")
                        continue
                    checked.add(name)
                    (spec, f) = spec_items[name]
                    if mine != spec:
                        errors.append(f"「{fn['name']} {step['marker']}」の項目「{name}」の必須・桁数が詳細仕様書 {f} と違う"
                                      f"（詳細設計: {fmt(mine)} / 詳細仕様: {fmt(spec)}）")
    need = {n for n, ((req, _), _) in spec_items.items() if req}
    if need - checked:
        errors.append(f"詳細仕様書6章の必須項目 {sorted(need - checked)} が、入力チェックの表（列 {DESIGN_CHECK_HEADERS}）に無い")
    return errors


def check(master, sys_dir, doc_dir, target, content):
    """Errors for one document against the master and its upstream documents (empty = consistent)."""
    if doc_dir == "requirements-definition":
        return check_requirements(master, content)
    if doc_dir == "basic-design":
        return check_basic(master, target, content)
    if doc_dir == "detailed-spec" and target.startswith("F-"):
        return check_spec(master, sys_dir, target, content)
    if doc_dir == "detailed-design":
        return check_design(master, sys_dir, target, content)
    return []
