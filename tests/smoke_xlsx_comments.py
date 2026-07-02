"""Smoke test: modern threaded comments in Excel (.xlsx).

add -> reply -> resolve/reopen -> delete via _xlsx_comments, asserting the threaded part,
the persons part, the regenerated legacy shadow (comments + VML + legacyDrawing), threading,
the resolved flag, byte-stability of unrelated parts, and that openpyxl still opens each step.

Run: python tests/smoke_xlsx_comments.py
"""
import os
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import _xlsx_comments as xc   # noqa: E402
import _errors                # noqa: E402
from openpyxl import Workbook, load_workbook   # noqa: E402


def _base(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"], ws["B1"] = "Parameter", "Value"
    ws["A2"], ws["B2"] = "North region", "$1.2M"
    wb.save(path)


def _opens(path):
    try:
        load_workbook(path)
        return True
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def main() -> int:
    fails = []
    tmp = tempfile.mkdtemp(prefix="xlsx_comments_")
    f = os.path.join(tmp, "t.xlsx")
    _base(f)
    before = _parts(f)

    try:
        xc.add_comment(f, {"sheet": "Nope", "cell": "B2"}, "x", author="A")
        fails.append("bad sheet: expected AnchorNotFound")
    except _errors.AnchorNotFound:
        pass
    except Exception as e:
        fails.append(f"bad sheet: wrong error {type(e).__name__}")

    # --- add ---
    cid = xc.add_comment(f, {"sheet": "Data", "cell": "B2"},
                         "Please confirm this figure matches the latest forecast.", author="Alex Morgan")
    recs = xc.list_comments(f)
    if len(recs) != 1:
        fails.append(f"add: expected 1 comment, got {len(recs)}")
    else:
        r = recs[0]
        if r["anchor"]["cell"] != "B2" or r["location"] != "Data!B2":
            fails.append(f"add: bad anchor/location {r['anchor']} {r['location']}")
        if r["author"] != "Alex Morgan":
            fails.append(f"add: author {r['author']!r}")
        if "$1.2M" not in r["context"]:
            fails.append(f"add: context missing cell value: {r['context']!r}")
    with zipfile.ZipFile(f) as z:
        names = set(z.namelist())
        for p in ("xl/threadedComments/threadedComment1.xml", "xl/persons/person.xml", "xl/comments1.xml"):
            if p not in names:
                fails.append(f"add: missing part {p}")
        if not any(n.endswith(".vml") for n in names):
            fails.append("add: no VML drawing written")
        ct = z.read("[Content_Types].xml").decode()
        for frag in ("ms-excel.threadedcomments", "ms-excel.person", "spreadsheetml.comments", "vmlDrawing"):
            if frag not in ct:
                fails.append(f"add: content-type missing {frag}")
        srels = z.read("xl/worksheets/_rels/sheet1.xml.rels").decode()
        for frag in ("relationships/threadedComment", "relationships/comments", "relationships/vmlDrawing"):
            if frag not in srels:
                fails.append(f"add: sheet rel missing {frag}")
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
        if "legacyDrawing" not in sheet_xml:
            fails.append("add: legacyDrawing not added to sheet")
        if "r:id=" not in sheet_xml:
            fails.append("add: legacyDrawing missing r:id (relationships prefix)")
        if "ns0:id=" in sheet_xml:
            fails.append("add: legacyDrawing serialised with ns0:id instead of r:id (#1 regression)")
    if (o := _opens(f)) is not True:
        fails.append(f"add: won't open: {o}")
    changed = sorted(n for n in before if before[n] != _parts(f).get(n))
    allowed = {"[Content_Types].xml", "xl/worksheets/sheet1.xml", "xl/_rels/workbook.xml.rels"}
    if not set(changed).issubset(allowed):
        fails.append(f"add: changed unexpected existing parts: {set(changed) - allowed}")
    if "[Content_Types].xml" not in changed or "xl/worksheets/sheet1.xml" not in changed:
        fails.append(f"add: expected content-types + sheet to change, got {changed}")

    # --- reply ---
    rid = xc.reply(f, cid, "Confirmed - it matches the final forecast.", author="Sam Lee")
    recs = {r["id"]: r for r in xc.list_comments(f)}
    if len(recs) != 2:
        fails.append(f"reply: expected 2, got {len(recs)}")
    if rid in recs:
        if recs[rid]["parent_id"] != cid or recs[rid]["thread_id"] != cid or not recs[rid]["is_reply"]:
            fails.append(f"reply: bad threading {recs[rid]}")
    else:
        fails.append("reply: reply id not listed")
    if (o := _opens(f)) is not True:
        fails.append(f"reply: won't open: {o}")

    # --- resolve / reopen ---
    xc.set_status(f, cid, True)
    if not all(r["resolved"] for r in xc.list_comments(f)):
        fails.append("resolve: not marked resolved")
    xc.set_status(f, rid, False)
    if any(r["resolved"] for r in xc.list_comments(f)):
        fails.append("reopen: still resolved")

    # --- delete reply then root ---
    xc.delete(f, rid)
    if [r["id"] for r in xc.list_comments(f)] != [cid]:
        fails.append("delete reply: wrong remaining set")
    xc.delete(f, cid)
    if xc.list_comments(f):
        fails.append("delete root: comments remain")
    if (o := _opens(f)) is not True:
        fails.append(f"final: won't open: {o}")

    # --- multi-comment regression: more than one comment per sheet must keep shapeId="0" on every
    #     legacy note. A per-comment shapeId (0,1,2..) points at a nonexistent VML shape, so Excel
    #     drops the whole comments part on open. Also covers comments on a non-first sheet. ---
    import re
    f2 = os.path.join(tmp, "multi.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    s2 = wb.create_sheet("Second")
    for i in range(1, 6):
        ws[f"B{i}"] = f"v{i}"
        s2[f"B{i}"] = f"w{i}"
    wb.save(f2)
    for i in range(1, 6):
        mcid = xc.add_comment(f2, {"sheet": "Data", "cell": f"B{i}"}, f"C{i}", author="Alex Morgan")
        xc.reply(f2, mcid, f"R{i}", author="Sam Lee")
    xc.add_comment(f2, {"sheet": "Second", "cell": "B1"}, "on the second sheet", author="Alex Morgan")

    if len(xc.list_comments(f2)) != 11:   # 5 roots + 5 replies on Data, 1 on Second
        fails.append(f"multi: expected 11 comments, got {len(xc.list_comments(f2))}")
    if (o := _opens(f2)) is not True:
        fails.append(f"multi: won't open: {o}")
    with zipfile.ZipFile(f2) as z:
        cparts = [n for n in z.namelist() if re.match(r"xl/comments\d+\.xml$", n)]
        if "xl/comments1.xml" not in cparts or "xl/comments2.xml" not in cparts:
            fails.append(f"multi: expected per-sheet comments1.xml + comments2.xml, got {sorted(cparts)}")
        for cpart in cparts:
            body = z.read(cpart).decode()
            shape_ids = re.findall(r'shapeId="([^"]*)"', body)
            if any(s != "0" for s in shape_ids):
                fails.append(f"multi: {cpart} has a non-zero shapeId {shape_ids}; Excel drops the "
                             f"comments unless every legacy note has shapeId=0")
            if body.count("<comment ") != body.count("<author>"):
                fails.append(f"multi: {cpart} comment/author count mismatch")

    # --- foreign relationship spelling: adding a comment must REUSE the sheet's existing
    #     comments/vml rel, not append a duplicate (duplicate rels make Excel drop all comments) ---
    f3 = os.path.join(tmp, "foreign.xlsx")
    _base(f3)
    xc.add_comment(f3, {"sheet": "Data", "cell": "B2"}, "one", author="Alex Morgan")
    rp = "xl/worksheets/_rels/sheet1.xml.rels"
    with zipfile.ZipFile(f3) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    data[rp] = (data[rp].decode()
                .replace('Target="../comments1.xml"', 'Target="/xl/comments1.xml"')
                .replace('Target="../drawings/vmlDrawing1.vml"', 'Target="/xl/drawings/vmlDrawing1.vml"')
                .encode())
    with zipfile.ZipFile(f3, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in infos:
            zo.writestr(i, data[i.filename])
    xc.add_comment(f3, {"sheet": "Data", "cell": "B3"}, "two", author="Alex Morgan")
    with zipfile.ZipFile(f3) as z:
        srels = z.read(rp).decode()
    for frag, name in (("relationships/comments", "comments"), ("relationships/vmlDrawing", "vmlDrawing")):
        if srels.count(frag) != 1:
            fails.append(f"foreign-rels: expected 1 {name} rel, got {srels.count(frag)} "
                         f"(a duplicate rel makes Excel drop the comments)")
    if (o := _opens(f3)) is not True:
        fails.append(f"foreign-rels: won't open: {o}")

    if fails:
        print("FAIL:")
        for x in fails:
            print("  -", x)
        return 1
    print("PASS — xlsx threaded comments (add / reply / resolve / reopen / delete)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
