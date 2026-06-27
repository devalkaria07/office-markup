"""Smoke test: modern threaded comments in Word (.docx).

Builds a tiny doc, then drives add -> reply -> resolve/reopen -> delete entirely through
_docx_comments, asserting structure, threading, the resolved flag, anchor round-trip, the
byte-stability guarantee (only comment-related parts change), and that python-docx still
opens every intermediate file.

Run: python tests/smoke_docx_comments.py
"""
import os
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
import _docx_comments as dc          # noqa: E402
import _errors                       # noqa: E402
from docx import Document            # noqa: E402


def _base(path):
    d = Document()
    d.add_paragraph("The Q3 revenue figures need review before we circulate this report.")
    d.add_paragraph("A second paragraph mentions revenue again to test ambiguity.")
    d.add_heading("Scope", level=1)
    d.save(path)


def _opens(path):
    try:
        Document(path)
        return True
    except Exception as e:
        return f"{type(e).__name__}: {e}"


def _parts(path):
    with zipfile.ZipFile(path) as z:
        return {i.filename: z.read(i.filename) for i in z.infolist()}


def _docxml(path):
    with zipfile.ZipFile(path) as z:
        return z.read("word/document.xml").decode("utf-8", "ignore")


def main() -> int:
    fails = []
    tmp = tempfile.mkdtemp(prefix="docx_comments_")
    f = os.path.join(tmp, "t.docx")
    _base(f)
    before = _parts(f)

    # --- ambiguity + not-found guards (before we touch anything) ---
    try:
        dc.add_comment(f, {"text": "revenue"}, "x", author="A")
        fails.append("ambiguous anchor: expected AmbiguousAnchor")
    except _errors.AmbiguousAnchor:
        pass
    except Exception as e:
        fails.append(f"ambiguous anchor: wrong error {type(e).__name__}: {e}")
    try:
        dc.add_comment(f, {"text": "does-not-exist-xyz"}, "x", author="A")
        fails.append("missing anchor: expected AnchorNotFound")
    except _errors.AnchorNotFound:
        pass
    except Exception as e:
        fails.append(f"missing anchor: wrong error {type(e).__name__}: {e}")

    # --- add ---
    cid = dc.add_comment(f, {"text": "revenue", "occurrence": 1},
                         "Please confirm this figure matches the latest forecast.",
                         author="Alex Morgan")
    recs = dc.list_comments(f)
    if len(recs) != 1:
        fails.append(f"add: expected 1 comment, got {len(recs)}")
    else:
        r = recs[0]
        if r["anchor_text"] != "revenue":
            fails.append(f"add: anchor_text={r['anchor_text']!r}")
        if r["author"] != "Alex Morgan" or r["initials"] != "AM":
            fails.append(f"add: author/initials wrong: {r['author']!r}/{r['initials']!r}")
        if r["resolved"] or r["is_reply"]:
            fails.append("add: should be unresolved top-level")
        if "Q3 revenue figures" not in r["context"]:
            fails.append(f"add: context missing surrounding text: {r['context']!r}")
    with zipfile.ZipFile(f) as z:
        names = set(z.namelist())
        for p in ("word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml",
                  "word/commentsExtensible.xml", "word/people.xml"):
            if p not in names:
                fails.append(f"add: missing part {p}")
        ct = z.read("[Content_Types].xml").decode()
        if "wordprocessingml.comments+xml" not in ct:
            fails.append("add: comments content-type not registered")
        rels = z.read("word/_rels/document.xml.rels").decode()
        for frag in ("relationships/comments", "office/2011/relationships/commentsExtended",
                     "office/2016/09/relationships/commentsIds",
                     "office/2018/08/relationships/commentsExtensible",
                     "office/2011/relationships/people"):
            if frag not in rels:
                fails.append(f"add: relationship missing: {frag}")
    if (o := _opens(f)) is not True:
        fails.append(f"add: won't open: {o}")
    after = _parts(f)
    changed = set(n for n in before if before[n] != after.get(n))
    # settings.xml changes because we upgrade compatibilityMode 14 -> 15 so Word allows Resolve
    allowed = {"word/document.xml", "[Content_Types].xml", "word/_rels/document.xml.rels", "word/settings.xml"}
    if not changed.issubset(allowed):
        fails.append(f"add: changed unexpected existing parts: {changed - allowed}")
    if "word/document.xml" not in changed or "word/settings.xml" not in changed:
        fails.append(f"add: expected document.xml + settings.xml to change, got {sorted(changed)}")

    # --- reply ---
    rid = dc.reply(f, cid, "Confirmed - it matches the final forecast.", author="Sam Lee")
    recs = {r["id"]: r for r in dc.list_comments(f)}
    if len(recs) != 2:
        fails.append(f"reply: expected 2 comments, got {len(recs)}")
    if rid in recs:
        rep = recs[rid]
        if rep["parent_id"] != cid or not rep["is_reply"] or rep["thread_id"] != cid:
            fails.append(f"reply: bad threading parent={rep['parent_id']} thread={rep['thread_id']}")
        if rep["author"] != "Sam Lee":
            fails.append("reply: author wrong")
    else:
        fails.append("reply: reply id not listed")
    if _docxml(f).count("commentReference") != 2:
        fails.append("reply: expected 2 commentReference anchors")
    if (o := _opens(f)) is not True:
        fails.append(f"reply: won't open: {o}")

    # --- resolve / reopen (via either id in the thread) ---
    dc.set_status(f, cid, True)
    if not all(r["resolved"] for r in dc.list_comments(f)):
        fails.append("resolve: thread not marked resolved")
    dc.set_status(f, rid, False)   # reopen via the reply id -> same thread
    if any(r["resolved"] for r in dc.list_comments(f)):
        fails.append("reopen: thread still resolved")
    if (o := _opens(f)) is not True:
        fails.append(f"resolve: won't open: {o}")

    # --- delete the reply, then the root ---
    dc.delete(f, rid)
    recs = dc.list_comments(f)
    if [r["id"] for r in recs] != [cid]:
        fails.append(f"delete reply: expected [{cid}], got {[r['id'] for r in recs]}")
    if _docxml(f).count("commentReference") != 1:
        fails.append("delete reply: expected 1 remaining anchor")
    dc.delete(f, cid)
    if dc.list_comments(f):
        fails.append("delete root: expected 0 comments")
    dx = _docxml(f)
    for tag in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        if tag in dx:
            fails.append(f"delete root: {tag} left behind in document.xml")
    if (o := _opens(f)) is not True:
        fails.append(f"delete: won't open: {o}")

    if fails:
        print("FAIL:")
        for x in fails:
            print("  -", x)
        return 1
    print("PASS — docx threaded comments (add / reply / resolve / reopen / delete)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
