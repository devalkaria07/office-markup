"""Desktop-Office verification for office-markup (Windows + real Office, dev-only, not shipped).

Drives REAL Word / Excel via COM to prove the v0.3.0 behaviours the smoke tests can't:
  A  header revisions — the skill's accept_all agrees with Word's own AcceptAllRevisions
     on twin copies of a Word-authored header change; the result reopens with 0 revisions.
  A2 authored header edit — a replacement the SKILL wrote into a header is accepted by
     Word itself with the expected header text.
  B  Excel notes coexistence — after a skill add/reply/delete cycle on a workbook that
     holds a classic note, the note text is intact, the thread behaves, and the workbook
     opens with no repair.
  C  batch — an apply_edits result renders in Word: redlines + comment thread, right authors.

    python dev/verify_desktop.py          # runs all checks, prints PASS/FAIL per check

Close stray Office instances first if files are locked (Get-Process word,excel | Stop-Process).
"""
import os
import shutil
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))

import _docx_batch as B          # noqa: E402
import _docx_comments as dc      # noqa: E402
import _docx_revisions as R      # noqa: E402
import _xlsx_comments as xc      # noqa: E402

WD_FORMAT_DOCX = 12
XL_XLSX = 51


def _word():
    import win32com.client as win32
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app


def _excel():
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    return xl


def check_a_header_parity(tmp, fails):
    """Word authors a header change; skill accept vs Word accept must agree."""
    src = os.path.join(tmp, "rev_header.docx")
    app = _word()
    try:
        doc = app.Documents.Add()
        doc.Range(0, 0).InsertAfter("Body text stays put.")
        doc.Sections(1).Headers(1).Range.Text = "Header line"
        doc.TrackRevisions = True
        hr = doc.Sections(1).Headers(1).Range
        hr.Collapse(0)
        hr.InsertAfter(" plus header change")
        if os.path.exists(src):
            os.remove(src)
        doc.SaveAs2(src, WD_FORMAT_DOCX)
        doc.Close(False)

        a, b = os.path.join(tmp, "hdr_skill.docx"), os.path.join(tmp, "hdr_word.docx")
        shutil.copy(src, a)
        shutil.copy(src, b)

        n = R.accept_all(a)                      # the skill
        if n < 1:
            fails.append(f"A: skill accept_all reported {n}, expected >=1")

        docb = app.Documents.Open(b)             # Word itself
        docb.AcceptAllRevisions()
        docb.Save()
        word_hdr = docb.Sections(1).Headers(1).Range.Text.strip()
        docb.Close(False)

        doca = app.Documents.Open(a)             # reopen the skill's result IN WORD
        skill_hdr = doca.Sections(1).Headers(1).Range.Text.strip()
        revs = doca.Revisions.Count
        doca.Close(False)

        if skill_hdr != word_hdr:
            fails.append(f"A: header text differs — skill {skill_hdr!r} vs Word {word_hdr!r}")
        if revs != 0:
            fails.append(f"A: {revs} revision(s) remain in the skill's result")
    finally:
        app.Quit()


def check_a2_authored_header_edit(tmp, fails):
    """The SKILL replaces text inside a header; Word's own accept produces the expected text."""
    f = os.path.join(tmp, "authored_hdr.docx")
    app = _word()
    try:
        doc = app.Documents.Add()
        doc.Range(0, 0).InsertAfter("Body text.")
        doc.Sections(1).Headers(1).Range.Text = "Confidential draft"
        if os.path.exists(f):
            os.remove(f)
        doc.SaveAs2(f, WD_FORMAT_DOCX)
        doc.Close(False)

        R.replace_tracked(f, {"text": "Confidential draft"}, "Approved final",
                          author="Desktop Verifier")

        doc = app.Documents.Open(f)
        # NB: doc.Revisions covers only the BODY story; header revisions live on the
        # header range's own collection.
        hdr_revs = doc.Sections(1).Headers(1).Range.Revisions.Count
        if hdr_revs < 2:
            fails.append(f"A2: expected >=2 header revisions (del+ins), Word sees {hdr_revs}")
        doc.AcceptAllRevisions()
        hdr = doc.Sections(1).Headers(1).Range.Text.strip()
        doc.Close(False)
        if hdr != "Approved final":
            fails.append(f"A2: header after Word's accept is {hdr!r}, expected 'Approved final'")
    finally:
        app.Quit()


def check_b_excel_notes(tmp, fails):
    """Classic note + skill-managed thread on one sheet; the note survives everything."""
    f = os.path.join(tmp, "mixed.xlsx")
    xl = _excel()
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Data"
        ws.Range("B2").Value = "$1.2M"
        ws.Range("D4").AddComment("Classic note text")
        if os.path.exists(f):
            os.remove(f)
        wb.SaveAs(f, XL_XLSX)
        wb.Close(False)

        cid = xc.add_comment(f, {"sheet": "Data", "cell": "B2"},
                             "Thread beside a note.", author="Desktop Verifier")
        xc.reply(f, cid, "And a reply.", author="Desktop Verifier")

        wb = xl.Workbooks.Open(f)                # would raise / repair-prompt if broken
        ws = wb.Worksheets("Data")
        note = ws.Range("D4").Comment
        if note is None or "Classic note text" not in note.Text():
            fails.append("B: classic note text lost after skill add/reply")
        if ws.Range("B2").CommentThreaded is None:
            fails.append("B: threaded comment not visible to Excel")
        wb.Close(False)

        xc.delete(f, cid)                        # last thread -> teardown path
        wb = xl.Workbooks.Open(f)
        ws = wb.Worksheets("Data")
        note = ws.Range("D4").Comment
        if note is None or "Classic note text" not in note.Text():
            fails.append("B: classic note lost after deleting the last thread")
        if ws.Range("B2").CommentThreaded is not None:
            fails.append("B: thread still present after delete")
        wb.Close(False)
    finally:
        xl.Quit()


def check_c_batch(tmp, fails):
    """apply_edits result renders in Word: redlines + comment thread, correct authors."""
    f = os.path.join(tmp, "batch.docx")
    from docx import Document
    d = Document()
    d.add_paragraph("The Q3 figures need review before release.")
    d.save(f)
    cid = dc.add_comment(f, {"text": "review"}, "Existing thread.", author="Alex Morgan")
    B.apply_edits(f, {"author": "Desktop Verifier", "edits": [
        {"action": "replace", "anchor": {"text": "Q3 figures"}, "text": "Q3 FINAL figures"},
        {"action": "delete", "anchor": {"text": " before release"}},
        {"action": "reply", "comment_id": cid, "text": "Handled in this pass."},
        {"action": "resolve", "comment_id": cid},
    ]})
    app = _word()
    try:
        doc = app.Documents.Open(f)
        if doc.Revisions.Count != 3:             # replace = del+ins, delete = del
            fails.append(f"C: Word sees {doc.Revisions.Count} revisions, expected 3")
        authors = {doc.Revisions(i + 1).Author for i in range(doc.Revisions.Count)}
        if authors != {"Desktop Verifier"}:
            fails.append(f"C: revision authors {authors}")
        if doc.Comments.Count != 2:
            fails.append(f"C: Word sees {doc.Comments.Count} comments, expected 2")
        doc.Close(False)
    finally:
        app.Quit()


def main() -> int:
    fails = []
    tmp = tempfile.mkdtemp(prefix="om_verify_")
    for name, fn in (("A header parity", check_a_header_parity),
                     ("A2 authored header edit", check_a2_authored_header_edit),
                     ("B excel notes coexistence", check_b_excel_notes),
                     ("C batch render", check_c_batch)):
        n0 = len(fails)
        try:
            fn(tmp, fails)
        except Exception as e:  # noqa: BLE001
            fails.append(f"{name.split()[0]}: crashed — {type(e).__name__}: {e}")
        print(("PASS  " if len(fails) == n0 else "FAIL  ") + name)
    if fails:
        print("\nFAILURES:")
        for x in fails:
            print("  -", x)
        return 1
    print("\nALL DESKTOP CHECKS PASSED —", tmp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
