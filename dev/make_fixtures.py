"""Generate ground-truth comment fixtures by driving REAL Office via COM (Windows-only,
dev-only, not shipped). Inspect the output with dev/inspect_fixture.py to learn the exact
parts/attributes/strings each per-format module must reproduce.

    python dev/make_fixtures.py [word|excel|powerpoint|word-revisions|word-header-revisions|excel-notes|all]
    (default: word)

Notes:
- Uses DispatchEx -> a private, hidden Office instance, so it won't touch anything you
  have open.
- Word: 'Comments.Add' and 'Replies.Add(Range, Text)' work. Marking a thread resolved is
  NOT exposed to COM ("command not available"), so the resolved STATE is validated against
  our own engine output in desktop Word, not captured here. The resolve mechanic is simply
  w15:done -> "1" on the thread root in word/commentsExtended.xml.
- CLASSIC Word comments (pre-2013, no w14:paraId): modern Word's COM 'Comments.Add' always
  emits modern five-part comments, so there is NO fixture job for the classic shape — the
  tests hand-inject it from the documented pre-2013 markup instead.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
FIXTURES = os.path.join(HERE, "fixtures")
WD_FORMAT_DOCX = 12  # wdFormatXMLDocument


def _save_word(doc, path):
    if os.path.exists(path):
        os.remove(path)
    doc.SaveAs2(path, WD_FORMAT_DOCX)
    return path


def word_fixtures(outdir=FIXTURES):
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    made = []
    try:
        doc = app.Documents.Add()
        rng = doc.Range(0, 0)
        rng.InsertAfter("The Q3 revenue figures need review before we circulate this report.")
        anchor = doc.Range(0, 17)  # "The revenue"
        c = doc.Comments.Add(anchor, "Please confirm this figure matches the latest forecast.")
        try:
            c.Author = "Alex Morgan"
        except Exception:
            pass
        made.append(_save_word(doc, os.path.join(outdir, "word_toplevel.docx")))

        rep = c.Replies.Add(c.Scope, "Confirmed - it matches the final forecast.")
        try:
            rep.Author = "Sam Lee"
        except Exception:
            pass
        made.append(_save_word(doc, os.path.join(outdir, "word_reply.docx")))
        doc.Close(False)
    finally:
        app.Quit()
    return made


def word_revision_fixtures(outdir=FIXTURES):
    """Drive real Word with Track Changes ON to emit one fixture per revision type — the ground
    truth _docx_revisions.py is built against. Inspect with dev/inspect_fixture.py."""
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    made = []

    def base(text="The quarterly revenue figures need careful review today."):
        doc = app.Documents.Add()
        doc.Content.InsertAfter(text)
        doc.TrackRevisions = True
        return doc

    def out(doc, name):
        p = _save_word(doc, os.path.join(outdir, name))
        doc.Close(False)
        made.append(p)

    try:
        d = base(); d.Range(0, 0).InsertAfter("URGENT "); out(d, "rev_insert.docx")
        d = base(); d.Range(0, 4).Delete(); out(d, "rev_delete.docx")
        d = base(); d.Range(4, 13).Font.Bold = True; out(d, "rev_rprchange.docx")
        d = base(); d.Paragraphs(1).Alignment = 1; out(d, "rev_pprchange.docx")
        d = base(); d.Range(20, 20).InsertParagraph(); out(d, "rev_ins_paramark.docx")

        d = app.Documents.Add()
        d.Content.InsertAfter("First paragraph here.\rSecond paragraph here.")
        d.TrackRevisions = True
        end1 = d.Paragraphs(1).Range.End
        d.Range(end1 - 1, end1).Delete()
        out(d, "rev_del_paramark.docx")

        d = app.Documents.Add(); tbl = d.Tables.Add(d.Range(0, 0), 2, 2)
        d.TrackRevisions = True; tbl.Rows.Add(); out(d, "rev_row_ins.docx")
        d = app.Documents.Add(); tbl = d.Tables.Add(d.Range(0, 0), 3, 2)
        d.TrackRevisions = True; tbl.Rows(2).Delete(); out(d, "rev_row_del.docx")

        d = base(); d.Paragraphs(1).Range.ListFormat.ApplyNumberDefault(); out(d, "rev_numbering.docx")

        d = app.Documents.Add()
        d.Content.InsertAfter("ALPHA one. BETA two. GAMMA three.")
        d.TrackRevisions = True
        d.Range(0, 10).Cut()
        d.Range(d.Content.End - 1, d.Content.End - 1).Paste()
        out(d, "rev_move.docx")

        d = app.Documents.Add()
        d.Fields.Add(d.Range(0, 0), -1, 'DATE \\@ "yyyy"', False)
        d.TrackRevisions = True
        d.Range(0, d.Content.End - 1).Delete()
        out(d, "rev_field_del.docx")

        d = app.Documents.Add()
        om = d.Range(0, 0).OMaths.Add(d.Range(0, 0))
        d.OMaths(1).Range.Text = "a+b"
        d.OMaths(1).BuildUp()
        d.TrackRevisions = True
        r = d.OMaths(1).Range
        d.Range(r.Start, r.Start + 1).Delete()
        out(d, "rev_math.docx")
    finally:
        app.Quit()
    return made


def _save(obj, path, fmt):
    if os.path.exists(path):
        os.remove(path)
    obj.SaveAs(path, fmt)
    return path


def word_header_revision_fixtures(outdir=FIXTURES):
    """Drive real Word to leave TRACKED CHANGES inside a header and a footnote — ground truth
    for the multi-story support (#8)."""
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    app = win32.DispatchEx("Word.Application")
    app.Visible = False
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    made = []
    try:
        # header: text laid down untracked, then edited with tracking on
        doc = app.Documents.Add()
        doc.Range(0, 0).InsertAfter("Body text stays put.")
        hdr = doc.Sections(1).Headers(1).Range          # 1 = wdHeaderFooterPrimary
        hdr.Text = "Header line"
        doc.TrackRevisions = True
        hr = doc.Sections(1).Headers(1).Range
        hr.Collapse(0)                                  # 0 = wdCollapseEnd
        hr.InsertAfter(" plus header change")
        made.append(_save_word(doc, os.path.join(outdir, "rev_header.docx")))
        doc.Close(False)

        # footnote: added untracked, then edited with tracking on
        doc = app.Documents.Add()
        doc.Range(0, 0).InsertAfter("Body sentence with a note.")
        fn_anchor = doc.Range(doc.Content.End - 2, doc.Content.End - 2)
        fn = doc.Footnotes.Add(fn_anchor, "", "A footnote.")
        doc.TrackRevisions = True
        fr = fn.Range
        fr.Collapse(0)
        fr.InsertAfter(" plus note change")
        made.append(_save_word(doc, os.path.join(outdir, "rev_footnote.docx")))
        doc.Close(False)
    finally:
        app.Quit()
    return made


def excel_notes_fixtures(outdir=FIXTURES):
    """Drive real Excel to create CLASSIC notes (Range.AddComment) — alone and mixed with a
    threaded comment on the same sheet. Ground truth for the coexistence work."""
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    made = []
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Data"
        ws.Range("A2").Value = "North region"
        ws.Range("B2").Value = "$1.2M"
        ws.Range("D4").AddComment("Classic note text")          # legacy note
        made.append(_save(wb, os.path.join(outdir, "excel_note_only.xlsx"), 51))
        ws.Range("B2").AddCommentThreaded("A threaded comment beside a note.")
        made.append(_save(wb, os.path.join(outdir, "excel_note_mixed.xlsx"), 51))
        wb.Close(False)
    finally:
        xl.Quit()
    return made


def excel_fixtures(outdir=FIXTURES):
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    made = []
    try:
        wb = xl.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Range("A2").Value = "North region"
        ws.Range("B2").Value = "$1.2M"
        ct = ws.Range("B2").AddCommentThreaded("Please confirm this figure matches the latest forecast.")
        made.append(_save(wb, os.path.join(outdir, "excel_toplevel.xlsx"), 51))  # 51 = xlOpenXMLWorkbook
        ct.AddReply("Confirmed - it matches the final forecast.")
        made.append(_save(wb, os.path.join(outdir, "excel_reply.xlsx"), 51))
        try:
            ct.Resolved = True
            made.append(_save(wb, os.path.join(outdir, "excel_resolved.xlsx"), 51))
        except Exception:
            pass
        wb.Close(False)
    finally:
        xl.Quit()
    return made


def powerpoint_fixtures(outdir=FIXTURES):
    import win32com.client as win32
    os.makedirs(outdir, exist_ok=True)
    ppt = win32.DispatchEx("PowerPoint.Application")
    try:
        ppt.Visible = True   # PowerPoint COM dislikes Visible=False
    except Exception:
        pass
    made = []
    try:
        pres = ppt.Presentations.Add()
        slide = pres.Slides.Add(1, 2)  # 2 = ppLayoutText
        try:
            slide.Shapes(1).TextFrame.TextRange.Text = "Quarterly Review Summary"
        except Exception:
            pass
        c = slide.Comments.Add(100, 100, "Alex Morgan", "AM", "Please confirm this figure.")
        made.append(_save(pres, os.path.join(outdir, "pptx_toplevel.pptx"), 24))  # 24 = ppSaveAsOpenXMLPresentation
        c.Replies.Add(100, 100, "Sam Lee", "SL", "Confirmed - it matches the final forecast.")
        made.append(_save(pres, os.path.join(outdir, "pptx_reply.pptx"), 24))
        pres.Close()
    finally:
        ppt.Quit()
    return made


def main(argv):
    which = (argv[0] if argv else "word").lower()
    jobs = {"word": word_fixtures, "excel": excel_fixtures, "powerpoint": powerpoint_fixtures,
            "word-revisions": word_revision_fixtures,
            "word-header-revisions": word_header_revision_fixtures,
            "excel-notes": excel_notes_fixtures}
    if which == "all":
        targets = list(jobs)
    elif which in jobs:
        targets = [which]
    else:
        print(__doc__)
        return 2
    for t in targets:
        try:
            made = jobs[t]()
            print(f"[{t}] wrote:")
            for m in made:
                print("  ", m)
        except NotImplementedError as e:
            print(f"[{t}] {e}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
