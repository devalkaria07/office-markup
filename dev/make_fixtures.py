"""Generate ground-truth comment fixtures by driving REAL Office via COM (Windows-only,
dev-only, not shipped). Inspect the output with dev/inspect_fixture.py to learn the exact
parts/attributes/strings each per-format module must reproduce.

    python dev/make_fixtures.py [word|excel|powerpoint|all]   (default: word)

Notes:
- Uses DispatchEx -> a private, hidden Office instance, so it won't touch anything you
  have open.
- Word: 'Comments.Add' and 'Replies.Add(Range, Text)' work. Marking a thread resolved is
  NOT exposed to COM ("command not available"), so the resolved STATE is validated against
  our own engine output in desktop Word, not captured here. The resolve mechanic is simply
  w15:done -> "1" on the thread root in word/commentsExtended.xml.
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


def _save(obj, path, fmt):
    if os.path.exists(path):
        os.remove(path)
    obj.SaveAs(path, fmt)
    return path


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
    jobs = {"word": word_fixtures, "excel": excel_fixtures, "powerpoint": powerpoint_fixtures}
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
