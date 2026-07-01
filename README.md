# office-markup

A small, surgical skill for two kinds of **Microsoft Office markup**: modern **threaded comments**
(Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`) and Word **tracked changes** (revisions / redlines).
It edits **only** the relevant parts inside the file's OOXML zip; every other part is copied through
byte-for-byte, so the rest of the document is never disturbed.

Works on any Office file (no branding required).

## Comments (Word, Excel, PowerPoint)

```bash
python scripts/list_comments.py  --file report.docx            # every thread (+ --json)
python scripts/add_comment.py    --file report.docx --anchor-text "this figure" \
    --author "Alex Morgan" --text "Please confirm."
python scripts/add_comment.py    --file book.xlsx --cell B2 --author "Alex Morgan" --text "Check."
python scripts/add_comment.py    --file deck.pptx --slide 2 --author "Alex Morgan" --text "Revise."
python scripts/reply_comment.py   --file report.docx --parent 0 --author "Sam Lee" --text "Confirmed."
python scripts/resolve_comment.py --file report.docx --comment 0 [--reopen]
python scripts/delete_comment.py  --file report.docx --comment 1
```

## Tracked changes (Word only)

```bash
python scripts/list_revisions.py  --file report.docx           # every change (+ --json)
python scripts/accept_change.py   --file report.docx --rev 3   # or --all [--author NAME]
python scripts/reject_change.py   --file report.docx --all
python scripts/track_changes.py   --file report.docx --on      # or --off
python scripts/insert_text.py     --file report.docx --anchor-text "revenue" \
    --text " (Q3)" --author "Alex Morgan"
python scripts/delete_text.py     --file report.docx --anchor-text "old clause" --author "Sam Lee"
```

## How it works
An Office file is a zip of XML parts. Comments live in their own parts; Word tracked changes live
inline in `document.xml` plus a one-line switch in `settings.xml`. This skill opens the zip, edits
just those pieces, and re-seals it — leaving everything else byte-for-byte identical. Every
accept/reject reproduces what Word's own Accept/Reject produces (verified against real Microsoft
Word). See `SKILL.md` for full usage and `RELEASING.md` for the release flow.

## Requirements
Python 3.9+, `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only for reading cell
values / slide titles and by the test suite. See `scripts/requirements.txt`.

## Tests
```bash
python evals/run_evals.py            # runs every tests/smoke_*.py
python scripts/release.py --check    # version sync + frontmatter + evals
```
