# office-markup

A small, surgical skill for **Microsoft Office markup**: modern **threaded comments**
(Word `.docx`, Excel `.xlsx`, PowerPoint `.pptx`), Word **tracked changes** (revisions /
redlines — across body, headers, footers and footnotes), and Word **text edits** (tracked or
direct, one at a time or a whole JSON batch in one all-or-nothing pass).
It edits **only** the relevant parts inside the file's OOXML zip; every other part is copied through
byte-for-byte, so the rest of the document is never disturbed. Excel classic notes are preserved
alongside threaded comments.

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
python scripts/list_revisions.py  --file report.docx           # every change, every story (+ --json)
python scripts/accept_change.py   --file report.docx --rev 3   # or --all [--author NAME]
python scripts/reject_change.py   --file report.docx --all
python scripts/track_changes.py   --file report.docx --on      # or --off
python scripts/insert_text.py     --file report.docx --anchor-text "revenue" \
    --text " (Q3)" --author "Alex Morgan"
python scripts/delete_text.py     --file report.docx --anchor-text "old clause" --author "Sam Lee"
python scripts/replace_text.py    --file report.docx --anchor-text "draft" \
    --text "final" --author "Alex Morgan"                      # one tracked change-X-to-Y
```

Anchors work in every story — add `--part header1` (or `footer2`, `footnotes`) to pin one.

## Direct edits & batch (Word only)

```bash
python scripts/edit_text.py   --file report.docx --anchor-text "teh" --replace-with "the"
python scripts/apply_edits.py --file report.docx --edits edits.json --author "Alex Morgan"
```

`edit_text` makes the change plainly — no tracking, nothing else touched. `apply_edits` runs a
JSON list of edits/comments/replies/resolves in one pass, **all-or-nothing**: if any entry
fails, the file is left untouched. Schema in `SKILL.md`.

## How it works
An Office file is a zip of XML parts. Comments live in their own parts; Word tracked changes live
inline in the story parts (`document.xml`, headers, footers, footnotes) plus a one-line switch in
`settings.xml`. This skill opens the zip, edits just those pieces, and re-seals it — leaving
everything else byte-for-byte identical (writes go to a temp file first and swap in atomically).
Every accept/reject reproduces what Word's own Accept/Reject produces (verified against real
Microsoft Word). See `SKILL.md` for full usage and `RELEASING.md` for the release flow.

## Requirements
Python 3.9+, `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only for reading cell
values / slide titles and by the test suite. See `scripts/requirements.txt`.

## Tests
```bash
python evals/run_evals.py            # runs every tests/smoke_*.py
python scripts/release.py --check    # version sync + frontmatter + evals
```
