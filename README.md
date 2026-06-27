# office-comments

A small, surgical skill for **modern threaded comments** in Microsoft Office files — Word
(`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`). List, add, reply, resolve/reopen and delete
comments by editing **only** the comment-related parts inside the file's OOXML zip; every other
part is copied through byte-for-byte, so the rest of the document is never disturbed.

Works on any Office file (no branding required).

## Quick start

```bash
# See every comment thread (ids, authors, anchors, resolved state)
python scripts/list_comments.py --file report.docx
python scripts/list_comments.py --file report.docx --json     # machine-readable

# Add a comment (anchor differs by format)
python scripts/add_comment.py  --file report.docx --anchor-text "this figure" \
    --author "Alex Morgan" --text "Please confirm."
python scripts/add_comment.py  --file book.xlsx   --cell B2 --author "Alex Morgan" --text "Check."
python scripts/add_comment.py  --file deck.pptx   --slide 2 --author "Alex Morgan" --text "Revise."

# Reply, resolve/reopen, delete (use an id from list_comments)
python scripts/reply_comment.py   --file report.docx --parent 0 --author "Sam Lee" --text "Confirmed."
python scripts/resolve_comment.py --file report.docx --comment 0 [--reopen]
python scripts/delete_comment.py  --file report.docx --comment 1
```

## How it works
An Office file is a zip of XML parts. Comments live in their own parts; this skill opens the zip,
edits just those parts (plus the single spot a comment anchors to), and re-seals it — leaving
everything else identical. See `SKILL.md` for full usage and `RELEASING.md` for the release flow.

## Requirements
Python 3.9+, `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only for reading cell
values / slide titles and by the test suite. See `scripts/requirements.txt`.

## Tests
```bash
python evals/run_evals.py            # runs every tests/smoke_*.py
python scripts/release.py --check    # version sync + frontmatter + evals
```
