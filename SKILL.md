---
name: office-comments
description: |
  Read, add, reply to, resolve/reopen, and delete modern THREADED comments in any Microsoft
  Office file — Word (.docx), Excel (.xlsx), and PowerPoint (.pptx). Surgically edits only the
  comment-related parts inside the file's OOXML zip and copies every other part byte-for-byte,
  so the rest of the document is never disturbed. Works on any Office file regardless of
  branding. Handles the whole conversation: a top-level comment anchored to a phrase (Word), a
  cell (Excel) or a slide (PowerPoint); threaded replies; resolving and reopening threads; and
  deleting a comment, a reply or a whole thread. Always run list_comments first to see every
  thread with its author, anchored text/context, resolved state and ids, then act on a comment
  by id. Triggers: comment, threaded comment, reply to a comment, resolve comment, reopen
  comment, delete comment, review comments, comment thread, in Word/Excel/PowerPoint or
  .docx/.xlsx/.pptx. Do NOT use for legacy single-note comments only, plain non-Office files,
  PDF, or Google Docs.
metadata:
  version: 0.1.0
  author: Deval Karia
---

# office-comments

## 1. Overview
Reads, adds, replies to, resolves/reopens, and deletes **modern threaded comments** in Word
(`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`). It edits ONLY the comment-related XML parts
inside the file's zip and copies every other part byte-for-byte — the rest of the document is
never touched. Works on ANY Office file (no branding required).

*Modern threaded comments* are the conversation-style comments (a comment plus replies, which can
be marked Resolved), not the old single-author sticky notes.

**Division of labour:** this skill is the *hands* — it finds comments and reports them with their
context, and places a new comment exactly where you say. The agent is the *brain* — decide which
comment is meant, what to say, and where a new one goes.

## 2. When to use
Use whenever the user wants to work with comments in an Office file: see what comments exist, add
one, reply in a thread, resolve/reopen a thread, or delete a comment/reply/thread.
Triggers: "what are the comments", "add a comment", "reply to …", "resolve/close this", "reopen",
"delete that comment".
Do NOT use for: legacy sticky-note-only comments, non-Office files, PDF, or Google Docs.

## 3. What's bundled
- `scripts/_ooxml_zip.py` — the binary-safe zip-patch engine (preserves all other parts).
- `scripts/_docx_comments.py`, `_xlsx_comments.py`, `_pptx_comments.py` — per-format logic.
- `scripts/list_comments.py`, `add_comment.py`, `reply_comment.py`, `resolve_comment.py`,
  `delete_comment.py` — the CLIs (auto-detect the format from the extension).
- `tests/` — smoke tests; `evals/run_evals.py` runs them all. `dev/` — Windows+Office-only
  fixture tooling (not shipped).
- Requires `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only to read cell values /
  slide titles and in tests.

## 4. List comments — ALWAYS do this first to get ids
    python scripts/list_comments.py --file FILE
    python scripts/list_comments.py --file FILE --json
`--json` emits `{"comments": [...]}`; each record has `id`, `thread_id`, `parent_id`, `is_reply`,
`author`, `date`, `text`, `resolved`, `anchor`, `anchor_text`, `context`, `location`. Read it to
pick the right `id` before replying/resolving/deleting.

## 5. Add a comment
    python scripts/add_comment.py --file FILE --author NAME --text "..." <anchor>
Anchor by format:
- Word: `--anchor-text "phrase" [--occurrence N]`  OR  `--paragraph N`
- Excel: `--cell B2 [--sheet "Sheet name"]`  (sheet defaults to the first)
- PowerPoint: `--slide N`
Prints the new comment id. If a Word anchor phrase appears more than once, the tool reports the
count — re-run with `--occurrence`. The phrase must exist verbatim in the document.

## 6. Reply
    python scripts/reply_comment.py --file FILE --parent ID --author NAME --text "..."
`ID` may be the thread root or any comment in the thread. Prints the new reply id.

## 7. Resolve / reopen
    python scripts/resolve_comment.py --file FILE --comment ID [--reopen]
Affects the whole thread.

## 8. Delete
    python scripts/delete_comment.py --file FILE --comment ID
A reply is removed on its own; deleting a thread's root removes the whole thread.

## 9. Per-format notes
- **Word**: comment + each reply anchor to a text span; comment ids are small integers (`0`, `1`…).
  Adding a comment upgrades the document to modern compatibility mode (`compatibilityMode=15`) when
  it's older — Word disables the **Resolve** option for comments in Compatibility Mode, so this is
  required for resolve/reopen to work. No-op for files that are already modern.
- **Excel**: comments are cell-anchored. The skill keeps the modern thread AND the legacy "note"
  copy that Excel needs to display a comment at all; ids are GUIDs.
- **PowerPoint**: comments are slide-anchored and replies nest inside the comment; ids are GUIDs.

## 10. Hard rules
- Never hand-edit the zip — always go through the scripts; they preserve every non-comment part.
- Target a comment by `id` from `list_comments`, never by guessing.
- Author name is required for add/reply; default to the current user unless told otherwise.
- Verify generated files in real desktop Office before shipping any change (see RELEASING.md).

## 11. Known limitations (v0.1.0)
- **Excel**: one comment thread per cell (reply to extend it). The skill refuses, with a clear
  message, to modify a sheet that already holds *classic* (non-threaded) comments or non-comment
  form-control drawings, rather than risk overwriting them.
- **Word**: resolve/reopen needs a *modern threaded* comment; classic (pre-2013) comments can't be
  resolved.
- These are safe refusals (no data loss), surfaced as clear errors.

## 12. Versioning & releases
The version lives in `SKILL.md` (`metadata.version`) and `scripts/_ooxml_zip.py` (`__version__`);
`scripts/release.py` asserts they match. See RELEASING.md.
