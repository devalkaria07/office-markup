---
name: office-markup
description: |
  Read and write two kinds of Microsoft Office markup: modern THREADED comments and Word TRACKED
  CHANGES (revisions / redlines). Comments work in Word (.docx), Excel (.xlsx) and PowerPoint
  (.pptx): list, add, reply, resolve/reopen, delete. Tracked changes are Word-only: list every
  revision, accept/reject one or all (optionally by author), turn Track Changes on/off, and author
  tracked insertions and deletions. Surgically edits only the relevant XML parts inside the file's
  OOXML zip and copies every other part byte-for-byte, so the rest of the document is untouched;
  works on any Office file. Run list_comments or list_revisions first to see everything with ids,
  then act by id. Triggers:
  comment, threaded comment, reply, resolve/reopen/review comments; tracked changes, track changes,
  revisions, redline, accept/reject changes, insert/delete tracked text, who changed what — in
  Word/Excel/PowerPoint or .docx/.xlsx/.pptx. Do NOT use for legacy single-note comments only,
  plain non-Office files, PDF, or Google Docs.
metadata:
  version: 0.2.1
  author: Deval Karia
---

# office-markup

## 1. Overview
Reads and writes two kinds of **Office markup** by surgically editing only the relevant XML parts
inside the file's zip and copying every other part byte-for-byte — the rest of the document is
never touched. Works on ANY Office file (no branding required).

1. **Modern threaded comments** — the conversation-style comments (a comment plus replies, which
   can be marked Resolved), in Word (`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`).
2. **Word tracked changes** (revisions / redlines) — Word recording every insert / delete / format
   / structure edit so a reviewer can accept or reject each one. **Word (`.docx`) only** — Excel's
   revision system is deprecated and PowerPoint has none.

**Division of labour:** this skill is the *hands* — it finds comments/changes and reports them with
their context, and makes the exact edit you name. The agent is the *brain* — decide which item is
meant, what to say or change, and where.

## 2. When to use
- **Comments** (any Office file): see what comments exist, add one, reply, resolve/reopen, or delete.
  Triggers: "what are the comments", "add a comment", "reply to …", "resolve/close this", "reopen".
- **Tracked changes** (Word only): see every tracked change, accept/reject some or all, turn Track
  Changes on/off, or author a tracked insertion/deletion. Triggers: "track changes", "what changed",
  "accept/reject changes", "redline", "who edited this".

Do NOT use for: legacy sticky-note-only comments, non-Office files, PDF, or Google Docs.

## 3. What's bundled
- `scripts/_ooxml_zip.py` — the binary-safe zip-patch engine (preserves all other parts).
- Comments: `scripts/_docx_comments.py`, `_xlsx_comments.py`, `_pptx_comments.py`.
- Tracked changes: `scripts/_docx_revisions.py`. Shared anchoring: `scripts/_docx_anchor.py`.
- Comment CLIs: `list_comments.py`, `add_comment.py`, `reply_comment.py`, `resolve_comment.py`,
  `delete_comment.py`.
- Tracked-change CLIs: `list_revisions.py`, `accept_change.py`, `reject_change.py`,
  `track_changes.py`, `insert_text.py`, `delete_text.py`.
- All CLIs auto-detect the format from the file extension (revision CLIs accept `.docx` only).
- `tests/` — smoke tests; `evals/run_evals.py` runs them all. `dev/` — Windows+Office-only fixture
  tooling (not shipped).
- Requires `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only to read cell values /
  slide titles and in tests.

---
## COMMENTS

## 4. List comments — ALWAYS do this first to get ids
    python scripts/list_comments.py --file FILE
    python scripts/list_comments.py --file FILE --json
`--json` emits `{"comments": [...]}`; each record has `id`, `thread_id`, `parent_id`, `is_reply`,
`author`, `date`, `text`, `resolved`, `anchor`, `anchor_text`, `context`, `location`. Read it to
pick the right `id` before replying/resolving/deleting.

## 5. Add / reply / resolve / delete a comment
    python scripts/add_comment.py --file FILE --author NAME --text "..." <anchor>
    python scripts/reply_comment.py --file FILE --parent ID --author NAME --text "..."
    python scripts/resolve_comment.py --file FILE --comment ID [--reopen]
    python scripts/delete_comment.py --file FILE --comment ID
Anchor for `add_comment` by format:
- Word: `--anchor-text "phrase" [--occurrence N]`  OR  `--paragraph N`
- Excel: `--cell B2 [--sheet "Sheet name"]`  (sheet defaults to the first)
- PowerPoint: `--slide N`
`add` prints the new comment id; if a Word anchor phrase appears more than once the tool reports the
count — re-run with `--occurrence`. Reply `ID` may be the thread root or any comment in it. Resolve
affects the whole thread. Deleting a reply removes just it; deleting a root removes the whole thread.

---
## TRACKED CHANGES (Word .docx only)

## 6. List revisions — ALWAYS do this first to get ids
    python scripts/list_revisions.py --file FILE.docx
    python scripts/list_revisions.py --file FILE.docx --json
`--json` emits `{"revisions": [...]}`; each record has `id`, `author`, `date`, `type`, `text`,
`paragraph`, `location`, `context`, `move_name`, `table`. **`id` is positional (document order)** —
after you accept or reject a change the ids renumber, so re-list before targeting the next one.
`type` is one of: insertion, deletion, run-format, paragraph-format, inserted/deleted-paragraph-mark,
inserted/deleted-row, inserted/deleted-cell, cell-merge, cell/row/table-format, move, math-deletion,
numbering.

## 7. Accept / reject changes
    python scripts/accept_change.py --file FILE.docx --rev ID          # one change (by list id)
    python scripts/accept_change.py --file FILE.docx --all [--author NAME]
    python scripts/reject_change.py --file FILE.docx (--rev ID | --all) [--author NAME]
Accept keeps a change; reject undoes it. `--all` applies to every change (optionally only those by
one `--author`) and prints how many. Accept/reject reproduce exactly what Word's own Accept/Reject
produces for that revision type.

## 8. Turn Track Changes on / off
    python scripts/track_changes.py --file FILE.docx (--on | --off)
Flips the document's `<w:trackRevisions/>` switch — the same toggle as Word's *Review > Track
Changes* button. It does not accept or reject any existing changes.

## 9. Author a tracked insertion / deletion
    python scripts/insert_text.py --file FILE.docx --author NAME --text "..." <anchor>
    python scripts/delete_text.py --file FILE.docx --author NAME <anchor>
Anchor with `--anchor-text "phrase" [--occurrence N]` or `--paragraph N`. `insert_text` adds the
text as a tracked insertion right after the anchor phrase (or at the end of the anchored paragraph);
`delete_text` marks the anchored phrase (or the paragraph's text) as a tracked deletion — the text
stays in the file as a redline until accepted. Both turn Track Changes on automatically.

## 10. Per-format notes
- **Word comments**: comment + each reply anchor to a text span; ids are small integers (`0`, `1`…).
  Adding a comment upgrades the document to modern compatibility mode (`compatibilityMode=15`) when
  it's older — Word disables **Resolve** in Compatibility Mode. No-op for already-modern files.
- **Excel comments**: cell-anchored; the skill keeps the modern thread AND the legacy "note" copy
  Excel needs to display a comment at all; ids are GUIDs.
- **PowerPoint comments**: slide-anchored, replies nest inside the comment; ids are GUIDs.
- **Word tracked changes**: live inline in `word/document.xml` plus a one-line switch in
  `settings.xml` — no separate parts, so accept/reject/author touch only those two files.

## 11. Hard rules
- Never hand-edit the zip — always go through the scripts; they preserve every unrelated part.
- Target a comment/revision by `id` from `list_comments` / `list_revisions`, never by guessing.
  Revision ids are positional — re-list after each accept/reject.
- Author name is required for add/reply and for authoring tracked insertions/deletions; default to
  the current user unless told otherwise.
- Verify generated files in real desktop Office before shipping any change (see RELEASING.md).

## 12. Known limitations (v0.2.1)
- **Excel comments**: one thread per cell (reply to extend it). The skill refuses, with a clear
  message, to touch a sheet that already holds *classic* (non-threaded) comments or non-comment
  form-control drawings, rather than risk overwriting them.
- **Word comments**: resolve/reopen needs a *modern threaded* comment; classic (pre-2013) comments
  can't be resolved.
- **Tracked changes are Word-only** (Excel's revision system is deprecated; PowerPoint has none).
  **Authoring** covers insertions and deletions of run content; reading + accept/reject cover every
  revision type Word produces in full. Rejecting a tracked *whole-column* add/remove restores a
  valid, grid-consistent table but may normalise column widths slightly differently from Word's
  layout engine (the table shape is preserved).
- These are safe behaviours (no data loss), surfaced as clear errors where relevant.

## 13. Versioning & releases
The version lives in `SKILL.md` (`metadata.version`) and `scripts/_ooxml_zip.py` (`__version__`);
`scripts/release.py` asserts they match. See RELEASING.md.
