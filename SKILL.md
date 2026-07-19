---
name: office-markup
description: |
  Read and write Microsoft Office review markup: modern THREADED comments in Word (.docx),
  Excel (.xlsx) and PowerPoint (.pptx) — list, add, reply, resolve/reopen, delete — and Word
  TRACKED CHANGES: list/accept/reject every revision across body, headers, footers and
  footnotes; toggle recording; author tracked insertions, deletions and replacements in any
  of those stories. Also DIRECT (untracked) text edits, and a batch command applying a
  JSON list of edits/comments in one call. Surgically patches only the relevant XML inside
  the OOXML zip — every other part stays byte-identical. Excel classic notes are preserved
  alongside threads. Run list_comments / list_revisions first, then act by id. Triggers:
  comment, reply, respond to / address review comments, resolve; tracked changes, redline,
  revisions, accept/reject changes; edit / replace / insert / delete text in Word; apply
  review feedback, batch of many edits, who changed what — in Word / Excel / PowerPoint or
  .docx/.xlsx/.pptx files. NOT for PDF or Google Docs.
metadata:
  version: 0.3.0
  author: Deval Karia
---

# office-markup

## 1. Overview
Reads and writes **Office markup** by surgically editing only the relevant XML parts inside
the file's zip and copying every other part byte-for-byte — the rest of the document is
never touched. Works on ANY Office file. Three capabilities:

1. **Modern threaded comments** — conversation-style comments (a comment plus replies, which
   can be marked Resolved), in Word (`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`).
2. **Word tracked changes** (revisions / redlines) — every insert / delete / format /
   structure edit Word records, across ALL stories: body, headers, footers, footnotes,
   endnotes. List, accept/reject (one or all), toggle recording, and AUTHOR tracked
   insertions, deletions and replacements. **Word (`.docx`) only.**
3. **Word direct edits + batch** — the same replace/insert/delete WITHOUT tracking (a clean
   edit, nothing else disturbed), and `apply_edits`: a whole JSON list of edits, comments,
   replies and resolves applied in one all-or-nothing pass.

**Division of labour:** this skill is the *hands* — it finds comments/changes and reports them
with their context, and makes the exact edit you name. The agent is the *brain* — decide which
item is meant, what to say or change, and where.

## 2. When to use
- **Comments** (any Office file): see what comments exist, add one, reply, resolve/reopen, or
  delete. Triggers: "what are the comments", "add a comment", "reply to …", "respond to /
  address the review comments", "resolve/close this", "reopen".
- **Tracked changes** (Word only): see every tracked change, accept/reject some or all, turn
  Track Changes on/off, or author a tracked insertion/deletion/replacement. Triggers: "track
  changes", "what changed", "accept/reject changes", "redline", "who edited this".
- **Edits** (Word only): "change X to Y" (tracked or clean), "fix this typo without touching
  anything else", "make these 30 edits", "apply this review feedback".

Do NOT use for: non-Office files, PDF, or Google Docs.

## 3. What's bundled
- `scripts/_ooxml_zip.py` — the binary-safe zip-patch engine (atomic writes; preserves all
  other parts).
- Comments: `scripts/_docx_comments.py`, `_xlsx_comments.py`, `_pptx_comments.py`.
- Tracked changes: `scripts/_docx_revisions.py`. Direct edits: `scripts/_docx_edit.py`.
  Batch: `scripts/_docx_batch.py`. Shared anchoring: `scripts/_docx_anchor.py`.
- Comment CLIs: `list_comments.py`, `add_comment.py`, `reply_comment.py`,
  `resolve_comment.py`, `delete_comment.py`.
- Tracked-change CLIs: `list_revisions.py`, `accept_change.py`, `reject_change.py`,
  `track_changes.py`, `insert_text.py`, `delete_text.py`, `replace_text.py`.
- Edit CLIs: `edit_text.py` (direct), `apply_edits.py` (batch).
- All CLIs auto-detect the format from the file extension (Word-only CLIs accept `.docx` only).
- `tests/` — smoke tests; `evals/run_evals.py` runs them all. `dev/` — Windows+Office-only
  fixture tooling (not shipped).
- Requires `lxml`. `python-docx` / `openpyxl` / `python-pptx` are used only to read cell
  values / slide titles and in tests.

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
Scans EVERY story: body, then headers, footers, footnotes, endnotes (in that fixed order).
`--json` emits `{"revisions": [...]}`; each record has `id`, `author`, `date`, `type`, `text`,
`paragraph`, `part`, `story`, `location`, `context`, `move_name`, `table`. `part`/`story` say
where the change lives (e.g. `word/header1.xml` / `header`); non-body locations are prefixed
with the part name. **`id` is positional (document order across ALL stories)** — after you
accept or reject a change the ids renumber, so re-list before targeting the next one.
`type` is one of: insertion, deletion, run-format, paragraph-format, inserted/deleted-paragraph-mark,
inserted/deleted-row, inserted/deleted-cell, cell-merge, cell/row/table-format, move, math-deletion,
numbering.

## 7. Accept / reject changes
    python scripts/accept_change.py --file FILE.docx --rev ID          # one change (by list id)
    python scripts/accept_change.py --file FILE.docx --all [--author NAME]
    python scripts/reject_change.py --file FILE.docx (--rev ID | --all) [--author NAME]
Accept keeps a change; reject undoes it. `--all` sweeps every story (optionally only changes by
one `--author`) and prints how many. Accept/reject reproduce exactly what Word's own
Accept/Reject produces for that revision type.

## 8. Turn Track Changes on / off
    python scripts/track_changes.py --file FILE.docx (--on | --off)
Flips the document's `<w:trackRevisions/>` switch — the same toggle as Word's *Review > Track
Changes* button. Creates the settings part when a file lacks one. It does not accept or reject
any existing changes.

## 9. Author a tracked insertion / deletion / replacement
    python scripts/insert_text.py  --file F.docx --author NAME --text "..." <anchor>
    python scripts/delete_text.py  --file F.docx --author NAME <anchor>
    python scripts/replace_text.py --file F.docx --author NAME --text "NEW" <anchor>
Anchor with `--anchor-text "phrase" [--occurrence N]` or `--paragraph N`, plus optional
`--part NAME` (e.g. `header1`, `footer2`, `footnotes`) to pin one story. The phrase search
covers EVERY story; when a phrase matches in more than one place the error reports the count
per part — pass `--occurrence` (numbered across the story order) or `--part`. `replace_text`
is ONE tracked change-X-to-Y (a deletion + insertion pair, formatting inherited — exactly what
a human edit with Track Changes on produces). All three turn Track Changes on automatically.

---
## DIRECT EDITS & BATCH (Word .docx only)

## 10. Direct (untracked) edits
    python scripts/edit_text.py --file F.docx <anchor> --replace-with "NEW"
    python scripts/edit_text.py --file F.docx <anchor> --insert-after " more"
    python scripts/edit_text.py --file F.docx <anchor> --insert-before ">> "
    python scripts/edit_text.py --file F.docx <anchor> --delete
The same anchoring (any story, `--part` supported), but the change is applied PLAINLY: no
revision markup, no author, settings untouched, nothing else in the file disturbed. Direct
edits never create redlines even when Track Changes is on. Use these when the user wants a
clean correction rather than a reviewable change.

## 11. Batch: many edits in one call
    python scripts/apply_edits.py --file F.docx --edits edits.json [--author NAME] [--dry-run] [--json]
`edits.json`:

    {
      "author": "Alex Morgan",
      "edits": [
        {"action": "replace", "anchor": {"text": "Q3 figures", "occurrence": 2}, "text": "Q3 (final) figures"},
        {"action": "insert",  "anchor": {"paragraph": 4}, "text": " See appendix.", "position": "after"},
        {"action": "delete",  "anchor": {"text": "DRAFT - "}, "tracked": false},
        {"action": "comment", "anchor": {"text": "revenue"}, "text": "Confirm against ledger."},
        {"action": "reply",   "comment_id": "3", "text": "Done - updated in this pass."},
        {"action": "resolve", "comment_id": "3"}
      ]
    }

Rules:
- Actions: `replace` / `insert` / `delete` (text edits), `comment` / `reply` / `resolve`.
- `tracked` defaults to **true** for text edits — reviewable is the default; a direct edit
  must say `"tracked": false` explicitly.
- Anchors take `text` (+ optional `occurrence`, `part`) or `paragraph`, like the CLIs.
  Exception: `comment` anchors search the BODY only — comments are body-anchored, so
  `part` is rejected there and `occurrence` counts body matches.
- Edits apply **sequentially in list order** — a later anchor sees earlier edits' results.
  Exception: text inside a TRACKED insertion is pending, not real, so it is NOT anchorable
  by later entries (anchor on surrounding text, or use `"tracked": false` when a later entry
  must target the new text).
- **All-or-nothing**: if any entry fails, the file is left completely untouched and the error
  names the failing entry. `--dry-run` executes in memory only and never writes.
- `reply`/`resolve` target comment ids that exist BEFORE the batch (run list_comments first).
- Author: per-edit `author` > spec top-level `author` > `--author`. Required for tracked
  edits, comments and replies — see Hard rule 3.

## 12. Per-format notes
- **Word comments**: comment + each reply anchor to a text span; ids are small integers (`0`, `1`…).
  Adding a comment upgrades the document to modern compatibility mode (`compatibilityMode=15`) when
  it's older — Word disables **Resolve** in Compatibility Mode. No-op for already-modern files.
- **Excel comments**: cell-anchored; the skill keeps the modern thread AND the legacy "note" copy
  Excel needs to display a comment at all; ids are GUIDs. **Classic (old-style) notes coexist**:
  notes already on a sheet are preserved untouched through every operation; only the specific
  cell that holds a note refuses a new thread (Excel's one-per-cell rule — the error says so).
- **PowerPoint comments**: slide-anchored, replies nest inside the comment; ids are GUIDs.
- **Word tracked changes**: live inline in the story parts (`document.xml`, `header*.xml`,
  `footer*.xml`, `footnotes.xml`, `endnotes.xml`) plus a one-line switch in `settings.xml` —
  only the parts actually touched are rewritten.

## 13. Hard rules
1. **NEVER rebuild or regenerate a document to make edits.** Even for many changes, use this
   skill (`apply_edits` for batches). Rebuilding a file with a generic library destroys
   formatting, headers, numbering, other people's comments and metadata — everything not
   explicitly re-created.
2. **Addressing existing comments:** run `list_comments` first, then reply INTO THE SAME
   THREAD by id (`reply_comment` / batch `reply`). Never post a new top-level comment that
   fakes a conversation, and never write another person's name inside your comment text.
3. **Author name: ask the user which author name to use** before the session's first comment
   or tracked edit — never assume one, never invent one, never default silently. (Direct
   edits carry no author.)
4. Target a comment/revision by `id` from `list_comments` / `list_revisions`, never by
   guessing. Revision ids are positional across all stories — re-list after each accept/reject.
5. Never hand-edit the zip — always go through the scripts; they preserve every unrelated part.
6. Verify generated files in real desktop Office before shipping any skill change (RELEASING.md).

## 14. Improve this skill — report genuine issues back
After the user's task is done: if, while using this skill, you hit a genuine **bug** (valid
input but an error, wrong output, or a file Office flags/repairs) or a genuine **capability
gap** that forced you into repetitive manual work, tell the user at the end of your reply in
2–4 lines: the command you ran, the kind of file, what happened, and what you expected — so
they can log it at `github.com/devalkaria07/office-markup`. Do NOT report cosmetic
preferences, speculative feature ideas, or anything at all when everything worked.

## 15. Known limitations (v0.3.0)
- **Excel**: one thread per cell (reply to extend it); a cell holding a classic note refuses a
  new thread (the note is preserved). Sheets whose legacy drawing layer holds **form controls**
  (checkboxes, buttons) are refused entirely — regenerating around them risks Excel dropping
  the controls.
- **Word comments**: resolve/reopen/reply need a *modern threaded* comment; classic (pre-2013)
  comments can be listed and deleted, nothing more.
- **Tracked changes are Word-only** (Excel's revision system is deprecated; PowerPoint has
  none). Anchors match only untracked text — text inside an existing tracked insertion or
  deletion is not anchorable. A delete/replace (tracked or direct) whose target CROSSES a
  hyperlink/content-control boundary — a link at either end of the phrase, a link in the
  middle of it, or a paragraph anchor on a paragraph mixing plain text with a link — is
  refused with guidance: edit the pieces separately (the text before the link, the link
  text, the text after).
- **PowerPoint legacy comments** (pre-"modern", `p:cm`): invisible to the skill — list shows
  none, and adding creates a modern part alongside them. Rare; documented, not handled.
- `edit_text` / `apply_edits` are Word-only.
- These are safe behaviours (no data loss), surfaced as clear errors where relevant.

## 16. Versioning & releases
The version lives in `SKILL.md` (`metadata.version`) and `scripts/_ooxml_zip.py` (`__version__`);
`scripts/release.py` asserts they match. See RELEASING.md.
