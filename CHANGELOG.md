# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/) and this project adheres to
[Semantic Versioning](https://semver.org/).

## [0.3.0] - 2026-07-19

### Added
- **Tracked changes now cover the whole document, not just the body.** `list_revisions` scans
  headers, footers, footnotes and endnotes too; every record says where it lives (`part`/`story`),
  and accept/reject — single or `--all` — work in every story. (#8)
- **Author tracked edits anywhere.** `insert_text` / `delete_text` (and the new commands below)
  find their anchor phrase across every story; when a phrase appears in more than one place the
  error reports the per-part counts, and `--part header1`-style pinning narrows the search.
- **`replace_text`** — "change X to Y" as ONE tracked change (a deletion + insertion pair with
  consecutive ids and inherited formatting), instead of two separate commands.
- **`edit_text`** — direct (untracked) replace / insert / delete: the file changes cleanly with
  nothing else touched, no revision markup, no author.
- **`apply_edits`** — a whole JSON list of edits, comments, replies and resolves applied in one
  pass. Sequential (later anchors see earlier results), **all-or-nothing** (any failure leaves
  the file untouched and names the failing entry), tracked-by-default, with `--dry-run`.
- **Excel: classic notes and threaded comments now coexist.** Sheets holding old-style notes are
  no longer refused: the notes (rows AND their drawing shapes) are preserved verbatim through
  every operation, and only the specific cell that holds a note refuses a new thread — Excel's
  own one-per-cell rule, said clearly. Sheets with form controls stay refused (their macro-wired
  drawing layer cannot be regenerated safely).
- New SKILL.md rules for agents: never rebuild a document to make edits (use the skill even for
  many changes); when addressing review comments, reply into the SAME thread by id — never fake
  a conversation; always ASK the user which author name to use; and report genuine bugs or
  capability gaps back to the user so they can be logged.

### Fixed
- **Word: text inside hyperlinks is now anchorable** for comments and tracked edits — it was
  invisible to the phrase search ("anchor text not found" on plainly visible text). Inserted
  text lands beside the link, as Word itself does. (#5)
- **Word: runs holding several text pieces (line breaks, tabs mid-run) now anchor correctly** —
  previously only the first piece was seen, text after a break was unfindable, and deleting text
  just before a break silently swallowed the break. Splits now cut precisely at text boundaries. (#6)
- **Word: deleting a classic (pre-2013) comment no longer destroys unrelated modern threads.**
  The old walk mis-seeded and removed every modern thread while leaving the target; a classic
  comment is now removed directly, and resolve/reply on one give a clear error instead of a
  misleading "not found". (#7)
- **Word: authoring a tracked edit on a file with no settings part now turns Track Changes on**
  by creating the part (content type + relationship included), as `track_changes --on` already
  did. (#4)
- **Excel: invalid cell references like "A0" are rejected upfront** (columns A..XFD, rows
  1..1,048,576) instead of writing broken geometry into the file; "a01" is canonicalized to "A1". (#3)
- **Excel: deleting the last threaded comment on a sheet that also holds classic notes no longer
  destroys the notes** — the teardown now rebuilds a notes-only layer instead of removing it
  unconditionally.
- Every save now writes to a temp file and swaps it into place — a crash mid-save can no longer
  leave a half-written Office file (falls back to in-place write when another program holds the
  file open for reading).

### Changed
- Phrase-occurrence counts can differ from v0.2.x when the phrase also appears inside hyperlinks
  or headers — those occurrences are now genuinely findable (that is the fix); pass
  `--occurrence`/`--part` when a phrase matches more than once.
- SKILL.md description and body rewritten around the wider scope (comments + tracked changes +
  edits + batch) with explicit triggers like "respond to review comments" and "apply review
  feedback".

## [0.2.1] - 2026-07-02

### Fixed
- **Excel: adding more than one comment to a sheet no longer corrupts the file.** Each legacy
  "note" was written with a per-comment `shapeId` (0, 1, 2, …), but those point at VML drawing
  shapes that don't exist (real shape ids start at 1025), so Excel discarded the entire comments
  part on open ("Removed Records: Comments") — leaving at most one thread. Every legacy note now
  uses `shapeId="0"`, exactly as Excel itself writes. Single-comment files were unaffected (their
  one shapeId happened to be the valid `0`). Verified in desktop Excel with many comments across
  multiple sheets; a multi-comment regression test now guards it.
- **Excel: no duplicate relationships when a file's comment wiring came from another tool.** Adding
  a comment now reuses the sheet's existing comments/VML relationship whatever its target spelling,
  instead of appending a second one when the spelling differed from the skill's `../…` (a duplicate
  relationship also made Excel drop the comments). The delete path unwires by relationship type too,
  so nothing is left dangling. Guarded by a regression test.
- **Excel: the VML note drawing stays valid past 1023 comments on a single sheet** — the shape-id
  block map is now computed from the notes present, not hardcoded.

## [0.2.0] - 2026-07-01

### Added
- **Word tracked changes (revisions / redlines).** List every tracked change; accept or reject one
  or all (optionally filtered by author); turn Track Changes on/off; and author tracked insertions
  and deletions. Word (`.docx`) only — Excel's revision system is deprecated and PowerPoint has none.
  Six new CLIs: `list_revisions`, `accept_change`, `reject_change`, `track_changes`, `insert_text`,
  `delete_text`.
- Every revision type Word produces is handled fully for **both accept and reject** — insertions,
  deletions (including deleted fields), run / paragraph / table-property changes, inserted and
  deleted paragraph marks, tracked table rows and cells, cell merges, moves, tracked equations, and
  legacy numbering changes. Each type's accept/reject was verified to match Microsoft Word's own
  Accept/Reject on a real-Word round-trip.
- Shared Word anchoring primitives extracted to `scripts/_docx_anchor.py`, used by both the comment
  and tracked-change modules.

### Changed
- **Renamed `office-comments` → `office-markup`** to reflect the broader scope (comments *and*
  tracked changes). The GitHub repository was renamed to match; the old URL redirects.

### Fixed
- **Excel**: the `<legacyDrawing>` relationship id now serialises as `r:id` instead of `ns0:id`,
  which some Excel builds could mishandle. Guarded by a smoke-test assertion. (#1)

## [0.1.0] - 2026-06-27

### Added
- Initial release. Read / add / reply / resolve-reopen / delete **modern threaded comments** in
  Word (`.docx`), Excel (`.xlsx`) and PowerPoint (`.pptx`).
- Binary-safe OOXML zip-patch engine (`_ooxml_zip.py`) that edits only the comment-related parts
  and preserves every other part byte-for-byte.
- Five extension-dispatching CLIs (`list` / `add` / `reply` / `resolve` / `delete`) with a
  machine-readable `--json` listing.
- Per-format support verified against real Microsoft 365 Office:
  - Word — five comment parts + anchor markers; threading via `paraIdParent`; resolve via
    `w15:done`. (Fix: `paraId`/`durableId` are kept below `0x80000000` — Word treats them as
    signed 32-bit ints and won't thread otherwise.)
  - Excel — modern `threadedComments` + the required legacy `comments`/VML shadow, regenerated
    from the thread after every change.
  - PowerPoint — modern `p188` comments with nested `replyLst`; slide anchor via `sldMk`; resolve
    via `status="resolved"`.

### Fixed
- **Word "Resolve" now works.** Adding a comment upgrades the document's `compatibilityMode` to 15
  when it's older (e.g. python-docx's default of 14). Word *disables* the "Resolve thread" option
  in Compatibility Mode, so without this a user couldn't resolve/reopen threads in Word (Excel and
  PowerPoint were unaffected). Verified live in desktop Word. `set_status` also now marks the whole
  thread done (root + replies), matching Word's own behaviour.

### Fixed (code review)
- **Word**: anchoring to a phrase inside a run that also holds a `w:tab` / `w:br` / inline image no
  longer duplicates that element when the run is split.
- **Word**: `list`'s reported paragraph index and `add --paragraph N` now use the same enumeration,
  so they round-trip even when the document contains tables.
- **Excel**: deleting the last comment on a sheet now tears down the legacy shadow completely (parts,
  `<legacyDrawing>`, relationships, content types) instead of leaving a dangling reference Excel could
  flag for repair.
- **Excel**: refuses (with a clear message) to add a second comment to a cell that already has a
  thread, and refuses to modify a sheet whose legacy parts hold classic comments or non-comment
  drawings (form controls) — preventing silent data loss — rather than overwriting them.
- **PowerPoint**: a freshly minted slide `creationId` is now kept unique across the deck (and non-zero).
- Hardened relationship-target resolution against malformed `..` segments.
