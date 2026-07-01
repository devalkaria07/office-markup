# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/) and this project adheres to
[Semantic Versioning](https://semver.org/).

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
