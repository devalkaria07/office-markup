# Releasing office-markup

Versions are SemVer. The version lives in **two** places that must always match:
- `SKILL.md` frontmatter `metadata.version`
- `scripts/_ooxml_zip.py` `__version__`

`scripts/release.py` refuses to package if they drift.

## Release steps

1. **Implement + bump** both version strings; add a dated `[X.Y.Z]` section to `CHANGELOG.md`.
2. **Green evals:** `python scripts/release.py --check` (version sync + frontmatter + all
   `tests/smoke_*.py`).
3. **Regenerate samples:** a comment / reply / resolved thread for each format, and (Word) a file
   with tracked changes — an authored insertion and deletion plus a mix of accepted/rejected changes.
4. **HOLD — verify in real desktop Office.** Do not release on automated tests alone; headless /
   library checks cannot see how Office actually renders markup.
   - **Comments** (Word, Excel, PowerPoint): the thread shows, **Reply** works, and **Resolve /
     Reopen** behave correctly — with no "repair" prompt.
   - **Tracked changes** (Word): the redlines show under *Review*; authored insertions/deletions
     carry the right author; and Word's own **Accept** / **Reject** on a copy agrees with the
     skill's — again with no "repair" prompt.
   - **Excel `r:id` fix**: open a regenerated `.xlsx` and confirm the comment shows and the file
     opens without repair.
5. **Package:** `python scripts/release.py --package` → builds `dist/office-markup-X.Y.Z.skill`.
6. **Tag + publish:** `git commit -am "Release vX.Y.Z" && git tag vX.Y.Z && git push --tags`,
   then attach the `.skill` to the GitHub release.

## Why the desktop-Office hold
These formats are Office-specific and full of details that only real Word / Excel / PowerPoint
exercise (e.g. Word silently refuses to thread a reply if an id has its high bit set; Excel will not
show a comment without its legacy "note" shadow; tracked-change accept/reject has per-type
structural rules). Every release is verified by round-tripping generated files through the actual
apps — for tracked changes, by confirming Word's own Accept / Reject matches the skill's — before
shipping.
