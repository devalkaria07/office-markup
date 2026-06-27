# Releasing office-comments

Versions are SemVer. The version lives in **two** places that must always match:
- `SKILL.md` frontmatter `metadata.version`
- `scripts/_ooxml_zip.py` `__version__`

`scripts/release.py` refuses to package if they drift.

## Release steps

1. **Implement + bump** both version strings; move the `[Unreleased]` notes into a dated
   `[X.Y.Z]` section in `CHANGELOG.md`.
2. **Green evals:** `python scripts/release.py --check` (version sync + frontmatter + all
   `tests/smoke_*.py`).
3. **Regenerate samples** for each format (a comment, a reply, a resolved thread).
4. **HOLD — verify in real desktop Office.** Open the freshly generated `.docx`, `.xlsx` and
   `.pptx` in Word, Excel and PowerPoint and confirm: the thread shows, **Reply** works, and
   **Resolve/Reopen** behave correctly — with no "repair" prompt. Do not release on automated
   tests alone; headless/library checks cannot see how Office actually renders comments.
5. **Package:** `python scripts/release.py --package` → builds `dist/office-comments-X.Y.Z.skill`.
6. **Tag + publish:** `git commit -am "Release vX.Y.Z" && git tag vX.Y.Z && git push --tags`,
   then attach the `.skill` to the GitHub release.

## Why the desktop-Office hold
These comment formats are Office-specific and full of details that only real Word/Excel/PowerPoint
exercise (e.g. Word silently refuses to thread a reply if an id has its high bit set; Excel will
not show a comment without its legacy "note" shadow). Every release is verified by round-tripping
generated files through the actual apps before shipping.
