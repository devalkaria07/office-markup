"""Release tooling for the office-comments skill.

A release is a tagged, packaged ``dist/office-comments-<version>.skill`` (a zip whose entries
are prefixed with ``office-comments/`` — matching how the skill is installed).

Single source of truth for the version: ``SKILL.md`` frontmatter ``metadata.version``.
``scripts/_ooxml_zip.py`` ``__version__`` MUST match it; this script asserts that and refuses to
package on drift.

Usage:
    python scripts/release.py --check        # version sync + frontmatter + evals
    python scripts/release.py --package      # --check, then build the .skill artifact
    python scripts/release.py                # same as --check
"""
from __future__ import annotations

import argparse
import re
import subprocess
import sys
import textwrap
import zipfile
from pathlib import Path

SKILL_DIR = Path(__file__).resolve().parent.parent
SKILL_MD = SKILL_DIR / "SKILL.md"
ENGINE_PY = SKILL_DIR / "scripts" / "_ooxml_zip.py"
DIST_DIR = SKILL_DIR / "dist"
PKG_PREFIX = "office-comments"

# `dev/` holds Windows+Office-only fixture tooling and must never ship.
_EXCLUDE_DIRS = {".git", "__pycache__", "dist", "build", ".pytest_cache", "dev"}
_EXCLUDE_SUFFIXES = {".pyc", ".pyo", ".skill", ".zip"}
_EXCLUDE_NAMES = {".gitignore", ".gitattributes", ".DS_Store", "Thumbs.db"}


def _read_skill_version() -> str:
    text = SKILL_MD.read_text(encoding="utf-8")
    parts = text.split("---", 2)
    front = parts[1] if len(parts) >= 3 else text
    m = re.search(r"^\s*version:\s*([^\s#]+)", front, re.MULTILINE)
    if not m:
        raise SystemExit("error: could not find metadata.version in SKILL.md")
    return m.group(1).strip()


def _read_engine_version() -> str:
    m = re.search(r"""__version__\s*=\s*["'](.+?)["']""", ENGINE_PY.read_text(encoding="utf-8"))
    if not m:
        raise SystemExit("error: could not find __version__ in scripts/_ooxml_zip.py")
    return m.group(1).strip()


def check_versions() -> str:
    sv, ev = _read_skill_version(), _read_engine_version()
    if sv != ev:
        raise SystemExit(f"error: version drift — SKILL.md={sv!r} but _ooxml_zip.__version__={ev!r}. "
                         f"Bump both (see RELEASING.md).")
    print(f"version sync OK: {sv}")
    return sv


def check_frontmatter() -> None:
    text = SKILL_MD.read_text(encoding="utf-8")
    parts = text.split("---", 2)
    if len(parts) < 3:
        raise SystemExit("error: SKILL.md has no YAML frontmatter")
    front = parts[1]
    if not re.search(r"^name:\s*\S", front, re.M):
        raise SystemExit("error: SKILL.md frontmatter missing 'name'")
    m = re.search(r"^description:\s*\|?\s*\n(.*?)(?=^\S|\Z)", front, re.S | re.M)
    desc = textwrap.dedent(m.group(1)).strip() if m else ""
    if not desc:
        m2 = re.search(r"^description:[ \t]*(.+)$", front, re.M)
        desc = m2.group(1).strip() if m2 else ""
    if not desc:
        raise SystemExit("error: SKILL.md frontmatter missing 'description'")
    if len(desc) > 1024:
        raise SystemExit(f"error: SKILL.md description is {len(desc)} chars (max 1024)")
    print(f"frontmatter OK: description {len(desc)} chars (<=1024)")


def run_evals() -> None:
    runner = SKILL_DIR / "evals" / "run_evals.py"
    if not runner.is_file():
        print("note: no evals/run_evals.py — skipping eval gate.")
        return
    print("running eval suite ...")
    rc = subprocess.call([sys.executable, str(runner)])
    if rc != 0:
        raise SystemExit(f"error: eval suite failed (exit {rc}); release blocked.")
    print("evals passed.")


def _included(path: Path) -> bool:
    rel = path.relative_to(SKILL_DIR)
    if any(part in _EXCLUDE_DIRS for part in rel.parts):
        return False
    if path.name in _EXCLUDE_NAMES or path.suffix.lower() in _EXCLUDE_SUFFIXES:
        return False
    return True


def package(version: str) -> Path:
    DIST_DIR.mkdir(exist_ok=True)
    out = DIST_DIR / f"office-comments-{version}.skill"
    n = 0
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for path in sorted(SKILL_DIR.rglob("*")):
            if not path.is_file() or not _included(path):
                continue
            z.write(path, f"{PKG_PREFIX}/{path.relative_to(SKILL_DIR).as_posix()}")
            n += 1
    print(f"wrote {out} ({n} files)")
    print(f"next: git commit -am 'Release v{version}' && git tag v{version}")
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--check", action="store_true", help="version sync + frontmatter + evals")
    ap.add_argument("--package", action="store_true", help="run --check then build the .skill artifact")
    ap.add_argument("--version", action="store_true", help="print the skill version")
    args = ap.parse_args(argv)

    if args.version:
        print(_read_skill_version())
        return 0

    version = check_versions()
    check_frontmatter()
    run_evals()
    if args.package:
        package(version)
    return 0


if __name__ == "__main__":
    sys.exit(main())
