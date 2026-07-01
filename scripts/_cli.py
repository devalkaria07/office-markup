"""Shared CLI plumbing: file-type dispatch, error -> exit-code mapping, human-readable listing.

Each per-action CLI (list/add/reply/resolve/delete) detects the format from the file
extension and dispatches to the matching module. Output is ASCII-only (Windows-console safe).
"""
from __future__ import annotations

import sys
import zipfile
from pathlib import Path

import _docx_comments
import _xlsx_comments
import _pptx_comments
import _docx_revisions
from _errors import (CommentError, UnsupportedFile, AnchorNotFound, AmbiguousAnchor,
                     CommentNotFound, RevisionNotFound, EXIT_OK, EXIT_GENERIC, EXIT_BAD_FILE,
                     EXIT_ANCHOR, EXIT_AMBIGUOUS, EXIT_NO_COMMENT, EXIT_NO_REVISION)

EXIT_OK = EXIT_OK  # re-export for the scripts

_MODULES = {".docx": _docx_comments, ".xlsx": _xlsx_comments, ".pptx": _pptx_comments}


def module_for(path):
    ext = Path(path).suffix.lower()
    mod = _MODULES.get(ext)
    if mod is None:
        raise UnsupportedFile(f"unsupported file type {ext!r}; use .docx, .xlsx or .pptx")
    if not Path(path).exists():
        raise UnsupportedFile(f"file not found: {path}")
    return mod


def docx_only(path):
    """Return the tracked-changes module for a .docx, or a clear error otherwise. Tracked changes
    are a Word-only feature (Excel's revision system is deprecated; PowerPoint has none)."""
    ext = Path(path).suffix.lower()
    if ext != ".docx":
        raise UnsupportedFile(f"tracked changes are a Word-only feature; {ext!r} is not supported "
                              f"(Excel revisions are deprecated, PowerPoint has none)")
    if not Path(path).exists():
        raise UnsupportedFile(f"file not found: {path}")
    return _docx_revisions


def run(fn) -> int:
    """Execute fn(); translate expected errors into stable exit codes."""
    try:
        return fn()
    except AmbiguousAnchor as e:
        print(f"error: ambiguous anchor - {e}", file=sys.stderr)
        return EXIT_AMBIGUOUS
    except AnchorNotFound as e:
        print(f"error: anchor not found - {e}", file=sys.stderr)
        return EXIT_ANCHOR
    except RevisionNotFound as e:
        print(f"error: {e}", file=sys.stderr)
        return EXIT_NO_REVISION
    except CommentNotFound as e:
        print(f"error: {e}", file=sys.stderr)
        return EXIT_NO_COMMENT
    except UnsupportedFile as e:
        print(f"error: {e}", file=sys.stderr)
        return EXIT_BAD_FILE
    except zipfile.BadZipFile:
        print("error: not a valid Office file (corrupt or not a .docx/.xlsx/.pptx)", file=sys.stderr)
        return EXIT_BAD_FILE
    except CommentError as e:
        print(f"error: {e}", file=sys.stderr)
        return EXIT_GENERIC
    except Exception as e:  # noqa: BLE001 - last-resort guard for a CLI boundary
        print(f"error: {type(e).__name__}: {e}", file=sys.stderr)
        return EXIT_GENERIC


def print_human(records) -> None:
    if not records:
        print("(no comments)")
        return
    by_thread = {}
    for r in records:
        by_thread.setdefault(r["thread_id"], []).append(r)
    roots = [r for r in records if not r["is_reply"]]
    for root in roots:
        state = "RESOLVED" if root["resolved"] else "open"
        print(f"[{root['id']}] {root['location']}  ({state})")
        if root.get("anchor_text"):
            snippet = root["anchor_text"].replace("\n", " ")[:70]
            print(f"     on: {snippet!r}")
        print(f"     {root['author']}: {root['text']}")
        for r in by_thread.get(root["thread_id"], []):
            if r["is_reply"]:
                print(f"       reply [{r['id']}] {r['author']}: {r['text']}")
        print()


def print_revisions(records) -> None:
    if not records:
        print("(no tracked changes)")
        return
    for r in records:
        who = r.get("author") or "?"
        loc = r.get("location") or ""
        head = f"[{r['id']}] {r['type']}  by {who}"
        if loc:
            head += f"  ({loc})"
        print(head)
        text = (r.get("text") or "").replace("\n", " ")
        if text:
            print(f"     {text[:70]!r}" + ("..." if len(text) > 70 else ""))
