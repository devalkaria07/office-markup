"""Batch edits for Word (.docx): one JSON spec, one pass, all-or-nothing.

The whole list runs inside a single zip-patch transaction (`patch_parts`): if ANY edit
fails — bad anchor, unknown comment id, invalid spec — the exception aborts before the
file is written, so the document on disk is completely untouched, and the error names
the failing edit. Edits apply sequentially in list order against the live document
tree, exactly as if the individual commands had been run one after another (a later
anchor sees the results of earlier edits).

Spec shape (see SKILL.md for the full contract):

    {
      "author": "Alex Morgan",                 # default author for every entry
      "edits": [
        {"action": "replace", "anchor": {"text": "old"}, "text": "new"},
        {"action": "insert",  "anchor": {"paragraph": 4}, "text": " More.", "position": "after"},
        {"action": "delete",  "anchor": {"text": "DRAFT - "}, "tracked": false},
        {"action": "comment", "anchor": {"text": "revenue"}, "text": "Check this."},
        {"action": "reply",   "comment_id": "3", "text": "Done."},
        {"action": "resolve", "comment_id": "3"}
      ]
    }

`tracked` defaults to TRUE for replace/insert/delete — reviewable edits are the default;
direct edits must say `"tracked": false` explicitly. Anchors may carry "occurrence" and
"part" (story pin) exactly like the single-edit commands. `reply`/`resolve` target ids
that exist BEFORE the batch (run list_comments first).
"""
from __future__ import annotations

from _ooxml_zip import read_parts, patch_parts, PartSet
from _errors import CommentError
import _docx_comments
import _docx_edit
import _docx_revisions

ACTIONS = ("replace", "insert", "delete", "comment", "reply", "resolve")


def _edit_author(edit, spec, cli_author):
    return edit.get("author") or spec.get("author") or cli_author


def validate(spec, cli_author=None) -> list[str]:
    """Collect EVERY problem in the spec (not just the first) before anything runs."""
    if not isinstance(spec, dict):
        return ["spec must be a JSON object"]
    edits = spec.get("edits")
    if not isinstance(edits, list) or not edits:
        return ['spec must contain a non-empty "edits" list']
    errors = []
    for i, e in enumerate(edits, 1):
        tag = f"edit {i}"
        if not isinstance(e, dict):
            errors.append(f"{tag}: must be an object")
            continue
        action = e.get("action")
        if action not in ACTIONS:
            errors.append(f"{tag}: unknown action {action!r} (use one of: {', '.join(ACTIONS)})")
            continue
        tracked = e.get("tracked", True)
        if not isinstance(tracked, bool):
            errors.append(f'{tag}: "tracked" must be true or false')
            tracked = True
        if action in ("replace", "insert", "delete", "comment"):
            a = e.get("anchor")
            if not isinstance(a, dict) or not (a.get("text") or a.get("paragraph") is not None):
                errors.append(f'{tag} ({action}): needs an "anchor" with "text" or "paragraph"')
            elif action == "comment" and a.get("part"):
                errors.append(f'{tag} (comment): comments are body-anchored; "part" is not '
                              f'supported for the comment action')
        if action in ("replace", "insert", "comment") and not e.get("text"):
            errors.append(f'{tag} ({action}): needs non-empty "text"')
        if action == "insert" and e.get("position", "after") not in ("after", "before"):
            errors.append(f'{tag} (insert): "position" must be "after" or "before"')
        if action in ("reply", "resolve") and e.get("comment_id") in (None, ""):
            errors.append(f'{tag} ({action}): needs "comment_id" (from list_comments)')
        if action == "reply" and not e.get("text"):
            errors.append(f'{tag} (reply): needs non-empty "text"')
        needs_author = (action in ("comment", "reply")
                        or (action in ("replace", "insert", "delete") and tracked))
        if needs_author and not _edit_author(e, spec, cli_author):
            errors.append(f'{tag} ({action}): needs an "author" — per edit, at the spec top '
                          f'level, or via --author (never assume one: ask the user)')
    return errors


def _dispatch(ps, e, spec, cli_author):
    action = e["action"]
    tracked = e.get("tracked", True)
    author = _edit_author(e, spec, cli_author)
    date = e.get("date")
    anchor = e.get("anchor")
    if action == "replace":
        if tracked:
            _docx_revisions._replace_tracked_in(ps, anchor, e["text"], author=author, date=date)
        else:
            _docx_edit._replace_in(ps, anchor, e["text"])
    elif action == "insert":
        before = e.get("position", "after") == "before"
        if tracked:
            _docx_revisions._insert_tracked_in(ps, anchor, e["text"], author=author,
                                               date=date, before=before)
        else:
            _docx_edit._insert_in(ps, anchor, e["text"], before=before)
    elif action == "delete":
        if tracked:
            _docx_revisions._delete_tracked_in(ps, anchor, author=author, date=date)
        else:
            _docx_edit._delete_in(ps, anchor)
    elif action == "comment":
        return _docx_comments._insert(ps, text=e["text"], author=author,
                                      initials=e.get("initials"), date=date,
                                      parent_id=None, anchor=anchor)
    elif action == "reply":
        return _docx_comments._insert(ps, text=e["text"], author=author,
                                      initials=e.get("initials"), date=date,
                                      parent_id=str(e["comment_id"]), anchor=None)
    elif action == "resolve":
        _docx_comments._set_status_in(ps, str(e["comment_id"]), bool(e.get("resolved", True)))
    return None


def apply_edits(path, spec, *, cli_author=None, dry_run=False) -> dict:
    """Apply the whole spec to `path`. All-or-nothing: any failure leaves the file
    untouched and raises with the failing edit's number. With dry_run=True the edits are
    executed against an in-memory copy and NOTHING is written — a pure validity check.
    Returns {"applied": n, "results": [{"index", "action", "id"?}, ...]}."""
    errors = validate(spec, cli_author)
    if errors:
        raise CommentError("invalid edits spec:\n  - " + "\n  - ".join(errors))
    results = []

    def mut(ps):
        for i, e in enumerate(spec["edits"], 1):
            try:
                out = _dispatch(ps, e, spec, cli_author)
            except CommentError as ex:
                # same exception type (so the exit code survives), failing edit named
                raise type(ex)(f"edit {i} ({e['action']}): {ex}") from ex
            rec = {"index": i, "action": e["action"]}
            if out:
                rec["id"] = out
            results.append(rec)

    if dry_run:
        infos, data = read_parts(path)
        mut(PartSet(infos, data))          # never flushed, never written
    else:
        patch_parts(path, mut)
    return {"applied": len(results), "results": results, "dry_run": bool(dry_run)}
