"""Direct (untracked) text edits for Word (.docx) — replace / insert / delete.

The surgical counterpart to _docx_revisions' tracked authoring: the same cross-story
anchoring (body, headers, footers, footnotes, endnotes), but the change is applied
PLAINLY — no revision markup, no author attribution, settings.xml untouched. Direct
edits never create redlines, even when the document's Track Changes switch is on
(tracking records edits made through Word's editing surface, not XML surgery).
Only the story part actually touched is re-serialized; everything else stays
byte-for-byte identical.
"""
from __future__ import annotations

import copy

from lxml import etree

from _ooxml_zip import patch_parts
from _docx_anchor import _w, _set_preserve, _child_of
from _docx_revisions import _anchor_in_stories, _span_runs


def _replace_in(ps, anchor, new_text):
    """Replace the anchored phrase with `new_text`, inheriting the phrase's first run's
    formatting. Non-text children fully covered by the phrase (a break the phrase wraps)
    are removed with it — that is what 'replace the phrase' means."""
    part, root, _p, first, last = _anchor_in_stories(ps, anchor)
    span = _span_runs(first, last, paragraph_anchor=anchor.get("paragraph") is not None)
    keep = span[0]
    for r in span[1:]:
        r.getparent().remove(r)
    rpr = keep.find(_w("rPr"))
    for child in list(keep):
        if child is not rpr:
            keep.remove(child)
    t = etree.SubElement(keep, _w("t"))
    t.text = new_text
    _set_preserve(t)
    ps.set_xml(part, root)


def _insert_in(ps, anchor, text, *, before=False):
    """Insert `text` as a plain run right after (default) or before the anchored phrase,
    cloning the adjacent run's formatting. Hoisted outside hyperlinks/content controls so
    new text lands beside the container, not inside it."""
    part, root, p, first, last = _anchor_in_stories(ps, anchor)
    edge = first if before else last
    r = etree.Element(_w("r"))
    rpr = edge.find(_w("rPr"))
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    _set_preserve(t)
    top = _child_of(p, edge)
    (top.addprevious if before else top.addnext)(r)
    ps.set_xml(part, root)


def _delete_in(ps, anchor):
    """Remove the anchored phrase outright (with any non-text children it fully covers)."""
    part, root, _p, first, last = _anchor_in_stories(ps, anchor)
    for r in _span_runs(first, last, paragraph_anchor=anchor.get("paragraph") is not None):
        r.getparent().remove(r)
    ps.set_xml(part, root)


def replace_direct(path, anchor, new_text) -> None:
    """Replace the anchored phrase with `new_text` — a clean, untracked edit."""
    patch_parts(path, lambda ps: _replace_in(ps, anchor, new_text))


def insert_direct(path, anchor, text, *, before=False) -> None:
    """Insert `text` right after (or before) the anchored phrase — untracked."""
    patch_parts(path, lambda ps: _insert_in(ps, anchor, text, before=before))


def delete_direct(path, anchor) -> None:
    """Delete the anchored phrase — untracked."""
    patch_parts(path, lambda ps: _delete_in(ps, anchor))
