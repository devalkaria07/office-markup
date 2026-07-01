"""Shared Word (.docx) anchoring primitives.

Locating an anchor phrase inside `word/document.xml` and splitting runs so a chosen character
span is covered by *whole* runs. These operate purely on the w:p / w:r / w:t run model and know
nothing about comments or revisions, so both features build on them:

    _docx_comments   wraps the isolated span in commentRangeStart/End markers
    _docx_revisions  wraps the isolated span in a w:ins / w:del element

Kept in one place so the run-splitting (the fiddly, well-tested part) has a single implementation.
"""
from __future__ import annotations

import copy

from lxml import etree

from _errors import AnchorNotFound, AmbiguousAnchor

# --- namespaces ---
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"


def _w(t):
    return f"{{{W}}}{t}"


def _set_preserve(t):
    t.set(f"{{{XML}}}space", "preserve")


def _para_runs(p):
    return [(r, r.find(_w("t"))) for r in p.findall(_w("r")) if r.find(_w("t")) is not None]


def _split_run(r, t, at):
    """Split run `r` at local offset `at`: the left slice (text[:at]) becomes a NEW run carrying
    only the run properties + that text, inserted before `r`; `r` keeps the right slice.

    Building the left run from just rPr + a fresh w:t (rather than deep-copying the whole run)
    avoids duplicating any non-text children the run may also hold — a `w:tab`, `w:br`, inline
    `w:drawing`, footnote ref, etc. — which a naive deepcopy would clone into both halves."""
    full = t.text or ""
    left = etree.Element(_w("r"))
    rpr = r.find(_w("rPr"))
    if rpr is not None:
        left.append(copy.deepcopy(rpr))
    lt = etree.SubElement(left, _w("t"))
    lt.text = full[:at]
    _set_preserve(lt)
    t.text = full[at:]
    _set_preserve(t)
    r.addprevious(left)


def _isolate(p, start, end):
    """Split runs so the [start, end) character span is covered by whole runs; return
    (first_run, last_run) of that span."""
    pos = 0
    for r, t in _para_runs(p):
        ln = len(t.text or "")
        if pos <= start < pos + ln:
            if start - pos > 0:
                _split_run(r, t, start - pos)
            break
        pos += ln
    pos = 0
    for r, t in _para_runs(p):
        ln = len(t.text or "")
        if pos < end <= pos + ln:
            if end - pos < ln:
                _split_run(r, t, end - pos)
            break
        pos += ln
    inside = []
    pos = 0
    for r, t in _para_runs(p):
        ln = len(t.text or "")
        if pos >= start and pos + ln <= end and ln:
            inside.append(r)
        pos += ln
    if not inside:
        raise AnchorNotFound("could not isolate the anchor text within the paragraph")
    return inside[0], inside[-1]


def _paragraphs(doc_root):
    """All paragraphs in document order (including those inside tables / text boxes). Used by both
    `list` (to report the paragraph index) and `add --paragraph N`, so the two refer to the same
    enumeration and round-trip even when the document contains tables."""
    body = doc_root.find(_w("body"))
    return list(body.iter(_w("p"))) if body is not None else []


def _find_phrase(doc_root, text, occurrence):
    """Locate the anchor text. Returns (paragraph, start, end). `occurrence` is 1-based;
    if None and the text matches more than once, raise AmbiguousAnchor."""
    matches = []
    for p in _paragraphs(doc_root):
        whole = "".join(t.text or "" for _, t in _para_runs(p))
        i = whole.find(text)
        while i != -1:
            matches.append((p, i, i + len(text)))
            i = whole.find(text, i + 1)
    if not matches:
        raise AnchorNotFound(f"anchor text not found: {text!r}")
    if occurrence is None:
        if len(matches) > 1:
            raise AmbiguousAnchor(f"{len(matches)} matches for {text!r}; pass an occurrence (1..{len(matches)})")
        return matches[0]
    if not (1 <= occurrence <= len(matches)):
        raise AnchorNotFound(f"occurrence {occurrence} out of range (1..{len(matches)}) for {text!r}")
    return matches[occurrence - 1]
