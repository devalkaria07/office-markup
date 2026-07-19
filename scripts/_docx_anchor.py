"""Shared Word (.docx) anchoring primitives.

Locating an anchor phrase inside a Word story part and splitting runs so a chosen character
span is covered by *whole* runs. These operate purely on the w:p / w:r / w:t model and know
nothing about comments or revisions, so both features build on them:

    _docx_comments   wraps the isolated span in commentRangeStart/End markers
    _docx_revisions  wraps the isolated span in a w:ins / w:del element

The text model (v0.3.0): a paragraph's anchorable text is the ordered list of SEGMENTS —
one per direct `w:t` child of each anchorable run, in document order. Runs are collected by
a walker that looks through *transparent* inline containers (`w:hyperlink`, `w:sdt`,
`w:smartTag`, ...) so visible link/content-control text is anchorable, but never into
revision wrappers (`w:ins`/`w:del`/moves — tracked text is not anchorable) nor into
`w:drawing`/`w:pict`/`w:object` (text-box paragraphs are enumerated separately by
`_paragraphs`; descending would double-count them). Splits happen at (text-node, offset)
boundaries and move everything strictly before the split point into the new left run, so
non-text children (`w:br`, `w:cr`, `w:tab`, inline drawings) always fall OUTSIDE a span
that does not textually cover them.

Kept in one place so the run-splitting (the fiddly, well-tested part) has a single
implementation.
"""
from __future__ import annotations

import copy

from lxml import etree

from _errors import AnchorNotFound, AmbiguousAnchor

# --- namespaces ---
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML = "http://www.w3.org/XML/1998/namespace"

# Inline containers whose runs are part of the paragraph's visible text flow. The walker
# recurses into these; everything else that isn't a plain `w:r` is skipped (which keeps
# revision wrappers, drawings, properties and foreign-namespace content invisible —
# exactly the pre-0.3.0 visibility, plus the containers listed here).
_TRANSPARENT = {"hyperlink", "smartTag", "sdt", "sdtContent", "fldSimple", "dir", "bdo",
                "customXml"}


def _w(t):
    return f"{{{W}}}{t}"


def _set_preserve(t):
    t.set(f"{{{XML}}}space", "preserve")


def _iter_anchor_runs(p):
    """Every anchorable `w:r` under paragraph `p`, in document order — direct runs plus
    runs inside transparent containers (hyperlinks, content controls, smart tags...)."""
    out = []

    def walk(el):
        for child in el:
            tag = child.tag
            if not isinstance(tag, str) or not tag.startswith("{" + W + "}"):
                continue                      # comments, PIs, foreign namespaces
            local = tag.rsplit("}", 1)[1]
            if local == "r":
                out.append(child)
            elif local in _TRANSPARENT:
                walk(child)

    walk(p)
    return out


def _text_segments(p):
    """The paragraph's text as ordered (run, w:t, length) segments — EVERY direct `w:t`
    child of every anchorable run, not just the first."""
    segs = []
    for r in _iter_anchor_runs(p):
        for t in r.findall(_w("t")):
            segs.append((r, t, len(t.text or "")))
    return segs


def _para_runs(p):
    """[(run, first w:t)] for each anchorable run that has text. Kept for consumers that
    need a paragraph's first/last text runs; offset math uses `_text_segments` instead."""
    out = []
    for r in _iter_anchor_runs(p):
        t = r.find(_w("t"))
        if t is not None:
            out.append((r, t))
    return out


def _pre_content(r, t):
    """True if run `r` has any non-rPr child strictly before text node `t`."""
    for child in r:
        if child is t:
            return False
        if child.tag != _w("rPr"):
            return True
    return False


def _post_content(r, t):
    """True if run `r` has any non-rPr child strictly after text node `t`."""
    seen = False
    for child in r:
        if seen and child.tag != _w("rPr"):
            return True
        if child is t:
            seen = True
    return False


def _split_run(r, t, at):
    """Split run `r` at offset `at` inside its text node `t`. A NEW left run — carrying a
    copy of the run properties, every child of `r` that precedes `t` (MOVED, not copied,
    so drawings/breaks are never duplicated), and text[:at] when at > 0 — is inserted
    before `r`; `r` keeps `t` (with text[at:]) and everything after it. With at == 0 this
    just peels the pre-`t` children off into the left run (a no-op if there are none)."""
    full = t.text or ""
    left = etree.Element(_w("r"))
    rpr = r.find(_w("rPr"))
    if rpr is not None:
        left.append(copy.deepcopy(rpr))
    moved = False
    for child in list(r):
        if child is t:
            break
        if child.tag == _w("rPr"):
            continue
        left.append(child)                    # lxml append MOVES the child
        moved = True
    if at > 0:
        lt = etree.SubElement(left, _w("t"))
        lt.text = full[:at]
        _set_preserve(lt)
        t.text = full[at:]
        _set_preserve(t)
        moved = True
    if moved:
        r.addprevious(left)


def _split_run_after(r, t):
    """Split run `r` right AFTER text node `t`: the new left run gets the run properties
    (copied) plus every child up to and including `t` (moved); `r` keeps the rest."""
    left = etree.Element(_w("r"))
    rpr = r.find(_w("rPr"))
    if rpr is not None:
        left.append(copy.deepcopy(rpr))
    for child in list(r):
        if child.tag == _w("rPr"):
            continue
        left.append(child)                    # moves
        if child is t:
            break
    r.addprevious(left)


def _isolate(p, start, end):
    """Split runs so the [start, end) character span is covered by whole runs; return
    (first_run, last_run) of that span. Non-text children outside the covered text stay
    outside the returned span; one sitting BETWEEN two covered text nodes of a run stays
    inside (the phrase genuinely wraps it)."""
    # Pass A — start boundary. Total text length is unchanged by splits, so absolute
    # offsets stay valid across passes; each pass walks fresh segments.
    pos = 0
    for r, t, ln in _text_segments(p):
        if pos <= start < pos + ln:
            if start - pos > 0 or _pre_content(r, t):
                _split_run(r, t, start - pos)
            break
        pos += ln
    # Pass B — end boundary.
    pos = 0
    for r, t, ln in _text_segments(p):
        if pos < end <= pos + ln:
            if end - pos < ln:
                _split_run(r, t, end - pos)
            elif _post_content(r, t):
                _split_run_after(r, t)
            break
        pos += ln
    # Pass C — collect: a run is inside iff every non-empty segment it holds lies within
    # [start, end) and it holds at least one non-empty segment.
    runs_meta = []                            # (run, all_inside, total_len) in order
    cur_run, cur_ok, cur_len = None, True, 0
    pos = 0
    for r, t, ln in _text_segments(p):
        if r is not cur_run:
            if cur_run is not None:
                runs_meta.append((cur_run, cur_ok, cur_len))
            cur_run, cur_ok, cur_len = r, True, 0
        if ln:
            cur_ok = cur_ok and (pos >= start and pos + ln <= end)
            cur_len += ln
        pos += ln
    if cur_run is not None:
        runs_meta.append((cur_run, cur_ok, cur_len))
    inside = [r for r, ok, total in runs_meta if ok and total]
    if not inside:
        raise AnchorNotFound("could not isolate the anchor text within the paragraph")
    return inside[0], inside[-1]


def _child_of(p, el):
    """The ancestor of `el` (possibly `el` itself) whose parent is `p`. Used to hoist an
    insertion out of a transparent container (hyperlink/content control) so new content
    lands beside the container, not inside it."""
    cur = el
    while cur.getparent() is not None and cur.getparent() is not p:
        cur = cur.getparent()
    return cur


def _paragraphs(doc_root):
    """All paragraphs in document order (including those inside tables / text boxes).
    Story-tolerant: a `w:document` root scopes to its `w:body`; header/footer/footnote/
    endnote roots (`w:hdr`, `w:ftr`, `w:footnotes`, `w:endnotes`) have no body and are
    walked directly. Used by both `list` (to report the paragraph index) and
    `add --paragraph N`, so the two share one enumeration."""
    body = doc_root.find(_w("body"))
    scope = body if body is not None else doc_root
    return list(scope.iter(_w("p")))


def _find_phrase(doc_root, text, occurrence):
    """Locate the anchor text. Returns (paragraph, start, end). `occurrence` is 1-based;
    if None and the text matches more than once, raise AmbiguousAnchor."""
    matches = []
    for p in _paragraphs(doc_root):
        whole = "".join((t.text or "") for _, t, _ in _text_segments(p))
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
