"""Word (.docx) tracked changes (revisions / redlines) — list / accept / reject / toggle / author.

Unlike comments (which live in their own parts inside the zip), tracked changes live INLINE in
the document's STORY parts — the body (`word/document.xml`) plus any headers, footers, footnotes
and endnotes — with a single on/off switch `<w:trackRevisions/>` in `word/settings.xml`. All
stories are scanned and edited (v0.3.0); only the parts actually touched are re-serialized.

Every accept/reject rule below was checked against REAL Microsoft Word (build 16.0.x): for each
revision type we drove Word to produce it, then compared our result to what Word's own
Accept / Reject produces. Where Word's COM automation won't emit a type (individual cell
insert/delete, cell merge, deleted math control), we hand-build the markup per ECMA-376 and
verify it the same way — Word's Accept/Reject is the oracle.

Detection of the "overloaded" elements: `w:ins` / `w:del` are *content wrappers* when their parent
is an ordinary content container (a paragraph, a cell, a moveFrom/To …), but *marker leaves* when
their parent is a properties element — `w:rPr` (a paragraph mark), `w:trPr` (a table row), etc.
The `*Change` family (`w:rPrChange`, `w:pPrChange`, `w:tblPrChange`, `w:tblGridChange`,
`w:tblPrExChange`, `w:tcPrChange`, `w:trPrChange`, `w:sectPrChange`) all share one shape — the old
properties nested inside — so one generic handler covers them.
"""
from __future__ import annotations

import copy
import re
import zipfile

from lxml import etree

from _ooxml_zip import patch_parts, abs_target, _rels_name
from _errors import RevisionNotFound, AnchorNotFound, AmbiguousAnchor
from _util import iso_z
from _docx_anchor import (W, _w, _set_preserve, _find_phrase, _isolate, _para_runs,
                          _paragraphs, _child_of, _text_segments)

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
DOC = "word/document.xml"
SETTINGS = "word/settings.xml"

# Story parts beyond the body, discovered through document.xml's relationships.
_REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
_STORY_RELS = (("header", _REL_BASE + "header"), ("footer", _REL_BASE + "footer"),
               ("footnotes", _REL_BASE + "footnotes"), ("endnotes", _REL_BASE + "endnotes"))


def _story_parts(get_bytes):
    """Every story part that can hold revisions, as [(part, story)] in a FIXED order:
    the document body, then headers (natural name order), footers, footnotes, endnotes.
    Only parts that actually exist are returned; duplicate targets (sections sharing one
    header part) are deduped. This order defines the global revision-id numbering."""
    stories = [(DOC, "document")]
    raw = get_bytes(_rels_name(DOC))
    buckets = {rel_type: [] for _, rel_type in _STORY_RELS}
    if raw:
        for rel in etree.fromstring(raw):
            if rel.get("Type") in buckets:
                buckets[rel.get("Type")].append(abs_target(DOC, rel.get("Target") or ""))

    def _key(name):
        m = re.search(r"(\d+)", name.rsplit("/", 1)[-1])
        return (int(m.group(1)) if m else 0, name)

    seen = {DOC}
    for story, rel_type in _STORY_RELS:
        for part in sorted(set(buckets[rel_type]), key=_key):
            if part not in seen and get_bytes(part) is not None:
                seen.add(part)
                stories.append((part, story))
    return stories


def _enumerate_stories(get_root, stories):
    """Yield (part, story, root, units) per story. Units keep their per-story document
    order but their `id`s are renumbered CONSECUTIVELY across the fixed story order, so
    one positional id space spans body + headers + footers + notes."""
    gid = 0
    for part, story in stories:
        root = get_root(part)
        units = _iter_units(root)
        for u in units:
            gid += 1
            u["id"] = gid
            u["part"] = part
            u["story"] = story
        yield part, story, root, units

# The `*Change` property-revision elements: parent is the properties element, and the old
# properties are nested inside (same local name as the parent, minus the "Change" suffix).
PROP_CHANGES = {"rPrChange", "pPrChange", "tblPrChange", "tblGridChange", "tblPrExChange",
                "tcPrChange", "trPrChange", "sectPrChange"}
# Optional properties containers that Word drops once they are emptied (e.g. a reject that
# restores "no formatting"). Required ones (tcPr, tblPr, tblGrid, sectPr) are kept.
PRUNABLE = {"rPr", "pPr", "trPr", "tblPrEx"}

_TYPE_LABEL = {
    "ins": "insertion", "del": "deletion", "move": "move",
    "ins-paramark": "inserted-paragraph-mark", "del-paramark": "deleted-paragraph-mark",
    "ins-row": "inserted-row", "del-row": "deleted-row",
    "cell-ins": "inserted-cell", "cell-del": "deleted-cell", "cell-merge": "cell-merge",
    "numbering-change": "numbering-change", "del-mathctrl": "math-control-deletion",
    "del-mathrun": "math-deletion", "ins-mathrun": "math-insertion", "ins-numbering": "numbering",
}
_PROP_LABEL = {"rPrChange": "run-format", "pPrChange": "paragraph-format",
               "tblPrChange": "table-format", "tblGridChange": "table-format",
               "tblPrExChange": "table-format", "tcPrChange": "cell-format",
               "trPrChange": "row-format", "sectPrChange": "section-format"}


def _ln(el):
    return etree.QName(el).localname


def _m(t):
    return f"{{{M}}}{t}"


# ---------------------------------------------------------------------------
# small tree operations
# ---------------------------------------------------------------------------

def _remove(el):
    parent = el.getparent()
    if parent is not None:
        parent.remove(el)


def _unwrap(el):
    """Replace `el` with its children, in place (promotes ALL children — runs, math, nested)."""
    parent = el.getparent()
    if parent is None:
        return
    idx = parent.index(el)
    for i, child in enumerate(list(el)):
        parent.insert(idx + i, child)
    parent.remove(el)


def _restore_deleted_text(scope):
    """Turn a deletion's `w:delText` back into `w:t` and `w:delInstrText` back into `w:instrText`
    (used when rejecting a deletion, so the text/field code comes back)."""
    for t in scope.iter(_w("delText")):
        t.tag = _w("t")
    for t in scope.iter(_w("delInstrText")):
        t.tag = _w("instrText")


def _strip_rsid_del(scope):
    """Drop w:rsidDel marks when a deletion is rejected: the content is no longer deleted, and a
    leftover rsidDel makes Word treat those runs as a dangling revision on reopen (surfaced with a
    restored field, which Word then re-updated under still-on tracking)."""
    for el in scope.iter():
        el.attrib.pop(_w("rsidDel"), None)


def _maybe_prune(el):
    if el is not None and _ln(el) in PRUNABLE and len(el) == 0 and not (el.text or "").strip():
        _remove(el)


def _owner(el, tag):
    p = el.getparent()
    while p is not None and _ln(p) != tag:
        p = p.getparent()
    return p


def _merge_next(p):
    """Join paragraph `p` into the following paragraph (used when a paragraph mark is removed):
    p's runs move to the front of the next paragraph, then p is dropped. No-op (returns False) if
    there is no following paragraph in the same container (e.g. the last pilcrow of a cell/body)."""
    if p is None:
        return False
    parent = p.getparent()
    nxt = p.getnext()
    while nxt is not None and _ln(nxt) != "p":
        nxt = nxt.getnext()
    if nxt is None or nxt.getparent() is not parent:
        return False
    npr = nxt.find(_w("pPr"))
    at = (list(nxt).index(npr) + 1) if npr is not None else 0
    for i, c in enumerate(c for c in list(p) if _ln(c) != "pPr"):
        nxt.insert(at + i, c)
    parent.remove(p)
    return True


def _remove_mark_leaf(leaf):
    """Remove a paragraph-mark ins/del leaf, pruning an rPr/pPr it leaves empty."""
    rpr = leaf.getparent()
    _remove(leaf)
    if rpr is not None and _ln(rpr) == "rPr":
        ppr = rpr.getparent()
        _maybe_prune(rpr)
        _maybe_prune(ppr)


# ---------------------------------------------------------------------------
# property-change (generic) + cells
# ---------------------------------------------------------------------------

def _reject_prop_change(change):
    """Restore the old properties stored inside a `*Change` element, then drop the change.
    On a paragraph's pPr we preserve the mark's own rPr (if it carries a separate ins/del
    revision) and any sectPr, since those are not part of the paragraph *formatting*."""
    parent = change.getparent()
    old = change.find(_w(_ln(change)[:-6]))   # strip "Change"
    keep = []
    if _ln(parent) == "pPr":
        cur_rpr = parent.find(_w("rPr"))
        if cur_rpr is not None and (cur_rpr.find(_w("ins")) is not None or
                                    cur_rpr.find(_w("del")) is not None):
            keep.append(cur_rpr)
        cur_sect = parent.find(_w("sectPr"))
        if cur_sect is not None:
            keep.append(cur_sect)
    for c in list(parent):
        parent.remove(c)
    if old is not None:
        for c in list(old):
            parent.append(c)
    for k in keep:
        for dup in parent.findall(k.tag):
            parent.remove(dup)
        parent.append(k)
    _maybe_prune(parent)


def _cell_span(tc):
    tcpr = tc.find(_w("tcPr"))
    gs = tcpr.find(_w("gridSpan")) if tcpr is not None else None
    v = gs.get(_w("val")) if gs is not None else None
    return int(v) if v and v.isdigit() else 1


def _cell_width(tc):
    tcpr = tc.find(_w("tcPr"))
    tw = tcpr.find(_w("tcW")) if tcpr is not None else None
    v = tw.get(_w("w")) if tw is not None else None
    return int(v) if v and v.lstrip("-").isdigit() else 0


def _grow_cell(tc, add_span, add_width):
    """Widen a cell to absorb a removed neighbour's grid column(s) + width, so the row keeps
    spanning the full table grid (exactly how Word rebalances a row after deleting a cell)."""
    tcpr = tc.find(_w("tcPr"))
    if tcpr is None:
        return
    new_span = _cell_span(tc) + add_span
    if new_span > 1:
        gs = tcpr.find(_w("gridSpan"))
        if gs is None:
            gs = etree.Element(_w("gridSpan"))
            tw = tcpr.find(_w("tcW"))
            (tw.addnext(gs) if tw is not None else tcpr.insert(0, gs))
        gs.set(_w("val"), str(new_span))
    tw = tcpr.find(_w("tcW"))
    if tw is not None and add_width:
        tw.set(_w("w"), str(_cell_width(tc) + add_width))


def _remove_cell(tc):
    """Remove a table cell, keeping the row grid-consistent by handing its grid span + width to a
    neighbouring cell (previous if any, else next). Falls back to a plain remove for a lone cell."""
    tr = _owner(tc, "tr")
    if tr is None:
        _remove(tc)
        return
    cells = tr.findall(_w("tc"))
    idx = cells.index(tc) if tc in cells else -1
    neighbor = cells[idx - 1] if idx > 0 else (cells[idx + 1] if 0 <= idx < len(cells) - 1 else None)
    if neighbor is not None:
        _grow_cell(neighbor, _cell_span(tc), _cell_width(tc))
    _remove(tc)


def _apply_cell_merge(el, action):
    """`w:cellMerge` records a tracked vertical merge: @w:vMerge = the new state, @w:vMergeOrig =
    the state to restore on reject. Accept applies the new vMerge; reject restores the original."""
    tcpr = el.getparent()
    want = el.get(_w("vMerge")) if action == "accept" else el.get(_w("vMergeOrig"))
    for existing in tcpr.findall(_w("vMerge")):
        tcpr.remove(existing)
    _remove(el)
    if want is not None:
        vm = etree.SubElement(tcpr, _w("vMerge"))
        vm.set(_w("val"), want)


# ---------------------------------------------------------------------------
# revision enumeration
# ---------------------------------------------------------------------------

def _in_tracked_row(el):
    tr = _owner(el, "tr")
    if tr is None:
        return False
    trpr = tr.find(_w("trPr"))
    return trpr is not None and (trpr.find(_w("ins")) is not None or trpr.find(_w("del")) is not None)


def _kind(el):
    """Classify a revision-marker element, or return None if `el` is not a primary marker."""
    tag = _ln(el)
    parent = el.getparent()
    ptag = _ln(parent) if parent is not None else None
    if tag == "ins":
        if parent is not None and parent.tag == _m("r"):
            return "ins-mathrun"
        if ptag == "rPr":
            return "ins-paramark"
        if ptag == "trPr":
            return "ins-row"
        if ptag == "tcPr":
            return "cell-ins"
        if ptag == "numPr":
            return "ins-numbering"
        return "ins"
    if tag == "del":
        if parent is not None and parent.tag == _m("r"):
            return "del-mathrun"
        if ptag == "rPr":
            gp = parent.getparent()
            if gp is not None and _ln(gp) == "ctrlPr":
                return "del-mathctrl"
            return "del-paramark"
        if ptag == "ctrlPr":
            return "del-mathctrl"
        if ptag == "trPr":
            return "del-row"
        if ptag == "tcPr":
            return "cell-del"
        return "del"
    if tag in PROP_CHANGES:
        return "prop-change"
    if tag == "cellIns":
        return "cell-ins"
    if tag == "cellDel":
        return "cell-del"
    if tag == "cellMerge":
        return "cell-merge"
    if tag == "numberingChange":
        return "numbering-change"
    return None


def _scan_moves(doc_root):
    """Group a move's scattered elements (moveFrom/To + their range markers) by move name.
    A move is ONE logical revision even though Word writes six-plus elements for it."""
    fromname = {s.get(_w("id")): s.get(_w("name")) for s in doc_root.iter(_w("moveFromRangeStart"))}
    toname = {s.get(_w("id")): s.get(_w("name")) for s in doc_root.iter(_w("moveToRangeStart"))}
    moves = {}
    cur_from = cur_to = None
    for el in doc_root.iter():
        tag = _ln(el)
        if tag == "moveFromRangeStart":
            cur_from = el.get(_w("name"))
            mv = moves.setdefault(cur_from, {"from": [], "to": [], "markers": [], "anchor": el,
                                             "author": el.get(_w("author")), "date": el.get(_w("date"))})
            mv["markers"].append(el)
        elif tag == "moveFromRangeEnd":
            nm = fromname.get(el.get(_w("id")))
            if nm in moves:
                moves[nm]["markers"].append(el)
            cur_from = None
        elif tag == "moveToRangeStart":
            cur_to = el.get(_w("name"))
            mv = moves.setdefault(cur_to, {"from": [], "to": [], "markers": [], "anchor": el,
                                           "author": el.get(_w("author")), "date": el.get(_w("date"))})
            mv["markers"].append(el)
        elif tag == "moveToRangeEnd":
            nm = toname.get(el.get(_w("id")))
            if nm in moves:
                moves[nm]["markers"].append(el)
            cur_to = None
        elif tag == "moveFrom":
            if cur_from in moves:
                moves[cur_from]["from"].append(el)
        elif tag == "moveTo":
            if cur_to in moves:
                moves[cur_to]["to"].append(el)
    return moves


def _iter_units(doc_root):
    """Every tracked change as a list of "units" in document order. Each unit is one accept/reject
    target. Moves are grouped by name; a table row's cell paragraph-marks are folded into the row."""
    moves = _scan_moves(doc_root)
    by_anchor = {id(mv["anchor"]): (name, mv) for name, mv in moves.items()}
    emitted_moves = set()
    move_children = set()
    for mv in moves.values():
        for group in (mv["from"], mv["to"], mv["markers"]):
            for e in group:
                move_children.add(id(e))

    units = []
    for el in doc_root.iter():
        if id(el) in by_anchor:
            name, mv = by_anchor[id(el)]
            if name not in emitted_moves:
                emitted_moves.add(name)
                units.append({"kind": "move", "name": name, "el": el, "move": mv,
                              "author": mv["author"], "date": mv["date"]})
            continue
        if id(el) in move_children:
            continue
        kind = _kind(el)
        if kind is None:
            continue
        if kind in ("ins-paramark", "del-paramark") and _in_tracked_row(el):
            continue   # folded into the row unit
        units.append({"kind": kind, "el": el,
                      "author": el.get(_w("author")), "date": el.get(_w("date"))})
    for i, u in enumerate(units, 1):
        u["id"] = i
    return units


# ---------------------------------------------------------------------------
# accept / reject one unit
# ---------------------------------------------------------------------------

def _apply(unit, action):
    kind, el = unit["kind"], unit.get("el")
    if kind == "move":
        mv = unit["move"]
        if action == "accept":
            for e in mv["from"]:
                _remove(e)
            for e in mv["to"]:
                _unwrap(e)
        else:
            for e in mv["from"]:
                _strip_rsid_del(e)
                _unwrap(e)
            for e in mv["to"]:
                _remove(e)
        for e in mv["markers"]:
            _remove(e)
    elif kind == "ins":
        _unwrap(el) if action == "accept" else _remove(el)
    elif kind == "del":
        if action == "accept":
            _remove(el)
        else:
            _restore_deleted_text(el)
            _strip_rsid_del(el)
            _unwrap(el)
    elif kind == "prop-change":
        _remove(el) if action == "accept" else _reject_prop_change(el)
    elif kind == "ins-paramark":
        _remove_mark_leaf(el) if action == "accept" else _merge_next(_owner(el, "p"))
    elif kind == "del-paramark":
        _merge_next(_owner(el, "p")) if action == "accept" else _remove_mark_leaf(el)
    elif kind == "ins-row":
        if action == "accept":
            _accept_row(el, "ins")
        else:
            _remove(_owner(el, "tr"))
    elif kind == "del-row":
        if action == "accept":
            _remove(_owner(el, "tr"))
        else:
            _accept_row(el, "del")   # reject a deleted row == keep it: strip the del leaves
    elif kind == "cell-ins":
        _remove(el) if action == "accept" else _remove_cell(_owner(el, "tc"))
    elif kind == "cell-del":
        if action == "accept":
            _remove_cell(_owner(el, "tc"))
        else:
            tc = _owner(el, "tc")
            _remove(el)
            _strip_rsid_del(tc)
    elif kind == "cell-merge":
        _apply_cell_merge(el, action)
    elif kind == "del-mathrun":
        _remove(el.getparent()) if action == "accept" else _unwrap(el)
    elif kind == "ins-mathrun":
        _unwrap(el) if action == "accept" else _remove(el.getparent())
    elif kind == "ins-numbering":
        _remove(el) if action == "accept" else _remove(_owner(el, "numPr") or el)
    elif kind in ("numbering-change", "del-mathctrl"):
        _remove(el)   # legacy / control markers: both accept and reject just drop the marker


def _accept_row(leaf, which):
    """Keep a tracked row (accept an insertion / reject a deletion): remove the trPr ins|del leaf
    and the matching paragraph-mark leaves Word puts in each cell, pruning containers left empty."""
    tr = _owner(leaf, "tr")
    trpr = leaf.getparent()
    _remove(leaf)
    _maybe_prune(trpr)
    if tr is None:
        return
    for rpr in list(tr.iter(_w("rPr"))):
        for m in rpr.findall(_w(which)):
            _remove(m)
        ppr = rpr.getparent()
        _maybe_prune(rpr)
        _maybe_prune(ppr)
    if which == "del":
        _strip_rsid_del(tr)


# ---------------------------------------------------------------------------
# records (list_revisions)
# ---------------------------------------------------------------------------

def _text_in(el):
    return "".join(x.text or "" for x in el.iter() if _ln(x) in ("t", "delText"))


def _record(unit, doc_root, paras):
    el = unit["el"]
    kind = unit["kind"]
    if kind == "move":
        text = _text_in(unit["move"]["from"][0]) if unit["move"]["from"] else ""
        typ = "move"
    else:
        typ = _PROP_LABEL[_ln(el)] if kind == "prop-change" else _TYPE_LABEL.get(kind, kind)
        text = _text_in(el)
    p = _owner(el, "p")
    para_idx = (paras.index(p) + 1) if (p is not None and p in paras) else None
    context = "".join(t.text or "" for t in p.iter(_w("t"), _w("delText"))) if p is not None else ""
    tr, tc = _owner(el, "tr"), _owner(el, "tc")
    table = None
    if tr is not None:
        row_i = None
        tbl = _owner(tr, "tbl")
        if tbl is not None:
            rows = tbl.findall(_w("tr"))
            row_i = rows.index(tr) + 1 if tr in rows else None
        col_i = None
        if tc is not None:
            cells = tr.findall(_w("tc"))
            col_i = cells.index(tc) + 1 if tc in cells else None
        table = {"row": row_i, "col": col_i}
    if not text and p is not None:
        text = context
    story = unit.get("story", "document")
    part = unit.get("part", DOC)
    if para_idx:
        location = f'para {para_idx}' + (f' (table r{table["row"]}c{table["col"]})' if table else '')
    else:
        location = "(document)" if story == "document" else f"({story})"
    if story != "document":
        location = f'{part.rsplit("/", 1)[-1]}: {location}'
    return {
        "id": unit["id"],
        "author": unit["author"],
        "date": unit["date"],
        "type": typ,
        "text": text,
        "paragraph": para_idx,
        "part": part,
        "story": story,
        "location": location,
        "context": context,
        "move_name": unit.get("name"),
        "table": table,
    }


def list_revisions(path) -> list[dict]:
    """Every tracked change in every story (body, headers, footers, footnotes, endnotes)
    as a flat record. `id` is a 1-based positional index across the fixed story order,
    used by accept/reject; each record also carries author, date, a human `type`, the
    affected `text`, `part`/`story` saying where it lives, the paragraph index +
    surrounding `context`, and table coordinates for table revisions."""
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        stories = _story_parts(lambda n: z.read(n) if n in names else None)
        roots = {part: etree.fromstring(z.read(part)) for part, _ in stories}
    out = []
    for part, _story, root, units in _enumerate_stories(roots.__getitem__, stories):
        paras = _paragraphs(root)
        out.extend(_record(u, root, paras) for u in units)
    return out


# ---------------------------------------------------------------------------
# settings.xml toggle
# ---------------------------------------------------------------------------

# Elements that follow <w:trackRevisions/> in the CT_Settings sequence (ECMA-376). We insert the
# toggle just before the first of these that is present, else append — keeping settings.xml valid.
_AFTER_TRACK = (
    "defaultTabStop", "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone",
    "doNotHyphenateCaps", "showEnvelope", "summaryLength", "clickAndTypeStyle", "defaultTableStyle",
    "evenAndOddHeaders", "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "displayHorizontalDrawingGridEvery",
    "displayVerticalDrawingGridEvery", "doNotUseMarginsForDrawingGridOrigin",
    "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin", "doNotShadeFormData",
    "noPunctuationKerning", "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
    "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture", "doNotValidateAgainstSchema",
    "saveInvalidXml", "ignoreMixedContent", "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml",
    "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags",
    "alwaysMergeEmptyNamespace", "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr",
    "compat", "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade", "captions",
    "readModeInkLockDown", "smartTagType", "shapeDefaults", "doNotEmbedSmartTags", "decimalSymbol",
    "listSeparator", "docId", "discardImageEditingData", "defaultImageDpi", "chartTrackingRefBased",
)


def _set_tracking_root(root, on):
    existing = root.find(_w("trackRevisions"))
    if on:
        if existing is not None:
            return False
        tr = etree.Element(_w("trackRevisions"))
        follower = next((root.find(_w(t)) for t in _AFTER_TRACK if root.find(_w(t)) is not None), None)
        if follower is not None:
            follower.addprevious(tr)
        else:
            root.append(tr)
        return True
    if existing is not None:
        root.remove(existing)
        return True
    return False


def _ensure_tracking_on(ps) -> None:
    """Make sure `<w:trackRevisions/>` is on — creating word/settings.xml (plus its
    content type and document relationship) when the package has none. Shared by
    set_tracking(on) and every tracked-authoring path, so authoring enables Track
    Changes even on a file with no settings part (#4)."""
    if ps.has(SETTINGS):
        root = ps.get_xml(SETTINGS)
        if _set_tracking_root(root, True):
            ps.set_xml(SETTINGS, root)
    else:
        root = etree.Element(_w("settings"), nsmap={"w": W})
        root.append(etree.Element(_w("trackRevisions")))
        ps.add_xml_part(SETTINGS, root)
        ps.ensure_content_type(
            "/" + SETTINGS,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
        ps.rels_for(DOC).add_rel(_REL_BASE + "settings", "settings.xml")


def set_tracking(path, on: bool) -> None:
    """Turn Word's Track Changes recording on or off (the `<w:trackRevisions/>` switch)."""
    def mut(ps):
        if on:
            _ensure_tracking_on(ps)
        elif ps.has(SETTINGS):
            root = ps.get_xml(SETTINGS)
            if _set_tracking_root(root, False):
                ps.set_xml(SETTINGS, root)
    patch_parts(path, mut)


# ---------------------------------------------------------------------------
# accept / reject
# ---------------------------------------------------------------------------

def _attached(el, root):
    p = el
    while p is not None:
        if p is root:
            return True
        p = p.getparent()
    return False


def _one(path, rev_id, action):
    def mut(ps):
        target, total = None, 0
        for part, _story, root, units in _enumerate_stories(ps.get_xml, _story_parts(ps.get_bytes)):
            total += len(units)
            if target is None:
                u = next((x for x in units if x["id"] == rev_id), None)
                if u is not None:
                    target = (part, root, u)
        if target is None:
            raise RevisionNotFound(f"no revision with id {rev_id} (found {total})")
        part, root, unit = target
        _apply(unit, action)
        ps.set_xml(part, root)
    patch_parts(path, mut)


def accept(path, rev_id) -> None:
    """Accept one tracked change (by its `list_revisions` id)."""
    _one(path, int(rev_id), "accept")


def reject(path, rev_id) -> None:
    """Reject one tracked change (by its `list_revisions` id)."""
    _one(path, int(rev_id), "reject")


def _all(path, action, author):
    out = {"n": 0}

    def mut(ps):
        for part, _story, root, units in _enumerate_stories(ps.get_xml, _story_parts(ps.get_bytes)):
            if author:
                units = [u for u in units if (u["author"] or "") == author]
            # bottom-up + attachment guard: removing/merging later revisions can't disturb
            # earlier ones, and any element already detached by a structural change is skipped.
            applied = False
            for u in reversed(units):
                primary = u["move"]["anchor"] if u["kind"] == "move" else u["el"]
                if not _attached(primary, root):
                    continue
                _apply(u, action)
                applied = True
            out["n"] += len(units)
            if applied:
                ps.set_xml(part, root)
    patch_parts(path, mut)
    return out["n"]


def accept_all(path, *, author=None) -> int:
    """Accept every tracked change (optionally only those by `author`). Returns how many."""
    return _all(path, "accept", author)


def reject_all(path, *, author=None) -> int:
    """Reject every tracked change (optionally only those by `author`). Returns how many."""
    return _all(path, "reject", author)


# ---------------------------------------------------------------------------
# authoring: create tracked insertions / deletions
# ---------------------------------------------------------------------------

def _next_rev_id(roots):
    """Next free revision id across EVERY story root passed in — w:id values stay unique
    document-wide even when body, headers and notes all carry revisions."""
    ids = []
    for doc_root in roots:
        for el in doc_root.iter():
            if _ln(el) in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange",
                           "cellIns", "cellDel", "cellMerge") or _ln(el) in PROP_CHANGES:
                v = el.get(_w("id"))
                if v and v.lstrip("-").isdigit():
                    ids.append(int(v))
    return (max(ids) + 1) if ids else 0


def _match_part(stories, wanted):
    """Resolve a user-facing part filter — 'header1', 'header1.xml', 'word/header1.xml',
    'document', 'footnotes', 'endnotes' — to one story part name."""
    w = (wanted or "").strip().lower()
    for part, story in stories:
        base = part.rsplit("/", 1)[-1].lower()
        candidates = {part.lower(), base, base.rsplit(".", 1)[0]}
        if story in ("document", "footnotes", "endnotes"):
            candidates.add(story)
        if w in candidates:
            return part
    have = ", ".join(p.rsplit("/", 1)[-1] for p, _ in stories)
    raise AnchorNotFound(f"no story part matching {wanted!r}; this document has: {have}")


def _anchor_in_stories(ps, anchor):
    """Resolve an anchor across every story part; returns (part, root, paragraph,
    first_run, last_run). `{"paragraph": N}` targets the document body unless
    `anchor["part"]` pins another story; `{"text": ...}` searches all stories in the
    fixed story order (or just the pinned one) — occurrences are numbered across that
    traversal, and an ambiguous phrase reports where each match lives."""
    stories = _story_parts(ps.get_bytes)
    if anchor.get("part"):
        pinned = _match_part(stories, anchor["part"])
        stories = [(p, s) for p, s in stories if p == pinned]
    if anchor.get("paragraph") is not None:
        part = stories[0][0]                     # pinned story, else the document body
        root = ps.get_xml(part)
        paras = _paragraphs(root)
        n = anchor["paragraph"]
        if not (1 <= n <= len(paras)):
            raise RevisionNotFound(f"paragraph {n} out of range (1..{len(paras)})")
        runs = _para_runs(paras[n - 1])
        if not runs:
            raise RevisionNotFound(f"paragraph {n} has no text to anchor to")
        return part, root, paras[n - 1], runs[0][0], runs[-1][0]
    text, occurrence = anchor["text"], anchor.get("occurrence")
    matches, counts = [], {}
    for part, _story in stories:
        root = ps.get_xml(part)
        for p in _paragraphs(root):
            whole = "".join((t.text or "") for _, t, _ in _text_segments(p))
            i = whole.find(text)
            while i != -1:
                matches.append((part, root, p, i, i + len(text)))
                counts[part] = counts.get(part, 0) + 1
                i = whole.find(text, i + 1)
    if not matches:
        raise AnchorNotFound(f"anchor text not found: {text!r}")
    if occurrence is None and len(matches) > 1:
        where = ", ".join(f'{c} in {p.rsplit("/", 1)[-1]}' for p, c in counts.items())
        raise AmbiguousAnchor(f"{len(matches)} matches for {text!r} ({where}); "
                              f"pass an occurrence (1..{len(matches)}) or a part")
    if occurrence is not None and not (1 <= occurrence <= len(matches)):
        raise AnchorNotFound(f"occurrence {occurrence} out of range (1..{len(matches)}) for {text!r}")
    part, root, p, s, e = matches[(occurrence or 1) - 1]
    first, last = _isolate(p, s, e)
    return part, root, p, first, last


def _all_story_roots(ps):
    return [ps.get_xml(part) for part, _ in _story_parts(ps.get_bytes)]


def _insert_tracked_in(ps, anchor, text, *, author, date=None, before=False):
    """PartSet-level core of insert_tracked (batch edits call this directly).
    Returns the attached `w:ins` element."""
    part, root, p, first, last = _anchor_in_stories(ps, anchor)
    ins = etree.Element(_w("ins"))
    ins.set(_w("id"), str(_next_rev_id(_all_story_roots(ps))))
    ins.set(_w("author"), author)
    ins.set(_w("date"), date or iso_z())
    r = etree.SubElement(ins, _w("r"))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    _set_preserve(t)
    # Hoist out of a hyperlink/content control: inserted text lands BESIDE the
    # container (Word's own behaviour — links don't grow), not inside it.
    if before:
        _child_of(p, first).addprevious(ins)
    else:
        _child_of(p, last).addnext(ins)
    ps.set_xml(part, root)
    _ensure_tracking_on(ps)
    return ins


def _boundary_error(paragraph_anchor):
    if paragraph_anchor:
        return AnchorNotFound(
            "this paragraph mixes plain text with a hyperlink/content control; delete or "
            "replace its pieces by phrase instead (the text before the link, then the "
            "link text, then the text after)")
    return AnchorNotFound(
        "anchor spans across a hyperlink/content-control boundary; delete or replace "
        "the pieces separately (the text before the link, the link text, the text after)")


def _span_runs(first, last, *, paragraph_anchor=False):
    """The run elements from `first` to `last` — refusing LOUDLY when the span crosses a
    container boundary: first/last under different parents, OR any element BETWEEN them
    that holds visible text (a hyperlink/content control/tracked insertion sitting inside
    the span would otherwise be silently skipped, leaving its text behind while the
    command reports success)."""
    parent = first.getparent()
    if parent is not last.getparent():
        raise _boundary_error(paragraph_anchor)
    sibs = list(parent)
    i0, i1 = sibs.index(first), sibs.index(last)
    window = sibs[i0:i1 + 1]
    for el in window:
        if el.tag != _w("r") and el.find(f".//{_w('t')}") is not None:
            raise _boundary_error(paragraph_anchor)
    return [el for el in window if el.tag == _w("r")]


def _delete_tracked_in(ps, anchor, *, author, date=None):
    """PartSet-level core of delete_tracked. Returns the attached `w:del` element."""
    part, root, p, first, last = _anchor_in_stories(ps, anchor)
    span = _span_runs(first, last, paragraph_anchor=anchor.get("paragraph") is not None)
    dl = etree.Element(_w("del"))
    dl.set(_w("id"), str(_next_rev_id(_all_story_roots(ps))))
    dl.set(_w("author"), author)
    dl.set(_w("date"), date or iso_z())
    first.addprevious(dl)
    for r in span:
        _remove(r)
        dl.append(r)
        for t in r.findall(_w("t")):
            t.tag = _w("delText")
            _set_preserve(t)
    ps.set_xml(part, root)
    _ensure_tracking_on(ps)
    return dl


def _replace_tracked_in(ps, anchor, text, *, author, date=None):
    """One tracked REPLACE: the anchored phrase becomes a deletion and the replacement
    text an insertion right after it — consecutive revision ids, formatting inherited
    from the phrase's first run. Exactly what a human editing with Track Changes on
    produces (Word shows it as 'replaced X with Y')."""
    dl = _delete_tracked_in(ps, anchor, author=author, date=date)
    ins = etree.Element(_w("ins"))
    # the w:del is already attached, so the id minted here is consecutive with its
    ins.set(_w("id"), str(_next_rev_id(_all_story_roots(ps))))
    ins.set(_w("author"), author)
    ins.set(_w("date"), date or iso_z())
    r = etree.SubElement(ins, _w("r"))
    src = dl.find(_w("r"))
    rpr = src.find(_w("rPr")) if src is not None else None
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, _w("t"))
    t.text = text
    _set_preserve(t)
    dl.addnext(ins)
    return ins


def insert_tracked(path, anchor, text, *, author, date=None) -> None:
    """Insert `text` as a tracked insertion, immediately after the anchored phrase (or at the end
    of the anchored paragraph) — in whichever story the anchor lives (body, header, footer,
    footnote, endnote). Turns Track Changes on as well."""
    patch_parts(path, lambda ps: _insert_tracked_in(ps, anchor, text, author=author, date=date))


def delete_tracked(path, anchor, *, author, date=None) -> None:
    """Mark the anchored phrase (or the anchored paragraph's text) as a tracked deletion —
    in whichever story the anchor lives."""
    patch_parts(path, lambda ps: _delete_tracked_in(ps, anchor, author=author, date=date))


def replace_tracked(path, anchor, text, *, author, date=None) -> None:
    """Replace the anchored phrase with `text` as ONE tracked change (deletion + insertion
    pair) — in whichever story the anchor lives. Turns Track Changes on as well."""
    patch_parts(path, lambda ps: _replace_tracked_in(ps, anchor, text, author=author, date=date))
