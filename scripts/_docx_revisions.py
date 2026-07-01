"""Word (.docx) tracked changes (revisions / redlines) — list / accept / reject / toggle / author.

Unlike comments (which live in their own parts inside the zip), tracked changes live INLINE in
`word/document.xml`, with a single on/off switch `<w:trackRevisions/>` in `word/settings.xml`.
So this module edits only those two parts and reuses the same byte-preserving zip engine.

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

import zipfile

from lxml import etree

from _ooxml_zip import patch_parts
from _errors import RevisionNotFound
from _util import iso_z
from _docx_anchor import W, _w, _set_preserve, _find_phrase, _isolate, _para_runs, _paragraphs

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
DOC = "word/document.xml"
SETTINGS = "word/settings.xml"

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
    return {
        "id": unit["id"],
        "author": unit["author"],
        "date": unit["date"],
        "type": typ,
        "text": text,
        "paragraph": para_idx,
        "location": (f'para {para_idx}' + (f' (table r{table["row"]}c{table["col"]})' if table else '')
                     if para_idx else "(document)"),
        "context": context,
        "move_name": unit.get("name"),
        "table": table,
    }


def list_revisions(path) -> list[dict]:
    """Every tracked change as a flat record. `id` is a 1-based document-order index used by
    accept/reject; each record also carries author, date, a human `type`, the affected `text`,
    the paragraph index + surrounding `context`, and table coordinates for table revisions."""
    with zipfile.ZipFile(path) as z:
        doc_root = etree.fromstring(z.read(DOC))
    paras = _paragraphs(doc_root)
    return [_record(u, doc_root, paras) for u in _iter_units(doc_root)]


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


def set_tracking(path, on: bool) -> None:
    """Turn Word's Track Changes recording on or off (the `<w:trackRevisions/>` switch)."""
    def mut(ps):
        if ps.has(SETTINGS):
            root = ps.get_xml(SETTINGS)
            if _set_tracking_root(root, on):
                ps.set_xml(SETTINGS, root)
        elif on:
            root = etree.Element(_w("settings"), nsmap={"w": W})
            root.append(etree.Element(_w("trackRevisions")))
            ps.add_xml_part(SETTINGS, root)
            ps.ensure_content_type(
                "/" + SETTINGS,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
            ps.rels_for(DOC).add_rel(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
                "settings.xml")
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
        doc_root = ps.get_xml(DOC)
        units = _iter_units(doc_root)
        unit = next((u for u in units if u["id"] == rev_id), None)
        if unit is None:
            raise RevisionNotFound(f"no revision with id {rev_id} (found {len(units)})")
        _apply(unit, action)
        ps.set_xml(DOC, doc_root)
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
        doc_root = ps.get_xml(DOC)
        units = _iter_units(doc_root)
        if author:
            units = [u for u in units if (u["author"] or "") == author]
        # bottom-up + attachment guard: removing/merging later revisions can't disturb earlier ones,
        # and any element already detached by a structural change is skipped.
        for u in reversed(units):
            primary = u["move"]["anchor"] if u["kind"] == "move" else u["el"]
            if not _attached(primary, doc_root):
                continue
            _apply(u, action)
        out["n"] = len(units)
        ps.set_xml(DOC, doc_root)
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

def _next_rev_id(doc_root):
    ids = []
    for el in doc_root.iter():
        if _ln(el) in ("ins", "del", "moveFrom", "moveTo", "rPrChange", "pPrChange",
                       "cellIns", "cellDel", "cellMerge") or _ln(el) in PROP_CHANGES:
            v = el.get(_w("id"))
            if v and v.lstrip("-").isdigit():
                ids.append(int(v))
    return (max(ids) + 1) if ids else 0


def _anchor_runs(doc_root, anchor):
    """Resolve an anchor (like add_comment) to (paragraph, first_run, last_run)."""
    if anchor.get("paragraph") is not None:
        paras = _paragraphs(doc_root)
        n = anchor["paragraph"]
        if not (1 <= n <= len(paras)):
            raise RevisionNotFound(f"paragraph {n} out of range (1..{len(paras)})")
        runs = _para_runs(paras[n - 1])
        if not runs:
            raise RevisionNotFound(f"paragraph {n} has no text to anchor to")
        return paras[n - 1], runs[0][0], runs[-1][0]
    p, s, e = _find_phrase(doc_root, anchor["text"], anchor.get("occurrence"))
    first, last = _isolate(p, s, e)
    return p, first, last


def insert_tracked(path, anchor, text, *, author, date=None) -> None:
    """Insert `text` as a tracked insertion, immediately after the anchored phrase (or at the end
    of the anchored paragraph). Turns Track Changes on as well."""
    def mut(ps):
        doc_root = ps.get_xml(DOC)
        _, _first, last = _anchor_runs(doc_root, anchor)
        ins = etree.Element(_w("ins"))
        ins.set(_w("id"), str(_next_rev_id(doc_root)))
        ins.set(_w("author"), author)
        ins.set(_w("date"), date or iso_z())
        r = etree.SubElement(ins, _w("r"))
        t = etree.SubElement(r, _w("t"))
        t.text = text
        _set_preserve(t)
        last.addnext(ins)
        ps.set_xml(DOC, doc_root)
        if ps.has(SETTINGS):
            sroot = ps.get_xml(SETTINGS)
            if _set_tracking_root(sroot, True):
                ps.set_xml(SETTINGS, sroot)
    patch_parts(path, mut)


def delete_tracked(path, anchor, *, author, date=None) -> None:
    """Mark the anchored phrase (or the anchored paragraph's text) as a tracked deletion."""
    def mut(ps):
        doc_root = ps.get_xml(DOC)
        p, first, last = _anchor_runs(doc_root, anchor)
        runs = _para_runs(p)
        seq = [r for r, _ in runs]
        i0, i1 = seq.index(first), seq.index(last)
        span = seq[i0:i1 + 1]
        dl = etree.Element(_w("del"))
        dl.set(_w("id"), str(_next_rev_id(doc_root)))
        dl.set(_w("author"), author)
        dl.set(_w("date"), date or iso_z())
        first.addprevious(dl)
        for r in span:
            _remove(r)
            dl.append(r)
            for t in r.findall(_w("t")):
                t.tag = _w("delText")
                _set_preserve(t)
        ps.set_xml(DOC, doc_root)
        if ps.has(SETTINGS):
            sroot = ps.get_xml(SETTINGS)
            if _set_tracking_root(sroot, True):
                ps.set_xml(SETTINGS, sroot)
    patch_parts(path, mut)
