"""Modern threaded comments for Word (.docx) — list / add / reply / resolve / delete.

A Word threaded comment lives across five parts plus anchor markers in document.xml,
exactly as confirmed from a real Microsoft 365 file (build 16.0.20026):

    word/comments.xml            w:comment (id/author/date/initials), body w:p w14:paraId
    word/commentsExtended.xml    w15:commentEx  paraId / paraIdParent (thread link) / done (resolved)
    word/commentsIds.xml         w16cid:commentId  paraId <-> durableId
    word/commentsExtensible.xml  w16cex:commentExtensible  durableId / dateUtc
    word/people.xml              w15:person  author + presenceInfo
    word/document.xml            commentRangeStart/End + a CommentReference run per comment

Per comment we mint: a w:id (int), a w14:paraId (8-hex) and a w16cid durableId (8-hex).
A reply is just another w:comment whose commentEx carries paraIdParent = the root's paraId,
plus its own parallel range markers around the same anchored span. Resolve flips the root
thread's w15:done to "1".
"""
from __future__ import annotations

import zipfile

from lxml import etree

from _ooxml_zip import patch_parts
from _errors import AnchorNotFound, CommentNotFound, CommentError
from _util import hex8, initials as _initials_of, iso_z, local_z
from _docx_anchor import W, _w, _set_preserve, _find_phrase, _isolate, _para_runs, _paragraphs

# --- namespaces (W + XML + the anchor primitives are imported from _docx_anchor, shared with revisions) ---
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W15 = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
W16CEX = "http://schemas.microsoft.com/office/word/2018/wordml/cex"

DOC = "word/document.xml"


def _w14(t): return f"{{{W14}}}{t}"
def _w15(t): return f"{{{W15}}}{t}"
def _wcid(t): return f"{{{W16CID}}}{t}"
def _wcex(t): return f"{{{W16CEX}}}{t}"

# part name, root tag, nsmap, content-type, relationship-type, rel target (verified from a real file)
_PARTS = {
    "comments": ("word/comments.xml", _w("comments"), {"w": W, "w14": W14},
                 "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
                 "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments", "comments.xml"),
    "ex": ("word/commentsExtended.xml", _w15("commentsEx"), {"w15": W15},
           "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
           "http://schemas.microsoft.com/office/2011/relationships/commentsExtended", "commentsExtended.xml"),
    "ids": ("word/commentsIds.xml", _wcid("commentsIds"), {"w16cid": W16CID},
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
            "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds", "commentsIds.xml"),
    "cex": ("word/commentsExtensible.xml", _wcex("commentsExtensible"), {"w16cex": W16CEX},
            "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml",
            "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible", "commentsExtensible.xml"),
    "people": ("word/people.xml", _w15("people"), {"w15": W15},
               "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml",
               "http://schemas.microsoft.com/office/2011/relationships/people", "people.xml"),
}


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _read_root(z, name):
    try:
        return etree.fromstring(z.read(name))
    except KeyError:
        return None


def _comment_text(cm):
    """Plain text of a comment body, excluding the leading annotation reference mark."""
    return "".join(t.text or "" for t in cm.iter(_w("t")))


def _ex_map(ex_root):
    """paraId -> {'parent': paraIdParent|None, 'done': bool} from commentsExtended.xml."""
    out = {}
    if ex_root is None:
        return out
    for ce in ex_root.findall(_w15("commentEx")):
        out[ce.get(_w15("paraId"))] = {
            "parent": ce.get(_w15("paraIdParent")),
            "done": ce.get(_w15("done")) in ("1", "true"),
        }
    return out


def _thread_root_paraid(ex, paraid):
    cur, guard = paraid, 0
    while ex.get(cur, {}).get("parent") and guard < 1000:
        cur = ex[cur]["parent"]
        guard += 1
    return cur


def _anchor_for(doc_root, cid):
    """Return (anchor_text, paragraph_index_1based, context_text) for a comment id, or
    ('', None, '') if the comment isn't anchored in the body."""
    start = end = None
    for el in doc_root.iter(_w("commentRangeStart")):
        if el.get(_w("id")) == str(cid):
            start = el
            break
    for el in doc_root.iter(_w("commentRangeEnd")):
        if el.get(_w("id")) == str(cid):
            end = el
            break
    if start is None or end is None:
        return "", None, ""
    seq = list(doc_root.iter())
    i0, i1 = seq.index(start), seq.index(end)
    anchor_text = "".join(e.text or "" for e in seq[i0:i1] if e.tag == _w("t"))
    p = start.getparent()
    while p is not None and p.tag != _w("p"):
        p = p.getparent()
    context = "".join(t.text or "" for t in p.iter(_w("t"))) if p is not None else ""
    idx = None
    if p is not None:
        paras = _paragraphs(doc_root)   # same enumeration add --paragraph N uses, so they round-trip
        idx = paras.index(p) + 1 if p in paras else None
    return anchor_text, idx, context


def list_comments(path) -> list[dict]:
    """Every comment (roots + replies) as flat records, each carrying its thread + context."""
    with zipfile.ZipFile(path) as z:
        comments = _read_root(z, "word/comments.xml")
        if comments is None:
            return []
        ex = _ex_map(_read_root(z, "word/commentsExtended.xml"))
        doc_root = _read_root(z, DOC)

    # paraId <-> w:id, so we can translate the paraIdParent thread links into comment ids
    paraid_of, id_of_paraid = {}, {}
    for cm in comments.findall(_w("comment")):
        cid = cm.get(_w("id"))
        p = cm.find(_w("p"))
        pid = p.get(_w14("paraId")) if p is not None else None
        paraid_of[cid] = pid
        if pid:
            id_of_paraid[pid] = cid

    records = []
    for cm in comments.findall(_w("comment")):
        cid = cm.get(_w("id"))
        pid = paraid_of.get(cid)
        parent_paraid = ex.get(pid, {}).get("parent") if pid else None
        parent_id = id_of_paraid.get(parent_paraid) if parent_paraid else None
        root_paraid = _thread_root_paraid(ex, pid) if pid else None
        thread_id = id_of_paraid.get(root_paraid, cid)
        resolved = ex.get(root_paraid, {}).get("done", False) if root_paraid else False
        anchor_text, para_idx, context = (_anchor_for(doc_root, cid) if doc_root is not None else ("", None, ""))
        records.append({
            "id": cid,
            "thread_id": thread_id,
            "parent_id": parent_id,
            "is_reply": parent_id is not None,
            "author": cm.get(_w("author")),
            "initials": cm.get(_w("initials")),
            "date": cm.get(_w("date")),
            "text": _comment_text(cm),
            "resolved": resolved,
            "anchor": {"kind": "docx", "paragraph": para_idx, "text": anchor_text},
            "anchor_text": anchor_text,
            "context": context,
            "location": (f'para {para_idx}: "{anchor_text[:40]}"' if para_idx else "(unanchored)"),
        })
    return records


# ---------------------------------------------------------------------------
# Building blocks for writes
# ---------------------------------------------------------------------------

def _ensure_part(ps, key):
    """Return the lxml root for a comment part, creating + wiring it (content type +
    relationship on document.xml) the first time."""
    name, tag, nsmap, ct, rel, target = _PARTS[key]
    if ps.has(name):
        return ps.get_xml(name)
    root = etree.Element(tag, nsmap=nsmap)
    ps.add_xml_part(name, root)
    ps.ensure_content_type("/" + name, ct)
    ps.rels_for(DOC).add_rel(rel, target)
    return root


def _build_comment_el(cid, paraid, textid, author, initials, date_local, text):
    cm = etree.Element(_w("comment"))
    cm.set(_w("id"), str(cid))
    cm.set(_w("author"), author)
    cm.set(_w("date"), date_local)
    cm.set(_w("initials"), initials)
    p = etree.SubElement(cm, _w("p"))
    p.set(_w14("paraId"), paraid)
    p.set(_w14("textId"), textid)
    ppr = etree.SubElement(p, _w("pPr"))
    etree.SubElement(ppr, _w("pStyle")).set(_w("val"), "CommentText")
    r1 = etree.SubElement(p, _w("r"))
    etree.SubElement(etree.SubElement(r1, _w("rPr")), _w("rStyle")).set(_w("val"), "CommentReference")
    etree.SubElement(r1, _w("annotationRef"))
    r2 = etree.SubElement(p, _w("r"))
    t = etree.SubElement(r2, _w("t"))
    t.text = text
    _set_preserve(t)
    return cm


def _ref_run(cid):
    r = etree.Element(_w("r"))
    rpr = etree.SubElement(r, _w("rPr"))
    etree.SubElement(rpr, _w("rStyle")).set(_w("val"), "CommentReference")
    etree.SubElement(rpr, _w("sz")).set(_w("val"), "22")
    etree.SubElement(rpr, _w("szCs")).set(_w("val"), "22")
    etree.SubElement(r, _w("commentReference")).set(_w("id"), str(cid))
    return r


def _range_el(tag, cid):
    el = etree.Element(_w(tag))
    el.set(_w("id"), str(cid))
    return el


def _existing(comments_root, ids_root, cex_root):
    wids = [int(c.get(_w("id"))) for c in comments_root.findall(_w("comment"))
            if (c.get(_w("id")) or "").isdigit()]
    paraids = {p.get(_w14("paraId")) for p in comments_root.iter(_w("p")) if p.get(_w14("paraId"))}
    durables = set()
    if ids_root is not None:
        for c in ids_root.findall(_wcid("commentId")):
            paraids.add(c.get(_wcid("paraId")))
            durables.add(c.get(_wcid("durableId")))
    if cex_root is not None:
        for c in cex_root.findall(_wcex("commentExtensible")):
            durables.add(c.get(_wcex("durableId")))
    return wids, paraids, durables


def _ensure_person(people_root, author):
    for person in people_root.findall(_w15("person")):
        if person.get(_w15("author")) == author:
            return
    person = etree.SubElement(people_root, _w15("person"))
    person.set(_w15("author"), author)
    pres = etree.SubElement(person, _w15("presenceInfo"))
    pres.set(_w15("providerId"), "None")   # unmanaged identity (no signed-in account)
    pres.set(_w15("userId"), author)


# ---------------------------------------------------------------------------
# Anchoring in document.xml
# ---------------------------------------------------------------------------

# _para_runs / _split_run / _isolate / _paragraphs / _find_phrase now live in _docx_anchor
# (shared with _docx_revisions) and are imported at the top of this module.


def _anchor_top_level(doc_root, anchor, cid):
    """Insert commentRangeStart/End + a reference run for a brand-new top-level comment."""
    if "paragraph" in anchor and anchor["paragraph"] is not None:
        paras = _paragraphs(doc_root)
        n = anchor["paragraph"]
        if not (1 <= n <= len(paras)):
            raise AnchorNotFound(f"paragraph {n} out of range (1..{len(paras)})")
        runs = _para_runs(paras[n - 1])
        if not runs:
            raise AnchorNotFound(f"paragraph {n} has no text to anchor to")
        first, last = runs[0][0], runs[-1][0]
    else:
        p, s, e = _find_phrase(doc_root, anchor["text"], anchor.get("occurrence"))
        first, last = _isolate(p, s, e)
    first.addprevious(_range_el("commentRangeStart", cid))
    last.addnext(_ref_run(cid))
    last.addnext(_range_el("commentRangeEnd", cid))


def _anchor_reply(doc_root, parent_id, cid):
    """Add parallel range markers + reference run for a reply, mirroring how Word nests a
    reply's anchor around the parent's anchored span."""
    parent_start = next((e for e in doc_root.iter(_w("commentRangeStart"))
                         if e.get(_w("id")) == str(parent_id)), None)
    parent_ref = next((e for e in doc_root.iter(_w("commentReference"))
                       if e.get(_w("id")) == str(parent_id)), None)
    if parent_start is None or parent_ref is None:
        raise CommentNotFound(f"parent comment {parent_id} is not anchored in the document")
    parent_start.addnext(_range_el("commentRangeStart", cid))
    parent_ref_run = parent_ref.getparent()
    parent_ref_run.addnext(_ref_run(cid))
    parent_ref_run.addnext(_range_el("commentRangeEnd", cid))


def _remove_anchor(doc_root, cid):
    for tag in ("commentRangeStart", "commentRangeEnd"):
        for el in list(doc_root.iter(_w(tag))):
            if el.get(_w("id")) == str(cid):
                el.getparent().remove(el)
    for ref in list(doc_root.iter(_w("commentReference"))):
        if ref.get(_w("id")) == str(cid):
            run = ref.getparent()
            run.getparent().remove(run)


# ---------------------------------------------------------------------------
# Public mutations
# ---------------------------------------------------------------------------

def _ensure_modern_comments(ps):
    """Make sure the document is in modern compatibility mode (compatibilityMode >= 15).

    Word DISABLES the "Resolve thread" option for comments in Compatibility Mode, and tools
    like python-docx stamp `compatibilityMode=14` (Word 2010) by default. Upgrading to 15 lets
    the user resolve/reopen threads. This is a no-op for files that are already modern, so it
    only ever rescues old / python-docx-generated documents."""
    if not ps.has("word/settings.xml"):
        return
    root = ps.get_xml("word/settings.xml")
    compat = root.find(_w("compat"))
    cs_el, cur = None, None
    if compat is not None:
        for cs in compat.findall(_w("compatSetting")):
            if cs.get(_w("name")) == "compatibilityMode":
                cs_el = cs
                try:
                    cur = int(cs.get(_w("val")))
                except (TypeError, ValueError):
                    cur = None
    if cur is not None and cur >= 15:
        return
    if compat is None:
        compat = etree.Element(_w("compat"))
        # CT_Settings is an ordered sequence; <w:compat> must precede rsids/mathPr/themeFontLang/docId.
        follower = next((root.find(_w(tag)) for tag in
                         ("rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
                          "shapeDefaults", "decimalSymbol", "listSeparator", "docId")
                         if root.find(_w(tag)) is not None), None)
        if follower is not None:
            follower.addprevious(compat)
        else:
            root.append(compat)
    if cs_el is None:
        cs_el = etree.SubElement(compat, _w("compatSetting"))
        cs_el.set(_w("name"), "compatibilityMode")
        cs_el.set(_w("uri"), "http://schemas.microsoft.com/office/word")
    cs_el.set(_w("val"), "15")
    ps.set_xml("word/settings.xml", root)


def _insert(ps, *, text, author, initials, date, parent_id, anchor):
    _ensure_modern_comments(ps)
    comments = _ensure_part(ps, "comments")
    ex = _ensure_part(ps, "ex")
    ids = _ensure_part(ps, "ids")
    cex = _ensure_part(ps, "cex")
    people = _ensure_part(ps, "people")
    doc_root = ps.get_xml(DOC)

    wids, paraids, durables = _existing(comments, ids, cex)
    cid = (max(wids) + 1) if wids else 0
    paraid = hex8(paraids)
    durable = hex8(durables)
    textid = hex8(paraids | {paraid})
    initials = initials or _initials_of(author)
    date_local = date or local_z()
    date_utc = date or iso_z()

    # parent thread link (replies)
    parent_paraid, parent_found = None, False
    if parent_id is not None:
        for cm in comments.findall(_w("comment")):
            if cm.get(_w("id")) == str(parent_id):
                parent_found = True
                p = cm.find(_w("p"))
                parent_paraid = p.get(_w14("paraId")) if p is not None else None
                break
        if not parent_found:
            raise CommentNotFound(f"no comment with id {parent_id}")
        if parent_paraid is None:
            raise CommentError(f"comment {parent_id} is a classic (non-threaded) comment; "
                               f"replies require a modern threaded comment")

    comments.append(_build_comment_el(cid, paraid, textid, author, initials, date_local, text))

    ce = etree.SubElement(ex, _w15("commentEx"))
    ce.set(_w15("paraId"), paraid)
    if parent_paraid:
        ce.set(_w15("paraIdParent"), parent_paraid)
    ce.set(_w15("done"), "0")

    cidel = etree.SubElement(ids, _wcid("commentId"))
    cidel.set(_wcid("paraId"), paraid)
    cidel.set(_wcid("durableId"), durable)

    cexel = etree.SubElement(cex, _wcex("commentExtensible"))
    cexel.set(_wcex("durableId"), durable)
    cexel.set(_wcex("dateUtc"), date_utc)

    _ensure_person(people, author)

    if parent_id is None:
        _anchor_top_level(doc_root, anchor, cid)
    else:
        _anchor_reply(doc_root, parent_id, cid)

    # mark every touched part dirty (new parts are already dirty; existing ones need this)
    ps.set_xml("word/comments.xml", comments)
    ps.set_xml("word/commentsExtended.xml", ex)
    ps.set_xml("word/commentsIds.xml", ids)
    ps.set_xml("word/commentsExtensible.xml", cex)
    ps.set_xml("word/people.xml", people)
    ps.set_xml(DOC, doc_root)
    return str(cid)


def add_comment(path, anchor, text, *, author, initials=None, date=None) -> str:
    """Add a new top-level comment. `anchor` = {"text": "...", "occurrence": n} or
    {"paragraph": n}. Returns the new comment id."""
    out = {}
    patch_parts(path, lambda ps: out.__setitem__(
        "id", _insert(ps, text=text, author=author, initials=initials, date=date,
                      parent_id=None, anchor=anchor)))
    return out["id"]


def reply(path, parent_id, text, *, author, initials=None, date=None) -> str:
    """Add a reply into the thread of `parent_id`. Returns the new reply id."""
    out = {}
    patch_parts(path, lambda ps: out.__setitem__(
        "id", _insert(ps, text=text, author=author, initials=initials, date=date,
                      parent_id=parent_id, anchor=None)))
    return out["id"]


def _set_status_in(ps, comment_id, resolved):
    """PartSet-level core of set_status (batch edits call this directly)."""
    if not ps.has("word/comments.xml"):
        raise CommentNotFound(f"no comment with id {comment_id}")
    comments = ps.get_xml("word/comments.xml")
    if not any(c.get(_w("id")) == str(comment_id) for c in comments.findall(_w("comment"))):
        raise CommentNotFound(f"no comment with id {comment_id}")
    if not ps.has("word/commentsExtended.xml"):
        raise CommentError(f"comment {comment_id} is a classic (non-threaded) comment; "
                           f"resolve/reopen needs a modern threaded comment")
    ex = ps.get_xml("word/commentsExtended.xml")
    # comment_id -> its paraId -> thread root paraId
    paraid = None
    for cm in comments.findall(_w("comment")):
        if cm.get(_w("id")) == str(comment_id):
            p = cm.find(_w("p"))
            paraid = p.get(_w14("paraId")) if p is not None else None
            break
    if paraid is None:
        raise CommentError(f"comment {comment_id} is a classic (non-threaded) comment; "
                           f"resolve/reopen needs a modern threaded comment")
    ex_map = _ex_map(ex)
    root_paraid = _thread_root_paraid(ex_map, paraid)
    # set done on the whole thread (root + replies), exactly as Word's UI Resolve does
    members = {pid for pid in ex_map if _thread_root_paraid(ex_map, pid) == root_paraid}
    members.add(root_paraid)
    hit = False
    for ce in ex.findall(_w15("commentEx")):
        if ce.get(_w15("paraId")) in members:
            ce.set(_w15("done"), "1" if resolved else "0")
            hit = True
    if not hit:
        raise CommentNotFound(f"no commentEx for id {comment_id}")
    ps.set_xml("word/commentsExtended.xml", ex)


def set_status(path, comment_id, resolved: bool) -> None:
    """Resolve (True) or reopen (False) the whole thread that `comment_id` belongs to."""
    patch_parts(path, lambda ps: _set_status_in(ps, comment_id, resolved))


def delete(path, comment_id) -> None:
    """Delete a comment. A reply is removed on its own; a thread root takes its whole
    thread (all replies) with it."""
    def mut(ps):
        comments = ps.get_xml("word/comments.xml")
        ex = ps.get_xml("word/commentsExtended.xml") if ps.has("word/commentsExtended.xml") else None
        ids = ps.get_xml("word/commentsIds.xml") if ps.has("word/commentsIds.xml") else None
        cex = ps.get_xml("word/commentsExtensible.xml") if ps.has("word/commentsExtensible.xml") else None
        doc_root = ps.get_xml(DOC)

        # map ids <-> paraIds
        paraid_of, id_of_paraid = {}, {}
        for cm in comments.findall(_w("comment")):
            p = cm.find(_w("p"))
            pid = p.get(_w14("paraId")) if p is not None else None
            paraid_of[cm.get(_w("id"))] = pid
            if pid:
                id_of_paraid[pid] = cm.get(_w("id"))
        if str(comment_id) not in paraid_of:
            raise CommentNotFound(f"no comment with id {comment_id}")

        exm = _ex_map(ex)
        target_para = paraid_of[str(comment_id)]
        # Collect the target plus any descendants (replies) by walking parent links.
        # A CLASSIC comment has no w14:paraId (target_para is None): it cannot have
        # threaded replies, so the walk starts empty — and the walk itself only ever
        # follows REAL parent links, because every modern thread root's parent is None
        # and a None seed would otherwise absorb every thread in the file (#7).
        doomed_paraids = {target_para} if target_para else set()
        changed = True
        while changed:
            changed = False
            for pid, info in exm.items():
                if info.get("parent") and info["parent"] in doomed_paraids and pid not in doomed_paraids:
                    doomed_paraids.add(pid)
                    changed = True
        # The selected comment is always doomed by its own id — even a classic one that
        # never appears in the paraId maps.
        doomed_ids = {str(comment_id)} | {id_of_paraid[pp] for pp in doomed_paraids
                                          if pp in id_of_paraid}

        # durableIds of doomed comments — read from commentsIds BEFORE we delete its rows
        doomed_durables = set()
        if ids is not None:
            doomed_durables = {c.get(_wcid("durableId")) for c in ids.findall(_wcid("commentId"))
                               if c.get(_wcid("paraId")) in doomed_paraids}

        for cm in list(comments.findall(_w("comment"))):
            if cm.get(_w("id")) in doomed_ids:
                comments.remove(cm)
        if ex is not None:
            for ce in list(ex.findall(_w15("commentEx"))):
                if ce.get(_w15("paraId")) in doomed_paraids:
                    ex.remove(ce)
        if ids is not None:
            for c in list(ids.findall(_wcid("commentId"))):
                if c.get(_wcid("paraId")) in doomed_paraids:
                    ids.remove(c)
        if cex is not None:
            for c in list(cex.findall(_wcex("commentExtensible"))):
                if c.get(_wcex("durableId")) in doomed_durables:
                    cex.remove(c)
        for did in doomed_ids:
            _remove_anchor(doc_root, did)

        ps.set_xml("word/comments.xml", comments)
        if ex is not None:
            ps.set_xml("word/commentsExtended.xml", ex)
        if ids is not None:
            ps.set_xml("word/commentsIds.xml", ids)
        if cex is not None:
            ps.set_xml("word/commentsExtensible.xml", cex)
        ps.set_xml(DOC, doc_root)
    patch_parts(path, mut)
