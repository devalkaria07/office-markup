"""Modern threaded comments for Excel (.xlsx) — list / add / reply / resolve / delete.

Excel keeps TWO synchronized representations and needs BOTH (confirmed against real Excel —
a threaded-only file has its comments dropped on open):

    xl/threadedComments/threadedComment{n}.xml   the modern thread (our source of truth)
    xl/persons/person.xml                        author registry
    xl/comments{n}.xml + xl/drawings/vmlDrawing{n}.vml + <legacyDrawing> in the sheet
                                                 the legacy "sticky note" that REGISTERS the
                                                 comment's existence on the cell

Threading is by `parentId` (reply -> root id); resolve is `done="1"` on the root threaded
comment. We treat the threaded part as authoritative and REGENERATE the legacy shadow from it
after every change, so the two never drift.
"""
from __future__ import annotations

import re
import zipfile

from lxml import etree

from _ooxml_zip import patch_parts, _rels_name
from _errors import AnchorNotFound, CommentNotFound, CommentError
from _util import guid, utc_now

# --- namespaces ---
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
X = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML = "http://www.w3.org/XML/1998/namespace"

CT_THREADED = "application/vnd.ms-excel.threadedcomments+xml"
CT_PERSON = "application/vnd.ms-excel.person+xml"
CT_COMMENTS = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
CT_VML = "application/vnd.openxmlformats-officedocument.vmlDrawing"
REL_THREADED = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
REL_PERSON = "http://schemas.microsoft.com/office/2017/10/relationships/person"
REL_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
REL_VML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

PERSONS_PART = "xl/persons/person.xml"

_LEGACY_NOTE = ("[Threaded comment]\n\nYour version of Excel allows you to read this threaded "
                "comment; however, any edits to it will get removed if the file is opened in a "
                "newer version of Excel. Learn more: https://go.microsoft.com/fwlink/?linkid=870924")

_VML_OPEN = ('<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
             ' xmlns:x="urn:schemas-microsoft-com:office:excel">')
_VML_SHAPETYPE = ('<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
                  'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
                  '<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>')
_VML_TAIL = "</xml>"


def _vml_document(shapes):
    """Assemble the VML note drawing. <o:idmap> must declare every 1024-wide shape-id block the
    notes actually span (ids start at _x0000_s1025), so a hardcoded data="1" only covers the first
    1023 notes; compute the block list to stay valid for very large comment counts."""
    blocks = sorted({(1025 + i) // 1024 for i in range(max(1, len(shapes)))})
    idmap = ('<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" '
             f'data="{",".join(map(str, blocks))}"/></o:shapelayout>')
    return _VML_OPEN + idmap + _VML_SHAPETYPE + "".join(shapes) + _VML_TAIL


# ---------------------------------------------------------------------------
# Cell-reference helpers
# ---------------------------------------------------------------------------

def _col_to_n(col):
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _parse_ref(ref):
    m = re.match(r"([A-Za-z]+)(\d+)$", ref.strip())
    if not m:
        raise AnchorNotFound(f"bad cell reference: {ref!r}")
    return _col_to_n(m.group(1)), int(m.group(2)) - 1   # (col0, row0)


def _xl_dt(date):
    return date or (utc_now().strftime("%Y-%m-%dT%H:%M:%S") + ".00")


# ---------------------------------------------------------------------------
# Relationship / sheet resolution (driven by a get_bytes(name)->bytes|None callable)
# ---------------------------------------------------------------------------

def _rels_by_type(get_bytes, part):
    raw = get_bytes(_rels_name(part))
    by_id, by_type = {}, {}
    if raw:
        for rel in etree.fromstring(raw):
            by_id[rel.get("Id")] = rel.get("Target")
            by_type.setdefault(rel.get("Type"), []).append((rel.get("Id"), rel.get("Target")))
    return by_id, by_type


def _abs_target(base_part, target):
    if target.startswith("/"):       # absolute package path (openpyxl writes these)
        return target[1:]
    base = base_part.rsplit("/", 1)[0]
    out = []
    for p in (base + "/" + target).split("/"):
        if p == "..":
            if out:
                out.pop()
        elif p not in ("", "."):
            out.append(p)
    return "/".join(out)


def _sheet_map(get_bytes):
    wb = etree.fromstring(get_bytes("xl/workbook.xml"))
    by_id, _ = _rels_by_type(get_bytes, "xl/workbook.xml")
    out = {}
    for sh in wb.iter(f"{{{X}}}sheet"):
        tgt = by_id.get(sh.get(f"{{{R}}}id"))
        if tgt:
            out[sh.get("name")] = _abs_target("xl/workbook.xml", tgt)
    return out


def _sheet_part_rel(get_bytes, sheet_part, rel_type):
    _, by_type = _rels_by_type(get_bytes, sheet_part)
    rels = by_type.get(rel_type)
    return _abs_target(sheet_part, rels[0][1]) if rels else None


def _sheet_number(sheet_part):
    m = re.search(r"sheet(\d+)\.xml$", sheet_part)
    return m.group(1) if m else "1"


# ---------------------------------------------------------------------------
# Reading
# ---------------------------------------------------------------------------

def _persons_map(get_bytes):
    raw = get_bytes(PERSONS_PART)
    out = {}
    if raw:
        for p in etree.fromstring(raw).iter(f"{{{TC}}}person"):
            out[p.get("id")] = p.get("displayName")
    return out


def list_comments(path) -> list[dict]:
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        get = lambda n: (z.read(n) if n in names else None)  # noqa: E731
        persons = _persons_map(get)
        sheet_map = _sheet_map(get)
        cell_sheets = {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            cell_sheets = {s: wb[s] for s in sheet_map if s in wb.sheetnames}
        except Exception:
            cell_sheets = {}

        records = []
        for sname, sheet_part in sheet_map.items():
            tpart = _sheet_part_rel(get, sheet_part, REL_THREADED)
            raw = get(tpart) if tpart else None
            if not raw:
                continue
            root = etree.fromstring(raw)
            tcs = list(root.iter(f"{{{TC}}}threadedComment"))
            done_of_root = {tc.get("id"): (tc.get("done") in ("1", "true"))
                            for tc in tcs if tc.get("parentId") is None}
            for tc in tcs:
                cid, parent, ref = tc.get("id"), tc.get("parentId"), tc.get("ref")
                root_id = parent or cid
                tnode = tc.find(f"{{{TC}}}text")
                val = None
                if sname in cell_sheets:
                    try:
                        val = cell_sheets[sname][ref].value
                    except Exception:
                        val = None
                records.append({
                    "id": cid,
                    "thread_id": root_id,
                    "parent_id": parent,
                    "is_reply": parent is not None,
                    "author": persons.get(tc.get("personId")),
                    "date": tc.get("dT"),
                    "text": tnode.text if tnode is not None else "",
                    "resolved": done_of_root.get(root_id, False),
                    "anchor": {"kind": "xlsx", "sheet": sname, "cell": ref},
                    "anchor_text": "" if val is None else str(val),
                    "context": "" if val is None else f"{ref} = {val!r}",
                    "location": f"{sname}!{ref}",
                })
    return records


# ---------------------------------------------------------------------------
# Writing
# ---------------------------------------------------------------------------

def _all_ids(ps):
    seen = set()
    if ps.has(PERSONS_PART):
        seen |= {p.get("id") for p in ps.get_xml(PERSONS_PART).iter(f"{{{TC}}}person")}
    for name in ps.names():
        if "threadedComment" in name and name.endswith(".xml"):
            seen |= {tc.get("id") for tc in ps.get_xml(name).iter(f"{{{TC}}}threadedComment")}
    return seen


def _ensure_person(ps, author):
    if ps.has(PERSONS_PART):
        root = ps.get_xml(PERSONS_PART)
        for p in root.findall(f"{{{TC}}}person"):
            if p.get("displayName") == author:
                return p.get("id")
    else:
        root = etree.Element(f"{{{TC}}}personList", nsmap={None: TC, "x": X})
        ps.add_xml_part(PERSONS_PART, root)
        ps.ensure_content_type("/" + PERSONS_PART, CT_PERSON)
        ps.rels_for("xl/workbook.xml").add_rel(REL_PERSON, "persons/person.xml")
    pid = guid(exclude=_all_ids(ps))
    p = etree.SubElement(root, f"{{{TC}}}person")
    p.set("displayName", author)
    p.set("id", pid)
    p.set("userId", author)
    p.set("providerId", "None")
    ps.set_xml(PERSONS_PART, root)
    return pid


def _threaded_for_sheet(ps, sheet_part):
    tpart = _sheet_part_rel(ps.get_bytes, sheet_part, REL_THREADED)
    if tpart and ps.has(tpart):
        return tpart, ps.get_xml(tpart)
    if not tpart:
        n = _sheet_number(sheet_part)
        tpart = f"xl/threadedComments/threadedComment{n}.xml"
        i = int(n)
        while ps.has(tpart):
            i += 1
            tpart = f"xl/threadedComments/threadedComment{i}.xml"
        ps.rels_for(sheet_part).add_rel(REL_THREADED, "../threadedComments/" + tpart.rsplit("/", 1)[1])
        ps.ensure_content_type("/" + tpart, CT_THREADED)
    root = etree.Element(f"{{{TC}}}ThreadedComments", nsmap={None: TC, "x": X})
    ps.add_xml_part(tpart, root)
    return tpart, root


def _thread_root_id(troot, cid):
    by_id = {tc.get("id"): tc for tc in troot.iter(f"{{{TC}}}threadedComment")}
    tc = by_id.get(cid)
    return (tc.get("parentId") or cid) if tc is not None else cid


def _find_comment(ps, cid):
    for sname, spart in _sheet_map(ps.get_bytes).items():
        tpart = _sheet_part_rel(ps.get_bytes, spart, REL_THREADED)
        if tpart and ps.has(tpart):
            for tc in ps.get_xml(tpart).iter(f"{{{TC}}}threadedComment"):
                if tc.get("id") == cid:
                    return sname, spart, tc.get("ref")
    return None


def _vml_shape(shape_id, col0, row0):
    anchor = f"{col0 + 1}, 15, {max(0, row0 - 1)}, 9, {col0 + 3}, 15, {row0 + 3}, 12"
    return (f'<v:shape id="_x0000_s{shape_id}" type="#_x0000_t202" '
            "style='position:absolute;margin-left:119.25pt;margin-top:6.75pt;width:108pt;"
            "height:59.25pt;z-index:1;visibility:hidden' fillcolor=\"infoBackground [80]\" "
            'strokecolor="none [81]" o:insetmode="auto">'
            '<v:fill color2="infoBackground [80]"/><v:shadow color="none [81]" obscured="t"/>'
            '<v:path o:connecttype="none"/>'
            "<v:textbox style='mso-direction-alt:auto'><div style='text-align:left'></div></v:textbox>"
            '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>'
            f'<x:Anchor>{anchor}</x:Anchor>'
            f'<x:AutoFill>False</x:AutoFill><x:Row>{row0}</x:Row><x:Column>{col0}</x:Column>'
            '</x:ClientData></v:shape>')


def _legacy_thread_text(troot, root_id):
    lines = [_LEGACY_NOTE, ""]
    for tc in troot.iter(f"{{{TC}}}threadedComment"):
        if tc.get("id") == root_id or tc.get("parentId") == root_id:
            tnode = tc.find(f"{{{TC}}}text")
            lines.append("Comment:" if tc.get("parentId") is None else "Reply:")
            lines.append("    " + (tnode.text if tnode is not None else ""))
    return "\n".join(lines)


def _check_legacy_clobber(ps, comments_part, vml_part):
    """Refuse to regenerate the legacy shadow when it holds non-comment content we'd destroy:
    classic (non-threaded) cell notes, or VML shapes that aren't comment notes (e.g. form controls)."""
    raw_c = ps.get_bytes(comments_part)
    if raw_c:
        for a in etree.fromstring(raw_c).iter(f"{{{X}}}author"):
            if not (a.text or "").startswith("tc="):
                raise CommentError("this sheet has classic (non-threaded) comments; office-markup "
                                   "does not modify sheets that mix classic and threaded comments")
    raw_v = ps.get_bytes(vml_part)
    if raw_v:
        text = raw_v.decode("utf-8", "ignore")
        if len(re.findall(r"<v:shape\b", text)) > text.count('ObjectType="Note"'):
            raise CommentError("this sheet has non-comment drawings (e.g. form controls) in its "
                               "legacy VML; office-markup won't risk overwriting them")


def _ensure_sheet_rel(ps, sheet_part, rel_type, target):
    """Return the Id of the sheet's relationship of `rel_type`, REUSING an existing one whatever its
    target spelling, and only adding `target` when none exists. Prevents a duplicate <Relationship>
    (which makes Excel drop the comments) when a foreign tool wrote the rel with a different spelling
    than our own `../...`."""
    _, by_type = _rels_by_type(ps.get_bytes, sheet_part)
    existing = by_type.get(rel_type)
    if existing:
        return existing[0][0]
    return ps.rels_for(sheet_part).add_rel(rel_type, target)


def _drop_sheet_rels(ps, sheet_part, rel_types):
    """Remove all of the sheet's relationships of the given types, whatever their target spelling —
    so a foreign `/xl/...`-spelled rel is unwired too, not left dangling to a dropped part."""
    _, by_type = _rels_by_type(ps.get_bytes, sheet_part)
    rels = ps.rels_for(sheet_part)
    for rt in rel_types:
        for _rid, tgt in by_type.get(rt, []):
            rels.remove_by_target(tgt)


def _remove_legacy_shadow(ps, sheet_part, comments_part, vml_part):
    """Tear the legacy shadow down when a sheet has no comments left: drop the comments + VML
    parts (and their content types), remove the <legacyDrawing>, and unwire the two rels."""
    ps.drop_part(comments_part)
    ps.drop_part(vml_part)
    sheet_root = ps.get_xml(sheet_part)
    ld = sheet_root.find(f"{{{X}}}legacyDrawing")
    if ld is not None:
        sheet_root.remove(ld)
        ps.set_xml(sheet_part, sheet_root)
    _drop_sheet_rels(ps, sheet_part, (REL_COMMENTS, REL_VML))


def _rebuild_legacy(ps, sheet_part, troot):
    """Regenerate xl/comments{n}.xml + xl/drawings/vmlDrawing{n}.vml from the current threads."""
    n = _sheet_number(sheet_part)
    comments_part = _sheet_part_rel(ps.get_bytes, sheet_part, REL_COMMENTS) or f"xl/comments{n}.xml"
    vml_part = _sheet_part_rel(ps.get_bytes, sheet_part, REL_VML) or f"xl/drawings/vmlDrawing{n}.vml"
    roots = [tc for tc in troot.iter(f"{{{TC}}}threadedComment") if tc.get("parentId") is None]

    _check_legacy_clobber(ps, comments_part, vml_part)
    if not roots:
        _remove_legacy_shadow(ps, sheet_part, comments_part, vml_part)
        return

    croot = etree.Element(f"{{{X}}}comments", nsmap={None: X, "mc": MC, "xr": XR})
    croot.set(f"{{{MC}}}Ignorable", "xr")
    authors = etree.SubElement(croot, f"{{{X}}}authors")
    clist = etree.SubElement(croot, f"{{{X}}}commentList")
    shapes = []
    for idx, rt in enumerate(roots):
        rid, ref = rt.get("id"), rt.get("ref")
        etree.SubElement(authors, f"{{{X}}}author").text = f"tc={rid}"
        c = etree.SubElement(clist, f"{{{X}}}comment")
        c.set("ref", ref)
        c.set("authorId", str(idx))
        # Excel writes shapeId="0" on EVERY note (the cell comes from ref + the VML shape's
        # Row/Column, not from shapeId). A per-comment 1..N points at a VML shape id that does not
        # exist (real ids are 1025+), so Excel drops the whole comments part on open. See tests.
        c.set("shapeId", "0")
        c.set(f"{{{XR}}}uid", rid)
        t = etree.SubElement(etree.SubElement(c, f"{{{X}}}text"), f"{{{X}}}t")
        t.text = _legacy_thread_text(troot, rid)
        t.set(f"{{{XML}}}space", "preserve")
        col, row = _parse_ref(ref)
        shapes.append(_vml_shape(1025 + idx, col, row))

    if ps.has(comments_part):
        ps.set_xml(comments_part, croot)
    else:
        ps.add_xml_part(comments_part, croot)
        ps.ensure_content_type("/" + comments_part, CT_COMMENTS)
    ps.set_bytes(vml_part, _vml_document(shapes).encode("utf-8"))

    # wiring (idempotent): vml content-type default, the two sheet rels, and <legacyDrawing>
    ps.ensure_default("vml", CT_VML)
    _ensure_sheet_rel(ps, sheet_part, REL_COMMENTS, "../" + comments_part[len("xl/"):])
    vml_rid = _ensure_sheet_rel(ps, sheet_part, REL_VML, "../" + vml_part[len("xl/"):])
    sheet_root = ps.get_xml(sheet_part)
    ld = sheet_root.find(f"{{{X}}}legacyDrawing")
    if ld is None:
        ld = etree.Element(f"{{{X}}}legacyDrawing", nsmap={"r": R})   # bind r: so r:id (not ns0:id) serialises
        nxt = next((sheet_root.find(f"{{{X}}}{tag}") for tag in ("tableParts", "extLst")
                    if sheet_root.find(f"{{{X}}}{tag}") is not None), None)
        (nxt.addprevious if nxt is not None else sheet_root.append)(ld)
    ld.set(f"{{{R}}}id", vml_rid)
    ps.set_xml(sheet_part, sheet_root)


def _insert(ps, *, sheet, cell, text, author, date, parent_id):
    sheet_map = _sheet_map(ps.get_bytes)
    if parent_id is None:
        sheet = sheet or next(iter(sheet_map))
        if sheet not in sheet_map:
            raise AnchorNotFound(f"no sheet named {sheet!r}; have {list(sheet_map)}")
        sheet_part = sheet_map[sheet]
        _parse_ref(cell)
        ref = cell.upper()
        # Excel allows ONE threaded comment thread per cell — refuse a duplicate root.
        existing = _sheet_part_rel(ps.get_bytes, sheet_part, REL_THREADED)
        if existing and ps.has(existing):
            for tc in ps.get_xml(existing).iter(f"{{{TC}}}threadedComment"):
                if tc.get("parentId") is None and tc.get("ref") == ref:
                    raise CommentError(f"cell {ref} already has a comment thread; reply to it instead")
    else:
        loc = _find_comment(ps, parent_id)
        if loc is None:
            raise CommentNotFound(f"no comment with id {parent_id}")
        sheet, sheet_part, ref = loc

    pid = _ensure_person(ps, author)
    tpart, troot = _threaded_for_sheet(ps, sheet_part)
    new_id = guid(exclude=_all_ids(ps))
    tc = etree.SubElement(troot, f"{{{TC}}}threadedComment")
    tc.set("ref", ref)
    tc.set("dT", _xl_dt(date))
    tc.set("personId", pid)
    tc.set("id", new_id)
    if parent_id is not None:
        tc.set("parentId", _thread_root_id(troot, parent_id))
    etree.SubElement(tc, f"{{{TC}}}text").text = text

    ps.set_xml(tpart, troot)
    _rebuild_legacy(ps, sheet_part, troot)
    return new_id


def add_comment(path, anchor, text, *, author, initials=None, date=None) -> str:
    """Add a top-level comment. `anchor` = {"sheet": name|None, "cell": "B2"}."""
    out = {}
    patch_parts(path, lambda ps: out.__setitem__(
        "id", _insert(ps, sheet=anchor.get("sheet"), cell=anchor["cell"], text=text,
                      author=author, date=date, parent_id=None)))
    return out["id"]


def reply(path, parent_id, text, *, author, initials=None, date=None) -> str:
    out = {}
    patch_parts(path, lambda ps: out.__setitem__(
        "id", _insert(ps, sheet=None, cell=None, text=text, author=author, date=date,
                      parent_id=parent_id)))
    return out["id"]


def set_status(path, comment_id, resolved: bool) -> None:
    def mut(ps):
        loc = _find_comment(ps, comment_id)
        if loc is None:
            raise CommentNotFound(f"no comment with id {comment_id}")
        _, sheet_part, _ = loc
        tpart = _sheet_part_rel(ps.get_bytes, sheet_part, REL_THREADED)
        troot = ps.get_xml(tpart)
        root_id = _thread_root_id(troot, comment_id)
        for tc in troot.iter(f"{{{TC}}}threadedComment"):
            if tc.get("id") == root_id:
                tc.set("done", "1" if resolved else "0")
                ps.set_xml(tpart, troot)
                return
        raise CommentNotFound(f"no comment with id {comment_id}")
    patch_parts(path, mut)


def delete(path, comment_id) -> None:
    """Delete a comment. A reply goes alone; a thread root takes its whole thread."""
    def mut(ps):
        loc = _find_comment(ps, comment_id)
        if loc is None:
            raise CommentNotFound(f"no comment with id {comment_id}")
        _, sheet_part, _ = loc
        tpart = _sheet_part_rel(ps.get_bytes, sheet_part, REL_THREADED)
        troot = ps.get_xml(tpart)
        by_id = {tc.get("id"): tc for tc in troot.iter(f"{{{TC}}}threadedComment")}
        target = by_id[comment_id]
        if target.get("parentId") is None:
            doomed = [tc for tc in by_id.values()
                      if tc.get("id") == comment_id or tc.get("parentId") == comment_id]
        else:
            doomed = [target]
        for tc in doomed:
            troot.remove(tc)
        if troot.findall(f"{{{TC}}}threadedComment"):
            ps.set_xml(tpart, troot)
            _rebuild_legacy(ps, sheet_part, troot)
        else:
            # last comment removed: tear down the threaded part + the legacy shadow + all wiring
            n = _sheet_number(sheet_part)
            comments_part = _sheet_part_rel(ps.get_bytes, sheet_part, REL_COMMENTS) or f"xl/comments{n}.xml"
            vml_part = _sheet_part_rel(ps.get_bytes, sheet_part, REL_VML) or f"xl/drawings/vmlDrawing{n}.vml"
            ps.drop_part(tpart)
            _drop_sheet_rels(ps, sheet_part, (REL_THREADED,))
            _remove_legacy_shadow(ps, sheet_part, comments_part, vml_part)
    patch_parts(path, mut)
