"""Binary-safe OOXML zip surgery — the shared engine for the office-markup skill.

An Office file (.docx / .xlsx / .pptx) is really a ZIP archive of XML "parts". To add
or change a comment we want to touch ONLY the comment-related parts and copy every
other part through byte-for-byte, preserving the original archive's entry order and
per-entry metadata. That is precisely what python-docx / openpyxl / python-pptx will
NOT do for modern *threaded* comments, so we do it here by hand.

The core pattern is generalized from oml-docs `_xlsx_common.inject_oml_theme`:

    read every entry  ->  mutate only the targets  ->  rewrite in the original order

On top of the raw read/rewrite this module adds the small amount of OPC "plumbing"
needed when a comment lives in a brand-new part: registering the part's content type
in ``[Content_Types].xml`` and wiring a relationship in the relevant ``.rels`` file.

Typical use (a per-format module supplies the `mutator`):

    def mutate(ps: PartSet) -> None:
        root = ps.get_xml("word/comments.xml")     # edit an existing part ...
        ...
        ps.set_xml("word/comments.xml", root)
        ps.add_part("word/commentsExtended.xml", new_root)        # ... or add a new one
        ps.ensure_content_type("/word/commentsExtended.xml", CT_COMMENTS_EXTENDED)
        ps.rels_for("word/document.xml").add_rel(REL_COMMENTS_EXTENDED, "commentsExtended.xml")

    patch_parts("Report.docx", mutate)

Byte-stability guarantee: a part is re-serialized ONLY if a mutator explicitly hands it
back via ``set_xml`` / ``set_bytes`` (or adds a relationship to its ``.rels``). Every
other part is written out with its original bytes, untouched.
"""
from __future__ import annotations

import io
import zipfile
from pathlib import Path
from typing import Callable

from lxml import etree

# Single source of truth for the skill version. SKILL.md `metadata.version` MUST match
# this; scripts/release.py asserts the two are in sync (see RELEASING.md).
__version__ = "0.2.0"

# OPC packaging namespaces — the "table of contents" and the cross-reference parts that
# every .docx/.xlsx/.pptx shares, regardless of which Office app produced it.
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

CONTENT_TYPES = "[Content_Types].xml"


def serialize(root) -> bytes:
    """Serialize an lxml element the way Office writes a part: a UTF-8 XML declaration
    with standalone="yes", and no pretty-print reflow (stray whitespace can change how
    some parts are read). Only ever applied to parts we actually modified."""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _split(part_name: str):
    """Split 'word/document.xml' -> ('word/', 'document.xml'); 'foo.xml' -> ('', 'foo.xml')."""
    if "/" in part_name:
        folder, fname = part_name.rsplit("/", 1)
        return folder + "/", fname
    return "", part_name


def _rels_name(part_name: str) -> str:
    """The relationships part that belongs to `part_name`.
    e.g. 'word/document.xml' -> 'word/_rels/document.xml.rels'."""
    folder, fname = _split(part_name)
    return f"{folder}_rels/{fname}.rels"


# ---------------------------------------------------------------------------
# Low-level read / write (the inject_oml_theme pattern, generalized)
# ---------------------------------------------------------------------------

def read_parts(path):
    """Return (infos, data): the ZipInfo list in original order and {name: bytes}."""
    with zipfile.ZipFile(path) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i.filename) for i in infos}
    return infos, data


def write_parts(path, infos, data, added=None) -> None:
    """Rewrite `path` atomically (build the whole archive in memory, then one write).
    Existing parts keep their original ZipInfo — entry order, compression type and the
    rest — so anything we didn't change is byte-identical. Brand-new parts in `added`
    ({name: bytes}) are appended afterwards with default DEFLATE compression."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in infos:
            zout.writestr(info, data[info.filename])
        for name, payload in (added or {}).items():
            zout.writestr(name, payload)
    Path(path).write_bytes(buf.getvalue())


# ---------------------------------------------------------------------------
# Relationships editor (one `_rels/<file>.rels` part)
# ---------------------------------------------------------------------------

class RelsEditor:
    """Reads/creates a single relationships part and adds <Relationship> entries.
    Created lazily by PartSet.rels_for(); only written back if something was added."""

    def __init__(self, partset: "PartSet", rels_name: str):
        self._ps = partset
        self.rels_name = rels_name
        self.changed = False
        raw = partset.get_bytes(rels_name)
        if raw is None:
            # No rels part yet (e.g. a part that never referenced anything). Start one.
            self.root = etree.fromstring(
                ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 f'<Relationships xmlns="{REL_NS}"/>').encode("utf-8"))
        else:
            self.root = etree.fromstring(raw)

    def _next_id(self) -> str:
        used = {r.get("Id") for r in self.root}
        n = 1
        while f"rId{n}" in used:
            n += 1
        return f"rId{n}"

    def add_rel(self, rel_type: str, target: str, target_mode: str | None = None) -> str:
        """Add a relationship (idempotent on (Type, Target)) and return its Id."""
        for r in self.root:
            if r.get("Type") == rel_type and r.get("Target") == target:
                return r.get("Id")
        rid = self._next_id()
        el = etree.SubElement(self.root, f"{{{REL_NS}}}Relationship")
        el.set("Id", rid)
        el.set("Type", rel_type)
        el.set("Target", target)
        if target_mode:
            el.set("TargetMode", target_mode)
        self.changed = True
        return rid

    def remove_by_target(self, target: str) -> None:
        """Remove any <Relationship> whose Target equals `target`."""
        for r in list(self.root):
            if r.get("Target") == target:
                self.root.remove(r)
                self.changed = True

    def flush(self) -> None:
        if self.changed:
            self._ps.set_bytes(self.rels_name, serialize(self.root))


# ---------------------------------------------------------------------------
# PartSet — the working set handed to a mutator
# ---------------------------------------------------------------------------

class PartSet:
    """Holds every part of one Office file and offers safe edit helpers. A part is only
    re-serialized if you explicitly set_xml/set_bytes it (or add a rel to its .rels)."""

    def __init__(self, infos, data):
        self.infos = infos          # original ZipInfo list (order preserved)
        self.data = data            # {name: bytes} for existing parts
        self.added: dict[str, bytes] = {}   # {name: bytes} for brand-new parts
        self._xml: dict[str, object] = {}   # name -> parsed root (read and/or written)
        self._dirty: set[str] = set()       # names that were set_xml'd (need re-serialize)
        self._rels: dict[str, RelsEditor] = {}

    # --- raw bytes ---
    def has(self, name: str) -> bool:
        return name in self.data or name in self.added or name in self._xml

    def names(self):
        return list(self.data.keys()) + list(self.added.keys())

    def get_bytes(self, name: str):
        if name in self.added:
            return self.added[name]
        return self.data.get(name)

    def set_bytes(self, name: str, payload: bytes) -> None:
        if name in self.data:
            self.data[name] = payload
        else:
            self.added[name] = payload

    # --- xml convenience ---
    def get_xml(self, name: str):
        """Parse a part to an lxml element (cached). Reading does NOT mark it dirty."""
        if name not in self._xml:
            raw = self.get_bytes(name)
            if raw is None:
                raise KeyError(name)
            self._xml[name] = etree.fromstring(raw)
        return self._xml[name]

    def set_xml(self, name: str, root) -> None:
        """Mark an existing part as changed; it is serialized on flush."""
        self._xml[name] = root
        self._dirty.add(name)

    def add_part(self, name: str, root_or_bytes) -> None:
        """Register a brand-new part from a finished lxml element or raw bytes (snapshot
        taken now). Use add_xml_part instead if you'll keep editing the element."""
        payload = (root_or_bytes if isinstance(root_or_bytes, (bytes, bytearray))
                   else serialize(root_or_bytes))
        self.added[name] = bytes(payload)

    def add_xml_part(self, name: str, root) -> None:
        """Register a brand-new part as an EDITABLE lxml root (serialized on flush).
        Use this (not add_part) when you create a part and then keep appending to it."""
        self._xml[name] = root
        self._dirty.add(name)

    def drop_part(self, name: str) -> None:
        """Remove a part entirely, including its [Content_Types].xml Override. The caller is
        responsible for removing any relationship that targets it (see RelsEditor.remove_by_target)."""
        self.infos = [i for i in self.infos if i.filename != name]
        self.data.pop(name, None)
        self.added.pop(name, None)
        self._xml.pop(name, None)
        self._dirty.discard(name)
        if self.has(CONTENT_TYPES):
            ct = self.get_xml(CONTENT_TYPES)
            overrides = [ov for ov in ct.findall(f"{{{CT_NS}}}Override") if ov.get("PartName") == "/" + name]
            for ov in overrides:
                ct.remove(ov)
            if overrides:
                self.set_xml(CONTENT_TYPES, ct)

    # --- OPC plumbing: content types + relationships ---
    def ensure_content_type(self, part_name: str, content_type: str) -> None:
        """Add an <Override PartName=.. ContentType=..> to [Content_Types].xml if absent.
        `part_name` is the absolute package path, e.g. '/word/commentsExtended.xml'."""
        ct = self.get_xml(CONTENT_TYPES)
        for ov in ct.findall(f"{{{CT_NS}}}Override"):
            if ov.get("PartName") == part_name:
                return
        el = etree.SubElement(ct, f"{{{CT_NS}}}Override")
        el.set("PartName", part_name)
        el.set("ContentType", content_type)
        self.set_xml(CONTENT_TYPES, ct)

    def ensure_default(self, ext: str, content_type: str) -> None:
        """Add a <Default Extension=.. ContentType=..> if that extension isn't declared."""
        ct = self.get_xml(CONTENT_TYPES)
        for d in ct.findall(f"{{{CT_NS}}}Default"):
            if (d.get("Extension") or "").lower() == ext.lower():
                return
        el = etree.Element(f"{{{CT_NS}}}Default")
        el.set("Extension", ext)
        el.set("ContentType", content_type)
        ct.insert(0, el)          # Defaults conventionally precede Overrides
        self.set_xml(CONTENT_TYPES, ct)

    def rels_for(self, part_name: str) -> RelsEditor:
        """RelsEditor for the part's relationships file (created in memory if missing)."""
        if part_name not in self._rels:
            self._rels[part_name] = RelsEditor(self, _rels_name(part_name))
        return self._rels[part_name]

    # --- flush (called by patch_parts) ---
    def _flush(self) -> None:
        for name in self._dirty:
            self.set_bytes(name, serialize(self._xml[name]))
        for rels in self._rels.values():
            rels.flush()


def patch_parts(path, mutator: Callable[[PartSet], None]) -> None:
    """Open `path`, hand a PartSet to `mutator`, then rewrite — preserving every part
    the mutator didn't touch byte-for-byte."""
    infos, data = read_parts(path)
    ps = PartSet(infos, data)
    mutator(ps)
    ps._flush()
    write_parts(path, ps.infos, ps.data, ps.added)
