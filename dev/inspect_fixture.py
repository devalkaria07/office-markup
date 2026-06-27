"""Read an Office file and print its comment-related innards — the ground-truth the
per-format modules must reproduce. Dev-only (not shipped).

    python dev/inspect_fixture.py <file.docx|.xlsx|.pptx>

Prints: the comment-related parts, the [Content_Types].xml entries and relationships
that wire them, each comment part's XML, and (for .docx) the anchor region inside
document.xml.
"""
import sys
import zipfile
from lxml import etree

KEYWORDS = ("comment", "people", "person", "author", "thread")


def _hits(name: str) -> bool:
    n = name.lower()
    return any(k in n for k in KEYWORDS)


def main(path: str) -> int:
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        parts = [n for n in names if _hits(n) and n.endswith(".xml")]

        print(f"== {path} ==")
        print("\n-- comment-related parts --")
        for n in parts:
            print("  ", n)

        print("\n-- [Content_Types].xml entries --")
        ct = z.read("[Content_Types].xml").decode("utf-8", "ignore")
        for line in ct.replace("><", ">\n<").splitlines():
            if _hits(line):
                print("  ", line.strip())

        print("\n-- relationships (all *.rels) --")
        for n in names:
            if n.endswith(".rels"):
                rels = z.read(n).decode("utf-8", "ignore")
                for line in rels.replace("><", ">\n<").splitlines():
                    if _hits(line):
                        print(f"   [{n}] {line.strip()}")

        for n in parts:
            print(f"\n-- {n} --")
            try:
                print(etree.tostring(etree.fromstring(z.read(n)),
                                     pretty_print=True, encoding="unicode"))
            except Exception as e:
                print("   parse error:", e)

        if path.lower().endswith(".docx"):
            _docx_anchor_region(z)
        elif path.lower().endswith(".pptx"):
            _list_slide_extlst(z)
    return 0


def _localname(el) -> str:
    return etree.QName(el).localname


def _docx_anchor_region(z):
    """Print each document.xml paragraph that carries a comment marker (the anchor)."""
    print("\n-- document.xml paragraphs containing comment markers --")
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.fromstring(z.read("word/document.xml"))
    markers = {"commentRangeStart", "commentRangeEnd", "commentReference"}
    seen = set()
    for p in root.iter(f"{{{W}}}p"):
        if any(_localname(c) in markers for c in p.iter()) and id(p) not in seen:
            seen.add(id(p))
            print(etree.tostring(p, pretty_print=True, encoding="unicode"))


def _list_slide_extlst(z):
    """Print each slide's <p:extLst> (where modern PowerPoint comments are discovered)."""
    print("\n-- slide extLst (comment discovery hooks) --")
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for n in sorted(x for x in z.namelist() if x.startswith("ppt/slides/slide") and x.endswith(".xml")):
        root = etree.fromstring(z.read(n))
        ext = root.find(f"{{{P}}}extLst")
        if ext is not None:
            print(f"   [{n}]")
            print(etree.tostring(ext, pretty_print=True, encoding="unicode"))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
