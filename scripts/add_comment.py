"""Add a new top-level comment to a Word/Excel/PowerPoint file.

Anchor depends on the format:
    Word (.docx):       --anchor-text "phrase" [--occurrence N]   OR   --paragraph N
    Excel (.xlsx):      --cell B2 [--sheet "Sheet1"]              (sheet defaults to the first)
    PowerPoint (.pptx): --slide N

    python scripts/add_comment.py --file r.docx --anchor-text "this figure" \
        --author "Alex Morgan" --text "Please confirm."

Prints the new comment id on success.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                       # noqa: E402
from _errors import AnchorNotFound  # noqa: E402


def _anchor(file, args):
    ext = Path(file).suffix.lower()
    if ext == ".docx":
        if args.paragraph is not None:
            return {"paragraph": args.paragraph}
        if args.anchor_text:
            return {"text": args.anchor_text, "occurrence": args.occurrence}
        raise AnchorNotFound("Word needs --anchor-text \"phrase\" (with optional --occurrence) or --paragraph N")
    if ext == ".xlsx":
        if not args.cell:
            raise AnchorNotFound("Excel needs --cell (e.g. --cell B2)")
        return {"sheet": args.sheet, "cell": args.cell}
    if ext == ".pptx":
        if args.slide is None:
            raise AnchorNotFound("PowerPoint needs --slide N")
        return {"slide": args.slide}
    return {}


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Add a threaded comment to an Office file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--text", required=True)
    ap.add_argument("--author", required=True)
    ap.add_argument("--initials")
    ap.add_argument("--anchor-text", dest="anchor_text", help="Word: text to attach the comment to")
    ap.add_argument("--occurrence", type=int, help="Word: which match of --anchor-text (1-based)")
    ap.add_argument("--paragraph", type=int, help="Word: anchor to the whole Nth paragraph")
    ap.add_argument("--sheet", help="Excel: sheet name (default: first sheet)")
    ap.add_argument("--cell", help="Excel: cell reference, e.g. B2")
    ap.add_argument("--slide", type=int, help="PowerPoint: slide number (1-based)")
    args = ap.parse_args(argv)

    def go():
        mod = _cli.module_for(args.file)
        cid = mod.add_comment(args.file, _anchor(args.file, args), args.text,
                              author=args.author, initials=args.initials)
        print(cid)
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
