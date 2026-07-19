"""Directly edit text in a Word (.docx) file — replace / insert / delete, NO tracking.

The change is applied plainly (no revision markup, no author attribution) and nothing
else in the file is touched. Works in any story (body, headers, footers, footnotes).
Use replace_text / insert_text / delete_text instead when the change should be
reviewable as a tracked change.

    python scripts/edit_text.py --file r.docx --anchor-text "teh" --replace-with "the"
    python scripts/edit_text.py --file r.docx --anchor-text "Total:" --insert-after " (net)"
    python scripts/edit_text.py --file r.docx --anchor-text "DRAFT - " --delete
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                       # noqa: E402
import _docx_edit                 # noqa: E402
from _errors import AnchorNotFound  # noqa: E402


def _anchor(args):
    if args.paragraph is not None:
        a = {"paragraph": args.paragraph}
    elif args.anchor_text:
        a = {"text": args.anchor_text, "occurrence": args.occurrence}
    else:
        raise AnchorNotFound('need --anchor-text "phrase" (optional --occurrence N) or --paragraph N')
    if args.part:
        a["part"] = args.part
    return a


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Direct (untracked) text edit in a Word file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--anchor-text", dest="anchor_text", help="the phrase to act on")
    ap.add_argument("--occurrence", type=int, help="which match of --anchor-text (1-based)")
    ap.add_argument("--paragraph", type=int, help="act on the Nth paragraph's text")
    ap.add_argument("--part", help="pin the search to one story part (e.g. header1, footnotes)")
    op = ap.add_mutually_exclusive_group(required=True)
    op.add_argument("--replace-with", dest="replace_with", help="replace the phrase with this text")
    op.add_argument("--insert-after", dest="insert_after", help="insert this text after the phrase")
    op.add_argument("--insert-before", dest="insert_before", help="insert this text before the phrase")
    op.add_argument("--delete", action="store_true", help="delete the phrase")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_path(args.file)
        anchor = _anchor(args)
        if args.replace_with is not None:
            _docx_edit.replace_direct(args.file, anchor, args.replace_with)
            print("replaced (direct)")
        elif args.insert_after is not None:
            _docx_edit.insert_direct(args.file, anchor, args.insert_after)
            print("inserted (direct)")
        elif args.insert_before is not None:
            _docx_edit.insert_direct(args.file, anchor, args.insert_before, before=True)
            print("inserted (direct)")
        else:
            _docx_edit.delete_direct(args.file, anchor)
            print("deleted (direct)")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
