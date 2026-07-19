"""Insert text into a Word (.docx) file as a tracked insertion.

The new text goes in right after the anchor phrase (or at the end of the anchored paragraph),
wrapped as a tracked change by --author, and Track Changes recording is turned on.

    python scripts/insert_text.py --file r.docx --anchor-text "revenue" \
        --text " (Q3, provisional)" --author "Alex Morgan"
    python scripts/insert_text.py --file r.docx --paragraph 2 --text " See appendix." \
        --author "Alex Morgan"
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                       # noqa: E402
from _errors import AnchorNotFound  # noqa: E402


def _anchor(args):
    if args.paragraph is not None:
        a = {"paragraph": args.paragraph}
    elif args.anchor_text:
        a = {"text": args.anchor_text, "occurrence": args.occurrence}
    else:
        raise AnchorNotFound('need --anchor-text "phrase" (optional --occurrence N) or --paragraph N')
    if getattr(args, "part", None):
        a["part"] = args.part
    return a


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Insert text as a tracked change in a Word file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--text", required=True)
    ap.add_argument("--author", required=True)
    ap.add_argument("--anchor-text", dest="anchor_text", help="insert right after this phrase")
    ap.add_argument("--occurrence", type=int, help="which match of --anchor-text (1-based)")
    ap.add_argument("--paragraph", type=int, help="insert at the end of the Nth paragraph")
    ap.add_argument("--part", help="pin the search to one story part (e.g. header1, footnotes)")
    ap.add_argument("--date", help="ISO-8601 timestamp (default: now, UTC)")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_only(args.file).insert_tracked(args.file, _anchor(args), args.text,
                                                 author=args.author, date=args.date)
        print("inserted (tracked)")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
