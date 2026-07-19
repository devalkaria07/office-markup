"""Replace text in a Word (.docx) file as ONE tracked change.

The anchored phrase becomes a tracked deletion and the replacement a tracked insertion
right after it — exactly what a human editing with Track Changes on produces. Works in
any story (body, headers, footers, footnotes); Track Changes recording is turned on.

    python scripts/replace_text.py --file r.docx --anchor-text "Q3 figures" \
        --text "Q3 (final) figures" --author "Alex Morgan"
    python scripts/replace_text.py --file r.docx --anchor-text "DRAFT" --part header1 \
        --text "FINAL" --author "Alex Morgan"
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
    if args.part:
        a["part"] = args.part
    return a


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Replace text as one tracked change in a Word file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--text", required=True, help="the replacement text")
    ap.add_argument("--author", required=True)
    ap.add_argument("--anchor-text", dest="anchor_text", help="the phrase to replace")
    ap.add_argument("--occurrence", type=int, help="which match of --anchor-text (1-based)")
    ap.add_argument("--paragraph", type=int, help="replace the Nth paragraph's text")
    ap.add_argument("--part", help="pin the search to one story part (e.g. header1, footnotes)")
    ap.add_argument("--date", help="ISO-8601 timestamp (default: now, UTC)")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_only(args.file).replace_tracked(args.file, _anchor(args), args.text,
                                                  author=args.author, date=args.date)
        print("replaced (tracked)")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
