"""Delete text from a Word (.docx) file as a tracked deletion.

The anchored phrase (or the anchored paragraph's text) is marked as deleted by --author, and
Track Changes recording is turned on. The text stays in the file as a redline until accepted.

    python scripts/delete_text.py --file r.docx \
        --anchor-text "before we circulate this report" --author "Sam Lee"
    python scripts/delete_text.py --file r.docx --paragraph 3 --author "Sam Lee"
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                       # noqa: E402
from _errors import AnchorNotFound  # noqa: E402


def _anchor(args):
    if args.paragraph is not None:
        return {"paragraph": args.paragraph}
    if args.anchor_text:
        return {"text": args.anchor_text, "occurrence": args.occurrence}
    raise AnchorNotFound('need --anchor-text "phrase" (optional --occurrence N) or --paragraph N')


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Delete text as a tracked change in a Word file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--author", required=True)
    ap.add_argument("--anchor-text", dest="anchor_text", help="the phrase to mark deleted")
    ap.add_argument("--occurrence", type=int, help="which match of --anchor-text (1-based)")
    ap.add_argument("--paragraph", type=int, help="mark the Nth paragraph's text deleted")
    ap.add_argument("--date", help="ISO-8601 timestamp (default: now, UTC)")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_only(args.file).delete_tracked(args.file, _anchor(args),
                                                 author=args.author, date=args.date)
        print("deleted (tracked)")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
