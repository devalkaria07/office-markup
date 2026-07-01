"""Reject tracked changes in a Word (.docx) file.

    python scripts/reject_change.py --file r.docx --rev 3       # reject one change (by list id)
    python scripts/reject_change.py --file r.docx --all         # reject every change
    python scripts/reject_change.py --file r.docx --all --author "Sam Lee"

Revision ids come from list_revisions and are positional (document order); after you accept or
reject one, re-list before targeting another. With --all, prints how many changes were rejected.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                     # noqa: E402
from _errors import CommentError  # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Reject tracked changes in a Word file.")
    ap.add_argument("--file", required=True)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--rev", type=int, help="reject a single change by its list id")
    g.add_argument("--all", action="store_true", help="reject every change")
    ap.add_argument("--author", help="with --all: only changes by this author")
    args = ap.parse_args(argv)

    def go():
        mod = _cli.docx_only(args.file)
        if args.all:
            n = mod.reject_all(args.file, author=args.author)
            print(f"rejected {n} change(s)")
        else:
            if args.author:
                raise CommentError("--author only applies together with --all")
            mod.reject(args.file, args.rev)
            print(f"rejected revision {args.rev}")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
