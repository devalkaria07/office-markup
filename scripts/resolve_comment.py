"""Resolve (or reopen) a comment thread.

    python scripts/resolve_comment.py --file r.docx --comment 0
    python scripts/resolve_comment.py --file r.docx --comment 0 --reopen

Resolving/reopening any comment in a thread affects the whole thread.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Resolve or reopen a comment thread.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--comment", required=True, help="id of any comment in the thread")
    ap.add_argument("--reopen", action="store_true", help="reopen instead of resolve")
    args = ap.parse_args(argv)

    def go():
        mod = _cli.module_for(args.file)
        mod.set_status(args.file, args.comment, resolved=not args.reopen)
        print("reopened" if args.reopen else "resolved")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
