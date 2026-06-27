"""Reply to an existing comment thread (by the parent comment's id).

    python scripts/reply_comment.py --file r.docx --parent 0 \
        --author "Sam Lee" --text "Confirmed."

Use list_comments.py --json first to find the id. Prints the new reply id.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Reply to a threaded comment in an Office file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--parent", required=True, help="id of the comment (or any comment in the thread) to reply to")
    ap.add_argument("--text", required=True)
    ap.add_argument("--author", required=True)
    ap.add_argument("--initials")
    args = ap.parse_args(argv)

    def go():
        mod = _cli.module_for(args.file)
        rid = mod.reply(args.file, args.parent, args.text, author=args.author, initials=args.initials)
        print(rid)
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
