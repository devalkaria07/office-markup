"""Delete a comment. A reply is removed on its own; deleting a thread's root removes the
whole thread (root + all replies).

    python scripts/delete_comment.py --file r.docx --comment 1
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Delete a comment or reply from an Office file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--comment", required=True, help="id of the comment or reply to delete")
    args = ap.parse_args(argv)

    def go():
        mod = _cli.module_for(args.file)
        mod.delete(args.file, args.comment)
        print("deleted")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
