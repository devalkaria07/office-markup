"""List threaded comments in a Word/Excel/PowerPoint file.

    python scripts/list_comments.py --file report.docx
    python scripts/list_comments.py --file book.xlsx --json

--json emits {"comments": [...]} where each record has id, thread_id, parent_id, author,
date, text, resolved, anchor, anchor_text, context and location — the machine-readable view
used to pick a comment id before replying/resolving/deleting.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="List threaded comments in an Office file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--json", action="store_true", help="machine-readable JSON output")
    args = ap.parse_args(argv)

    def go():
        records = _cli.module_for(args.file).list_comments(args.file)
        if args.json:
            print(json.dumps({"comments": records}, indent=2, ensure_ascii=False))
        else:
            _cli.print_human(records)
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
