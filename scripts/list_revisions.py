"""List tracked changes (revisions) in a Word (.docx) file.

    python scripts/list_revisions.py --file report.docx
    python scripts/list_revisions.py --file report.docx --json

--json emits {"revisions": [...]} where each record has id, author, date, type, text, paragraph,
location, context, move_name and table — the machine-readable view used to pick a revision id
before accepting or rejecting. The id is positional (document order); re-list after each change.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="List tracked changes in a Word file.")
    ap.add_argument("--file", required=True)
    ap.add_argument("--json", action="store_true", help="machine-readable JSON output")
    args = ap.parse_args(argv)

    def go():
        records = _cli.docx_only(args.file).list_revisions(args.file)
        if args.json:
            print(json.dumps({"revisions": records}, indent=2, ensure_ascii=False))
        else:
            _cli.print_revisions(records)
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
