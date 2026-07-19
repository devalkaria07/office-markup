"""Apply a whole LIST of edits to one Word (.docx) file in a single pass.

The edits file is JSON (schema in SKILL.md): tracked or direct replace/insert/delete,
plus comment / reply / resolve — applied sequentially in list order, ALL-OR-NOTHING
(if any entry fails, the file is left completely untouched and the error names the
failing entry). Tracked is the default; direct edits say "tracked": false.

    python scripts/apply_edits.py --file r.docx --edits edits.json --author "Alex Morgan"
    python scripts/apply_edits.py --file r.docx --edits edits.json --dry-run
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli                       # noqa: E402
import _docx_batch                # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Apply a JSON list of edits to a Word file (all-or-nothing).")
    ap.add_argument("--file", required=True)
    ap.add_argument("--edits", required=True, help="path to the JSON edits file")
    ap.add_argument("--author", help="author for entries that don't set their own")
    ap.add_argument("--dry-run", dest="dry_run", action="store_true",
                    help="validate and execute in memory only; the file is not modified")
    ap.add_argument("--json", dest="as_json", action="store_true",
                    help="print the result (applied count, per-edit ids) as JSON")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_path(args.file)
        spec = json.loads(Path(args.edits).read_text(encoding="utf-8-sig"))
        out = _docx_batch.apply_edits(args.file, spec, cli_author=args.author,
                                      dry_run=args.dry_run)
        if args.as_json:
            print(json.dumps(out, indent=2))
        else:
            note = " (dry run - file untouched)" if args.dry_run else ""
            print(f"applied {out['applied']} edit(s){note}")
            for r in out["results"]:
                if "id" in r:
                    print(f"  edit {r['index']} ({r['action']}): id {r['id']}")
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
