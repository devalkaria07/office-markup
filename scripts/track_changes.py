"""Turn Word's Track Changes recording on or off in a Word (.docx) file.

    python scripts/track_changes.py --file r.docx --on
    python scripts/track_changes.py --file r.docx --off

This flips the document's <w:trackRevisions/> switch — the same toggle as Word's Review >
Track Changes button. It does not accept or reject any existing changes.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import _cli   # noqa: E402


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Toggle Track Changes recording in a Word file.")
    ap.add_argument("--file", required=True)
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--on", action="store_true", help="start recording tracked changes")
    g.add_argument("--off", action="store_true", help="stop recording tracked changes")
    args = ap.parse_args(argv)

    def go():
        _cli.docx_only(args.file).set_tracking(args.file, bool(args.on))
        print("track changes: " + ("on" if args.on else "off"))
        return _cli.EXIT_OK

    return _cli.run(go)


if __name__ == "__main__":
    sys.exit(main())
