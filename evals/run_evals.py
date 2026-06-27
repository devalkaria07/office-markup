"""Eval suite — runs every tests/smoke_*.py and fails if any fails.

This is the gate `scripts/release.py --check` calls before packaging. Each smoke test drives a
real Office file through add -> reply -> resolve/reopen -> delete and asserts the structure,
threading, resolved flag, byte-stability and that the relevant Office library still opens it.

Run: python evals/run_evals.py
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path

TESTS = Path(__file__).resolve().parent.parent / "tests"


def main() -> int:
    smoke = sorted(TESTS.glob("smoke_*.py"))
    if not smoke:
        print("no smoke tests found")
        return 1
    failures = []
    for t in smoke:
        print(f"\n=== {t.name} ===")
        rc = subprocess.call([sys.executable, str(t)])
        if rc != 0:
            failures.append(t.name)
    print("\n" + "=" * 40)
    if failures:
        print(f"EVAL FAILURES ({len(failures)}/{len(smoke)}): {', '.join(failures)}")
        return 1
    print(f"ALL {len(smoke)} EVALS PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
