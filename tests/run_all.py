"""Run every test file and print one summary line each.

    python tests/run_all.py

Most tests read the real production register from data/source/ (not in the repository);
test_end_to_end.py makes its own sample registers and runs without it.
"""
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TESTS = ["test_settings.py", "test_chart.py", "test_production.py", "test_report.py",
         "test_report2.py", "test_excel.py", "test_pdf.py", "test_word.py", "test_web.py", "test_end_to_end.py"]


def main() -> int:
    failed = 0
    for name in TESTS:
        start = time.time()
        p = subprocess.run([sys.executable, "-W", "ignore", str(ROOT / "tests" / name)], cwd=ROOT,
                           capture_output=True, text=True, encoding="utf-8", errors="replace",
                           env={**__import__("os").environ, "PYTHONIOENCODING": "utf-8"})
        lines = p.stdout.splitlines()
        passed = sum(ln.startswith("PASS") for ln in lines)
        fails = [ln for ln in lines if ln.startswith("FAIL")]
        ok = p.returncode == 0
        failed += not ok
        print(f"{'ok  ' if ok else 'FAIL'} {name:24} {passed:4} checks  {time.time() - start:5.1f}s")
        for ln in fails:
            print(f"       {ln[:150]}")
        if not ok and not fails:
            print("       " + (p.stderr.strip().splitlines() or ["(no output)"])[-1][:150])
    print("\nall test files pass" if not failed else f"\n{failed} test file(s) failed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
