"""Run case-management regressions in temporary synthetic projects; --help supported."""
from pathlib import Path
import argparse
import subprocess
import sys


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--verbose', action='store_true')
    args = parser.parse_args()
    test = Path(__file__).resolve().parents[2] / 'investigator-2-0-case-management' / 'scripts' / 'test_case_management.py'
    return subprocess.call([sys.executable, '-B', str(test), '-v' if args.verbose else '-q'])


if __name__ == '__main__':
    raise SystemExit(main())
