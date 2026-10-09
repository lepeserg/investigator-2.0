"""Release infrastructure: publication allowlist and duplicated requirement lists."""
from pathlib import Path
import shutil
import subprocess
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT/'investigator-2-0'/'scripts'
sys.path.insert(0, str(SCRIPTS))

# Files of these kinds are always meant for publication. The root .gitignore is an
# allowlist, so a new file that lacks its own entry silently stays out of the release.
PUBLISHED_GLOBS = (
    'investigator-2-0/scripts/*.py',
    'investigator-2-0/references/*.md',
    'tests/*.py',
    'requirements/*.txt',
    '.github/workflows/*.yml',
)


def _requirement_lines(path):
    lines = []
    for raw in path.read_text(encoding='utf-8').splitlines():
        line = raw.split('#', 1)[0].strip()
        if line:
            lines.append(line)
    return lines


class PublicationAllowlistTests(unittest.TestCase):
    def test_published_kinds_are_not_ignored(self):
        git = shutil.which('git')
        if not git or not (ROOT/'.git').exists():
            self.skipTest('not a git checkout (e.g. unpacked ZIP release)')
        paths = sorted({p.relative_to(ROOT).as_posix()
                        for pattern in PUBLISHED_GLOBS for p in ROOT.glob(pattern) if p.is_file()})
        self.assertTrue(paths)
        # --no-index: judge by the rules only, so tracked files are checked as well.
        result = subprocess.run([git, 'check-ignore', '--no-index', '--stdin'], cwd=ROOT,
                                input='\n'.join(paths) + '\n', capture_output=True,
                                text=True, encoding='utf-8', timeout=60)
        self.assertIn(result.returncode, (0, 1), result.stderr)
        ignored = [line for line in result.stdout.splitlines() if line]
        self.assertEqual(ignored, [], 'Files are excluded by the .gitignore allowlist; '
                         'add explicit "!/<path>" entries after content review: ' + ', '.join(ignored))


class RequirementListsTests(unittest.TestCase):
    def test_scripts_requirements_match_base_profile(self):
        # scripts/requirements.txt stays a literal copy because check_skill.py parses
        # package names from it and does not follow "-r" includes.
        self.assertEqual(_requirement_lines(SCRIPTS/'requirements.txt'),
                         _requirement_lines(ROOT/'requirements'/'base.txt'))


if __name__ == '__main__':
    unittest.main()
