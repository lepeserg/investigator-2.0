"""Release infrastructure: publication allowlist, duplicated requirement lists, check_account."""
from pathlib import Path
import contextlib
import io
import shutil
import subprocess
import sys
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT/'investigator-2-0'/'scripts'
sys.path.insert(0, str(SCRIPTS))
import check_account  # noqa: E402

# Files of these kinds are always meant for publication. The root .gitignore is an
# allowlist, so a new file that lacks its own entry silently stays out of the release.
PUBLISHED_GLOBS = (
    'investigator-2-0/scripts/*.py',
    'investigator-2-0*/SKILL.md',
    'investigator-2-0*/references/*.md',
    'investigator-2-0*/agents/*.yaml',
    'docs/*.md',
    '*.py',
    '*.cmd',
    '*.ps1',
    'README.md',
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


class CheckAccountTests(unittest.TestCase):
    BIK = '044525225'
    # Public correspondent account paired with the BIK above (Bank of Russia branch rule).
    CORR = '30101810400000000225'

    def _client_account(self, prefix):
        draft = '40817810' + '0' + '38000000001'
        key = check_account.control_key(draft, prefix)
        return draft[:check_account.KEY_POS] + key + draft[check_account.KEY_POS + 1:]

    def test_normalize_strips_separators(self):
        self.assertEqual(check_account.normalize('301 018.10-4000 0000 0225'), self.CORR)

    def test_correspondent_account_uses_bank_of_russia_branch(self):
        res = check_account.check(self.CORR, self.BIK)
        self.assertEqual((res['code'], res['условный'], res['валюта']), (0, '025', 'рубли'))
        self.assertEqual(res['ключ_нужен'], self.CORR[check_account.KEY_POS])

    def test_client_account_uses_last_three_bik_digits(self):
        account = self._client_account(self.BIK[6:9])
        res = check_account.check(account, self.BIK)
        self.assertEqual((res['code'], res['условный']), (0, '225'))

    def test_single_wrong_digit_is_detected(self):
        account = self._client_account(self.BIK[6:9])
        broken = account[:-1] + str((int(account[-1]) + 1) % 10)
        res = check_account.check(broken, self.BIK)
        self.assertEqual(res['code'], 1)
        self.assertEqual(res['ключ_нужен'], check_account.control_key(broken, '225'))

    def test_match_only_by_alternative_branch(self):
        account = self._client_account('0' + self.BIK[4:6])
        if check_account.matches(account, self.BIK[6:9]):
            self.skipTest('synthetic account matches both branches')
        self.assertEqual(check_account.check(account, self.BIK)['code'], 2)

    def _main(self, *args):
        with patch.object(sys, 'argv', ['check_account.py', *args]), \
                contextlib.redirect_stdout(io.StringIO()):
            return check_account.main()

    def test_cli_exit_codes(self):
        good = self._client_account('225')
        alt = self._client_account('025')
        self.assertEqual(self._main(self.BIK, good), 0)
        self.assertEqual(self._main(good, self.BIK, '--quiet'), 0)
        self.assertEqual(self._main(self.BIK, '12345'), 3)
        self.assertEqual(self._main(good), 3)
        if not check_account.matches(alt, '225'):
            self.assertEqual(self._main(self.BIK, alt), 2)
            self.assertEqual(self._main(self.BIK, alt, '--strict'), 1)
            self.assertEqual(self._main(self.BIK, good, alt), 2)


if __name__ == '__main__':
    unittest.main()
