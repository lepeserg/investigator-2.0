"""Profile-driven requirement selection in check_env.py."""
from pathlib import Path
import importlib.util
import json
import shutil
import unittest

ROOT = Path(__file__).resolve().parents[1]
_spec = importlib.util.spec_from_file_location(
    'check_env', ROOT/'investigator-2-0'/'scripts'/'check_env.py')
check_env = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(check_env)


class ProfileSelectionTests(unittest.TestCase):
    def setUp(self):
        self.root = ROOT/'.test-tmp'/'check_env'
        if not self.root.resolve().is_relative_to(ROOT.resolve()):
            raise RuntimeError('Test directory escaped repository')
        shutil.rmtree(self.root, ignore_errors=True)
        self.root.mkdir(parents=True)

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def write(self, rel, data):
        path = self.root/rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(data) if not isinstance(data, str) else data, encoding='utf-8')

    def test_expand_always_includes_base_and_drops_unknown(self):
        self.assertEqual(check_env.expand_profiles([]), {'base'})
        self.assertEqual(check_env.expand_profiles(['ocr', 'bogus']), {'base', 'ocr'})
        self.assertEqual(check_env.expand_profiles(['all']), set(check_env.PROFILES))

    def test_missing_marker_falls_back_to_base(self):
        profiles, source = check_env.detect_profiles(str(self.root))
        self.assertEqual(profiles, {'base'})
        self.assertIn('по умолчанию', source)

    def test_corrupt_marker_falls_back_to_base(self):
        self.write('.venv/investigator-install.json', '{not json')
        self.assertEqual(check_env.detect_profiles(str(self.root))[0], {'base'})

    def test_bootstrap_stamp_and_setup_report_are_merged(self):
        self.write('.venv/investigator-install.json', {'profiles': ['base', 'ocr'], 'fingerprint': 'x'})
        self.write('runtime/setup-report.json', {'profile': 'documents', 'checks': 'passed'})
        profiles, source = check_env.detect_profiles(str(self.root))
        self.assertEqual(profiles, {'base', 'ocr', 'documents'})
        self.assertIn('investigator-install.json', source)

    def test_cli_profile_overrides_marker(self):
        self.write('.venv/investigator-install.json', {'profiles': ['base', 'audio']})
        profiles, source = check_env.resolve_profiles(['ocr'], str(self.root))
        self.assertEqual(profiles, {'base', 'ocr'})
        self.assertIn('--profile', source)

    def test_unselected_profile_does_not_fail(self):
        checks = [('base', True), ('ocr', False), ('audio', False)]
        self.assertTrue(check_env.overall_ok(checks, {'base'}))
        self.assertEqual(check_env.classify(False, 'ocr', {'base'}), 'skip')

    def test_selected_profile_gap_fails(self):
        checks = [('base', True), ('ocr', False)]
        self.assertFalse(check_env.overall_ok(checks, {'base', 'ocr'}))
        self.assertFalse(check_env.overall_ok([('base', False)], {'base'}))
        self.assertEqual(check_env.classify(False, 'ocr', {'base', 'ocr'}), 'miss')

    def test_hints_use_current_installers(self):
        for profile in check_env.PROFILES:
            hint = check_env.install_hint(profile)
            self.assertIn('start.cmd', hint)
            self.assertNotIn('runtime/python', hint)
        self.assertIn('bootstrap.py --profile ocr', check_env.install_hint('ocr'))


if __name__ == '__main__':
    unittest.main()
