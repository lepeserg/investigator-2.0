"""Profile-driven requirement selection in check_env.py."""
from pathlib import Path
import contextlib
import importlib.util
import io
import json
import shutil
import unittest
from unittest import mock

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

    def test_missing_marker_checks_legacy_full_set(self):
        profiles, source = check_env.detect_profiles(str(self.root))
        self.assertEqual(profiles, {'base', 'documents', 'ocr'})
        self.assertNotIn('audio', profiles)
        self.assertEqual(source, check_env.LEGACY_SOURCE)

    def test_corrupt_marker_checks_legacy_full_set(self):
        self.write('.venv/investigator-install.json', '{not json')
        self.assertEqual(check_env.detect_profiles(str(self.root))[0], {'base', 'documents', 'ocr'})

    def test_marker_keeps_profile_selection(self):
        self.write('.venv/investigator-install.json', {'profiles': ['base']})
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
        # bootstrap.py не принимает documents — подсказка не должна его предлагать.
        self.assertNotIn('bootstrap', check_env.install_hint('documents'))
        self.assertIn('setup-windows.ps1 -Profile documents', check_env.install_hint('documents'))


class MainRequirementTests(unittest.TestCase):
    """main() с подменёнными находками: какие пробелы роняют код возврата."""

    def run_main(self, argv, word=(True, 'word'), lo=None, gs='gs', pdfium=True,
                 detected=None):
        def has_module(name, heavy=False):
            return pdfium if name == 'pypdfium2' else True
        patches = [
            mock.patch.object(check_env, 'find_word_com', return_value=word),
            mock.patch.object(check_env, 'find_libreoffice', return_value=lo),
            mock.patch.object(check_env, 'find_ghostscript', return_value=gs),
            mock.patch.object(check_env, 'find_tesseract', return_value=None),
            mock.patch.object(check_env, 'find_es', return_value=None),
            mock.patch.object(check_env, 'find_ffmpeg', return_value=None),
            mock.patch.object(check_env, '_has_module', side_effect=has_module),
        ]
        if detected is not None:
            patches.append(mock.patch.object(check_env, 'detect_profiles', return_value=detected))
        out = io.StringIO()
        with contextlib.ExitStack() as stack:
            for p in patches:
                stack.enter_context(p)
            stack.enter_context(contextlib.redirect_stdout(out))
            code = check_env.main(argv)
        return code, out.getvalue()

    def test_windows_word_com_required_even_with_libreoffice(self):
        code, out = self.run_main(['--profile', 'documents'], word=(False, 'нет Word'), lo='soffice')
        self.assertEqual(code, 1)
        self.assertIn('[MISS] pywin32 / Word', out)

    def test_windows_word_without_libreoffice_is_ok(self):
        code, _ = self.run_main(['--profile', 'documents'], word=(True, 'ok'), lo=None)
        self.assertEqual(code, 0)

    def test_non_windows_requires_libreoffice(self):
        self.assertEqual(self.run_main(['--profile', 'documents'], word=(None, 'не Windows'), lo=None)[0], 1)
        self.assertEqual(self.run_main(['--profile', 'documents'], word=(None, 'не Windows'), lo='soffice')[0], 0)

    def test_no_markers_require_legacy_set(self):
        legacy = (check_env.expand_profiles(check_env.LEGACY_PROFILES), check_env.LEGACY_SOURCE)
        code, out = self.run_main([], detected=legacy)
        self.assertEqual(code, 1)  # tesseract не найден — OCR обязателен, как раньше
        self.assertIn('Метки профилей', out)
        self.assertNotIn('[MISS] ffmpeg', out)  # audio по-прежнему не обязателен

    def test_ocr_engine_ghostscript_or_pypdfium2(self):
        def ocr_gap(gs, pdfium):
            _, out = self.run_main(['--profile', 'ocr'], gs=gs, pdfium=pdfium)
            return '[MISS] движок PDF' in out
        self.assertFalse(ocr_gap('gs', False))
        self.assertFalse(ocr_gap(None, True))
        self.assertTrue(ocr_gap(None, False))

    def test_unknown_arguments_are_ignored(self):
        code, _ = self.run_main(['--bogus', 'x', '--profile', 'base'])
        self.assertEqual(code, 0)


if __name__ == '__main__':
    unittest.main()
