"""Exercise installer failure and download boundaries without installing software."""
from pathlib import Path
import copy
import hashlib
import io
import json
import os
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import setup_windows as setup
import system_tools


class WindowsSetupTests(unittest.TestCase):
    def setUp(self):
        self.config = json.loads((ROOT/'config.example.json').read_text(encoding='utf-8'))
        (ROOT/'.test-tmp').mkdir(exist_ok=True)

    def test_native_plan_only_contains_needed_capability(self):
        with patch.object(system_tools, 'probe', return_value=False), patch.object(system_tools, 'word_registered', return_value=False):
            self.assertEqual(setup.missing_tools(setup.selected('base'), {}), [])
            self.assertEqual(setup.missing_tools(setup.selected('ocr'), {}), ['tesseract'])
            self.assertEqual(setup.missing_tools(setup.selected('audio'), {}), ['ffmpeg'])
            self.assertEqual(setup.missing_tools(setup.selected('documents'), {}), ['soffice'])

    def test_existing_word_does_not_request_libreoffice(self):
        with patch.object(system_tools, 'word_registered', return_value=True), patch.object(system_tools, 'probe') as probe:
            self.assertEqual(setup.missing_tools(setup.selected('documents'), {}), [])
            probe.assert_not_called()

    def test_cancel_does_not_install_or_create_files(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as tmp:
            with patch.object(setup, 'missing_tools', return_value=['ffmpeg']), patch.object(setup, 'install_package') as install, patch.object(setup.bootstrap, 'ensure') as ensure:
                self.assertEqual(setup.prepare('audio', self.config, Path(tmp), confirm=lambda _: 'n'), 2)
                install.assert_not_called()
                ensure.assert_not_called()
                self.assertEqual(list(Path(tmp).iterdir()), [])

    def test_check_passes_offline_and_never_calls_install(self):
        with patch.object(setup, 'missing_tools', return_value=['tesseract']), patch.object(setup, 'languages_available', return_value=False), patch.object(setup.bootstrap, 'ensure', side_effect=RuntimeError('missing')) as ensure, patch.object(setup, 'install_package') as install, patch.object(setup, 'download_language') as download:
            issues = setup.check(setup.selected('ocr'), self.config)
            self.assertTrue(issues)
            self.assertTrue(all(call.kwargs['offline'] for call in ensure.call_args_list))
            install.assert_not_called()
            download.assert_not_called()

    def test_failed_native_install_stops_before_python_install(self):
        with patch.object(setup, 'missing_tools', return_value=['ffmpeg']), patch.object(setup.shutil, 'which', return_value='winget'), patch.object(setup, 'install_package', side_effect=subprocess.CalledProcessError(1, 'winget')), patch.object(setup.bootstrap, 'ensure') as ensure:
            with self.assertRaises(subprocess.CalledProcessError):
                setup.prepare('audio', self.config, confirm=lambda _: 'y')
            ensure.assert_not_called()

    def test_successful_installer_but_missing_program_is_failure(self):
        with patch.object(setup, 'missing_tools', return_value=['ffmpeg']), patch.object(setup.shutil, 'which', return_value='winget'), patch.object(setup, 'install_package'), patch.object(system_tools, 'probe', return_value=False), patch.object(setup.bootstrap, 'ensure') as ensure:
            with self.assertRaisesRegex(RuntimeError, 'still unavailable'):
                setup.prepare('audio', self.config, confirm=lambda _: 'y')
            ensure.assert_not_called()

    def test_install_package_prefers_user_scope(self):
        done = subprocess.CompletedProcess([], 0)
        with patch.object(setup.shutil, 'which', return_value='winget'), patch.object(setup.subprocess, 'run', return_value=done) as run:
            setup.install_package('ffmpeg')
        self.assertEqual(run.call_count, 1)
        command = run.call_args.args[0]
        self.assertEqual(command[-2:], ['--scope', 'user'])
        self.assertIn(system_tools.PACKAGES['ffmpeg'], command)

    def test_install_package_falls_back_to_default_scope(self):
        calls = []
        def fake_run(command, **kwargs):
            calls.append((command, kwargs))
            return subprocess.CompletedProcess(command, 1 if '--scope' in command else 0)
        with patch.object(setup.shutil, 'which', return_value='winget'), patch.object(setup.subprocess, 'run', side_effect=fake_run), patch('sys.stdout', new_callable=io.StringIO) as out:
            setup.install_package('tesseract')
        self.assertEqual(len(calls), 2)
        self.assertNotIn('--scope', calls[1][0])
        self.assertTrue(calls[1][1].get('check'))
        self.assertIn('--scope user', out.getvalue())

    def test_install_package_fallback_failure_propagates(self):
        def fake_run(command, **kwargs):
            if kwargs.get('check'):
                raise subprocess.CalledProcessError(5, command)
            return subprocess.CompletedProcess(command, 1)
        with patch.object(setup.shutil, 'which', return_value='winget'), patch.object(setup.subprocess, 'run', side_effect=fake_run), patch('sys.stdout', new_callable=io.StringIO):
            with self.assertRaises(subprocess.CalledProcessError):
                setup.install_package('soffice')

    def test_hash_mismatch_preserves_existing_file_and_removes_partial(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as tmp:
            folder = Path(tmp)
            target = folder/'rus.traineddata'
            target.write_bytes(b'old-data')
            with self.assertRaisesRegex(RuntimeError, 'Checksum mismatch'):
                setup.download_language('rus', folder, opener=lambda *a, **k: io.BytesIO(b'wrong'))
            self.assertEqual(target.read_bytes(), b'old-data')
            self.assertEqual(list(folder.iterdir()), [target])

    def test_verified_download_reused_without_network(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as tmp:
            payload = b'synthetic-model'
            with patch.dict(setup.LANGUAGES, {'rus': hashlib.sha256(payload).hexdigest()}):
                setup.download_language('rus', Path(tmp), opener=lambda *a, **k: io.BytesIO(payload))
                def no_network(*a, **k):
                    raise AssertionError('Unexpected repeat download')
                setup.download_language('rus', Path(tmp), opener=no_network)

    def test_project_languages_and_explicit_override(self):
        with tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp') as tmp, patch.dict(os.environ, {}, clear=True):
            root = Path(tmp)
            folder = root/'runtime'/'tessdata'
            folder.mkdir(parents=True)
            for lang in setup.LANGUAGES:
                (folder/f'{lang}.traineddata').touch()
            self.assertEqual(system_tools.environment(self.config, root)['TESSDATA_PREFIX'], str(folder))
            config = copy.deepcopy(self.config)
            config['tools']['tessdata_prefix'] = str(root/'explicit-data')
            self.assertEqual(system_tools.environment(config, root)['TESSDATA_PREFIX'], config['tools']['tessdata_prefix'])
            os.environ['TESSDATA_PREFIX'] = str(root/'environment-data')
            self.assertEqual(system_tools.environment(self.config, root)['TESSDATA_PREFIX'], os.environ['TESSDATA_PREFIX'])


if __name__ == '__main__':
    unittest.main()
