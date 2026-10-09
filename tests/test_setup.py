"""Behavior checks for configuration, launch boundaries and setup failures."""
from pathlib import Path
import copy
import json
import subprocess
import sys
import tempfile
import unittest
import unittest.mock
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import bootstrap
import local_config
import run


class SetupTests(unittest.TestCase):
    def setUp(self):
        self.template = json.loads((ROOT/'config.example.json').read_text(encoding='utf-8'))
        self.temp_root = ROOT/'.test-tmp'
        if not self.temp_root.resolve().is_relative_to(ROOT.resolve()):
            raise RuntimeError('Test directory escaped repository')
        self.temp_root.mkdir(exist_ok=True)

    def test_empty_profile_does_not_invent_identity(self):
        local_config.validate(self.template, self.template)
        self.assertTrue(all(value == '' for value in self.template['signer'].values()))
        self.assertIsNone(self.template['correspondence']['requested_reply_days'])

    def test_project_data_cannot_be_saved_under_public_repository(self):
        config = copy.deepcopy(self.template)
        config['projects_root'] = str(ROOT/'private-cases')
        with self.assertRaises(ValueError):
            local_config.validate(config, self.template)

    def test_unknown_secret_field_rejected(self):
        config = copy.deepcopy(self.template)
        config['token'] = 'synthetic-value'
        with self.assertRaises(ValueError):
            local_config.validate(config, self.template)

    def test_incorrect_types_rejected(self):
        config = copy.deepcopy(self.template)
        config['correspondence']['requested_reply_days'] = True
        with self.assertRaises(ValueError):
            local_config.validate(config, self.template)

    def test_profile_mapping_does_not_invent_authority_or_identity(self):
        config = copy.deepcopy(self.template)
        mapped = local_config.legacy_constants(config)
        self.assertNotIn('следователь', mapped)
        self.assertNotIn('руководитель', mapped)
        config['signer']['full_name'] = 'УЧЕБНЫЙ ПОДПИСАНТ'
        config['signer']['rank'] = 'УЧЕБНОЕ ЗВАНИЕ'
        mapped = local_config.legacy_constants(config)
        self.assertEqual(mapped['следователь']['фио'], 'УЧЕБНЫЙ ПОДПИСАНТ')
        self.assertNotIn('инициалы_фам', mapped['следователь'])
        self.assertNotIn('орган_родительный', mapped)

    def test_corrupt_install_stamp_is_not_treated_as_ready(self):
        with tempfile.TemporaryDirectory(dir=self.temp_root) as folder:
            root = Path(folder)
            (root/'requirements').mkdir()
            (root/'requirements'/'base.txt').write_text('test-package==1\n')
            (root/'.venv').mkdir()
            (root/'.venv'/'investigator-install.json').write_text('[]')
            with self.assertRaises(RuntimeError):
                bootstrap.ensure(offline=True, root=root)

    def test_cannot_launch_script_outside_tool_directory(self):
        with self.assertRaises(ValueError):
            run.select_script('../../bootstrap.py')
        self.assertEqual(run.select_script('extract_docx.py').name, 'extract_docx.py')

    def _run_main(self, argv, ensure):
        with patch.object(sys, 'argv', ['run.py', *argv]), \
                patch.object(run.local_config, 'load', return_value={}), \
                patch.object(run.bootstrap, 'ensure', ensure), \
                patch.object(run, 'environment', return_value={}), \
                patch.object(run.subprocess, 'call', return_value=0) as call, \
                patch('sys.stderr') as stderr:
            code = run.main()
        return code, call, ''.join(str(c.args[0]) for c in stderr.write.call_args_list if c.args)

    def test_run_does_not_install_by_default(self):
        for argv in (['extract_docx.py'], ['--offline', 'extract_docx.py']):
            ensure = unittest.mock.Mock(side_effect=RuntimeError('Environment is not prepared.'))
            code, call, err = self._run_main(argv, ensure)
            ensure.assert_called_once_with('base', offline=True)
            call.assert_not_called()
            self.assertNotEqual(code, 0)
            self.assertIn('start.cmd', err)
            self.assertIn('--install', err)

    def test_run_installs_only_with_explicit_flag(self):
        ensure = unittest.mock.Mock(return_value=Path('python'))
        code, call, _ = self._run_main(['--install', 'extract_docx.py', '--help'], ensure)
        ensure.assert_called_once_with('base', offline=False)
        self.assertEqual(code, 0)
        self.assertEqual(call.call_args.args[0][2:], ['--help'])

    def test_offline_never_creates_environment_or_installs(self):
        with tempfile.TemporaryDirectory(dir=self.temp_root) as folder:
            root = Path(folder)
            (root/'requirements').mkdir()
            (root/'requirements'/'base.txt').write_text('test-package==1\n')
            with patch.object(bootstrap.venv, 'EnvBuilder') as builder, patch.object(bootstrap.subprocess, 'run') as process:
                with self.assertRaises(RuntimeError):
                    bootstrap.ensure(offline=True, root=root)
                builder.assert_not_called()
                process.assert_not_called()
            self.assertFalse((root/'.venv').exists())

    def test_failed_install_is_not_marked_successful(self):
        with tempfile.TemporaryDirectory(dir=self.temp_root) as folder:
            root = Path(folder)
            (root/'requirements').mkdir()
            (root/'requirements'/'base.txt').write_text('test-package==1\n')
            executable = bootstrap.python_path(root)
            executable.parent.mkdir(parents=True)
            executable.touch()
            stamp = root/'.venv'/'investigator-install.json'
            stamp.write_text('{"profiles": ["base"], "fingerprint": "stale"}')
            with patch.object(bootstrap, 'run_checked', side_effect=subprocess.CalledProcessError(1, 'synthetic-install')):
                with self.assertRaises(subprocess.CalledProcessError):
                    bootstrap.ensure(root=root)
            self.assertFalse(stamp.exists())


if __name__ == '__main__':
    unittest.main()
