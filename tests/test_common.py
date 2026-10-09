"""_common: unified tool lookup, lock check and atomic save."""
from pathlib import Path
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT/'investigator-2-0/scripts'))
import _common

ENV_KEYS = ('ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432', 'LOCALAPPDATA')


class FindToolTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        env = {k: v for k, v in os.environ.items() if k not in ENV_KEYS}
        for p in (patch.dict(os.environ, env, clear=True),
                  patch.object(_common.shutil, 'which', side_effect=self._which),
                  patch.object(_common, '_config_extra_path', return_value=[]),
                  # keep real C:\ and /usr/bin paths of the host machine out of the test
                  patch.object(_common, '_install_roots', side_effect=self._roots)):
            p.start()
            self.addCleanup(p.stop)
        self.on_path = {}
        self.roots = []

    def _which(self, name, path=None):
        if path is None:
            return self.on_path.get(name)
        for d in path.split(os.pathsep):
            c = Path(d)/name
            if c.is_file():
                return str(c)
        return None

    def _roots(self):
        return list(self.roots)

    def _touch(self, *parts):
        p = self.root.joinpath(*parts)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b'')
        return str(p)

    def test_path_wins(self):
        self.roots = [str(self.root)]
        self._touch('Tesseract-OCR', 'tesseract.exe')
        self.on_path['tesseract'] = '/bin/tesseract'
        self.assertEqual(_common.find_tool('tesseract'), '/bin/tesseract')

    def test_tesseract_in_program_files(self):
        self.roots = [str(self.root/'none'), str(self.root/'pf')]
        exe = self._touch('pf', 'Tesseract-OCR', 'tesseract.exe')
        self.assertEqual(_common.find_tool('tesseract'), exe)

    def test_soffice_alias_and_libreoffice_name_on_path(self):
        self.on_path['libreoffice'] = '/opt/lo'
        self.assertEqual(_common.find_tool('libreoffice'), '/opt/lo')

    def test_soffice_in_install_dir(self):
        self.roots = [str(self.root)]
        exe = self._touch('LibreOffice', 'program', 'soffice.exe')
        with patch.dict(_common._TOOLS['soffice'], fixed=(), globs=()):
            self.assertEqual(_common.find_tool('soffice'), exe)

    def test_soffice_glob_takes_newest(self):
        old = self._touch('LibreOffice 7', 'program', 'soffice.exe')
        new = self._touch('LibreOffice 24', 'program', 'soffice.exe')
        pattern = str(self.root/'LibreOffice*'/'program'/'soffice.exe')
        with patch.dict(_common._TOOLS['soffice'], fixed=(), globs=(pattern,)):
            self.assertEqual(_common.find_tool('soffice'), max(old, new))

    def test_winget_links_and_config_extra_path(self):
        os.environ['LOCALAPPDATA'] = str(self.root)
        exe = self._touch('Microsoft', 'WinGet', 'Links', 'antiword')
        self.assertEqual(_common.find_tool('antiword'), exe)
        extra = self._touch('extra', 'tesseract')
        with patch.object(_common, '_config_extra_path', return_value=[str(Path(extra).parent)]):
            self.assertEqual(_common.find_tool('tesseract'), extra)

    def test_order_matches_run_py_path(self):
        """extra_path → Program Files → WinGet\\Links, as system_tools.environment builds PATH."""
        os.environ['LOCALAPPDATA'] = str(self.root/'local')
        self.roots = [str(self.root/'pf')]
        links = self._touch('local', 'Microsoft', 'WinGet', 'Links', 'tesseract')
        pf = self._touch('pf', 'Tesseract-OCR', 'tesseract.exe')
        self.assertEqual(_common.find_tool('tesseract'), pf)
        extra = self._touch('extra', 'tesseract')
        with patch.object(_common, '_config_extra_path', return_value=[str(Path(extra).parent)]):
            self.assertEqual(_common.find_tool('tesseract'), extra)
        os.remove(pf)
        self.assertEqual(_common.find_tool('tesseract'), links)

    def test_not_found(self):
        with patch.dict(_common._TOOLS['antiword'], fixed=()):
            self.assertIsNone(_common.find_tool('antiword'))



class SearchLocationTests(unittest.TestCase):
    def test_install_roots_include_env_and_defaults(self):
        with patch.dict(os.environ, {'ProgramW6432': 'D:\\PF'}):
            roots = _common._install_roots()
        self.assertIn(r'C:\Program Files', roots)
        self.assertIn(r'C:\Program Files (x86)', roots)
        self.assertIn('D:\\PF', roots)

    def test_candidates_cover_all_previous_copies(self):
        """Union of locations from check_env, ocr_smart, opis_verify, make_statcards."""
        with patch.dict(os.environ, {'LOCALAPPDATA': 'L'}):
            tess = _common.tool_candidates('tesseract')
            lo = _common.tool_candidates('soffice')
        for c in (r'C:\Program Files', r'C:\Program Files (x86)'):
            self.assertIn(os.path.join(c, 'Tesseract-OCR', 'tesseract.exe'), tess)
            self.assertIn(os.path.join(c, 'LibreOffice', 'program', 'soffice.exe'), lo)
        self.assertIn(os.path.join('L', 'LibreOffice', 'program', 'soffice.exe'), lo)
        for c in ('/usr/bin/soffice', '/usr/bin/libreoffice',
                  '/Applications/LibreOffice.app/Contents/MacOS/soffice'):
            self.assertIn(c, lo)
        self.assertEqual(len(_common._TOOLS['soffice']['globs']), 2)


class LockAndSaveTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.dir = Path(temp.name)
        self.path = str(self.dir/'doc.docx')
        doc = Document()
        doc.add_paragraph('исходный')
        doc.save(self.path)

    def test_word_lock_file_raises(self):
        (self.dir/'~$doc.docx').write_bytes(b'')
        with self.assertRaises(_common.DocxLockedError) as ctx:
            _common.check_locked(self.path)
        self.assertIn('ОТКРЫТ В WORD', str(ctx.exception))

    def test_missing_file_is_not_locked(self):
        _common.check_locked(str(self.dir/'absent.docx'))

    def test_save_atomic_runs_hook_and_replaces(self):
        doc = Document(self.path)
        doc.add_paragraph('новый')
        seen = []
        _common.save_atomic(doc, self.path, tmp_prefix='~t_', before_replace=seen.append)
        self.assertEqual(len(seen), 1)
        self.assertTrue(Path(seen[0]).name.startswith('~t_'))
        self.assertIn('новый', [p.text for p in Document(self.path).paragraphs])
        self.assertEqual([p.name for p in self.dir.iterdir() if p.name.startswith('~')], [])

    def test_save_atomic_permission_error_is_locked_and_cleans_tmp(self):
        doc = Document(self.path)
        with patch.object(_common.os, 'replace', side_effect=PermissionError(13, 'busy')):
            with self.assertRaises(_common.DocxLockedError):
                _common.save_atomic(doc, self.path)
        self.assertEqual([p.name for p in self.dir.iterdir() if p.name.startswith('~')], [])

    def test_save_atomic_broken_zip(self):
        doc = Document(self.path)
        with patch.object(doc, 'save', side_effect=lambda t: Path(t).write_bytes(b'not zip')):
            with self.assertRaises(_common.BrokenSaveError):
                _common.save_atomic(doc, self.path)
        self.assertEqual([p.name for p in self.dir.iterdir() if p.name.startswith('~')], [])

    def test_docx_edit_reexports_same_error(self):
        import docx_edit
        self.assertIs(docx_edit.DocxLockedError, _common.DocxLockedError)
        self.assertIs(docx_edit.word_lock_path, _common.word_lock_path)


if __name__ == '__main__':
    unittest.main()
