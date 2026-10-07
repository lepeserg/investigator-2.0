"""Synthetic document regression tests; no real case or legal conclusion is used."""
from pathlib import Path
import hashlib
import shutil
import sys
import tempfile
import unittest
from docx import Document

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'investigator-2-0/scripts'))
import docx_edit
import docx_integrity
import make_docx


class DocumentTests(unittest.TestCase):
    def setUp(self):
        (ROOT/'.test-tmp').mkdir(exist_ok=True)
        self.temp=tempfile.TemporaryDirectory(dir=ROOT/'.test-tmp')
        self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name)

    def test_working_copy_replaces_body_table_footer_and_preserves_source(self):
        source=self.root/'source.docx'; copy=self.root/'working.docx'
        doc=Document();doc.add_paragraph('SYNTHETIC_OLD subject')
        doc.add_table(rows=1,cols=2).cell(0,0).text='SYNTHETIC_OLD table'
        doc.sections[0].footer.paragraphs[0].text='SYNTHETIC_OLD footer'
        doc.save(source);before=source.read_bytes();shutil.copy2(source,copy)
        report=docx_edit.edit_file(str(copy),[{'op':'replace_all','old':'SYNTHETIC_OLD','new':'SYNTHETIC_NEW','expect':3}],backup=False)
        self.assertEqual(source.read_bytes(),before)
        after=Document(copy)
        self.assertIn('SYNTHETIC_NEW',after.paragraphs[0].text)
        self.assertIn('SYNTHETIC_NEW',after.tables[0].cell(0,0).text)
        self.assertIn('SYNTHETIC_NEW',after.sections[0].footer.paragraphs[0].text)
        self.assertEqual(len(after.tables),1)
        self.assertTrue(docx_integrity.check(str(copy)))

    def test_failed_preflight_leaves_file_unchanged(self):
        file=self.root/'draft.docx';doc=Document();doc.add_paragraph('SYNTHETIC TEXT');doc.save(file)
        before=file.read_bytes()
        with self.assertRaises(ValueError):
            docx_edit.edit_file(str(file),[{'op':'replace','old':'ABSENT TEXT','new':'NEW TEXT','expect':1}],backup=False)
        self.assertEqual(file.read_bytes(),before)

    def test_missing_letterhead_is_explicit_not_a_false_success(self):
        out=self.root/'request.docx'
        with self.assertRaises(ValueError):
            make_docx.build_letterhead_doc(str(out),[],blank_template_path=None)
        self.assertFalse(out.exists())


if __name__=='__main__':unittest.main()
