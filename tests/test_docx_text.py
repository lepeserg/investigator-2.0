"""Shared raw-XML text extraction (word_text.xml_text) and the scripts that use it."""
from pathlib import Path
import sys
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'investigator-2-0/scripts'))
import word_text  # noqa: E402
from word_text import xml_text, docx_xml_text  # noqa: E402

NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
TRACKED = ('<w:p><w:r><w:t xml:space="preserve">Потерпевший </w:t></w:r>'
           '<w:del w:id="1"><w:r><w:delText>Старов</w:delText></w:r></w:del>'
           '<w:ins w:id="2"><w:r><w:t>Новиков</w:t></w:r></w:ins></w:p>')
FIELD = ('<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
         '<w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>'
         '<w:r><w:t>7</w:t></w:r></w:p>')
TABS = ('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs></w:pPr>'
        '<w:r><w:t>1.</w:t></w:r><w:r><w:tab/></w:r>'
        '<w:r><w:t xml:space="preserve">Протокол &amp; опись &quot;А&quot;</w:t></w:r></w:p>')


class XmlTextTests(unittest.TestCase):
    def test_deleted_text_dropped_unless_requested(self):
        self.assertEqual(xml_text(TRACKED), 'Потерпевший Новиков\n')
        self.assertEqual(xml_text(TRACKED, include_deleted=True), 'Потерпевший СтаровНовиков\n')

    def test_field_codes_always_dropped(self):
        self.assertEqual(xml_text(FIELD), '7\n')
        self.assertEqual(xml_text(FIELD, include_deleted=True), '7\n')

    def test_tab_runs_and_tab_stops(self):
        self.assertEqual(xml_text(TABS), '1.\tПротокол & опись "А"\n')
        self.assertEqual(xml_text(TABS, tab=''), '1.Протокол & опись "А"\n')

    def test_t_pattern_does_not_catch_neighbours(self):
        xml = ('<w:p><w:r><w:t/><w:t xml:space="preserve"/><w:tab w:val="x"/>'
               '<w:t>A</w:t></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p>')
        self.assertEqual(xml_text(xml), '\tA\nB\n')

    def test_text_outside_w_t_is_ignored(self):
        xml = ('<w:p><w:r><w:drawing><wp:positionH><wp:posOffset>12345</wp:posOffset>'
               '</wp:positionH></w:drawing><w:t>X</w:t></w:r></w:p>')
        self.assertEqual(xml_text(xml), 'X\n')

    def test_paragraph_mode(self):
        xml = '<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p/><w:p w:rsidR="1"/>' + TRACKED
        self.assertEqual(xml_text(xml, paragraphs=True), ['A', '', '', 'Потерпевший Новиков'])
        self.assertEqual(xml_text('<w:r><w:t>tail</w:t></w:r>', paragraphs=True), ['tail'])
        self.assertEqual(xml_text('', paragraphs=True), [])

    def test_entities_decoded_once(self):
        self.assertEqual(xml_text('<w:t>&amp;lt; &#1040;</w:t>'), '&lt; А')

    def test_docx_xml_text_reads_body_headers_footers(self):
        with tempfile.TemporaryDirectory() as tmp:
            f = Path(tmp) / 'p.docx'
            wrap = '<w:document xmlns:w="' + NS + '"><w:body>{}</w:body></w:document>'
            with zipfile.ZipFile(f, 'w') as z:
                z.writestr('word/document.xml', wrap.format(TRACKED))
                z.writestr('word/header1.xml', '<w:hdr><w:p><w:r><w:t>H</w:t></w:r></w:p></w:hdr>')
                z.writestr('word/footnotes.xml', '<w:p><w:r><w:t>NOTE</w:t></w:r></w:p>')
            self.assertEqual(docx_xml_text(f), 'Потерпевший Новиков\n\nH\n')
            self.assertEqual(docx_xml_text(f, parts='header\\d*'), 'H\n')
            self.assertEqual(docx_xml_text(f, paragraphs=True), ['Потерпевший Новиков', 'H'])
            self.assertEqual(word_text.docx_xml_text(f, include_deleted=True).split('\n')[0],
                             'Потерпевший СтаровНовиков')


if __name__ == '__main__':
    unittest.main()
