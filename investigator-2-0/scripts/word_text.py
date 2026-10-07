"""Read DOCX text in XML order with explicit tracked-change views.

This module reads text only; it does not accept/reject changes or interpret images.
Without an explicit view, tracked changes stop reading until the user chooses.
Views are reading aids and do not establish an approved document version.
"""
from collections import Counter
import re
import zipfile
from xml.etree import ElementTree as ET

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
MC = '{http://schemas.openxmlformats.org/markup-compatibility/2006}'
VIEWS = ('marked', 'current', 'original')
CHANGES = {'ins': 'ВСТАВКА', 'del': 'УДАЛЕНИЕ',
           'moveTo': 'ПЕРЕМЕЩЕНО СЮДА', 'moveFrom': 'ПЕРЕМЕЩЕНО ОТСЮДА'}
IGNORED = {'pPr', 'rPr', 'tblPr', 'tblGrid', 'trPr', 'tcPr', 'sectPr', 'instrText'}
REVISION_TAGS = set(CHANGES) | {'pPrChange', 'rPrChange', 'tblPrChange',
                              'trPrChange', 'tcPrChange', 'sectPrChange'}


class RevisionChoiceRequired(ValueError):
    """The document contains changes, but the user has not selected a view."""

    def __init__(self, counts):
        self.counts = dict(counts)
        super().__init__('Есть непринятые исправления Word; уточните редакцию у пользователя. '
                         'После выбора: --revisions current|original|marked. '
                         'Представление не подтверждает согласование документа.')


def _render(node, view):
    tag = node.tag.removeprefix(W)
    if tag in IGNORED:
        return ''
    if tag in ('t', 'delText'):
        return node.text or ''
    if tag == 'tab':
        return '\t'
    if tag in ('br', 'cr'):
        return '\n'
    if node.tag == MC + 'AlternateContent':
        branch = node.find(MC + 'Choice')
        if branch is None:
            branch = node.find(MC + 'Fallback')
        return _render(branch, view) if branch is not None else ''
    if tag in CHANGES:
        insertion = tag in ('ins', 'moveTo')
        if (view == 'current' and not insertion) or (view == 'original' and insertion):
            return ''
        content = ''.join(_render(child, view) for child in node)
        if view == 'marked':
            ending = '\n' if content.endswith('\n') else ' '
            return ' [' + CHANGES[tag] + ': ' + content.rstrip('\n') + ']' + ending
        return content
    content = ''.join(_render(child, view) for child in node)
    if tag == 'p':
        return content + '\n'
    if tag == 'tc':
        return content.rstrip('\n') + '\t'
    if tag == 'tr':
        return content.rstrip('\t') + '\n'
    if tag == 'tbl':
        return '\n[ТАБЛИЦА]\n' + content + '[КОНЕЦ ТАБЛИЦЫ]\n'
    return content


def read_docx(path, revisions=None, include_auxiliary=False):
    """Return (text, metadata) without changing the DOCX package.

    Metadata includes revision counts, part names, and the selected reading view.
    Auxiliary parts are headers, footers and foot/endnotes, used by corpus search.
    """
    if revisions is not None and revisions not in VIEWS:
        raise ValueError('Неизвестный режим исправлений: ' + str(revisions))
    texts, counts, parts, roots = [], Counter(), [], []
    with zipfile.ZipFile(path) as archive:
        names = ['word/document.xml']
        names += sorted(name for name in archive.namelist()
                        if re.fullmatch(r'word/(header\d*|footer\d*|footnotes|endnotes)\.xml', name))
        for name in names:
            root = ET.fromstring(archive.read(name))
            for node in root.iter():
                tag = node.tag.removeprefix(W)
                if tag in REVISION_TAGS:
                    counts[tag] += 1
            if name == 'word/document.xml' or include_auxiliary:
                roots.append((name, root))
                parts.append(name)
    if counts and revisions is None:
        raise RevisionChoiceRequired(counts)
    for name, root in roots:
        content = _render(root, revisions or 'marked')
        if name != 'word/document.xml':
            content = '\n[ЧАСТЬ: ' + name + ']\n' + content
        texts.append(content)
    return '\n'.join(texts), {'revisions': dict(counts), 'view': revisions, 'parts': parts}
