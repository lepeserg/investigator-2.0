"""Read DOCX text in XML order with explicit tracked-change views.

This module reads text only; it does not accept/reject changes or interpret images.
Without an explicit view, tracked changes stop reading until the user chooses.
Views are reading aids and do not establish an approved document version.
"""
from collections import Counter
import functools
import html
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


# Быстрое извлечение текста из сырого WordprocessingML (регулярками, без разбора дерева) —
# общее для check_tom, consistency_check, docx_integrity и style_lint. Единые правила:
#   · текст берётся только из <w:t> (и <w:delText>, если include_deleted); <w:tab>, <w:tabs>,
#     <w:t/> и прочие теги на «w:t» не цепляются; разметка рисунков и чисел вне w:t отбрасывается;
#   · коды полей <w:instrText> не входят никогда;
#   · <w:tab/> → tab ('\t'), табуляторы абзаца <w:tabs>…</w:tabs> — не текст;
#   · </w:p> и пустой <w:p/> — конец абзаца ('\n' либо отдельный элемент списка);
#   · сущности XML декодируются (html.unescape);
#   · префикс берётся из объявления пространства имён WordprocessingML (xmlns:<префикс>=...):
#     генераторы на ElementTree пишут «ns0:t» вместо «w:t». Нет объявления (кусок вроде
#     одного <w:p>) — префикс «w». Чужие пространства (a:t рисунков и т.п.) не читаются.
# Исправления Word здесь не выбираются: вставки читаются как текст, удалённое — по флагу.
# Полноценное чтение с выбором редакции — read_docx.
_OPEN = r"(?:\s[^>]*?)?(?<!/)>"
_W_NS_RX = re.compile(r"""xmlns:([A-Za-z_][\w.-]*)\s*=\s*["'](?:"""
                      r"http://schemas\.openxmlformats\.org/wordprocessingml/2006/main"
                      r"|http://purl\.oclc\.org/ooxml/wordprocessingml/main)[\"']")
DOCX_TEXT_PARTS = r"document|header\d*|footer\d*"


@functools.lru_cache(maxsize=None)
def _xml_text_rx(prefixes):
    w = "(?:" + "|".join(re.escape(p) for p in prefixes) + "):"
    return re.compile(
        r"(?P<tabs><" + w + r"tabs" + _OPEN + r".*?</" + w + r"tabs>)"
        r"|<(?P<pfx>" + w + r")(?P<kind>t|delText|instrText)" + _OPEN
        + r"(?P<body>.*?)</(?P=pfx)(?P=kind)>"
        r"|(?P<tab><" + w + r"tab(?:\s[^>]*)?/>)"
        r"|</" + w + r"p>|<" + w + r"p(?:\s[^>]*)?/>", re.S)


def xml_text(xml, *, include_deleted=False, paragraphs=False, tab="\t"):
    """Текст куска WordprocessingML (document.xml, колонтитул, отдельный <w:p>).

    paragraphs=False → строка, абзацы завершаются '\n';
    paragraphs=True  → список текстов абзацев (хвост вне абзаца — только если непуст).
    include_deleted=True — оставить удалённый при рецензировании текст (w:delText)."""
    out, paras = [], []
    prefixes = tuple(sorted(set(_W_NS_RX.findall(xml)))) or ("w",)
    for m in _xml_text_rx(prefixes).finditer(xml):
        kind = m.group("kind")
        if kind:
            if kind == "t" or (kind == "delText" and include_deleted):
                out.append(html.unescape(m.group("body")))
        elif m.group("tabs"):
            continue
        elif m.group("tab"):
            out.append(tab)
        elif paragraphs:
            paras.append("".join(out))
            out = []
        else:
            out.append("\n")
    tail = "".join(out)
    if not paragraphs:
        return tail
    if tail:
        paras.append(tail)
    return paras


def docx_xml_text(path, parts=DOCX_TEXT_PARTS, *, sep="\n", errors="replace", **kwargs):
    """xml_text частей word/<parts>.xml (порядок — как в архиве; parts — регулярка имени
    без пути и расширения). Строки частей склеиваются sep, при paragraphs=True списки
    абзацев сцепляются. errors — режим decode("utf-8", errors): consistency_check передаёт
    "ignore" (как было у него), чтобы U+FFFD не разрывал ФИО. Пакет только читается."""
    with zipfile.ZipFile(path) as archive:
        texts = [xml_text(archive.read(name).decode("utf-8", errors), **kwargs)
                 for name in archive.namelist()
                 if re.fullmatch(r"word/(?:" + parts + r")\.xml", name)]
    if kwargs.get("paragraphs"):
        return [para for part in texts for para in part]
    return sep.join(texts)


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
