# -*- coding: utf-8 -*-
"""docx_edit.py — БЫСТРАЯ правка существующего .docx «на месте».

Зачем: после live-допроса (и вообще при точечных правках заполненного бланка)
НЕЛЬЗЯ пересобирать документ из карточки — затрёшь ручной труд. Этот модуль
правит ТОТ ЖЕ файл, сохраняя форматирование, колонтитулы и разбивку по страницам.

КОГДА ПРАВИТЬ (этот модуль), а когда РЕГЕНЕРИРОВАТЬ (make_docx):
  ПРАВИТЬ  — .docx уже существует и, вероятно, несёт ручные правки, а изменение локальное:
             поправить ФИО/дату/номер, исправить опечатку, заменить/вставить/дописать абзац,
             поправить ячейку статкарточки.
  РЕГЕНЕРИРОВАТЬ — первая сборка (правок ещё нет); меняется тип/структура документа целиком;
             неверный бланк/шаблон; файл не открывается даже через docx_recover.
  Эмпирика: «поменяй X на Y / добавь предложение» -> ПРАВИТЬ; «пересобери / не тот шаблон» ->
  РЕГЕНЕРИРОВАТЬ. При сомнении и возможных ручных правках — ПРАВИТЬ, не затирать файл молча.

Ключевое отличие от наивной замены: поиск НЕ ЗАВИСИТ от того, как Word разбил
строку на runs и какие пробелы вставил. Любая последовательность пробелов в
искомой строке матчится с любой последовательностью пробелов (включая
неразрывные \\u00A0 и мягкие переносы) в документе. Замена точечная — меняется
только найденный фрагмент, остальной текст абзаца и его формат не трогаются.

ВНЕ ОБЛАСТИ: SmartArt, диаграммы и WordArt — их текст лежит НЕ в `word/document.xml`
(`word/diagrams/*.xml`, `word/charts/*.xml`, атрибут `v:textpath/@string`), и обход туда не
заходит: править только вручную. Также вне области — фраза, разорванная Word'ом между РАЗНЫМИ
абзацами (внутри абзаца разбивка по runs терпится). Legacy бинарный .doc сначала
конвертировать в .docx (см. docx_recover.py / Word COM / LibreOffice).

API:
    from docx_edit import edit_file, find_text, count_text, glue_signature, DocxLockedError, open_doc
    rep = edit_file(path, [
        {"op": "replace", "old": "Флотская, д. 20", "new": "Октябрьская, д. 8", "expect": 1},
        {"op": "replace", "old": "ЦСМЧ", "new": "ЦМСЧ"},
        # ⚠ у replace_paragraph ключи ДРУГИЕ: "anchor" (а не "old") и "new" — по аналогии с replace
        # естественно написать "old", и раньше это падало «нет ключа 'anchor'» (п. 5.3 отчёта 05.09.2026)
        {"op": "replace_paragraph", "anchor": "Приложение: на 3 л.", "new": "Приложение: на 5 л."},
        # список -> КАЖДАЯ строка отдельным абзацем (формат наследуется от якоря); строка -> один абзац
        {"op": "insert_after", "anchor": "Прошу направить", "text": ["– справку об отборе;",
                                                                    "– копию приказа."]},
        {"op": "append", "text": "Дополнительный абзац.", "before_signature": True},
    ])
    # rep -> [{"op":"replace","matched":True/False,"old":...,"count":N,"replaced":0/1}, ...]
    # count — сколько ВХОЖДЕНИЙ было в документе, replaced — сколько заменено (replace правит ПЕРВОЕ).
    # "expect": N — ожидаемое число вхождений; не совпало -> ValueError ДО записи (файл цел).
    # "expect_after": N — то же, но с учётом ПРЕДЫДУЩИХ операций набора (после вставки блока).
    #
    # РАЗМЕТКА И СТРУКТУРА (полный список ключей — в docstring edit_file):
    #   clone_block        — продублировать диапазон абзацев XML-клоном (жир и пустые абзацы целы);
    #   keep_with_next     — флаг keepNext на абзацах диапазона;
    #   insert_blank_after — N пустых абзацев после якоря;
    #   bold_paragraph     — весь абзац жирным; bold_lead — жирным начало до разделителя.
    # ⛔ Всё это раньше делали python-docx'ом напрямую — мимо lock-check, .bak и атомарной записи.

CLI (быстро, без написания питона):
    python docx_edit.py "файл.docx" "старое=>новое" "ещё старое=>ещё новое" ...
    python docx_edit.py "файл.docx" --paras              # СТРУКТУРА абзацев (Шаг 0 сценария Г): № · где · длина · текст
    python docx_edit.py "файл.docx" --check "фрагмент"   # есть ли фрагмент и СКОЛЬКО вхождений
    python docx_edit.py "файл.docx" --dates              # ВСЕ даты: тело + таблицы + надписи + колонтитулы
    python docx_edit.py "файл.docx" --page-numbers       # нумерация страниц, если их больше двух
    python docx_edit.py "файл.docx" --page-numbers --force   # поставить, не считая страниц
    python docx_edit.py "файл.docx" --glue-signature     # скрепить подписной блок keepNext (§16.17)

Перед записью делает «.bak_<timestamp>» рядом (с ретенцией prune_backups) и пишет
АТОМАРНО (tmp в той же папке -> проверка zip -> os.replace). Если файл открыт в
Word — бросает DocxLockedError и НИЧЕГО не пишет (попроси закрыть и повтори)."""
import os, sys, re, time, zipfile, shutil, copy
os.environ.setdefault("PYTHONUTF8", "1")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from make_docx import (  # noqa: E402
    replace_block_between_anchors, set_cell_text, _set_run_font,
    FIRST_LINE_INDENT, BODY_ALIGN, prune_backups,
)
# Занятость файла и атомарное сохранение — общие для навыка (_common). Старые имена
# DocxLockedError / word_lock_path / _check_locked / _save_atomic остаются в docx_edit:
# их импортируют снаружи (DocxLockedError) и тесты (_save_atomic).
import _common  # noqa: E402
from _common import (  # noqa: E402
    DocxLockedError, check_locked as _check_locked, save_atomic as _common_save_atomic,
)
word_lock_path = _common.word_lock_path  # публичное имя (references/25-assembly.md)


def _copy_run_font(dst_run, src_run):
    """Скопировать шрифт (rFonts + размер) из src_run в dst_run — для наследования шрифта формы."""
    s = src_run._r.find(qn("w:rPr"))
    if s is None:
        return
    d = dst_run._r.get_or_add_rPr()
    for tag in ("w:rFonts", "w:sz", "w:szCs"):
        el = s.find(qn(tag))
        if el is not None:
            ex = d.find(qn(tag))
            if ex is not None:
                d.remove(ex)
            d.append(copy.deepcopy(el))


def set_statcard_cell(cell, value, font_from_cell=None):
    'Вспомогательная функция. Частный пример исключён из публичной версии.'
    p = cell.paragraphs[-1]
    if "".join(r.text for r in p.runs).strip() == str(value).strip():
        return cell  # идемпотентность
    nonempty = [r for r in p.runs if r.text.strip()]
    if nonempty:
        nonempty[0].text = str(value)
        for r in nonempty[1:]:
            r.text = ""
        return cell
    if p.runs:
        # run(ы) в строке-значении ЕСТЬ, но пустые (форматированный плейсхолдер бланка, напр. Arial 14пт).
        # Пишем ВНУТРЬ существующего run — наследуем ЕГО шрифт (эталон make_statcards._set_cell_value);
        # НЕ создаём новый run и НЕ копируем из метки-6пт (иначе значение выйдет 6-м кеглем — «съехавший шрифт»).
        target = p.runs[len(p.runs) // 2]
        target.text = str(value)
        # ⚠ Если пустой run БЕЗ собственного w:rPr (типовой случай очищенной графы бланка), наследовать
        # нечего — значение уйдёт шрифтом документа (TNR), а соседние графы Arial 8. Тогда всё же
        # берём шрифт из font_from_cell. Аудит 30.07.2026: font_from молча игнорировался.
        if target._r.find(qn("w:rPr")) is None and font_from_cell is not None:
            donor_run = None
            for pp in reversed(font_from_cell.paragraphs):
                for r in pp.runs:
                    if r.text.strip() and r._r.find(qn("w:rPr")) is not None:
                        donor_run = r
                        break
                if donor_run:
                    break
            if donor_run is not None:
                _copy_run_font(target, donor_run)
        return cell
    # run'ов в строке-значении НЕТ вовсе — создаём с УНАСЛЕДОВАННЫМ шрифтом (не TNR-дефолт!).
    # Донор шрифта — run из строки-ЗНАЧЕНИЯ соседней графы (font_from) или своей ячейки: абзацы С КОНЦА.
    donor_run = None
    for candidate in ([font_from_cell] if font_from_cell is not None else []) + [cell]:
        for pp in reversed(candidate.paragraphs):
            for r in pp.runs:
                if r.text.strip():
                    donor_run = r
                    break
            if donor_run:
                break
        if donor_run:
            break
    nr = p.add_run(str(value))
    if donor_run is not None:
        _copy_run_font(nr, donor_run)
    return cell


_WS = re.compile(r"\s+")


# ---------- открытие / сохранение ----------

def open_doc(path):
    return Document(path)


def _save_atomic(doc, path, *, backup=True):
    """Атомарно сохранить в тот же файл: tmp ~edit_* в ТОЙ ЖЕ папке -> проверка zip ->
    метаданные + бэкап -> os.replace (_common.save_atomic; tmp удаляется при любом сбое).
    backup=True: копия исходника в <path>.bak_<timestamp> перед заменой (+ prune_backups)."""
    return _common_save_atomic(doc, path, tmp_prefix="~edit_",
                               before_replace=lambda tmp: _finish_save(tmp, path, backup))


def _finish_save(tmp, path, backup):
    """Метаданные -> бэкап. Замену tmp->path и очистку tmp делает _common.save_atomic."""
    # Свойства документа → профиль владельца (автор = следователь; никаких следов
    # python-docx и чужих авторов донора — стоячее указание владельца 07.08.2026,
    # см. doc_meta.py). Идемпотентно: на чистом файле no-op.
    try:
        import doc_meta
        doc_meta.clean(tmp)
    except Exception:
        pass  # метаданные не должны ронять сохранение содержимого
    if backup and os.path.exists(path):
        stamp = time.strftime("%Y%m%d_%H%M%S")
        bak = path + ".bak_" + stamp
        i = 1
        while os.path.exists(bak):
            bak = "%s.bak_%s_%d" % (path, stamp, i); i += 1
        shutil.copyfile(path, bak)
        prune_backups(path)  # ретенция: не плодить .bak_* без ограничения


# ---------- обход абзацев ----------

def _join_where(*parts):
    """Склеить подсказки о контейнере: «колонтитул-верх › таблица 0, стр. 1, кол. 3»."""
    return " › ".join(p for p in parts if p)


def _walk_paragraphs(container, where=""):
    """Рекурсивно: абзацы контейнера + всех таблиц ЛЮБОЙ вложенности + НАДПИСЕЙ (текстбоксов).

    Отдаёт ПАРЫ (абзац, подсказка о контейнере): «» — сам контейнер, «таблица 0, стр. 1, кол. 3»
    — ячейка таблицы, «надпись» — текстбокс; вложенность склеивается через « › »."""
    for p in container.paragraphs:
        yield p, where
    for ti, t in enumerate(container.tables):
        for ri, row in enumerate(t.rows):
            for ci, cell in enumerate(row.cells):
                yield from _walk_paragraphs(
                    cell, _join_where(where, "таблица %d, стр. %d, кол. %d" % (ti, ri + 1, ci + 1)))
    for p in _textbox_paragraphs(container):
        yield p, _join_where(where, "надпись")


def _textbox_paragraphs(container):
    """Абзацы внутри НАДПИСЕЙ (`w:txbxContent`) — рамки бланков, где живут дата и исх. №.

    Зачем: в бланке 384 ВСО дата и исходящий номер стоят в надписи, а не в теле. Раньше
    docx_edit туда не заходил, и они МОЛЧА оставались донорскими (боевой отчёт 29.07.2026).
    Обходим и обычные (`w:txbxContent`), и AlternateContent-варианты Word."""
    el = getattr(container, "_element", None)
    if el is None:
        return
    try:
        boxes = el.findall(".//" + qn("w:txbxContent"))
    except Exception:
        return
    for box in boxes:
        for p_el in box.findall(qn("w:p")):
            yield Paragraph(p_el, container)


_HDR_FTR_ATTRS = ("header", "footer", "first_page_header", "first_page_footer",
                  "even_page_header", "even_page_footer")


def _hdrftr_exists(part):
    """Есть ли у части СВОЁ определение в этой секции — БЕЗ его создания.

    ⚠ КРИТИЧНО. В python-docx обращение к `header.paragraphs` вызывает
    `_get_or_add_definition()` и МАТЕРИАЛИЗУЕТ пустой колонтитул: в документ добавляются
    word/header1..3.xml + footer1..3.xml и ссылки в sectPr. На бланке с полем 1 см это
    сдвигает форму — именно так «разъехались» статкарты (боевой отчёт 29.07.2026; регресс
    внесён моей же правкой v17.4, где обход колонтитулов делался через getattr(sec, attr)).
    `is_linked_to_previous` читает ссылку в sectPr и ничего не создаёт — проверяем ТОЛЬКО им."""
    try:
        return not part.is_linked_to_previous
    except Exception:
        return False


def _all_paragraphs_located(doc):
    '[(абзац, где)] по ВСЕМ контейнерам документа.'
    # ⛔ ДЕДУПЛИКАЦИЯ ПО XML-ЭЛЕМЕНТУ — обязательна.
    # `_walk_paragraphs` спускается в ячейки таблиц (там вложенный вызов уже отдаёт надписи ячейки),
    # а затем `_textbox_paragraphs` ищет `.//w:txbxContent` по ВСЕМУ поддереву — включая те же
    # таблицы. Абзац надписи ВНУТРИ таблицы попадал в список ДВАЖДЫ, `replace_all` прогонялся по
    # нему два раза, и защита сентинелом (она конечна внутри ОДНОГО вызова) обходилась:
    # самоссылочная замена вкладывалась сама в себя — «исх. исх. СК РОССИИ» на реальном бланке ВСО.
    # Порядок первого вхождения СОХРАНЯЕМ: `replace` правит первое вхождение — тело, а не шапку.
    seen_el, out = set(), []
    for p, w in _walk_paragraphs(doc):
        key = id(p._element)
        if key in seen_el:
            continue
        seen_el.add(key)
        # «тело» — только абзац САМОГО тела; ячейка таблицы и надпись несут свою метку
        out.append((p, w or "тело"))
    seen = set()
    for si, sec in enumerate(doc.sections):
        for attr in _HDR_FTR_ATTRS:
            part = getattr(sec, attr, None)
            if part is None:
                continue
            if not _hdrftr_exists(part):
                continue  # ⛔ НЕ трогать: обращение к .paragraphs создало бы пустой колонтитул
            # у связанных секций (is_linked_to_previous) объект тот же — не дублировать
            key = id(getattr(part, "_element", part))
            if key in seen:
                continue
            seen.add(key)
            where = ("колонтитул-верх" if "header" in attr else "колонтитул-низ")
            if len(doc.sections) > 1:
                where += f" (секция {si + 1})"
            try:
                out.extend((p, _join_where(where, w)) for p, w in _walk_paragraphs(part))
            except Exception:
                pass  # экзотические части — не роняем правку
    return out


def _all_paragraphs(doc):
    return [p for p, _ in _all_paragraphs_located(doc)]


# ---------- поиск / замена, «терпящая разбивку Word» ----------

def _pattern(old):
    """Шаблон: токены искомой строки, склеенные «любым пробелом» (\\s+).
    Так строка находится независимо от переносов и неразрывных пробелов."""
    tokens = _WS.split(old.strip())
    return re.compile(r"\s*".join(re.escape(t) for t in tokens))


def _para_full(para):
    return "".join(r.text for r in para.runs)


def _norm(s):
    """Нормализовать пробелы: любая их последовательность -> один, края обрезаны."""
    return _WS.sub(" ", s or "").strip()


def _para_text(para):
    'Полный текст абзаца для ПОИСКА: runs, а если их нет — `p.text` (гиперссылки, поля).'
    return _para_full(para) or para.text or ""


def _para_matches(para, needle, contains=True):
    """ЕДИНАЯ проверка «абзац подходит под якорь/фрагмент» — та же нормализация, что у count_text.

    contains=True  — фрагмент внутри абзаца (пробелы любые, в т.ч. неразрывные и переносы);
    contains=False — точное равенство ПОСЛЕ нормализации пробелов."""
    text = _para_text(para)
    if not contains:
        return _norm(text) == _norm(needle)
    n = _norm(needle)
    if not n:
        return False
    return _pattern(n).search(text) is not None


def _hits(paras, needle, contains=True):
    """Номера абзацев, подходящих под якорь (единая нормализация `_para_matches`)."""
    n = _norm(needle)
    if not contains:
        return [i for i, p in enumerate(paras) if _norm(_para_text(p)) == n]
    if not n:
        return []
    pat = _pattern(n)          # компилируем ОДИН раз на весь список абзацев
    return [i for i, p in enumerate(paras) if pat.search(_para_text(p))]


def _nearest(paras, needle, k=3, cutoff=0.45):
    """[(№, текст, доля совпадения)] — 2–3 ближайших абзаца по нечёткому совпадению.

    Зачем: раньше при ненайденном якоре инструмент говорил только «не найдено», и якорь
    подбирали вслепую (провал 05.09.2026). Мера — длина наибольшего общего фрагмента,
    делённая на длину якоря: она честно показывает «якорь почти такой, отличается хвостом»."""
    import difflib
    n = _norm(needle).lower()
    if not n:
        return []
    scored = []
    for i, p in enumerate(paras):
        t = _norm(_para_text(p))
        if not t:
            continue
        m = difflib.SequenceMatcher(None, n, t.lower()).find_longest_match(0, len(n), 0, len(t))
        score = m.size / float(len(n))
        if score >= cutoff:
            scored.append((score, i, t))
    scored.sort(key=lambda x: (-x[0], x[1]))
    return [(i, t, s) for s, i, t in scored[:k]]


def _hint_not_found(paras, needle, k=3):
    """Хвост сообщения об ошибке: ближайшие абзацы (или совет напечатать структуру)."""
    near = _nearest(paras, needle, k=k)
    if not near:
        return " — похожих абзацев не нашлось; напечатай структуру: `--paras`"
    return (" — ближайшие абзацы: "
            + "; ".join("[%d] «%s» (%d%%)" % (i, t[:70], round(s * 100)) for i, t, s in near))


def find_text(doc, sub):
    """True, если фрагмент есть в документе (с учётом разбивки по runs/пробелам)."""
    pat = _pattern(sub)
    return any(pat.search(_para_text(p)) for p in _all_paragraphs(doc))


def count_text(doc, sub):
    """СКОЛЬКО раз фрагмент встречается в документе (тело + таблицы + надписи + колонтитулы).

    Зачем: `replace` правит ТОЛЬКО ПЕРВОЕ вхождение. В документе с повторяющимися блоками
    (ОЗ по совокупности, обложки томов, пачка однотипных уведомлений) «правка легла» ещё не
    значит «правка легла везде» — реальный провал 22.08.2026: место нахождения свидетеля
    легло в один перечень из трёх. Считать ожидаемое число вхождений ДО правки и сверять
    (`{"op": ..., "expect": N}` в edit_file)."""
    doc = doc if hasattr(doc, "sections") else Document(doc)
    pat = _pattern(sub)
    return sum(len(pat.findall(_para_text(p))) for p in _all_paragraphs(doc))


def list_paragraphs(path, width=60):
    """Структура абзацев: [(№, где, длина, первые `width` знаков)] — Шаг 0 сценария Г (§25.04).

    Зачем: якорь подбирают наугад, потому что не смотрят на документ. А многострочный
    заголовок бланка — это ОДИН абзац с мягкими переносами (в тексте они не видны, зато
    видна длина), поэтому якорь по второй строке не находится никогда. Печатать структуру
    ДО первой правки, а не после отказа:  python docx_edit.py "файл.docx" --paras"""
    doc = path if hasattr(path, "sections") else Document(path)
    out = []
    for i, (p, where) in enumerate(_all_paragraphs_located(doc)):
        t = _WS.sub(" ", _para_full(p) or p.text).strip()
        out.append((i, where, len(t), t[:width]))
    return out


def _replace_span(runs, s, e, new):
    """Заменить глобальный диапазон [s, e) текста абзаца на new, правя runs."""
    pos = 0
    inserted = False
    for r in runs:
        t = r.text
        rs, re_ = pos, pos + len(t)
        kept = "".join(ch for i, ch in enumerate(t) if not (s <= rs + i < e))
        if not inserted and rs <= s < re_ + (1 if e > rs else 0):
            # позиция вставки = число символов этого run, стоящих ДО s
            before = sum(1 for i in range(len(t)) if rs + i < s)
            kept = kept[:before] + new + kept[before:]
            inserted = True
        r.text = kept
        pos = re_
    if not inserted and runs:  # s в самом конце абзаца
        runs[-1].text += new


def _replace_one(para, pat, new):
    full = _para_full(para)
    m = pat.search(full)
    if not m:
        return False
    _replace_span(para.runs, m.start(), m.end(), new)
    return True


_SENTINEL = ""          # PUA-символ: в процессуальных документах не встречается
_RUNAWAY_LIMIT = 500


def _replace_all_in_para(para, pat, new):
    'Заменить ВСЕ вхождения в абзаце. Через сентинел — защита от САМОССЫЛОЧНОЙ замены.'
    n = 0
    while _replace_one(para, pat, _SENTINEL):
        n += 1
        if n > _RUNAWAY_LIMIT:
            raise ValueError(
                "replace_all: превышен предохранитель в %d замен в одном абзаце — похоже на "
                "зацикливание. Правка НЕ внесена." % _RUNAWAY_LIMIT)
    if n:
        spat = re.compile(re.escape(_SENTINEL))
        left = n
        while _replace_one(para, spat, new) and left >= 0:
            left -= 1
    return n


# ---------- операции над абзацами (порт из v15) ----------

def _iter_paragraphs(doc):
    yield from _all_paragraphs(doc)


def _bold_majority(para):
    'Жирным ли БОЛЬШИНСТВО СИМВОЛОВ абзаца (не большинство runs).'
    runs = [r for r in para.runs if r.text.strip()]
    if not runs:
        return False
    total = sum(len(r.text) for r in runs)
    bold = sum(len(r.text) for r in runs if r.bold)
    return total > 0 and bold * 2 > total


def replace_paragraph(doc, anchor, new_text, *, contains=True, lead=None, allow_multiple=False):
    """Полностью переписать абзац, найденный по anchor (подстрока, либо точное равенство).

    ⛔ ЯКОРЬ ДОЛЖЕН БЫТЬ УНИКАЛЕН. Если он встречается в нескольких абзацах, вызов
    отклоняется (ValueError) со списком всех попаданий — раньше правка молча уходила в ПЕРВОЕ
    вхождение. Реальный провал 19.08.2026: якорь «384 военного следственного отдела» был взят
    в расчёте на подписной блок, но раньше встретился в абзаце о порядке обжалования — и абзац
    об обжаловании был УНИЧТОЖЕН в двух уведомлениях, замечено только на читке.
    allow_multiple=True — осознанно править первое вхождение.

    ⚠ ШРИФТ НАСЛЕДУЕТСЯ ОТ ПРАВИМОГО АБЗАЦА, а не задаётся TNR 13 (боевой отчёт 29.07.2026:
    константа роняла межстрочный 1,0 в протоколе и 8 пт в статкарте). `_set_run_font` —
    только если в абзаце вообще не было ни одного run со своим форматом."""
    # ⚠ Совпадение — через `_para_matches` (та же нормализация, что у count_text/--check):
    # иначе якорь с неразрывным пробелом --check находит, а правка «не видит».
    hits = [p for p in _iter_paragraphs(doc) if _para_matches(p, anchor, contains)]
    if len(hits) > 1 and not allow_multiple:
        lst = "\n".join("   %d) %s" % (i + 1, (p.text.strip()[:120] or "«пустой абзац»"))
                        for i, p in enumerate(hits))
        raise ValueError(
            "replace_paragraph: якорь %r встречается в %d абзацах — правка НЕ внесена.\n%s\n"
            "Взять более длинный и уникальный якорь (правка уходит в ПЕРВОЕ вхождение и молча "
            "затирает не тот абзац) либо передать allow_multiple=True осознанно."
            % (anchor, len(hits), lst))
    for p in hits:
        majority_bold = _bold_majority(p)
        # запомнить формат ДО очистки: первый содержательный run абзаца — эталон шрифта
        donor = next((r for r in p.runs if r.text.strip()), None) or next((r for r in p.runs), None)
        src_rPr = copy.deepcopy(donor._r.find(qn("w:rPr"))) if donor is not None else None

        def _fmt(run, bold):
            if src_rPr is not None:
                old = run._r.find(qn("w:rPr"))
                if old is not None:
                    run._r.remove(old)
                run._r.insert(0, copy.deepcopy(src_rPr))
            else:
                _set_run_font(run)
            run.bold = bold

        for r in list(p.runs):
            r.text = ""
        r0 = p.runs[0] if p.runs else p.add_run("")
        if lead:
            rest = new_text[len(lead):] if new_text.startswith(lead) else new_text
            r0.text = lead
            _fmt(r0, True)
            r1 = p.add_run(rest)
            _fmt(r1, False)
        else:
            r0.text = new_text
            _fmt(r0, majority_bold)
        return True
    raise ValueError("Абзац-якорь не найден: «%s»%s"
                     % (anchor, _hint_not_found(list(_iter_paragraphs(doc)), anchor)))


def _clone_format(dst_para, src_para):
    """Скопировать формат АБЗАЦА-ОБРАЗЦА (pPr) и шрифт его первого содержательного run.

    Зачем: `_set_run_font` + константы FIRST_LINE_INDENT/BODY_ALIGN навязывают TNR 13 / отступ
    1,25 / по ширине. В протоколе это ломает межстрочный 1,0 и интервал 0, а в статкарте нужен
    вообще 8 пт с наследованием Arial (боевой отчёт 29.07.2026). Правильно — наследовать от
    соседнего абзаца того же документа."""
    if src_para is None:
        return False
    src_pPr = src_para._p.find(qn("w:pPr"))
    if src_pPr is not None:
        old = dst_para._p.find(qn("w:pPr"))
        if old is not None:
            dst_para._p.remove(old)
        dst_para._p.insert(0, copy.deepcopy(src_pPr))
    donor = next((r for r in src_para.runs if r.text.strip()), None)
    if donor is None:
        donor = next((r for r in src_para.runs), None)
    if donor is not None:
        for r in dst_para.runs:
            src_rPr = donor._r.find(qn("w:rPr"))
            if src_rPr is not None:
                old = r._r.find(qn("w:rPr"))
                if old is not None:
                    r._r.remove(old)
                r._r.insert(0, copy.deepcopy(src_rPr))
    return src_pPr is not None or donor is not None


def _new_para(doc, text, *, like=None):
    """Новый абзац. `like` — абзац-ОБРАЗЕЦ, от которого наследуется формат (предпочтительно).
    Без образца — прежние умолчания формата ВСО (TNR 13, отступ 1,25, по ширине).

    ⛔ Только СТРОКА. Список сюда уходил в `add_run(list)` и склеивался в один абзац молча
    (п. 5.2 отчёта 05.09.2026): python-docx перебирает переданное «посимвольно», а у списка
    «символы» — целые строки. Многоабзацная вставка — через `insert_paragraph` (он принимает
    список) или `replace_block`."""
    if not isinstance(text, str):
        raise ValueError("_new_para: ожидается СТРОКА, получено %s — список строк склеился бы в "
                         "один абзац. Для перечня используй insert_paragraph(..., [строки]) или "
                         "replace_block." % type(text).__name__)
    np = doc.add_paragraph()
    run = np.add_run(text)
    if like is not None and _clone_format(np, like):
        return np
    _set_run_font(run)
    np.paragraph_format.first_line_indent = FIRST_LINE_INDENT
    np.alignment = BODY_ALIGN
    return np


def delete_paragraph(doc, anchor, *, contains=True, require=True, max_count=1):
    """Удалить абзац(ы) по якорю. Возвращает число удалённых.

    Зачем: до 05.08.2026 операции удаления НЕ БЫЛО (были replace/insert/block), поэтому
    подмена заполнителя вроде пустого «Ответ:» делалась вставкой + ручной чисткой через
    python-docx МИМО всех защит docx_edit (lock-check, .bak, атомарная запись). Реальный
    след: после insert_after остался дубль «Ответ:», который вычищали руками.

    max_count ограничивает число удалений (по умолчанию 1) — защита от сноса половины
    документа при слишком общем якоре."""
    n = 0
    all_paras = list(_iter_paragraphs(doc))
    for p in all_paras:
        if not _para_matches(p, anchor, contains):
            continue
        el = p._element
        el.getparent().remove(el)
        n += 1
        if n >= max_count:
            break
    if n == 0 and require:
        raise ValueError("delete_paragraph: абзац-якорь НЕ найден — правка не внесена: "
                         + repr(anchor) + _hint_not_found(all_paras, anchor))
    return n


def replace_placeholder(doc, anchor, text, *, contains=True, require=True):
    """Заменить абзац-ЗАПОЛНИТЕЛЬ на содержательный, сохранив его формат.

    Разница с `replace_paragraph`: тот переписывает текст в существующих runs; этот нужен,
    когда заполнитель пустой или почти пустой («Ответ:», «___», «[текст]») — пишем текст
    и наследуем формат самого заполнителя, ничего не задваивая. Именно отсутствие такой
    операции и приводило к паре «вставили новый + остался пустой старый»."""
    return replace_paragraph(doc, anchor, text, contains=contains)


def _as_lines(text):
    """Текст операции -> СПИСОК строк-абзацев. Список/кортеж — по элементу на абзац; строка — один.

    ⛔ Провал 05.09.2026 (перечень приложений к запросу в КВТКУ, п. 5.2 отчёта). Список из трёх
    пунктов уходил в `add_run(list)`: python-docx перебирает переданное ПОСИМВОЛЬНО, а у списка
    «символы» — целые строки, поэтому все три склеивались в ОДИН абзац
    («– справку о результатах отбора;– справку о размере выплат;– копию приказа.») — молча, без
    ошибки. Приходилось делать replace_paragraph на первый пункт и по insert_after на остальные."""
    if isinstance(text, (list, tuple)):
        return [str(x) for x in text]
    return [text if isinstance(text, str) else str(text)]


def insert_paragraph(doc, anchor, text, *, position="after", like_anchor=True):
    """Вставить новый абзац (или НЕСКОЛЬКО) относительно абзаца-якоря (position 'after'/'before').

    text — строка (один абзац) ЛИБО список/кортеж строк: тогда КАЖДАЯ строка становится
    ОТДЕЛЬНЫМ абзацем, в переданном порядке (перечень приложений, пункты списка).
    like_anchor=True (по умолчанию) — формат наследуется от АБЗАЦА-ЯКОРЯ, а не задаётся
    константами: так сохраняются межстрочный интервал, кегль и отступы окружающего текста.

    Возвращает число вставленных абзацев (>= 1)."""
    lines = _as_lines(text)
    if not lines:
        raise ValueError("insert_paragraph: пустой список строк — вставлять нечего")
    all_paras = list(_iter_paragraphs(doc))
    for p in all_paras:
        if _para_matches(p, anchor, True):
            like = p if like_anchor else None
            ref = p._p
            for ln in lines:
                np = _new_para(doc, ln, like=like)
                if position == "before":
                    p._p.addprevious(np._p)      # каждый встаёт ВПРИТЫК перед якорем — порядок цел
                else:
                    ref.addnext(np._p)           # цепочкой после якоря — порядок тоже цел
                    ref = np._p
            return len(lines)
    raise ValueError("Абзац-якорь не найден: «%s»%s"
                     % (anchor, _hint_not_found(all_paras, anchor)))


SIGN_ANCHORS = ("Следователь", "Руководитель", "Заместитель", "Врио", "Дознаватель")


def append_paragraph(doc, text, *, before_signature=False):
    """Добавить абзац. before_signature=True — вставить ПЕРЕД подписным блоком."""
    if before_signature:
        for p in doc.paragraphs:
            if any(a in p.text for a in SIGN_ANCHORS):
                np = _new_para(doc, text)
                p._p.addprevious(np._p)
                return True
    _new_para(doc, text)
    return True


def replace_block(doc, start_anchor, end_anchor, new_lines):
    """Заменить абзацы СТРОГО между якорями (границы целы). Для блока показаний и т.п."""
    return replace_block_between_anchors(doc, start_anchor, end_anchor, new_lines)


# ---------- ДИАПАЗОН АБЗАЦЕВ: клонирование, keepNext, пустые отбивки, жир ----------
#
# Локальный пример исключён из публичной поставки.
#   · раздел ОЗ дублировали ЧЕРЕЗ ТЕКСТ (собрали `p.text` и отдали в `block`) — внутриабзацный
#     жир схлопнулся в один прогон, а фильтр `if x.text.strip()` выбросил ПУСТЫЕ абзацы, то есть
#     всю разметку раздела;
#   · `keepNext` и полужирную разметку выставляли python-docx'ом НАПРЯМУЮ и сохраняли документ
#     сами — мимо lock-check, `.bak`, атомарной записи и постгейта. Дважды за сессию.
# Поэтому: клонирование — копированием XML (`copy.deepcopy` элемента `w:p`), без прохода через
# текст; разметка — штатными операциями набора.


def _resolve_range(paras, *, start=None, end=None, start_index=None, end_index=None, what="range"):
    """Границы диапазона абзацев: по якорям либо по индексам `--paras`. Вернуть (si, ei), si<=ei."""
    if start_index is not None or end_index is not None:
        if start_index is None or end_index is None:
            raise ValueError("%s: индексы задаются ПАРОЙ — и 'start_index', и 'end_index'" % what)
        si, ei = int(start_index), int(end_index)
        if not (0 <= si < len(paras)) or not (0 <= ei < len(paras)):
            raise ValueError("%s: индекс вне диапазона (абзацев %d) — напечатай структуру `--paras`"
                             % (what, len(paras)))
        if ei < si:
            raise ValueError("%s: 'end_index' выше 'start_index' — границы перепутаны местами" % what)
        return si, ei
    if not start or not end:
        raise ValueError("%s: задай границы — пару якорей 'start'/'end' либо пару "
                         "'start_index'/'end_index' (номера из `--paras`)" % what)
    sh = _hits(paras, start, True)
    if not sh:
        raise ValueError("%s: якорь начала %r НЕ найден%s" % (what, start, _hint_not_found(paras, start)))
    si = sh[0]
    eh = [i for i in _hits(paras, end, True) if i >= si]
    if not eh:
        raise ValueError("%s: якорь конца %r не найден ПОСЛЕ начала (абз. %d)%s"
                         % (what, end, si, _hint_not_found(paras, end)))
    return si, eh[0]


def _span_elements(s_el, e_el):
    """Соседние XML-элементы от s_el до e_el включительно (оба должны быть в одном родителе)."""
    parent = s_el.getparent()
    if parent is None or e_el.getparent() is not parent:
        raise ValueError("clone_block: границы диапазона лежат в РАЗНЫХ контейнерах (тело/ячейка/"
                         "надпись) — клонировать такой диапазон нельзя, возьми границы внутри одного")
    kids = list(parent)
    a, b = kids.index(s_el), kids.index(e_el)
    return kids[a:b + 1]


def clone_block(doc, start=None, end=None, *, start_index=None, end_index=None,
                position="after", times=1):
    """Продублировать диапазон абзацев КОПИРОВАНИЕМ XML — жир, пустые абзацы и разметка целы.

    Границы включительно: пара якорей 'start'/'end' либо пара индексов из `--paras`.
    position='after' (по умолчанию) — копия сразу после диапазона; 'before' — перед ним.
    times — сколько копий (для ОЗ по совокупности: доказательственный раздел на каждый состав).

    Возвращает {"count", "cloned", "paras_before", "paras_after", "start_index", "end_index"}."""
    if int(times) < 1:
        raise ValueError("clone_block: times должно быть >= 1")
    paras = _all_paragraphs(doc)
    before = len(paras)
    si, ei = _resolve_range(paras, start=start, end=end, start_index=start_index,
                            end_index=end_index, what="clone_block")
    s_el, e_el = paras[si]._element, paras[ei]._element
    span = _span_elements(s_el, e_el)
    for _ in range(int(times)):
        nodes = [copy.deepcopy(el) for el in span]
        if position == "before":
            for nd in nodes:                 # каждый встаёт ВПРИТЫК перед началом — порядок сохраняется
                s_el.addprevious(nd)
        else:
            ref = e_el
            for nd in nodes:
                ref.addnext(nd)
                ref = nd
    after = len(_all_paragraphs(doc))
    return {"count": int(times), "cloned": len(span), "paras_before": before,
            "paras_after": after, "start_index": si, "end_index": ei}


def keep_with_next(doc, *, anchor=None, start=None, end=None, start_index=None, end_index=None,
                   contains=True, with_last=False, allow_multiple=False):
    """Поставить флаг `keepNext` на абзацы диапазона — чтобы Word не разорвал их по страницам.

    Один абзац — 'anchor' (флаг «не отрывать от следующего» ставится на него).
    Диапазон — пара 'start'/'end' либо пара индексов из `--paras`.
    ⚠ На ПОСЛЕДНЕМ абзаце диапазона флага НЕТ (тянуть его не за что, а лишний `keepNext`
    подтащил бы к блоку посторонний следующий абзац) — как в `glue_signature`. Нужен и на
    последнем (диапазон продолжается дальше) — with_last=True.
    Отличие от `glue_signature`: тот сам ИЩЕТ подписной блок по содержанию, а здесь границы
    задаёт вызывающий — для таблиц-перечней, шапок и заголовков с текстом.

    Возвращает {"glued": сколько абзацев помечено, "start_index", "end_index"}."""
    paras = _all_paragraphs(doc)
    if anchor is not None and start is None and end is None \
            and start_index is None and end_index is None:
        hits = _hits(paras, anchor, contains)
        if not hits:
            raise ValueError("keep_with_next: якорь %r НЕ найден%s"
                             % (anchor, _hint_not_found(paras, anchor)))
        if len(hits) > 1 and not allow_multiple:
            raise ValueError(
                "keep_with_next: якорь %r встречается в %d абзацах — флаг лёг бы на ПЕРВЫЙ. "
                "Взять уникальный якорь, задать диапазон 'start'/'end' либо allow_multiple=True"
                % (anchor, len(hits)))
        idx = [hits[0]]
    else:
        si, ei = _resolve_range(paras, start=start, end=end, start_index=start_index,
                                end_index=end_index, what="keep_with_next")
        idx = list(range(si, ei + 1))
        if not with_last and len(idx) > 1:
            idx = idx[:-1]
    for i in idx:
        paras[i].paragraph_format.keep_with_next = True
    return {"glued": len(idx), "start_index": idx[0], "end_index": idx[-1]}


def insert_blank_after(doc, anchor, count=1, *, contains=True, like_anchor=True):
    """Вставить N ПУСТЫХ абзацев после абзаца-якоря (отбивка разметки ОЗ, отступ перед подписью).

    Формат наследуется от якоря (кегль и интервалы отбивки должны совпадать с окружением).
    Зачем операция: пустые абзацы — часть разметки, а не мусор; раньше их добавляли
    python-docx'ом напрямую, мимо `.bak` и постгейта."""
    n = int(count)
    if n < 1:
        raise ValueError("insert_blank_after: count должно быть >= 1")
    paras = _all_paragraphs(doc)
    hits = _hits(paras, anchor, contains)
    if not hits:
        raise ValueError("insert_blank_after: якорь %r НЕ найден%s"
                         % (anchor, _hint_not_found(paras, anchor)))
    p = paras[hits[0]]
    ref = p._p
    for _ in range(n):
        np = _new_para(doc, "", like=(p if like_anchor else None))
        ref.addnext(np._p)
        ref = np._p
    return {"inserted": n, "index": hits[0]}


def bold_paragraph(doc, anchor, *, bold=True, contains=True, allow_multiple=False):
    """Сделать ВЕСЬ абзац полужирным (или снять жир, bold=False). Заголовки разделов ОЗ.

    ⚠ Правит только начертание: текст, шрифт и абзацный формат не трогаются."""
    paras = _all_paragraphs(doc)
    hits = _hits(paras, anchor, contains)
    if not hits:
        raise ValueError("bold_paragraph: якорь %r НЕ найден%s"
                         % (anchor, _hint_not_found(paras, anchor)))
    if len(hits) > 1 and not allow_multiple:
        raise ValueError(
            "bold_paragraph: якорь %r встречается в %d абзацах — жир лёг бы на ПЕРВЫЙ. "
            "Взять более длинный уникальный якорь либо allow_multiple=True" % (anchor, len(hits)))
    n = 0
    for i in (hits if allow_multiple else hits[:1]):
        p = paras[i]
        if not p.runs:
            raise ValueError("bold_paragraph: в абзаце [%d] нет ни одного run — нечего выделять" % i)
        for r in p.runs:
            r.bold = bool(bold)
        n += 1
    return {"count": n, "indices": (hits if allow_multiple else hits[:1])}


def _split_run_at(para, run, offset):
    """Разрезать run по смещению внутри его текста; вернуть ПРАВУЮ часть (новый run) или None.

    Формат (`w:rPr`) копируется целиком — начертание, шрифт и кегль правой части те же."""
    from docx.text.run import Run
    t = run.text
    if offset <= 0 or offset >= len(t):
        return None
    new_r = copy.deepcopy(run._r)
    run._r.addnext(new_r)
    right = Run(new_r, para)
    right.text = t[offset:]
    run.text = t[:offset]
    return right


def bold_lead(doc, anchor, *, sep=",", include_sep=False, contains=True,
              rest_bold=None, allow_multiple=False, occurrence=1):
    """Выделить полужирным НАЧАЛО абзаца — до заданного разделителя (по умолчанию первой запятой).

    sep — разделитель; occurrence — какое по счёту его вхождение считать границей (для зачинов
    с запятой внутри инициалов); include_sep=True — разделитель тоже жирным.
    rest_bold=None — остаток абзаца не трогать; False — принудительно снять с него жир.

    Возвращает {"index", "lead": выделенный текст, "cut": смещение границы}."""
    paras = _all_paragraphs(doc)
    hits = _hits(paras, anchor, contains)
    if not hits:
        raise ValueError("bold_lead: якорь %r НЕ найден%s" % (anchor, _hint_not_found(paras, anchor)))
    if len(hits) > 1 and not allow_multiple:
        raise ValueError(
            "bold_lead: якорь %r встречается в %d абзацах — разметка легла бы в ПЕРВЫЙ. "
            "Взять более длинный уникальный якорь либо allow_multiple=True" % (anchor, len(hits)))
    p = paras[hits[0]]
    full = _para_full(p)
    if not full:
        raise ValueError("bold_lead: в абзаце [%d] нет текста в runs — размечать нечего" % hits[0])
    pos, k = -1, 0
    while k < int(occurrence):
        pos = full.find(sep, pos + 1)
        if pos < 0:
            raise ValueError(
                "bold_lead: разделитель %r (вхождение %d) в абзаце [%d] не найден — абзац: «%s». "
                "Задай другой sep или occurrence." % (sep, occurrence, hits[0], _norm(full)[:100]))
        k += 1
    cut = pos + (len(sep) if include_sep else 0)
    if cut <= 0:
        raise ValueError("bold_lead: граница пришлась на самое начало абзаца — выделять нечего")
    # разрезать run, на который попала граница, и раскрасить обе половины
    acc = 0
    for r in list(p.runs):
        ln = len(r.text)
        if acc < cut < acc + ln:
            _split_run_at(p, r, cut - acc)
            break
        acc += ln
    acc = 0
    for r in p.runs:
        ln = len(r.text)
        if acc + ln <= cut:
            r.bold = True
        elif rest_bold is not None:
            r.bold = bool(rest_bold)
        acc += ln
    return {"index": hits[0], "lead": full[:cut], "cut": cut}


# ---------- СКЛЕЙКА ПОДПИСНОГО БЛОКА (`keepNext`) ----------
#
# Зачем: любая правка, меняющая ЧИСЛО абзацев (`delete_paragraph`, `block` с иным числом строк),
# сдвигает вёрстку — и подписной блок уезжает к нижней кромке листа. Провал 25.08.2026
# Локальный пример исключён из публичной поставки.
# `подпись-разорвана` (§16.17) появился ТОЛЬКО на ПОВТОРНОМ прогоне `style_lint`, а чинить было
# нечем — готовый хелпер `_glue_block` живёт ВНУТРИ `make_docx.py` и наружу не выведен, поэтому
# `keep_with_next` выставляли вручную через python-docx МИМО защит docx_edit (lock-check, .bak,
# атомарная запись). Здесь та же склейка, но для УЖЕ СУЩЕСТВУЮЩЕГО файла.
#
# ⚠ Поиск блока — ОДИН В ОДИН со `style_lint._sig_block` и `_SIG_LINE`: разойдутся — склейка ляжет
#   не на те абзацы, и флаг останется гореть после починки. Меняешь здесь — сверь там.
# ⛔ Область — ТЕЛО документа (`doc.paragraphs`). Подпись внутри ячейки таблицы (бланки-протоколы)
#   не обрабатывается: `keepNext` внутри ячейки разрыва не держит, там границы задаются явно.
_SIG_FIO_RX = re.compile(
    r"(?:\t| {2,})(?:[А-ЯЁ]\.\s?[А-ЯЁ]\.\s*[А-ЯЁ][а-яё]+|[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ]\.\s?[А-ЯЁ]\.)\s*$")
_SIG_BLOCK_MAX = 6      # длиннее блок не бывает: орган (2–3 строки) + звание + пустые отбивки
_SIG_LEAD_BLANKS = 2    # пустые отбивки ПЕРЕД блоком: в make_docx они часть блока и несут keepNext


def _sig_text(para):
    """Текст абзаца ВМЕСТЕ с табуляциями (в подписной строке фамилия отбита табом)."""
    return _para_full(para) or para.text or ""


def _signature_block(paras, i, *, lead_blanks=_SIG_LEAD_BLANKS):
    """Индексы абзацев подписного блока, кончающегося строкой Ф.И.О. (индекс `i`).

    Вверх идём, пока строки короткие и не кончаются на «.:;» — такая строка это уже текст
    документа, а не шапка подписи. Пустые отбивки ВНУТРИ блока (их в бланках 1–2) в блок входят:
    разрыв, легший ровно на них, рвёт подпись так же.
    ⚠ Ведущие пустые `style_lint` в блок не берёт, а склеить их всё равно надо — в `make_docx`
    отбивки часть блока (`_glue_block(made)`), и разрыв на них уносит подпись целиком."""
    blk, j, empty = [i], i - 1, 0
    while j >= 0 and len(blk) < _SIG_BLOCK_MAX:
        t = _sig_text(paras[j]).strip()
        if not t:
            empty += 1
            if empty > 2:
                break
            blk.insert(0, j)
            j -= 1
            continue
        if len(t) > 120 or t[-1] in ".:;":
            break
        blk.insert(0, j)
        j -= 1
    while blk and not _sig_text(paras[blk[0]]).strip():
        blk.pop(0)
    n = 0
    while blk and blk[0] > 0 and n < lead_blanks and not _sig_text(paras[blk[0] - 1]).strip():
        blk.insert(0, blk[0] - 1)
        n += 1
    return blk


def find_signature_blocks(doc, *, anchor=None, start=None, end=None, with_prev=0):
    """[[номера абзацев ТЕЛА]] подписных блоков — БЕЗ изменения документа (для предполёта).

    Границы блока: по умолчанию определяются ПО СОДЕРЖАНИЮ (строка «звание · Фамилия И.О.»),
    а не по номеру абзаца — номер уезжает после первой же структурной правки.
    `anchor` — фрагмент ПОСЛЕДНЕЙ строки блока (когда подпись нетиповая и регулярка её не видит).
    `start`/`end` — явные границы ПАРОЙ, когда автоопределение не сработало вовсе.
    `with_prev` — сколько абзацев ТЕКСТА перед блоком тоже скрепить (§16.17: последний абзац
    перед подписью). ⚠ Длинный абзац с `keepNext` уедет на следующий лист целиком — ставить осознанно."""
    paras = list(doc.paragraphs)
    blocks = []
    if start or end:
        if not (start and end):
            raise ValueError("glue_signature: явные границы задаются ПАРОЙ — и 'start', и 'end'")
        heads = [i for i, p in enumerate(paras) if start in _sig_text(p)]
        tails = [i for i, p in enumerate(paras) if end in _sig_text(p)]
        if not heads or not tails:
            return []
        a = heads[0]
        after = [i for i in tails if i >= a]
        if not after:
            raise ValueError("glue_signature: 'end' стоит ВЫШЕ 'start' — границы перепутаны местами")
        blocks.append(list(range(a, after[0] + 1)))
    else:
        for i, p in enumerate(paras):
            t = _sig_text(p)
            if anchor is not None:
                if anchor not in t:
                    continue
            elif not _SIG_FIO_RX.search(t):
                continue
            blk = _signature_block(paras, i)
            if len(blk) > 1:
                blocks.append(blk)
    if with_prev:
        for blk in blocks:
            j, n = blk[0] - 1, 0
            while j >= 0 and n < with_prev:
                blk.insert(0, j)
                j -= 1
                n += 1
    return blocks


def glue_signature(doc, *, anchor=None, start=None, end=None, with_prev=0, require=True):
    """Скрепить абзацы подписного блока флагом `keepNext`, чтобы Word не разорвал его по страницам.

    ⚠ Флаг ставится на все абзацы блока, КРОМЕ ПОСЛЕДНЕГО: последний тянуть не за что, а лишний
    `keepNext` на нём подтащил бы к подписи посторонний СЛЕДУЮЩИЙ абзац (`make_docx._glue_block`).
    Если на последнем абзаце флаг уже стоял (ручная починка в Word) — СНИМАЕТСЯ.
    Возвращает {"count": блоков, "glued": скреплено абзацев, "cleared": снято флагов, "blocks": […]}.
    ⛔ Разрыв страницы программно не виден (§16.17): операция убирает ПРИЧИНУ (нечем удержать),
    а не подтверждает отсутствие разрыва. Достоверная проверка одна — рендер в PDF."""
    paras = list(doc.paragraphs)
    blocks = find_signature_blocks(doc, anchor=anchor, start=start, end=end, with_prev=with_prev)
    if not blocks:
        if require:
            raise ValueError(
                "glue_signature: подписной блок НЕ найден — правка не внесена. Автоопределение ищет "
                "в ТЕЛЕ строку «звание · Фамилия И.О.» (перед Ф.И.О. — табуляция или два пробела). "
                "Напечатай структуру (`--paras`) и задай границы явно: "
                '{"op": "glue_signature", "start": "…", "end": "…"}')
        return {"count": 0, "glued": 0, "cleared": 0, "blocks": []}
    glue, tails = set(), set()
    for blk in blocks:
        glue.update(blk[:-1])
        tails.add(blk[-1])
    for i in sorted(glue):
        paras[i].paragraph_format.keep_with_next = True
    cleared = 0
    for i in sorted(tails - glue):          # хвост одного блока может быть серединой соседнего
        if paras[i].paragraph_format.keep_with_next:
            paras[i].paragraph_format.keep_with_next = False
            cleared += 1
    return {"count": len(blocks), "glued": len(glue), "cleared": cleared, "blocks": blocks}



# ---------- НУМЕРАЦИЯ СТРАНИЦ (требование владельца 19.08.2026) ----------
#
# Правило: у документа ДЛИННЕЕ ДВУХ СТРАНИЦ номер страницы стоит в ВЕРХНЕМ колонтитуле,
# ПО ЦЕНТРУ, Times New Roman 13, и на ПЕРВОЙ странице не печатается.
# Формат-спецификация — references/02-format.md §2.24.
#
# ⚠ Первая страница исключается стандартным механизмом Word «Особый колонтитул для первой
# страницы» (different first page). Если у документа уже был общий верхний колонтитул
# (например, № уголовного дела), его содержимое СНАЧАЛА копируется в колонтитул первой
# страницы — иначе включение режима стёрло бы шапку с первого листа.

_PAGE_FLD = "PAGE"


def _hdr_has_page_field(hdr):
    """Есть ли уже поле PAGE в этом колонтитуле."""
    xml = hdr._element.xml if hdr is not None else ""
    return ("w:instrText" in xml and _PAGE_FLD in xml) or ("w:fldSimple" in xml and _PAGE_FLD in xml)


def _make_page_number_par(hdr, font, size):
    """Абзац по центру с полем { PAGE } нужным шрифтом."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.shared import Pt

    par = hdr.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.first_line_indent = Pt(0)

    run = par.add_run()
    run.font.name = font
    run.font.size = Pt(size)
    # кириллический шрифт задаётся отдельно, иначе Word подставит свой
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(a), font)

    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), " %s   \\* MERGEFORMAT " % _PAGE_FLD)
    inner = OxmlElement("w:r")
    inner.append(copy.deepcopy(rPr))
    txt = OxmlElement("w:t")
    txt.text = "1"
    inner.append(txt)
    fld.append(inner)
    run._element.addnext(fld)
    return par


def add_page_numbers(path, *, font="Times New Roman", size=13, min_pages=3,
                     force=False, backup=True):
    """Проставить номера страниц по формату владельца. Возвращает отчёт (dict).

    min_pages — порог: при меньшем числе страниц нумерация НЕ ставится (по умолчанию 3,
    то есть «длиннее двух страниц»). force=True — поставить независимо от числа страниц.
    """
    _check_locked(path)
    pages = None if force else page_count(path)   # force — не поднимать Word ради счёта
    if not force and pages is not None and pages < min_pages:
        return {"applied": False, "pages": pages,
                "reason": "страниц %d — меньше порога %d, нумерация не нужна" % (pages, min_pages)}
    if not force and pages is None:
        return {"applied": False, "pages": None,
                "reason": "число страниц не определено (нет Word и нет docProps/app.xml) — "
                          "посчитать рендером и запустить с force=True, если страниц больше двух"}

    doc = Document(path)
    touched = 0
    for sec in doc.sections:
        if _hdr_has_page_field(sec.header) or _hdr_has_page_field(sec.first_page_header):
            continue
        # ⛔ Сначала сохранить шапку первой страницы, потом включать «особый первый».
        if not sec.different_first_page_header_footer:
            src = sec.header
            had = [p for p in src.paragraphs if p.text.strip()] if src.is_linked_to_previous is False else []
            sec.different_first_page_header_footer = True
            if had:
                first = sec.first_page_header
                for p in had:
                    first._element.append(copy.deepcopy(p._element))
        else:
            sec.different_first_page_header_footer = True
        _make_page_number_par(sec.header, font, size)
        touched += 1

    if touched:
        _save_atomic(doc, path, backup=backup)
    return {"applied": bool(touched), "pages": pages, "sections": touched,
            "reason": "" if touched else "нумерация уже стоит"}


def page_count(path):
    """Число страниц. Word (свой процесс) → docProps/app.xml → None.

    ⛔ Только DispatchEx: Word — одноэкземплярный COM-сервер, Dispatch подключился бы к
    сеансу владельца и закрыл его документы (правило 39).
    """
    try:
        import pythoncom
        import win32com.client as w32
        pythoncom.CoInitialize()
        app = w32.DispatchEx("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                   AddToRecentFiles=False, Visible=False)
            try:
                d.Repaginate()
                return int(d.ComputeStatistics(2))       # wdStatisticPages
            finally:
                d.Close(False)
        finally:
            app.Quit()
    except Exception:
        pass
    try:
        with zipfile.ZipFile(path) as z:
            if "docProps/app.xml" in z.namelist():
                m = re.search(r"<Pages>(\d+)</Pages>", z.read("docProps/app.xml").decode("utf-8", "replace"))
                if m:
                    return int(m.group(1))
    except Exception:
        pass
    return None

# ---------- предполётная проверка: блочная замена не должна накрывать уже изменённый абзац ----------

# Локальный пример исключён из публичной поставки.
# «применено», после чего операция `block` в ТОМ ЖЕ списке заменила всё между якорями — включая
# только что изменённые абзацы. Правки исчезли, а отчёт честно показывал успех. Это худший из
# возможных исходов: молчаливая потеря при зелёном рапорте. Обнаружилось только при чтении файла.
# Поэтому перекрытие теперь ловится ДО записи, а не диагностируется после.
_PARA_OPS = {"replace": "old", "replace_all": "old", "replace_paragraph": "anchor",
             "delete_paragraph": "anchor", "replace_placeholder": "anchor",
             "insert_after": "anchor", "insert_before": "anchor",
             "glue_signature": "anchor", "insert_blank_after": "anchor",
             "bold_paragraph": "anchor", "bold_lead": "anchor"}


def _para_index(paras, needle):
    """Номер первого абзаца, содержащего needle; -1 — не найден (единая нормализация)."""
    if not needle:
        return -1
    hits = _hits(paras, needle, True)
    return hits[0] if hits else -1


def _preflight_overlap(paras, ops):
    """Вернуть список сообщений о перекрытии. Пусто — перекрытий нет."""
    problems = []
    for bi, op in enumerate(ops):
        if op.get("op") != "block":
            continue
        i_start = _para_index(paras, op.get("start"))
        i_end = _para_index(paras, op.get("end"))
        if i_start < 0 or i_end < 0 or i_end <= i_start:
            continue                      # границы не найдены — это забота самой replace_block
        for oi, prev in enumerate(ops[:bi]):
            key = _PARA_OPS.get(prev.get("op", "replace"))
            if not key:
                continue
            idx = _para_index(paras, prev.get(key))
            if i_start < idx < i_end:     # строго МЕЖДУ якорями — ровно то, что затирает block
                problems.append(
                    "операция %d (%s, якорь %r, абзац %d) попадает ВНУТРЬ блока операции %d "
                    "(block %r … %r, абзацы %d–%d) — блочная замена затрёт её результат"
                    % (oi + 1, prev.get("op", "replace"), prev.get(key), idx,
                       bi + 1, op.get("start"), op.get("end"), i_start, i_end))
    return problems


# ---------- предполётная проверка якорей, числа вхождений и повторной замены ----------
#
# Локальный пример исключён из публичной поставки.
# с якорем «т. 1 л.д. 140–141» (3 вхождения) бросил ValueError на середине списка ops — ранее
# применённые операции ТОГО ЖЕ вызова не сохранились (запись атомарная и одна, в самом конце),
# и весь набор пришлось прогонять заново. И так по одной проблеме за прогон. Поэтому ВСЕ якоря
# и все счётчики проверяются ДО первой мутации и все проблемы выдаются ОДНИМ списком.
# ⚠ Дробить набор на отдельные вызовы «чтобы сохранить сделанное» ЗАПРЕЩЕНО — правило 24 (в):
# все правки одного файла идут в один проход, иначе расширяется окно гонки с открытым Word.

_ANCHOR_OPS = {"replace_paragraph": "anchor", "replace_placeholder": "anchor",
               "delete_paragraph": "anchor", "insert_after": "anchor", "insert_before": "anchor",
               "insert_blank_after": "anchor", "bold_paragraph": "anchor", "bold_lead": "anchor"}


def _anchor_hits(paras, anchor, contains=True):
    'Номера абзацев, попадающих под якорь — ОДНА функция с replace_paragraph и count_text.'
    return _hits(paras, anchor, contains)


def _maybe_created(prev_ops, needle):
    """Мог ли якорь/фрагмент быть СОЗДАН более ранней операцией того же набора.

    Без этой поправки предполёт ругался бы на законный порядок «сначала вставили абзац,
    потом правим его по якорю»: список абзацев снят ДО применения операций."""
    n = _norm(needle)
    if not n:
        return False
    for op in prev_ops:
        for k in ("new", "text"):
            v = op.get(k)
            if isinstance(v, str) and n in _norm(v):
                return True
        for ln in (op.get("lines") or []):
            if isinstance(ln, str) and n in _norm(ln):
                return True
    return False


def _occurrences(paras, old):
    pat = _pattern(old)
    return sum(len(pat.findall(_para_text(p))) for p in paras)


# ---------- (1) `old` последующей замены ВНУТРИ вставляемого блока ----------
#
# Локальный пример исключён из публичной поставки.
# Локальный пример исключён из публичной поставки.
# рассчитанная на 7 вхождений в СТАРОМ тексте. Замена накрыла и только что вставленный блок:
# 12 замен вместо 7, раздел испорчен, документ пересобран с нуля.
# `_preflight_overlap` этого не видел: он смотрит только, попадает ли ЯКОРЬ ранней операции
# внутрь блока, и никогда не заглядывает в САМ ТЕКСТ вставляемых строк.
# Здесь проверка обратного направления: что́ вставили — то последующая массовая замена и затрёт.

_INSERTING_OPS = ("block", "insert_after", "insert_before", "append")


def _inserted_lines(op):
    """Строки, которые операция ДОБАВЛЯЕТ в документ (для проверки последующих замен)."""
    kind = op.get("op", "replace")
    if kind == "block":
        return [str(x) for x in (op.get("lines") or []) if isinstance(x, str)]
    if kind in ("insert_after", "insert_before", "append"):
        v = op.get("text")
        if isinstance(v, (list, tuple)):     # список -> по абзацу на строку, проверять надо КАЖДУЮ
            return [str(x) for x in v if str(x)]
        return [v] if isinstance(v, str) and v else []
    return []


def _preflight_block_content(ops):
    """Сообщения о том, что `old` замены встречается в тексте РАНЕЕ вставляемого блока."""
    problems = []
    for i, op in enumerate(ops):
        kind = op.get("op", "replace")
        if kind not in ("replace", "replace_all") or op.get("allow_repeat"):
            continue
        old = str(op.get("old") or "")
        if not _norm(old):
            continue
        pat = _pattern(old)
        for j, prev in enumerate(ops[:i]):
            if prev.get("op") not in _INSERTING_OPS:
                continue
            where = []
            total = 0
            for k, ln in enumerate(_inserted_lines(prev)):
                n = len(pat.findall(ln))
                if n:
                    total += n
                    where.append("строка %d: «%s»" % (k + 1, _norm(ln)[:70]))
            if total:
                problems.append(
                    "операция %d (%s): %r встречается %d раз(а) в тексте, который ВСТАВЛЯЕТ "
                    "операция %d (%s) — %s. Замена накрыла бы и только что вставленный блок "
                    '(в ОЗ Фамилия так затёрло 5 законных упоминаний: 12 замен вместо 7). '
                    "Правильный порядок — сначала замена по старому тексту, потом вставка блока; "
                    "если перекрытие осознанное — allow_repeat=True"
                    % (i + 1, kind, old[:60], total, j + 1, prev.get("op"),
                       "; ".join(where[:3]) + ("; …" if len(where) > 3 else "")))
    return problems


# ---------- (2) МОДЕЛЬ ДОКУМЕНТА для `expect_after` ----------
#
# Локальный пример исключён из публичной поставки.
# вхождений оставалось ровно сколько нужно — `expect` считается по ИСХОДНОМУ файлу. Приходилось
# снимать `expect` целиком и терять защиту на всём наборе. `expect_after` считает вхождения
# на момент выполнения ЭТОЙ операции, с учётом уже смоделированных предыдущих.
# Модель ТЕКСТОВАЯ (список нормализованных строк-абзацев): она отвечает на вопрос «сколько
# вхождений будет», а не воспроизводит форматирование.


def _model_texts(paras):
    return [_norm(_para_text(p)) for p in paras]


def _model_pattern(needle):
    """Шаблон по нормализованной строке; None — искать нечего.

    ⛔ Без этой отсечки пустой якорь давал `_pattern("")`, который совпадает С ЛЮБЫМ абзацем:
    модель «правила» первый попавшийся, и `expect_after` считал ерунду."""
    n = _norm(needle if isinstance(needle, str) else "")
    return _pattern(n) if n else None


def _model_hits(texts, needle):
    pat = _model_pattern(needle)
    return [i for i, t in enumerate(texts) if pat.search(t)] if pat else []


def _model_apply(texts, op):
    """Применить ТЕКСТОВЫЙ эффект операции к модели (список строк-абзацев). Правит на месте."""
    kind = op.get("op", "replace")
    try:
        if kind in ("replace", "replace_all"):
            pat = _model_pattern(op.get("old"))
            if pat is None:
                return texts
            new = str(op.get("new") or "")
            for i, t in enumerate(texts):
                if not pat.search(t):
                    continue
                if kind == "replace":
                    texts[i] = pat.sub(lambda m: new, t, count=1)
                    break
                texts[i] = pat.sub(lambda m: new, t)
        elif kind in ("replace_paragraph", "replace_placeholder"):
            key = "new" if kind == "replace_paragraph" else "text"
            hits = _model_hits(texts, op.get("anchor"))
            if hits:
                texts[hits[0]] = _norm(str(op.get(key) or ""))
        elif kind == "delete_paragraph":
            limit = int(op.get("max_count", 1))
            pat = _model_pattern(op.get("anchor"))
            if pat is None:
                return texts
            out, n = [], 0
            for t in texts:
                if n < limit and pat.search(t):
                    n += 1
                    continue
                out.append(t)
            texts[:] = out
        elif kind in ("insert_after", "insert_before"):
            hits = _model_hits(texts, op.get("anchor"))
            if hits:
                at = hits[0] + (1 if kind == "insert_after" else 0)
                # список -> НЕСКОЛЬКО абзацев: иначе `expect_after` следующих операций считал по
                # модели из одного склеенного абзаца и расходился с документом
                texts[at:at] = [_norm(x) for x in _as_lines(op.get("text") or "")]
        elif kind == "insert_blank_after":
            hits = _model_hits(texts, op.get("anchor"))
            if hits:
                for _ in range(int(op.get("count", 1))):
                    texts.insert(hits[0] + 1, "")
        elif kind == "append":
            texts.append(_norm(str(op.get("text") or "")))
        elif kind == "block":
            si = _model_index(texts, op.get("start"))
            ei = _model_index(texts, op.get("end"), after=si)
            if si is not None and ei is not None and ei > si:
                texts[si + 1:ei] = [_norm(str(x)) for x in (op.get("lines") or [])]
        elif kind == "clone_block":
            si = _model_index(texts, op.get("start"))
            ei = _model_index(texts, op.get("end"), after=si)
            if op.get("start_index") is not None and op.get("end_index") is not None:
                si, ei = int(op["start_index"]), int(op["end_index"])
            if si is not None and ei is not None and ei >= si:
                span = texts[si:ei + 1]
                at = si if op.get("position") == "before" else ei + 1
                for _ in range(int(op.get("times", 1))):
                    texts[at:at] = list(span)
    except Exception:
        pass          # модель — вспомогательная; её сбой не должен ронять предполёт
    return texts


def _model_index(texts, needle, after=None):
    pat = _model_pattern(needle)
    if pat is None:
        return None
    start = 0 if after is None else after + 1
    for i in range(start, len(texts)):
        if pat.search(texts[i]):
            return i
    return None


def _model_count(texts, op):
    """Сколько вхождений/попаданий якоря увидит операция в смоделированном состоянии.
    None — для этой операции счёт не определён (`expect_after` тогда не поддерживается)."""
    kind = op.get("op", "replace")
    if kind in ("replace", "replace_all"):
        pat = _model_pattern(op.get("old"))
        return sum(len(pat.findall(t)) for t in texts) if pat else None
    key = _ANCHOR_OPS.get(kind)
    if key and op.get(key):
        if not op.get("contains", True):
            return sum(1 for t in texts if t == _norm(op[key]))
        pat = _model_pattern(op[key])
        return sum(1 for t in texts if pat.search(t)) if pat else None
    return None


def _prev_adds_signature(prev_ops):
    """Мог ли подписной блок ПОЯВИТЬСЯ по ходу набора (список абзацев снят ДО применения ops)."""
    for op in prev_ops:
        for k in ("new", "text"):
            v = op.get(k)
            if isinstance(v, str) and _SIG_FIO_RX.search(v):
                return True
        for ln in (op.get("lines") or []):
            if isinstance(ln, str) and _SIG_FIO_RX.search(ln):
                return True
    return False


def _preflight_ops(paras, ops, doc=None):
    """Проверить набор ДО записи: якоря, ожидаемое число вхождений, повторная замена.

    Возвращает список сообщений; пусто — можно применять."""
    problems = _preflight_block_content(ops)
    model = _model_texts(paras)      # состояние документа С УЧЁТОМ уже смоделированных операций
    for i, op in enumerate(ops):
        kind = op.get("op", "replace")
        prev = ops[:i]
        n = None

        # (1) ЯКОРЬ: найден и уникален (П-37). Неуникальный якорь правит/удаляет ПЕРВОЕ вхождение.
        key = _ANCHOR_OPS.get(kind)
        if key and isinstance(op.get(key), str) and op[key]:
            anchor = op[key]
            hits = _anchor_hits(paras, anchor, op.get("contains", True))
            n = len(hits)
            if _maybe_created(prev, anchor):
                n = None                     # абзац появится по ходу набора — считать нечего
            elif not hits:
                if op.get("require", True):
                    problems.append(
                        "операция %d (%s): якорь %r НЕ найден ни в одном абзаце%s"
                        % (i + 1, kind, anchor, _hint_not_found(paras, anchor)))
            else:
                limit = op.get("max_count", 1) if kind == "delete_paragraph" else 1
                if len(hits) > limit and not op.get("allow_multiple"):
                    lst = "; ".join("абз. %d: %s" % (j, (paras[j].text.strip()[:70] or "«пусто»"))
                                    for j in hits[:5])
                    problems.append(
                        "операция %d (%s): якорь %r встречается в %d абзацах (%s%s) — правка ушла "
                        "бы в ПЕРВЫЙ. Взять более длинный уникальный якорь либо allow_multiple=True"
                        % (i + 1, kind, anchor, len(hits), lst,
                           "; …" if len(hits) > 5 else ""))

        # (2) ПОВТОРНАЯ ДОПОЛНЯЮЩАЯ ЗАМЕНА (П-36): `old` — начало более длинного `new`,
        # а `new` в документе УЖЕ есть. Реальный провал 22.08.2026: замена
        # «т. 2 л.д. 50–55, 61–63, 78–87» ушла в уже дополненную строку «…78–87, 176–180»
        # и дала «176–180, 176–180».
        if kind in ("replace", "replace_all") and not op.get("allow_repeat"):
            old_n = _WS.sub(" ", str(op.get("old") or "")).strip()
            new_n = _WS.sub(" ", str(op.get("new") or "")).strip()
            if old_n and new_n and old_n != new_n and old_n in new_n:
                pat_new = _pattern(new_n)
                if any(pat_new.search(_para_full(p)) for p in paras):
                    problems.append(
                        "операция %d (%s): новое значение УЖЕ есть в документе, а %r — его начало/часть. "
                        "Замена легла бы в уже дополненную строку и задвоила хвост. Взять более длинный "
                        "`old` (до границы строки) либо replace_paragraph по уникальному якорю; "
                        "осознанный повтор — allow_repeat=True" % (i + 1, kind, old_n[:60]))

        # (3) СКЛЕЙКА ПОДПИСИ: блок должен находиться ДО записи, иначе операция — тихий no-op.
        if kind == "glue_signature":
            if doc is None:
                pass                         # предполёт без документа — проверит сама операция
            else:
                try:
                    blocks = find_signature_blocks(
                        doc, anchor=op.get("anchor"), start=op.get("start"),
                        end=op.get("end"), with_prev=op.get("with_prev", 0))
                except ValueError as e:
                    problems.append("операция %d (glue_signature): %s" % (i + 1, e))
                    continue
                n = len(blocks)
                if not blocks:
                    if op.get("require", True) and not _prev_adds_signature(prev):
                        problems.append(
                            "операция %d (glue_signature): подписной блок в ТЕЛЕ документа НЕ найден "
                            "(ищется строка «звание · Фамилия И.О.») — напечатай структуру "
                            "(`--paras`) и задай границы явно: \"start\"/\"end\"" % (i + 1))
                    n = None
                elif op.get("anchor") and len(blocks) > 1 and not op.get("allow_multiple"):
                    problems.append(
                        "операция %d (glue_signature): якорь %r попал в %d подписных блоков — взять "
                        "уникальный якорь, задать границы \"start\"/\"end\" либо allow_multiple=True"
                        % (i + 1, op.get("anchor"), len(blocks)))

        # (4) ОЖИДАЕМОЕ ЧИСЛО ВХОЖДЕНИЙ (П-39): считать ДО правки и сверять.
        if "expect" in op:
            if kind in ("replace", "replace_all"):
                n = None if _maybe_created(prev, str(op.get("old") or "")) \
                    else _occurrences(paras, op["old"])
            elif kind == "glue_signature":
                pass                         # n — число НАЙДЕННЫХ подписных блоков, посчитано в (3)
            elif not key:
                problems.append("операция %d (%s): `expect` для этой операции не поддерживается"
                                % (i + 1, kind))
                continue
            if n is not None and n != op["expect"]:
                problems.append(
                    "операция %d (%s): ожидалось вхождений %d, фактически %d%s%s"
                    % (i + 1, kind, op["expect"], n,
                       " — `replace` правит только ПЕРВОЕ вхождение, для всех нужен replace_all"
                       if kind == "replace" and n > op["expect"] else "",
                       " (считается по ИСХОДНОМУ файлу; если число должно измениться предыдущими "
                       "операциями набора — бери `expect_after`)" if prev else ""))

        # (5) ОЖИДАЕМОЕ ЧИСЛО НА МОМЕНТ ЭТОЙ ОПЕРАЦИИ (`expect_after`).
        # Локальный пример исключён из публичной поставки.
        # завалил набор, хотя ПОСЛЕ вставки блока вхождений оставалось ровно сколько нужно.
        # Из-за этого `expect` снимали целиком и теряли защиту на всех операциях набора.
        if "expect_after" in op:
            na = _model_count(model, op)
            if na is None:
                problems.append("операция %d (%s): `expect_after` для этой операции не "
                                "поддерживается (нужна замена или операция с якорем)" % (i + 1, kind))
            elif na != op["expect_after"]:
                problems.append(
                    "операция %d (%s): `expect_after` %d, а с учётом предыдущих операций набора "
                    "выходит %d — сверь, что именно добавляют/убирают операции 1–%d"
                    % (i + 1, kind, op["expect_after"], na, i))

        _model_apply(model, op)      # состояние для `expect_after` следующих операций
    return problems


# ---------- имена ключей операции: синонимы и внятный отказ ----------
#
# Локальный пример исключён из публичной поставки.
# ждёт "new", `block` — "lines", `replace` — "old"/"new", — и набор с привычным "text" падал
# ГОЛЫМ `KeyError: 'new'` из середины применения: ни номера операции, ни списка ожидаемых ключей.
# Диагностика такого отказа стоит хода целиком. Теперь: синоним подставляется, а чего не хватает
# по-настоящему — говорится ДО открытия файла и с перечнем ожидаемых ключей.
_OP_REQUIRED = {
    "replace": ("old", "new"),
    "replace_all": ("old", "new"),
    "replace_paragraph": ("anchor", "new"),
    "replace_placeholder": ("anchor", "text"),
    "insert_after": ("anchor", "text"),
    "insert_before": ("anchor", "text"),
    "delete_paragraph": ("anchor",),
    "append": ("text",),
    "block": ("start", "end", "lines"),
    "glue_signature": (),
    "cell": ("table", "row", "col", "text"),
    # границы задаются парой якорей ЛИБО парой индексов — проверяет сама операция (_resolve_range)
    "clone_block": (),
    "keep_with_next": (),
    "insert_blank_after": ("anchor",),
    "bold_paragraph": ("anchor",),
    "bold_lead": ("anchor",),
    "stamp": (),                 # угловик по эталону: строки — из константы.json либо "lines"
}
# Синоним берётся ТОЛЬКО когда канонического ключа нет: явно переданный ключ всегда сильнее.
_OP_SYNONYMS = {
    "old": ("старое", "find", "search"),
    "new": ("text", "новое", "value"),
    "text": ("new", "текст", "value"),
    "anchor": ("якорь", "target"),
    "lines": ("text", "new", "строки"),
    "start": ("начало", "from"),
    "end": ("конец", "to"),
}
# Ключи, куда СПИСОК передавать нельзя (склеился бы в один абзац). insert_after/insert_before
# и block здесь отсутствуют СОЗНАТЕЛЬНО: там многоабзацная вставка штатная.
_STR_ONLY = {
    "replace": ("old", "new"),
    "replace_all": ("old", "new"),
    "replace_paragraph": ("new",),
    "replace_placeholder": ("text",),
    "append": ("text",),
    "cell": ("text",),
}


def _normalize_op(op, i):
    """Привести ключи операции к каноническим; нехватку объявить ValueError, а не KeyError."""
    kind = op.get("op", "replace")
    req = _OP_REQUIRED.get(kind)
    if req is None:
        return op                      # неизвестная операция — про неё скажет сам edit_file
    out = dict(op)
    for k in req:
        if k in out:
            continue
        for alt in _OP_SYNONYMS.get(k, ()):
            if alt in out:
                out[k] = out[alt]
                break
    if kind == "block" and isinstance(out.get("lines"), str):
        out["lines"] = out["lines"].split("\n")   # одна строка вместо списка — не повод падать
    # ⛔ СПИСОК ТАМ, ГДЕ ЖДУТ ОДНУ СТРОКУ, — отклонять, а не склеивать молча. python-docx перебирает
    # переданное посимвольно, и список строк схлопывается в один абзац без единой жалобы (п. 5.2
    # отчёта 05.09.2026). Многоабзацная вставка — это insert_after/insert_before (там список
    # поддержан) либо `block` c "lines".
    for k in _STR_ONLY.get(kind, ()):
        if isinstance(out.get(k), (list, tuple)):
            raise ValueError(
                "операция %d (%s): ключ %r получил СПИСОК из %d строк, а ждёт одну строку — "
                "абзацы склеились бы в один молча. Для перечня возьми "
                "{\"op\": \"insert_after\", \"anchor\": …, \"text\": [строки]} (каждая строка — свой "
                "абзац) либо {\"op\": \"block\", \"start\": …, \"end\": …, \"lines\": [строки]}. "
                "Файл НЕ открывался." % (i + 1, kind, k, len(out[k])))
    miss = [k for k in req if k not in out]
    if miss:
        raise ValueError(
            "операция %d (%s): нет ключ%s %s. Ожидаются ключи: %s (принимаются синонимы: %s). "
            "Передано: %s. Файл НЕ открывался, ни одна правка не применена."
            % (i + 1, kind, "а" if len(miss) == 1 else "ей",
               ", ".join(repr(k) for k in miss), ", ".join(req),
               "; ".join("%s ← %s" % (k, "/".join(_OP_SYNONYMS.get(k, ()) or ("—",))) for k in req),
               ", ".join(sorted(k for k in op if k != "op")) or "«ничего»"))
    return out


# ---------- единая точка входа ----------

def edit_file(path, ops, backup=True, allow_overlap=False, preflight=True):
    """Открыть существующий .docx, применить операции, атомарно сохранить В ТОТ ЖЕ файл.
    Перед записью проверяет блокировку Word и делает .bak_<timestamp> (+ prune).

    ⛔ ПРОВЕРЯЕМЫЙ РЕЗУЛЬТАТ МАССОВОЙ ПРАВКИ (правило 53: сверяется СЧЁТ, а не статус):
        `replace_all` -> "indices" (номера затронутых абзацев) и "paras_touched";
        `block`       -> "paras_before"/"paras_after" (абзацев между якорями до и после) и
                         "total_before"/"total_after" (по документу). «matched: True» —
        не доказательство: 05.09.2026 регулярка не захватила две строки описи, отчёт был зелёный.

    ⚠ "expect_after":N — то же, что `expect`, но счёт НА МОМЕНТ ЭТОЙ операции, с учётом уже
    смоделированных предыдущих операций набора. `expect` считает по ИСХОДНОМУ файлу и потому
    валит законный порядок «вставили блок -> правим по нему»; тогда берётся `expect_after`.

    ⛔ ПЕРЕКРЫТИЕ ЗАПРЕЩЕНО: если операция `block` накрывает абзац, изменённый более ранней
    операцией того же списка, вызов отклоняется ДО записи (ValueError). Такая правка терялась
    молча, а отчёт показывал успех. Осознанное перекрытие — allow_overlap=True.

    ⛔ ПРЕДПОЛЁТ (`preflight=True`): якоря, `expect` и повторная дополняющая замена проверяются
    ДО первой мутации, и ВСЕ проблемы выдаются одним списком — чтобы набор не приходилось
    прогонять заново по одной ошибке за прогон.

    ⚠ ИМЕНА КЛЮЧЕЙ: канонические — см. выше, но "text"/"new"/"lines" (и «старое»/«новое»/«якорь»)
    принимаются как синонимы, а нехватка ключа объявляется ValueError с перечнем ожидаемых
    ДО открытия файла — вместо голого KeyError из середины применения (_normalize_op)."""
    ops = [_normalize_op(o, i) for i, o in enumerate(ops)]
    _check_locked(path)
    doc = Document(path)
    paras = _all_paragraphs(doc)
    if not allow_overlap:
        overlaps = _preflight_overlap(paras, ops)
        if overlaps:
            raise ValueError(
                "перекрывающиеся правки в одном вызове — файл НЕ изменён:\n  · "
                + "\n  · ".join(overlaps)
                + "\nПорядок: сначала блочная замена, ОТДЕЛЬНЫМ вызовом — правки шапки и подписей "
                  "(§25.04). Если перекрытие осознанное — allow_overlap=True.")
    if preflight:
        problems = _preflight_ops(paras, ops, doc=doc)
        if problems:
            raise ValueError(
                "предполётная проверка не пройдена — файл НЕ изменён, НИ ОДНА правка не применена:\n  · "
                + "\n  · ".join(problems)
                + "\nПочини ВЕСЬ список и прогони набор ОДНИМ вызовом (правило 24 (в)); "
                  "структура абзацев — `python docx_edit.py \"файл.docx\" --paras`.")
    report = []
    for op in ops:
        kind = op.get("op", "replace")
        if kind == "replace_all":
            pat = _pattern(op["old"])
            n = 0
            touched = []          # ⚠ правило 53: после массовой правки сверяется СЧЁТ, а не статус
            for pi, p in enumerate(paras):
                k = _replace_all_in_para(p, pat, op["new"])
                if k:
                    touched.append(pi)
                n += k
            if n == 0 and op.get("require", True):
                raise ValueError(
                    "replace_all: текст НЕ найден (require=True) — правка не внесена, файл НЕ перезаписан: "
                    + repr(op["old"]) + ". Часто фраза разбита Word'ом по абзацам/runs — проверь через find_text/--check "
                    "или передай require=False, если пропуск допустим.")
            report.append({"op": kind, "old": op["old"], "matched": n > 0, "count": n,
                           "replaced": n, "indices": touched, "paras_touched": len(touched)})
        elif kind == "replace":
            pat = _pattern(op["old"])
            # СНАЧАЛА посчитать ВСЕ вхождения (П-36/П-39): `replace` правит только ПЕРВОЕ, и
            # «matched: true» при count > 1 означает, что остальные вхождения ОСТАЛИСЬ старыми.
            total = sum(len(pat.findall(_para_full(p))) for p in paras)
            done = False
            for p in paras:
                if _replace_one(p, pat, op["new"]):
                    done = True
                    break
            if not done and op.get("require", True):
                raise ValueError(
                    "replace: текст НЕ найден (require=True) — правка не внесена, файл НЕ перезаписан: "
                    + repr(op["old"]) + ". Часто фраза разбита Word'ом по абзацам/runs — проверь через find_text/--check "
                    "или передай require=False, если пропуск допустим.")
            report.append({"op": kind, "old": op["old"], "matched": done,
                           "count": total, "replaced": 1 if done else 0})
        elif kind == "replace_paragraph":
            replace_paragraph(doc, op["anchor"], op["new"], contains=op.get("contains", True),
                              lead=op.get("lead"), allow_multiple=op.get("allow_multiple", False))
            report.append({"op": kind, "anchor": op["anchor"], "matched": True})
        elif kind == "delete_paragraph":
            n = delete_paragraph(doc, op["anchor"], contains=op.get("contains", True),
                                 require=op.get("require", True),
                                 max_count=op.get("max_count", 1))
            report.append({"op": kind, "anchor": op["anchor"], "matched": n > 0, "count": n})
        elif kind == "replace_placeholder":
            replace_placeholder(doc, op["anchor"], op["text"], contains=op.get("contains", True))
            report.append({"op": kind, "anchor": op["anchor"], "matched": True})
        elif kind in ("insert_after", "insert_before"):
            # "inserted" — сколько АБЗАЦЕВ реально легло (список -> по абзацу на строку): правило 53,
            # сверяется счёт, а не статус «matched: True».
            n = insert_paragraph(doc, op["anchor"], op["text"],
                                 position="after" if kind == "insert_after" else "before")
            report.append({"op": kind, "anchor": op["anchor"], "matched": True, "inserted": n})
        elif kind == "append":
            append_paragraph(doc, op["text"], before_signature=op.get("before_signature", False))
            report.append({"op": kind, "matched": True})
        elif kind == "block":
            # ⚠ правило 53 («после массовой правки сверяется СЧЁТ, а не статус») получает
            # машинную опору: сколько абзацев было МЕЖДУ якорями и сколько стало. Провал
            # 05.09.2026 (опись в постановлении о выделении): регулярка не захватила две
            # строки, `block` вернул matched: True, потеря вскрылась сверкой 77/439 против 79/441.
            body = doc.paragraphs
            _si = _para_index(body, op["start"])
            _ei = next((i for i in _hits(body, op["end"], True) if i > _si), -1) if _si >= 0 else -1
            _before_range = (_ei - _si - 1) if (_si >= 0 and _ei > _si) else None
            _before_total = len(paras)
            replace_block(doc, op["start"], op["end"], op["lines"])
            _after = _all_paragraphs(doc)
            report.append({"op": kind, "matched": True,
                           "paras_before": _before_range, "paras_after": len(op["lines"]),
                           "total_before": _before_total, "total_after": len(_after)})
        elif kind == "glue_signature":
            r = glue_signature(doc, anchor=op.get("anchor"), start=op.get("start"),
                               end=op.get("end"), with_prev=op.get("with_prev", 0),
                               require=op.get("require", True))
            report.append({"op": kind, "matched": r["count"] > 0, "count": r["count"],
                           "glued": r["glued"], "cleared": r["cleared"], "blocks": r["blocks"]})
        elif kind == "clone_block":
            r = clone_block(doc, op.get("start"), op.get("end"),
                            start_index=op.get("start_index"), end_index=op.get("end_index"),
                            position=op.get("position", "after"), times=op.get("times", 1))
            report.append(dict({"op": kind, "matched": True}, **r))
        elif kind == "keep_with_next":
            r = keep_with_next(doc, anchor=op.get("anchor"), start=op.get("start"),
                               end=op.get("end"), start_index=op.get("start_index"),
                               end_index=op.get("end_index"), contains=op.get("contains", True),
                               with_last=op.get("with_last", False),
                               allow_multiple=op.get("allow_multiple", False))
            report.append(dict({"op": kind, "matched": True}, **r))
        elif kind == "insert_blank_after":
            r = insert_blank_after(doc, op["anchor"], op.get("count", 1),
                                   contains=op.get("contains", True))
            report.append(dict({"op": kind, "anchor": op["anchor"], "matched": True}, **r))
        elif kind == "bold_paragraph":
            r = bold_paragraph(doc, op["anchor"], bold=op.get("bold", True),
                               contains=op.get("contains", True),
                               allow_multiple=op.get("allow_multiple", False))
            report.append(dict({"op": kind, "anchor": op["anchor"], "matched": True}, **r))
        elif kind == "bold_lead":
            r = bold_lead(doc, op["anchor"], sep=op.get("sep", ","),
                          include_sep=op.get("include_sep", False),
                          contains=op.get("contains", True), rest_bold=op.get("rest_bold"),
                          allow_multiple=op.get("allow_multiple", False),
                          occurrence=op.get("occurrence", 1))
            report.append(dict({"op": kind, "anchor": op["anchor"], "matched": True}, **r))
        elif kind == "cell":
            tbl = doc.tables[op["table"]]
            cell = tbl.rows[op["row"]].cells[op["col"]]
            if op.get("statcard"):
                ff = op.get("font_from")
                font_cell = tbl.rows[ff[0]].cells[ff[1]] if ff else None
                set_statcard_cell(cell, op["text"], font_from_cell=font_cell)
            else:
                set_cell_text(cell, op["text"], bold=op.get("bold", False))
            report.append({"op": kind, "matched": True})
        elif kind == "stamp":
            r = set_stamp(doc, op.get("lines"), sizes=op.get("sizes"),
                          anchor=op.get("anchor", "СК РОССИИ"), require=op.get("require", True))
            report.append(dict({"op": kind, "matched": r["boxes"] > 0}, **r))
        else:
            raise ValueError("Неизвестная операция: " + str(kind))
        # обновляем список абзацев после структурных операций (вставка/блок/клон)
        if kind in ("insert_after", "insert_before", "append", "block",
                    "clone_block", "insert_blank_after", "delete_paragraph", "stamp"):
            paras = _all_paragraphs(doc)
    _save_atomic(doc, path, backup=backup)
    return report


# ---------- угловой штамп (надпись бланка) ----------

_STAMP_KINDS = ("h1", "bold", "plain", "blank")


def _stamp_sizes_default():
    """Кегли по видам строк (полупункты) из константы «угловик» → «кегль»; нет ключа — None."""
    import json
    path = os.environ.get("SK_CONSTANTS") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "константы.json")
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)["угловик"].get("кегль") or None
    except Exception:
        return None


def _stamp_lines_default():
    """Строки эталонного угловика из ПЕРСОНАЛЬНОГО слоя: константы.json → «угловик» → «строки».

    Ядро текста штампа не знает (реквизиты — персональный слой). Нет ключа — ValueError с
    подсказкой, а не молчаливый пропуск."""
    import json
    path = os.environ.get("SK_CONSTANTS") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "константы.json")
    try:
        with open(path, encoding="utf-8") as fh:
            lines = json.load(fh)["угловик"]["строки"]
    except Exception as e:
        raise ValueError(
            "stamp: строки эталонного угловика не заданы — передай \"lines\" в операции либо заведи "
            "в константы.json ключ «угловик» → «строки» ([[текст, вид], …], вид: h1/bold/plain/blank): "
            + repr(e))
    return lines


def _p_text_el(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t")))


def _p_is_bold(p_el):
    for r in p_el.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None and rpr.find(qn("w:b")) is not None:
            return True
    return False


def _p_has_image(p_el):
    return any(el.tag in (qn("w:drawing"), qn("w:pict")) for el in p_el.iter())


def set_stamp(doc, lines=None, *, sizes=None, anchor="СК РОССИИ", require=True):
    """Переписать ТЕКСТ углового штампа (надписи бланка) по эталону — {"op": "stamp"}.

    Надпись бланка 384 ВСО устроена так: [абзац с гербом] → текстовые абзацы реквизитов →
    таблица «дата | № | исходящий» → хвост. Переписываются ТОЛЬКО текстовые абзацы между гербом
    и таблицей; герб, таблица (в ней уже стоят дата и номер) и всё после неё не трогаются.
    Формат каждой новой строки клонируется из старого абзаца того же ВИДА (h1 — первый жирный,
    bold — жирный, plain — нежирный, blank — пустой), поэтому кегль и шрифт остаются бланковыми.
    Обрабатываются ВСЕ надписи с якорем (у шаблона их две — DrawingML и VML-запасная).

    Зачем (09.09.2026): эталон угловика — бумажный бланк (скан в «Шаблоны и бланки»), а доноры
    корпуса несут старую редакцию («ВСУ СК России по ЦВО», «(по гарнизону)», адрес без
    республики). Строки эталона живут в персональном слое (константы.json → «угловик»)."""
    lines = lines or _stamp_lines_default()
    if sizes is None:
        sizes = _stamp_sizes_default() or {}
    norm_lines = []
    for item in lines:
        if isinstance(item, str):
            text, kind = item, "plain"
        else:
            text, kind = (list(item) + ["plain"])[:2]
        text = str(text)
        if not text.strip():
            kind = "blank"
        if kind not in _STAMP_KINDS:
            raise ValueError("stamp: неизвестный вид строки %r (ожидается h1/bold/plain/blank)" % (kind,))
        norm_lines.append((text, kind))

    boxes = [b for b in doc.element.body.iter(qn("w:txbxContent"))
             if anchor in "".join(t.text or "" for t in b.iter(qn("w:t")))]
    if not boxes:
        if require:
            raise ValueError(
                "stamp: надпись с якорем %r не найдена — в документе нет углового штампа бланка "
                "(или он вставлен картинкой); файл НЕ изменён" % anchor)
        return {"boxes": 0, "lines": len(norm_lines)}

    for box in boxes:
        kids = list(box)
        start = 0
        for i, el in enumerate(kids):
            if el.tag == qn("w:p") and _p_has_image(el):
                start = i + 1
                break
        end = len(kids)
        for i in range(start, len(kids)):
            if kids[i].tag == qn("w:tbl"):
                end = i
                break
        old = [el for el in kids[start:end] if el.tag == qn("w:p")]
        if not old:
            raise ValueError("stamp: в надписи нет текстовых абзацев между гербом и таблицей — "
                             "структура не бланковая, файл НЕ изменён")
        texted = [q for q in old if _p_text_el(q).strip()]
        bold = [q for q in texted if _p_is_bold(q)]
        plain = [q for q in texted if not _p_is_bold(q)]
        blank = [q for q in old if not _p_text_el(q).strip()]
        if not texted:
            raise ValueError("stamp: в надписи нет ни одного абзаца с текстом — образца формата нет")
        src = {
            "h1": bold[0] if bold else texted[0],
            "bold": bold[1] if len(bold) > 1 else (bold[0] if bold else texted[0]),
            "plain": plain[0] if plain else texted[-1],
            "blank": blank[0] if blank else texted[0],
        }
        new_els = []
        for text, kind in norm_lines:
            el = copy.deepcopy(src[kind])
            for a in list(el.attrib):              # клон не должен нести чужие paraId/textId
                if a.endswith("}paraId") or a.endswith("}textId"):
                    del el.attrib[a]
            runs = el.findall(qn("w:r"))
            if kind == "blank":
                for r in runs:
                    el.remove(r)
                new_els.append(el)
                continue
            keep = None
            for r in runs:
                if keep is None and r.find(qn("w:t")) is not None:
                    keep = r
                else:
                    el.remove(r)
            if keep is None:
                keep = el.makeelement(qn("w:r"), {})
                el.append(keep)
            for child in list(keep):
                if child.tag != qn("w:rPr"):
                    keep.remove(child)
            # ⛔ Кегль — ЯВНО. Строка без w:sz наследует Normal, а `_setup_document_style` генератора
            # делает Normal = TNR 13: адрес раздувался и переносился на вторую строку (поймано
            # рендером пробного письма 09.09.2026). Бланк — 12 пт (sz 24), как на бумаге.
            rpr = keep.find(qn("w:rPr"))
            if rpr is None:
                rpr = keep.makeelement(qn("w:rPr"), {})
                keep.insert(0, rpr)
            # Кегль по виду строки — константа «угловик» → «кегль» (сняты с бумажного бланка по
            # ширине строк: h1 12 / bold 11 / plain 10 пт); нет значения — 12 пт, если кегля не было.
            want = sizes.get(kind) if isinstance(sizes, dict) else None
            if want is not None or rpr.find(qn("w:sz")) is None:
                for tag in (qn("w:sz"), qn("w:szCs")):
                    for old_el in rpr.findall(tag):
                        rpr.remove(old_el)
                at = 0
                for k, ch in enumerate(list(rpr)):
                    if ch.tag in (qn("w:rFonts"), qn("w:b"), qn("w:bCs"), qn("w:i"), qn("w:iCs")):
                        at = k + 1
                val = str(int(want)) if want is not None else "24"
                rpr.insert(at, rpr.makeelement(qn("w:sz"), {qn("w:val"): val}))
                rpr.insert(at + 1, rpr.makeelement(qn("w:szCs"), {qn("w:val"): val}))
            t_el = keep.makeelement(qn("w:t"), {})
            t_el.text = text
            if text != text.strip():
                t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            keep.append(t_el)
            new_els.append(el)
        pos = list(box).index(old[0])
        for q in old:
            box.remove(q)
        for k, el in enumerate(new_els):
            box.insert(pos + k, el)
    return {"boxes": len(boxes), "lines": len(norm_lines)}


_MONTHS = ("январ|феврал|март|апрел|ма[йя]|июн|июл|август|сентябр|октябр|ноябр|декабр")
_DATE_PATTERNS = (
    ("ДД.ММ.ГГГГ", re.compile(r"\b\d{1,2}\.\d{1,2}\.(?:19|20)\d{2}\b")),
    # ⚠ _MONTHS ОБЯЗАТЕЛЬНО в (?:…): у «|» самый низкий приоритет, и без группы шаблон распадался
    # на «(«ДД»…январ) | (феврал) | … | (декабр…\d{4})» — дата в кавычках не распознавалась,
    # зато любое слово «феврал»/«март» считалось датой (аудит 30.07.2026).
    ("«ДД» месяца ГГГГ", re.compile(r"«\s*\d{1,2}\s*»\s*(?:" + _MONTHS + r")[а-яё]*\s+\d{4}", re.I)),
    ("ДД месяца ГГГГ", re.compile(r"\b\d{1,2}\s+(?:" + _MONTHS + r")[а-яё]*\s+\d{4}", re.I)),
    ("ДД месяца (без года)", re.compile(r"\b\d{1,2}\s+(?:" + _MONTHS + r")[а-яё]*(?!\s+\d{4})\b", re.I)),
    ("ДД.ММ.ГГ", re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{2}\b(?!\d)")),
)


def find_dates(path):
    """Полный аудит дат по ВСЕМ контейнерам: тело, таблицы (в т.ч. бланк-шапка), НАДПИСИ, колонтитулы.

    ⚠ Дата углового штампа бланка ВСО живёт в НАДПИСИ — она в выводе помечена «надпись»
    (проверено 07.09.2026 на копии реального запроса). Но `--dates` показывает РАСПОЗНАННЫЙ
    текст: дату, которая пойдёт в опись или в перечень постановления, подтверждай РЕНДЕРОМ
    листа, а дату, противоречащую хронологии дела, не «поправляй по смыслу».

    Возвращает [{"where":.., "format":.., "date":.., "context":..}] в порядке обхода.
        python docx_edit.py "файл.docx" --dates"""
    doc = path if hasattr(path, "sections") else Document(path)
    found, seen = [], set()
    for para, where in _all_paragraphs_located(doc):
        text = _para_full(para) or para.text
        if not text or not text.strip():
            continue
        for label, rx in _DATE_PATTERNS:
            for m in rx.finditer(text):
                key = (where, m.group(0), m.start(), id(para))
                if key in seen:
                    continue
                seen.add(key)
                s = max(0, m.start() - 30)
                found.append({"where": where, "format": label, "date": m.group(0).strip(),
                              "context": ("…" if s else "") + text[s:m.end() + 30].strip() + "…"})
    return found


def scan_traces(path, forbidden, raise_on_hit=True):
    """Авто-скан на следы документа-донора (Критическое правило 13). Ищет ЗАПРЕЩЁННЫЕ значения
    (чужие фамилии, № дела/КРСП, № в/ч, подписанты, адреса) в тексте .docx — абзацы + ячейки
    таблиц, с нормализацией пробелов и регистра. Возвращает список найденных.
    raise_on_hit=True (по умолчанию) — бросает ValueError при находке, ПРЕКРАЩАЯ выдачу."""
    doc = path if hasattr(path, "tables") else Document(path)
    # ⚠ ВСЕ контейнеры: тело + таблицы + КОЛОНТИТУЛЫ + НАДПИСИ. Раньше сканировались только
    # тело и таблицы — поэтому паспортные данные донора в трёх запросах остались незамеченными
    # (боевой отчёт 29.07.2026, ошибка 12).
    parts = [_para_full(p) or p.text for p in _all_paragraphs(doc)]
    hay = _WS.sub(" ", " ".join(parts)).lower()
    hits = []
    for f in forbidden:
        if not f:
            continue
        needle = _WS.sub(" ", str(f)).strip().lower()
        if not needle:
            continue
        # граница слова/цифры: «Дзен» не матчит «Дзержинский», «Дон» не матчит «Донецк»,
        # «11» не матчит «1123» — иначе короткие/числовые токены ложно блокируют чистый документ.
        if re.search(r"(?<!\w)" + re.escape(needle) + r"(?!\w)", hay):
            hits.append(f)
    if hits and raise_on_hit:
        raise ValueError("Следы донора (правило 13) — документ НЕ выдавать, вычистить: "
                         + ", ".join(map(str, hits)))
    return hits


# Что автоматически считать «реквизитом донора»: паспорт, счёт, в/ч, звание, УИН/СНИЛС,
# Локальный пример исключён из публичной поставки.
# (52 17 № 624896, ОУФМС по Омской области) уехал в три запроса (отчёт 29.07.2026, ошибка 12).
_DONOR_PATTERNS = (
    ("паспорт", re.compile(r"\b\d{2}\s?\d{2}\s*№?\s*\d{6}\b")),
    ("орган выдачи", re.compile(r"\b(?:ОУФМС|УФМС|ТП УФМС|ОВМ|МП ОВД)[^,.;)]{0,60}")),
    ("счёт", re.compile(r"\b\d{20}\b")),
    ("карта", re.compile(r"\b\d{4}\s?\*{4,}\s?\d{4}\b|\*\d{4}\b")),
    ("воинская часть", re.compile(r"войсково[йе]\s+части?\s+\d{4,6}|в/ч\s*\d{4,6}", re.I)),
    ("личный номер", re.compile(r"\b[А-Я]{2}-?\d{6}\b")),
    ("звание", re.compile(r"\b(?:рядовой|ефрейтор|младший сержант|сержант|старший сержант|"
                          r"старшина|прапорщик|младший лейтенант|лейтенант|старший лейтенант|"
                          r"капитан|майор|подполковник|полковник)\b", re.I)),
    ("телефон", re.compile(r"\+7\s?\(?\d{3}\)?[\s-]?\d{3}[\s-]?\d{2}[\s-]?\d{2}")),
    ("СНИЛС/УИН", re.compile(r"\b\d{3}-\d{3}-\d{3}\s?\d{2}\b|\bУИН\s*«?\d{15,25}»?")),
)


def donor_traces(donor_path, *, extra=()):
    """Собрать список ЗАПРЕЩЁННЫХ значений ИЗ ДОНОРА автоматически — для scan_traces.

    Возвращает {категория: [значения]} + плоский список. Смысл: не полагаться на память
    «что заменить», а вытащить из донора всё, что является персональным реквизитом, и затем
    убедиться, что ни одно из этих значений не осталось в новом документе."""
    doc = Document(donor_path)
    hay = _WS.sub(" ", " ".join(_para_full(p) or p.text for p in _all_paragraphs(doc)))
    out, flat = {}, []
    for name, rx in _DONOR_PATTERNS:
        vals = sorted({m.group(0).strip() for m in rx.finditer(hay)})
        if vals:
            out[name] = vals
            flat.extend(vals)
    for e in extra:
        if e:
            flat.append(e)
    out["все"] = sorted(set(flat))
    return out


_FIO_INITIALS = re.compile(r"([А-ЯЁ][а-яё]+(?:ов|ев|ёв|ин|ын|ский|цкий|ко|ук|юк|ян|дзе|швили|"
                           r"а|я|ых|их)?)\s+([А-ЯЁ]\.\s?[А-ЯЁ]\.)")
# Подписной блок ВСО: «И.О. Фамилия» — ОБРАТНЫЙ порядок. Именно в подписи чаще всего остаются
# инициалы донора, а прежний шаблон его не видел вовсе (аудит 30.07.2026).
_INITIALS_FIO = re.compile(r"([А-ЯЁ]\.\s?[А-ЯЁ]\.)\s+([А-ЯЁ][а-яё]{2,})")
# Слова, которые не являются фамилией, но стоят перед инициалами (должность/роль) — иначе
# «Следователь Д.А.» выдавалось как ФИО.
_NOT_SURNAME = {"следователь", "руководитель", "заместитель", "врио", "дознаватель", "прокурор",
                "начальник", "командир", "защитник", "адвокат", "эксперт", "специалист",
                "потерпевший", "свидетель", "подозреваемый", "обвиняемый", "понятой", "переводчик"}


def find_fio_with_initials(path):
    'Найти все связки ФИО+инициалы в документе — В ОБОИХ порядках.'
    doc = path if hasattr(path, "tables") else Document(path)
    hay = _WS.sub(" ", " ".join(_para_full(p) or p.text for p in _all_paragraphs(doc)))
    found = {}
    for rx, surname_group in ((_FIO_INITIALS, 1), (_INITIALS_FIO, 2)):
        for m in rx.finditer(hay):
            if m.group(surname_group).lower() in _NOT_SURNAME:
                continue
            key = _WS.sub(" ", m.group(0)).strip()
            found[key] = found.get(key, 0) + 1
    return dict(sorted(found.items(), key=lambda kv: -kv[1]))


_DOC_EXT = (".docx", ".docm", ".dotx", ".doc")


def _cli_reorder(argv):
    """Поставить позиционный <файл.docx> первым, где бы он ни стоял в строке.

    ⛔ Провал 05.09.2026: `docx_edit.py --glue-signature "файл.docx"` падал с
    PackageNotFoundError («Package not found at '--glue-signature'»), потому что путь брался
    жёстко из argv[0]. Сообщение читается как «файл битый», а не как «флаг стоит не на месте»,
    и ход уходит на диагностику несуществующей поломки документа. Работал только один порядок
    аргументов — `<файл> --glue-signature`.
    ⚠ Кандидатом в путь считается аргумент, который НЕ флаг (не с «-») и НЕ пара «старое=>новое»:
    иначе замена «Приложение.docx=>Приложение № 2.docx» была бы принята за имя файла."""
    for i, a in enumerate(argv):
        if a.startswith("-") or "=>" in a:
            continue
        if a.lower().endswith(_DOC_EXT):
            return [a] + list(argv[:i]) + list(argv[i + 1:])
    return list(argv)


def _cli(argv):
    if argv and argv[0] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        print("Использование: python docx_edit.py <файл.docx> \"старое=>новое\" ...")
        print("              python docx_edit.py <файл.docx> --paras")
        print("              python docx_edit.py <файл.docx> --check \"фрагмент\"")
        print("              python docx_edit.py <файл.docx> --glue-signature")
        print("              (порядок не важен: <файл.docx> может стоять и после флага)")
        return 0
    argv = _cli_reorder(argv)
    if len(argv) < 2:
        print("Использование: python docx_edit.py <файл.docx> \"старое=>новое\" ...")
        print("              python docx_edit.py <файл.docx> --paras")
        print("              python docx_edit.py <файл.docx> --check \"фрагмент\"")
        print("              (порядок не важен: <файл.docx> может стоять и после флага)")
        return 2
    path = argv[0]
    if argv[1] == "--check":
        d = Document(path)
        many = False
        for frag in argv[2:]:
            n = count_text(d, frag)
            many = many or n > 1
            print(("ЕСТЬ (вхождений: %d)  " % n if n else "НЕТ                   ") + frag[:60])
        if many:
            print("⚠ Вхождений больше одного, а `replace` правит ТОЛЬКО ПЕРВОЕ: бери replace_all "
                  "либо более длинный уникальный фрагмент и передай \"expect\": N.")
        return 0
    if argv[1] in ("--paras", "--paragraphs"):
        rows = list_paragraphs(path)
        print("Абзацев (тело + таблицы + надписи + колонтитулы): %d\n" % len(rows))
        w = min(40, max([len(r[1]) for r in rows] or [4]))
        for i, where, ln, txt in rows:
            print("  [%4d] %-*s %4d зн.  %s" % (i, w, where, ln, txt or "«пусто»"))
        print("\n⚠ Длина сильно больше показанного текста — это ОДИН абзац с мягкими переносами "
              "(многострочный заголовок бланка): якорь по второй строке не найдётся никогда.")
        print("⚠ Колонка «где»: «тело» — абзац самого тела; «таблица N, стр. R, кол. C» — ячейка "
              "таблицы (второе вхождение даты в шапке бланка — почти всегда именно она, а не тело: "
              "не ставь на неё \"expect\": 2, не посмотрев сюда).")
        return 0
    if argv[1] in ("--page-numbers", "--pages"):
        force = "--force" in argv[2:]
        rep = add_page_numbers(path, force=force)
        if rep["applied"]:
            print("нумерация проставлена: верхний колонтитул, по центру, Times New Roman 13, "
                  "кроме первой страницы (разделов: %d)" % rep["sections"])
        else:
            print("нумерация НЕ ставилась: " + rep["reason"])
        if rep.get("pages") is not None:
            print("страниц в документе: %d" % rep["pages"])
        return 0
    if argv[1] in ("--glue-signature", "--glue"):
        try:
            rep = edit_file(path, [{"op": "glue_signature"}])
        except DocxLockedError as e:
            print("ЗАНЯТО:", e)
            return 1
        except ValueError as e:
            print("ОШИБКА:", e)
            return 1
        r = rep[0]
        print("подписных блоков: %d · скреплено абзацев (keepNext): %d · снято лишних флагов: %d"
              % (r["count"], r["glued"], r["cleared"]))
        print("абзацы блоков (нумерация ТЕЛА): " + "; ".join(str(b) for b in r["blocks"]))
        print("⛔ Разрыв страницы программно не виден: прогони style_lint заново и, если документ "
              "многостраничный, посмотри рендер в PDF (§16.17).")
        return 0
    if argv[1] == "--dates":
        rows = find_dates(path)
        if not rows:
            print("Дат в документе не найдено.")
            return 0
        print(f"Найдено дат: {len(rows)} (тело + таблицы + надписи + колонтитулы)\n")
        for r in rows:
            print(f"  [{r['where']}] {r['date']}  ({r['format']})")
            print(f"      → {r['context']}")
        print("\nПеред добавлением строки с датой убедись, что даты ещё нет в бланке/таблице/колонтитуле.")
        return 0
    ops = []
    for a in argv[1:]:
        if "=>" not in a:
            print("Пропускаю (нет '=>'):", a)
            continue
        old, new = a.split("=>", 1)
        # CLI — интерактивный инструмент: терпимо (require=False), чтобы применить найденное и напечатать «НЕТ»
        # по ненайденному (а не падать на первом промахе; программный edit_file по умолчанию строгий).
        ops.append({"op": "replace", "old": old, "new": new, "require": False})
    try:
        rep = edit_file(path, ops)
    except DocxLockedError as e:
        print("ЗАНЯТО:", e)
        return 1
    except ValueError as e:
        print("ОШИБКА:", e)
        return 1
    for r in rep:
        label = r.get("old", r.get("anchor", r.get("op", "")))
        print(("OK   " if r.get("matched") else "НЕТ  ") + str(label)[:60])
    print("Готово. Бэкап: " + os.path.basename(path) + ".bak_<timestamp>")
    return 0


if __name__ == "__main__":
    sys.exit(_cli(sys.argv[1:]))
