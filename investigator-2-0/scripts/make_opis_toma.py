# -*- coding: utf-8 -*-
"""make_opis_toma.py — ПЕЧАТНАЯ опись тома (.docx) и обложки на N томов ОТ ДОНОРА.

Зачем: опись тома до сих пор существовала только как .txt — подшить в дело нечего,
и перед сдачей она переписывалась в Word руками. Обложка при этом собиралась ровно
одна («Обложка УД том 1»), а томов в деле три: два тома уезжали прокурору без обложки.
Оба пробела — механические, оба закрываются одной командой.

⛔ Документ собирается ОТ ДОНОРА (правило 12): скрипт открывает готовую опись, клонирует
её строку таблицы со всем форматированием (шрифт, границы, ширины колонок) и подставляет
позиции. Ничего не рисуется с нуля — поэтому вид совпадает с тем, что уже подшивалось.

ВХОД — перечень позиций, по одной в строке (файл или stdin):
    Постановление о возбуждении уголовного дела от 06.03.2026 | 1-3
    Письмо военному прокурору Казанского гарнизона от 06.03.2026 | 4
    Уведомление о возбуждении уголовного дела | 5
Разделитель — «|», табуляция или «;». Второе поле (страницы) можно не указывать.
Ведущая нумерация «12.» в начале строки отбрасывается — нумерует скрипт.

Флаги:
    --delo N        номер уголовного дела (в шапку описи)
    --tom N         номер тома
    --items PATH    файл с позициями («-» — читать stdin)
    --out PATH      куда сохранить (по умолчанию «Опись том N.docx» рядом с --items)
    --donor PATH    иной донор описи
    --start N       начать нумерацию позиций не с 1
    --oblozhka PATH обложка тома 1 — донор для размножения
    --oblozhki N    сколько всего томов: сделать обложки на тома 2…N рядом с донором

ВЫХОД: .docx описи и/или файлы обложек; в консоль — что создано и сколько позиций.
Код возврата: 0 — сделано; 1 — ошибка (нет донора, пустой перечень, нет метки «ТОМ № 1»)."""
import argparse
import copy
import io
import json
import os
import re
import shutil
import sys

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

DONOR_OPIS = None
SPLIT = re.compile(r"\s*[|;\t]\s*")
LEAD_NUM = re.compile(r"^\s*\d{1,3}[.)]\s+")
TOM_RE = re.compile(r"ТОМ\s*№\s*\d+", re.I)


# ── реквизиты подписи ─────────────────────────────────────────────────────────────────────
def _constants():
    """Читает «константы.json» тем же поиском, что и make_docx (рядом со скриптом,
    на уровень выше, SK_CONSTANTS). Возвращает {} — тогда подпись берётся из донора."""
    env = os.environ.get("SK_CONSTANTS")
    cands = [env] if env else []
    cands += [os.path.join(HERE, "константы.json"),
              os.path.join(HERE, "..", "константы.json"),
              os.path.join(HERE, "..", "..", "константы.json")]
    for c in cands:
        if c and os.path.isfile(c):
            try:
                with io.open(c, encoding="utf-8") as fh:
                    return json.load(fh)
            except Exception:
                return {}
    return {}


def signature_lines():
    """Строки подписного блока следователя: [роль, орган (род.), звание, ФИО].

    Сборку роли и органа берём у make_docx (`_role_org_lines`) — там она уже
    выправлена по эталону 01-identity §1.6 (роль с заглавной и БЕЗ отдела, орган
    в родительном падеже отдельной строкой). Если импорт не удался — тот же
    результат собирается локально, чтобы скрипт оставался автономным.
    """
    c = _constants()
    sled = c.get("следователь", {}) if isinstance(c.get("следователь"), dict) else {}
    dolzh = sled.get("должность", "следователь")
    org_gen = c.get("орган_родительный") or c.get("орган_полное") or ""
    zvanie = sled.get("звание", "")
    fio = sled.get("фио", "")
    parts = fio.split()
    if len(parts) >= 3:
        name = "%s.%s. %s" % (parts[1][:1], parts[2][:1], parts[0])
    else:
        name = "%s %s" % (sled.get("инициалы_имо", ""), parts[0] if parts else "")
    try:
        import make_docx
        role, org = make_docx._role_org_lines(dolzh, org_gen)
    except Exception:
        m = re.match(r"^(.*?\S)\s+\d.*$", dolzh or "")
        role = (m.group(1) if m else dolzh) or ""
        role = role[:1].upper() + role[1:]
        org = org_gen
    return [role, org, zvanie, name.strip()]


# ── работа с .docx ────────────────────────────────────────────────────────────────────────
def set_text(par, text):
    """Ставит тексту абзаца новое значение, СОХРАНЯЯ формат первого run'а."""
    runs = par.runs
    if not runs:
        par.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


def fill_row(row, values):
    """Заполняет ячейки строки, сохраняя формат первого run'а каждой ячейки."""
    for cell, val in zip(row.cells, values):
        pars = cell.paragraphs
        set_text(pars[0], val)
        for extra in pars[1:]:
            extra._p.getparent().remove(extra._p)


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def strip_bookmarks(doc):
    """Снимает ВСЕ закладки (`bookmarkStart`/`bookmarkEnd`).

    ⛔ Обязательно: строки донора удаляются и клонируются, и половина закладки
    (`bookmarkEnd`) уезжает вместе с удалённой строкой, а вторая остаётся. Word
    на такой паре падает при сохранении БЕЗ сохранения (docx_integrity ловит это
    как «закладки не сходятся»). В описи закладки не несут смысла — снимаем все.
    """
    for tag in ("bookmarkStart", "bookmarkEnd"):
        for el in list(doc.element.iter(W_NS + tag)):
            el.getparent().remove(el)


def _post_table_paragraphs(doc, table):
    """Абзацы, идущие ПОСЛЕ таблицы (подписной блок донора)."""
    from docx.text.paragraph import Paragraph
    out, seen = [], False
    for ch in doc.element.body.iterchildren():
        if ch is table._tbl:
            seen = True
            continue
        if seen and ch.tag.endswith("}p"):
            out.append(Paragraph(ch, doc))
    return out


def rebuild_signature(doc, table, lines):
    """Перебирает подписной блок донора на реквизиты из констант.

    Абзацы не создаются с нуля: берутся ДВА образца из самого донора — обычная
    строка блока и строка «звание …… И.О. Фамилия» (в ней сохраняется ровно та
    последовательность табуляций, которой донор отбивает ФИО к правому полю).
    """
    paras = _post_table_paragraphs(doc, table)
    filled = [p for p in paras if p.text.strip()]
    if len(filled) < 2 or not lines:
        return False
    role, org, zvanie, fio = lines
    sample_line = filled[0]._p
    sample_sign = filled[-1]._p
    sep = "\t\t\t\t\t\t"
    m = re.match(r"^(\S.*?)((?:\t|\s{3,})+)(\S.*)$", filled[-1].text)
    if m:
        sep = m.group(2)
    body = doc.element.body
    for p in paras:                      # снять донорский блок целиком
        body.remove(p._p)
    from docx.text.paragraph import Paragraph
    sect = body.find(
        "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sectPr")

    def _add(sample, text):
        el = copy.deepcopy(sample)
        if sect is not None:
            sect.addprevious(el)
        else:
            body.append(el)
        set_text(Paragraph(el, doc), text)

    _add(sample_line, "")
    _add(sample_line, "")
    _add(sample_line, role)
    _add(sample_line, org)
    _add(sample_line, "")
    _add(sample_sign, "%s%s%s" % (zvanie, sep, fio))
    return True


def parse_items(text):
    """Строки перечня -> [(наименование, страницы)]. Пустые строки пропускаются."""
    items = []
    for line in text.splitlines():
        line = LEAD_NUM.sub("", line.strip())
        if not line:
            continue
        parts = [p.strip() for p in SPLIT.split(line) if p.strip()]
        if len(parts) >= 2:
            items.append((" ".join(parts[:-1]), parts[-1]))
        else:
            items.append((parts[0], ""))
    return items


def make_opis(donor, out, delo, tom, items, start=1):
    """Собирает опись тома от донора. Возвращает путь к созданному файлу."""
    import docx
    shutil.copyfile(donor, out)
    doc = docx.Document(out)
    if not doc.tables:
        raise ValueError("в доноре нет таблицы — взят не тот файл "
                         "(опись обязана иметь таблицу «№ п/п | документ | стр.»)")
    table = doc.tables[0]

    for par in doc.paragraphs:                       # шапка: номер дела и номер тома
        if par.text.strip().startswith("№") and delo:
            set_text(par, "№\u00a0%s" % delo)
        elif TOM_RE.search(par.text):
            set_text(par, TOM_RE.sub("ТОМ № %d" % tom, par.text.strip()))

    if len(table.rows) < 2:
        raise ValueError("в таблице донора нет строки-образца данных")
    sample = copy.deepcopy(table.rows[1]._tr)
    dot = "." if table.rows[1].cells[0].text.strip().endswith(".") else ""
    for row in list(table.rows)[1:]:                 # снять данные донора
        row._tr.getparent().remove(row._tr)
    for i, (name, pages) in enumerate(items, start):
        tr = copy.deepcopy(sample)
        table._tbl.append(tr)
        fill_row(table.rows[-1], ["%d%s" % (i, dot), name, pages])

    rebuild_signature(doc, table, signature_lines())
    strip_bookmarks(doc)
    doc.save(out)
    try:
        import doc_meta
        doc_meta.clean(out)
    except Exception:
        pass
    return out


def make_oblozhki(donor, count):
    """Размножает обложку тома 1 на тома 2…count. Возвращает список созданных файлов."""
    import docx
    import docx_edit
    base = os.path.dirname(os.path.abspath(donor))
    name = os.path.basename(donor)
    doc_text = "\n".join(p.text for p in docx.Document(donor).paragraphs)
    if not re.search(r"ТОМ\s*№\s*1\b", doc_text, re.I):
        raise ValueError("в доноре нет метки «ТОМ № 1» — это не обложка тома 1")
    made = []
    for n in range(2, count + 1):
        new_name = re.sub(r"(?i)(том\s*)1\b", lambda m: m.group(1) + str(n), name, count=1)
        if new_name == name:
            root, ext = os.path.splitext(name)
            new_name = "%s том %d%s" % (root, n, ext)
        out = os.path.join(base, new_name)
        shutil.copyfile(donor, out)
        docx_edit.edit_file(out, [{"op": "replace", "old": "ТОМ № 1",
                                   "new": "ТОМ № %d" % n, "expect": 1}], backup=False)
        made.append(out)
    return made


def main():
    ap = argparse.ArgumentParser(description="Печатная опись тома и обложки на N томов")
    ap.add_argument("--delo", help="номер уголовного дела")
    ap.add_argument("--tom", type=int, help="номер тома")
    ap.add_argument("--items", help="файл с позициями описи («-» — stdin)")
    ap.add_argument("--out", help="куда сохранить опись")
    ap.add_argument("--donor", help="донор описи (.docx), обязателен при --items")
    ap.add_argument("--start", type=int, default=1, help="с какого номера нумеровать позиции")
    ap.add_argument("--oblozhka", help="донор обложки — файл «… том 1 ….docx»")
    ap.add_argument("--oblozhki", type=int, help="сколько всего томов (сделать обложки 2…N)")
    args = ap.parse_args()

    did = False
    if args.oblozhka or args.oblozhki:
        if not (args.oblozhka and args.oblozhki):
            print("Для обложек нужны оба флага: --oblozhka <файл тома 1> --oblozhki <N>")
            return 1
        if not os.path.isfile(args.oblozhka):
            print("Нет файла обложки-донора: %s" % args.oblozhka)
            return 1
        if args.oblozhki < 2:
            print("Томов меньше двух — размножать нечего.")
            return 1
        try:
            made = make_oblozhki(args.oblozhka, args.oblozhki)
        except Exception as e:
            print("Обложки не сделаны: %s" % e)
            return 1
        for p in made:
            print("обложка: %s" % p)
        did = True

    if args.items:
        if not args.tom:
            print("Не указан --tom (номер тома).")
            return 1
        if not args.donor or not os.path.isfile(args.donor):
            print("Нет донора описи: %s\nУкажи свой через --donor." % args.donor)
            return 1
        if args.items == "-":
            text = sys.stdin.read()
            base = os.getcwd()
        else:
            if not os.path.isfile(args.items):
                print("Нет файла с позициями: %s" % args.items)
                return 1
            with io.open(args.items, encoding="utf-8") as fh:
                text = fh.read()
            base = os.path.dirname(os.path.abspath(args.items))
        items = parse_items(text)
        if not items:
            print("Перечень пуст — описи не из чего собирать.")
            return 1
        out = args.out or os.path.join(base, "Опись том %d.docx" % args.tom)
        try:
            path = make_opis(args.donor, out, args.delo, args.tom, items, args.start)
        except Exception as e:
            print("Опись не собрана: %s" % e)
            return 1
        print("опись: %s (позиций: %d)" % (path, len(items)))
        did = True

    if not did:
        ap.print_help()
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
