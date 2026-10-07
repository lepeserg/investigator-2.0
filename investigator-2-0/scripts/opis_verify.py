# -*- coding: utf-8 -*-
"""opis_verify.py — КАРТА «СТРАНИЦА → ТИП ДОКУМЕНТА» по собранному тому PDF.

Зачем. Позиция описи «постановление о ВУД, сопроводительные письма и уведомление на 5 л.»
была названа неверно: пятый лист оказался РАПОРТОМ об обнаружении признаков преступления.
Наименование выводилось из соседних документов, а не из содержимого листа. Скрипт делает
обратное: читает каждую страницу и говорит, что на ней на самом деле, — опись строится по факту.

Что делает: рендерит ВЕРХНЮЮ ТРЕТЬ каждой страницы, прогоняет Tesseract (-l rus) и печатает
«страница → распознанный заголовок → предполагаемый тип документа». В конце — блоки подряд
идущих страниц одного типа: это и есть заготовка позиций описи.

⚠ OCR кириллицы врёт в словах, поэтому тип определяется по УСТОЙЧИВЫМ ЯКОРЯМ (ПОСТАНОВЛЕНИЕ,
ПРОТОКОЛ ДОПРОСА, РАПОРТ, УВЕДОМЛЕНИЕ, РАСПИСКА, СОПРОВОДИТЕЛЬНОЕ и др.) с допуском на опечатки
и латинские двойники букв. Что не опознано — помечается «?»: это НЕ значит «продолжение
предыдущего», это значит «посмотри лист сам».

ВХОД: PDF тома (сканы). ВЫХОД: таблица в stdout, при --csv — ещё и файл «страница;заголовок;тип».

Использование:
    python opis_verify.py \"Том 1 — материалы УД.pdf\"
    python opis_verify.py \"Том 1.pdf\" --pages 1-40          # только часть тома
    python opis_verify.py \"Том 1.pdf\" --csv karta_tom1.csv  # карта в файл
    python opis_verify.py \"Том 1.pdf\" --dpi 400 --part 0.4  # мелкий шрифт / высокая шапка
    python opis_verify.py \"Том 1.pdf\" --no-deep             # быстрее: только шапка

Скорость: около 1,5 с на лист (300 dpi + Tesseract), неопознанные листы читаются дважды.
Том на 150 листов — примерно 5 минут; на первый прогон берите --pages.

Код возврата: 0 — карта построена; 1 — есть неопознанные страницы (нужен глаз человека);
2 — ошибка вызова (нет файла, нет Tesseract).
Зависимости: pymupdf (fitz); Tesseract OCR с языковым пакетом rus (системная программа)."""

import argparse
import difflib
import os
import re
import shutil
import subprocess
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Латинские двойники кириллицы — Tesseract подставляет их постоянно.
LAT2CYR = {"A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "K": "К", "M": "М",
           "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У", "3": "З", "0": "О"}

# Якоря: (шаблон, тип). Порядок важен — первый сработавший выигрывает,
# поэтому частные формулировки стоят выше общих.
ANCHORS = [
    ("ПРОТОКОЛДОПОЛНИТЕЛЬНОГОДОПРОСА", "протокол дополнительного допроса"),
    ("ПРОТОКОЛДОПРОСА", "протокол допроса"),
    ("ПРОТОКОЛОЧНОЙСТАВКИ", "протокол очной ставки"),
    ("ПРОТОКОЛОСМОТРА", "протокол осмотра"),
    ("ПРОТОКОЛОБЫСКА", "протокол обыска"),
    ("ПРОТОКОЛВЫЕМКИ", "протокол выемки"),
    ("ПРОТОКОЛОЗНАКОМЛЕНИЯ", "протокол ознакомления"),
    ("ПРОТОКОЛПРЕДЪЯВЛЕНИЯ", "протокол предъявления"),
    ("ПРОТОКОЛЗАДЕРЖАНИЯ", "протокол задержания"),
    ("ПРОТОКОЛ", "протокол (какой — уточнить)"),
    ("ОБВИНИТЕЛЬНОЕЗАКЛЮЧЕНИЕ", "обвинительное заключение"),
    ("ПОСТАНОВЛЕНИЕ", "постановление"),
    ("РАПОРТ", "рапорт"),
    ("УВЕДОМЛЕНИЕ", "уведомление"),
    ("РАСПИСКА", "расписка"),
    ("СОПРОВОДИТЕЛЬНОЕПИСЬМО", "сопроводительное письмо"),
    ("ХОДАТАЙСТВО", "ходатайство"),
    ("ОБЪЯСНЕНИЕ", "объяснение"),
    ("ЗАКЛЮЧЕНИЕЭКСПЕРТА", "заключение эксперта"),
    ("ДОСУДЕБНОЕСОГЛАШЕНИЕ", "досудебное соглашение"),
    ("СПРАВКАМЕМОРАНДУМ", "справка-меморандум"),
    ("СПРАВКА", "справка"),
    ("ТРЕБОВАНИЕ", "требование"),
    ("ЗАПРОС", "запрос"),
    ("ОРДЕР", "ордер"),
    ("ПОДПИСКА", "подписка"),
    ("ОБЯЗАТЕЛЬСТВООЯВКЕ", "обязательство о явке"),
    # «ОПИСЬ» в якорях НЕ держим: слово короткое, а обрывок «(подпись)» на любом бланке
    # читается как «одпись» и давал ложную «опись». Внутренняя опись тома и так первый лист.
    ("ПОВЕСТКА", "повестка"),
    ("ПОРУЧЕНИЕ", "поручение"),
]


def find_tesseract():
    if shutil.which("tesseract"):
        return "tesseract"
    for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.isfile(c):
            return c
    return None


def normalize(s):
    """Верхний регистр, латиница → кириллица, прочь всё, кроме букв."""
    s = s.upper()
    s = "".join(LAT2CYR.get(ch, ch) for ch in s)
    return re.sub(r"[^А-ЯЁ]", "", s)


# Строки, которые НЕ являются заголовком, хотя начинаются похоже: резолютивные и вводные
# слова тела документа. «ПОСТАНОВИЛ:» на третьем листе иначе делает лист новым постановлением.
NOT_TITLE = ("ПОСТАНОВИЛ", "УСТАНОВИЛ", "ОПРЕДЕЛИЛ", "ПОСТАНОВЛЯЮ", "ПРИЛОЖЕНИЕ",
             "ПРОТОКОЛЬНО")


def head_match(hay, needle, cutoff):
    """Похожесть НАЧАЛА строки на якорь.

    Якорь сверяется только с НАЧАЛОМ строки — иначе типом становится любое упоминание слова
    в теле: строка «Копия настоящего постановления направлена военному прокурору» на третьем
    листе постановления делала этот лист новым «постановлением». Хвост окна на два знака длиннее
    якоря — этого хватает, чтобы пережить лишний знак, приклеенный OCR спереди.
    """
    n = len(needle)
    if n == 0 or len(hay) < n * 0.7:
        return 0.0
    sm = difflib.SequenceMatcher(autojunk=False)
    sm.set_seq2(needle)
    sm.set_seq1(hay[:n + 2])
    r = sm.ratio()
    return r if r >= cutoff else 0.0


def guess_type(lines, cutoff, max_extra=None):
    'Тип документа по якорям. Возвращает (тип, строка-заголовок, схожесть).'
    for ln in lines:
        norm = normalize(ln)
        if any(head_match(norm, stop, 0.86) for stop in NOT_TITLE):
            continue
        for anchor, kind in ANCHORS:      # частные формулировки стоят выше общих
            if max_extra is not None and len(norm) > len(anchor) + max_extra:
                continue
            score = head_match(norm, anchor, cutoff)
            if score:
                return kind, ln, score
    return None, None, 0.0


def ocr_top(tess, doc, index, dpi, part, tmp):
    """OCR верхней части страницы. Возвращает все непустые строки сверху вниз."""
    import fitz
    page = doc.load_page(index)
    r = page.rect
    clip = fitz.Rect(r.x0, r.y0, r.x1, r.y0 + r.height * part)
    png = os.path.join(tmp, "p%05d.png" % (index + 1))
    page.get_pixmap(dpi=dpi, clip=clip).save(png)
    res = subprocess.run([tess, png, "stdout", "-l", "rus", "--psm", "6"],
                         capture_output=True)
    try:
        os.remove(png)
    except OSError:
        pass
    out = []
    for ln in res.stdout.decode("utf-8", "replace").splitlines():
        ln = " ".join(ln.split())
        if len(ln) >= 3:
            out.append(ln)
    return out


# Насколько глубоко опускаться, если в шапке ничего не опознано. Проверено на реальном томе:
# на бланке РАПОРТА сверху лежит блок резолюций руководителя, и слово «РАПОРТ» оказывается
# десятой строкой и ниже трети листа — на пяти верхних строках верхней трети лист не опознаётся.
DEEP_PART = 0.6


def read_page_type(tess, doc, index, args, tmp):
    """Тип листа: сначала шапка, потом — глубже. Возвращает (тип, заголовок, схожесть, пометка)."""
    lines = ocr_top(tess, doc, index, args.dpi, args.part, tmp)
    kind, head, score = guess_type(lines[:args.lines], args.cutoff)
    if kind or args.no_deep:
        return kind, head, score, "", lines
    deep_cut, deep_extra = min(args.cutoff + 0.04, 0.98), 10
    kind, head, score = guess_type(lines, deep_cut, deep_extra)
    if kind:
        return kind, head, score, " (ниже %d-й строки)" % args.lines, lines
    if args.part < DEEP_PART:
        deep = ocr_top(tess, doc, index, args.dpi, DEEP_PART, tmp)
        kind, head, score = guess_type(deep, deep_cut, deep_extra)
        if kind:
            return kind, head, score, " (ниже шапки)", deep
    return None, None, 0.0, "", lines


def parse_pages(spec, total):
    if not spec:
        return list(range(1, total + 1))
    pages = []
    for part in spec.replace(" ", "").split(","):
        m = re.match(r"^(\d+)(?:[-–](\d+))?$", part)
        if not m:
            raise ValueError("непонятный диапазон страниц: %s" % part)
        a = int(m.group(1))
        b = int(m.group(2) or a)
        pages += [p for p in range(a, b + 1) if 1 <= p <= total]
    return pages


def main():
    ap = argparse.ArgumentParser(
        prog="opis_verify.py",
        description="Карта «страница → заголовок → тип документа» по PDF тома: "
                    "опись строится по содержимому листа, а не по соседям.",
        epilog="Код возврата: 0 — все страницы опознаны; 1 — есть «?»; 2 — ошибка вызова.")
    ap.add_argument("pdf", help="PDF собранного тома")
    ap.add_argument("--pages", help="страницы: «1-40» или «1,5,7-9» (по умолчанию все)")
    ap.add_argument("--dpi", type=int, default=300, help="разрешение рендера (по умолчанию 300)")
    ap.add_argument("--part", type=float, default=0.33,
                    help="доля высоты страницы сверху (по умолчанию 0.33 — верхняя треть)")
    ap.add_argument("--lines", type=int, default=5,
                    help="сколько первых строк шапки разбирать (по умолчанию 5)")
    ap.add_argument("--no-deep", action="store_true",
                    help="не опускаться ниже шапки, если тип не опознан (быстрее, но грубее)")
    ap.add_argument("--cutoff", type=float, default=0.82,
                    help="порог похожести на якорь, 0..1 (по умолчанию 0.82)")
    ap.add_argument("--csv", help="выгрузить карту в файл «страница;заголовок;тип»")
    args = ap.parse_args()

    if not os.path.isfile(args.pdf):
        sys.stderr.write("нет файла: %s\n" % args.pdf)
        return 2
    tess = find_tesseract()
    if not tess:
        sys.stderr.write("не найден tesseract — поставьте Tesseract OCR с пакетом rus "
                         "(см. check_env.py)\n")
        return 2
    try:
        import fitz
    except ImportError:
        sys.stderr.write("нужен pymupdf (pip install pymupdf)\n")
        return 2

    doc = fitz.open(args.pdf)
    try:
        pages = parse_pages(args.pages, doc.page_count)
    except ValueError as e:
        sys.stderr.write("%s\n" % e)
        doc.close()
        return 2

    print("Файл: %s   ·   страниц в PDF: %d   ·   разбираю: %d"
          % (os.path.basename(args.pdf), doc.page_count, len(pages)))
    print("Верхняя часть листа: %d%%   ·   %d dpi   ·   OCR: rus" % (args.part * 100, args.dpi))
    print("-" * 100)

    tmp = tempfile.mkdtemp(prefix="opis_verify_")
    rows = []
    try:
        for p in pages:
            kind, head, score, note, lines = read_page_type(tess, doc, p - 1, args, tmp)
            if head is None:
                head = lines[0] if lines else "(пусто — лист не распознан)"
            rows.append((p, head, kind, score, note))
            print("стр. %4d | %-58s | %s%s" % (p, head[:58], kind or "?", note))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
        doc.close()

    unknown = [r[0] for r in rows if not r[2]]
    print("-" * 100)
    print("БЛОКИ ПОДРЯД ИДУЩИХ СТРАНИЦ ОДНОГО ТИПА (заготовка позиций описи):")
    blocks = []
    for p, head, kind, score, note in rows:
        key = kind or "?"
        if blocks and blocks[-1][2] == key and p == blocks[-1][1] + 1:
            blocks[-1][1] = p
        else:
            blocks.append([p, p, key])
    for a, b, key in blocks:
        print("  стр. %d–%d  (%d л.)  %s" % (a, b, b - a + 1, key))
    if unknown:
        print("\n⚠ Не опознано (посмотрите лист сам): стр. %s"
              % ", ".join(str(u) for u in unknown))

    if args.csv:
        with open(args.csv, "w", encoding="utf-8-sig", newline="") as f:
            f.write("страница;заголовок;тип;схожесть\n")
            for p, head, kind, score, note in rows:
                f.write("%d;%s;%s;%.2f\n"
                        % (p, head.replace(";", ","), (kind or "?") + note, score))
        print("\nКарта выгружена: %s" % args.csv)

    return 1 if unknown else 0


if __name__ == "__main__":
    sys.exit(main())
