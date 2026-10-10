# -*- coding: utf-8 -*-
"""check_tom.py — СВЕРКА ПЕРЕЧНЯ, ОПИСИ И СОБРАННЫХ ТОМОВ выделенного уголовного дела.

Что проверяет:
  1. сквозная нумерация позиций перечня — без разрывов и без повторов;
  2. непрерывность листов дела (л.д.) — без дыр и без наложений;
  3. «на N л.» каждой позиции = ширине её диапазона л.д.;
  4. сумма «на N л.» по перечню = последнему листу описи = числу страниц PDF тома;
  5. ⛔ ни одна страница источника (том T, л. X–Y) не использована дважды;
  6. нет повторяющихся наименований позиций (в описи их не различить);
  7. формат дат ДД.ММ.ГГГГ и короткое тире «–» в диапазонах.

ВХОД:
  --perechen — .txt описи («  1) наименование … на 3 л.» + строка
               «      л.д. 1–3   |   источник: том 1, л. 1–3»)
               ЛИБО .docx постановления о выделении (позиции «1) … на 3 л.;» абзацами);
  --toma     — папка с PDF томов либо сами файлы (том определяется по «Том N» в имени).

ВЫХОД: таблица расхождений в stdout. Код возврата: 0 — всё сходится; 1 — есть расхождения;
2 — ошибка вызова (нет файла и т. п.).

ПЛАН ТОМА (--plan-save / --plan-check / --map / --izyat)
Владелец печатает один том, пока правится соседний. Чтобы не переписывать напечатанное,
рядом с PDF кладётся план тома `<имя тома>.plan.json` — перечень «лист → источник → поворот».
  · --plan-save  — записать план по текущей описи (поворот берётся из PDF);
  · --plan-check — сличить текущую опись с сохранённым планом и назвать тома, которые
                   ДЕЙСТВИТЕЛЬНО надо переписать (состав или поворот изменились);
                   том, открытый в просмотрщике, помечается «занят» — не падаем;
  · --map        — карта тома: страница/лист → позиция описи → источник
                   («листы с 44 по 50 лишние» — какие это позиции);
  · --izyat      — что изъять из УЖЕ НАПЕЧАТАННОГО тома: диапазоны л.д. печатной версии,
                   считанные по сохранённому плану, а не по памяти.

Зависимости: pymupdf (fitz) — только когда задан --toma; python-docx — только для .docx
(есть zip-фолбэк без библиотеки)."""

import argparse
import datetime
import json
import os
import re
import sys
import zipfile

from word_text import xml_text

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DASH = "–"                      # короткое тире, оно же «–»
DASHES = "–—−-"       # любое тире/дефис/минус — как их пишут в диапазонах

# «  12) наименование … на 3 л.»
POS_RE = re.compile(r"^\s{0,6}(\d{1,4})\)\s+(.*)$")
# «      л.д. 23–28   |   источник: том 1, л. 29–34» (источник необязателен)
LD_RE = re.compile(
    r"л\.\s?д\.\s*(\d{1,4})\s*[%s]\s*(\d{1,4})"
    r"(?:.*?источник:\s*том\s*(\d{1,3})\s*,\s*л\.\s*(\d{1,4})\s*[%s]\s*(\d{1,4}))?" % (DASHES, DASHES)
)
TOM_HEAD_RE = re.compile(r"^\s*Том\s+(\d{1,3})\b")
NA_L_RE = re.compile(r"на\s+(\d{1,4})\s+л\.")
TOTAL_RE = re.compile(r"ВСЕГО:\s*(\d{1,4})\s*позиц\w*,\s*(\d{1,5})\s*лист")
# Дата-кандидат: разделители ОДИНАКОВЫЕ и вокруг нет цифр/точек/дробей — иначе в кандидаты
# Локальный пример исключён из публичной поставки.
DATE_ANY_RE = re.compile(r"(?<![\d.\-/])(\d{1,2})([.\-/])(\d{1,2})\2(\d{2,4})(?![\d.\-/])")
DATE_OK_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
MONTHS = ("январ", "феврал", "март", "апрел", "мая", "май", "июн", "июл",
          "август", "сентябр", "октябр", "ноябр", "декабр")
DATE_WORD_RE = re.compile(r"\b\d{1,2}\s+(%s)\w*\s+\d{4}" % "|".join(MONTHS), re.I)
BAD_RANGE_RE = re.compile(r"\d\s*[—−]\s*\d")   # длинное тире/минус между цифрами


class Pos(object):
    """Одна позиция перечня."""

    def __init__(self, num, name, line, tom):
        self.num = num          # номер позиции
        self.name = name        # наименование
        self.line = line        # строка/абзац источника — для сообщений
        self.tom = tom          # том, объявленный заголовком описи (или None)
        self.deklar = None      # «на N л.»
        self.ld = None          # (начало, конец) листов дела
        self.src = None         # (том, начало, конец) страниц источника


def _norm(s):
    return re.sub(r"\s+", " ", s).strip().strip(";.").lower()


def _read_text(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:3] == b"\xef\xbb\xbf":
        raw = raw[3:]
    return raw.decode("utf-8", "replace")


def _docx_paragraphs(path):
    """Абзацы .docx. Сначала python-docx, при его отсутствии — разбор XML из zip."""
    try:
        import docx
        return [p.text for p in docx.Document(path).paragraphs]
    except ImportError:
        pass
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
    return _xml_paragraphs(xml)


def _xml_paragraphs(xml):
    """Тексты абзацев из document.xml — общими правилами word_text.xml_text
    (удалённое при рецензировании и коды полей не входят; <w:tab/> здесь не пишется)."""
    return xml_text(xml, paragraphs=True, tab="")


def parse_perechen(path):
    """Разбирает опись .txt либо постановление .docx. Возвращает (позиции, итог, вид)."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".docx":
        lines = _docx_paragraphs(path)
        kind = "постановление"
    else:
        lines = _read_text(path).splitlines()
        kind = "опись"

    positions, total, tom = [], None, None
    for i, raw in enumerate(lines, 1):
        line = raw.replace("\t", " ").rstrip()
        m = TOM_HEAD_RE.match(line)
        if m:
            tom = int(m.group(1))
            continue
        m = TOTAL_RE.search(line)
        if m:
            total = (int(m.group(1)), int(m.group(2)))
            continue
        m = POS_RE.match(line)
        if m and NA_L_RE.search(m.group(2)):
            # в постановлении номер позиции идёт от «1)», а «2. Выделенному делу…»
            # резолютивной части сюда не попадёт — там нет «на N л.»
            positions.append(Pos(int(m.group(1)), m.group(2).strip(), i, tom))
            continue
        m = LD_RE.search(line)
        if m and positions:
            p = positions[-1]
            p.ld = (int(m.group(1)), int(m.group(2)))
            if m.group(3):
                p.src = (int(m.group(3)), int(m.group(4)), int(m.group(5)))
            continue
        # продолжение длинного наименования (перенос строки внутри позиции)
        if positions and line.strip() and not line.startswith(("=", "#", "-")) \
                and positions[-1].ld is None and not POS_RE.match(line) \
                and kind == "опись" and line.startswith("      "):
            positions[-1].name += " " + line.strip()

    for p in positions:
        nums = NA_L_RE.findall(p.name)
        if nums:
            p.deklar = int(nums[-1])      # «на N л.» — последнее в наименовании
    return positions, total, kind


def load_toma(paths):
    """Возвращает [(номер тома, имя файла, страниц, полный путь)] по PDF, по возрастанию тома."""
    import pymupdf as fitz
    files = []
    for p in paths:
        if os.path.isdir(p):
            files += [os.path.join(p, f) for f in sorted(os.listdir(p))
                      if f.lower().endswith(".pdf")]
        elif os.path.isfile(p):
            files.append(p)
    out = []
    for f in files:
        m = re.search(r"[Тт]ом\s*(\d{1,3})", os.path.basename(f))
        num = int(m.group(1)) if m else len(out) + 1
        doc = fitz.open(f)
        out.append((num, os.path.basename(f), doc.page_count, os.path.abspath(f)))
        doc.close()
    out.sort(key=lambda t: t[0])
    return out


def detect_numbering(positions):
    """«potomno» — в каждом томе листы с 1; «skvoznaya» — общий счёт через все тома."""
    starts = {}
    for p in positions:
        if p.ld and p.tom is not None and p.tom not in starts:
            starts[p.tom] = p.ld[0]
    if len(starts) < 2:
        return "skvoznaya"
    return "potomno" if all(v == 1 for v in starts.values()) else "skvoznaya"


def check(positions, total, kind, toma, numbering):
    """Возвращает список расхождений: (раздел проверки, текст)."""
    bad = []

    # 1. Нумерация позиций
    if not positions:
        bad.append(("перечень", "не нашёл ни одной позиции вида «1) … на N л.» — "
                                "проверьте, тот ли это файл"))
        return bad
    seen = {}
    for p in positions:
        if p.num in seen:
            bad.append(("нумерация", "позиция № %d встречается дважды (строки %d и %d)"
                        % (p.num, seen[p.num], p.line)))
        seen[p.num] = p.line
    expect = 1
    for p in positions:
        if p.num != expect:
            if p.num > expect:
                bad.append(("нумерация", "разрыв нумерации: после № %d идёт № %d "
                            "(пропущено %d)" % (expect - 1, p.num, p.num - expect)))
            else:
                bad.append(("нумерация", "нумерация пошла назад: после № %d идёт № %d"
                            % (expect - 1, p.num)))
            expect = p.num
        expect += 1

    # 2. «на N л.» у каждой позиции
    for p in positions:
        if p.deklar is None:
            bad.append(("объём", "№ %d: не указано «на N л.» — %s" % (p.num, p.name[:70])))

    # 3-4. Листы дела
    with_ld = [p for p in positions if p.ld]
    if with_ld:
        prev_end, prev_tom, prev_num = 0, with_ld[0].tom, None
        for p in with_ld:
            if numbering == "potomno" and p.tom != prev_tom:
                prev_end, prev_tom = 0, p.tom
            if p.ld[0] != prev_end + 1:
                what = "дыра" if p.ld[0] > prev_end + 1 else "наложение"
                bad.append(("л.д.", "%s в листах дела: № %d начинается с л.д. %d, "
                            "а предыдущая (№ %s) кончилась на %d"
                            % (what, p.num, p.ld[0], prev_num, prev_end)))
            if p.ld[1] < p.ld[0]:
                bad.append(("л.д.", "№ %d: диапазон л.д. %d%s%d вывернут"
                            % (p.num, p.ld[0], DASH, p.ld[1])))
            width = p.ld[1] - p.ld[0] + 1
            if p.deklar is not None and width != p.deklar:
                bad.append(("объём", "№ %d: «на %d л.», а л.д. %d%s%d — это %d л."
                            % (p.num, p.deklar, p.ld[0], DASH, p.ld[1], width)))
            prev_end, prev_num = max(prev_end, p.ld[1]), p.num
    elif kind == "опись":
        bad.append(("л.д.", "в описи нет ни одной строки «л.д. A–B» — "
                            "непрерывность листов проверить нечем"))

    # 5. Ни одна страница источника не использована дважды
    used = {}
    for p in positions:
        if not p.src:
            continue
        t, a, b = p.src
        if b < a:
            bad.append(("источник", "№ %d: диапазон источника том %d, л. %d%s%d вывернут"
                        % (p.num, t, a, DASH, b)))
            a, b = b, a
        if p.deklar is not None and (b - a + 1) != p.deklar:
            bad.append(("источник", "№ %d: «на %d л.», а взято из источника %d л. "
                        "(том %d, л. %d%s%d)" % (p.num, p.deklar, b - a + 1, t, a, DASH, b)))
        for page in range(a, b + 1):
            key = (t, page)
            if key in used:
                bad.append(("источник", "⛔ том %d, л. %d использован дважды: "
                            "в позициях № %d и № %d" % (t, page, used[key], p.num)))
            else:
                used[key] = p.num

    # 6. Итог: сумма «на N л.» = последний л.д. = страницы PDF
    summa = sum(p.deklar for p in positions if p.deklar is not None)
    per_tom_pos = {}
    for p in positions:
        if p.deklar is not None and p.tom is not None:
            per_tom_pos[p.tom] = per_tom_pos.get(p.tom, 0) + p.deklar
    if with_ld:
        last_ld = max(p.ld[1] for p in with_ld) if numbering == "skvoznaya" else None
        if numbering == "skvoznaya" and last_ld != summa:
            bad.append(("итог", "сумма «на N л.» по перечню = %d, а последний лист описи = %d"
                        % (summa, last_ld)))
        if numbering == "potomno":
            for t in sorted(per_tom_pos):
                lds = [p.ld[1] for p in with_ld if p.tom == t]
                if lds and max(lds) != per_tom_pos[t]:
                    bad.append(("итог", "том %d: сумма «на N л.» = %d, "
                                "а последний лист описи тома = %d" % (t, per_tom_pos[t], max(lds))))
    if total:
        if total[0] != len(positions):
            bad.append(("итог", "строка «ВСЕГО» обещает %d позиций, в перечне их %d"
                        % (total[0], len(positions))))
        if total[1] != summa:
            bad.append(("итог", "строка «ВСЕГО» обещает %d л., сумма «на N л.» = %d"
                        % (total[1], summa)))

    if toma:
        pages_all = sum(t[2] for t in toma)
        if pages_all != summa:
            bad.append(("тома", "в PDF томов %d стр., а сумма «на N л.» по перечню = %d "
                        "(разница %+d)" % (pages_all, summa, pages_all - summa)))
        for num, fname, pages, _path in toma:
            if num in per_tom_pos and per_tom_pos[num] != pages:
                bad.append(("тома", "том %d: в PDF %d стр., а по перечню %d л. (%s)"
                            % (num, pages, per_tom_pos[num], fname)))
        if per_tom_pos:
            missing = sorted(set(per_tom_pos) - {t[0] for t in toma})
            if missing:
                bad.append(("тома", "в перечне есть тома %s, а PDF на них нет"
                            % ", ".join(str(m) for m in missing)))

    # 7. Повторяющиеся наименования
    names = {}
    for p in positions:
        key = _norm(p.name)
        if key in names:
            bad.append(("повторы", "№ %d и № %d названы одинаково — в описи их не различить: %s"
                        % (names[key], p.num, p.name[:90])))
        else:
            names[key] = p.num

    # 8. Формат дат и тире
    for p in positions:
        for m in DATE_ANY_RE.finditer(p.name):
            if not DATE_OK_RE.match(m.group(0)):
                bad.append(("формат", "№ %d: дата «%s» — не ДД.ММ.ГГГГ" % (p.num, m.group(0))))
        for m in DATE_WORD_RE.finditer(p.name):
            bad.append(("формат", "№ %d: дата прописью «%s» — нужна ДД.ММ.ГГГГ"
                        % (p.num, m.group(0))))
        m = BAD_RANGE_RE.search(p.name)
        if m:
            bad.append(("формат", "№ %d: в диапазоне «%s» длинное тире — нужно короткое «%s»"
                        % (p.num, m.group(0), DASH)))
    return bad


# ======================================================================================
# ПЛАН ТОМА: «лист → позиция описи → источник → поворот»
# ======================================================================================

PLAN_VERSION = 1
PLAN_SUFFIX = ".plan.json"


def tom_spans(positions, numbering):
    """{номер тома: (первый л.д., последний л.д., [позиции])}. Тома без л.д. пропускаются."""
    spans = {}
    for p in positions:
        if not p.ld:
            continue
        t = p.tom if p.tom is not None else 1
        if t not in spans:
            spans[t] = [p.ld[0], p.ld[1], []]
        spans[t][0] = min(spans[t][0], p.ld[0])
        spans[t][1] = max(spans[t][1], p.ld[1])
        spans[t][2].append(p)
    if numbering == "potomno":
        for t in spans:
            spans[t][0] = 1
    return dict((t, (v[0], v[1], v[2])) for t, v in spans.items())


def page_rotations(path):
    """Повороты страниц PDF: [0, 90, …]. Файл недоступен — None (не падаем)."""
    try:
        import pymupdf as fitz
        doc = fitz.open(path)
        rots = [doc[i].rotation for i in range(doc.page_count)]
        doc.close()
        return rots
    except Exception as e:
        print("⚠ повороты страниц не прочитаны, сверка поворотов пропущена: %s (%s)" % (path, e),
              file=sys.stderr)
        return None


def is_busy(path):
    """Файл занят другим приложением (том открыт в просмотрщике)?"""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "r+b"):
            return False
    except PermissionError:
        return True
    except OSError:
        return True


def build_plan(tom, positions, numbering, perechen, fname=None, pages=None, rots=None):
    """План одного тома: по листу на каждый лист дела."""
    spans = tom_spans(positions, numbering)
    if tom not in spans:
        return None
    first, last, poss = spans[tom]
    sheets = []
    for p in sorted(poss, key=lambda q: q.ld[0]):
        for i, ld in enumerate(range(p.ld[0], p.ld[1] + 1)):
            page = ld - first + 1
            sheets.append({
                "page": page,
                "ld": ld,
                "pos": p.num,
                "name": p.name,
                "src_tom": p.src[0] if p.src else None,
                "src_page": (p.src[1] + i) if p.src else None,
                "rot": (rots[page - 1] if rots and 0 < page <= len(rots) else None),
            })
    return {
        "version": PLAN_VERSION,
        "created": datetime.datetime.now().replace(microsecond=0).isoformat(),
        "perechen": os.path.basename(perechen),
        "tom": tom,
        "file": fname,
        "numbering": numbering,
        "first_ld": first,
        "last_ld": last,
        "sheets_count": len(sheets),
        "pdf_pages": pages,
        "sheets": sheets,
    }


def plan_key(plan):
    """Ключ переписывания PDF: последовательность «источник → поворот». Имена не в счёт."""
    return [(s.get("src_tom"), s.get("src_page"), s.get("rot")) for s in plan["sheets"]]


def plan_names(plan):
    return [(s.get("pos"), s.get("name")) for s in plan["sheets"]]


def plan_file(tom, fname, plan_dir, toma_dir):
    base = (fname or ("Том %d" % tom)) + PLAN_SUFFIX
    return os.path.normpath(os.path.join(plan_dir or toma_dir or ".", base))


def load_plan(path):
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:3] == b"\xef\xbb\xbf":
        raw = raw[3:]
    return json.loads(raw.decode("utf-8"))


def save_plan(path, plan):
    """Пишет план. Занят — сообщает, а не падает. Возвращает None либо текст ошибки."""
    try:
        with open(path, "w", encoding="utf-8", newline="\n") as f:
            json.dump(plan, f, ensure_ascii=False, indent=1)
            f.write("\n")
        return None
    except PermissionError:
        return "файл плана занят другим приложением: %s" % path
    except OSError as e:
        return "не записать план %s: %s" % (path, e)


def compact(nums):
    """[1,2,3,7,9,10] → «1–3, 7, 9–10»."""
    out, nums = [], sorted(set(nums))
    i = 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        out.append(str(nums[i]) if i == j else "%d%s%d" % (nums[i], DASH, nums[j]))
        i = j + 1
    return ", ".join(out) if out else "—"


def plan_diff(old, new):
    """(надо ли переписывать PDF, [пояснения])."""
    ko, kn = plan_key(old), plan_key(new)
    if ko == kn:
        why = []
        if plan_names(old) != plan_names(new):
            why.append("состав тот же, изменились только наименования/нумерация позиций "
                       "— PDF переписывать не нужно, переписать опись")
        return False, why
    so = set((s.get("src_tom"), s.get("src_page")) for s in old["sheets"])
    sn = set((s.get("src_tom"), s.get("src_page")) for s in new["sheets"])
    why = []
    ubrano, dobavleno = so - sn, sn - so
    if ubrano:
        why.append("убрано %d л. источника" % len(ubrano))
    if dobavleno:
        why.append("добавлено %d л. источника" % len(dobavleno))
    rot_old = dict(((s.get("src_tom"), s.get("src_page")), s.get("rot")) for s in old["sheets"])
    rot_new = dict(((s.get("src_tom"), s.get("src_page")), s.get("rot")) for s in new["sheets"])
    rot = [k for k in set(rot_old) & set(rot_new)
           if rot_old[k] is not None and rot_new[k] is not None and rot_old[k] != rot_new[k]]
    if rot:
        why.append("изменён поворот у %d л." % len(rot))
    if not ubrano and not dobavleno and not rot:
        why.append("тот же состав, изменился ПОРЯДОК листов")
    why.append("листов было %d, стало %d" % (len(old["sheets"]), len(new["sheets"])))
    return True, why


SPEC_RE = re.compile(r"^(?:(?:т|t|том)\s*(\d{1,3})\s*[:.])?\s*(\d{1,5})(?:\s*[%s]\s*(\d{1,5}))?$"
                     % DASHES, re.I)


def parse_spec(spec):
    """«44-50» → (None, 44, 50) — листы дела; «т2:44-50» → (2, 44, 50) — страницы тома 2."""
    m = SPEC_RE.match(spec.strip())
    if not m:
        return None
    tom = int(m.group(1)) if m.group(1) else None
    a = int(m.group(2))
    b = int(m.group(3)) if m.group(3) else a
    return (tom, a, b) if a <= b else (tom, b, a)


def cmd_map(specs, positions, numbering, out):
    """Карта тома: страница/лист → позиция описи → источник."""
    spans = tom_spans(positions, numbering)
    plans = dict((t, build_plan(t, positions, numbering, "—")) for t in spans)
    problems = []
    for spec in specs:
        parsed = parse_spec(spec)
        if not parsed:
            problems.append(("карта", "не разобрал «%s» — пишите «44-50», «142» или «т2:1-3»" % spec))
            continue
        tom, a, b = parsed
        rows = []
        for t in sorted(plans):
            for s in plans[t]["sheets"]:
                hit = (s["page"] if tom is not None else s["ld"])
                if (tom is None or t == tom) and a <= hit <= b:
                    rows.append((t, s))
        out.append("")
        out.append("КАРТА «%s» — %s" % (
            spec, ("страницы тома %d" % tom) if tom is not None else "листы дела"))
        if not rows:
            out.append("  · ничего не нашлось: проверьте номер тома и диапазон")
            problems.append(("карта", "по «%s» в описи ничего нет" % spec))
            continue
        cur = prev = None
        for t, s in rows:
            if cur is None or s["pos"] != cur[1]["pos"] or t != cur[0]:
                if cur:
                    out.append(_map_row(cur, prev))
                cur = (t, s)
            prev = (t, s)
        out.append(_map_row(cur, prev))
    return problems


def _map_row(first, last):
    t, s = first
    _t2, s2 = last
    ld = "%d" % s["ld"] if s["ld"] == s2["ld"] else "%d%s%d" % (s["ld"], DASH, s2["ld"])
    pg = "%d" % s["page"] if s["page"] == s2["page"] else "%d%s%d" % (s["page"], DASH, s2["page"])
    src = "—"
    if s["src_tom"] is not None:
        src = "том %d, л. %s" % (s["src_tom"], "%d" % s["src_page"] if s["src_page"] == s2["src_page"]
                                 else "%d%s%d" % (s["src_page"], DASH, s2["src_page"]))
    return ("  л.д. %s  (том %d, стр. %s)  →  поз. № %d  %s\n      источник: %s"
            % (ld, t, pg, s["pos"], s["name"][:95], src))


def cmd_izyat(tom, old, new, out):
    """Что изъять из напечатанного тома — в л.д. НАПЕЧАТАННОЙ версии."""
    old_keys = [(s.get("src_tom"), s.get("src_page")) for s in old["sheets"]]
    new_keys = set((s.get("src_tom"), s.get("src_page")) for s in new["sheets"])
    izyat = [s["ld"] for s, k in zip(old["sheets"], old_keys)
             if k[1] is None or k not in new_keys]
    old_set = set(old_keys)
    vlozhit = [s for s in new["sheets"]
               if s.get("src_page") is None or (s.get("src_tom"), s.get("src_page")) not in old_set]
    out.append("")
    out.append("ТОМ %d — ИЗЪЯТИЕ ИЗ НАПЕЧАТАННОЙ ВЕРСИИ" % tom)
    out.append("  напечатано %d л.  ·  в новом составе %d л."
               % (len(old["sheets"]), len(new["sheets"])))
    out.append("  ИЗЪЯТЬ (л.д. напечатанной версии): %s" % compact(izyat))
    out.append("  всего к изъятию: %d л.   ·   остаётся из напечатанного: %d л."
               % (len(izyat), len(old["sheets"]) - len(izyat)))
    if vlozhit:
        by_pos = {}
        for s in vlozhit:
            by_pos.setdefault(s["pos"], []).append(s)
        out.append("  ВЛОЖИТЬ (нет в напечатанном): %d л." % len(vlozhit))
        for num in sorted(by_pos):
            g = by_pos[num]
            src = ("том %d, л. %s" % (g[0]["src_tom"],
                   compact([x["src_page"] for x in g])) if g[0]["src_tom"] is not None else "—")
            out.append("      поз. № %d  %s — %d л. (источник: %s)"
                       % (num, g[0]["name"][:80], len(g), src))
    else:
        out.append("  ВЛОЖИТЬ: нечего")
    out.append("  ПРОВЕРКА: %d − %d + %d = %d л. (столько должно остаться в томе)"
               % (len(old["sheets"]), len(izyat), len(vlozhit),
                  len(old["sheets"]) - len(izyat) + len(vlozhit)))
    return izyat, vlozhit


def run_plan_modes(args, positions, numbering, toma, perechen, out):
    """--plan-save / --plan-check / --izyat. Возвращает список расхождений."""
    problems = []
    toma_dir = None
    for p in args.toma:
        toma_dir = p if os.path.isdir(p) else os.path.dirname(os.path.abspath(p))
        break
    if not toma_dir and not args.plan_dir:
        out.append("  ⛔ не задано, где лежат планы: нужен --toma (папка с PDF) либо --plan-dir")
        return [("план", "не задано, где лежат планы томов: нужен --toma либо --plan-dir")]
    spans = tom_spans(positions, numbering)
    by_num = dict((t[0], t) for t in toma)
    nums = sorted(set(list(spans.keys()) + list(by_num.keys())))

    perepisat, zanyaty = [], []
    for tom in nums:
        entry = by_num.get(tom)
        fname = entry[1] if entry else None
        pages = entry[2] if entry else None
        path = entry[3] if entry else None
        rots = page_rotations(path) if (path and (args.plan_save or args.plan_check
                                                  or args.izyat)) else None
        new = build_plan(tom, positions, numbering, perechen, fname, pages, rots)
        if new is None:
            problems.append(("план", "том %d: в описи нет ни одной позиции с «л.д. A%sB»"
                             % (tom, DASH)))
            continue
        pf = plan_file(tom, fname, args.plan_dir, toma_dir)
        old = None
        if os.path.isfile(pf):
            try:
                old = load_plan(pf)
            except Exception as e:
                problems.append(("план", "том %d: план не читается (%s): %s" % (tom, e, pf)))

        if args.plan_check or args.izyat:
            if old is None:
                problems.append(("план", "том %d: сохранённого плана нет — сначала "
                                 "--plan-save на напечатанной версии (%s)" % (tom, pf)))
            elif args.plan_check:
                changed, why = plan_diff(old, new)
                busy = is_busy(path) if path else False
                mark = "ИЗМЕНИЛСЯ — переписать" if changed else "не изменился — не трогать"
                out.append("  том %d: %s%s" % (tom, mark, ("   [" + "; ".join(why) + "]")
                                               if why else ""))
                if changed:
                    perepisat.append(tom)
                    if busy:
                        zanyaty.append(tom)
                        out.append("      ⚠ файл занят (том открыт в просмотрщике) — "
                                   "переписать нельзя, закройте файл: %s" % (fname or path))
            if args.izyat and old is not None:
                cmd_izyat(tom, old, new, out)

        if args.plan_save:
            err = save_plan(pf, new)
            if err:
                problems.append(("план", "том %d: %s" % (tom, err)))
            else:
                out.append("  том %d: план сохранён (%d л.) — %s"
                           % (tom, new["sheets_count"], os.path.basename(pf)))

    if args.plan_check:
        out.append("")
        out.append("ПЕРЕПИСАТЬ ТОМА: %s" % (", ".join(str(t) for t in perepisat)
                                            if perepisat else "—"))
        if zanyaty:
            out.append("ЗАНЯТЫ (закрыть просмотрщик): %s"
                       % ", ".join(str(t) for t in zanyaty))
    return problems


def main():
    ap = argparse.ArgumentParser(
        prog="check_tom.py",
        description="Сверка перечня выделенного дела, описи и собранных PDF томов.",
        epilog="Код возврата: 0 — всё сходится, 1 — есть расхождения, 2 — ошибка вызова.")
    ap.add_argument("--perechen", required=True,
                    help="опись .txt либо постановление о выделении .docx")
    ap.add_argument("--toma", nargs="*", default=[],
                    help="папка с PDF томов либо сами файлы")
    ap.add_argument("--numbering", choices=("auto", "skvoznaya", "potomno"), default="auto",
                    help="нумерация листов: сквозная через все тома или в каждом томе с 1")
    ap.add_argument("--quiet", action="store_true", help="печатать только расхождения")
    ap.add_argument("--plan-save", action="store_true",
                    help="записать план тома (лист → источник → поворот) рядом с PDF")
    ap.add_argument("--plan-check", action="store_true",
                    help="сличить опись с сохранённым планом и назвать тома к переписыванию")
    ap.add_argument("--plan-dir", default=None,
                    help="папка для файлов плана (по умолчанию — рядом с PDF тома)")
    ap.add_argument("--map", nargs="+", metavar="ДИАПАЗОН", default=[],
                    help="карта тома: «44-50» (листы дела), «т2:1-3» (страницы тома 2), «142»")
    ap.add_argument("--izyat", action="store_true",
                    help="что изъять из напечатанного тома — по сохранённому плану")
    args = ap.parse_args()

    if not os.path.isfile(args.perechen):
        sys.stderr.write("нет файла перечня: %s\n" % args.perechen)
        return 2
    positions, total, kind = parse_perechen(args.perechen)
    try:
        toma = load_toma(args.toma) if args.toma else []
    except ImportError:
        sys.stderr.write("для сверки с томами нужен pymupdf (pip install pymupdf)\n")
        return 2

    numbering = detect_numbering(positions) if args.numbering == "auto" else args.numbering
    bad = check(positions, total, kind, toma, numbering)

    if not args.quiet:
        print("Перечень: %s (%s)" % (os.path.basename(args.perechen), kind))
        print("Позиций: %d   ·   сумма «на N л.»: %d   ·   нумерация листов: %s"
              % (len(positions), sum(p.deklar or 0 for p in positions),
                 "в каждом томе с 1" if numbering == "potomno" else "сквозная"))
        if total:
            print("Строка «ВСЕГО»: %d позиций, %d л." % total)
        for num, fname, pages, _path in toma:
            print("Том %d: %d стр.   %s" % (num, pages, fname))
        print("-" * 78)

    # Режимы плана тома и карты — поверх сверки, на её код возврата не влияют,
    # пока сами не найдут расхождения.
    if args.plan_save or args.plan_check or args.izyat:
        out = ["ПЛАН ТОМОВ:"]
        bad += run_plan_modes(args, positions, numbering, toma, args.perechen, out)
        print("\n".join(out))
    if args.map:
        out = []
        bad += cmd_map(args.map, positions, numbering, out)
        print("\n".join(out))

    if not bad:
        print("Расхождений не найдено.")
        return 0
    order = ["перечень", "нумерация", "л.д.", "объём", "источник", "итог", "тома",
             "повторы", "формат", "план", "карта"]
    print("РАСХОЖДЕНИЯ (%d):" % len(bad))
    for group in order:
        rows = [t for g, t in bad if g == group]
        if not rows:
            continue
        print("\n[%s]" % group.upper())
        for r in rows:
            print("  · %s" % r)
    return 1


if __name__ == "__main__":
    sys.exit(main())
