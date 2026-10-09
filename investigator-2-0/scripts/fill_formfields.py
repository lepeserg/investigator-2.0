# -*- coding: utf-8 -*-
"""fill_formfields.py — заполнение бланков-ФОРМ Word (статкарты Ф-1, Ф-1.2, Ф-2, Ф-3, Ф-5,
Ф-6, Ф-11, Ф-12, Ф-13) через Word COM ПО ПОЛЯМ ФОРМЫ, а не по координатам ячеек.

ПОДКОМАНДЫ
  inventory <бланк.docx> [--json out.json]
      Инвентарь полей: индекс, тип, имя, текущее значение, варианты списка, номер графы
      (своя строка → шапка блока → буква подграфы по сетке таблицы), блок и заголовок
      колонки. Снимается с ЛЮБОГО бланка, не только Ф-1/Ф-1.2.
      ⛔ КАРТА ДЛЯ ЗАПОЛНЕНИЯ — ЭТО НЕ ВСЕ ПОЛЯ, А ТОЛЬКО ОДНОЗНАЧНЫЕ ГРАФЫ. У блочных
      бланков одна графа приходится на десяток полей (Ф-1.1: 189 полей, однозначных
      граф 6) — такие поля адресуются ТОЛЬКО номером поля. Итоговая строка inventory
      печатает оба числа; чего в ней нет, того нет и в карте.

  fill <бланк.docx> <данные.json> <результат.docx> [--force]
      Копирует бланк в результат и заполняет копию. ОРИГИНАЛ НЕ ТРОГАЕТСЯ.
      Ключи данных: номер поля («3») либо номер графы («2.4», если она одна).
      Значение для списка проверяется по перечню; для текста — по maxLength.

  verify <файл.docx>
      Что заполнено, что пусто, и нет ли значений вне перечня списка.

ОКРУЖЕНИЕ
  Только Windows + установленный Word + pywin32. В Cowork не работает — там docx_edit.
  ⚠ После сбоя остаётся осиротевший WINWORD.EXE и держит файл: скрипт всегда прибирает
    за собой в finally, а `--kill-word` снимает зависшие процессы принудительно.
  ⚠ «Защищённый просмотр»: Word виснет на файлах с сетевых и съёмных дисков —
    скрипт отказывается работать с не-локальным путём и просит скопировать на C:."""
import os
import re
import sys
import json
import shutil
import argparse
import subprocess

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
try:
    from sk_config import setup_utf8          # локальный тулсет владельца, если рядом
    setup_utf8()
except Exception:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
KIND = {70: "текст", 71: "флажок", 83: "список"}
WD_TEXT, WD_CHECK, WD_DROP = 70, 71, 83


# ─────────────────────────── окружение ───────────────────────────

def _need_com():
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except Exception:
        sys.exit("Нужен pywin32 и установленный Word. В Cowork этого нет — "
                 "там путь через docx_edit (координаты ячеек, §09.16).")


def _check_local(path):
    """Word виснет на «Защищённом просмотре» файлов с сетевых/съёмных дисков."""
    p = os.path.abspath(path)
    if p.startswith("\\\\"):
        sys.exit(f"Сетевой путь — Word откроет в «Защищённом просмотре» и зависнет.\n"
                 f"Скопируй файл на локальный диск: {p}")
    drive = os.path.splitdrive(p)[0].upper()
    if drive and drive not in ("C:", "D:"):
        print(f"⚠ Диск {drive} — если это съёмный носитель, Word может открыть файл "
              f"в «Защищённом просмотре» и зависнуть. Надёжнее скопировать на C:.")
    return p


def kill_word():
    """Снять осиротевшие WINWORD.EXE (держат файл и роняют следующий запуск).

    ⛔ ТОЛЬКО ФОНОВЫЕ, И ПОДТВЕРЖДЁННЫЕ ДВУМЯ СПОСОБАМИ. Прежняя версия делала
    `taskkill /F /IM WINWORD.EXE` — то есть убивала и рабочий Word владельца со всеми
    несохранёнными документами (жалоба 14.08.2026).

    ⚠ Одного признака мало: `tasklist /FI "WINDOWTITLE eq N/A"` НЕ надёжен — живой документ
    владельца в одном вызове показывается безымянным, в другом с полным именем (§16.30).
    Поэтому «нет окна» подтверждаем ещё и `MainWindowHandle == 0` через PowerShell, и убиваем
    только пересечение двух списков. В сомнении — НЕ убиваем.

    ⚠ БЕЗ text=True: taskkill пишет в консольной кодировке (cp866), и декодирование
    его вывода как UTF-8 роняло уборку `UnicodeDecodeError` (поймано прогоном 06.08.2026)."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE",
                              "/FI", "WINDOWTITLE eq N/A", "/FO", "CSV", "/NH"],
                             capture_output=True, timeout=30).stdout.decode("cp866", "replace")
        by_title = {int(p) for p in re.findall(r'"WINWORD\.EXE","(\d+)"', out)}
    except Exception:
        return False
    try:
        ps = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "Get-Process WINWORD -ErrorAction SilentlyContinue | "
             "Where-Object {$_.MainWindowHandle -eq 0} | ForEach-Object {$_.Id}"],
            capture_output=True, timeout=60).stdout.decode("utf-8", "replace")
        by_handle = {int(x) for x in re.findall(r"\d+", ps)}
    except Exception:
        return False                      # второй способ не отработал — не убиваем ничего

    killed = 0
    for pid in sorted(by_title & by_handle):          # только то, с чем согласны ОБА способа
        try:
            r = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
            killed += (r.returncode == 0)
        except Exception:
            pass
    return killed > 0


def _word_pids():
    """PID всех запущенных WINWORD.EXE (чтобы отличить СВОЙ Word от пользовательского)."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
                             capture_output=True, timeout=30).stdout.decode("cp866", "replace")
    except Exception:
        return set()
    return {int(m) for m in re.findall(r'"WINWORD\.EXE","(\d+)"', out)}


def _hwnd_pid(app):
    """PID процесса по окну СВОЕГО COM-экземпляра (Application.Hwnd).

    None, если Hwnd недоступен или нет pywin32. В объектной модели Word свойство
    Hwnd документировано у Window, а не у Application, поэтому у Word оно, скорее
    всего, недоступно и функция возвращает None — тогда основным путём фактически
    становится разница снимков tasklist (см. _Word.__enter__)."""
    try:
        import win32process
        pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]
        return pid if pid > 0 else None
    except Exception:
        return None


class _Word:
    """Word COM с гарантированной уборкой (в т.ч. при исключении).

    ⛔ КРИТИЧНО (поймано 14.08.2026 по жалобе «документы произвольно закрываются
    без сохранения»). Word — ОДНОЭКЗЕМПЛЯРНЫЙ COM-сервер: `Dispatch("Word.Application")`
    ПОДКЛЮЧАЕТСЯ к уже открытому Word пользователя. Прежний код делал с ним
    `Visible = False` (окна с документами исчезали с экрана), `DisplayAlerts = 0`
    (запрос «сохранить?» подавлен) и в `__exit__` — `Quit()`. То есть скрипт закрывал
    рабочий сеанс владельца со всеми несохранёнными документами.
    Теперь: **DispatchEx** — всегда СВОЙ отдельный процесс; `Quit()` только для него;
    пользовательские экземпляры не трогаем; глобальные Options возвращаем как были.
    Не заменять `DispatchEx` на `Dispatch` ни при каких условиях.

    Свой PID (для taskkill зависшего процесса в `__exit__`) определяется через
    Application.Hwnd, а если его нет (у Word это, вероятно, обычный случай) — по
    разнице снимков tasklist, и только когда новый PID ровно один. Иначе не убиваем
    ничего: свой зависший Word может остаться висеть, но чужой не пострадает."""

    def __enter__(self):
        import pythoncom
        import win32com.client as win32
        self._pc = pythoncom
        pythoncom.CoInitialize()
        self._foreign = _word_pids()          # Word'ы владельца — их не трогаем
        # DispatchEx = ВСЕГДА новый процесс, не подключение к чужому сеансу
        self.app = win32.DispatchEx("Word.Application")  # НЕ gencache: виснет на первом запуске
        # ⛔ СВОЙ PID пытаемся взять по окну своего экземпляра. Разница двух снимков
        # tasklist — гонка: Word, открытый владельцем между снимками, мог попасть в «свои».
        # NB: у Word.Application свойства Hwnd, вероятно, нет (оно у объекта Window), так что
        # на практике обычно срабатывает запасной путь ниже — он и есть основной.
        pid = _hwnd_pid(self.app)
        if pid:
            self._mine = {pid}
        else:
            # запасной путь (Hwnd недоступен): разница снимков, но ТОЛЬКО если новый PID
            # ровно один; ноль или больше одного — не знаем, какой наш, и не убиваем ничего
            # (свой зависший Word тогда может остаться — это осознанная плата за безопасность)
            new_pids = _word_pids() - self._foreign
            self._mine = new_pids if len(new_pids) == 1 else set()
        self.app.Visible = False
        self.app.DisplayAlerts = 0
        # ⛔ Автозамена Word ПОРТИТ значение: «е» превращалось в "е" (прямые кавычки),
        # а это нарушение формата (только ёлочки) и флаг style_lint. Поймано на боевой
        # Локальный пример исключён из публичной поставки.
        # Options у Word — ГЛОБАЛЬНЫЕ и переживают перезапуск: запоминаем и вернём как было,
        # иначе меняем настройки владельца молча и навсегда.
        self._opts = {}
        for opt, val in (("AutoFormatAsYouTypeReplaceQuotes", False),
                         ("AutoFormatAsYouTypeReplaceHyphens", False),
                         ("AutoFormatAsYouTypeReplaceOrdinals", False),
                         ("AutoFormatAsYouTypeReplaceSymbols", False)):
            try:
                self._opts[opt] = getattr(self.app.Options, opt)
                setattr(self.app.Options, opt, val)
            except Exception:
                pass
        return self

    def __exit__(self, *exc):
        for opt, val in self._opts.items():
            try:
                setattr(self.app.Options, opt, val)
            except Exception:
                pass
        try:
            self.app.Quit()          # это НАШ экземпляр (DispatchEx), сеанс владельца не трогаем
        except Exception:
            pass
        # если наш процесс завис — снимаем ТОЛЬКО его, по PID, и никогда чужой
        for pid in getattr(self, "_mine", set()):
            try:
                subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
            except Exception:
                pass
        try:
            self._pc.CoUninitialize()
        except Exception:
            pass
        return False


# ─────────────────────────── привязка к графам ───────────────────────────
#
# Локальный пример исключён из публичной поставки.
# Прежняя версия искала номер графы только в ПЕРВОЙ и ПОСЛЕДНЕЙ ячейке своей строки.
# Замер на `Ф-11 нов.docx` (ныне архив): полей 189, графа определилась у 86, но ОДНОЗНАЧНО адресуемых
# (графа встречается ровно у одного поля, только такую и принимает `_resolve_key`) — ТРИ.
# Причина: у блочных бланков (Ф-1.1, Ф-1.3, Ф-2, Ф-3) номер графы стоит в ШАПКЕ БЛОКА
# строкой выше, буква подграфы («а) № прест.», «б) Статья, часть») — в ячейке шапки НАД
# колонкой, а строки данных объединены иначе, чем шапка, и крайняя ячейка пуста.
# Теперь: (1) вся строка, а не края; (2) шапка блока выше по таблице; (3) буква подграфы —
# по СЕТКЕ таблицы (`w:gridSpan`), а не по порядковому номеру ячейки: номера ячеек в шапке
# и в данных не совпадают, сетка совпадает всегда.
# ЗАМЕР 26.08.2026 (папка `Справочная информация\НОВЫЕ СТАТКАРТЫ`, 19 бланков):
#   Ф-11 нов.docx  графа у 86 из 189 → у 189 из 189; однозначно адресуемых 3 → 6;
#   все 19 бланков графа у 780 из 1445 → у 1331; однозначных 149 → 221;
#   попутно исправлены НЕВЕРНЫЕ подписи: Ф-1.3 (значения `001`/`12` шли за номер графы),
#   Ф-3 (гр. 3.2.б читалась как «3»), Ф-2 (гр. 2.13.а и 2.13.б — обе как «2.1»).
# ⚠ ЦИФРЫ ВЫШЕ — ИСТОРИЧЕСКИЕ: замер снят до переноса 26.08.2026, все 19 бланков теперь
#   в архиве `НОВЫЕ СТАТКАРТЫ\Старые стат. карты\` (разложены по папкам дел). У РАБОЧЕГО
#   бланка `Исходные бланки (с флешки)\Ф-11.docx` полей 210, а не 189 — с «189» не сверяться.
# ⛔ И ГЛАВНОЕ — честность: где графа не выведена, поле помечается пустой графой, а
# `inventory` печатает итог «графа определена у N из M · однозначно адресуемых K».
# Карта заполнения — это K, а не M: у блочных бланков одна графа приходится на десяток
# полей, и такие поля адресуются ТОЛЬКО номером поля. ⛔ Обещать карту по всем полям нельзя.

# ⛔ Номер графы — только вида «1», «2.19», «3.2.в»: РАЗДЕЛ У СТАТКАРТ ОДНОЗНАЧНЫЙ (1-3).
# Прежний `\d+(\.\d+)*` принимал за номер графы ЗНАЧЕНИЕ поля: на Ф-1.3 номер лица `001`
# и срок `12` уходили в графу, а строка со значением становилась «шапкой блока» для всех
# строк ниже (замер 26.08.2026: 4 поля из 28 были подписаны неверно). Ведущий ноль и
# двузначный раздел отвергаем; голая цифра без точки принимается только в крайней ячейке
# строки — там, где в бланках и стоит номер графы.
_RE_GRAPH_FULL = re.compile(r"[1-9](?:\.\d{1,2})*(?:\.[а-яё])?$")
_RE_GRAPH_HEAD = re.compile(r"^([1-9](?:\.\d{1,2})*(?:\.[а-яё])?)\s")
_RE_LETTER = re.compile(r"^\s*([а-яё])\s*\)")          # «б) Статья…», «д )Р ас-крыто»
_RE_LETTER_ANY = re.compile(r"(?:^|\s)([а-яё])\s*\)")  # сколько букв в ячейке всего
_UP_ROWS = 15                                          # насколько высоко искать шапку блока


def _cell_text(el):
    return " ".join((t.text or "").replace("\xa0", " ") for t in el.iter(W + "t")).strip()


def _norm_num(t):
    """«3 . 2 .б» → «3.2.б», «2.1 3» → «2.13», «2. 19 .а» → «2.19.а».

    ⛔ Word рвёт номер графы на несколько ранов (правка, курсив, проверка орфографии), а
    `w:t` мы склеиваем через пробел. Из-за этого гр. 3.2.б на Ф-3 читалась как «3», а
    гр. 2.13.а и 2.13.б на Ф-2 — обе как «2.1», то есть значение уходило В ЧУЖУЮ ГРАФУ.
    Пробелы снимаем только в ячейке, где, кроме цифр, точек и одной буквы, ничего нет.
    Нормализация — ТОЛЬКО для сопоставления, в отчёт идёт исходный текст ячейки."""
    s = re.sub(r"\s*\.\s*", ".", (t or "")).strip()
    if re.fullmatch(r"[0-9.\s]+[а-яё]?", s):
        s = re.sub(r"\s+", "", s)
    return s


def _tc_spans(tr):
    """Ячейки строки с координатами в СЕТКЕ таблицы: [(tc, начало, конец)].

    ⚠ Считать по `w:gridSpan`, а не по порядку: в шапке блока 6 ячеек, в строке данных 9,
    и «третья ячейка» шапки стоит над пятой ячейкой данных."""
    out, col = [], 0
    for c in tr:
        if c.tag != W + "tc":
            continue
        gs = c.find(W + "tcPr/" + W + "gridSpan")
        val = gs.get(W + "val") if gs is not None else None
        n = int(val) if (val or "").isdigit() else 1
        out.append((c, col, col + n))
        col += n
    return out


def _dist(a, b, x, y):
    """Расстояние между диапазонами сетки. ⛔ 0 — ТОЛЬКО пересечение.

    У соседних встык диапазонов зазор равен нулю ([1,2) и [0,1)), поэтому к зазору
    прибавляем единицу: иначе ячейка «3.2» слева считалась накрывающей нашу колонку
    и подписывала поле заголовком соседа (поймано на Ф-1.3 26.08.2026)."""
    if a < y and x < b:
        return 0
    return 1 + ((x - b) if b <= x else (a - y))


def _nearest(cands, a, b):
    """Из [(текст, начало, конец, значение)] — ближайшее к диапазону (a, b)."""
    if not cands:
        return ""
    return min(cands, key=lambda c: (_dist(a, b, c[1], c[2]), c[1]))[3]


def _has_field(tr):
    return next(tr.iter(W + "ffData"), None) is not None


def _is_graph(text, idx, n):
    """Текст ячейки — номер графы? Голая цифра засчитывается только в КРАЙНЕЙ ячейке."""
    t = _norm_num(text)
    if not t or not _RE_GRAPH_FULL.fullmatch(t):
        return False
    return "." in t or idx in (0, n - 1)


def _header_above(rows, ri, a, b):
    """Шапка блока выше -> (номер блока, буква подграфы, заголовок колонки, есть ли буквы).

    Шапка блока — ближайшая выше строка, где ОТДЕЛЬНАЯ ячейка равна номеру графы («2.3»).
    Буква берётся из ячейки ТОЙ ЖЕ шапки, накрывающей нашу колонку по сетке, и только если
    в ячейке ровно ОДНА пометка «х)»: ячейка-легенда «а) Номер преступления б) Квалификация»
    относится сразу к двум графам, по ней букву выводить нельзя.
    Четвёртое значение — были ли в шапке буквы вообще: если букв нет, блок и есть графа
    целиком (гр. 2.2, 2.3.е), если буквы есть, а наша колонка ни под одну не попала —
    ⛔ графа НЕ выводится."""
    block, letter, column, lettered = "", "", "", False
    for k in range(ri - 1, max(-1, ri - 1 - _UP_ROWS), -1):
        tr = rows[k]
        spans = _tc_spans(tr)
        texts = [(_cell_text(c), x, y) for c, x, y in spans]
        # заголовок колонки — из строки БЕЗ полей формы (иначе поймаем чужое значение)
        if not column and not _has_field(tr):
            over = [t for t, x, y in texts if t and _dist(a, b, x, y) == 0]
            if over:
                column = over[0]
        nums = [t for j, (t, x, y) in enumerate(texts) if _is_graph(t, j, len(texts))]
        if not nums:
            continue
        block = _norm_num(nums[0])
        for t, x, y in texts:
            m = _RE_LETTER.match(t or "")
            if not (m and len(_RE_LETTER_ANY.findall(t)) == 1):
                continue
            lettered = True
            if _dist(a, b, x, y) == 0:
                letter = m.group(1)
                if not column:
                    column = t
                break
        break
    return block, letter, column, lettered


def graph_labels(path):
    """Для каждого ffData ПО ПОРЯДКУ: номер графы, блок, заголовок колонки, текст ячейки.
    Порядок ffData в XML совпадает с порядком doc.FormFields(i).

    Ключи записи:
      `графа`   — номер графы, если выведен однозначно; иначе ПУСТО (⛔ не угадывать);
      `блок`    — номер графы-блока из шапки, когда подграфа не выведена (справочно);
      `столбец` — заголовок колонки из шапки (по сетке) — чем поле является по смыслу;
      `источник`— откуда взята графа: `строка` (своя строка) или `шапка` (шапка блока);
      `ячейка`  — собственный текст ячейки поля.
    """
    try:
        import docx
    except Exception:
        return []
    d = docx.Document(path)
    cache = {}
    out = []

    for ff in d.element.body.iter(W + "ffData"):
        tc = ff
        while tc is not None and tc.tag != W + "tc":
            tc = tc.getparent()
        rec = {"графа": "", "блок": "", "столбец": "", "источник": "", "ячейка": ""}
        if tc is None:
            out.append(rec)                       # поле вне таблицы — графы у него нет
            continue
        rec["ячейка"] = _cell_text(tc)[:120]
        tr = tc.getparent()
        if tr is None or tr.tag != W + "tr":
            out.append(rec)
            continue
        tbl = tr.getparent()
        rows = cache.get(id(tbl))
        if rows is None:
            rows = [r for r in tbl if r.tag == W + "tr"]
            cache[id(tbl)] = rows
        try:
            ri = rows.index(tr)
        except ValueError:
            ri = 0
        spans = _tc_spans(tr)
        a, b = next(((x, y) for c, x, y in spans if c is tc), (0, 1))
        texts = [(_cell_text(c), x, y) for c, x, y in spans]

        # 1) номер графы отдельной ячейкой в СВОЕЙ строке — берём ближайшую по сетке
        full = [(t, x, y, t) for j, (t, x, y) in enumerate(texts)
                if _is_graph(t, j, len(texts))]
        label = _nearest(full, a, b)
        # 2) номер графы в начале текста ячейки этой же строки («2.3 а) Номер преступления»)
        if not label:
            head = [(t, x, y, _RE_GRAPH_HEAD.match(_norm_num(t)).group(1)) for t, x, y in texts
                    if _RE_GRAPH_HEAD.match(_norm_num(t))]
            label = _nearest(head, a, b)
        if label:
            # своя строка сказала всё — шапку выше не подставляем: у построчных бланков
            # (Ф-1, Ф-3) «строка выше» — это ПРЕДЫДУЩАЯ графа, и в отчёте она только врёт
            rec["графа"], rec["источник"] = _norm_num(label), "строка"
            out.append(rec)
            continue
        # 3) шапка блока выше + буква подграфы по сетке
        block, letter, column, lettered = _header_above(rows, ri, a, b)
        rec["блок"], rec["столбец"] = block, column[:60]
        if block:
            if letter:
                rec["графа"], rec["источник"] = f"{block}.{letter}", "шапка"
            elif not lettered:          # в шапке букв нет — блок и есть графа целиком
                rec["графа"], rec["источник"] = block, "шапка"
        if rec["графа"] and rec["графа"] == rec["блок"]:
            rec["блок"] = ""
        out.append(rec)
    return out


def map_report(rows):
    """Итог по карте: сколько полей адресуемо по графе, а сколько — только по номеру."""
    seen = {}
    for r in rows:
        g = r.get("графа") or ""
        if g:
            seen[g] = seen.get(g, 0) + 1
    with_graph = sum(seen.values())
    uniq = sum(1 for r in rows if r.get("графа") and seen[r["графа"]] == 1)
    return {"полей": len(rows), "с графой": with_graph, "однозначных": uniq,
            "различных граф": len(seen)}


def _open(app, path, readonly):
    """Открыть документ и УБЕДИТЬСЯ, что вернулся именно Document.

    ⚠ При позднем связывании `Documents.Open` изредка отдаёт объект без нужных свойств
    (ловилось, когда в одном процессе подряд поднимались ДВА экземпляра Word). Проверяем
    сразу, иначе падение вылезет позже и непонятно где."""
    doc = app.Documents.Open(FileName=path, ReadOnly=readonly,
                             AddToRecentFiles=False, Visible=False)
    try:
        doc.FormFields.Count
    except Exception:
        raise RuntimeError(f"Word открыл файл, но объект документа неполноценный: {path}. "
                           f"Обычно помогает `--kill-word` (висит прежний экземпляр).")
    return doc


def _read_open_doc(doc, ctx):
    """Инвентарь полей УЖЕ открытого документа."""
    rows = []
    for i in range(1, doc.FormFields.Count + 1):
        f = doc.FormFields(i)
        rec = {"i": i, "тип": KIND.get(f.Type, str(f.Type)), "имя": f.Name}
        try:
            rec["значение"] = f.Result
        except Exception:
            rec["значение"] = None
        if f.Type == WD_DROP:
            try:
                rec["варианты"] = [f.DropDown.ListEntries(k).Name
                                   for k in range(1, f.DropDown.ListEntries.Count + 1)]
            except Exception:
                rec["варианты"] = []
        c = ctx[i - 1] if i - 1 < len(ctx) else {"графа": "", "блок": "", "столбец": "",
                                                 "источник": "", "ячейка": ""}
        rec.update(c)
        rows.append(rec)
    return rows


def read_fields(path):
    """Инвентарь полей: COM (индекс/тип/имя/значение/варианты) + графа из XML."""
    _need_com()
    path = _check_local(path)
    ctx = graph_labels(path)
    with _Word() as w:
        doc = _open(w.app, path, True)
        try:
            rows = _read_open_doc(doc, ctx)
        finally:
            doc.Close(0)
    if ctx and len(ctx) != len(rows):
        print(f"⚠ ffData в XML {len(ctx)} ≠ полей у Word {len(rows)} — привязка к графам ненадёжна")
    return rows


# ─────────────────────────── заполнение ───────────────────────────

SYMBOL_FONTS = {"MT Extra", "Symbol", "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings"}


def fix_symbol_font(path):
    'Вернуть значения полей из СИМВОЛЬНОГО шрифта в шрифт формы.'
    try:
        import docx
    except Exception:
        return 0
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    doc = docx.Document(path)

    cnt = {}
    for r in doc.element.body.iter(W + "r"):
        rf = r.find(W + "rPr/" + W + "rFonts")
        nm = rf.get(W + "ascii") if rf is not None else None
        if nm and nm not in SYMBOL_FONTS:
            cnt[nm] = cnt.get(nm, 0) + 1
    font = max(cnt, key=cnt.get) if cnt else "Arial"

    hits, inside = [], False
    for r in doc.element.body.iter(W + "r"):
        fc = r.find(W + "fldChar")
        if fc is not None:
            t = fc.get(W + "fldCharType")
            if t == "separate":
                inside = True
                continue
            if t == "end":
                inside = False
                continue
        if not inside:
            continue
        if not "".join(t.text or "" for t in r.iter(W + "t")).strip():
            continue
        rf = r.find(W + "rPr/" + W + "rFonts")
        if rf is not None and rf.get(W + "ascii") in SYMBOL_FONTS:
            hits.append(rf)

    if hits:
        for rf in hits:
            rf.set(W + "ascii", font)
            rf.set(W + "hAnsi", font)
            rf.attrib.pop(W + "hint", None)
        doc.save(path)
    return len(hits)


def _resolve_key(key, fields):
    """Ключ данных → индекс поля. Либо номер поля, либо НЕДВУСМЫСЛЕННЫЙ номер графы."""
    k = str(key).strip()
    if k.isdigit():
        i = int(k)
        if not (1 <= i <= len(fields)):
            raise KeyError(f"поля №{i} нет (всего {len(fields)})")
        return i
    k = _norm_num(k)
    cand = [f["i"] for f in fields if _norm_num(f.get("графа") or "") == k]
    if not cand:
        near = [f["i"] for f in fields
                if _norm_num(f.get("блок") or "") == k
                or _norm_num(f.get("графа") or "").startswith(k + ".")]
        if near:
            raise KeyError(f"графа «{k}» отдельным полем не адресуется; к ней относятся "
                           f"поля {near[:12]} — указывай номер поля (см. inventory)")
        raise KeyError(f"графа «{k}» среди полей не найдена. ⛔ Значащие графы Ф-1.1, Ф-2 "
                       f"и Ф-3 в полях формы не лежат — они обычные ячейки таблицы "
                       f"(§09.19.5), их путь — docx_edit + set_statcard_cell (§09.16)")
    if len(cand) > 1:
        show = "; ".join(f"{f['i']}·{(f.get('столбец') or f.get('ячейка') or '')[:24]}"
                         for f in fields if f["i"] in cand[:8])
        raise KeyError(f"графа «{k}» — это {len(cand)} полей [{show}]; "
                       f"указывай номер поля, а не графу")
    return cand[0]


def fill(blank, data_json, out_path, force=False):
    _need_com()
    blank = _check_local(blank)
    out_path = _check_local(out_path)
    if os.path.exists(out_path) and not force:
        sys.exit(f"Файл уже существует (перезапись только с --force): {out_path}")
    with open(data_json, encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        sys.exit("Данные должны быть объектом {ключ: значение}")

    ctx_blank = graph_labels(blank)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    shutil.copy2(blank, out_path)        # ⛔ оригинал бланка не трогаем НИКОГДА
    written = []

    # ⚠ ОДИН сеанс Word на всю операцию: и инвентарь бланка, и запись в копию.
    # Раньше сперва звался read_fields (свой сеанс Word), затем открывался ВТОРОЙ — и второй
    # цеплялся к закрывающемуся экземпляру: `Documents.Open` отдавал неполноценный объект
    # («AttributeError: Open.FormFields»). Поймано на боевой сборке 06.08.2026.
    with _Word() as w:
        doc = _open(w.app, out_path, False)
        try:
            fields = _read_open_doc(doc, ctx_blank)
            plan = []
            for key, val in data.items():
                if str(key).startswith("_"):     # «_комментарий» и т.п.
                    continue
                try:
                    i = _resolve_key(key, fields)
                except KeyError as e:
                    raise SystemExit("⛔ " + str(e).strip("'"))
                f = fields[i - 1]
                val = "" if val is None else str(val)
                if f["тип"] == "список":
                    variants = f.get("варианты") or []
                    norm = [v.strip() for v in variants]
                    if val.strip() not in norm:
                        raise SystemExit(f"⛔ Поле {i} (графа {f['графа'] or '—'}): значение {val!r} "
                                         f"вне перечня списка {variants}. Код не выдумывать.")
                    plan.append((i, "список", norm.index(val.strip()) + 1, val))
                else:
                    plan.append((i, "текст", val, val))

            for i, kind, payload, shown in plan:
                f = doc.FormFields(i)
                if kind == "список":
                    f.DropDown.Value = payload
                else:
                    f.Result = payload
                if str(f.Result).strip() != str(shown).strip():
                    raise SystemExit(f"Поле {i}: Word изменил значение из-за формата поля; проверьте тип и формат. Рабочая копия не сохранена.")
                written.append((i, fields[i - 1].get("графа", ""), shown, f.Result))
            doc.Save()
        except SystemExit:
            doc.Close(0)
            if os.path.exists(out_path):
                os.remove(out_path)      # при отказе файла-полуфабриката не оставляем
            raise
        else:
            doc.Close(0)

    # значения, попавшие в символьный шрифт заглушки, вернуть в шрифт формы
    n_sym = fix_symbol_font(out_path)
    if n_sym:
        print(f"⚠ значений было в символьном шрифте заглушки: {n_sym} — переведены в шрифт формы")

    # свойства файла → профиль владельца: бланки несут ЧУЖОГО автора и организацию
    # («ХХХХХ», Company «ГВП» — аудит 07.08.2026); автор документа = следователь
    try:
        import doc_meta
        fixes = doc_meta.clean(out_path)
        if fixes:
            print(f"свойства файла приведены к профилю владельца ({len(fixes)} правок)")
    except Exception:
        pass

    def _cmp(s):
        """Сравнение read-back: COM отдаёт значение с НОРМАЛИЗОВАННЫМИ кавычками
        («е» → "е"), хотя в документе сохраняются ёлочки — проверено по word/document.xml
        06.08.2026. Без этой нормализации выдавалось ложное «в поле легло …»."""
        return re.sub(r"[«»„“”\"']", "", str(s)).strip()

    print(f"заполнено полей: {len(written)}   →  {out_path}")
    for i, g, want, got in written:
        mark = "" if _cmp(got) == _cmp(want) else f"  ⚠ в поле легло {got!r}"
        print(f"  поле {i:>3}  графа {g or '—':<8} = {want!r}{mark}")
    return out_path


def verify(path):
    fields = read_fields(path)
    filled = [f for f in fields if (f.get("значение") or "").strip()]
    empty = [f for f in fields if not (f.get("значение") or "").strip()]
    bad = []
    for f in fields:
        if f["тип"] == "список":
            v = (f.get("значение") or "").strip()
            variants = [x.strip() for x in (f.get("варианты") or [])]
            if v and variants and v not in variants:
                bad.append(f)
    print(f"{os.path.basename(path)}: полей {len(fields)} | заполнено {len(filled)} | пусто {len(empty)}")
    if bad:
        print(f"\n⛔ значения ВНЕ перечня списка: {len(bad)}")
        for f in bad:
            print(f"  поле {f['i']} (графа {f['графа'] or '—'}): {f['значение']!r} ∉ {f['варианты']}")
    print("\nзаполненные:")
    for f in filled:
        print(f"  поле {f['i']:>3}  графа {f['графа'] or '—':<8} = {f['значение']!r}")
    return 1 if bad else 0


# ─────────────────────────── CLI ───────────────────────────

def main():
    ap = argparse.ArgumentParser(description="Заполнение бланков-форм Word по ПОЛЯМ ФОРМЫ (Word COM)")
    ap.add_argument("cmd", choices=["inventory", "fill", "verify"])
    ap.add_argument("args", nargs="*")
    ap.add_argument("--json", help="inventory: куда сохранить карту полей")
    ap.add_argument("--force", action="store_true", help="fill: перезаписать результат")
    ap.add_argument("--kill-word", action="store_true", help="сначала снять зависшие WINWORD.EXE")
    a = ap.parse_args()

    if a.kill_word:
        print("WINWORD.EXE снят" if kill_word() else "зависших WINWORD.EXE не было")

    if a.cmd == "inventory":
        if len(a.args) < 1:
            sys.exit("inventory <бланк.docx> [--json out.json]")
        rows = read_fields(a.args[0])
        rep = map_report(rows)
        print(f"{os.path.basename(a.args[0])}: полей формы {len(rows)}\n")
        print(f"{'i':>3} {'тип':<8} {'графа':<9} {'значение':<20} варианты / контекст")
        print("-" * 112)
        for r in rows:
            extra = ("[" + ", ".join((r.get("варианты") or [])[:10]) + "]") if r["тип"] == "список" \
                else (r.get("ячейка") or r.get("столбец") or "")[:58]
            if not r.get("графа") and r.get("блок"):
                extra = f"блок {r['блок']} · " + extra
            print(f"{r['i']:>3} {r['тип']:<8} {r['графа'] or '—':<9} "
                  f"{(r.get('значение') or '')[:18]:<20} {extra[:68]}")
        # ⛔ Итог — не украшение: карта заполнения = «однозначных», а не «полей».
        print(f"\nграфа определена у {rep['с графой']} из {rep['полей']} полей · "
              f"различных граф {rep['различных граф']} · "
              f"ОДНОЗНАЧНО адресуемых по графе {rep['однозначных']}")
        if rep["однозначных"] < rep["полей"]:
            print(f"⛔ остальные {rep['полей'] - rep['однозначных']} полей адресуй ТОЛЬКО "
                  f"номером поля (ключ «57»): одна графа = несколько полей.")
        if rep["с графой"] < rep["полей"]:
            print("⛔ у полей с графой «—» графа не выведена — не выдумывать: "
                  "значащие графы Ф-1.1/Ф-2/Ф-3 лежат в обычных ячейках таблицы (§09.19.5).")
        if a.json:
            json.dump(rows, open(a.json, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
            print(f"\nкарта сохранена: {a.json}")
        return 0

    if a.cmd == "fill":
        if len(a.args) < 3:
            sys.exit("fill <бланк.docx> <данные.json> <результат.docx> [--force]")
        fill(a.args[0], a.args[1], a.args[2], force=a.force)
        return 0

    if len(a.args) < 1:
        sys.exit("verify <файл.docx>")
    return verify(a.args[0])


if __name__ == "__main__":
    try:
        sys.exit(main())
    finally:
        # страховка: даже при исключении не оставляем Word держать файл
        kill_word()
