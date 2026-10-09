# -*- coding: utf-8 -*-
"""consistency_check.py — СКВОЗНАЯ СОГЛАСОВАННОСТЬ ФАКТА по всем документам дела.

Закрывает критическое правило 32: один и тот же факт (сумма, № счёта, адрес, дата, звание,
паспорт, № дела, ФИО с инициалами) не должен расходиться между документами дела. Расхождения
всплывали на сборке ОЗ и на мере пресечения — когда исправлять уже дорого.

Что делает: обходит .docx/.doc папки дела (тело + таблицы + КОЛОНТИТУЛЫ + НАДПИСИ), вытаскивает
значения по категориям и показывает те, у которых В РАЗНЫХ ФАЙЛАХ РАЗНЫЕ значения. Отдельно
сверяет объём приложений («Приложение… на N л.») с фактическим числом страниц названного файла.

⛔ ГРАНИЦА КАТЕГОРИИ «СУММЫ»: НА СЕРИЙНОМ ДЕЛЕ ОНА НЕ ПРИМЕНЯЕТСЯ И ЭТО ГОВОРИТСЯ ВСЛУХ.
Замер 04.09.2026 (дело 937, два протокола): категория выдала 27 «различающихся сумм», ложными
оказались все 27 — это суммы разных эпизодов серии. Привязку суммы к эпизоду (фамилия в том же
абзаце) пробовали и ОТВЕРГЛИ замером: у одного лица в одном абзаце законно стоят и доля,
и взятка, и итог, поэтому привязка дала уже 37 ложных позиций вместо 27. Поэтому: суммы
сравниваются между документами ТОЛЬКО на одно-двухфигурантном деле; как только в денежных
абзацах названо три и более лица, категория молчит и печатает, что не проверяла. Видимость
проверки хуже её отсутствия (правило 46).

Использование:
    python consistency_check.py "<папка дела>"
    python consistency_check.py "<папка дела>" --only суммы,счета
    python consistency_check.py "<папка дела>" --only приложения          # только «на N л.»
    python consistency_check.py "<папка дела>" --show-files      # к каждому значению — файлы
    python consistency_check.py "<папка дела>" --min-files 2     # игнорировать одиночные
    python consistency_check.py --pair ФАЙЛ1 ФАЙЛ2               # сверить два документа
    python consistency_check.py "<папка>" --namesakes "<файл>"   # белый список однофамильцев
    python consistency_check.py "<папка>" --no-attachments       # без сверки «на N л.»

Код возврата: 0 — расхождений нет; 1 — есть (для гейта перед выдачей окончания).
Зависимости: python-docx (для .docx есть zip-фолбэк), antiword — для .doc (иначе .doc пропускается)."""

import argparse
import collections
import json
import os
import re
import subprocess
import sys
import zipfile

from word_text import docx_xml_text, xml_text

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

SKIP_DIRS = {"__pycache__", ".git", "_Дубли", "_КАРАНТИН"}

# Категория -> (регэксп, нормализатор ключа). Ключ нужен, чтобы «1 100 000» и «1100000»
# считались одним значением, а «1 400 000» — другим.
# ⚠ Маркер «руб/рублей/коп» ОБЯЗАТЕЛЕН: без него в суммы лезли номера счетов, КБК, ОКТМО и УИН
# («417 116 031», «123 104», «7 384») — замер на реальной папке дела дал 21 «сумму» вместо 13.
_MONEY = re.compile(r"\b\d{1,3}(?:[  ]\d{3}){1,3}\s*(?:руб|рублей|коп)", re.IGNORECASE)
_ACCOUNT = re.compile(r"\b\d{20}\b")
_PASSPORT = re.compile(r"\b\d{2}\s?\d{2}\s*№?\s*\d{6}\b")
_CASE_NO = re.compile(r"\b\d\.\d{2}\.\d{4}\.\d{3,4}\.\d{6}\b")
_KRSP = re.compile(r"КРСП[^\d]{0,12}(\d{2,6})", re.IGNORECASE)
_UNIT = re.compile(r"(?:войсково[йе]\s+части?|в/ч)\s*(\d{4,6})", re.IGNORECASE)
# ⚠ _RANK требует слово «юстиции» — то есть ловит ТОЛЬКО звания юристов (следователь,
# руководитель, прокурор). Звание обвиняемого/свидетеля этой категорией НЕ проверяется
# (правило 41 — статус и звание на дату документа — остаётся за человеком).
_RANK = re.compile(r"\b(рядовой|ефрейтор|младший сержант|сержант|старший сержант|старшина|"
                   r"прапорщик|старший прапорщик|младший лейтенант|лейтенант|старший лейтенант|"
                   r"капитан|майор|подполковник|полковник)\s+юстиции\b", re.IGNORECASE)
_FIO_INI = re.compile(r"\b([А-ЯЁ][а-яё]{2,})\s+([А-ЯЁ]\.\s?[А-ЯЁ]\.)")
_ADDRESS = re.compile(r"ул\.\s*[А-ЯЁ][^,;.]{2,40},\s*д\.\s*\d+[А-Яа-я]?", re.IGNORECASE)

CATEGORIES = {
    "суммы":    (_MONEY,    lambda m: re.sub(r"\D", "", m.group(0))),
    "счета":    (_ACCOUNT,  lambda m: m.group(0)),
    "паспорта": (_PASSPORT, lambda m: re.sub(r"\D", "", m.group(0))),
    "дела":     (_CASE_NO,  lambda m: m.group(0)),
    "КРСП":     (_KRSP,     lambda m: m.group(1)),
    "в/части":  (_UNIT,     lambda m: m.group(1)),
    "звания":   (_RANK,     lambda m: m.group(1).lower()),
    "ФИО":      (_FIO_INI,  lambda m: m.group(1)),        # ключ — фамилия, значение — с инициалами
    "адреса":   (_ADDRESS,  lambda m: re.sub(r"\s+", " ", m.group(0)).lower()),
}


# ── Общие мелочи ──────────────────────────────────────────────────────────────────────────────

_SURNAME_SUF = ("ыми", "ому", "ему", "ым", "ом", "ем", "ой", "ей", "у", "ю", "а", "я", "е", "ы", "и")


def _stem(surname):
    """Зеркало `_surname_stem` из style_lint: без него склонение разводит одно лицо по разным
    ключам и настоящая недозамена проходит мимо."""
    s = (surname or "").strip()
    for suf in _SURNAME_SUF:
        if s.lower().endswith(suf) and len(s) - len(suf) >= 4:
            return s[:-len(suf)].lower()
    return s.lower()


def _norm_fio(value):
    '«Фамилия  А. И.» → «Фамилия|а.и.» — сравнимый ключ лица (основа фамилии + инициалы).'
    m = _FIO_INI.search(value or "")
    if not m:
        return re.sub(r"\s+", "", (value or "")).lower()
    return _stem(m.group(1)) + "|" + re.sub(r"\s+", "", m.group(2)).lower()


# Удалённый при рецензировании текст (w:delText) и коды полей (w:instrText) в документ не входят:
# без их вырезания старые ФИО/даты из правок давали ложные расхождения. Правила — общие,
# word_text.xml_text (как у check_tom, docx_integrity, style_lint).
def _xml_to_text(xml):
    return xml_text(xml)


_WARNED = set()


def _warn_unread(path, why):
    '«файл не прочитан» в stderr — один раз на файл (extract зовётся несколькими проверками).'
    if path not in _WARNED:
        _WARNED.add(path)
        print("⚠ не прочитан, в сверку не вошёл: %s (%s)" % (path, why), file=sys.stderr)


def _docx_text(path):
    try:
        return docx_xml_text(path)
    except Exception as e:
        _warn_unread(path, e)
        return None


def _doc_text(path):
    try:
        r = subprocess.run(["antiword", "-w", "0", "-m", "UTF-8.txt", path],
                           capture_output=True, timeout=60)
        if r.returncode == 0 and r.stdout:
            return r.stdout.decode("utf-8", "replace")
    except Exception as e:
        _warn_unread(path, "antiword: %s" % e)
    return None


def extract(path):
    low = path.lower()
    if low.endswith(".docx"):
        return _docx_text(path)
    if low.endswith(".doc"):
        return _doc_text(path)
    return None


# ── ПОДПИСАНТ: берём из навыка, а не считаем однофамильцем фигуранта ──────────────────────────
# Локальный пример исключён из публичной поставки.
# всплывала бы в КАЖДОМ прогоне (провал 04.09.2026). Источник правды — `константы.json` рядом
# с навыком, фолбэк — таблицы `references/01-identity.md` («| Сокращённое ФИО | Х / Y |»).

_SKILL_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_IDENTITY_SHORT = re.compile(r"\|\s*Сокращённое ФИО\s*\|\s*([^|]+)\|")


def load_signers(skill_root=None):
    """{нормализованное «фамилия|и.о.»: как написано} для следователя, руководителя и зама."""
    root = skill_root or _SKILL_ROOT
    out = {}

    def add(value):
        for part in re.split(r"[/,]", value or ""):
            part = part.strip()
            if _FIO_INI.search(part):
                out[_norm_fio(part)] = part

    try:
        with open((os.environ.get("SK_CONSTANTS") if skill_root is None else None) or os.path.join(root, "константы.json"), encoding="utf-8") as f:
            data = json.load(f)

        def walk(node):
            if isinstance(node, dict):
                for k, v in node.items():
                    if k == "инициалы_фам" and isinstance(v, str):
                        add(v)
                    else:
                        walk(v)
            elif isinstance(node, list):
                for v in node:
                    walk(v)

        walk(data)
    except Exception:
        pass
    if not out:
        try:
            with open(os.path.join(root, "references", "01-identity.md"), encoding="utf-8") as f:
                for line in f:
                    m = _IDENTITY_SHORT.search(line)
                    if m:
                        add(m.group(1))
        except Exception:
            pass
    return out


def load_own_addresses(skill_root=None):
    """{нормализованный адрес} — адреса САМОГО ВСО и постоянных адресатов.

    ⛔ Адрес отдела попадал в «различающиеся адреса дела» наравне с адресами эпизодов
    Адреса берутся из локального профиля. Источник — `константы.json`
    и `адресаты.json` рядом с навыком: там они уже заведены и правятся в одном месте."""
    root = skill_root or _SKILL_ROOT
    out = set()
    for name in ("константы.json", "адресаты.json"):
        try:
            with open((os.environ.get("SK_CONSTANTS" if name == "константы.json" else "SK_ADDRESSEES") if skill_root is None else None) or os.path.join(root, name), encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue

        def walk(node):
            if isinstance(node, dict):
                for v in node.values():
                    walk(v)
            elif isinstance(node, list):
                for v in node:
                    walk(v)
            elif isinstance(node, str):
                for m in _ADDRESS.finditer(node):
                    out.add(re.sub(r"\s+", " ", m.group(0)).lower())

        walk(data)
    return out


# ── БЕЛЫЙ СПИСОК ОДНОФАМИЛЬЦЕВ по делу ───────────────────────────────────────────────────────
# Подтверждённая пара гасится один раз, а не разбирается заново каждую сессию.

NAMESAKES_FILE = "однофамильцы.txt"
CARD_FILE = "карточка.json"
CARD_KEY = "однофамильцы"


def load_namesakes(root, explicit=None):
    """[set(нормализованных лиц)] — группы подтверждённых однофамильцев."""
    groups = []

    def add_line(line):
        line = (line or "").split("#")[0].strip()
        if not line:
            return
        people = {_norm_fio(x) for x in re.split(r"[|;]", line) if _FIO_INI.search(x or "")}
        if len(people) >= 2:
            groups.append(people)

    paths = [explicit] if explicit else []
    if root:
        paths.append(os.path.join(root, NAMESAKES_FILE))
    for p in paths:
        if p and os.path.isfile(p):
            try:
                with open(p, encoding="utf-8") as f:
                    for line in f:
                        add_line(line)
            except Exception:
                pass
    card = os.path.join(root, CARD_FILE) if root else None
    if card and os.path.isfile(card):
        try:
            with open(card, encoding="utf-8") as f:
                data = json.load(f)
            for grp in (data.get(CARD_KEY) or []):
                if isinstance(grp, str):
                    add_line(grp)
                elif isinstance(grp, (list, tuple)):
                    add_line(" | ".join(str(x) for x in grp))
        except Exception:
            pass
    return groups


# ── Сбор значений ────────────────────────────────────────────────────────────────────────────

def iter_docs(root):
    """[(абсолютный путь, относительное имя)] по папке дела."""
    out = []
    for dp, dn, fn in os.walk(root):
        dn[:] = [d for d in dn if d not in SKIP_DIRS and not d.startswith("~$")]
        for f in sorted(fn):
            if f.startswith("~$") or not f.lower().endswith((".doc", ".docx")):
                continue
            path = os.path.join(dp, f)
            out.append((path, os.path.relpath(path, root)))
    return out


def collect(docs, only=None):
    """(data, прочитано, пропущено .doc, лица в денежных абзацах).

    data — {категория: {ключ: {значение: set(файлы)}}}.
    ⚠ Разбор идёт ПО АБЗАЦАМ, а не по всему тексту одной строкой: нужно знать, сколько РАЗНЫХ
    лиц названо рядом с суммами, — по этому числу решается, применима ли категория «суммы»."""
    data = {c: collections.defaultdict(lambda: collections.defaultdict(set))
            for c in CATEGORIES}
    money_people = set()
    scanned = skipped_doc = 0
    for path, rel in docs:
        txt = extract(path)
        if txt is None:
            if path.lower().endswith(".doc"):
                skipped_doc += 1
            continue
        scanned += 1
        for para in txt.split("\n"):
            flat = re.sub(r"[\s ]+", " ", para).strip()
            if not flat:
                continue
            has_money = bool(_MONEY.search(flat))
            if has_money:
                money_people.update(_stem(n[0]) for n in _FIO_INI.findall(flat))
            for cat, (rx, keyf) in CATEGORIES.items():
                if only and cat not in only:
                    continue
                for m in rx.finditer(flat):
                    val = re.sub(r"\s+", " ", m.group(0)).strip()
                    data[cat][keyf(m)][val].add(rel)
    return data, scanned, skipped_doc, money_people


# Категории двух типов:
#  · ТОЖДЕСТВО — ключ определяет сущность (фамилия, счёт, № дела, ЭПИЗОД для сумм): расхождение =
# Локальный пример исключён из публичной поставки.
#  · РАЗБРОС — сам набор разных значений и есть сигнал (адреса, звания): показываем ВСЕ
#    различающиеся значения с числом файлов. ⚠ Ранний вариант требовал ≥2 файлов на значение и
#    потому МОЛЧА пропускал ровно тот случай, ради которого писался (одиночное значение из
#    банковской выписки).
# ⛔ «Суммы» из РАЗБРОСА ИЗЪЯТЫ: на серийном деле общий разброс — это суммы разных эпизодов,
# и он давал 27 ложных позиций из 27 (замер 04.09.2026). Общий разброс сумм показывается только
# на ОДНО-ЭПИЗОДНОМ деле — см. `_spread_sums_allowed`.
# Сколько разных лиц рядом с суммами делает дело СЕРИЙНЫМ. Одно-двухфигурантное дело даёт
# 1–2 фамилии (обвиняемый и потерпевший), и там разброс сумм — настоящий сигнал: именно он ловил
# «1 400 000 в постановлении против 1 100 000 по банку». Три и больше — серия эпизодов.
_SERIAL_PEOPLE_MIN = 3
_SPREAD = ("суммы", "адреса", "звания")
_SPREAD_NOTE = {
    "суммы":  "сверить с ПЕРВИЧНЫМ документом (выписка, чек), а не с фабулой ВУД — правило 31",
    "адреса": "сверить с протоколом/выпиской: адрес операции часто уточняется офисом и домом. "
              "⚠ На многофигурантном деле разные адреса — это разные эпизоды плюс адрес самого "
              "ВСО; расхождением считать только адрес ОДНОГО и того же события",
    "звания": "ТОЛЬКО звания юристов (подписантов) — звание фигуранта здесь НЕ проверяется; "
              "разные звания подписанта МОГУТ быть верны — сверить с ДАТОЙ документа (§16.27)",
}
SHOW_LIMIT = 12   # сколько значений печатать по умолчанию (одиночные — сверх лимита, всегда)


def report(data, show_files=False, min_files=1, show_all=False,
           signers=None, namesakes=None, money_people=None, own_addresses=None):
    signers = signers or {}
    namesakes = namesakes or []
    own_addresses = own_addresses or set()
    money_people = money_people or set()
    serial = len(money_people) >= _SERIAL_PEOPLE_MIN
    problems = 0
    suppressed = []
    for cat, keys in data.items():
        if cat == "суммы" and serial:
            continue        # серийное дело: категория не применяется, сообщение — ниже
        rows = []
        for key, variants in sorted(keys.items()):
            if cat == "ФИО":
                # ⛔ ПОДПИСАНТ не участвует: следователь/руководитель/зам — не однофамилец
                # Локальный пример исключён из публичной поставки.
                own = [v for v in variants if _norm_fio(v) in signers]
                if own:
                    variants = {v: f for v, f in variants.items() if _norm_fio(v) not in signers}
                    suppressed.append(f"ФИО «{key}»: вариант(ы) {', '.join(sorted(own))} — "
                                      "подписант из навыка, не однофамилец фигуранта")
                # ⛔ Подтверждённые однофамильцы по делу — гасятся один раз белым списком.
                norm = {_norm_fio(v) for v in variants}
                if len(norm) >= 2 and any(norm <= g for g in namesakes):
                    suppressed.append(f"ФИО «{key}»: {', '.join(sorted(variants))} — "
                                      f"подтверждённые однофамильцы ({NAMESAKES_FILE})")
                    continue
            if len(variants) <= 1:
                continue
            note = "написано по-разному"
            if cat == "ФИО":
                # Одна фамилия с разными инициалами — это либо НЕДОЗАМЕНА (редкий вариант),
                # либо ДВА РАЗНЫХ ЛИЦА (оба варианта встречаются часто). Различаем по доле:
                # у владельца это реальный случай — фигурант и его супруга.
                counts = sorted((len(f) for f in variants.values()), reverse=True)
                ratio = counts[1] / max(counts[0], 1)
                # Нужны ОБА условия: доля и абсолют. Только доля ошибалась на малых числах —
                # «3 против 1» это 0,33, но одиночное вхождение почти всегда недозамена.
                two_people = ratio >= 0.25 and counts[1] >= 3
                note = ("вероятно ДВА РАЗНЫХ ЛИЦА с одной фамилией — убедиться, что не перепутаны"
                        if two_people else
                        "похоже на НЕДОЗАМЕНУ инициалов: фамилию заменили, инициалы донора остались")
                # Подсказка нужна на ЛЮБОМ вердикте: подтверждённая пара гасится один раз.
                note += ("; лица разные и это проверено — внести строкой «"
                         + " | ".join(sorted(variants)) + "» в " + NAMESAKES_FILE
                         + " рядом с делом, чтобы не разбирать её снова")
            rows.append((key, variants, note))
        if cat in _SPREAD and len(keys) > 1:
            merged = {}
            for k, variants in keys.items():
                if cat == "адреса" and k in own_addresses:
                    suppressed.append(f"адрес «{k}» — адрес самого ВСО либо постоянного "
                                      "адресата (константы.json / адресаты.json), не эпизод")
                    continue
                for val, files in variants.items():
                    merged.setdefault(val, set()).update(files)
            merged = {v: f for v, f in merged.items() if len(f) >= min_files}
            if len(merged) > 1:
                rows.append(("‹различающиеся значения в деле›", merged, _SPREAD_NOTE[cat]))
        if not rows:
            continue
        print(f"\n=== {cat.upper()} — позиций к проверке: {len(rows)} ===")
        for key, variants, note in rows:
            problems += 1
            print(f"  • {key}  ({note}):")
            # РЕДКИЕ — ПЕРВЫМИ. Прежняя сортировка (по убыванию числа файлов) при срезе
            # отрезала как раз одиночные значения — ровно тех кандидатов в расхождение,
            # ради которых написан инструмент (ревизия 22.08.2026).
            items = sorted(variants.items(), key=lambda kv: (len(kv[1]), str(kv[0])))
            singles = sum(1 for _, files in items if len(files) == 1)
            limit = len(items) if show_all else max(SHOW_LIMIT, singles)
            for val, files in items[:limit]:
                print(f"      «{val}»  — в {len(files)} файл(ах)")
                if show_files:
                    for f in sorted(files)[:6]:
                        print(f"          {f}")
            hidden = len(items) - limit
            if hidden > 0:
                print(f"      … скрыто ещё {hidden} значений (самые частые; "
                      f"одиночные показаны все). Показать всё: --all")
    if serial and data.get("суммы"):
        print(f"\n⛔ СУММЫ НЕ ПРОВЕРЯЛИСЬ: дело серийное — рядом с суммами названо разных лиц "
              f"{len(money_people)} (порог {_SERIAL_PEOPLE_MIN}). Разброс сумм на серии — это "
              "суммы РАЗНЫХ эпизодов, а не расхождение (замер: 27 позиций, ложных 27). Сумму "
              "конкретного эпизода сверять руками по ПЕРВИЧНОМУ документу — выписка, чек, "
              "протокол осмотра (правило 31), а не по фабуле ВУД. Свести два конкретных "
              "протокола: --pair ФАЙЛ1 ФАЙЛ2.")
    if suppressed:
        print("\n— погашено белым списком и подписантом:")
        for s in suppressed:
            print("   ·", s)
    return problems



# ── ПРИЛОЖЕНИЕ… НА N Л. против фактического числа страниц ─────────────────────────────────────
# Правило 20 говорит про каскад ФАКТОВ, объём под него не подпадал: обвинение выросло с 6 до 9 л.,
# а в сопроводе 84/40 осталось «Приложение: по тексту на 6 л.», в сопроводе прокурору — «на 7 л.»
# Локальный пример исключён из публичной поставки.

_SHEETS = re.compile(r"на\s+(\d{1,3})\s*л(?:\.|истах?\b)", re.IGNORECASE)
_PRIL_HEAD = re.compile(r"^\s*Приложени[ея]", re.IGNORECASE)
_ITEM_MARK = re.compile(r"^\s*(?:\d{1,2}\s*[).]|[-–—•])\s*")
_ITEM_SPLIT = re.compile(r"(?:(?<=\s)|^)\d{1,2}\s*[).]\s")
_SIGN_HEAD = re.compile(r"^\s*(Руководитель|Заместитель|Врио|Следователь|Старший следователь|"
                        r"И\.о\.)", re.IGNORECASE)
_BY_TEXT = re.compile(r"по\s+тексту", re.IGNORECASE)
_WORD_RX = re.compile(r"[А-Яа-яЁёA-Za-z]{5,}")
_DESC_STOP = {"приложение", "приложения", "приложении", "тексту", "текст", "экземпляр",
              "экземплярах", "листах", "адрес", "прошу", "копия", "копии", "всего", "также"}
# Свойства файла достоверны, только если документ в последний раз сохранял Word и цифры сходятся
# между собой: «Pages 1» при «Characters 16576» — это устаревшая метка (файл после Word правил
# скрипт). Границы знаков на страницу — по корпусу А4/14 пт.
_CHARS_PER_PAGE = (400, 4000)
_META_TOLERANCE = (0.7, 1.4)      # во сколько раз фактический текст может расходиться с меткой


def _pages(path):
    """Число страниц документа либо None, если сверить нечем (метка отсутствует/устарела)."""
    low = path.lower()
    if low.endswith(".pdf"):
        try:
            with open(path, "rb") as f:
                n = len(re.findall(rb"/Type\s*/Page[^s]", f.read()))
            return n or None
        except Exception:
            return None
    if not low.endswith(".docx"):
        return None
    try:
        with zipfile.ZipFile(path) as z:
            app = z.read("docProps/app.xml").decode("utf-8", "replace")
    except Exception:
        return None

    def tag(name):
        m = re.search(r"<%s>([^<]*)</%s>" % (name, name), app)
        return m.group(1) if m else ""

    if "Word" not in tag("Application"):
        return None                        # файл не сохранялся Word — числа страниц нет
    try:
        pg, ch = int(tag("Pages")), int(tag("Characters"))
    except Exception:
        return None
    if pg <= 0 or ch <= 0:
        return None
    if not (_CHARS_PER_PAGE[0] <= ch / pg <= _CHARS_PER_PAGE[1]):
        return None                        # метка внутренне противоречива — не доверяем
    txt = extract(path)
    if txt is None:
        return None
    real = len(re.sub(r"\s", "", txt))
    if not (_META_TOLERANCE[0] * ch <= real <= _META_TOLERANCE[1] * ch):
        return None                        # текст ушёл от метки: документ правили после Word
    return pg


def _sig_words(text):
    return [w.lower() for w in _WORD_RX.findall(text or "") if w.lower() not in _DESC_STOP]


def _match_attachment(desc, siblings):
    """(имя файла, страниц) для названного приложения либо None.

    ⛔ Сопоставление только по числу общих слов ошибалось: «Ходатайство подозреваемого о
    заключении досудебного соглашения» садилось на «Постановление о возбуждении ходатайства
    о ДСС» — общих слов хватало. Поэтому ЖАНР (первое значимое слово описания) обязан совпасть
    с первым значимым словом имени файла."""
    words = _sig_words(desc)
    if not words:
        return None
    head = words[0][:5]
    dset = {w[:5] for w in words}
    best = []
    for name, pages in siblings.items():
        fw = _sig_words(os.path.splitext(name)[0])
        if not fw or fw[0][:5] != head:
            continue
        best.append((len({w[:5] for w in fw} & dset), name, pages))
    best.sort(reverse=True)
    if not best or best[0][0] < 2:
        return None
    if len(best) > 1 and best[1][0] == best[0][0]:
        return None                        # два одинаково похожих файла — не гадаем
    return best[0][1], best[0][2]


def _folder_pages(folder):
    """{имя файла: страниц} по всем файлам папки, у которых число страниц достоверно."""
    out = {}
    try:
        names = os.listdir(folder)
    except Exception:
        return out
    for n in names:
        if n.startswith("~$"):
            continue
        p = os.path.join(folder, n)
        if not os.path.isfile(p):
            continue
        v = _pages(p)
        if v:
            out[n] = v
    return out


def _attachment_items(text):
    """[(описание, N)] из блока «Приложение…» документа."""
    items, inside = [], False
    for para in (text or "").split("\n"):
        line = re.sub(r"[\s ]+", " ", para).strip()
        if not line:
            continue
        if _PRIL_HEAD.match(line):
            inside = True
        elif inside and (_SIGN_HEAD.match(line)
                         or not (_ITEM_MARK.match(line) or _SHEETS.search(line))):
            inside = False
        if not (inside and _SHEETS.search(line)):
            continue
        # ⚠ Весь перечень бывает ОДНИМ абзацем: «1. Постановление … на 1 л. 2. Опись … на 63 л.».
        # Без разрезания по номерам описание позиции склеивалось с соседней (ложный флаг 05.09).
        for chunk in _ITEM_SPLIT.split(line):
            for m in _SHEETS.finditer(chunk):
                desc = chunk[:m.start()]
                desc = re.sub(r"^Приложени[ея]\s*:?\s*", "", _ITEM_MARK.sub("", desc),
                              flags=re.IGNORECASE).strip(" ,;:")
                items.append((desc, int(m.group(1))))
    return items


def check_attachments(docs):
    """([строки расхождений], [строки «сверить вручную»]) по «Приложение… на N л.»."""
    mismatched, manual = [], []
    folders = {}
    for path, rel in docs:
        txt = extract(path)
        if txt is None:
            continue
        items = _attachment_items(txt)
        if not items:
            continue
        folder = os.path.dirname(path)
        if folder not in folders:
            folders[folder] = _folder_pages(folder)
        self_name = os.path.basename(path)
        sib = {n: v for n, v in folders[folder].items() if n != self_name}
        if not sib:
            continue
        for desc, n in items:
            got = _match_attachment(desc, sib)
            if got:
                name, pages = got
                if pages != n:
                    mismatched.append(f"{rel}\n      «{desc[:70]}» — заявлено на {n} л., "
                                      f"а «{name}» — {pages} стр.")
            elif _BY_TEXT.search(desc) or not desc:
                if n not in set(sib.values()):
                    manual.append(f"{rel}: «{desc or 'Приложение'}» на {n} л. — ни один документ "
                                  f"папки не на {n} стр. (есть: {sorted(set(sib.values()))})")
    return mismatched, manual


def main():
    ap = argparse.ArgumentParser(description="Сквозная согласованность фактов по папке дела")
    ap.add_argument("root", nargs="?", help="папка дела (не нужна при --pair)")
    ap.add_argument("--pair", nargs=2, metavar=("ФАЙЛ1", "ФАЙЛ2"),
                    help="сверить два конкретных документа (они могут лежать в разных ветках "
                         "корпуса — обход по папке дела их не свёл бы)")
    ap.add_argument("--only", default=None,
                    help="категории через запятую: " + ",".join(CATEGORIES)
                         + ",приложения (сверка «на N л.»; входит по умолчанию, "
                           "а при --only её надо назвать явно)")
    ap.add_argument("--show-files", action="store_true")
    ap.add_argument("--min-files", type=int, default=1)
    ap.add_argument("--namesakes", default=None,
                    help=f"файл белого списка однофамильцев (по умолчанию {NAMESAKES_FILE} "
                         f"и {CARD_FILE} рядом с делом)")
    ap.add_argument("--no-attachments", action="store_true",
                    help="не сверять «Приложение… на N л.» с числом страниц файлов")
    ap.add_argument("--all", action="store_true",
                    help="печатать ВСЕ значения без лимита (по умолчанию частые усечены "
                         "до 12, одиночные показываются всегда)")
    a = ap.parse_args()

    if a.pair:
        missing = [p for p in a.pair if not os.path.isfile(p)]
        if missing:
            print("Не файл:", "; ".join(missing))
            return 2
        docs = [(p, os.path.basename(p)) for p in a.pair]
        card_root = os.path.dirname(os.path.abspath(a.pair[0]))
    else:
        if not a.root or not os.path.isdir(a.root):
            print("Не папка:", a.root)
            return 2
        docs = iter_docs(a.root)
        card_root = a.root

    signers = load_signers()
    own_addresses = load_own_addresses()
    namesakes = load_namesakes(card_root, a.namesakes)
    only = set(x.strip() for x in a.only.split(",")) if a.only else None
    data, scanned, skipped, money_people = collect(docs, only)
    print(f"consistency_check: просмотрено документов {scanned}"
          + (f"; .doc пропущено без antiword: {skipped}" if skipped else ""))
    if a.pair:
        print("режим --pair: сверяются ровно два файла — "
              + " ↔ ".join(os.path.basename(p) for p in a.pair))
    if not signers:
        print("⚠ подписант не прочитан (нет константы.json / 01-identity.md) — своя фамилия "
              "может всплыть как однофамилец фигуранта")
    n = report(data, a.show_files, a.min_files, a.all, signers, namesakes, money_people,
               own_addresses)

    manual = []
    if not a.no_attachments and not (only and "приложения" not in only):
        bad, manual = check_attachments(docs)
        if bad:
            print("\n=== ПРИЛОЖЕНИЯ — объём не сходится с файлом (правило 20) ===")
            for line in bad:
                print("  •", line)
            print("      ⚠ число страниц берётся из свойств файла (последнее сохранение в Word). "
                  "Документ, правленный после этого скриптом, к сверке НЕ допускается — такие "
                  "файлы молчат, а не врут (правило 46).")
        n += len(bad)

    if manual:
        print("\n— приложения, которые машина сверить не может (проверить вручную):")
        for m in manual:
            print("   ·", m)

    if n:
        print(f"\n{'=' * 62}\nРАСХОЖДЕНИЙ: {n}. Правило 32 — устранить ДО выдачи окончания.")
        print("Строить документ на ПЕРВИЧНОМ доказательстве (правило 31), затем синхронно "
              "поправить остальные документы дела.")
        return 1
    print("\nРасхождений не найдено. ⚠ Это механическая проверка: смысловые расхождения "
          "(один факт разными словами) ловит человек.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
