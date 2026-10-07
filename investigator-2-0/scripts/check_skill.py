# -*- coding: utf-8 -*-
"""check_skill.py — САМОПРОВЕРКА навыка перед релизом (не документа, а самого скилла).

Зачем: навык растёт правками по боевым отчётам, и накапливается «гниль ссылок» — правило ссылается
на §, которого нет; в индексе указан скрипт, которого нет; два раздела получили один номер; в
`SKILL.md` сломан YAML (навык тогда молча не грузится). Глазами это не ловится.

Проверяет:
  1. YAML-frontmatter `SKILL.md` парсится, есть `name` и `description` (инвариант: кавычки при «: »).
  2. Все файлы `references/*.md`, упомянутые в тексте, существуют; все существующие — упомянуты.
  3. Все `scripts/*.py`, упомянутые в тексте, существуют; все существующие — упомянуты.
  4. Ссылки на разделы вида «§07.05.8», «§06.01.13», «§16.21» разрешаются в реальные заголовки.
  5. Нет ДУБЛЕЙ номеров разделов внутри одного файла.
  6. Ссылки на «правило N» не выходят за фактическое число критических правил в `SKILL.md`.
  7. Нет управляющих символов (следы битой генерации) и BOM — в md И в скриптах.
  8. Все .py компилируются (`compile()`), у каждого есть docstring.
  9. ПАСПОРТА ЖАНРОВ: у обязательных жанровых файлов паспорт есть, стоит ПЕРВЫМ (до первого
     нумерованного раздела) и полон — все пять блоков; у прочих жанровых — замечание.
 10. Один номер раздела не занят в ДВУХ файлах сразу (проверка 5 смотрит только внутри файла).
 11. Ссылки нестандартной формы (§08f.01, §08e.2, §07.10А) — их регулярка §NN.NN не видит.
 12. `константы.json` и `адресаты.json` парсятся как JSON (иначе генератор молча уходит
     на фолбэк с плейсхолдерами), шаблон гербового бланка на месте и является .docx.
 13. `scripts/requirements.txt` покрывает реальные импорты скриптов (python-docx, pymupdf,
     pywin32) — иначе на новой машине путь отваливается молча.
 14. У каждого CLI-скрипта есть разбор `--help`.
 15. Строки "text" в ```python-шаблонах references (то, что генератор кладёт В ДОКУМЕНТ) не несут
     длинного тире вне плейсхолдера {…} и сдвоенного «ст. ст.» — иначе линтер ловит это уже в каждом
     документе заново (09.09.2026: 27 тире в шаблонах семи жанров, донорский хвост с «ст. ст.» в шести).

Запуск:
    python check_skill.py [<корень навыка>]        # по умолчанию — папка на уровень выше scripts/
Код возврата: 0 — чисто; 1 — есть замечания.
"""
import os
import re
import sys
import json
import ast
from pathlib import Path
from urllib.parse import unquote


def _check_bundle_structure(root):
    """Version 2 entrypoints, strict YAML/JSON and Markdown paths, excluding runtime."""
    errors = []
    bundle = Path(root).resolve().parent
    names = ['investigator-2-0'] + ['investigator-2-0-' + x for x in
             ('analysis', 'correspondence', 'interrogations', 'actions', 'decisions', 'charging', 'stat-cards')]
    try:
        import yaml
    except ImportError:
        return ['Для полной проверки восьми навыков требуется PyYAML (запуск через run.py)']
    class UniqueLoader(yaml.SafeLoader):
        pass
    def unique_pairs(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError('Повторный ключ: ' + str(key))
            result[key] = value
        return result
    def mapping(loader, node):
        return unique_pairs(loader.construct_pairs(node, deep=True))
    UniqueLoader.add_constructor(yaml.resolver.BaseResolver.DEFAULT_MAPPING_TAG, mapping)
    for name in names:
        folder = bundle / name
        entry = folder / 'SKILL.md'
        if not entry.is_file():
            errors.append(f'{name}: отсутствует обязательный SKILL.md')
            continue
        try:
            match = re.match(r'^---\s*\n(.*?)\n---(?:\n|$)', _read(str(entry)), re.S)
            if not match:
                raise ValueError('нет YAML-frontmatter')
            meta = yaml.load(match.group(1), Loader=UniqueLoader)
            if not isinstance(meta, dict) or meta.get('name') != name:
                raise ValueError('неверное имя навыка')
            if not isinstance(meta.get('description'), str) or not meta['description'].strip():
                raise ValueError('description должен быть непустой строкой')
        except Exception as exc:
            errors.append(f'{name}/SKILL.md: {exc}')
        for path in folder.rglob('*'):
            if not path.is_file() or '__pycache__' in path.parts:
                continue
            try:
                if path.suffix == '.json':
                    json.loads(_read(str(path)), object_pairs_hook=unique_pairs,
                               parse_constant=lambda x: (_ for _ in ()).throw(ValueError('Недопустимая JSON-константа: '+x)))
                elif path.suffix in ('.yaml', '.yml'):
                    yaml.load(_read(str(path)), Loader=UniqueLoader)
                elif path.suffix == '.py':
                    ast.parse(_read(str(path)), filename=str(path))
                elif path.suffix == '.md':
                    content = re.sub(r'```.*?```', '', _read(str(path)), flags=re.S)
                    for link in re.findall(r'\]\(([^)]+)\)', content):
                        target = link.strip().strip('<>')
                        if re.match(r'^[A-Za-z][A-Za-z0-9+.-]*:', target) or target.startswith('#'):
                            continue
                        target = unquote(target.split('#', 1)[0])
                        if not (path.parent / target).exists():
                            errors.append(f'{path.relative_to(bundle)}: не найден Markdown-переход {link}')
            except Exception as exc:
                errors.append(f'{path.relative_to(bundle)}: {exc}')
    return errors

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Номер раздела: 07.05 · 07.05.8 · 07.10А · 07.01.5А · 14.01.10Б · 16.03.6
# Литера может стоять после ЛЮБОГО компонента, а не только после «NN.NN» — иначе «07.01.5А»
# усечётся до «07.01» и все подразделы схлопнутся в ложный «дубль номера» (поймано самопроверкой).
# Мажор — ОДНА или две цифры: `01-identity.md` (§1.5, §1.17) и `02-format.md` (§2.11, §2.24)
# нумеруются без ведущего нуля, и при `\d{2}` 47 их заголовков и все ссылки на них
# не проверялись ВООБЩЕ — ни на дубль номера, ни на разрешимость (ревизия 22.08.2026).
_SEC = r"\d{1,2}\.\d{2}[А-ЯA-Z]?(?:\.\d+[А-ЯA-Z]?)*"
# объявление подпункта: номер в начале строки (## / ** / - / § ), дальше точка или скобка
SUBSEC_DECL = re.compile(r"^[\s#*\-–—>]*§?\s?(\d{2}\.\d{2}\.\d{1,2})\s*[.)]", re.M)
SEC_REF = re.compile(r"§\s?(" + _SEC + r")")
# В заголовке может стоять диапазон («### 16.03.6–16.03.7. …») — забираем ВСЕ номера строки
SEC_HEAD_LINE = re.compile(r"^#{1,4}\s+((?:" + _SEC + r")(?:\s*[–—-]\s*" + _SEC + r")*)\.", re.M)
SEC_IN_HEAD = re.compile(_SEC)


def _norm_sec(num):
    """«2.24» → «02.24»: ведущий ноль в мажоре, чтобы ссылка и заголовок сходились
    независимо от того, как их записали."""
    parts = num.split(".")
    if parts and parts[0].isdigit():
        parts[0] = parts[0].zfill(2)
    return ".".join(parts)
# Не принимать хвост пути соседнего навыка за файл общего references/.
FILE_REF = re.compile(r"(?<![\w./\\-])references[/\\]([0-9A-Za-z_\-]+\.md)`?")
# Ссылкой НА СКРИПТ НАВЫКА считается только явный путь `scripts/xxx.py`. Голое `xxx.py` в бэктиках —
# это, как правило, ЛОКАЛЬНЫЙ тулсет владельца (22-tools-and-backup.md), его в навыке нет и быть не должно.
SCRIPT_PATH_REF = re.compile(r"scripts[/\\]([0-9A-Za-z_]+\.py)")
SCRIPT_BARE_REF = re.compile(r"`([0-9A-Za-z_]+\.py)`")
RULE_REF = re.compile(r"[Пп]равил[оаяеу]м?\s+(\d{1,2})\b")

# ── ПАСПОРТА ЖАНРОВ ───────────────────────────────────────────────────────────────────────
# Повод: знание, решающее МАРШРУТ работы, физически лежало в конце большого жанрового файла
# (§09.19.4 — «какой путь для какой формы»), а модель читает первые экраны и начинает работать.
# 14.08.2026 так была потеряна половина сессии: Ф-1 и Ф-1.2 собраны запасным координатным путём
# вместо штатных форм ГВП. Лечение структурное: у жанрового файла есть короткая ОБЯЗАТЕЛЬНАЯ
# голова из пяти блоков, и она стоит ПЕРЕД телом. Проверяем механически — иначе паспорт
# заведут в одном файле и забудут в остальных.
PASSPORT_HEAD = "## ПАСПОРТ ЖАНРА"
PASSPORT_BLOCKS = [                       # (номер блока, ключевое слово в заголовке)
    ("1", "МАРШРУТ"),
    ("2", "СОСТАВ"),
    ("3", "ФАТАЛЬНЫЕ ОШИБКИ"),
    ("4", "ПРАВИЛА ЯДРА"),
    ("5", "ГЕЙТ"),
]
# Паспорт ОБЯЗАТЕЛЕН у КАЖДОГО жанрового файла с телом (отсутствие/неполнота = релиз не выпускать).
# Очередь закрыта 17.08.2026: все 18 жанровых файлов оснащены.
PASSPORT_REQUIRED = {
    "05-genres-lifecycle.md",
    "06a-doprosy.md", "06b-ppm-opoznanie.md", "06c-osmotry-vyemka-obysk.md",
    "06d-ekspertizy.md", "06e-veshchdoki.md",
    "07a-pds-dopros.md", "07b-okonchanie-oz.md", "07c-izmenenie-obvineniya.md",
    "07d-mery-presecheniya.md",
    "08a-zaprosy.md", "08b-soprovody.md", "08c-porucheniya-povestki.md",
    "08d-peredacha-obshchie.md", "08e-raport.md",
    "09-genres-stat-cards.md",
    "18-bank-analysis.md", "23-genres-144-proverka.md",
}
# Паспорт ЖЕЛАТЕЛЕН (замечание, не блокирует). Держать список для будущих жанровых файлов.
PASSPORT_WANTED = set()
# ⛔ НЕ включать сюда файлы-КРОСС-ССЫЛКИ без собственного тела: `05b-prodlenie-proverki-krsp.md` —
# указатель на §23.03–23.05, у него нет ни маршрута, ни состава комплекта, ни своих ошибок.
# Паспорт там был бы церемонией: пять пустых блоков вместо одной строки «см. 23».
# Файлы-ИНДЕКСЫ (`06-genres-investigative.md`, `07-genres-charging.md`,
# `08-genres-correspondence.md`) — по той же причине.
# Номер раздела бывает и вида `08e.1` (буква после первой пары цифр) — иначе проверка
# «паспорт стоит первым» молча не работала бы на `08e-raport.md`.
FIRST_SECTION = re.compile(r"^## \d{2}[A-Za-zА-Яа-я]?\.\d", re.M)

# ── КАРТОЧКИ ЖАНРА ────────────────────────────────────────────────────────────────────────
# Повод (замечание владельца 17.08.2026): паспорт заводился НА ФАЙЛ, а жанров в файле до
# пятнадцати. Один маршрут и один состав комплекта на ВУД и на прекращение — бессмыслица:
# у них общего только файл. Поэтому у КАЖДОГО раздела-жанра своя короткая карточка.
# ⛔ Не всякий раздел `## NN.NN` — жанр: бывают сквозные требования, правовые нюансы, отсылки
# и техника заполнения. Такие помечаются строкой «> ⓘ Не жанр — …», иначе проверка требовала
# бы карточку там, где её нечего класть.
CARD_HEAD = "### ● Карточка жанра"
NOT_GENRE = "ⓘ Не жанр"
CARD_FIELDS = ["**Маршрут.**", "**Комплект.**", "**⛔ Фатальные ошибки.**", "**Гейт.**"]
# Файлы, где разделы — виды ДОКУМЕНТОВ. `18-bank-analysis.md` сюда НЕ входит: это слой
# фактуры для чужих документов (фабулы, ОЗ, протокол осмотра), а не перечень жанров.
CARDS_REQUIRED = PASSPORT_REQUIRED - {"18-bank-analysis.md"}
SECTION_SPLIT = re.compile(r"(?m)^## (\d{2}[A-Za-zА-Яа-я]?\.[\d.]*\d)\.")


def _check_cards(refs_dir):
    """У каждого раздела-жанра есть карточка из четырёх полей либо пометка «не жанр»."""
    problems = []
    for name in sorted(CARDS_REQUIRED):
        path = os.path.join(refs_dir, name)
        if not os.path.isfile(path):
            continue
        t = _read(path)
        marks = list(SECTION_SPLIT.finditer(t))
        for i, m in enumerate(marks):
            end = marks[i + 1].start() if i + 1 < len(marks) else len(t)
            body = t[m.end():end]
            num = m.group(1)
            if NOT_GENRE in body[:400]:
                continue
            if CARD_HEAD not in body:
                problems.append(f"references/{name}: у раздела §{num} нет карточки жанра "
                                f"и нет пометки «{NOT_GENRE}»")
                continue
            head = body[body.index(CARD_HEAD):]
            nxt = head.find("\n## ")
            head = head[:nxt] if nxt > 0 else head
            miss = [f for f in CARD_FIELDS if f not in head]
            if miss:
                problems.append(f"references/{name}: в карточке §{num} нет полей — "
                                + " · ".join(miss))
    return problems


def _check_passports(refs_dir):
    """Паспорт жанра: есть · стоит ПЕРВЫМ · полон (пять блоков) · несёт запрет Шага 3."""
    problems, notes = [], []
    if not os.path.isdir(refs_dir):
        return problems, notes
    for name in sorted(PASSPORT_REQUIRED | PASSPORT_WANTED):
        path = os.path.join(refs_dir, name)
        if not os.path.isfile(path):
            problems.append(f"паспорта: в списке значится references/{name}, а файла нет")
            continue
        t = _read(path)
        required = name in PASSPORT_REQUIRED
        pos = t.find(PASSPORT_HEAD)
        if pos < 0:
            msg = f"references/{name}: нет ПАСПОРТА ЖАНРА (Шаг 3 — голова файла из пяти блоков)"
            (problems if required else notes).append(msg)
            continue
        # паспорт обязан стоять ДО тела: иначе он не голова, а очередной раздел в середине
        first = FIRST_SECTION.search(t)
        if first and first.start() < pos:
            problems.append(f"references/{name}: ПАСПОРТ ЖАНРА стоит ПОСЛЕ раздела "
                            f"«{t[first.start():first.start() + 30].strip()}» — он должен быть "
                            f"первым, иначе его не прочитают до начала работы")
        body = t[pos:first.start()] if (first and first.start() > pos) else t[pos:]
        missing = [f"{n}. {kw}" for n, kw in PASSPORT_BLOCKS
                   if not re.search(r"^### %s\..*%s" % (n, kw), body, re.M)]
        if missing:
            msg = (f"references/{name}: в паспорте нет блоков — " + " · ".join(missing))
            (problems if required else notes).append(msg)
        if "Шаг 3" not in body:
            notes.append(f"references/{name}: в паспорте нет ссылки на Шаг 3 "
                         f"(«читается ПОЛНОСТЬЮ и ДО первого действия»)")
    return problems, notes


def _strip_docs(src):
    """Убрать строковые литералы и комментарии, сохранив нумерацию строк.

    Нужно, чтобы проверка опасных вызовов не срабатывала на ОПИСАНИИ этих вызовов
    в docstring (сам запрет объясняется в комментарии — иначе вечный ложный флаг)."""
    import io as _io
    import tokenize as _tk
    try:
        toks = list(_tk.generate_tokens(_io.StringIO(src).readline))
    except Exception:
        return src
    lines = src.split("\n")
    out = [list(l) for l in lines]
    prev = _tk.NEWLINE
    for t in toks:
        kill = t.type == _tk.COMMENT or (
            # ⚠ гасим ТОЛЬКО docstring — самостоятельный строковый литерал. Обычные строки
            # (аргументы вызова) обязаны остаться: именно по ним и ищем `Dispatch("Word.Application")`.
            t.type == _tk.STRING and prev in (_tk.NEWLINE, _tk.NL, _tk.INDENT, _tk.DEDENT))
        if t.type not in (_tk.NL, _tk.COMMENT):
            prev = t.type
        if not kill:
            continue
        (r1, c1), (r2, c2) = t.start, t.end
        for r in range(r1 - 1, min(r2, len(out))):
            a = c1 if r == r1 - 1 else 0
            b = c2 if r == r2 - 1 else len(out[r])
            for i in range(a, min(b, len(out[r]))):
                out[r][i] = " "
    return "\n".join("".join(l) for l in out)


def _read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


def check(root):
    problems, notes = [], []
    refs_dir = os.path.join(root, "references")
    scr_dir = os.path.join(root, "scripts")
    skill_md = os.path.join(root, "SKILL.md")

    md_files = {}
    for d in (root, refs_dir):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.endswith(".md"):
                md_files[os.path.join(d, f)] = _read(os.path.join(d, f))
    py_files = sorted(f for f in os.listdir(scr_dir) if f.endswith(".py")) if os.path.isdir(scr_dir) else []

    # 1. YAML
    txt = md_files.get(skill_md, "")
    m = re.match(r"^---\n(.*?)\n---\n", txt, re.S)
    if not m:
        problems.append("SKILL.md: нет YAML-frontmatter")
    else:
        body = m.group(1)
        for key in ("name:", "description:"):
            if key not in body:
                problems.append(f"SKILL.md: в YAML нет «{key}»")
        desc = body.split("description:", 1)[1].strip() if "description:" in body else ""
        if ": " in desc and not desc.startswith(('"', "'")):
            problems.append("SKILL.md: description содержит «: » и НЕ в кавычках — YAML упадёт, "
                            "навык молча не загрузится (инвариант skill-yaml-description-quoting)")
        try:
            import yaml
            yaml.safe_load(body)
        except ImportError:
            notes.append("PyYAML не установлен — YAML проверен только эвристикой")
        except Exception as e:
            problems.append(f"SKILL.md: YAML не парсится: {e}")

    all_text = "\n".join(md_files.values())

    # Навыки версии 2.0 лежат рядом с общим справочником. Проверяем переходы
    # из каждого модуля: прежняя проверка references/ их не охватывала.
    if os.path.basename(os.path.abspath(root)) == 'investigator-2-0':
        problems.extend(_check_bundle_structure(root))
        bundle = os.path.dirname(os.path.abspath(root))
        for folder in sorted(os.listdir(bundle)):
            if not (folder == 'investigator-2-0' or folder.startswith('investigator-2-0-')):
                continue
            entry = os.path.join(bundle, folder, 'SKILL.md')
            if not os.path.isfile(entry):
                problems.append(f'{folder}: отсутствует SKILL.md')
                continue
            entry_text = _read(entry)
            declared = re.search(r'^name:\s*([^\n]+)', entry_text, re.M)
            if not declared or declared.group(1).strip() != folder:
                problems.append(f'{folder}: имя навыка не совпадает с папкой')
            for link in re.findall(r'\]\(([^)]+)\)', entry_text):
                if '://' in link or link.startswith('#'):
                    continue
                target = os.path.normpath(os.path.join(os.path.dirname(entry), unquote(link.strip().strip('<>').split('#')[0])))
                if not os.path.exists(target):
                    problems.append(f'{folder}: не найден переход {link}')

    # 2. Файлы references
    existing_refs = {f for f in os.listdir(refs_dir) if f.endswith(".md")} if os.path.isdir(refs_dir) else set()
    mentioned_refs = set(FILE_REF.findall(all_text))
    for f in sorted(mentioned_refs - existing_refs):
        problems.append(f"ссылка на несуществующий файл: references/{f}")
    for f in sorted(existing_refs - mentioned_refs):
        notes.append(f"файл не упомянут нигде в тексте: references/{f}")

    # 3. Скрипты
    # плейсхолдеры из документации («…считается только явный путь `scripts/xxx.py`») — не файлы
    PLACEHOLDERS = {"xxx.py", "yyy.py", "name.py", "script.py", "имя.py", "файл.py", "nesushchestvuyushchiy.py"}
    path_refs = set(SCRIPT_PATH_REF.findall(all_text)) - PLACEHOLDERS
    bare_refs = set(SCRIPT_BARE_REF.findall(all_text))
    for f in sorted(path_refs - set(py_files)):
        problems.append(f"ссылка `scripts/{f}` — такого скрипта в навыке нет")
    for f in sorted(set(py_files) - path_refs - bare_refs - {"__init__.py"}):
        notes.append(f"скрипт не упомянут нигде в тексте: scripts/{f}")

    # 4-5. Разделы: собрать все заголовки и найти дубли
    heads = {}
    for path, t in md_files.items():
        seen = {}
        for line_nums in SEC_HEAD_LINE.findall(t):
            for num in SEC_IN_HEAD.findall(line_nums):
                num = _norm_sec(num)
                heads.setdefault(num, path)
                seen[num] = seen.get(num, 0) + 1
        for num, n in sorted(seen.items()):
            if n > 1:
                problems.append(f"{os.path.basename(path)}: номер раздела §{num} встречается {n} раз(а)")

    refd = {_norm_sec(x) for x in SEC_REF.findall(all_text)}
    unresolved = sorted(n for n in refd if n not in heads)
    # §NN.NN.N может ссылаться на подпункт внутри раздела §NN.NN — считаем разрешённой,
    # если существует любой заголовок-префикс
    # §NN.NN.N — подпункт внутри раздела. РАНЬШЕ: считался разрешённым, если есть заголовок-префикс
    # §NN.NN, то есть выдуманный номер подпункта (§08.17.9 при живом §08.17) гейт пропускал молча —
    # проверено экспериментом 05.09.2026. ТЕПЕРЬ: требуем, чтобы подпункт был где-то ОБЪЯВЛЕН —
    # то есть номер стоит в начале строки (заголовок, жирный лид, пункт списка) и за ним точка/скобка.
    declared3 = set(SUBSEC_DECL.findall(all_text))
    truly, orphan3 = [], []
    for n in unresolved:
        parent = any(h == n or n.startswith(h + ".") for h in heads)
        if not parent:
            truly.append(n)
        elif n.count(".") == 2 and n not in declared3:
            orphan3.append(n)
    for n in truly:
        problems.append(f"ссылка на несуществующий раздел: §{n}")
    for n in orphan3:
        problems.append(f"ссылка на подпункт §{n}: раздел §{n.rsplit('.', 1)[0]} есть, "
                        f"а самого подпункта {n} нигде не объявлено")

    # 6. После разделения на навыки полный реестр правил находится в 24-rules.md.
    # Проверяем существование каждого номера, а не только верхнюю границу.
    rules_path = os.path.join(refs_dir, "24-rules.md")
    rules_text = md_files.get(rules_path, txt)
    rules = {int(r) for r in re.findall(r"^(\d{1,2})\.\s+", rules_text, re.M)}
    for r in sorted({int(x) for x in RULE_REF.findall(all_text)}):
        if r not in rules:
            problems.append(f"ссылка на «правило {r}», но его нет в реестре правил")

    # 6-bis. Нумерация версий в CHANGELOG — ДЕСЯТИЧНАЯ (минор одна цифра, после .9 растёт мажор).
    # Ловушка добавлена после ВТОРОГО нарушения этого правила (v16.9→v16.10, затем v17.9→v17.10):
    # документированного правила оказалось мало — минор наращивался механически, без проверки границы.
    chlog = os.path.join(root, "CHANGELOG.md")
    if os.path.isfile(chlog):
        ct = _read(chlog)
        vers = re.findall(r"^##\s+v(\d+)\.(\d+)", ct, re.M)
        for maj, mnr in vers:
            if len(mnr) > 1:
                problems.append(f"CHANGELOG: версия v{maj}.{mnr} — минор ДВУЗНАЧНЫЙ, нарушена "
                                f"десятичная схема; после .9 растёт мажор (должно быть v{int(maj)+1}.0)")
        seen_v, dup = set(), []
        for maj, mnr in vers:
            key = (int(maj), int(mnr))
            if key in seen_v:
                dup.append(f"v{maj}.{mnr}")
            seen_v.add(key)
        for d in sorted(set(dup)):
            problems.append(f"CHANGELOG: версия {d} встречается дважды")
        if vers:
            ordered = [(int(a), int(b)) for a, b in vers]
            if ordered != sorted(ordered, reverse=True):
                notes.append("CHANGELOG: версии идут не строго по убыванию — проверить порядок записей")

    # 7. Мусорные символы
    # ⛔ Скрипты проверяются НАРАВНЕ с md: 08.09.2026 в make_povestka.py прошёл
    # гейтом символ 0x08 внутри регулярного выражения — оно перестало совпадать МОЛЧА,
    # и повестки уходили с годом «20__». Проверка смотрела только md_files.
    _ctrl = dict(md_files)
    for _f in py_files:
        _p = os.path.join(scr_dir, _f)
        _ctrl[_p] = _read(_p)
    for path, t in _ctrl.items():
        bad = sorted({hex(ord(c)) for c in t if ord(c) < 32 and c not in "\n\t"})
        if bad:
            problems.append(f"{os.path.basename(path)}: управляющие символы {bad}")
        if t.startswith("﻿"):
            problems.append(f"{os.path.basename(path)}: BOM в начале файла")

    # 15. Строки "text" в ```python-шаблонах references — то, что генератор кладёт В ДОКУМЕНТ.
    # 09.09.2026: 27 длинных тире вне плейсхолдеров и донорский хвост с «ст. ст.» жили в шаблонах
    # шести-семи жанров, и линтер ловил их уже В ДОКУМЕНТЕ, на каждом документе заново (правило 3,
    # флаг `сдвоенное-ст`). Плейсхолдер {…} — подсказка, а не текст: внутри него тире допустимо.
    _tpl_str = re.compile(r'"text":\s*"((?:[^"\\]|\\.)*)"')
    _dbl_st = re.compile(r"\bст\.\s*ст\.")
    for path, t in md_files.items():
        if os.path.dirname(path) != refs_dir:
            continue
        for blk in re.findall(r"```python\n(.*?)```", t, re.S):
            for m in _tpl_str.finditer(blk):
                s = m.group(1)
                depth, outside = 0, []
                for ch in s:
                    if ch == "{":
                        depth += 1
                    elif ch == "}":
                        depth = max(0, depth - 1)
                    elif depth == 0:
                        outside.append(ch)
                out = "".join(outside)
                if "—" in out:
                    problems.append(f"{os.path.basename(path)}: длинное тире в строке python-шаблона "
                                    f"(уйдёт в документ, правило 3): «{s[:70]}»")
                if _dbl_st.search(out):
                    problems.append(f"{os.path.basename(path)}: «ст. ст.» в строке python-шаблона "
                                    f"(линтер `сдвоенное-ст`): «{s[:70]}»")

    # 8. Компиляция скриптов
    for f in py_files:
        p = os.path.join(scr_dir, f)
        src = _read(p)
        try:
            compile(src, p, "exec")
        except SyntaxError as e:
            problems.append(f"scripts/{f}: синтаксическая ошибка — строка {e.lineno}: {e.msg}")
        if not ast.get_docstring(ast.parse(src)):
            notes.append(f"scripts/{f}: нет модульного docstring")

    # 9. ОПАСНАЯ АВТОМАТИЗАЦИЯ OFFICE (правило 39).
    # Повод: 14.08.2026 fill_formfields.py через Dispatch подключался к РАБОЧЕМУ Word
    # владельца и в __exit__ делал Quit() при DisplayAlerts=0 — закрывал его документы
    # без сохранения. Ловим механически, чтобы не вернулось.
    DANGER = [
        (r'Dispatch\(\s*["\'](?:Word|Excel|PowerPoint)\.Application',
         'Dispatch подключается к РАБОЧЕМУ Office пользователя — только DispatchEx (свой процесс)'),
        (r'taskkill[^\n]*["\']/IM["\']',
         'taskkill /IM убьёт и рабочий Office владельца — снимать только свои PID и только окна без заголовка'),
    ]
    for f in py_files:
        src = _read(os.path.join(scr_dir, f))
        code = _strip_docs(src)            # без строк и комментариев: ищем ВЫЗОВЫ, а не текст о них
        for pat, why in DANGER:
            for m in re.finditer(pat, code):
                problems.append(f"scripts/{f}:{code[:m.start()].count(chr(10)) + 1}: {why}")

    # ⚠ И в references тоже: модель копирует РЕЦЕПТЫ из ```python-блоков дословно.
    # Именно так опасный `Dispatch("Excel.Application")` пережил проверку скриптов
    # и 14.08.2026 трижды закрыл Excel владельца (§09.18.2).
    for path in md_files:
        text = _read(path)
        for block in re.finditer(r"```(?:python|py)\s*\n(.*?)```", text, re.S):
            body = block.group(1)
            for pat, why in DANGER:
                for m in re.finditer(pat, body):
                    line = text[:block.start(1)].count("\n") + body[:m.start()].count("\n") + 1
                    problems.append(f"{os.path.basename(path)}:{line} (пример кода): {why}")

    # 10. ПУТИ К ФАЙЛАМ НА ДИСКЕ, упомянутые в тексте навыка (в обратных кавычках).
    # check_skill исторически проверял только ВНУТРЕННИЕ ссылки (references/scripts/§/правила),
    # а внешние пути устаревали молча — после переездов папок навык месяцами вёл к небытию
    # (реестр без `вх\`, бланк 384 ВСО, папка «ПКГ ВСО»).
    # ⚠ ИСКЛЮЧЕНИЕ: путь, НАМЕРЕННО помеченный как мёртвый («мёртв», «прежний адрес»,
    # «больше нет», «устаревш»), — это предупреждение владельцу, а не ошибка. Без исключения
    # получаем вечный ложный флаг (проверено: такой путь в навыке ровно один —
    # `Примеры Казань\Бланк 384 ВСО СК России.docx`, о нём прямо написано, что он мёртв).
    # Локальный пример исключён из публичной поставки.
    # абсолютные C:\ — с плейсхолдерами {…}/<…>/* не связываемся.
    base = os.environ.get("KRASNAYA_BASE", "")
    DEAD_MARK = re.compile(r"мёртв|мертв|прежний адрес|прежний путь|больше нет|устаревш|НЕ существует",
                           re.I)
    PLACEHOLDER = ("{", "<", "*", "NN", "vNN")
    if os.path.isdir(base):
        # «…\» в текстах означает РАЗНОЕ: корень рабочей папки, каталог ВЫШЕ него
        # Локальный пример исключён из публичной поставки.
        # Путь считается живым, если существует хотя бы при одном прочтении, — иначе
        # получаем 12 ложных флагов на чистом навыке (замерено 14.08.2026).
        roots = [base, os.path.dirname(base), os.path.expanduser("~"),
                 os.environ.get("ProgramFiles", r"C:\Program Files")]
        seen = set()
        for path in md_files:
            # ⛔ CHANGELOG не проверяем: там пути — ИСТОРИЧЕСКИЙ снимок («переехало оттуда сюда»),
            # старый адрес в записи о переезде обязан остаться и ошибкой не является.
            if os.path.basename(path).lower() == "changelog.md":
                continue
            for ln, line_txt in enumerate(_read(path).splitlines(), 1):
                if DEAD_MARK.search(line_txt):
                    continue                      # намеренно мёртвый путь — не флажим
                for raw in re.findall(r"`([^`\n]+)`", line_txt):
                    raw = raw.strip()
                    if "\\" not in raw or any(x in raw for x in PLACEHOLDER):
                        continue
                    tail = raw[2:] if raw.startswith("…\\") else raw
                    if "…" in tail:               # многоточие ВНУТРИ — это шаблон, не путь
                        continue
                    if raw.startswith("…\\"):
                        cands = [os.path.join(r, tail) for r in roots if r]
                    elif re.match(r"^[A-Za-z]:\\", raw):
                        cands = [raw]
                    else:
                        continue
                    cands = [c.rstrip("\\") for c in cands]
                    if any(os.path.exists(c) for c in cands):
                        continue
                    if cands[0] in seen:
                        continue
                    seen.add(cands[0])
                    notes.append(f"{os.path.basename(path)}:{ln}: путь не найден на диске — "
                                 f"`{raw}` (если он мёртв намеренно — пометь в строке словом "
                                 f"«мёртв»/«прежний адрес», и проверка его пропустит)")

    # 11. ПАСПОРТА ЖАНРОВ (см. комментарий у PASSPORT_REQUIRED)
    p_pr, p_no = _check_passports(refs_dir)
    problems.extend(p_pr)
    notes.extend(p_no)

    # 11-bis. КАРТОЧКИ ЖАНРА у каждого раздела-жанра (см. комментарий у CARD_HEAD)
    problems.extend(_check_cards(refs_dir))

    # 12. ДУБЛИ НОМЕРОВ РАЗДЕЛОВ МЕЖДУ ФАЙЛАМИ.
    # Проверка 5 смотрела только ВНУТРИ файла, и две коллизии жили незамеченными до 17.08.2026:
    # §06.21 был занят и «получением образцов» (06e), и «общими требованиями» (06f);
    # §08.23 — и «чужой регистрацией» (08d, на него ссылается правило 25 из пяти мест),
    # и «сопроводом по приостановленному делу» (08b). Голая ссылка «§08.23» вела в два места
    # сразу — ровно та «гниль ссылок», ради которой писался этот скрипт.
    by_num = {}
    for path, t in md_files.items():
        if os.path.basename(path).lower() == "changelog.md":
            continue                      # исторический снимок, старые номера там законны
        for num in re.findall(r"^##\s+(" + _SEC + r")\.", t, re.M):
            by_num.setdefault(num, set()).add(os.path.basename(path))
    for num, files in sorted(by_num.items()):
        if len(files) > 1:
            problems.append(f"номер раздела §{num} занят в НЕСКОЛЬКИХ файлах ({', '.join(sorted(files))}) "
                            f"— ссылка «§{num}» ведёт в два места сразу; переномеруй тот файл, "
                            f"на который не ссылаются извне")

    # 13. ССЫЛКИ НЕСТАНДАРТНОЙ ФОРМЫ: §08f.01, §08e.2, §07.10А — буква стоит после первой пары
    # цифр или в конце. Регулярка _SEC их не берёт, то есть 8 ссылок на §08f.01 никем не
    # проверялись. Разрешаем по заголовкам-«как есть».
    odd_heads = set()
    for path, t in md_files.items():
        for h in re.findall(r"^#{1,4}\s+([0-9]{2}[A-Za-zА-Яа-я]?\.[0-9A-Za-zА-Яа-я.]*)", t, re.M):
            odd_heads.add(h.rstrip("."))
    ODD_REF = re.compile(r"§\s?(\d{2}[A-Za-zА-Яа-я]\.[\d.]*\d|\d{2}\.\d+[A-Za-zА-Яа-я])")
    for r in sorted({m for path, t in md_files.items() for m in ODD_REF.findall(t)}):
        if not any(h == r or h.startswith(r + ".") for h in odd_heads):
            problems.append(f"ссылка на несуществующий раздел: §{r} (форма с буквой — "
                            f"проверка 4 её не видит)")

    # 14. ФАЙЛЫ, КОТОРЫЕ РАНЬШЕ НЕ ПРОВЕРЯЛИСЬ ВООБЩЕ (ревизия 22.08.2026):
    #     константы.json, адресаты.json, scripts/requirements.txt, шаблон бланка.
    import json as _json
    for jname in ("константы.json", "адресаты.json"):
        jp = os.path.join(root, jname)
        if not os.path.isfile(jp):
            notes.append(f"нет файла персонального слоя: {jname} "
                         f"(генератор возьмёт фолбэк с плейсхолдерами)")
            continue
        try:
            data = _json.loads(_read(jp))
        except Exception as e:
            problems.append(f"{jname}: не парсится как JSON ({e}) — генератор молча "
                            f"уйдёт на фолбэк с плейсхолдерами")
            continue
        if not isinstance(data, dict) or not data:
            problems.append(f"{jname}: ожидался непустой объект JSON")

    req = os.path.join(scr_dir, "requirements.txt")
    if not os.path.isfile(req):
        problems.append("нет scripts/requirements.txt — на новой машине нечем поднять окружение")
    else:
        # Имена пакетов берём из СТРОК, отбросив комментарии: закомментированный
        # «# pywin32» не должен считаться указанным.
        declared = set()
        for line in _read(req).splitlines():
            line = line.split("#", 1)[0].strip()
            if not line:
                continue
            declared.add(re.split(r"[;<>=\[\s!]", line, 1)[0].strip().lower())
        # Пакеты, без которых скрипты не запустятся. Ключ — импорт, значение — имя на PyPI.
        need = {"docx": "python-docx", "fitz": "pymupdf", "win32com": "pywin32"}
        srcs = {f: _read(os.path.join(scr_dir, f)) for f in py_files}
        for mod, pipname in need.items():
            used = any(re.search(r"(?:^|\W)import\s+%s\b|from\s+%s\b" % (mod, mod), t)
                       for t in srcs.values())
            if used and pipname.lower() not in declared:
                problems.append(f"requirements.txt: не указан «{pipname}», хотя скрипты "
                                f"импортируют «{mod}» — на новой машине путь молча отвалится")

    tpl = os.path.join(scr_dir, "templates", "blank_384_vso.docx")
    if os.path.isfile(tpl):
        import zipfile as _zf
        if not _zf.is_zipfile(tpl):
            problems.append("scripts/templates/blank_384_vso.docx — не .docx (битый zip)")
    else:
        notes.append("Локальный гербовый бланк не включён в публичную поставку; без его подключения документы на бланке "
                     "собрать не получится")

    # 15. У КАЖДОГО CLI-скрипта должна быть справка: `--help` печатает и выходит,
    #     а не трактует флаг как путь и не создаёт файл-мусор (ревизия 22.08.2026).
    for f in py_files:
        t = _read(os.path.join(scr_dir, f))
        if '__name__ == "__main__"' not in t and "__name__ == '__main__'" not in t:
            continue                      # библиотека без CLI — справка не нужна
        if "--help" not in t and "ArgumentParser" not in t:
            problems.append(f"scripts/{f}: нет разбора «--help» — флаг уйдёт в позиционный "
                            f"аргумент (создаст файл/папку с именем «--help»)")

    return problems, notes, len(md_files), len(py_files)


def main():
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        print("Запуск:  python check_skill.py [корень навыка]")
        print("Без аргумента корнем считается папка на уровень выше scripts/.")
        print("Код возврата: 0 — чисто, 1 — есть проблемы, 2 — это не корень навыка.")
        return 0
    root = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if not os.path.isdir(os.path.join(root, "references")):
        print(f"Не похоже на корень навыка (нет references/): {root}")
        return 2
    problems, notes, n_md, n_py = check(root)
    print(f"check_skill: {n_md} md-файлов, {n_py} скриптов\n")
    if problems:
        print(f"ПРОБЛЕМЫ ({len(problems)}) — чинить до релиза:")
        for p in problems:
            print("  ✗ " + p)
    else:
        print("Проблем не найдено.")
    if notes:
        print(f"\nЗамечания ({len(notes)}) — не блокируют:")
        for n in notes:
            print("  · " + n)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
