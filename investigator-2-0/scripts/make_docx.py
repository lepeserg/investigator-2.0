# -*- coding: utf-8 -*-
"""
Генератор .docx с правильным форматированием для процессуальных документов ВСО
СК России (формат по Приказу СК России № 40).

ЭТАЛОННОЕ ЯДРО. Персональные реквизиты (наименование органа, подписанты,
прокурор, адрес, e-mail, телефон исполнителя) НЕ зашиты в код — они читаются из
файла констант «константы.json» персонального слоя. Путь к нему берётся из:
    1) аргумента функций (constants=...);
    2) переменной окружения SK_CONSTANTS;
    3) файла «константы.json» рядом со скриптом / на два уровня выше (папка _система);
    4) фолбэк-словаря с ПЛЕЙСХОЛДЕРАМИ ({ОРГАН_ПОЛН}, {СЛЕДОВАТЕЛЬ_ФИО}, …),
       чтобы генератор не падал без констант, а выдавал документ с явными
       метками «подставить реквизит».

Две основные функции:
    build_doc()             – обычные процессуальные документы на чистом листе
                              (постановления, ходатайства, ОЗ).
    build_letterhead_doc()  – документы на бланке органа (с гербом и реквизитами):
                              запросы, сопроводы, повестки, справки, уведомления,
                              поручения, представления.

Использование (через analysis tool в claude.ai):
    from make_docx import build_doc
    build_doc(
        out_path="/path/to/result.docx",
        body_blocks=[
            {"type": "header_centered", "text": "ПОСТАНОВЛЕНИЕ"},
            {"type": "header_centered_thin", "text": "о возбуждении уголовного дела и принятии его к производству"},
            {"type": "place_date", "place": "г. {ГОРОД}", "date": "13.04.2026", "time": "11 ч. 50 мин."},
            {"type": "para", "text": "Старший следователь-криминалист ..."},
            {"type": "section", "text": "УСТАНОВИЛ:"},
            {"type": "para", "text": "Поводом для возбуждения уголовного дела ..."},
            {"type": "section", "text": "ПОСТАНОВИЛ:"},
            {"type": "para_no_indent", "text": "1. Возбудить уголовное дело ..."},
            {"type": "signature", "kind": "следователь"},  # или "руководитель", "заместитель"
        ]
    )

Параметры стиля (Приказ СК России № 40):
    шрифт: Times New Roman 13
    интервал: одинарный (1.0)
    поля: берутся из константы.json (Казань: левое 2,5 / верхнее 2 / нижнее 2 / правое 1 см)
    абзацный отступ: 1,25 см
"""

import os
import json
import shutil
import time
import tempfile
import zipfile
import re
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ============================================================================
# Константы персонального слоя (наименование органа, подписанты, реквизиты)
# ============================================================================
#
# Фолбэк-словарь с ПЛЕЙСХОЛДЕРАМИ: если «константы.json» не найден, документ всё
# равно собирается, но с явными метками «{...}» — их видно и не спутать с фактом.
# Ключи гранулярные (для разбивки подписи на строки); из плоской схемы
# «константы.json» (лицо одной строкой) они выводятся в load_constants().
_CONSTANTS_FALLBACK = {
    "орган_полное": "{ОРГАН_ПОЛН}",
    "орган_родительный": "{ОРГАН_ПОЛН_РОД}",
    "орган_сокращённо": "{ОРГАН_СОКР}",
    "город": "{ГОРОД}",
    "орган_адрес": "{АДРЕС_ОРГАНА}",
    "орган_email": "{EMAIL_ОРГАНА}",
    # Следователь
    "следователь_должность": "{СЛЕДОВАТЕЛЬ_ДОЛЖНОСТЬ}",
    "следователь_звание": "{СЛЕДОВАТЕЛЬ_ЗВАНИЕ}",
    "следователь_фио": "{СЛЕДОВАТЕЛЬ_ФИО}",
    # Руководитель
    "руководитель_должность": "Руководитель",
    "руководитель_звание": "{РУКОВОДИТЕЛЬ_ЗВАНИЕ}",
    "руководитель_фио": "{РУКОВОДИТЕЛЬ_ФИО}",
    # Заместитель руководителя
    "заместитель_должность": "Заместитель руководителя",
    "заместитель_звание": "{ЗАМ_ЗВАНИЕ}",
    "заместитель_фио": "{ЗАМ_РУКОВОДИТЕЛЯ}",
    # Прокуратура (надзор / утверждение ОЗ)
    "прокуратура_надзорная": "{ПРОКУРАТУРА_НАДЗ}",
    "прокурор_должность": "Военный прокурор",
    "прокурор_звание": "{ПРОКУРОР_ЗВАНИЕ}",
    "прокурор_фио": "{ПРОКУРОР_ФИО}",
    # Исполнитель в колонтитуле исходящих (запросы/поручения)
    "исполнитель_имя_отчество": "{ИСПОЛНИТЕЛЬ_ИМЯ}",
    "исполнитель_телефон": "{ТЕЛ_ИСПОЛНИТЕЛЯ}",
    # Год для грифов «___ ____ 20__ г.»
    "год": "20__",
}

# Кэш загруженных констант по пути (чтобы не читать файл на каждый блок).
_CONSTANTS_CACHE = {}


def _find_constants_path(explicit=None):
    """Определяет путь к «константы.json»: аргумент -> SK_CONSTANTS -> рядом со
    скриптом / на два уровня выше (папка _система) -> None (тогда фолбэк)."""
    requested = explicit or os.environ.get("SK_CONSTANTS")
    if requested:
        if not os.path.isfile(requested):
            raise FileNotFoundError("Не найден явно выбранный профиль реквизитов: " + str(requested))
        return os.path.abspath(requested)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(here, "константы.json"),
        os.path.join(here, "..", "константы.json"),
        os.path.join(here, "..", "..", "константы.json"),
        os.path.join(here, "..", "..", "_система", "константы.json"),
    ]
    for c in candidates:
        if os.path.isfile(c):
            return os.path.abspath(c)
    return None


def _split_person(value):
    """Разбирает свободную строку реквизита лица «Фамилия И.О., звание, должность»
    в (фио, звание, должность). Плоская схема «константы.json» хранит
    следователя/руководителя/прокурора одной строкой через запятую; порядок
    полей — ФИО, звание, [должность]. Пустые поля -> "". Устойчив к недостаче
    компонентов (для руководителя должность часто не указана)."""
    if not value:
        return "", "", ""
    parts = [p.strip() for p in str(value).split(",")]
    fio = parts[0] if len(parts) > 0 else ""
    zvanie = parts[1] if len(parts) > 1 else ""
    dolzhnost = parts[2] if len(parts) > 2 else ""
    return fio, zvanie, dolzhnost


def load_constants(constants_path=None):
    """Загружает реквизиты персонального слоя из «константы.json» (или фолбэк).

    Возвращает НОРМАЛИЗОВАННЫЙ словарь с гранулярными ключами, которые нужны
    генератору для сборки подписных блоков и грифов. Если исходный JSON хранит
    лицо одной строкой (следователь/руководитель/прокурор) — раскладывает её на
    ФИО/звание/должность. Отсутствующие значения заменяются плейсхолдерами из
    _CONSTANTS_FALLBACK, так что документ всегда собирается."""
    path = _find_constants_path(constants_path)
    explicit_profile = bool(constants_path or os.environ.get('SK_CONSTANTS'))
    stamp = (os.stat(path).st_mtime_ns, os.stat(path).st_size) if path else None
    cache_key = (path or "__fallback__", explicit_profile, stamp)
    if cache_key in _CONSTANTS_CACHE:
        return _CONSTANTS_CACHE[cache_key]

    c = dict(_CONSTANTS_FALLBACK)
    raw = {}
    if path:
        try:
            with open(path, "r", encoding="utf-8") as f:
                raw = json.load(f)
        except (OSError, ValueError) as exc:
            raise ValueError('Не удалось прочитать профиль реквизитов: ' + str(path)) from exc
        if not isinstance(raw, dict):
            raise ValueError('Профиль реквизитов должен быть JSON-объектом: ' + str(path))
        if raw.get('автоматические_реквизиты_дела') is False and not explicit_profile:
            # Historical identities remain available to explicit legacy callers only.
            # Do not silently assign a prior investigator/authority to a new case.
            raw = {}

    # Прямые гранулярные ключи (если персональный слой сразу хранит гранулярно).
    for k in list(_CONSTANTS_FALLBACK.keys()):
        if raw.get(k):
            c[k] = raw[k]

    # Совместимость с плоской схемой «константы.json»: орган/адрес/email/лица
    # одной строкой. Раскладываем на гранулярные поля, если те не заданы явно.
    if raw.get("орган_полное"):
        c["орган_полное"] = raw["орган_полное"]
    if raw.get("орган_адрес"):
        c["орган_адрес"] = raw["орган_адрес"]
    if raw.get("орган_email"):
        c["орган_email"] = raw["орган_email"]

    def _person_fields(value):
        'Вспомогательная функция. Частный пример исключён из публичной версии.'
        if isinstance(value, dict):
            fio = (value.get("инициалы_фам") or value.get("фио") or "").strip()
            return fio, (value.get("звание") or "").strip(), (value.get("должность") or "").strip()
        return _split_person(value)

    def _fill(person_keys, dolzh_key, zvanie_key, fio_key, default_dolzh=None):
        # person_keys: ключ ИЛИ список ключей-кандидатов (берётся первый найденный).
        keys = [person_keys] if isinstance(person_keys, str) else list(person_keys)
        val = next((raw[k] for k in keys if raw.get(k)), None)
        if val is not None:
            fio, zvanie, dolzh = _person_fields(val)
            if fio:
                c[fio_key] = fio
            if zvanie:
                c[zvanie_key] = zvanie
            if dolzh:
                c[dolzh_key] = dolzh
            elif default_dolzh and not raw.get(dolzh_key):
                c[dolzh_key] = default_dolzh

    _fill("следователь", "следователь_должность", "следователь_звание", "следователь_фио")
    _fill("руководитель", "руководитель_должность", "руководитель_звание", "руководитель_фио",
          default_dolzh="Руководитель")
    _fill(["зам", "руководитель_заместитель"], "заместитель_должность", "заместитель_звание",
          "заместитель_фио", default_dolzh="Заместитель руководителя")
    # Прокурор: вложенный объект {фио, звание, должность} или плоская строка.
    if raw.get("прокурор"):
        fio, zvanie, dolzh = _person_fields(raw["прокурор"])
        if fio:
            c["прокурор_фио"] = fio
        if zvanie:
            c["прокурор_звание"] = zvanie
        if dolzh:
            c["прокурор_должность"] = dolzh
            # Надзорная прокуратура (орган) — из явного ключа «прокуратура»,
            # иначе производно от должности: «военный прокурор X гарнизона»
            # -> «военная прокуратура X гарнизона».
            nadz = raw["прокурор"].get("прокуратура") if isinstance(raw["прокурор"], dict) else None
            c["прокуратура_надзорная"] = nadz or dolzh.replace("военный прокурор", "военная прокуратура")

    _CONSTANTS_CACHE[cache_key] = c
    return c


def _fio_last_first(fio):
    'Вспомогательная функция. Частный пример исключён из публичной версии.'
    fio = (fio or "").strip()
    if not fio:
        return ""
    m = re.match(r"^([А-ЯЁA-Z][а-яёa-z\-]+)\s+([А-ЯЁA-Z]\.\s*[А-ЯЁA-Z]\.?)\s*$", fio)
    if m:
        return (m.group(2).replace(" ", "") + " " + m.group(1)).strip()
    return fio


# ============================================================================
# Неразрывные пробелы (борьба с висячими предлогами) — §2.14
# ============================================================================
_NBSP = " "
_SHORT_WORDS = [
    "в", "и", "с", "к", "о", "у", "а", "я", "во", "со", "ко", "об", "изо",
    "на", "по", "за", "из", "от", "до", "не", "ни", "же", "бы", "ли", "то",
    "для", "под", "над", "при", "без", "про", "но", "или", "да", "что", "как",
]
_ABBR = ["г", "ул", "д", "кв", "пос", "стр", "обл", "ст", "ч", "п", "пп", "абз", "т", "л", "руб", "коп"]


def _apply_nbsp(text):
    """ОТКЛЮЧЕНО по требованию пользователя: неразрывный пробел (U+00A0) в текст
    документов НЕ вставляем. Висячие предлоги пользователь убирает вручную
    комбинацией Shift+Enter (мягкий перенос) при вычитке. Функция оставлена как
    no-op, чтобы не трогать места вызова."""
    return text


def _normalize_text(text):
    """Финальная нормализация текста под формат ВСО (Приказ № 40 + формат-профиль):
    (1) тире — ТОЛЬКО короткое «–» (U+2013); длинное «—»/«―» и «--» запрещены;
        обычные дефисы внутри слов (в/ч, ПОВСК-2) не трогаем;
    (2) ё→е, Ё→Е — единая финальная замена по всему тексту (правило `02a-format-profile`
        «Правило ё→е»); раньше профиль это обещал, но генератор не делал (аудит 18.07.2026);
    (3) английские двойные кавычки “ ” и „ ‟ → ёлочки « »."""
    if not text:
        return text
    s = (text.replace("—", "–")
             .replace("―", "–")
             .replace("--", "–")
             .replace("ё", "е")
             .replace("Ё", "Е")
             .replace("“", "«").replace("”", "»")
             .replace("„", "«").replace("‟", "»"))
    if '"' in s:  # прямые ASCII-кавычки → ёлочки, чередуя открывающую/закрывающую в пределах текста
        out, open_q = [], True
        for ch in s:
            if ch == '"':
                out.append("«" if open_q else "»")
                open_q = not open_q
            else:
                out.append(ch)
        s = "".join(out)
    return s


# ============================================================================
# Параметры формата (Приказ СК России № 40)
# ============================================================================
FONT_NAME = "Times New Roman"
FONT_SIZE = Pt(13)
LEFT_MARGIN = Cm(2.5)
TOP_MARGIN = Cm(2.0)
BOTTOM_MARGIN = Cm(2.0)
RIGHT_MARGIN = Cm(1.0)
FIRST_LINE_INDENT = Cm(1.25)

# Размер страницы – A4 (КРИТИЧНО: без явной установки python-docx ставит US Letter).
PAGE_WIDTH = Cm(21.0)
PAGE_HEIGHT = Cm(29.7)

# Расстояние от края страницы до колонтитула (верхний и нижний).
HEADER_DISTANCE = Cm(0.5)
FOOTER_DISTANCE = Cm(0.5)

# Выравнивание основного текста.
BODY_ALIGN = WD_ALIGN_PARAGRAPH.JUSTIFY

# Междустрочный интервал «точно» для подписей/реквизитов адресата.
LINE_EXACT = Pt(12)

# --- Параметры бланка (letterhead) ---
# Левый отступ блока адресата на бланке (справа, на уровне углового штампа).
ADDRESSEE_INDENT = Cm(7.0)
# Сколько строк отвести под шапку бланка до начала текста.
LETTERHEAD_LEADING_LINES = 15
# Опустить блок адресата на N строк, чтобы он встал на уровень текста «СК РОССИИ».
ADDRESSEE_DROP_LINES = 2

# Путь к бланку по умолчанию (templates/blank_384_vso.docx рядом со скриптом).
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_LETTERHEAD_PATH = os.path.join(_SCRIPT_DIR, "templates", "blank_384_vso.docx")


# ============================================================================
# Подписные блоки и грифы — СТРОЯТСЯ из констант персонального слоя
# ============================================================================

def _role_org_lines(dolzh, org_gen):
    """Двухстрочная шапка подписи «Роль / орган (род. падеж)» — как в эталоне
    01-identity §1.6:
        Следователь по особо важным делам
        384 военного следственного отдела СК России
    Строка 1 — роль С ЗАГЛАВНОЙ буквы, БЕЗ отдела. Строка 2 — орган в
    РОДИТЕЛЬНОМ падеже (`орган_родительный`), одинаковый для всех ролей.
    Должности руководителя/зама в константах содержат отдел («… 384 …отдела…»)
    — он отбрасывается из строки 1, чтобы отдел не задвоился строкой органа."""
    m = re.match(r"^(.*?\S)\s+\d.*$", dolzh or "")
    role = m.group(1) if m else (dolzh or "")
    if role:
        role = role[:1].upper() + role[1:]
    return [role, org_gen]


def _signature_blocks(c):
    """Возвращает словарь подписных блоков (список строк) из констант c.
    Строка «звание\\tФамилия И.О.» выравнивает ФИО по правому полю (в _add_signature)."""
    org_gen = c.get("орган_родительный", c.get("орган_полное", "{ОРГАН_ПОЛН}"))

    def _block(dolzh, zvanie, fio):
        return _role_org_lines(dolzh, org_gen) + [
            "",
            "%s\t\t\t\t\t\t\t\t%s" % (zvanie, _fio_last_first(fio)),
        ]

    return {
        "следователь": _block(
            c.get("следователь_должность", "{СЛЕДОВАТЕЛЬ_ДОЛЖНОСТЬ}"),
            c.get("следователь_звание", "{СЛЕДОВАТЕЛЬ_ЗВАНИЕ}"),
            c.get("следователь_фио", "{СЛЕДОВАТЕЛЬ_ФИО}"),
        ),
        "руководитель": _block(
            c.get("руководитель_должность", "Руководитель"),
            c.get("руководитель_звание", "{РУКОВОДИТЕЛЬ_ЗВАНИЕ}"),
            c.get("руководитель_фио", "{РУКОВОДИТЕЛЬ_ФИО}"),
        ),
        "заместитель": _block(
            c.get("заместитель_должность", "Заместитель руководителя"),
            c.get("заместитель_звание", "{ЗАМ_ЗВАНИЕ}"),
            c.get("заместитель_фио", "{ЗАМ_РУКОВОДИТЕЛЯ}"),
        ),
    }


def _prosecutor_role_lines(dolzh):
    """«военный прокурор Казанского гарнизона» -> [«Военный прокурор», «Казанского гарнизона»]
    по эталону 01-identity §1.6: роль С ЗАГЛАВНОЙ, гарнизон отдельной строкой. Строку
    «военная прокуратура …» в гриф НЕ добавлять — она нужна только адресату/органу
    (иначе задвоение «Казанского гарнизона» и строчная буква — баг, найденный аудитом 18.07.2026)."""
    d = (dolzh or "").strip()
    m = re.match(r"(?iu)^(военный\s+прокурор|прокурор)\s+(.+)$", d)
    if m:
        role = m.group(1)
        role = role[:1].upper() + role[1:].lower()
        return [role, m.group(2).strip()]
    return [d[:1].upper() + d[1:]] if d else [d]


def _prosecutor_approval(c):
    """Гриф «УТВЕРЖДАЮ» прокурора (для ОЗ) из констант. Роль+гарнизон — двумя строками
    по эталону §1.6, БЕЗ задвоения строкой «военная прокуратура …» (см. _prosecutor_role_lines)."""
    return [
        "«УТВЕРЖДАЮ»",
        "",
    ] + _prosecutor_role_lines(c.get("прокурор_должность", "Военный прокурор")) + [
        "",
        "%s\t\t\t\t%s" % (
            c.get("прокурор_звание", "{ПРОКУРОР_ЗВАНИЕ}"),
            _fio_last_first(c.get("прокурор_фио", "{ПРОКУРОР_ФИО}")),
        ),
        "",
        "«___» _____________ %s г." % c.get("год", "20__"),
    ]


def _leader_approval(c):
    """Гриф «СОГЛАСЕН» руководителя (для ОЗ и ходатайств в суд) из констант."""
    return [
        "«СОГЛАСЕН»",
        "",
    ] + _role_org_lines(
        c.get("руководитель_должность", "Руководитель"),
        c.get("орган_родительный", c.get("орган_полное", "{ОРГАН_ПОЛН}")),
    ) + [
        "",
        "%s\t\t\t\t%s" % (
            c.get("руководитель_звание", "{РУКОВОДИТЕЛЬ_ЗВАНИЕ}"),
            _fio_last_first(c.get("руководитель_фио", "{РУКОВОДИТЕЛЬ_ФИО}")),
        ),
        "",
        "«___» _____________ %s г." % c.get("год", "20__"),
    ]


def _set_run_font(run):
    """Устанавливает Times New Roman 13 для конкретного run."""
    run.font.name = FONT_NAME
    run.font.size = FONT_SIZE
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)


def _setup_document_style(doc):
    """Поля, шрифт по умолчанию, межстрочный интервал."""
    section = doc.sections[0]
    # Формат страницы — A4 (по умолчанию python-docx ставит Letter, что ломает
    # поля и правую табуляцию подписи).
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    # Поля берём из персонального слоя (константы.json «поля_страницы» = [Л, В, Н, П] в см;
    # у Казани 2,5/2/2/1, у Гаджиево 3,0/2/2/1,5). Если поля не заданы — дефолт-константы.
    _lm, _tm, _bm, _rm = LEFT_MARGIN, TOP_MARGIN, BOTTOM_MARGIN, RIGHT_MARGIN
    try:
        import json as _json
        _cp = _find_constants_path()
        _pf = None
        if _cp:
            with open(_cp, encoding="utf-8") as _constants_file:
                _pf = _json.load(_constants_file).get("поля_страницы")
        if isinstance(_pf, (list, tuple)) and len(_pf) == 4:
            _lm, _tm, _bm, _rm = (Cm(float(_pf[0])), Cm(float(_pf[1])),
                                  Cm(float(_pf[2])), Cm(float(_pf[3])))
    except Exception:
        pass
    section.left_margin = _lm
    section.top_margin = _tm
    section.bottom_margin = _bm
    section.right_margin = _rm
    section.header_distance = HEADER_DISTANCE
    section.footer_distance = FOOTER_DISTANCE

    # Стиль Normal — по умолчанию
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = FONT_SIZE
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), FONT_NAME)
    rFonts.set(qn("w:hAnsi"), FONT_NAME)
    rFonts.set(qn("w:cs"), FONT_NAME)
    rFonts.set(qn("w:eastAsia"), FONT_NAME)

    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)
    # ВНИМАНИЕ: опцию совместимости `doNotExpandShiftReturn` НЕ ставим. Она делает
    # строку перед Shift+Enter выровненной по левому краю — документ становится
    # «лесенкой» (следователь: «загубил, страшный, выровняй по ширине»).
    # Нужно наоборот: строка перед мягким переносом ТЯНЕТСЯ по ширине (поведение
    # Word по умолчанию) — блок остаётся прямоугольным. См. §2.14 формата.


def _set_char_spacing(run, pt):
    """Разрядка (межзнаковый интервал) текста run на pt пунктов.

    Основание: п. 3.2.5 (вид документа) и п. 3.6.4.4 (распорядительное слово)
    Приказа СК России № 40 — печатаются «вразрядку, разреженный на 2,5 пт».
    В OOXML значение w:spacing задаётся в двадцатых долях пункта (2,5 пт = 50)."""
    rPr = run._element.get_or_add_rPr()
    sp = rPr.find(qn("w:spacing"))
    if sp is None:
        sp = OxmlElement("w:spacing")
        rPr.append(sp)
    sp.set(qn("w:val"), str(int(round(pt * 20))))


def _add_para(doc, text, *, alignment=None, indent_first=FIRST_LINE_INDENT, bold=False, spacing_after=0, exact_pt=None, highlight=False, char_spacing_pt=None, left_indent=None):
    """Добавляет обычный абзац с настройками формата.

    highlight=True — заливает текст жёлтым (маркер «не забыть подставить данные»).
    char_spacing_pt — разрядка текста (межзнаковый интервал) в пунктах; для
        заголовков-видов документа и распорядительных слов — 2.5 (п. 3.2.5, 3.6.4.4).
    left_indent — левый отступ абзаца (Cm); для грифов/решений, смещённых вправо.
    """
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    pf = p.paragraph_format
    pf.first_line_indent = indent_first
    if left_indent is not None:
        pf.left_indent = left_indent
    if exact_pt is not None:
        pf.line_spacing = Pt(exact_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(spacing_after)
    run = p.add_run(_normalize_text(_apply_nbsp(text)))
    run.bold = bold
    if highlight:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    _set_run_font(run)
    if char_spacing_pt:
        _set_char_spacing(run, char_spacing_pt)
    return p


def _add_centered(doc, text, *, bold=False, spacing_after=0, char_spacing_pt=None, left_indent=None):
    """Центрированный заголовок (опционально со смещением вправо через left_indent)."""
    return _add_para(
        doc, text,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        indent_first=Cm(0),
        bold=bold,
        spacing_after=spacing_after,
        char_spacing_pt=char_spacing_pt,
        left_indent=left_indent,
    )


def _add_runs(doc, segments, *, alignment=None, indent_first=FIRST_LINE_INDENT, left_indent=None, exact_pt=None):
    """Абзац из нескольких run-ов с индивидуальным форматированием.
    segments — список словарей: {"text": str, "bold": bool, "highlight": bool, "size": pt, "char_spacing_pt": pt}."""
    p = doc.add_paragraph()
    if alignment is not None:
        p.alignment = alignment
    pf = p.paragraph_format
    pf.first_line_indent = indent_first
    if left_indent is not None:
        pf.left_indent = left_indent
    if exact_pt is not None:
        pf.line_spacing = Pt(exact_pt)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    for seg in segments:
        run = p.add_run(_normalize_text(_apply_nbsp(seg.get("text", ""))))
        run.bold = bool(seg.get("bold", False))
        if seg.get("highlight"):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        _set_run_font(run)
        if seg.get("size"):
            run.font.size = Pt(seg["size"])
        if seg.get("char_spacing_pt"):
            _set_char_spacing(run, seg["char_spacing_pt"])
    return p


def _add_blank(doc):
    """Пустая строка."""
    return _add_para(doc, "", indent_first=Cm(0))


def _add_place_date(doc, place, date, time=None):
    """
    Место и дата — в ОДНОЙ строке: место слева, дата прижата к правому краю
    через правую табуляцию (а не пачкой пробелов/табов), поэтому дата НИКОГДА
    не переносится на новую строку.
        г. {ГОРОД}                                                   13.04.2026
                                                            11 ч. 50 мин.   <-- если time задано
    """
    sec = doc.sections[0]
    usable = sec.page_width - sec.left_margin - sec.right_margin  # EMU
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.first_line_indent = Cm(0)
    pf.tab_stops.add_tab_stop(usable, WD_TAB_ALIGNMENT.RIGHT)
    run_left = p.add_run(_normalize_text(_apply_nbsp(place)))
    _set_run_font(run_left)
    run_tab = p.add_run("\t")
    _set_run_font(run_tab)
    run_date = p.add_run(_normalize_text(date or ""))
    _set_run_font(run_date)

    if time:
        p2 = doc.add_paragraph()
        pf2 = p2.paragraph_format
        pf2.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf2.first_line_indent = Cm(0)
        pf2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p2.add_run(_normalize_text(time))
        _set_run_font(r)


SIG_LEAD_BLANKS = 2  # пустых строк перед подписным блоком (можно уменьшать для подгонки на 1 страницу)


def _glue_block(paras):
    """⛔ Скрепляет абзацы блока флагом `keepNext`, чтобы Word не разрывал его
    между страницами.

    Провал 24.08.2026: подпись ВУД разорвалась, на третьей странице осталась одна
    строка звания. Флаг `style_lint` «подпись-разорвана» ловил следствие, но чинить
    приходилось вручную в Word на каждом многостраничном документе. Правильное место
    починки — генератор.

    `keepNext` ставится на все абзацы блока, КРОМЕ последнего: последний тянуть не за
    что, а лишний флаг на нём подтащил бы к подписи посторонний следующий абзац.
    """
    for p in paras[:-1]:
        p.paragraph_format.keep_with_next = True
    return paras


def _add_signature(doc, kind, constants=None):
    """
    Подписной блок.
    kind: 'следователь' | 'руководитель' | 'заместитель'.
    Реквизиты берутся из констант персонального слоя.
    ⛔ Блок скрепляется `keepNext` (см. `_glue_block`) — не разрывается по страницам.
    """
    blocks = _signature_blocks(load_constants(constants))
    if kind not in blocks:
        raise ValueError(
            f"Неизвестный тип подписи: {kind}. "
            f"Допустимые: {list(blocks.keys())}"
        )
    sec = doc.sections[0]
    usable = sec.page_width - sec.left_margin - sec.right_margin  # EMU до правого поля
    made = []
    # Пустые строки между основным текстом и подписным блоком — точно 12 пт
    for _ in range(SIG_LEAD_BLANKS):
        made.append(_add_para(doc, "", indent_first=Cm(0), exact_pt=12))
    # Подписной блок (должность + фамилия) — точно 12 пт
    for line in blocks[kind]:
        if "\t" in line:
            # Строка «звание … Фамилия»: фамилия прижата к правому полю
            # правой табуляцией (а не пачкой табов, которая не доходит до края).
            left, _, rest = line.partition("\t")
            right = rest.replace("\t", "").strip()
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.line_spacing = Pt(12)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.tab_stops.add_tab_stop(usable, WD_TAB_ALIGNMENT.RIGHT)
            run = p.add_run(_normalize_text(left + "\t" + right))
            _set_run_font(run)
            made.append(p)
        else:
            made.append(_add_para(doc, line, indent_first=Cm(0), exact_pt=12))
    _glue_block(made)


def _add_sigblock(doc, lines, lead_blanks=0):
    """Произвольный подписной блок: список строк; строка с табом «звание\\tФамилия»
    выравнивает фамилию по правому полю. Для коротких подписей.
    ⛔ Блок скрепляется `keepNext` (см. `_glue_block`) — не разрывается по страницам."""
    sec = doc.sections[0]
    usable = sec.page_width - sec.left_margin - sec.right_margin
    made = []
    for _ in range(lead_blanks):
        made.append(_add_para(doc, "", indent_first=Cm(0), exact_pt=12))
    for line in lines:
        if "\t" in line:
            left, _, rest = line.partition("\t")
            right = rest.replace("\t", "").strip()
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.line_spacing = Pt(12)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.tab_stops.add_tab_stop(usable, WD_TAB_ALIGNMENT.RIGHT)
            run = p.add_run(_normalize_text(left + "\t" + right))
            _set_run_font(run)
            made.append(p)
        else:
            made.append(_add_para(doc, line, indent_first=Cm(0), exact_pt=12))
    _glue_block(made)


def _add_prosecutor_approval(doc, constants=None):
    """Гриф «УТВЕРЖДАЮ» прокурора (на обвинительном заключении).
    ⛔ Блок скрепляется `keepNext` — гриф не должен разрываться по страницам."""
    _glue_block([
        _add_para(doc, line, indent_first=Cm(0), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        for line in _prosecutor_approval(load_constants(constants))
    ])


def _add_leader_approval(doc, constants=None):
    """Гриф «СОГЛАСЕН» руководителя (на ОЗ и ходатайствах в суд).
    ⛔ Блок скрепляется `keepNext` — гриф не должен разрываться по страницам."""
    _glue_block([
        _add_para(doc, line, indent_first=Cm(0))
        for line in _leader_approval(load_constants(constants))
    ])


def _add_addressee_right(doc, lines):
    """Адресат — обычными абзацами, БЛОК СПРАВА, текст по ЛЕВОМУ краю (СТАНДАРТ).

    ⚠️ ТАБЛИЧНЫЙ адресат НЕ ИСПОЛЬЗОВАТЬ НИКОГДА — он ломает вёрстку документа
    (правка пользователя). Оба типа блока адресата ("addressee_right"
    и "addressee_block") рендерятся этой функцией — обычным текстом, без таблицы.

    Эталон адресата: текст выровнен по ЛЕВОМУ краю, а сам блок сдвинут вправо
    ЛЕВЫМ ОТСТУПОМ абзаца 10,25 см (НЕ выравниванием вправо!), межстрочный
    «точно 12 пт». Пустые строки внутри адресата передавать как "" в lines —
    они сохранят тот же отступ. Функция `_add_addressee_block` (таблица 1×2)
    оставлена только для letterhead с гербовым бланком; напрямую как адресат не
    применяется.
    """
    for line in lines:
        _add_para(doc, line, indent_first=Cm(0), left_indent=Cm(10.25), exact_pt=12)


def _add_addressee_block(doc, lines, *, offset_lines=6, left_w=8.0):
    """Адресат штатной компоновкой ВСО: без-рамочная таблица 1×2.

    Левая ячейка — пустая (место под угловой штамп/гербовый бланк). Правая —
    адресат, текст по ЛЕВОМУ краю, опущенный на уровень «СК РОССИИ» бланка
    (offset_lines пустых строк сверху).

    Args:
        lines: строки адресата, например
            ["Военному прокурору", "{ПРОКУРАТУРА_НАДЗ}", "",
             "майору юстиции", "", "И.О. Фамилия"].
        offset_lines: число пустых строк сверху адресата (по умолчанию 6).
        left_w: ширина левой (пустой) колонки в см.
    """
    sec = doc.sections[0]
    usable_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0  # EMU→см
    right_w = max(usable_cm - left_w, 6.0)

    table = doc.add_table(rows=1, cols=2)
    _remove_table_borders(table)
    table.autofit = False
    table.columns[0].width = Cm(left_w)
    table.columns[1].width = Cm(right_w)
    left_cell, right_cell = table.rows[0].cells
    left_cell.width = Cm(left_w)
    right_cell.width = Cm(right_w)

    out_lines = [""] * offset_lines + list(lines)
    first = True
    for line in out_lines:
        p = right_cell.paragraphs[0] if first else right_cell.add_paragraph()
        first = False
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.line_spacing = Pt(12)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        run = p.add_run(_normalize_text(_apply_nbsp(line)))
        _set_run_font(run)


def _remove_table_borders(table):
    """Убирает все границы таблицы-раскладки — и на уровне таблицы, и на уровне
    каждой ячейки (w:val=\"nil\"), чтобы рамка не печаталась."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn("w:tcBorders")):
                tcPr.remove(old)
            tcb = OxmlElement("w:tcBorders")
            for edge in ("top", "left", "bottom", "right"):
                el = OxmlElement(f"w:{edge}")
                el.set(qn("w:val"), "nil")
                tcb.append(el)
            tcPr.append(tcb)


def _add_anketa(doc, items):
    """Анкета допрашиваемого ровной колонкой — безрамочная таблица 2×N.

    Левая графа — «N. Поле», правая — значение. Значения выровнены по единому
    левому краю; длинные поля и значения переносятся ВНУТРИ своей графы и не
    ломают колонку (в отличие от старого «поле\\tзначение», где табы прыгали
    по сетке и колонка получалась кривой)."""
    # ЕДИНЫЙ СТАНДАРТ АНКЕТЫ (утв. пользователем) —
    # для допросов, протоколов, объяснений: скрытая (безрамочная) таблица 2×N,
    # текст ПО ЛЕВОМУ КРАЮ (не по ширине — иначе «растягивает»), одинарный
    # межстрочный, «интервал после абзаца» 4 пт для читаемости (НЕ пустыми
    # строками!), у последней строки 0; колонки 6,5 / 10 см; хвостовые «;»
    # у значений убираются; предлоги/сокращения склеиваются неразрывным пробелом
    # (без висячих). Анкета должна помещаться на страницу — при переполнении
    # уменьшать «интервал после».
    import re as _re
    _NB = " "
    def _nb(t):
        # НЕ использовать неразрывные пробелы (правило пользователя):
        # анкета и весь текст — только обычные пробелы. Висячие предлоги при
        # необходимости убираются вручную мягким переносом (Shift+Enter), не nbsp.
        return t or ""
    n = len(items)
    table = doc.add_table(rows=n, cols=2)
    _remove_table_borders(table)
    table.autofit = False
    left_w, right_w = Cm(6.5), Cm(10.0)
    table.columns[0].width = left_w
    table.columns[1].width = right_w
    for i, pair in enumerate(items):
        label, value = pair[0], (pair[1] or "")
        value = value.rstrip()
        if value.endswith(";"):
            value = value[:-1].rstrip()
        is_last = (i == n - 1)
        cells = table.rows[i].cells
        cells[0].width = left_w
        cells[1].width = right_w
        for cell, text in ((cells[0], label), (cells[1], value)):
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.left_indent = Cm(0)
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pf.space_before = Pt(0)
            pf.space_after = Pt(0 if is_last else 4)
            run = p.add_run(_normalize_text(_nb(text or "")))
            _set_run_font(run)


def _sanitize_letterhead(image_path):
    """Убирает «вшитую» чёрную рамку по краям гербового бланка и лишние поля.

    Сканирует от каждого края внутрь (до 12 %) и ищет первую «длинную» тёмную
    линию (рамку) — закрашивает её и прилегающие тёмные линии белым, затем
    обрезает по содержимому с равномерным белым полем. Эмблема и текст не
    трогаются: они не образуют сплошной линии во всю ширину/высоту.
    Если PIL недоступен — возвращает исходный путь без изменений.
    """
    try:
        from PIL import Image, ImageChops, ImageDraw
    except ImportError:
        return image_path
    import tempfile, os
    g = Image.open(image_path).convert("L")
    W, H = g.size
    px = g.load()
    rd = lambda y: sum(1 for x in range(W) if px[x, y] < 110) / W
    cd = lambda x: sum(1 for y in range(H) if px[x, y] < 110) / H
    rgb = Image.open(image_path).convert("RGB")
    d = ImageDraw.Draw(rgb)
    # верх
    for y0 in range(int(H * 0.12)):
        if rd(y0) > 0.5:
            y = y0
            while y < H and rd(y) > 0.3:
                d.line([(0, y), (W, y)], fill=(255, 255, 255)); y += 1
            break
    # низ
    for y0 in range(H - 1, int(H * 0.88), -1):
        if rd(y0) > 0.5:
            y = y0
            while y >= 0 and rd(y) > 0.3:
                d.line([(0, y), (W, y)], fill=(255, 255, 255)); y -= 1
            break
    # лево
    for x0 in range(int(W * 0.12)):
        if cd(x0) > 0.5:
            x = x0
            while x < W and cd(x) > 0.3:
                d.line([(x, 0), (x, H)], fill=(255, 255, 255)); x += 1
            break
    # право
    for x0 in range(W - 1, int(W * 0.88), -1):
        if cd(x0) > 0.5:
            x = x0
            while x >= 0 and cd(x) > 0.3:
                d.line([(x, 0), (x, H)], fill=(255, 255, 255)); x -= 1
            break
    bg = Image.new("RGB", rgb.size, (255, 255, 255))
    bbox = ImageChops.difference(rgb, bg).getbbox()
    if bbox:
        pad = 12
        l, tp, r, b = bbox
        rgb = rgb.crop((max(0, l - pad), max(0, tp - pad),
                        min(W, r + pad), min(H, b + pad)))
    fd, out = tempfile.mkstemp(suffix=".png"); os.close(fd)
    rgb.save(out)
    return out


def _add_letterhead(doc, image_path, *, width_cm=8.0, addressee_lines=None,
                    addressee_offset_lines=5):
    """Гербовый бланк в левом верхнем углу (для запросов, сопроводов и пр.).

    Вставляет изображение углового штампа слева. Если задан addressee_lines —
    справа от бланка размещается адресат (штатная компоновка исходящего письма).

    Args:
        image_path: путь к файлу гербового бланка (PNG).
        width_cm: ширина бланка в сантиметрах (по умолчанию 8,0).
        addressee_lines: список строк адресата справа (опционально).
        addressee_offset_lines: сколько пустых строк добавить сверху адресата,
            чтобы он начинался на уровне строки «СК РОССИИ» (под орлом), а не
            у самого верха эмблемы. По умолчанию 5.
    """
    image_path = _sanitize_letterhead(image_path)
    if addressee_lines:
        if addressee_offset_lines:
            addressee_lines = [""] * addressee_offset_lines + list(addressee_lines)
        # Двухколоночная безрамочная таблица: слева бланк, справа адресат.
        table = doc.add_table(rows=1, cols=2)
        _remove_table_borders(table)
        table.autofit = False
        left_w = min(width_cm + 0.4, 10.0)
        right_w = 16.0 - left_w
        table.columns[0].width = Cm(left_w)
        table.columns[1].width = Cm(right_w)

        left_cell, right_cell = table.rows[0].cells
        left_cell.width = Cm(left_w)
        right_cell.width = Cm(right_w)

        # Левая ячейка — бланк
        p_img = left_cell.paragraphs[0]
        p_img.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_img.paragraph_format.first_line_indent = Cm(0)
        run_img = p_img.add_run()
        run_img.add_picture(image_path, width=Cm(width_cm))

        # Правая ячейка — адресат
        first = True
        for line in addressee_lines:
            p = right_cell.paragraphs[0] if first else right_cell.add_paragraph()
            first = False
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.first_line_indent = Cm(0)
            pf.line_spacing = Pt(12)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            run = p.add_run(_normalize_text(_apply_nbsp(line)))
            _set_run_font(run)
    else:
        # Только бланк, слева вверху
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        run = p.add_run()
        run.add_picture(image_path, width=Cm(width_cm))


def _add_sig_line(doc, role, name):
    """Подпись участника: статус слева, короткая линия по центру, ФИО справа.

    Линия (место для подписи) — короткая, центрирована на середине рабочего поля;
    ФИО прижато к правому краю через правую табуляцию."""
    sec = doc.sections[0]
    usable_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0  # EMU→см
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.tab_stops.add_tab_stop(Cm(usable_cm / 2.0), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(Cm(usable_cm), WD_TAB_ALIGNMENT.RIGHT)
    run = p.add_run(_normalize_text(role + "\t" + "____________" + "\t" + name))
    _set_run_font(run)
    return p


# ============================================================================
# Диспетчер блоков (общий для build_doc и build_letterhead_doc)
# ============================================================================

def _add_approval_indented(doc, lines, indent_cm=8.25):
    """Гриф «ПРОДЛЕВАЮ»/«УТВЕРЖДАЮ» (вариант А, ст. 144) — блок с левым отступом
    (правая половина листа), выровнен по левому краю. В строке вида
    'звание\\tФамилия' правый таб ставит фамилию к правому полю (рабочая ширина —
    из фактических полей секции)."""
    sec = doc.sections[0]
    usable_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0  # EMU→см
    for line in lines:
        p = _add_para(doc, line, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                      indent_first=Cm(0), left_indent=Cm(indent_cm))
        if p is not None and "\t" in line:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(usable_cm), WD_TAB_ALIGNMENT.RIGHT)


def _process_blocks(doc, body_blocks, constants=None):
    """Проходит по body_blocks и добавляет их в документ.

    Управление вертикальными отступами (пустыми строками):
    п. 3.6.4.4 Приказа СК России № 40 — распорядительное слово («ПРИКАЗЫВАЮ»,
    аналог нашего УСТАНОВИЛ:/ПОСТАНОВИЛ:) печатается отдельной строкой по центру
    и ОТДЕЛЯЕТСЯ ОТ ЧАСТЕЙ ТЕКСТА межстрочным интервалом; текст отделяется от
    шапки/заголовка (§2.3 формата). Здесь это зашито в код:
      • после блока «место/дата» (place_date) — одна пустая строка перед преамбулой;
      • перед и после section (УСТАНОВИЛ:/ПОСТАНОВИЛ:) — по одной пустой строке.
    Логика идемпотентна: подряд идущие пустые строки схлопываются в одну,
    пустые строки в начале документа подавляются (двойных отбивок не будет).
    """
    state = {"last_blank": True, "pending_blank": False}

    def _emit_blank(exact_pt=12):
        if state["last_blank"]:
            return
        _add_para(doc, "", indent_first=Cm(0), exact_pt=exact_pt)
        state["last_blank"] = True

    for block in body_blocks:
        t = block["type"]

        # Явная пустая строка из шаблона — со схлопыванием дублей.
        if t == "blank":
            _emit_blank(exact_pt=None)
            state["pending_blank"] = False
            continue

        # Отдать отложенную пустую строку «снизу» (после place_date / section).
        if state["pending_blank"]:
            _emit_blank()
            state["pending_blank"] = False

        if t == "header_centered":
            # Вид документа (ПОСТАНОВЛЕНИЕ, ПРОТОКОЛ) — прописными, полужирный,
            # вразрядку 2,5 пт (п. 3.2.5 Приказа № 40).
            _add_centered(doc, block["text"], bold=True, char_spacing_pt=2.5)
        elif t == "header_centered_thin":
            _add_centered(doc, block["text"], bold=False)
        elif t == "page_break":
            # Разрыв страницы — для нескольких документов (комплекта запросов) в одном файле.
            doc.add_page_break()
            state["pending_blank"] = False
            continue
        elif t == "grif_line":
            # Строка грифа-решения руководителя (постановление о продлении проверки КРСП):
            # по центру, но со смещением блока ВПРАВО (гриф-решение размещается в
            # правой части листа, а не по центру страницы).
            # Левый отступ по умолчанию 6,75 см; строка с датой решения — 1,5 см.
            _add_centered(doc, block["text"], bold=False,
                          left_indent=Cm(block.get("indent_cm", 6.75)))
        elif t == "approval_indented":
            # Гриф «ПРОДЛЕВАЮ»/«УТВЕРЖДАЮ» с левым отступом (вариант А, ст. 144).
            _add_approval_indented(doc, block["lines"],
                                   indent_cm=block.get("indent_cm", 8.25))
        elif t == "place_date":
            _add_place_date(doc, block.get("place", ""), block.get("date", ""), block.get("time"))
        elif t == "para":
            _add_para(doc, block["text"], alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, highlight=block.get("highlight", False), bold=block.get("bold", False))
        elif t == "para_no_indent":
            _add_para(doc, block["text"], indent_first=Cm(0), alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, highlight=block.get("highlight", False), bold=block.get("bold", False))
        elif t == "runs":
            # Абзац со смешанным форматированием: segments=[{"text","bold","highlight"}, ...].
            _ALIGN_MAP = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
                          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
            _add_runs(doc, block["segments"],
                      alignment=_ALIGN_MAP.get(block.get("align")),
                      indent_first=Cm(block.get("indent_cm", 0)),
                      left_indent=Cm(block["left_cm"]) if "left_cm" in block else None)
        elif t == "li":
            # Нумерованный пункт перечня — с абзацным отступом первой строки, как у основного текста.
            _add_para(doc, block["text"], indent_first=FIRST_LINE_INDENT, alignment=WD_ALIGN_PARAGRAPH.JUSTIFY, highlight=block.get("highlight", False), bold=block.get("bold", False))
        elif t == "prilozhenie":
            # Реквизит «Приложение» — с прописной буквы, без отступа, по левому краю,
            # межстрочный «точно 12 пт» (п. 3.2.14 Приказа СК России № 40).
            _add_para(doc, block["text"], indent_first=Cm(0), alignment=WD_ALIGN_PARAGRAPH.LEFT, exact_pt=12, highlight=block.get("highlight", False))
        elif t == "obrashenie":
            # Обращение к адресату («Уважаемый Имя Отчество!») — ВСЕГДА по центру,
            # обычным начертанием, без отступа (§2.14).
            _add_centered(doc, block["text"], bold=False)
        elif t == "h_left":
            # Левый жирный подзаголовок (для справок, аналитических записок), без отступа.
            _add_para(doc, block["text"], indent_first=Cm(0), alignment=WD_ALIGN_PARAGRAPH.LEFT, bold=True)
        elif t == "li_left":
            # Нумерованный пункт без выравнивания по ширине (для анкеты «№ — значение»), с абзацным отступом.
            _add_para(doc, block["text"], indent_first=FIRST_LINE_INDENT, alignment=WD_ALIGN_PARAGRAPH.LEFT, highlight=block.get("highlight", False))
        elif t == "anketa":
            # Анкета лица ровной колонкой (безрамочная таблица 2×N): [["1. Поле","значение"], ...].
            _add_anketa(doc, block["items"])
        elif t == "section":
            # Распорядительное слово (УСТАНОВИЛ:/ПОСТАНОВИЛ:) — Казань/02a:
            # обычным текстом (НЕ жирным), по центру, БЕЗ разрядки, интервал 6 пт до и после.
            _sp = doc.add_paragraph()
            _sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            _spf = _sp.paragraph_format
            _spf.first_line_indent = Cm(0)
            _spf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            _spf.space_before = Pt(6)
            _spf.space_after = Pt(6)
            _sr = _sp.add_run(_normalize_text(_apply_nbsp(block["text"])))
            _sr.bold = False
            _set_run_font(_sr)
        elif t == "signature":
            _add_signature(doc, block["kind"], constants=constants)
        elif t == "sigblock":
            _add_sigblock(doc, block["lines"], lead_blanks=block.get("lead_blanks", 0))
        elif t == "prosecutor_approval":
            _add_prosecutor_approval(doc, constants=constants)
        elif t == "leader_approval":
            _add_leader_approval(doc, constants=constants)
        elif t == "addressee_right":
            _add_addressee_right(doc, block["lines"])
        elif t == "addressee_block":
            # НИКОГДА не таблицей: табличный адресат ломает вёрстку документа
            # (правка пользователя «запомни в скилл про таблицу и не используй её
            # никогда»). "addressee_block" и "addressee_right" рендерятся ОДИНАКОВО —
            # обычными абзацами. offset_lines/left_w не применяются.
            _add_addressee_right(doc, block["lines"])
        elif t == "letterhead":
            _add_letterhead(
                doc,
                block["image_path"],
                width_cm=block.get("width_cm", 8.0),
                addressee_lines=block.get("addressee_lines"),
                addressee_offset_lines=block.get("addressee_offset_lines", 5),
            )
        elif t == "sig_line":
            _add_sig_line(doc, block["role"], block["name"])
        elif t == "raw":
            _add_para(doc, block["text"], indent_first=Cm(0))
        else:
            raise ValueError(f"Неизвестный тип блока: {t}")

        state["last_blank"] = False
        # Пустую строку «снизу» откладываем до следующего видимого блока.
        if t in ("section", "place_date"):
            state["pending_blank"] = True


# ============================================================================
# Ретенция бэкапов и безопасная перезапись
# ============================================================================

# ⛔ 10 → 3 (05.09.2026). Десять .bak лежат в ПАПКЕ РЕШЕНИЯ вперемешку с процессуальными
# документами и читаются как содержимое дела — это прямо против правила 47 (папка-событие
# содержит документы события, а не служебный мусор). Три копии закрывают реальную задачу
# бэкапа: откат последней правки и предыдущей. Глубже отката за сутки не случалось ни разу,
# а возвратный путь всё равно фиксируется журналом, а не грудой .bak.
BAK_KEEP = 3             # сколько .bak хранить рядом с одним файлом
BAK_MAX_AGE_DAYS = 30    # и не старше скольких дней (иначе .bak копятся бесконечно)


def prune_backups(out_path, keep=BAK_KEEP, max_age_days=BAK_MAX_AGE_DAYS):
    """Ретенция .bak: оставить последние `keep` и не старше `max_age_days`, остальное удалить.
    Без неё .bak_* множатся при каждой правке/пересборке и засоряют папку дела."""
    import glob as _g, time as _t
    baks = sorted((p for p in _g.glob(out_path + ".bak_*") if os.path.isfile(p)),
                  key=lambda p: os.path.getmtime(p))
    drop = set(baks[:-keep]) if keep and len(baks) > keep else set()
    now = _t.time()
    for b in baks:
        try:
            if (now - os.path.getmtime(b)) > max_age_days * 86400:
                drop.add(b)
        except OSError:
            pass
    for b in drop:
        try:
            os.remove(b)
        except OSError:
            pass


def _finalize_meta(path):
    """Свойства файла → профиль владельца (автор = следователь, приложение = его Word).

    Убирает следы инструмента: dc:creator="python-docx", description="generated by
    python-docx", даты дефолтного шаблона 2013-12-23, «Microsoft Macintosh Word 14» —
    всё это стояло в СВОЙСТВАХ боевых документов (аудит 07.08.2026). Документ выпускает
    и подписывает следователь — он и автор; python-docx лишь средство набора, как Word.
    Идемпотентно; сбой метаданных не роняет сохранение содержимого. См. doc_meta.py."""
    try:
        import doc_meta
        doc_meta.clean(path)
    except Exception:
        pass


def _backup_if_exists(out_path):
    """Перед перезаписью существующего файла — сделать .bak с временной меткой.
    Страховка: регенерация из шаблона не должна МОЛЧА затирать возможные ручные правки.
    (Для точечной правки существующего документа используй scripts/docx_edit.py, а не пересборку.)
    ВАЖНО: если бэкап НЕ удался — перезапись отменяется (иначе оригинал теряется молча)."""
    import time as _t, shutil as _sh
    if out_path and os.path.exists(out_path):
        _stamp = _t.strftime("%Y%m%d_%H%M%S")
        _bak = out_path + ".bak_" + _stamp
        _i = 1
        while os.path.exists(_bak):
            _bak = "%s.bak_%s_%d" % (out_path, _stamp, _i); _i += 1
        try:
            _sh.copyfile(out_path, _bak)
        except OSError as e:
            raise IOError("Не удалось сделать .bak перед перезаписью «%s»: %s. "
                          "Перезапись отменена, чтобы не потерять оригинал." % (out_path, e)) from e
        prune_backups(out_path)


# ============================================================================
# Публичные функции сборки
# ============================================================================

def build_doc(out_path, body_blocks, constants=None):
    """Собирает обычный процессуальный документ (на чистом листе) и сохраняет.

    Использовать для постановлений, обвинительного заключения и ходатайств.
    Протоколы следственных действий и формы со сложными таблицами готовятся
    правкой отдельной копии бланка по references/25-assembly.md, а не этой функцией.

    Каждый блок — словарь. Основные типы:

    Заголовки и подзаголовки:
    - {"type": "header_centered", "text": "ПОСТАНОВЛЕНИЕ"}
        — заголовок по центру, жирный, вразрядку.
    - {"type": "header_centered_thin", "text": "о возбуждении уголовного дела"}
        — подзаголовок по центру, обычный.
    - {"type": "section", "text": "УСТАНОВИЛ:"}
        — распорядительное слово по центру, обычным начертанием, БЕЗ разрядки,
          интервал 6 пт до и после (УСТАНОВИЛ:, ПОСТАНОВИЛ: и т.п.).

    Место и дата:
    - {"type": "place_date", "place": "г. {ГОРОД}", "date": "13.04.2026"} — без времени.
    - {"type": "place_date", "place": "г. {ГОРОД}", "date": "13.04.2026", "time": "11 ч. 50 мин."}
        — с временем (на отдельной строке справа).

    Текст:
    - {"type": "para", "text": "..."} — обычный абзац с отступом 1,25 см, по ширине.
    - {"type": "para_no_indent", "text": "..."} — абзац без отступа (пункты резолютивной части).
    - {"type": "li", "text": "..."} — нумерованный пункт с абзацным отступом.
    - {"type": "prilozhenie", "text": "Приложение: ... на ___ л. в ___ экз."}
        — реквизит «Приложение»: без отступа, по левому краю, «точно 12 пт».
    - {"type": "obrashenie", "text": "Уважаемый Имя Отчество!"} — обращение по центру.
    - {"type": "runs", "segments": [{"text","bold","highlight",...}], "align": "..."}
        — абзац со смешанным форматированием.
    - {"type": "anketa", "items": [["1. Поле", "значение"], ...]} — анкета лица (таблица 2×N).
    - {"type": "raw", "text": "..."} — текст без обработки.
    - {"type": "blank"} / {"type": "page_break"} — пустая строка / разрыв страницы.

    Адресат, подпись, грифы:
    - {"type": "letterhead", "image_path": "scripts/blank.png", "addressee_lines": [...], "width_cm": 7.0}.
    - {"type": "addressee_right", "lines": [...]} — адресат справа (для запросов и сопроводов).
    - {"type": "signature", "kind": "следователь|руководитель|заместитель"}.
    - {"type": "sigblock", "lines": [...]} — произвольный подписной блок.
    - {"type": "prosecutor_approval"} — гриф «УТВЕРЖДАЮ» прокурора (для ОЗ).
    - {"type": "leader_approval"} — гриф «СОГЛАСЕН» руководителя (для ОЗ и ходатайств в суд).

    Args:
        out_path: путь для сохранения .docx.
        body_blocks: список блоков (см. выше).
        constants: путь к «константы.json» (по умолчанию — автопоиск/SK_CONSTANTS/фолбэк).

    Returns:
        out_path (для удобства chain).
    """
    _backup_if_exists(out_path)
    doc = Document()
    _setup_document_style(doc)
    _process_blocks(doc, body_blocks, constants=constants)
    doc.save(out_path)
    _finalize_meta(out_path)
    return out_path


def _para_has_content(el):
    """True, если в абзаце есть текст ИЛИ графика (герб/штамп/textbox бланка)."""
    if any((t.text or "").strip() for t in el.iter(qn("w:t"))):
        return True
    if el.findall(".//" + qn("w:drawing")) or el.findall(".//" + qn("w:pict")):
        return True
    return False


def _split_addressee(body_blocks):
    """
    Выделяет блок адресата (он должен идти первым) из остального тела.

    Поддерживает:
    - явный блок {"type": "addressee"/"addressee_right"/"addressee_block", "lines": [...]};
    - устаревшие шаблоны, где адресат набран ведущими para_no_indent
      (с возможными пустыми строками внутри) до начала основного текста.

    Возвращает (addressee_lines, rest_blocks).
    """
    rest = list(body_blocks)
    if rest and rest[0].get("type") in ("addressee", "addressee_right", "addressee_block"):
        return list(rest[0].get("lines", [])), rest[1:]
    if rest and rest[0].get("type") == "para_no_indent":
        lines = []
        while rest and rest[0].get("type") in ("para_no_indent", "blank"):
            b = rest.pop(0)
            lines.append("" if b["type"] == "blank" else b.get("text", ""))
        while lines and lines[-1] == "":
            lines.pop()
        return lines, rest
    return [], rest


def _set_executor_footer(doc, name, phone, size=Pt(10)):
    """Исполнитель в нижнем колонтитуле (для запросов и писем): ФИО + телефон."""
    footer = doc.sections[0].footer
    footer.is_linked_to_previous = False
    for i, txt in enumerate((name, phone)):
        p = footer.paragraphs[0] if i == 0 else footer.add_paragraph()
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = LINE_EXACT
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        r = p.add_run(_normalize_text(txt))
        _set_run_font(r)
        r.font.size = size
    return footer


def build_letterhead_doc(out_path, body_blocks, blank_template_path=None,
                         leading_lines=LETTERHEAD_LEADING_LINES, executor=None,
                         constants=None, *, letterhead_profile="auto"):
    """
    Собирает документ на бланке органа (герб + реквизиты) и сохраняет.

    Использовать для запросов, сопроводов, повесток, справок, уведомлений,
    поручений (ст. 152), представлений (ст. 158 ч. 2 УПК РФ).

    ВЫРАВНИВАНИЕ БЛАНКА (важно):
    - Блок адресата ставится ПЕРВЫМ (тип "addressee"/"addressee_right"/"addressee_block"
      или ведущие para_no_indent). Размещается СПРАВА за счёт левого отступа ~7 см и
      оказывается НА УРОВНЕ реквизитов отправителя (углового штампа слева).
    - Основной текст начинается через два пробела (две пустые строки) от нижнего
      края исходящего («№ ___ от ___»). За это отвечает leading_lines.
    - Ручные пустые строки после адресата НЕ нужны — отступ ставится автоматически.

    Args:
        out_path: путь для сохранения .docx.
        body_blocks: список блоков (см. build_doc); первым — адресат.
        blank_template_path: обязательный явный путь к проверенному бланку соответствующего жанра.
        leading_lines: строк под шапку бланка до текста (по умолчанию 15).
        executor: колонтитул исполнителя (только запросы/поручения). Варианты:
            None (по умолчанию) — без колонтитула (сопроводы прокурору и руководителю);
            True — данные исполнителя взять из констант (исполнитель_имя_отчество / _телефон);
            (имя, телефон) — явный кортеж.
        constants: путь к «константы.json».
        letterhead_profile: auto (по структуре), vsu-cvo-requests или legacy.
            Для запросов ВСУ выбирай vsu-cvo-requests явно; имя копии не имеет значения.

    Returns:
        out_path (для удобства chain).
    """
    if blank_template_path is None:
        raise ValueError('Укажите blank_template_path по жанру: для запросов – blank_vsu_cvo_requests.docx; старый бланк автоматически не выбирается.')

    if not os.path.isfile(blank_template_path):
        raise FileNotFoundError(
            f"Бланк не найден: {blank_template_path}\n"
            f"Положи файл бланка (blank_384_vso.docx) в папку templates/ рядом со скриптом, "
            f"либо передай явный путь параметром blank_template_path."
        )

    doc = Document(blank_template_path)
    _setup_document_style(doc)  # применяет A4 и поля поверх бланка

    addressee_lines, rest = _split_addressee(body_blocks)

    # В бланке ВСУ адресат находится в правой ячейке на уровне угловика.
    # Размещение адресата отдельными абзацами ниже таблицы нарушает бланк.
    if letterhead_profile not in ("auto", "vsu-cvo-requests", "legacy"):
        raise ValueError("Неизвестный профиль бланка: " + str(letterhead_profile))
    left_text = doc.tables[0].cell(0, 0).text.upper() if doc.tables else ""
    structural_vsu = ("УПРАВЛЕНИЕ ПО ЦЕНТРАЛЬНОМУ" in left_text
                      and "ВОЕННОМУ ОКРУГУ" in left_text)
    # The filename is retained only as a compatibility guard for malformed old inputs.
    named_vsu = os.path.basename(blank_template_path).lower() == "blank_vsu_cvo_requests.docx"
    vsu_request_blank = letterhead_profile == "vsu-cvo-requests" or (
        letterhead_profile == "auto" and (structural_vsu or named_vsu))
    if vsu_request_blank:
        if not structural_vsu:
            raise ValueError("Профиль ВСУ выбран для бланка с неподтверждённым угловиком")
        if not doc.tables or not doc.tables[0].rows or len(doc.tables[0].rows[0].cells) < 2:
            raise ValueError("Бланк ВСУ не содержит правой ячейки для адресата")
        if not addressee_lines:
            raise ValueError("Для запроса на бланке ВСУ нужен блок адресата первым в body_blocks")
        cell = doc.tables[0].cell(0, 1)
        if cell._tc is doc.tables[0].cell(0, 0)._tc or len(cell.paragraphs) < 3:
            raise ValueError("Бланк ВСУ не содержит отдельного проверенного поля адресата")
        slots = cell.paragraphs[3:]
        for i, line in enumerate(addressee_lines):
            p = slots[i] if i < len(slots) else cell.add_paragraph()
            p.text = _normalize_text(line)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(12)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for run in p.runs:
                _set_run_font(run)
        for p in slots[len(addressee_lines):]:
            p.text = ""
        for p in list(doc.paragraphs):
            if not _para_has_content(p._p):
                p._p.getparent().remove(p._p)
        _process_blocks(doc, rest, constants=constants)
        if executor:
            if executor is True:
                raise ValueError("Для запроса на бланке ВСУ укажи исполнителя и телефон явным кортежем")
            name, phone = executor
            _set_executor_footer(doc, name, phone)
        _backup_if_exists(out_path)
        doc.save(out_path)
        _finalize_meta(out_path)
        return out_path

    # Убираем пустые абзацы-распорки шаблона (абзац с гербом/штампом сохраняем),
    # чтобы самим управлять отступом до текста.
    for p in list(doc.paragraphs):
        if not _para_has_content(p._p):
            p._p.getparent().remove(p._p)

    # Абзац с плавающим бланком остаётся первой строкой.
    lines_used = 1
    for _ in range(ADDRESSEE_DROP_LINES):              # опустить адресат до уровня «СК РОССИИ»
        _add_blank(doc)
        lines_used += 1
    for line in addressee_lines:                       # адресат – вверху, справа
        _add_para(doc, line, indent_first=Cm(0),
                  alignment=WD_ALIGN_PARAGRAPH.LEFT, left_indent=ADDRESSEE_INDENT,
                  exact_pt=12)
        lines_used += 1
    while lines_used < leading_lines:                  # «два пробела» от исходящего
        _add_blank(doc)
        lines_used += 1

    _process_blocks(doc, rest, constants=constants)
    # Колонтитул с исполнителем – только запросы/поручения. На сопроводах
    # прокурору и руководителю ВСО передавай executor=None (по умолчанию).
    if executor:
        if executor is True:
            c = load_constants(constants)
            # ⛔ 10.09.2026: константы.json владельца несёт «телефон_исполнителя» и «следователь.фио»,
            # а здесь ждали «исполнитель_телефон» / «исполнитель_имя_отчество» — executor=True
            # МОЛЧА отдавал плейсхолдеры в подвал (поймано рендером пробного письма). Читаем оба
            # написания, имя-отчество выводим из ФИО следователя («Фамилия Имя Отчество»).
            name = c.get("исполнитель_имя_отчество")
            if not name:
                fio = (c.get("следователь") or {}).get("фио", "") if isinstance(c.get("следователь"), dict) else ""
                parts = fio.split()
                name = " ".join(parts[1:3]) if len(parts) >= 3 else "{ИСПОЛНИТЕЛЬ_ИМЯ}"
            phone = c.get("исполнитель_телефон") or c.get("телефон_исполнителя") or "{ТЕЛ_ИСПОЛНИТЕЛЯ}"
        else:
            name, phone = executor
        _set_executor_footer(doc, name, phone)
    _backup_if_exists(out_path)
    doc.save(out_path)
    _finalize_meta(out_path)
    return out_path


def add_case_header(path, number):
    """№ уголовного дела в верхнем колонтитуле ПЕРВОЙ страницы (по правому краю).
    Применять к постановлению о возбуждении уголовного дела (ВУД) после сборки."""
    d = Document(path)
    s = d.sections[0]
    s.different_first_page_header_footer = True
    p = s.first_page_header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    for r in list(p.runs):
        r.text = ""
    if p.runs:
        p.runs[0].text = f"№ {number}"
        _set_run_font(p.runs[0])
    else:
        _set_run_font(p.add_run(f"№ {number}"))
    d.save(path)
    return path


# ============================================================================
# Надёжная работа с .docx на смонтированной папке (mount периодически даёт «битый zip»)
# ============================================================================

def safe_doc(path, retries=6, delay=0.4):
    """Надёжно открыть .docx: копируем в /tmp, проверяем zipfile.is_zipfile,
    повторяем при «битом zip» (артефакт облачной синхронизации/блокировки Word)."""
    last = None
    for i in range(retries):
        tmp = os.path.join(tempfile.gettempdir(), f"_safe_{os.getpid()}_{i}.docx")
        try:
            shutil.copyfile(path, tmp)
            if zipfile.is_zipfile(tmp):
                return Document(tmp)
        except Exception as e:
            last = e
        time.sleep(delay)
    raise IOError(
        f"Не удалось прочитать валидный .docx за {retries} попыток: {path} ({last}). "
        f"Если рядом есть файл ~$<имя> – документ открыт в Word (снять блокировку). "
        f"Если файл действительно битый – см. scripts/docx_recover.py.")


def safe_save_copy(doc, dst, retries=6, delay=0.4):
    """Сохранить в /tmp, проверить валидность, затем копировать на mount с повтором
    и повторной проверкой (запись на mount тоже иногда даёт битый zip)."""
    tmp = os.path.join(tempfile.gettempdir(), f"_save_{os.getpid()}.docx")
    doc.save(tmp)
    if not zipfile.is_zipfile(tmp):
        raise IOError(f"python-docx сохранил битый файл: {tmp}")
    last = None
    for i in range(retries):
        try:
            shutil.copyfile(tmp, dst)
            if zipfile.is_zipfile(dst):
                return dst
        except Exception as e:
            last = e
        time.sleep(delay)
    raise IOError(f"Не удалось записать валидный .docx на {dst} за {retries} попыток ({last}).")


# ============================================================================
# Редактирование готового бланка (паттерн «заполнить чужой шаблон», а не строить с нуля)
# ============================================================================

def set_cell_text(cell, text, *, bold=False):
    """Заменить текст ячейки таблицы на РОВНО один run, сохранив шрифт Times New Roman 13."""
    cell.text = ""                       # очистить содержимое ячейки -> один абзац
    p = cell.paragraphs[0]
    for r in list(p.runs):               # убрать пустой run, оставленный сеттером cell.text (иначе runs[0] пуст)
        r._r.getparent().remove(r._r)
    r = p.add_run(text)
    r.bold = bold
    _set_run_font(r)
    return cell


def replace_block_between_anchors(doc, start_text, end_text, new_lines):
    """Заменить абзацы СТРОГО между абзацем с start_text и абзацем с end_text
    (границы не трогаем) на new_lines. Для замены блока показаний в готовом шаблоне."""
    ps = doc.paragraphs
    si = next((i for i, p in enumerate(ps) if start_text in p.text), None)
    ei = next((i for i, p in enumerate(ps) if end_text in p.text and (si is None or i > si)), None)
    if si is None or ei is None or ei <= si:
        raise ValueError("Якоря не найдены или в неверном порядке")
    for p in ps[si + 1:ei]:
        p._element.getparent().remove(p._element)
    anchor = ps[ei]._element
    for line in new_lines:
        np = doc.add_paragraph()
        anchor.addprevious(np._element)
        r = np.add_run(line); _set_run_font(r)
        np.paragraph_format.first_line_indent = FIRST_LINE_INDENT
        np.alignment = BODY_ALIGN
    return doc


def replace_name(doc, old, new):
    """Глобальная замена ФИО (защитника/обвиняемого) во всём документе: абзацы,
    таблицы, колонтитулы. Склеивает текст абзаца, если вхождение разорвано по run."""
    def fix(p):
        if old in p.text:
            full = p.text.replace(old, new)
            for r in list(p.runs):
                r.text = ""
            if p.runs:
                p.runs[0].text = full; _set_run_font(p.runs[0])
            else:
                _set_run_font(p.add_run(full))
    def walk(c):
        for p in c.paragraphs:
            fix(p)
        for t in c.tables:
            for row in t.rows:
                for cell in row.cells:
                    walk(cell)
    walk(doc)
    for sec in doc.sections:
        for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
            for p in hf.paragraphs:
                fix(p)
    return doc


def replace_text_preserving_layout(doc, old, new, *, include_headers=False):
    """Заменить old -> new, СОХРАНЯЯ вёрстку: не склеивает все runs абзаца в первый
    (в отличие от replace_name) и по умолчанию НЕ трогает колонтитулы.

    Зачем: на сложных формах (статкарты Ф-1/Ф-1.2, бланки с таблицами) склейка runs
    ломает форматирование и разбивку по страницам, а обработка колонтитулов раздувает
    форму (наблюдалось 4 -> 6 страниц). Здесь:
      - если вхождение целиком в одном run – правим только этот run;
      - если вхождение разорвано по нескольким runs (часто: «в/ч 00000» = ['в/ч 00','0','00'],
        дата = ['1','5','.0','6']) – переписываем только перекрытые runs, остальные не трогаем;
      - подписи с табами: <w:tab/> – отдельный элемент, run.text его не содержит, поэтому
        замена идёт по текстовым фрагментам и табуляция сохраняется.
    Для статкарт и иных форм вызывать с include_headers=False (по умолчанию).
    """
    if old == new or not old:
        return 0

    def fix_paragraph(p):
        runs = p.runs
        if not runs:
            return 0
        n = 0
        search_from = 0
        guard = 0
        while guard < 1000:
            guard += 1
            texts = [r.text for r in runs]
            full = "".join(texts)
            pos = full.find(old, search_from)
            if pos == -1:
                break
            end = pos + len(old)
            starts = []
            acc = 0
            for t in texts:
                starts.append(acc)
                acc += len(t)

            def run_at(off):
                for i, t in enumerate(texts):
                    if starts[i] <= off < starts[i] + len(t):
                        return i
                return len(texts) - 1

            i0 = run_at(pos)
            i1 = run_at(end - 1)
            if i0 == i1:
                r = runs[i0]
                ls = pos - starts[i0]
                r.text = r.text[:ls] + new + r.text[ls + len(old):]
            else:
                r0, r1 = runs[i0], runs[i1]
                ls = pos - starts[i0]
                le = end - starts[i1]
                r0.text = r0.text[:ls] + new
                for j in range(i0 + 1, i1):
                    runs[j].text = ""
                r1.text = r1.text[le:]
            n += 1
            # Шагнуть ЗА вставленное: если old входит в new, повтор иначе не заменится
            # (прежний guard «if old in new: break» обрывал абзац после 1-й замены),
            # а поиск с той же позиции дал бы бесконечный цикл.
            search_from = pos + len(new)
        return n

    def walk(container):
        c = 0
        for p in container.paragraphs:
            c += fix_paragraph(p)
        for t in container.tables:
            for row in t.rows:
                for cell in row.cells:
                    c += walk(cell)
        return c

    total = walk(doc)
    if include_headers:
        for sec in doc.sections:
            for hf in (sec.header, sec.footer, sec.first_page_header, sec.first_page_footer):
                for pp in hf.paragraphs:
                    total += fix_paragraph(pp)
    return total


# ============================================================================
# Самопроверка
# ============================================================================

if __name__ == "__main__":
    import sys
    try:  # иначе самопроверка падает на «×»/«–» в cp1251-консоли Windows (аудит 18.07.2026)
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    # Разбор аргументов ДО любой записи на диск: «--help»/опечатка не должны
    # создавать папку с демо-документом (ревизия 22.08.2026).
    _arg = sys.argv[1] if len(sys.argv) > 1 else None
    if _arg in ("-h", "--help", "/?", "help"):
        print(__doc__)
        print("Запуск как скрипта — самопроверка (собирает демо-документы):")
        print("    python make_docx.py            # положить демо в системный TEMP")
        print("    python make_docx.py <папка>    # положить демо в указанную папку")
        sys.exit(0)
    if _arg is not None and _arg.startswith("-"):
        sys.exit("Неизвестный аргумент: %s\nСправка: python make_docx.py --help" % _arg)

    out_dir = _arg if _arg else tempfile.gettempdir()
    os.makedirs(out_dir, exist_ok=True)

    test_vud = [
            {"type": "title", "text": "УЧЕБНЫЙ ПРИМЕР"},
            {"type": "para", "text": "[Синтетический текст для проверки сборки документа.]"},
        ]
    out_vud = os.path.join(out_dir, "test_vud.docx")
    build_doc(out_vud, test_vud)
    print(f"[1] Постановление: {out_vud}")

    from docx import Document as Doc
    d = Doc(out_vud)
    s = d.sections[0]
    is_a4 = abs(s.page_width.cm - 21.0) < 0.01 and abs(s.page_height.cm - 29.7) < 0.01
    print(f"  Страница: {s.page_width.cm:.2f} × {s.page_height.cm:.2f} см "
          f"({'A4 OK' if is_a4 else 'НЕ A4 !'})")
    print(f"  Поля: Л={s.left_margin.cm:.2f} / В={s.top_margin.cm:.2f} / "
          f"Н={s.bottom_margin.cm:.2f} / П={s.right_margin.cm:.2f}")
    print(f"  Шрифт: {d.styles['Normal'].font.name} {d.styles['Normal'].font.size.pt:.0f} pt")
    print(f"  Параграфов: {len(d.paragraphs)}")
