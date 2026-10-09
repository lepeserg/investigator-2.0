# -*- coding: utf-8 -*-
"""_common.py — общие помощники скриптов investigator-sk (не запускается сам по себе).

  find_tool(name)          — поиск внешней программы (tesseract / soffice / antiword):
                             PATH → config.local.json tools.extra_path → стандартные
                             папки установки Windows → WinGet\\Links → маски (порядок run.py).
  check_locked(path)       — «документ открыт в Word / занят» → DocxLockedError.
  save_atomic(doc, path)   — tmp в той же папке → проверка zip → os.replace; tmp удаляется
                             при любом сбое.

Импорт из соседнего скрипта (как make_docx в docx_edit):
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from _common import find_tool
"""
import glob
import json
import os
import shutil
import zipfile

_SCRIPTS = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(os.path.dirname(_SCRIPTS))  # scripts → investigator-2-0 → корень


# ---------- поиск внешних программ ----------

# Имена для shutil.which и пути ОТНОСИТЕЛЬНО папок установки (Program Files*, LOCALAPPDATA).
_TOOLS = {
    "tesseract": {
        "names": ("tesseract",),
        "rel": (("Tesseract-OCR", "tesseract.exe"),),
        "globs": (),
        "fixed": (),
    },
    "soffice": {
        # soffice в PATH обычно НЕТ (установщик LibreOffice себя туда не прописывает) —
        # рабочая ветка это перебор стандартных папок установки.
        "names": ("soffice", "soffice.exe", "libreoffice"),
        "rel": (("LibreOffice", "program", "soffice.exe"),),
        "globs": (r"C:\Program Files\LibreOffice*\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice*\program\soffice.exe"),
        "fixed": ("/usr/bin/soffice", "/usr/bin/libreoffice",
                  "/Applications/LibreOffice.app/Contents/MacOS/soffice"),
    },
    "antiword": {
        "names": ("antiword",),
        "rel": (("antiword", "antiword.exe"),),
        "globs": (),
        "fixed": (r"C:\antiword\antiword.exe",),
    },
}
_ALIASES = {"libreoffice": "soffice"}


def _install_roots():
    """Папки установки: жёсткие C:\\Program Files* + переменные окружения (как system_tools)."""
    roots = [r"C:\Program Files", r"C:\Program Files (x86)"]
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            roots.append(base)
    return list(dict.fromkeys(roots))


def _config_extra_path():
    """tools.extra_path из config.local.json в корне сборки (если есть). run.py и так кладёт
    его в PATH; здесь — для прямого запуска скрипта мимо run.py."""
    try:
        with open(os.path.join(_ROOT, "config.local.json"), encoding="utf-8-sig") as f:
            extra = json.load(f)["tools"]["extra_path"]
        return [p for p in extra if isinstance(p, str) and p]
    except Exception:
        return []


def _winget_links():
    """%LOCALAPPDATA%\\Microsoft\\WinGet\\Links (ярлыки winget, как system_tools.search_paths)."""
    local = os.environ.get("LOCALAPPDATA")
    return [os.path.join(local, "Microsoft", "WinGet", "Links")] if local else []


def tool_candidates(name):
    """Упорядоченные пути-кандидаты в папках установки (без проверки существования)."""
    spec = _TOOLS[_ALIASES.get(name, name)]
    return list(dict.fromkeys(_install_candidates(spec) + list(spec["fixed"])))


def _install_candidates(spec):
    return [os.path.join(root, *rel) for root in _install_roots() for rel in spec["rel"]]


def _which_in(names, dirs):
    dirs = [d for d in dict.fromkeys(dirs) if os.path.isdir(d)]
    if dirs:
        for n in names:
            p = shutil.which(n, path=os.pathsep.join(dirs))
            if p:
                return p
    return None


def find_tool(name):
    """Полный путь к программе или None. name: 'tesseract' | 'soffice' ('libreoffice') | 'antiword'.

    Порядок — как PATH, который собирает run.py (system_tools.environment/search_paths), чтобы
    при прямом запуске и через run.py выбиралась одна и та же программа:
    PATH → config tools.extra_path → папки установки (Program Files*, LOCALAPPDATA) →
    WinGet\\Links → фиксированные пути → маски (берётся последняя по сортировке — самая новая
    версия). PATH первым — как в прежних копиях поиска (через run.py он уже начинается с
    тех же папок в том же порядке)."""
    spec = _TOOLS[_ALIASES.get(name, name)]
    for n in spec["names"]:
        p = shutil.which(n)
        if p:
            return p
    p = _which_in(spec["names"], _config_extra_path())
    if p:
        return p
    for c in _install_candidates(spec):
        if os.path.isfile(c):
            return c
    p = _which_in(spec["names"], _winget_links())
    if p:
        return p
    for c in spec["fixed"]:
        if os.path.isfile(c):
            return c
    for pat in spec["globs"]:
        hits = [h for h in glob.glob(pat) if os.path.isfile(h)]
        if hits:
            return sorted(hits)[-1]
    return None


# ---------- занятость документа (Word lock-файл ~$ + проба r+b) ----------

class DocxLockedError(IOError):
    """Файл занят (обычно открыт в Word). Закрыть документ и повторить."""


class BrokenSaveError(IOError):
    """python-docx сохранил битый (не zip) файл — оригинал не тронут."""


def word_lock_path(path):
    """Путь к lock-файлу Word (~$...) если документ открыт, иначе None.
    Word роняет первые 2 символа имени при len>=2, поэтому проверяем оба варианта."""
    d, name = os.path.split(path)
    d = d or "."
    cands = ["~$" + name]
    if len(name) >= 2:
        cands.append("~$" + name[2:])
    for c in cands:
        if os.path.exists(os.path.join(d, c)):
            return os.path.join(d, c)
    # ⛔ Эвристический поиск «любого ~$*.docx с похожим именем» удалён (боевой отчёт
    # 05.08.2026): он объявлял занятым документ, который никто не открывал. Проверяем ТОЛЬКО
    # два точных варианта имени; остальное ловит попытка открыть файл на запись (check_locked).
    return None


def busy_message(path, exc=None, *, stage="записать"):
    """Текст отказа по занятому файлу — РАЗЛИЧАЯ Word и любой другой держатель дескриптора.

    ⛔ Провал 07.09.2026: при ЛЮБОМ занятом дескрипторе говорилось «открыт в Word», а держал
    его другой процесс (незакрытый zipfile). Признак Word — ЕГО lock-файл `~$<имя>` рядом."""
    name = os.path.basename(path)
    lp = word_lock_path(path)
    code = ""
    if exc is not None:
        wn, se = getattr(exc, "winerror", None), getattr(exc, "strerror", None)
        parts = [p for p in (("WinError %s" % wn) if wn else None, se) if p]
        if parts:
            code = " [" + "; ".join(parts) + "]"
    if lp:
        return ("Не удалось %s «%s»%s — документ ОТКРЫТ В WORD (рядом его lock-файл «%s»). "
                "Закрой документ и повтори — правки НЕ внесены."
                % (stage, name, code, os.path.basename(lp)))
    return ("Не удалось %s «%s»%s — файл занят ДРУГИМ ПРОЦЕССОМ. Lock-файла Word «~$…» рядом НЕТ, "
            "поэтому закрывать Word, скорее всего, бесполезно: ищи незакрытый дескриптор у себя "
            "(zipfile/open без close, антивирус, индексатор, открытый предпросмотр). "
            "Правки НЕ внесены." % (stage, name, code))


def check_locked(path):
    """Двойная проверка занятости: lock-файл ~$ и попытка открыть на запись."""
    if not os.path.exists(path):
        return
    if word_lock_path(path):
        raise DocxLockedError(busy_message(path, stage="править"))
    try:
        with open(path, "r+b"):
            pass
    except PermissionError as e:
        raise DocxLockedError(busy_message(path, e, stage="открыть на запись"))


# ---------- атомарное сохранение ----------

def save_atomic(doc, path, *, tmp_prefix="~edit_", before_replace=None):
    """Атомарно сохранить doc в path: tmp в ТОЙ ЖЕ папке -> проверка zip ->
    before_replace(tmp) (метаданные, бэкап) -> os.replace(tmp, path).

    Битый zip → BrokenSaveError; PermissionError при замене → DocxLockedError.
    Временный файл удаляется при ЛЮБОМ сбое (doc.save, битый zip, хук, os.replace, SystemExit)."""
    d = os.path.dirname(os.path.abspath(path)) or "."
    pid = os.getpid()
    k = 0
    tmp = os.path.join(d, "%s%d_%d.docx" % (tmp_prefix, pid, k))
    while os.path.exists(tmp):  # уникальность при конкурентных правках в одном процессе
        k += 1
        tmp = os.path.join(d, "%s%d_%d.docx" % (tmp_prefix, pid, k))
    replaced = False
    try:
        doc.save(tmp)
        if not zipfile.is_zipfile(tmp):
            raise BrokenSaveError("python-docx сохранил битый файл — оригинал НЕ тронут.")
        if before_replace is not None:
            before_replace(tmp)
        try:
            os.replace(tmp, path)
        except PermissionError as e:
            raise DocxLockedError(busy_message(path, e, stage="записать")) from e
        replaced = True
    finally:
        if not replaced:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
    return path
