# -*- coding: utf-8 -*-
"""check_env.py - proverka toolchain dlja investigator-sk.

Document-ingestion core: Python-pakety, Tesseract (+ rus), Ghostscript.
Donor-poisk: Everything (`es`) — otdel'no proverjaetsja «ne zapushhen» (Error 8)
i «zapushhen, no index PUST» (es otvechaet uspehom i pustym spiskom).
Audio/video (optional): ffmpeg, torch+CUDA, whisperx, HF token dlja diarizacii.
Sistemnye programmy ishhutsja ne tol'ko v PATH, no i v standartnyh papkah
ustanovki (chtoby ne bylo lozhnogo MISS, esli PATH eshhjo ne obnovilsja).

Obyazatel'nost' komponentov zadajut USTANOVLENNYE profili (base/documents/ocr/audio):
metka .venv/investigator-install.json (bootstrap.py) i runtime/setup-report.json
(setup_windows.py), libo argument --profile. Komponenty nevybrannyh profilej
pechatajutsja kak «ne ustanovleno (profil' X ne vybran)» i kod vozvrata ne ronjajut.
Bez metok (staraja/ruchnaja ustanovka) proverjaetsja prezhnij polnyj nabor:
base + documents + ocr (audio, kak i ran'she, ne objazatelen).
"""
import sys
import os
import glob
import shutil
import importlib
import importlib.util

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _common import find_tool  # noqa: E402


def _utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def find_tesseract():
    return find_tool("tesseract")


def find_ghostscript():
    p = shutil.which("gswin64c") or shutil.which("gs")
    if p:
        return p
    for pat in (r"C:\Program Files\gs\*\bin\gswin64c.exe",
                r"C:\Program Files (x86)\gs\*\bin\gswin32c.exe"):
        hits = glob.glob(pat)
        if hits:
            return sorted(hits)[-1]
    return None


def find_ffmpeg():
    p = shutil.which("ffmpeg")
    if p:
        return p
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        hits = glob.glob(os.path.join(local, r"Microsoft\WinGet\Packages\Gyan.FFmpeg*\**\bin\ffmpeg.exe"),
                         recursive=True)
        if hits:
            return sorted(hits)[-1]
    return None


def find_libreoffice():
    """soffice в PATH обычно НЕТ — рабочая ветка это поиск по стандартным папкам
    установки (единый перебор: _common.find_tool)."""
    return find_tool("soffice")


def find_es():
    """es.exe (Everything CLI) — канал донор-поиска. В PATH попадает не всегда:
    winget кладёт его в LOCALAPPDATA\\Microsoft\\WinGet\\Packages\\voidtools.Everything.Cli*."""
    for name in ("es", "es.exe"):
        p = shutil.which(name)
        if p:
            return p
    cands = []
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        cands += glob.glob(os.path.join(local, r"Microsoft\WinGet\Packages\voidtools.Everything*", "**", "es.exe"),
                           recursive=True)
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        base = os.environ.get(env)
        if base:
            cands.append(os.path.join(base, "Everything", "es.exe"))
    for c in cands:
        if os.path.exists(c):
            return c
    return None


def probe_es(es_path):
    """Различает ТРИ состояния канала Everything — вторая беда тише и опаснее первой:
      'down'    — Everything не запущен: `Error 8: Everything IPC not found`, код возврата 8;
      'empty'   — Everything ЗАПУЩЕН, но проиндексировал НОЛЬ файлов (типично при запуске
                  без прав администратора). `es` при этом отвечает УСПЕХОМ и пустым списком,
                  то есть «ничего не найдено» неотличимо от «донора нет»;
      'ok'      — в индексе есть файлы.
    Возвращает (state, note)."""
    import subprocess

    def _run(args):
        return subprocess.run([es_path] + args, capture_output=True, text=True,
                              errors="replace", timeout=40)

    try:
        out = _run(["-timeout", "5000", "-no-digit-grouping", "-get-result-count"])
    except Exception as e:
        return "unknown", "es не запустился: %s" % e
    txt = ((out.stdout or "") + (out.stderr or "")).strip()
    first = txt.splitlines()[0].strip() if txt else ""
    if out.returncode == 8 or "IPC not found" in txt:
        return "down", "Everything НЕ ЗАПУЩЕН (%s)" % (first or "код возврата 8")
    digits = "" if (out.returncode != 0 or first.lower().startswith("error")) \
        else "".join(ch for ch in first if ch.isdigit())
    if digits:
        n = int(digits)
        if n == 0:
            return "empty", "Everything запущен, но в индексе 0 файлов (запуск без прав администратора)"
        return "ok", "в индексе %d файлов" % n
    # Сборка es без `-get-result-count` (или непонятный ответ): пробуем получить хотя бы одну
    # строку выдачи. Пустая выдача при коде 0 — это и есть ТИХИЙ отказ, а не «файлов нет».
    try:
        out2 = _run(["-timeout", "5000", "-n", "1"])
    except Exception as e:
        return "unknown", "es не запустился: %s" % e
    txt2 = ((out2.stdout or "") + (out2.stderr or "")).strip()
    if out2.returncode == 8 or "IPC not found" in txt2:
        return "down", "Everything НЕ ЗАПУЩЕН (Error 8: IPC not found)"
    if out2.returncode != 0 or txt2.lower().startswith("error"):
        return "unknown", (txt2.splitlines()[0].strip() if txt2 else "код возврата %d" % out2.returncode)
    if txt2:
        return "ok", "индекс не пуст (проверено выдачей одной строки)"
    return "empty", "es отвечает успехом, но не выдал НИ ОДНОГО файла — индекс пуст"


def find_word_com():
    """Word COM (pywin32) — на нём держатся fill_formfields, post_release_gvp и
    конвертация в docx_edit. Раньше check_env о нём молчал и печатал READY."""
    if not sys.platform.startswith("win"):
        return None, "не Windows — путь через Word COM недоступен (это нормально)"
    try:
        # проверка наличия pywin32: модуль только импортируется, не используется
        importlib.import_module("win32com.client")
    except Exception:
        return False, "нет pywin32 -> pip install pywin32"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"):
            pass
        return True, "pywin32 + Word.Application зарегистрирован"
    except Exception:
        return False, "pywin32 есть, но Word не зарегистрирован (класс Word.Application не найден)"


# --- Профили установки ------------------------------------------------------
# Обязательность компонентов определяется УСТАНОВЛЕННЫМИ профилями, а не «всем
# сразу»: раньше ocrmypdf (профиль ocr) считался core, и после установки base
# проверка всегда печатала MISSING PIECES — на неё переставали смотреть.
PROFILES = ("base", "documents", "ocr", "audio")
REPO_ROOT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
INSTALL_STAMP = os.path.join(".venv", "investigator-install.json")    # пишет bootstrap.py
SETUP_REPORT = os.path.join("runtime", "setup-report.json")           # пишет setup_windows.py
# Без меток установщика (старая или ручная установка) проверяется прежний полный
# набор, как до введения профилей: document core + Word/LibreOffice + OCR.
# Audio и раньше был необязательным.
LEGACY_PROFILES = ("base", "documents", "ocr")
LEGACY_SOURCE = ("метки профилей не найдены — проверяется полный набор старой установки "
                 "(base, documents, ocr)")


def expand_profiles(names):
    """Нормализует список профилей: 'all' -> все, base добавляется всегда,
    неизвестные имена отбрасываются. Чистая функция."""
    out = {"base"}
    for name in names or ():
        name = str(name).strip().lower()
        if name == "all":
            out.update(PROFILES)
        elif name in PROFILES:
            out.add(name)
    return out


def detect_profiles(root=REPO_ROOT):
    """Установленные профили по меткам установщика.
    Возвращает (profiles, source); без меток — прежний полный набор LEGACY_PROFILES."""
    import json
    names, sources = [], []
    stamp = os.path.join(root, INSTALL_STAMP)
    try:
        with open(stamp, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("profiles"), list):
            names += data["profiles"]
            sources.append(INSTALL_STAMP.replace(os.sep, "/"))
    except (OSError, ValueError):
        pass
    report = os.path.join(root, SETUP_REPORT)
    try:
        with open(report, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict) and isinstance(data.get("profile"), str):
            names.append(data["profile"])
            sources.append(SETUP_REPORT.replace(os.sep, "/"))
    except (OSError, ValueError):
        pass
    if not sources:
        return expand_profiles(LEGACY_PROFILES), LEGACY_SOURCE
    return expand_profiles(names), ", ".join(sources)


def resolve_profiles(cli_profiles, root=REPO_ROOT):
    """--profile имеет приоритет над метками установщика."""
    if cli_profiles:
        return expand_profiles(cli_profiles), "аргумент --profile"
    return detect_profiles(root)


def install_hint(profile):
    """Текущая команда установки профиля (вместо устаревших runtime/python и
    scripts/requirements.txt)."""
    if profile == "documents":
        return "start.cmd -> пункт «Documents» (или setup-windows.ps1 -Profile documents)"
    return "start.cmd (setup-windows.ps1 -Profile %s) или py -3.12 bootstrap.py --profile %s" % (profile, profile)


def classify(ok, profile, profiles):
    """Статус одной проверки: 'ok' | 'miss' (профиль выбран, компонента нет —
    роняет код возврата) | 'skip' (профиль не выбран — не роняет). Чистая функция."""
    if ok:
        return "ok"
    return "miss" if profile in profiles else "skip"


def overall_ok(checks, profiles):
    """checks: [(profile, ok)]. True, если все компоненты ВЫБРАННЫХ профилей на месте."""
    return all(classify(ok, prof, profiles) != "miss" for prof, ok in checks)


def _print_check(checks, profiles, ok, profile, label, hint=None):
    checks.append((profile, bool(ok)))
    state = classify(ok, profile, profiles)
    if state == "ok":
        print("  [OK]   %s" % label)
    elif state == "miss":
        print("  [MISS] %s   -> %s" % (label, hint or install_hint(profile)))
    else:
        print("  [-]    %s: не установлено (профиль %s не выбран)" % (label, profile))


def _has_module(name, heavy=False):
    """heavy=False — только find_spec (не поднимает torch/whisperx зря)."""
    try:
        if heavy:
            importlib.import_module(name)
            return True
        # importlib.util импортирован на уровне модуля: локальный `import importlib.util`
        # делал имя importlib локальным во всей функции, и ветка heavy всегда падала
        # с UnboundLocalError → «пакета нет» при установленном пакете (поймано CI на Windows).
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def _parse_args(argv):
    import argparse
    ap = argparse.ArgumentParser(prog="check_env.py", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter,
                                 epilog="Без --profile профили берутся из .venv/investigator-install.json "
                                        "и runtime/setup-report.json (без меток — base, documents, ocr). "
                                        "Код возврата: 0 — компоненты выбранных профилей на месте, 1 — чего-то не хватает.")
    ap.add_argument("--profile", action="append", choices=PROFILES + ("all",),
                    help="проверять как обязательный профиль (можно несколько раз)")
    if argv and argv[0] in ("/?", "help"):
        argv = ["--help"]
    # Неизвестные аргументы игнорируются, как и до введения --profile.
    return ap.parse_known_args(argv)[0]


def main(argv=None):
    _utf8()
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    profiles, source = resolve_profiles(args.profile)
    checks = []
    order = [p for p in PROFILES if p in profiles]
    print("Профили: %s (источник: %s)" % (", ".join(order), source))
    if source == LEGACY_SOURCE:
        print("Метки профилей (.venv/investigator-install.json, runtime/setup-report.json) не найдены — "
              "проверяется полный набор: base, documents, ocr.")

    print("== Python packages (base) ==")
    for mod, pip in [("fitz", "pymupdf"), ("docx", "python-docx"),
                     ("pypdf", "pypdf"), ("PIL", "Pillow")]:
        _print_check(checks, profiles, _has_module(mod, heavy=True), "base", pip)

    # Ниже — НЕ core: ни один скрипт навыка их не импортирует (ревизия 22.08.2026),
    # их отсутствие не роняет проверку.
    print("== Python packages (optional) ==")
    for mod, pip in [("pdfplumber", "pdfplumber"), ("pytesseract", "pytesseract"),
                     ("cv2", "opencv-python")]:
        if _has_module(mod):
            print("  [OK]   %s" % pip)
        else:
            print("  [-]    %s (не используется скриптами навыка; ставить только по нужде)" % pip)

    print("== Word COM / рендер (профиль documents: бланки-формы, статкарты ГВП, docx -> PDF) ==")
    # Как в исходной проверке: на Windows Word COM обязателен сам по себе —
    # LibreOffice НЕ заменяет его для бланков-форм (fill_formfields, post_release_gvp),
    # он нужен только для рендера docx -> PDF. Вне Windows Word COM недоступен,
    # и для рендера обязателен LibreOffice.
    w_ok, w_note = find_word_com()
    lo = find_libreoffice()
    if w_ok is None:
        print("  [-]    win32com: %s" % w_note)
        _print_check(checks, profiles, lo, "documents",
                     "libreoffice/soffice: %s" % (lo or "не найден"))
    else:
        _print_check(checks, profiles, w_ok, "documents", "pywin32 / Word: %s" % w_note,
                     "Word COM обязателен для бланков-форм; LibreOffice его не заменяет. "
                     + install_hint("documents"))
        print("  %s libreoffice/soffice: %s" % (
            "[OK]  " if lo else "[-]   ",
            lo or "нет; при доступном Word использовать отдельный экземпляр Word COM"))

    print("== OCR (профиль ocr) ==")
    _print_check(checks, profiles, _has_module("ocrmypdf", heavy=True), "ocr", "ocrmypdf")
    tess = find_tesseract()
    _print_check(checks, profiles, tess, "ocr", "tesseract: %s" % (tess or "not found"),
                 "start.cmd (setup-windows.ps1 -Profile ocr) — ставит Tesseract и языки")
    rus = False
    if tess:
        import subprocess
        try:
            out = subprocess.run([tess, "--list-langs"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            rus = any(line.strip() == "rus" for line in (out.stdout or "").splitlines())
        except Exception:
            pass
    _print_check(checks, profiles, rus, "ocr", "tesseract Russian (rus)",
                 "start.cmd (setup-windows.ps1 -Profile ocr) — докачает rus/eng/osd")
    # Движок PDF для OCRmyPDF: достаточно ЛЮБОГО из двух — Ghostscript или pypdfium2.
    gs = find_ghostscript()
    print("  %s ghostscript: %s" % ("[OK]  " if gs else "[-]   ", gs or "optional: используется pypdfium2"))
    pdfium = _has_module("pypdfium2", heavy=True)
    print("  %s pypdfium2: альтернативный движок PDF для OCRmyPDF 17+" % ("[OK]  " if pdfium else "[-]   "))
    _print_check(checks, profiles, bool(gs or pdfium), "ocr",
                 "движок PDF для OCRmyPDF (ghostscript или pypdfium2)")

    print("== Донор-поиск (Everything / es) ==")
    es_state, es_note = "down", "es.exe не найден"
    es = find_es()
    if not es:
        print("  [MISS] es.exe не найден (ни в PATH, ни в папках winget/Program Files)")
    else:
        es_state, es_note = probe_es(es)
        mark = {"ok": "OK", "empty": "ПУСТО", "down": "MISS", "unknown": "?"}[es_state]
        print("  [%s] %s: %s" % (mark, es, es_note))
    if es_state != "ok":
        print("  ⛔ Команду `es` НЕ повторять и владельца запускать приложение НЕ просить.")
        print("     Обход: python scripts/corpus_search.py \"<фрагмент>\" --root \"<корень корпуса>\" --files-only")
        print("     плюс обход каталогов из Python (os.walk/glob). Программный обход РАЗРЕШЁН —")
        print("     запрет «обходить папки руками» касается только ручного клика по проводнику.")
    if es_state == "empty":
        print("  ⛔ ОПАСНО: es отвечает УСПЕХОМ и пустым списком — «не найдено» здесь не значит")
        print("     «донора нет». Everything без прав администратора индексирует ноль файлов.")

    print("== Audio/Video (профиль audio) ==")
    audio_sel = "audio" in profiles
    ff = find_ffmpeg()
    _print_check(checks, profiles, ff, "audio", "ffmpeg: %s" % (ff or "not found"))
    torch_ok = False
    if audio_sel:
        try:
            import torch
            cuda = torch.cuda.is_available()
            dev = torch.cuda.get_device_name(0) if cuda else "CPU only"
            torch_ok = True
            print("  [OK]   torch %s (CUDA: %s, %s)" % (torch.__version__, cuda, dev))
        except Exception:
            _print_check(checks, profiles, False, "audio", "torch")
    else:
        torch_ok = _has_module("torch")
        _print_check(checks, profiles, torch_ok, "audio", "torch")
    whisper_ok = _has_module("whisperx", heavy=audio_sel)
    _print_check(checks, profiles, whisper_ok, "audio", "whisperx")
    av_ok = bool(ff and torch_ok and whisper_ok)
    hf = bool(os.environ.get("HF_TOKEN") or os.environ.get("HUGGINGFACE_TOKEN"))
    if not hf and sys.platform == 'win32':
        try:
            from setup_hf_access import saved_token
            hf = bool(saved_token())
        except Exception:
            pass
    print("  %s доступ для диаризации сохранён (сеть: setup_hf_access.py --verify)" % ("[OK]  " if hf else "[-]   "))

    ok = overall_ok(checks, profiles)
    print()
    print("Выбранные профили (%s): " % ", ".join(order)
          + ("READY" if ok else "MISSING PIECES (see above)"))
    skipped = [p for p in PROFILES if p not in profiles]
    if skipped:
        print("Не выбраны (на код возврата не влияют): %s. Добавить: start.cmd или "
              "setup-windows.ps1 -Profile <имя>; Python-пакеты вне Windows: "
              "py -3.12 bootstrap.py --profile <base|ocr|audio> (профиль documents — "
              "только через start.cmd / setup-windows.ps1)" % ", ".join(skipped))
    print("Донор-поиск (es): " + {
        "ok": "READY",
        "empty": "СЛОМАН ТИХО — индекс пуст, ответы es пустые; идти через corpus_search.py + os.walk",
        "down": "Everything не запущен; идти через corpus_search.py + os.walk (es не повторять)",
        "unknown": "состояние неясно; не полагаться, идти через corpus_search.py + os.walk",
    }[es_state])
    if av_ok:
        print("Audio/Video dependencies: READY; модели и диаризацию проверять отдельно")
    elif audio_sel:
        print("Audio/Video dependencies: MISSING -> " + install_hint("audio"))
    else:
        print("Audio/Video dependencies: не установлено (профиль audio не выбран)")
    if not ok:
        need = [p for p in order if any(prof == p and not good for prof, good in checks)]
        print("Доустановка: start.cmd (меню) или setup-windows.ps1 -Profile <%s>; "
              "проверка без установки: setup-windows.ps1 -Profile <имя> -CheckOnly. "
              "Python-пакеты вне Windows: py -3.12 bootstrap.py --profile <base|ocr|audio> "
              "(documents bootstrap не принимает)." % "|".join(need))

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
