# -*- coding: utf-8 -*-
"""check_env.py - proverka toolchain dlja investigator-sk.

Document-ingestion core: Python-pakety, Tesseract (+ rus), Ghostscript.
Donor-poisk: Everything (`es`) — otdel'no proverjaetsja «ne zapushhen» (Error 8)
i «zapushhen, no index PUST» (es otvechaet uspehom i pustym spiskom).
Audio/video (optional): ffmpeg, torch+CUDA, whisperx, HF token dlja diarizacii.
Sistemnye programmy ishhutsja ne tol'ko v PATH, no i v standartnyh papkah
ustanovki (chtoby ne bylo lozhnogo MISS, esli PATH eshhjo ne obnovilsja).
"""
import sys
import os
import glob
import shutil
import importlib


def _utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def find_tesseract():
    p = shutil.which("tesseract")
    if p:
        return p
    for c in (r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"):
        if os.path.exists(c):
            return c
    return None


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
    установки. Раньше пути были записаны с ЗАДВОЕННЫМИ слэшами внутри raw-строк
    (r"C:\\Program Files\\..."), такого пути на диске не существует, и
    LibreOffice не находился никогда (ревизия 22.08.2026)."""
    for name in ("soffice", "soffice.exe", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    cands = [r"C:\Program Files\LibreOffice\program\soffice.exe",
             r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]
    for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", "LOCALAPPDATA"):
        base = os.environ.get(env)
        if base:
            cands.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    for c in cands:
        if os.path.exists(c):
            return c
    for pat in (r"C:\Program Files\LibreOffice*\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice*\program\soffice.exe"):
        hits = glob.glob(pat)
        if hits:
            return sorted(hits)[-1]
    return None


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
        import win32com.client  # noqa: F401
    except Exception:
        return False, "нет pywin32 -> pip install pywin32"
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"):
            pass
        return True, "pywin32 + Word.Application зарегистрирован"
    except Exception:
        return False, "pywin32 есть, но Word не зарегистрирован (класс Word.Application не найден)"


def main():
    _utf8()
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        print("Запуск без аргументов: полная проверка окружения (~15 c, поднимает torch/CUDA).")
        print("Отдельная строка «Донор-поиск (es)»: READY / не запущен / индекс пуст.")
        print("Код возврата: 0 — document core готов, 1 — чего-то не хватает.")
        sys.exit(0)
    core_ok = True

    print("== Python packages (document core) ==")
    for mod, pip in [
        ("fitz", "pymupdf"), ("docx", "python-docx"),
        ("pypdf", "pypdf"), ("PIL", "Pillow"), ("ocrmypdf", "ocrmypdf"),
    ]:
        try:
            importlib.import_module(mod)
            print("  [OK]   %s" % pip)
        except Exception:
            core_ok = False
            print("  [MISS] %s   -> pip install %s" % (pip, pip))

    # Ниже — НЕ core: ни один скрипт навыка их не импортирует (ревизия 22.08.2026),
    # их отсутствие больше не роняет проверку в MISSING PIECES.
    print("== Python packages (optional) ==")
    for mod, pip in [("pdfplumber", "pdfplumber"), ("pytesseract", "pytesseract"),
                     ("cv2", "opencv-python")]:
        try:
            importlib.import_module(mod)
            print("  [OK]   %s" % pip)
        except Exception:
            print("  [-]    %s (не используется скриптами навыка; ставить только по нужде)" % pip)

    print("== Word COM (бланки-формы, статкарты ГВП, конвертация в docx_edit) ==")
    w_ok, w_note = find_word_com()
    if w_ok is None:
        print("  [-]    win32com: %s" % w_note)
    else:
        print("  [%s] pywin32 / Word: %s" % ("OK" if w_ok else "MISS", w_note))
        if not w_ok:
            core_ok = False

    print("== System tools (OCR) ==")
    tess = find_tesseract()
    print("  [%s] tesseract: %s" % ("OK" if tess else "MISS", tess or "not found"))
    rus = False
    if tess:
        import subprocess
        try:
            out = subprocess.run([tess, "--list-langs"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            rus = any(line.strip() == "rus" for line in (out.stdout or "").splitlines())
        except Exception:
            pass
    print("  [%s] tesseract Russian (rus)" % ("OK" if rus else "MISS"))
    gs = find_ghostscript()
    print("  [%s] ghostscript: %s" % ("OK" if gs else "-", gs or "optional: используется pypdfium2"))
    try:
        importlib.import_module('pypdfium2')
        pdfium = True
        print('  [OK]   pypdfium2: альтернативный движок PDF для OCRmyPDF 17+')
    except ImportError:
        pdfium = False
    if not (tess and rus and (gs or pdfium)):
        core_ok = False

    print("== Render check (docx -> PDF, визуальная верификация форм) ==")
    lo = find_libreoffice()
    print("  [%s] libreoffice/soffice: %s" % ("OK" if lo else "-", lo or "нет; при доступном Word использовать отдельный экземпляр Word COM"))
    if not lo and not w_ok:
        core_ok = False
    try:
        importlib.import_module("fitz")
        print("  [OK]   pymupdf (fitz) – чтение PDF при сверке")
    except Exception:
        print("  [MISS] pymupdf (fitz)   -> pip install pymupdf")

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

    print("== Audio/Video stack (optional) ==")
    ff = find_ffmpeg()
    print("  [%s] ffmpeg: %s" % ("OK" if ff else "MISS", ff or "not found"))
    av_ok = bool(ff)
    try:
        import torch
        cuda = torch.cuda.is_available()
        dev = torch.cuda.get_device_name(0) if cuda else "CPU only"
        print("  [OK]   torch %s (CUDA: %s, %s)" % (torch.__version__, cuda, dev))
    except Exception:
        av_ok = False
        print("  [MISS] torch   -> install whisperx in runtime/python")
    try:
        importlib.import_module("whisperx")
        print("  [OK]   whisperx")
    except Exception:
        av_ok = False
        print("  [MISS] whisperx   -> install whisperx in runtime/python")
    hf = bool(os.environ.get("HF_TOKEN") or os.environ.get("HUGGINGFACE_TOKEN"))
    if not hf and sys.platform == 'win32':
        from setup_hf_access import saved_token
        hf = bool(saved_token())
    print("  [%s] доступ для диаризации сохранён (сеть: setup_hf_access.py --verify)" % ("OK" if hf else "-"))

    print()
    print("Document pipeline: " + ("READY" if core_ok else "MISSING PIECES (see above)"))
    print("Донор-поиск (es): " + {
        "ok": "READY",
        "empty": "СЛОМАН ТИХО — индекс пуст, ответы es пустые; идти через corpus_search.py + os.walk",
        "down": "Everything не запущен; идти через corpus_search.py + os.walk (es не повторять)",
        "unknown": "состояние неясно; не полагаться, идти через corpus_search.py + os.walk",
    }[es_state])
    print("Audio/Video dependencies: " + ("READY; модели и диаризацию проверять отдельно" if av_ok else "not installed (install whisperx in runtime/python)"))
    if not core_ok:
        print("OCR: используйте run.py и окружение runtime/python; зависимости перечислены в scripts/requirements.txt.")

    sys.exit(0 if core_ok else 1)


if __name__ == "__main__":
    main()
