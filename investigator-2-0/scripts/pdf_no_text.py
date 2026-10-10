# -*- coding: utf-8 -*-
"""pdf_no_text.py — отчёт «PDF БЕЗ ТЕКСТОВОГО СЛОЯ» по корпусу дел (+ пакетный OCR).

Зачем: в корпусе лежат сотни сканов, которые никогда не распознавались. `corpus_search`
их НЕ ВИДИТ — он читает текст, а в таком PDF текста нет, есть картинка страницы. Из-за
этого отрицательный результат поиска выглядит как «документа в деле нет», и об этом
не сигналит ни один инструмент (провал 04.09.2026: 198 файлов на 10 849 страниц, среди
них сканы дел на 650 и 272 листа — правило 30 на них не работало вовсе).

Скрипт обходит корпус, открывает каждый PDF и смотрит, есть ли на страницах извлекаемый
текст. Страницы проверяются ВЫБОРКОЙ (равномерно по документу), поэтому обход тысяч
файлов занимает минуты, а не часы.

Классификация:
    нет слоя   — ни на одной проверенной странице нет извлекаемого текста;
    частично   — текст есть меньше чем на половине проверенных страниц (сшивка скана
                 с обычным документом: часть листов ищется, часть нет — самый коварный
                 случай, поиск даёт ложное «нашлось не всё»);
    есть       — в отчёт не попадает.

⛔ По умолчанию скрипт НИЧЕГО НЕ РАСПОЗНАЁТ — только считает и печатает. Пакетный OCR
на десять тысяч страниц идёт часы и переписывает файлы дела, поэтому запускается
отдельным осознанным решением владельца: флаг --ocr.
⚠ OCR помогает только на МАШИНОПИСНЫХ листах. Рукописные объяснения и заявления
(их в корпусе много) дают на выходе мусор — текстовый слой появится, читать его нельзя;
поиск по такому файлу останется бесполезным, а «слой есть» усыпит бдительность.

Флаги:
    --root PATH     корень корпуса (по умолчанию — текущая папка)
    --sample N      сколько страниц проверять в документе (по умолчанию 12; 0 — все)
    --min-chars N   от скольких символов на странице считать, что текст есть (по умолчанию 20)
    --min-pages N   не показывать документы короче N страниц
    --limit N       показать только первые N строк отчёта (итог считается по всем)
    --partial       включить в отчёт и «частично» распознанные
    --csv PATH      выгрузить полный список в CSV (;-разделитель, UTF-8 с BOM для Excel)
    --ocr           РАСПОЗНАТЬ найденные файлы через ocrmypdf (правка файлов на месте)
    --ocr-limit N   распознать только первые N файлов списка
    --ocr-lang L    языки для ocrmypdf (по умолчанию rus+eng)
    --dry-run       с --ocr: показать команды, но не запускать

ВЫХОД: список «страниц · тип · путь», отсортированный по объёму, и итог
«N файлов, M страниц». Код возврата: 0 — таких PDF нет; 1 — есть (годится как гейт);
2 — ошибка запуска (нет корня, нет PyMuPDF, нет ocrmypdf)."""
import argparse
import importlib
import io
import os
import subprocess
import sys
import time

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

SKIP_DIRS = {".git", "__pycache__", "_to_delete", "node_modules", ".venv", "venv"}


def iter_pdfs(root):
    """Все .pdf под корнем, кроме служебных папок и вордовских времянок «~$…»."""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS and not d.startswith("~$")]
        for name in filenames:
            if name.lower().endswith(".pdf") and not name.startswith("~$"):
                yield os.path.normpath(os.path.join(dirpath, name))


def sample_indexes(n_pages, sample):
    """Номера страниц для проверки: первая, последняя и равномерно между ними."""
    if sample <= 0 or n_pages <= sample:
        return list(range(n_pages))
    step = (n_pages - 1) / float(sample - 1)
    return sorted({int(round(i * step)) for i in range(sample)})


def inspect(path, sample=12, min_chars=20):
    """Смотрит PDF. Возвращает (страниц, тип, проверено, со_слоем) либо (0, 'ошибка', …).

    Тип: «нет слоя» · «частично» · «есть». Ошибка чтения — тоже результат: битый
    или запаролённый файл поиском тоже не берётся, и знать о нём надо.
    """
    import pymupdf as fitz
    try:      # ворох «MuPDF error: syntax error…» от кривых сканов забивает отчёт
        fitz.TOOLS.mupdf_display_errors(False)
    except Exception:
        pass
    try:
        doc = fitz.open(path)
    except Exception:
        return 0, "ошибка", 0, 0
    try:
        n = doc.page_count
        if n == 0:
            return 0, "ошибка", 0, 0
        idx = sample_indexes(n, sample)
        with_text = 0
        for i in idx:
            try:
                if len(doc.load_page(i).get_text("text").strip()) >= min_chars:
                    with_text += 1
            except Exception:
                pass
        if with_text == 0:
            kind = "нет слоя"
        elif with_text * 2 < len(idx):
            kind = "частично"
        else:
            kind = "есть"
        return n, kind, len(idx), with_text
    finally:
        doc.close()


def scan(root, sample=12, min_chars=20, progress=True):
    """Обходит корпус. Возвращает (список результатов, сколько PDF всего обойдено)."""
    rows, total = [], 0
    t0 = time.time()
    for path in iter_pdfs(root):
        total += 1
        pages, kind, checked, with_text = inspect(path, sample, min_chars)
        rows.append({"path": path, "pages": pages, "kind": kind,
                     "checked": checked, "with_text": with_text})
        if progress and total % 100 == 0:
            sys.stderr.write("  …обойдено %d PDF (%.0f с)\r" % (total, time.time() - t0))
            sys.stderr.flush()
    if progress:
        sys.stderr.write(" " * 60 + "\r")
        sys.stderr.flush()
    return rows, total


def run_ocr(rows, lang="rus+eng", limit=0, dry=False):
    """Пакетное распознавание найденных файлов через ocrmypdf, ПРАВКА НА МЕСТЕ.

    Идёт последовательно и печатает время на файл: на сотнях листов это десятки
    минут, и владелец должен видеть, где остановиться (Ctrl+C прерывает между
    файлами, начатый файл ocrmypdf доводит сам либо не трогает оригинал).
    """
    todo = rows[:limit] if limit else rows
    ok = fail = 0
    for i, r in enumerate(todo, 1):
        cmd = ["ocrmypdf", "--language", lang, "--skip-text", "--quiet",
               r["path"], r["path"]]
        print("[%d/%d] %s (%d стр.)" % (i, len(todo), r["path"], r["pages"]))
        if dry:
            print("        " + " ".join(cmd))
            continue
        t0 = time.time()
        try:
            res = subprocess.run(cmd, capture_output=True)
        except FileNotFoundError:
            print("ocrmypdf не найден в PATH — распознавание невозможно.")
            return 2
        if res.returncode == 0:
            ok += 1
            print("        готово за %.0f с" % (time.time() - t0))
        else:
            fail += 1
            err = (res.stderr or b"").decode("utf-8", "ignore").strip().splitlines()
            print("        ОШИБКА: %s" % (err[-1] if err else res.returncode))
    if dry:
        print("\n--dry-run: ничего не распознано, показаны только команды.")
        return 0
    print("\nOCR: распознано %d, с ошибкой %d." % (ok, fail))
    return 0 if not fail else 1


def main():
    ap = argparse.ArgumentParser(description="Отчёт «PDF без текстового слоя» по корпусу")
    ap.add_argument("--root", default=".", help="корень корпуса")
    ap.add_argument("--sample", type=int, default=12, help="страниц на документ (0 — все)")
    ap.add_argument("--min-chars", type=int, default=20, help="символов на странице = текст есть")
    ap.add_argument("--min-pages", type=int, default=0, help="не показывать короче N страниц")
    ap.add_argument("--limit", type=int, default=0, help="показать первые N строк")
    ap.add_argument("--partial", action="store_true", help="включить «частично» распознанные")
    ap.add_argument("--csv", help="выгрузить список в CSV")
    ap.add_argument("--ocr", action="store_true", help="распознать найденное (ocrmypdf)")
    ap.add_argument("--ocr-limit", type=int, default=0, help="распознать первые N файлов")
    ap.add_argument("--ocr-lang", default="rus+eng", help="языки ocrmypdf")
    ap.add_argument("--dry-run", action="store_true", help="с --ocr: только показать команды")
    args = ap.parse_args()

    if not os.path.isdir(args.root):
        print("Нет такой папки: %s" % args.root)
        return 2
    try:
        # проверка наличия PyMuPDF: модуль только импортируется, не используется
        importlib.import_module("pymupdf")
    except ImportError:
        print("Не установлен PyMuPDF (pip install pymupdf) — читать PDF нечем.")
        return 2

    rows, total = scan(args.root, args.sample, args.min_chars)
    kinds = ("нет слоя", "частично") if args.partial else ("нет слоя",)
    bad = [r for r in rows if r["kind"] in kinds and r["pages"] >= args.min_pages]
    bad.sort(key=lambda r: -r["pages"])
    broken = [r for r in rows if r["kind"] == "ошибка"]

    print("Корпус: %s\nОбойдено PDF: %d" % (os.path.abspath(args.root), total))
    if not bad:
        print("PDF без текстового слоя не найдено.")
    else:
        print("\nБЕЗ ТЕКСТОВОГО СЛОЯ (поиск по содержимому их не видит):\n")
        shown = bad[:args.limit] if args.limit else bad
        for r in shown:
            print("%6d стр. · %-9s · %s" % (r["pages"], r["kind"], r["path"]))
        if len(shown) < len(bad):
            print("   … и ещё %d файлов" % (len(bad) - len(shown)))
        print("\nИТОГО: %d файлов, %d страниц." % (len(bad), sum(r["pages"] for r in bad)))
    if broken:
        print("\nНе открылись (битые или под паролем): %d — %s"
              % (len(broken), broken[0]["path"]))

    if args.csv:
        with io.open(args.csv, "w", encoding="utf-8-sig", newline="") as fh:
            fh.write("страниц;тип;проверено страниц;из них с текстом;путь\n")
            for r in bad:
                fh.write("%d;%s;%d;%d;%s\n" % (r["pages"], r["kind"], r["checked"],
                                               r["with_text"], r["path"].replace(";", ",")))
        print("CSV: %s" % os.path.abspath(args.csv))

    if args.ocr and bad:
        print("\nOCR: %d файлов, %d страниц. На сотнях страниц это часы; "
              "прервать можно Ctrl+C между файлами."
              % (len(bad[:args.ocr_limit] if args.ocr_limit else bad),
                 sum(r["pages"] for r in (bad[:args.ocr_limit] if args.ocr_limit else bad))))
        return run_ocr(bad, args.ocr_lang, args.ocr_limit, args.dry_run)

    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
