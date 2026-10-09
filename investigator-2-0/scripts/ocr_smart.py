# -*- coding: utf-8 -*-
"""ocr_smart.py — умное распознавание скана для investigator-sk.
Порядок движков: marker/surya (GPU, лучшее качество печати) → Tesseract (rus) как откат.

  python ocr_smart.py <файл.pdf|изображение> [начало] [конец]

Печатает распознанный текст (marker → markdown) с шапкой [OCR: <движок>].
Для больших сканов вызывай постранично: ocr_smart.py \"<pdf>\" 1 8, затем 9 16 и т.д."""
import sys, os, subprocess, tempfile, glob, shutil


def _utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def find_marker():
    p = os.environ.get("MARKER_SINGLE")
    if p and os.path.isfile(p):
        return p
    cands = []
    env = os.environ.get("MARKER_ENV")
    if env:
        cands.append(os.path.join(env, "Scripts", "marker_single.exe"))
    cands.append(os.path.expanduser(r"~\marker-env\Scripts\marker_single.exe"))
    for c in cands:
        if os.path.isfile(c):
            return c
    return None


def find_tesseract():
    if shutil.which("tesseract"):
        return "tesseract"
    p = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    return p if os.path.isfile(p) else None


# Предел времени: marker на большом скане идёт десятки минут, но не часами; зависание без предела
# блокировало бы сессию. Tesseract считается постранично — ему хватает минут на страницу.
MARKER_TIMEOUT = 3600
TESSERACT_PAGE_TIMEOUT = 600


def resolve_page_range(start, end, page_count):
    """Единая семантика диапазона для всех движков: (первая, последняя) с 1, включительно.

    Только начало → до конца документа; только конец → с первой страницы; конец обрезается по
    page_count. Без start и end → None (весь документ). page_count None (неизвестно) допустим
    лишь при явном конце."""
    if start is None and end is None:
        return None
    if start is not None and start < 1 or end is not None and end < 1:
        raise ValueError('Нумерация страниц начинается с 1')
    s = start or 1
    if page_count is None:
        if end is None:
            raise ValueError('Не удалось определить число страниц для диапазона «с %d до конца»' % s)
        e = end
    else:
        e = min(end or page_count, page_count)
    if s > e:
        raise ValueError('Диапазон страниц пуст или выходит за пределы PDF')
    return s, e


def _is_pdf(src):
    return os.path.splitext(src)[1].lower() == ".pdf"


def _pdf_page_count(src):
    import pymupdf as fitz
    with fitz.open(src) as doc:
        return doc.page_count


def run_marker(marker, src, start, end):
    tmp = tempfile.mkdtemp(prefix="ocr_marker_")
    try:
        cmd = [marker, src, "--force_ocr", "--disable_ocr_math",
               "--output_format", "markdown", "--output_dir", tmp]
        # Для изображения диапазон не имеет смысла (одна страница) — page_range не передаём.
        if (start or end) and _is_pdf(src):
            count = None
            try:
                count = _pdf_page_count(src)
            except ImportError:
                pass
            a, b = resolve_page_range(start, end, count)
            cmd += ["--page_range", "%d-%d" % (a - 1, b - 1)]
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                           errors="replace", timeout=MARKER_TIMEOUT)
        mds = glob.glob(os.path.join(tmp, "**", "*.md"), recursive=True)
        if r.returncode == 0 and mds:
            with open(mds[0], encoding="utf-8") as f:
                return f.read()
        sys.stderr.write((r.stderr or "")[-600:] + "\n")
        return None
    except subprocess.TimeoutExpired:
        sys.stderr.write("marker не завершился за %d с (MARKER_TIMEOUT) — процесс прерван. "
                         "Дели скан на диапазоны страниц поменьше.\n" % MARKER_TIMEOUT)
        return None
    except Exception as e:
        sys.stderr.write("marker error: %s\n" % e)
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


class OCRFailure(RuntimeError):
    """A failed engine invocation, distinct from a successfully read blank page."""


def _recognize_page(tess, image, base, label):
    """Check the engine and its expected text artifact before reading it."""
    try:
        process = subprocess.run([tess, image, base, "-l", "rus"], capture_output=True,
                                 timeout=TESSERACT_PAGE_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise OCRFailure('%s: Tesseract не завершился за %d с — процесс прерван' %
                         (label, TESSERACT_PAGE_TIMEOUT))
    if process.returncode != 0:
        diagnostic = (process.stderr or process.stdout or b'').decode('utf-8', 'replace').strip()
        raise OCRFailure('%s: Tesseract завершился с кодом %s. %s' %
                         (label, process.returncode, diagnostic[-600:]))
    output = base + '.txt'
    if not os.path.isfile(output):
        raise OCRFailure(label + ': Tesseract не создал ожидаемый текстовый файл')
    with open(output, encoding='utf-8') as handle:
        return handle.read()


def _import_pymupdf():
    """PyMuPDF нужен для разбора PDF на страницы. Без него — понятная ошибка, а не traceback."""
    try:
        import pymupdf as fitz
    except ImportError:
        raise OCRFailure('для постраничного OCR PDF нужен PyMuPDF (pymupdf), он не установлен. '
                         'Установи базовый профиль: start.cmd или py -3.12 bootstrap.py --profile base')
    return fitz


def check_image_range(start, end):
    """Изображение — одна страница: допустима только страница 1 (одинаково для всех движков)."""
    if start not in (None, 1) or end not in (None, 1):
        raise ValueError('Для изображения допустима только страница 1')


def run_tesseract(tess, src, start, end):
    out = []
    tmp = tempfile.mkdtemp(prefix="ocr_tess_")
    try:
        if start is not None and start < 1 or end is not None and end < 1:
            raise ValueError('Нумерация страниц начинается с 1')
        if _is_pdf(src):
            fitz = _import_pymupdf()
            with fitz.open(src) as doc:
                s, e = resolve_page_range(start or 1, end, doc.page_count)
                for i in range(s - 1, e):
                    png = os.path.join(tmp, "p%04d.png" % (i + 1))
                    doc.load_page(i).get_pixmap(dpi=300).save(png)
                    base = os.path.join(tmp, "o%04d" % (i + 1))
                    text = _recognize_page(tess, png, base, '%s, страница %d' % (src, i + 1))
                    out.append("\n----- стр. %d -----\n" % (i + 1) + text)
        else:
            check_image_range(start, end)
            out.append(_recognize_page(tess, src, os.path.join(tmp, 'o'), src))
        return "".join(out)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    _utf8()
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    if sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        sys.exit(0)
    src = sys.argv[1]
    try:
        start = int(sys.argv[2]) if len(sys.argv) > 2 else None
        end = int(sys.argv[3]) if len(sys.argv) > 3 else None
        if start is not None and start < 1 or end is not None and end < 1:
            raise ValueError('Нумерация страниц начинается с 1')
        if start is not None and end is not None and start > end:
            raise ValueError('Начало диапазона позже конца')
        # Проверка ДО выбора движка: marker для изображения диапазон игнорирует, и без
        # этого marker и Tesseract вели бы себя по-разному на одной команде.
        if not _is_pdf(src):
            check_image_range(start, end)
    except ValueError as exc:
        sys.stderr.write('Ошибка диапазона: %s\n' % exc)
        return 2
    if not os.path.isfile(src):
        sys.stderr.write("Нет файла: %s\n" % src)
        sys.exit(2)

    marker = find_marker()
    if marker:
        text = run_marker(marker, src, start, end)
        if text is not None:
            print("[OCR: marker/surya]")
            sys.stdout.write(text)
            return 0
        sys.stderr.write("marker недоступен/упал — откат на Tesseract\n")

    tess = find_tesseract()
    if not tess:
        sys.stderr.write("Ни marker, ни Tesseract не найдены. Поставь marker-env либо Tesseract+rus.\n")
        sys.exit(3)
    try:
        text = run_tesseract(tess, src, start, end)
    except (OCRFailure, OSError, ValueError) as exc:
        sys.stderr.write('OCR НЕ ЗАВЕРШЕНО: %s\nЧастичный результат не выдаётся как полный.\n' % exc)
        return 1
    print("[OCR: Tesseract rus (откат)]")
    sys.stdout.write(text)
    return 0


if __name__ == "__main__":
    sys.exit(main())
