# -*- coding: utf-8 -*-
"""
pdf_ingest.py - приём входящих PDF/сканов для investigator.

Подкоманды:
  info   <pdf>                              -> JSON о структуре (стр., текст-слой, скан?, dpi, размер)
  render <pdf> <start> <end> [dpi] [outdir] -> рендер страниц в JPEG, печатает пути
  text   <pdf> [start] [end]                -> извлечение текстового слоя (если есть)
  ocr    <pdf> <out.pdf> [lang]             -> добавить искомый текстовый слой (ocrmypdf)

ДВИЖОК И ОКРУЖЕНИЕ.
  - Если установлен PyMuPDF (fitz) — используется он (локальная машина с полным тулчейном).
  - Если PyMuPDF НЕТ (песочница claude.ai/Cowork/analysis tool) — АВТООТКАТ на poppler
    (pdfinfo/pdftotext/pdftoppm), который в песочнице есть. Подкоманды info/render/text
    работают в ОБЕИХ средах без изменения команды.
  - Принудительно выбрать движок: переменная окружения PDF_INGEST_ENGINE=poppler|fitz.
  - Подкоманда ocr (искомый текст-слой) требует OCRmyPDF 17+ + Tesseract(rus) + pypdfium2 —
    это ЛОКАЛЬНАЯ машина; в песочнице её обычно нет, там путь «render + чтение зрением».

Примеры:
  python pdf_ingest.py info "C:\\dela\\materialy.pdf"
  python pdf_ingest.py render "C:\\dela\\materialy.pdf" 1 8 200
  python pdf_ingest.py text "C:\\dela\\dogovor.pdf"
  python pdf_ingest.py ocr "C:\\dela\\skan.pdf" "C:\\dela\\skan_ocr.pdf" rus
"""
import sys
import os
import json
import glob
import subprocess


def _stdout_utf8():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def _engine():
    """Какой движок использовать: 'fitz' или 'poppler'."""
    forced = (os.environ.get("PDF_INGEST_ENGINE") or "").strip().lower()
    if forced in ("fitz", "poppler"):
        return forced
    try:
        import pymupdf as fitz  # noqa: F401
        return "fitz"
    except Exception:
        return "poppler"


# ---------------------------------------------------------------- poppler (fallback)
def _pop_pagecount(path):
    try:
        r = subprocess.run(["pdfinfo", path], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=60)
        for line in r.stdout.splitlines():
            if line.lower().startswith("pages:"):
                return int(line.split(":", 1)[1].strip())
    except Exception:
        pass
    try:
        from pypdf import PdfReader
        return len(PdfReader(path).pages)
    except Exception:
        return None


def _pop_text(path, start=None, end=None):
    cmd = ["pdftotext", "-layout"]
    if start:
        cmd += ["-f", str(start)]
    if end:
        cmd += ["-l", str(end)]
    cmd += [path, "-"]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                           timeout=600)
    except subprocess.TimeoutExpired:
        sys.stderr.write("pdftotext не ответил за 600 с — текст не извлечён. "
                         "Задай диапазон страниц поменьше.\n")
        return ""
    return r.stdout or ""


def _info_poppler(path):
    size_mb = round(os.path.getsize(path) / (1024.0 * 1024.0), 1)
    pages = _pop_pagecount(path)
    sample = _pop_text(path, 1, 5)
    text_layer = len(sample.strip()) > 200
    scanned = not text_layer
    too_big = size_mb > 100
    rec = "render+read" if (scanned or too_big or (pages or 0) > 20) else "read_or_text"
    return {
        "path": path, "pages": pages, "size_mb": size_mb, "engine": "poppler",
        "text_layer_pages": "n/a", "image_only_pages": "n/a",
        "is_scanned": scanned, "approx_dpi": None,
        "too_big_for_builtin_read": too_big, "recommendation": rec,
    }


def _render_poppler(path, start, end, dpi, outdir):
    os.makedirs(outdir, exist_ok=True)
    root = os.path.join(outdir, "p")
    try:
        subprocess.run(["pdftoppm", "-jpeg", "-r", str(dpi), "-f", str(start), "-l", str(end),
                        path, root], check=True, timeout=1800)
    except subprocess.TimeoutExpired:
        sys.stderr.write("pdftoppm не ответил за 30 мин — рендер прерван. "
                         "Задай диапазон страниц поменьше или уменьши DPI.\n")
        sys.exit(3)
    for f in sorted(glob.glob(root + "-*.jpg")):
        try:
            num = int(os.path.splitext(os.path.basename(f))[0].rsplit("-", 1)[-1])
        except ValueError:
            continue
        if start <= num <= end:
            print(f)


def _text_poppler(path, start, end):
    sys.stdout.write(_pop_text(path, start, end))


# ---------------------------------------------------------------- fitz (PyMuPDF)
def _info_fitz(path):
    import pymupdf as fitz
    doc = fitz.open(path)
    n = doc.page_count
    size_mb = round(os.path.getsize(path) / (1024.0 * 1024.0), 1)
    step = max(1, n // 10)
    sample = sorted(set(list(range(min(15, n))) + list(range(0, n, step)) + ([n - 1] if n else [])))
    text_pages = 0
    image_only = 0
    dpi_guess = None
    for i in sample:
        page = doc.load_page(i)
        txt = page.get_text("text").strip()
        imgs = page.get_images(full=True)
        if len(txt) > 40:
            text_pages += 1
        elif imgs:
            image_only += 1
            if dpi_guess is None:
                try:
                    pix = fitz.Pixmap(doc, imgs[0][0])
                    dpi_guess = round(pix.width / (page.rect.width / 72.0))
                except Exception:
                    pass
    doc.close()
    scanned = image_only >= max(1, text_pages)
    too_big = size_mb > 100
    rec = "render+read" if (scanned or too_big or n > 20) else "read_or_text"
    return {
        "path": path, "pages": n, "size_mb": size_mb, "engine": "fitz",
        "sampled_pages": len(sample), "text_layer_pages": text_pages,
        "image_only_pages": image_only, "is_scanned": scanned, "approx_dpi": dpi_guess,
        "too_big_for_builtin_read": too_big, "recommendation": rec,
    }


def _render_fitz(path, start, end, dpi, outdir):
    import pymupdf as fitz
    os.makedirs(outdir, exist_ok=True)
    doc = fitz.open(path)
    end = min(end, doc.page_count)
    for i in range(start - 1, end):
        page = doc.load_page(i)
        pix = page.get_pixmap(dpi=dpi)
        fp = os.path.join(outdir, "p%04d.jpg" % (i + 1))
        pix.pil_save(fp, format="JPEG", quality=82)
        print(fp)
    doc.close()


def _text_fitz(path, start, end):
    import pymupdf as fitz
    doc = fitz.open(path)
    end = min(end, doc.page_count)
    parts = []
    for i in range(start - 1, end):
        parts.append("\n----- стр. %d -----\n" % (i + 1))
        parts.append(doc.load_page(i).get_text("text"))
    doc.close()
    sys.stdout.write("".join(parts))


# ---------------------------------------------------------------- команды
def _default_outdir(path):
    stem = os.path.splitext(os.path.basename(path))[0]
    return os.path.join(os.path.dirname(os.path.abspath(path)), "_render_" + stem)


def cmd_info(args):
    path = args[0]
    out = _info_fitz(path) if _engine() == "fitz" else _info_poppler(path)
    print(json.dumps(out, ensure_ascii=False, indent=2))


def cmd_render(args):
    path = args[0]
    start = int(args[1])
    end = int(args[2])
    dpi = int(args[3]) if len(args) > 3 else 200
    outdir = args[4] if len(args) > 4 else _default_outdir(path)
    if _engine() == "fitz":
        _render_fitz(path, start, end, dpi, outdir)
    else:
        _render_poppler(path, start, end, dpi, outdir)


def cmd_text(args):
    path = args[0]
    start = int(args[1]) if len(args) > 1 else 1
    end = int(args[2]) if len(args) > 2 else None
    if _engine() == "fitz":
        _text_fitz(path, start, end or 10 ** 9)
    else:
        _text_poppler(path, start, end)


def cmd_ocr(args):
    path = args[0]
    out = args[1]
    lang = args[2] if len(args) > 2 else "rus"
    # Searchable PDF does not require Ghostscript/PDF-A conversion.
    cmd = [sys.executable, "-m", "ocrmypdf", "-l", lang, "--skip-text", "--output-type", "pdf", path, out]
    try:
        r = subprocess.run(cmd)
    except FileNotFoundError:
        sys.stderr.write("Не удалось запустить локальный Python для OCRmyPDF. Используйте run.py и check_env.py.\n")
        sys.exit(3)
    sys.exit(r.returncode)


def main():
    _stdout_utf8()
    if len(sys.argv) > 1 and sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__)
        sys.exit(0)
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    cmd = sys.argv[1]
    args = sys.argv[2:]
    table = {"info": cmd_info, "render": cmd_render, "text": cmd_text, "ocr": cmd_ocr}
    fn = table.get(cmd)
    if not fn:
        sys.stderr.write("Unknown command: %s\n" % cmd)
        print(__doc__)
        sys.exit(2)
    fn(args)


if __name__ == "__main__":
    main()
