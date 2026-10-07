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


def run_marker(marker, src, start, end):
    tmp = tempfile.mkdtemp(prefix="ocr_marker_")
    try:
        cmd = [marker, src, "--force_ocr", "--disable_ocr_math",
               "--output_format", "markdown", "--output_dir", tmp]
        if start or end:
            a = start or 1
            b = end or a
            cmd += ["--page_range", "%d-%d" % (a - 1, b - 1)]
        r = subprocess.run(cmd, capture_output=True, text=True)
        mds = glob.glob(os.path.join(tmp, "**", "*.md"), recursive=True)
        if r.returncode == 0 and mds:
            with open(mds[0], encoding="utf-8") as f:
                return f.read()
        sys.stderr.write((r.stderr or "")[-600:] + "\n")
        return None
    except Exception as e:
        sys.stderr.write("marker error: %s\n" % e)
        return None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def run_tesseract(tess, src, start, end):
    import fitz
    out = []
    tmp = tempfile.mkdtemp(prefix="ocr_tess_")
    try:
        if os.path.splitext(src)[1].lower() == ".pdf":
            doc = fitz.open(src)
            s = start or 1
            e = min(end or doc.page_count, doc.page_count)
            for i in range(s - 1, e):
                png = os.path.join(tmp, "p%04d.png" % (i + 1))
                doc.load_page(i).get_pixmap(dpi=300).save(png)
                base = os.path.join(tmp, "o%04d" % (i + 1))
                subprocess.run([tess, png, base, "-l", "rus"], capture_output=True)
                out.append("\n----- стр. %d -----\n" % (i + 1))
                if os.path.exists(base + ".txt"):
                    with open(base + ".txt", encoding="utf-8") as f:
                        out.append(f.read())
            doc.close()
        else:
            base = os.path.join(tmp, "o")
            subprocess.run([tess, src, base, "-l", "rus"], capture_output=True)
            if os.path.exists(base + ".txt"):
                with open(base + ".txt", encoding="utf-8") as f:
                    out.append(f.read())
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
    start = int(sys.argv[2]) if len(sys.argv) > 2 else None
    end = int(sys.argv[3]) if len(sys.argv) > 3 else None
    if not os.path.isfile(src):
        sys.stderr.write("Нет файла: %s\n" % src)
        sys.exit(2)

    marker = find_marker()
    if marker:
        text = run_marker(marker, src, start, end)
        if text is not None:
            print("[OCR: marker/surya]")
            sys.stdout.write(text)
            return
        sys.stderr.write("marker недоступен/упал — откат на Tesseract\n")

    tess = find_tesseract()
    if not tess:
        sys.stderr.write("Ни marker, ни Tesseract не найдены. Поставь marker-env либо Tesseract+rus.\n")
        sys.exit(3)
    print("[OCR: Tesseract rus (откат)]")
    sys.stdout.write(run_tesseract(tess, src, start, end))


if __name__ == "__main__":
    main()
