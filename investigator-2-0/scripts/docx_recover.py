# -*- coding: utf-8 -*-
"""
docx_recover.py – восстановление «битых» .docx (обрезаны при облачной синхронизации
или блокировке Word: потеряна центральная директория zip / EOCD).

Стратегия:
1. Если zipfile.is_zipfile – файл цел, ничего не делаем.
2. Иначе сканируем сырые байты по сигнатурам локальных заголовков PK\\x03\\x04,
   извлекаем записи (stored/deflate) и пересобираем валидный .zip.
3. Если восстановлен хотя бы word/document.xml – сохраняем <имя>_recovered.docx.

Использование:
    python docx_recover.py "путь/к/битому.docx"
"""
import sys, os, struct, zlib, zipfile

LFH = b"PK\x03\x04"


def _iter_entries(data):
    n = len(data); i = 0
    while True:
        j = data.find(LFH, i)
        if j < 0 or j + 30 > n:
            break
        (_sig, _ver, _flags, method, _t, _d, _crc, comp, _unc,
         fnlen, eflen) = struct.unpack("<IHHHHHIIIHH", data[j:j + 30])
        ns = j + 30
        raw = data[ns:ns + fnlen]
        ds = ns + fnlen + eflen
        try:
            name = raw.decode("utf-8")
        except UnicodeDecodeError:
            name = raw.decode("cp866", "replace")
        if comp:
            blob = data[ds:ds + comp]; i = ds + comp
        else:
            nxt = data.find(LFH, ds); blob = data[ds:(nxt if nxt > 0 else n)]; i = ds + len(blob)
        yield name, method, blob


def recover(path):
    if zipfile.is_zipfile(path):
        print("Файл валиден, восстановление не требуется."); return path
    data = open(path, "rb").read()
    out = path.rsplit(".", 1)[0] + "_recovered.docx"
    got = {}
    for name, method, blob in _iter_entries(data):
        if not name or name.endswith("/"):
            continue
        try:
            got[name] = blob if method == 0 else zlib.decompress(blob, -15)
        except Exception:
            continue
    if "word/document.xml" not in got:
        raise SystemExit("Не удалось восстановить word/document.xml – файл слишком повреждён.")
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in got.items():
            z.writestr(name, content)
    print(f"Восстановлено записей: {len(got)} -> {out}")
    return out


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    if len(sys.argv) < 2:
        print("usage: python docx_recover.py <битый.docx>"); sys.exit(1)
    if sys.argv[1] in ("-h", "--help", "/?", "help"):
        print(__doc__); sys.exit(0)
    if sys.argv[1].startswith("-"):
        sys.exit("Неизвестный аргумент: %s\nСправка: python docx_recover.py --help" % sys.argv[1])
    if not os.path.isfile(sys.argv[1]):
        sys.exit("Нет файла: %s" % sys.argv[1])
    recover(sys.argv[1])
