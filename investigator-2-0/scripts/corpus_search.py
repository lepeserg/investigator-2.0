# -*- coding: utf-8 -*-
"""corpus_search.py — полнотекстовый поиск по КОРПУСУ ДЕЛ (.docx / .doc / .txt / .md).

Зачем: Everything (`es`) ищет по ИМЕНАМ файлов и не видит содержимое документов.
Этот скрипт открывает каждый .docx (word/document.xml) и .doc (cp1251) и ищет
по тексту. Именно он закрывает критическое правило 30 скилла: прежде чем объявить
факт неустановленным или спросить у пользователя — искать по корпусу.

Параметры:
    ЗАПРОС...        одно или несколько слов; файл попадает в выдачу, если содержит ЛЮБОЕ
                     (по умолчанию) либо ВСЕ (--all) слова
    --root PATH      корень корпуса дел (по умолчанию — текущая папка);
                     синонимы --path / --dir принимаются наравне (провал 05.09.2026: `--path`
                     давал ошибку парсера, ход терялся на опечатке в имени флага)
    --context N      сколько символов контекста вокруг совпадения (по умолчанию 180)
    --max-hits N     сколько фрагментов показывать на файл (по умолчанию 3)
    --all            файл должен содержать ВСЕ слова запроса
    --ignore-case    регистронезависимый поиск
    --files-only     только список файлов, без контекста
    --regex RE       дополнительно вытащить все совпадения регулярки (например, ФИО рядом с позывным)
    --unique         для --regex: вывести уникальные значения группы 1 сводным списком
    --ext .docx,.doc какие расширения обходить

Код возврата: 0 — что-то найдено, 1 — не найдено (удобно для гейта)."""

import argparse
import os
import re
import sys
import zipfile

try:  # иначе кириллица в выводе ломается на Windows-консоли (cp866/cp1251)
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

DEFAULT_EXT = (".docx", ".doc", ".txt", ".md", ".rtf")

# ⛔ Ниже этого числа обойдённых файлов вывод «В КОРПУСЕ НЕ НАЙДЕНО» не печатается:
# он означал бы, что искали не там (см. блок в конце main).
MIN_CORPUS_FILES = 200
# Подсказка для сообщения об ошибке: типовой корень корпуса дел.
DEFAULT_ROOT = os.getcwd()
SKIP_DIRS = {".git", "__pycache__", "_to_delete", "~$"}


def extract_text(path):
    """Достать текст из файла. Возвращает str или None, если не удалось."""
    low = path.lower()
    try:
        if low.endswith(".docx"):
            with zipfile.ZipFile(path) as z:
                parts = []
                for name in z.namelist():
                    # тело + колонтитулы + сноски: там тоже бывают реквизиты
                    if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", name):
                        parts.append(z.read(name).decode("utf-8", "ignore"))
            xml = "\n".join(parts)
            xml = re.sub(r"</w:p>", "\n", xml)
            txt = re.sub(r"<[^>]+>", "", xml)
            return (txt.replace("&amp;", "&").replace("&lt;", "<")
                       .replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'"))
        if low.endswith(".doc"):
            # бинарный .doc: сначала antiword (чистый текст, есть локально; в Cowork нет),
            # иначе — грубое извлечение читаемого текста в cp1251
            try:
                import subprocess
                r = subprocess.run(["antiword", "-w", "0", "-m", "UTF-8.txt", path],
                                   capture_output=True, timeout=60)
                if r.returncode == 0 and r.stdout:
                    return r.stdout.decode("utf-8", "replace")
            except Exception:
                pass
            raw = open(path, "rb").read().decode("cp1251", "ignore")
            return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]+", " ", raw)
        if low.endswith((".txt", ".md", ".rtf")):
            return open(path, encoding="utf-8", errors="ignore").read()
    except Exception:
        return None
    return None


def norm(s):
    """Схлопнуть пробелы/переносы — Word рвёт слова по runs и строкам."""
    return re.sub(r"[\s\u00a0]+", " ", s)


def main():
    ap = argparse.ArgumentParser(description="Полнотекстовый поиск по корпусу дел")
    ap.add_argument("query", nargs="+", help="слово или фраза (можно несколько)")
    # --path / --dir — синонимы --root (dest остаётся `root`). Ошибка парсера на `--path`
    # стоила отдельного хода: скрипт не запускался, а сообщение argparse про «unrecognized
    # arguments» выглядит как поломка инструмента, а не как опечатка в имени флага.
    ap.add_argument("--root", "--path", "--dir", default=".", dest="root",
                    help="корень корпуса дел (синонимы: --path, --dir)")
    ap.add_argument("--context", type=int, default=180)
    ap.add_argument("--max-hits", type=int, default=3)
    ap.add_argument("--all", action="store_true", help="файл должен содержать ВСЕ слова")
    ap.add_argument("--ignore-case", action="store_true")
    ap.add_argument("--files-only", action="store_true")
    ap.add_argument("--regex", default=None, help="дополнительно вытащить совпадения регулярки")
    ap.add_argument("--unique", action="store_true", help="сводка уникальных значений группы 1 --regex")
    ap.add_argument("--ext", default=",".join(DEFAULT_EXT))
    a = ap.parse_args()

    exts = tuple(e.strip().lower() for e in a.ext.split(",") if e.strip())
    flags = re.IGNORECASE if a.ignore_case else 0
    # пробел в запросе матчим как любой пробельный набор — Word переносит строки
    pats = [re.compile(r"[\s\u00a0]+".join(re.escape(w) for w in q.split(" ")), flags)
            for q in a.query]
    extra = re.compile(a.regex, flags) if a.regex else None

    scanned = 0
    files_hit = 0
    uniq = set()

    for root, dirs, files in os.walk(a.root):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS and not d.startswith("~$")]
        for fn in sorted(files):
            if fn.startswith("~$") or not fn.lower().endswith(exts):
                continue
            path = os.path.join(root, fn)
            txt = extract_text(path)
            if not txt:
                continue
            scanned += 1
            flat = norm(txt)
            found = [p for p in pats if p.search(flat)]
            ok = (len(found) == len(pats)) if a.all else bool(found)
            if not ok:
                continue
            files_hit += 1
            rel = os.path.relpath(path, a.root)
            print(f"\n### {rel}")
            if a.files_only:
                continue
            shown = 0
            for p in found:
                for m in p.finditer(flat):
                    if shown >= a.max_hits:
                        break
                    s = max(0, m.start() - a.context)
                    e = min(len(flat), m.end() + a.context)
                    frag = flat[s:e].strip()
                    print(f"    …{frag}…")
                    shown += 1
                if shown >= a.max_hits:
                    break
            if extra:
                for m in extra.finditer(flat):
                    val = m.group(1) if m.groups() else m.group(0)
                    uniq.add(val.strip())
                    if not a.unique:
                        print(f"    [regex] {val.strip()[:160]}")

    print(f"\n{'=' * 60}")
    print(f"Просмотрено файлов: {scanned}   |   Найдено в файлах: {files_hit}")
    if extra and a.unique:
        print(f"\nУникальных значений по --regex: {len(uniq)}")
        for v in sorted(uniq):
            print(f"   • {v[:180]}")
    if files_hit == 0:
        # ⛔ «В КОРПУСЕ НЕ НАЙДЕНО» — сильное утверждение: на нём висит правило 30
        # (право поставить «подлежит установлению»). Оно ЛОЖНО, если обход шёл не по
        # корпусу. Реальный провал 24.08.2026: запуск из папки scripts дал
        # «Просмотрено файлов: 2» и «В КОРПУСЕ НЕ НАЙДЕНО» — вывод был бы неверным.
        # Порог 200 файлов: любая настоящая папка дел крупнее, корпус целиком — тысячи.
        if scanned < MIN_CORPUS_FILES:
            print(f"⛔ ВЫВОД О КОРПУСЕ НЕ ДЕЛАЕТСЯ: обойдено всего {scanned} файл(ов) — это не корпус.")
            print(f"   Похоже, --root указывает не туда (сейчас: {os.path.abspath(a.root)}).")
            print(f"   Повторить с корнем корпуса, например:")
            print(f'     python corpus_search.py "{a.query[0]}" --root "{DEFAULT_ROOT}"')
            print("   ⛔ Пометку «подлежит установлению» по такому прогону НЕ ставить (правило 30).")
            sys.exit(2)
        print("В КОРПУСЕ НЕ НАЙДЕНО — только после этого допустимо ставить пометку "
              "«подлежит установлению» либо спрашивать у пользователя (правило 30).")
    sys.exit(0 if files_hit else 1)


if __name__ == "__main__":
    main()
