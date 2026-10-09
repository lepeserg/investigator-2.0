# -*- coding: utf-8 -*-
"""docx_integrity.py — ПОСТГЕЙТ ЦЕЛОСТНОСТИ .docx (чек-лист §17.1 пункт 14).

Только стандартная библиотека — работает и в Cowork (там нет ни Word, ни pywin32).

ПОДКОМАНДЫ
  check <файл|папка> [...]        Проверить. Код возврата 1, если есть дефекты.
  fix   <файл|папка> [...]        Починить безопасно исправимое (см. ниже) и перепроверить.
  compare <до.docx> <после.docx>  Сверить состав zip до/после правки: не появились ли
                                  материализованные колонтитулы, не пропали ли части,
                                  сошлись ли поля формы и защита.

ТО ЖЕ ИЗ PYTHON (имя функции = имя подкоманды, менять CLI не нужно):
  from docx_integrity import check, fix, compare, inspect, repair
  r = check(path)          # -> CheckResult; bool(r) == True ТОЛЬКО если дефектов нет
  if not r: print(r.problems)          # список «файл: дефект»
  r.raise_if_bad()                     # или сразу исключение
  fix(path)                # починить безопасное и перепроверить -> CheckResult
⛔ `if check(path):` на кортеже всегда истинно — поэтому check возвращает объект,
  который сам по себе ложен при дефектах. Отчёт «сохранено» состоянием файла не является.

ЧТО ПРОВЕРЯЕТ
  · файл открывается как zip, `testzip()` == None;
  · все .xml/.rels разбираются парсером;
  · каждая часть объявлена в `[Content_Types].xml`;
  · цели связей (`.rels`) существуют;
  · `w:bookmarkStart` == `w:bookmarkEnd`, `w:fldChar begin` == `end`;
  · ссылки `r:id` в document.xml имеют связь;
  · (справочно) число полей формы `w:ffData` и наличие `w:documentProtection`.

ЧТО ЧИНИТ (`fix`) — только то, что не меняет видимый текст:
  · осиротевший `bookmarkStart` → сразу после него ставится парный `bookmarkEnd`
    (закладка нулевой длины: имя сохраняется, содержимое не двигается);
  · потерянные объявления пространств имён в `docProps/core.xml`.
Перед записью пересобранный файл проверяется заново И сверяется извлечённый текст
до символа — не совпало, файл НЕ трогается.

⛔ Не «чинит» битый zip и не восстанавливает усечённые файлы — это ручная операция."""
import os
import re
import sys
import zipfile
import argparse
import shutil
import time
from xml.etree import ElementTree as ET

from word_text import xml_text

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

NS_DECL = {
    "cp": 'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"',
    "dc": 'xmlns:dc="http://purl.org/dc/elements/1.1/"',
    "dcterms": 'xmlns:dcterms="http://purl.org/dc/terms/"',
    "dcmitype": 'xmlns:dcmitype="http://purl.org/dc/dcmitype/"',
    "xsi": 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"',
}
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
BM_START = re.compile(r"<w:bookmarkStart\b[^>]*?/>")
ID_RE = re.compile(r'w:id="(\d+)"')


def _text(doc):
    """Текст для сверки до/после починки: общие правила word_text.xml_text, причём удалённое
    при рецензировании (w:delText) тоже сверяется — починка не вправе трогать и его."""
    return xml_text(doc, include_deleted=True)


def parts_of(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist()}


def inspect(path):
    """-> (список проблем, справка)."""
    problems, info = [], {}
    try:
        z = zipfile.ZipFile(path)
    except Exception as e:
        return [f"не открывается как .docx: {e}"], info
    names = set(z.namelist())
    if z.testzip():
        problems.append(f"битая запись в архиве: {z.testzip()}")

    for n in names:
        if n.endswith((".xml", ".rels")):
            try:
                ET.fromstring(z.read(n))
            except Exception as e:
                problems.append(f"невалидный XML в {n}: {e}")

    if "[Content_Types].xml" not in names:
        problems.append("нет [Content_Types].xml")
    else:
        ctx = ET.fromstring(z.read("[Content_Types].xml"))
        defaults = {d.get("Extension", "").lower() for d in ctx.findall(CT_NS + "Default")}
        overrides = {o.get("PartName", "").lstrip("/") for o in ctx.findall(CT_NS + "Override")}
        for n in names:
            if n == "[Content_Types].xml" or n.endswith((".rels", "/")):
                continue
            if n not in overrides and os.path.splitext(n)[1].lstrip(".").lower() not in defaults:
                problems.append(f"часть не объявлена в [Content_Types]: {n}")

    for n in [x for x in names if x.endswith(".rels")]:
        base = os.path.dirname(os.path.dirname(n))
        try:
            rx = ET.fromstring(z.read(n))
        except Exception:
            continue
        for rel in rx.findall(REL_NS + "Relationship"):
            t = rel.get("Target", "")
            if rel.get("TargetMode") == "External" or t.startswith(("#", "http")):
                continue
            full = os.path.normpath(os.path.join(base, t)).replace("\\", "/").lstrip("/")
            if full not in names:
                problems.append(f"связь в пустоту: {n} -> {t}")

    if "word/document.xml" in names:
        doc = z.read("word/document.xml").decode("utf-8", "replace")
        bs = len(re.findall(r"<w:bookmarkStart\b", doc))
        be = len(re.findall(r"<w:bookmarkEnd\b", doc))
        if bs != be:
            problems.append(f"закладки не сходятся: bookmarkStart={bs}, bookmarkEnd={be} "
                            f"(незакрытая закладка роняет Word при сохранении)")
        fb = doc.count('<w:fldChar w:fldCharType="begin"')
        fe = doc.count('<w:fldChar w:fldCharType="end"')
        if fb != fe:
            problems.append(f"поля Word не сходятся: begin={fb}, end={fe}")
        rels = set()
        if "word/_rels/document.xml.rels" in names:
            rx = ET.fromstring(z.read("word/_rels/document.xml.rels"))
            rels = {r.get("Id") for r in rx.findall(REL_NS + "Relationship")}
        missing = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', doc)) - rels
        if missing:
            problems.append(f"ссылки r:id без связи: {sorted(missing)[:5]}")
        info["полей формы"] = doc.count("<w:ffData")
        info["защита формы"] = "<w:documentProtection" in doc
        info["закладок"] = bs

    info["частей"] = len(names)
    info["колонтитулов"] = len([n for n in names if re.match(r"word/(header|footer)\d+\.xml$", n)])
    z.close()
    return problems, info


def _fix_bookmarks(doc):
    ends = set(re.findall(r'<w:bookmarkEnd\b[^>]*?w:id="(\d+)"', doc))
    out, pos, added = [], 0, 0
    for m in BM_START.finditer(doc):
        mid = ID_RE.search(m.group(0))
        if not mid or mid.group(1) in ends:
            continue
        out.append(doc[pos:m.end()])
        out.append(f'<w:bookmarkEnd w:id="{mid.group(1)}"/>')
        pos = m.end()
        added += 1
    out.append(doc[pos:])
    return "".join(out), added


def _fix_core(core):
    m = re.search(r"<([A-Za-z0-9]+:coreProperties)\b[^>]*>", core)
    if not m:
        return core, []
    head, used, added = m.group(0), set(re.findall(r"([A-Za-z0-9]+):", core)), []
    for p in sorted(used):
        if p in NS_DECL and f"xmlns:{p}=" not in head:
            head = head[:-1] + " " + NS_DECL[p] + ">"
            added.append(p)
    return core[:m.start()] + head + core[m.end():], added


def _fix_truncated(xml):
    """Достроить XML, оборванный на середине (частая беда `docProps/app.xml`).

    Найдено 14.08.2026: в 5 документах-клонах одного донора `app.xml` обрывался на
    `<CharactersW` — корень `</Properties>` так и не закрывался. Word это терпит,
    парсер — нет. В app.xml лежит только статистика (страницы, слова, компания),
    поэтому обрезать хвост и закрыть открытые теги безопасно."""
    cut = xml.rfind(">")
    if cut < 0:
        return xml, False
    body = xml[:cut + 1]
    stack = []
    for m in re.finditer(r"<(/?)([A-Za-z0-9:._-]+)([^>]*?)(/?)>", body):
        closing, name, _, self_closing = m.groups()
        if name.lower() in ("?xml", "!doctype"):
            continue
        if closing:
            if stack and stack[-1] == name:
                stack.pop()
        elif not self_closing:
            stack.append(name)
    if not stack:
        return body, body != xml
    return body + "".join(f"</{t}>" for t in reversed(stack)), True


def repair(path, apply=True):
    parts = parts_of(path)
    with zipfile.ZipFile(path) as z:
        infos = {i.filename: i for i in z.infolist()}
    changes = []
    doc0 = parts.get("word/document.xml", b"").decode("utf-8", "surrogatepass")
    doc1, added = _fix_bookmarks(doc0)
    if added:
        changes.append(f"закрыто закладок: {added}")
        parts["word/document.xml"] = doc1.encode("utf-8", "surrogatepass")
    if "docProps/core.xml" in parts:
        c0 = parts["docProps/core.xml"].decode("utf-8", "surrogatepass")
        try:
            ET.fromstring(c0)
        except Exception:
            c1, ns = _fix_core(c0)
            try:
                ET.fromstring(c1)
                changes.append(f"восстановлены пространства имён: {', '.join(ns)}")
                parts["docProps/core.xml"] = c1.encode("utf-8", "surrogatepass")
            except Exception:
                changes.append("⚠ свойства битые, автоматически не чинятся")
    for part in ("docProps/app.xml", "docProps/core.xml"):
        if part not in parts:
            continue
        raw = parts[part].decode("utf-8", "surrogatepass")
        try:
            ET.fromstring(raw)
            continue
        except Exception:
            pass
        fixed, done = _fix_truncated(raw)
        if done:
            try:
                ET.fromstring(fixed)
                parts[part] = fixed.encode("utf-8", "surrogatepass")
                changes.append(f"достроен оборванный {os.path.basename(part)}")
            except Exception:
                pass
    if not changes or not apply:
        return changes, False

    # tmp собираем в ПАПКЕ САМОГО ДОКУМЕНТА, а не в системном TEMP: при документе на
    # другом томе (F:, сетевая шара) shutil.move из %TEMP% = копирование + удаление,
    # то есть НЕ атомарно, и обрыв посередине оставлял битый файл без исходника
    # (ревизия 22.08.2026). На одном томе os.replace атомарен.
    _dir = os.path.dirname(os.path.abspath(path)) or "."
    tmp = os.path.join(_dir, "~fix_%d.docx" % os.getpid())
    _k = 0
    while os.path.exists(tmp):
        _k += 1
        tmp = os.path.join(_dir, "~fix_%d_%d.docx" % (os.getpid(), _k))
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for n, data in parts.items():
            zi = infos.get(n)
            nz = zipfile.ZipInfo(n, date_time=zi.date_time) if zi else zipfile.ZipInfo(n)
            nz.compress_type = zipfile.ZIP_DEFLATED
            out.writestr(nz, data)
    probs, _ = inspect(tmp)
    new_doc = parts_of(tmp)["word/document.xml"].decode("utf-8", "surrogatepass")
    if probs or _text(new_doc) != _text(doc0):
        os.remove(tmp)
        return changes + [f"ОТМЕНА, файл не тронут: {probs or 'текст изменился'}"], False
    try:
        with open(path, "ab"):
            pass
    except OSError:
        os.remove(tmp)
        return changes + ["пропуск: файл занят (открыт в Word)"], False
    # Страховка перед заменой — по образцу docx_edit._save_atomic: если что-то пойдёт
    # не так, исходник (уже повреждённый, но единственный) остаётся рядом.
    bak = None
    try:
        stamp = time.strftime("%Y%m%d_%H%M%S")
        bak = path + ".bak_" + stamp
        _i = 1
        while os.path.exists(bak):
            bak = "%s.bak_%s_%d" % (path, stamp, _i); _i += 1
        shutil.copyfile(path, bak)
        try:  # ретенция .bak_* — та же, что у make_docx/docx_edit; без python-docx пропускаем
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            import make_docx
            make_docx.prune_backups(path)
        except Exception:
            pass
    except Exception:
        bak = None
    try:
        os.replace(tmp, path)
    except OSError as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return changes + ["ОТМЕНА, файл не тронут: не удалось заменить (%s)" % e], False
    return changes + ([f"страховка: {os.path.basename(bak)}"] if bak else []), True


def compare(before, after):
    """Что изменилось в СОСТАВЕ файла после правки (§17.1 п. 14)."""
    a, b = parts_of(before), parts_of(after)
    out = []
    new = sorted(set(b) - set(a))
    gone = sorted(set(a) - set(b))
    hdr = [n for n in new if re.match(r"word/(header|footer)\d+\.xml$", n)]
    if hdr:
        out.append(f"⚠ ПОЯВИЛИСЬ колонтитулы ({len(hdr)}): {', '.join(os.path.basename(x) for x in hdr)}"
                   " — обращение к свойству python-docx материализует их")
    if gone:
        out.append(f"⛔ ПРОПАЛИ части: {', '.join(gone)}")
    if [n for n in new if n not in hdr]:
        out.append(f"добавились части: {', '.join(n for n in new if n not in hdr)}")
    ia, ib = inspect(before)[1], inspect(after)[1]
    for k in ("полей формы", "защита формы", "закладок"):
        if ia.get(k) != ib.get(k):
            out.append(f"⚠ {k}: было {ia.get(k)}, стало {ib.get(k)}")
    return out or ["состав не изменился"]


# ───────────── импортируемое API: имя функции = имя подкоманды ─────────────
#
# Локальный пример исключён из публичной поставки.
# `ImportError: cannot import name 'check'` — `check` было только именем подкоманды CLI.
# Постгейт при этом пропускали. Теперь подкоманды `check` и `fix` существуют и как функции,
# поведение CLI не тронуто.

__all__ = ["check", "fix", "compare", "inspect", "repair", "parts_of", "CheckResult"]


class CheckResult:
    """Итог постгейта. ⛔ bool(результат) истинен ТОЛЬКО когда дефектов нет.

    .ok        — файл(ы) целы;
    .report    — {путь: [дефекты]};
    .problems  — плоский список «путь: дефект»;
    .info      — {путь: справка (части, закладки, поля формы, защита)};
    .fixed     — {путь: [что починено]} (заполняет `fix`, у `check` пусто).
    """

    def __init__(self, report, info=None, fixed=None):
        self.report = report
        self.info = info or {}
        self.fixed = fixed or {}
        self.problems = [f"{p}: {t}" for p, ts in report.items() for t in ts]
        self.ok = not self.problems
        self.files = list(report)

    def __bool__(self):
        return self.ok

    def __iter__(self):                      # чтобы `probs, info = ...` тоже работало
        return iter((self.problems, self.info))

    def __repr__(self):
        return ("<цел: %d файл(ов)>" % len(self.files) if self.ok
                else "<ДЕФЕКТЫ %d: %s>" % (len(self.problems), "; ".join(self.problems[:3])))

    def raise_if_bad(self):
        if not self.ok:
            raise RuntimeError("постгейт целостности не пройден: " + "; ".join(self.problems))
        return self


def check(*paths):
    """Постгейт целостности ИЗ PYTHON — то же, что CLI `check` (правило 24, §17.1 п. 14).
    Принимает файлы, папки и списки путей. -> CheckResult (ложен при дефектах)."""
    report, info = {}, {}
    for f in _collect(_flatten(paths)):
        probs, i = inspect(f)
        report[f], info[f] = probs, i
    return CheckResult(report, info)


def fix(*paths):
    """То же, что CLI `fix`: починить безопасно исправимое и ПЕРЕПРОВЕРИТЬ. -> CheckResult."""
    report, info, fixed = {}, {}, {}
    for f in _collect(_flatten(paths)):
        probs, i = inspect(f)
        if probs:
            changes, _done = repair(f)
            fixed[f] = changes
            probs, i = inspect(f)
        report[f], info[f] = probs, i
    return CheckResult(report, info, fixed)


def _flatten(paths):
    out = []
    for p in paths:
        if isinstance(p, (list, tuple, set)):
            out += list(p)
        else:
            out.append(p)
    return out


def _collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for dp, dn, fn in os.walk(p):
                files += [os.path.join(dp, f) for f in fn
                          if f.lower().endswith(".docx") and not f.startswith("~$")]
        else:
            files.append(p)
    return files


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[1])
    ap.add_argument("cmd", choices=["check", "fix", "compare"])
    ap.add_argument("paths", nargs="+")
    a = ap.parse_args()

    if a.cmd == "compare":
        if len(a.paths) != 2:
            print("compare требует два файла: до и после")
            return 2
        for line in compare(*a.paths):
            print(" ", line)
        return 0

    bad = 0
    for f in _collect(a.paths):
        probs, info = inspect(f)
        if a.cmd == "fix" and probs:
            changes, done = repair(f)
            probs, info = inspect(f)
            print(f"{'ПОЧИНЕН' if done else 'нужно  '} {f}\n     {'; '.join(changes)}")
        if probs:
            bad += 1
            print(f"✗ {f}")
            for p in probs:
                print(f"     {p}")
    print(f"\nпроверено {len(_collect(a.paths))}, с дефектами {bad}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
