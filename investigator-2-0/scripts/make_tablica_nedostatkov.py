# -*- coding: utf-8 -*-
"""make_tablica_nedostatkov.py — «ТАБЛИЦА устранения недостатков» (ответ на замечания
военного прокурора / руководителя по изучаемому или возвращённому делу).

  · A4 АЛЬБОМНАЯ (29,7 x 21,0), поля 1,5 см со всех сторон;
  · «ТАБЛИЦА» — по центру, полужирным, TNR 12;
  · «устранения недостатков (замечаний …) по уголовному делу № … в отношении {ФИО в Р.п.}»
    — по центру, полужирным, TNR 12;
  · таблица 4 колонки: № п/п | Замечание (вопрос) | Принятые меры по устранению |
    Лист дела / документ; ширины 1,2 / 6,3 / 13,2 / 5,5 см; шапка — серая заливка D9D9D9,
    полужирная, по центру; все границы; TNR 12; «Меры» — по ширине (самая широкая колонка);
  · ПОСЛЕ таблицы — примечание о доказательственной базе («Событие преступления и вина …
    подтверждаются: …»);
  · подпись следователя (реквизиты — из констант персонального слоя).

Генератор, а не правка донора: таблица строится программно, поэтому «разъехаться» ей негде
(запрет генератора касается протоколов СД и бланков со сложными удостоверительными таблицами).
Первый готовый файл дальше можно использовать как донор.

Запуск:
    python make_tablica_nedostatkov.py --json данные.json --out \"…\\Таблица устранения недостатков.docx\"
    python make_tablica_nedostatkov.py --demo --out \"…\\demo.docx\""""
import os
import sys
import json
import argparse

os.environ.setdefault("PYTHONUTF8", "1")
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document                                    # noqa: E402
from docx.shared import Cm, Pt                                # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH                 # noqa: E402
from docx.enum.section import WD_ORIENT                       # noqa: E402
from docx.oxml.ns import qn                                   # noqa: E402
from docx.oxml import OxmlElement                             # noqa: E402

from make_docx import (                                       # noqa: E402
    _add_signature, _normalize_text, _backup_if_exists, FONT_NAME,
)

FONT_SIZE_TBL = Pt(12)
HEADER_FILL = "D9D9D9"
COL_WIDTHS_CM = (1.2, 6.3, 13.2, 5.5)          # снято с эталона
HEADERS = ("№\nп/п", "Замечание (вопрос) {кого}", "Принятые меры по устранению",
           "Лист дела / документ")


def _run(par, text, *, bold=False, size=FONT_SIZE_TBL):
    r = par.add_run(text)
    r.bold = bold
    r.font.name = FONT_NAME
    r.font.size = size
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rPr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(a), FONT_NAME)
    return r


def _shade(cell, fill=HEADER_FILL):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _borders(table, sz=4):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), '000000')
        borders.append(el)
    tblPr.append(borders)


def _fixed_widths(table, widths_cm):
    """Жёстко зафиксировать ширины колонок.

    ⚠ `table.autofit = False` + `cell.width` НЕДОСТАТОЧНО: Word (и python-docx при чтении)
    берёт ширину из сетки `w:gridCol`, и без `w:tblLayout fixed` колонки выходят равными
    (проверено: 6,67 см вместо 1,2/6,3/13,2/5,5). Поэтому пишем и сетку, и tcW."""
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    total = OxmlElement("w:tblW")
    total.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    total.set(qn("w:type"), "dxa")
    tblPr.append(total)
    grid = table._tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        for gc, w in zip(cols, widths_cm):
            gc.set(qn("w:w"), str(int(w * 567)))       # см -> twips (1 см = 567)
    for row in table.rows:
        for c, w in zip(row.cells, widths_cm):
            c.width = Cm(w)


def _cell(cell, text, *, bold=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, width_cm=None):
    cell.text = ""
    p = cell.paragraphs[0]
    for r in list(p.runs):          # пустой run от `cell.text = ""` — убрать, иначе в ячейке
        r._element.getparent().remove(r._element)   # остаётся run без шрифта (ловушка «TNR по умолчанию»)
    p.alignment = align
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    lines = str(text).split("\n")
    for i, line in enumerate(lines):
        if i:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
        _run(p, _normalize_text(line), bold=bold)
    if width_cm:
        cell.width = Cm(width_cm)
    return cell


def _title(doc, text, *, bold=True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(0)
    _run(p, _normalize_text(text), bold=bold)
    return p


def build_tablica(out_path, data, constants=None):
    """Собрать «Таблицу устранения недостатков» по данным `data` (см. JSON в шапке модуля)."""
    items = data.get("пункты") or []
    if not items:
        raise ValueError("Нет ни одного пункта: без замечаний таблица бессмысленна "
                         "(ключ «пункты» — список {замечание, меры, листы}).")
    case_no = data.get("дело") or ""
    fio_rod = data.get("фигурант_род") or ""
    kogo = (data.get("кто_замечания") or "").strip()

    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    for m in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, m, Cm(1.5))
    st = doc.styles["Normal"]
    st.font.name = FONT_NAME
    st.font.size = FONT_SIZE_TBL

    # Заголовок
    _title(doc, "ТАБЛИЦА")
    подзаг = "устранения недостатков"
    if kogo:
        подзаг += f" (замечаний {kogo})"
    if case_no:
        подзаг += f" по уголовному делу № {case_no}"
    if fio_rod:
        подзаг += f" в отношении {fio_rod}"
    _title(doc, подзаг)

    # Таблица
    tbl = doc.add_table(rows=1, cols=4)
    tbl.autofit = False
    _borders(tbl)
    hdr = list(HEADERS)
    hdr[1] = hdr[1].format(кого=kogo).replace("  ", " ").strip()
    for i, h in enumerate(hdr):
        c = tbl.rows[0].cells[i]
        _cell(c, h, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, width_cm=COL_WIDTHS_CM[i])
        _shade(c)
    for n, it in enumerate(items, 1):
        row = tbl.add_row()
        _cell(row.cells[0], str(it.get("№", n)), align=WD_ALIGN_PARAGRAPH.CENTER,
              width_cm=COL_WIDTHS_CM[0])
        _cell(row.cells[1], it.get("замечание", ""), width_cm=COL_WIDTHS_CM[1])
        _cell(row.cells[2], it.get("меры", ""), width_cm=COL_WIDTHS_CM[2])
        _cell(row.cells[3], it.get("листы", ""), align=WD_ALIGN_PARAGRAPH.LEFT,
              width_cm=COL_WIDTHS_CM[3])
    _fixed_widths(tbl, COL_WIDTHS_CM)

    # Примечание ПОСЛЕ таблицы
    note = (data.get("примечание") or "").strip()
    if note:
        doc.add_paragraph()
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        _run(p, _normalize_text(note))

    _add_signature(doc, data.get("подпись") or "следователь", constants)

    _backup_if_exists(out_path)
    doc.save(out_path)
    return out_path


DEMO = {
    "дело": "[номер дела]",
    "фигурант_род": "[ФИО в родительном падеже]",
    "кто_замечания": "[должность и ФИО проверяющего]",
    "примечание": "Учебный проект; выводы необходимо сверять с материалами.",
    "пункты": [],
}


def main():
    ap = argparse.ArgumentParser(description="Таблица устранения недостатков (.docx)")
    ap.add_argument("--json", help="файл с данными (UTF-8)")
    ap.add_argument("--out", required=True, help="куда сохранить .docx (папка дела!)")
    ap.add_argument("--demo", action="store_true", help="собрать демонстрационный образец")
    ap.add_argument("--constants", help="путь к константы.json (по умолчанию — авто)")
    a = ap.parse_args()
    if a.demo:
        data = DEMO
    elif a.json:
        with open(a.json, encoding="utf-8") as f:
            data = json.load(f)
    else:
        ap.error("нужен --json или --demo")
    p = build_tablica(a.out, data, a.constants)
    print("Готово:", p)
    print("Проверь: альбомная A4, поля 1,5 см, шапка с серой заливкой, TNR 12, подпись следователя.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
