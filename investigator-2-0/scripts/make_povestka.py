# -*- coding: utf-8 -*-
"""Сборщик повестки о вызове на допрос (ст. 188 УПК РФ) в формате ВСО.

Жанр: повестка свидетелю / потерпевшему / подозреваемому / обвиняемому,
вызываемому НАПРЯМУЮ (гражданское лицо). Военнослужащий вызывается через
командование воинской части (ч. 7 ст. 188 УПК РФ) — это отдельный документ
(вызов через командира), здесь не собирается.

Норма сверена 23.06.2026 (consultant.ru):
  ч. 1 ст. 188 — повестка указывает: кто и в каком качестве, к кому и куда,
                 дата и время, последствия неявки без уважительных причин;
  ч. 3 ст. 188 — обязанность явиться либо заранее уведомить о причинах неявки;
                 при неявке без уважительных причин — привод либо иные меры
                 принуждения по ст. 111 УПК РФ;
  привод свидетеля — ч. 7 ст. 56 и ст. 113 УПК РФ.

Род («Гр-ну/Гр-ке», «проживающему/проживающей», «получил/получила») определяется
по отчеству в дательном падеже; если отчество нестандартное — задать sex явно,
иначе сборщик остановится с ошибкой.

Жёлтой заливкой автоматически помечаются поля, оставшиеся под заполнение:
№ дела с хвостом «___», кабинет «___», телефон «___»."""

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_docx import build_doc


_MALE_TAILS = ("вичу", "ичу", "оглы", "улы")
_FEM_TAILS = ("вне", "ичне", "нишне", "кызы")


def detect_sex(fio_dat):
    """Определить пол по отчеству в ДАТЕЛЬНОМ падеже: «…овичу» → м, «…овне» → ж.

    Возвращает "м" / "ж" / None (отчества нет или оно нестандартное — пол задаёт
    вызывающий явно параметром sex).
    """
    words = fio_dat.strip().strip(",.").lower().split()
    if not words:
        return None
    last = words[-1]
    if last.endswith(_MALE_TAILS):
        return "м"
    if last.endswith(_FEM_TAILS):
        return "ж"
    return None


def build_povestka(out_path, *, fio_dat, fio_short, address, delo_num, qualifier,
                   date_yavka, time_yavka, date_doc, sex=None, kabinet="___",
                   phone="{ТЕЛ_ИСПОЛНИТЕЛЯ}", place="{ГОРОД}",
                   org_rod="{ОРГАН_ПОЛН_РОД}",
                   inv_rank="{СЛЕДОВАТЕЛЬ_ЗВАНИЕ}", inv_name="{СЛЕДОВАТЕЛЬ_ФИО}",
                   office_address="{АДРЕС_ОРГАНА}", email="{EMAIL_ОРГАНА}"):
    sex = (sex or detect_sex(fio_dat) or "").lower()[:1]
    if sex not in ("м", "ж"):
        raise ValueError(
            "build_povestka: не определён пол вызываемого — отчество в %r нестандартное. "
            "Задать явно: sex=\"м\" либо sex=\"ж\" (от этого зависят «Гр-ну/Гр-ке», "
            "«проживающему/проживающей», «получил/получила»)." % fio_dat)
    grazhd = "Гр-ну " if sex == "м" else "Гр-ке "
    prozhiv = "проживающему" if sex == "м" else "проживающей"
    poluchil = "получил" if sex == "м" else "получила"
    year_doc = re.search("(19|20)[0-9]{2}", date_doc)
    year_doc = year_doc.group(0) if year_doc else "20__"

    blocks = [
        {"type": "header_centered", "text": "ПОВЕСТКА"},
        {"type": "header_centered_thin", "text": "о вызове на допрос"},
        {"type": "blank"},
        {"type": "place_date", "place": place, "date": date_doc},

        {"type": "para_no_indent", "text": grazhd + fio_dat + ","},
        {"type": "para_no_indent", "text": prozhiv + " по адресу: " + address + "."},
        {"type": "blank"},

        {"type": "para", "text": "На основании ст. 188 УПК РФ Вы вызываетесь на допрос "
            "в качестве " + qualifier + " по уголовному делу № " + delo_num + " к следователю "
            + org_rod + " " + inv_rank + " " + inv_name + "."},

        {"type": "para", "text": "Явиться надлежит " + date_yavka + " к " + time_yavka +
            " по адресу: " + office_address + ", служебный кабинет № " + kabinet + "."},

        {"type": "para", "text": "При себе иметь документ, удостоверяющий личность (паспорт)."},

        {"type": "para", "text": "В случае неявки без уважительных причин Вы можете быть "
            "подвергнуты приводу либо к Вам могут быть применены иные меры процессуального "
            "принуждения, предусмотренные ст. 111 УПК РФ (ч. 3 ст. 188 УПК РФ). При "
            "невозможности явиться в назначенный срок необходимо заранее уведомить "
            "следователя о причинах неявки по телефону " + phone + " или по адресу "
            "электронной почты: " + email + "."},

        {"type": "signature", "kind": "следователь"},

        {"type": "blank"},
        {"type": "raw", "text": "– – – – – – – – – – – – – – – – – – – – линия отреза "
            "– – – – – – – – – – – – – – – – – – – –"},
        {"type": "blank"},
        {"type": "header_centered", "text": "РАСПИСКА"},
        {"type": "para_no_indent", "text": "Повестку о вызове на допрос в качестве " +
            qualifier + " по уголовному делу № " + delo_num + " на " + date_yavka +
            " к " + time_yavka + " " + poluchil + "."},
        {"type": "blank"},
        {"type": "sig_line", "role": "«___» __________ " + year_doc + " года",
         "name": fio_short},
    ]

    # Жёлтая заливка полей, оставшихся под заполнение.
    for b in blocks:
        if b.get("type") in ("para", "para_no_indent"):
            txt = b["text"]
            if "000___" in txt or "№ ___" in txt or "телефону ___" in txt:
                b["highlight"] = True

    return build_doc(out_path, blocks)
