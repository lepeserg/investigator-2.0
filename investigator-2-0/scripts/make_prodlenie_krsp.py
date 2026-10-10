# -*- coding: utf-8 -*-
"""
Генератор постановления о возбуждении ходатайства о продлении срока проверки
сообщения о преступлении (ч. 3 ст. 144 УПК РФ).

Использует ОБЩИЙ генератор build_doc из make_docx.py (единый формат скилла:
Times New Roman 13; поля и стиль связок — из персонального слоя (константы.json / 02a-format-profile.md: Казань 2,5/2,0/2,0/1,0, УСТАНОВИЛ:/ПОСТАНОВИЛ: без разрядки),
автоматические отбивки вокруг УСТАНОВИЛ:/ПОСТАНОВИЛ: и после места/даты,
красная строка). Отдельной копии форматирования здесь больше нет —
любые правки формата в make_docx.py применяются и к этому жанру.

Для нового продления заполнить словарь DATA (≈7 полей) и запустить:

    python3 make_prodlenie_krsp.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from make_docx import build_doc

# Реквизиты органа берутся из персонального слоя (константы.json / 01-identity.md).
# Здесь — плейсхолдеры (родительный/именительный падеж); подставь реальное
# наименование органа при вызове генератора.
ORG = "{ОРГАН_ПОЛН_РОД}"           # напр.: «военного следственного отдела ... по ... гарнизону»
ORG_IMENIT = "{ОРГАН_ПОЛН_ИМЕНИТ}"  # напр.: «военного следственного отдела ... по ... гарнизону» (им. п.)


def build_prodlenie(out_path, d):
    """Собирает постановление о продлении срока проверки КРСП через build_doc."""
    krsp = f"(КРСП № {d['krsp_no']} от {d['krsp_date']})"
    blocks = []

    # ── гриф-решение руководителя (сверху, по центру) ──
    grif = [
        "Срок проверки сообщения о преступлении",
        f"продлить до {d['days_word']} суток, то есть",
        f"до {d['date_to']}.",
        "",
        f"Руководитель {ORG_IMENIT}",
        f"{d['leader_rank']}            {d['leader_name']}",
        "",
        f"«{d['day2']}» {d['monthyear']}",
    ]
    # Гриф-решение руководителя — по центру, но со смещением ВПРАВО (правка от
    # 06.06.2026, действующий стандарт формат-профиля): блок решения/подписи
    # руководителя — левый отступ 6,75 см, строка с датой решения «__» — 1,5 см.
    date_line = f"«{d['day2']}» {d['monthyear']}"
    for line in grif:
        if line == "":
            blocks.append({"type": "blank"})
        elif line == date_line:
            blocks.append({"type": "grif_line", "text": line, "indent_cm": 1.5})
        else:
            blocks.append({"type": "grif_line", "text": line, "indent_cm": 6.75})
    blocks.append({"type": "blank"})

    # ── заголовок ──
    blocks.append({"type": "header_centered", "text": "ПОСТАНОВЛЕНИЕ"})
    blocks.append({"type": "header_centered_thin", "text": "о возбуждении ходатайства о продлении срока"})
    blocks.append({"type": "header_centered_thin", "text": "проверки сообщения о преступлении"})
    blocks.append({"type": "blank"})
    blocks.append({"type": "place_date", "place": "{ГОРОД}", "date": d["date_full"]})

    # ── преамбула + УСТАНОВИЛ ──
    blocks.append({"type": "para", "text": (
        f"Старший следователь-криминалист {ORG} {d['inv_rank']} {d['inv_name']}, "
        f"рассмотрев сообщение о преступлении, предусмотренном {d['statya']}, признаки которого "
        f"усматриваются в действиях военнослужащего в/части {d['vch']} {d['fio']} {krsp}, –")})
    blocks.append({"type": "section", "text": "УСТАНОВИЛ:"})
    for para in d["fabula"]:
        blocks.append({"type": "para", "text": para})
    blocks.append({"type": "para", "text": (
        "Для принятия законного и обоснованного решения, в целях проведения полной, объективной "
        "и всесторонней проверки, необходимо:")})
    for m in d["meropriyatiya"]:
        blocks.append({"type": "para", "text": m})
    blocks.append({"type": "para", "text": (
        "Принимая во внимание, что выполнить указанные мероприятия в установленный трёхсуточный срок "
        "не представляется возможным, необходимо продлить срок проверки сообщения о преступлении "
        f"до {d['days_num']} суток, то есть до {d['date_to']}.")})
    blocks.append({"type": "para", "text": "На основании изложенного и руководствуясь ч. 3 ст. 144 УПК РФ,"})

    # ── ПОСТАНОВИЛ ──
    blocks.append({"type": "section", "text": "ПОСТАНОВИЛ:"})
    blocks.append({"type": "para", "text": (
        f"Ходатайствовать перед руководителем {ORG} о продлении срока проверки сообщения о преступлении, "
        f"предусмотренном {d['statya']}, признаки которого усматриваются в действиях военнослужащего "
        f"в/части {d['vch']} {d['fio']} {krsp}, до {d['days_word']} суток, то есть до {d['date_to']}.")})

    # ── подпись следователя ──
    blocks.append({"type": "signature", "kind": "следователь"})

    return build_doc(out_path, blocks)


# ════════════════ ЗАПОЛНИТЬ ТОЛЬКО ЭТО (пример условный, вымышленные данные) ════════════════
DATA = {
    "krsp_no": "[номер КРСП]",
    "krsp_date": "ДД.ММ.ГГГГ",
    "statya": "[квалификация по проверенным материалам]",
    "vch": "в/ч 00000",
    "fio": "[ФИО лица]",
    "date_full": "ДД.ММ.ГГГГ",
    "day2": "ДД",
    "monthyear": "[месяц ГГГГ года]",
    "date_to": "ДД.ММ.ГГГГ",
    "days_word": "[срок прописью]",
    "days_num": "[число дней]",
    "fabula": ["[условное обстоятельство 1]", "[условное обстоятельство 2]"],
    "meropriyatiya": ["[проверочное мероприятие 1]", "[проверочное мероприятие 2]", "[проверочное мероприятие 3]"],
    "inv_rank": "[звание следователя]",
    "inv_name": "[ФИО следователя]",
    "leader_rank": "[звание руководителя]",
    "leader_name": "[ФИО руководителя]",
}
# ══════════════════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    # ⚠ было "/tmp/prodlenie_krsp.docx" — POSIX-путь, которого на Windows НЕТ:
    # запуск без аргумента падал FileNotFoundError на машине владельца и работал только в Cowork.
    # gettempdir() даёт /tmp в песочнице и %TEMP% на Windows.
    # Разбор аргументов ДО записи на диск: «--help»/опечатка не должны создавать
    # файл-мусор с текстом постановления (ревизия 22.08.2026).
    _arg = sys.argv[1] if len(sys.argv) > 1 else None
    if _arg in ("-h", "--help", "/?", "help"):
        print(__doc__)
        print("    python make_prodlenie_krsp.py <путь к .docx>   # положить в указанный файл")
        sys.exit(0)
    if _arg is not None and _arg.startswith("-"):
        sys.exit("Неизвестный аргумент: %s\nСправка: python make_prodlenie_krsp.py --help" % _arg)

    out = _arg if _arg else os.path.join(tempfile.gettempdir(), "prodlenie_krsp.docx")
    build_prodlenie(out, DATA)
    print("OK:", out)
