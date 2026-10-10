# -*- coding: utf-8 -*-
"""post_release_gvp.py — ПОСТОБРАБОТКА статкарт ГВП, выпущенных кнопкой «КнопкаВыпуска».

  1. ⛔ У Ф-1 ТЕРЯЮТСЯ графы 2.8.а (квалификация) и 2.8.б (квалифицирующие признаки) —
     они заполняются из справочных списков формы, и при выпуске обнуляются. Заодно
     пустеют служебные O47 (код статьи для DBF) и O48 (номер дела без разделителей).
     Карта уходит прокурору БЕЗ КВАЛИФИКАЦИИ.
  2. Лист выпуска защищён (SHA-512), из-за чего не поправить вёрстку.
  3. Дата в графе 2.19.а рендерится как `######` — не влезает в колонку.
  4. Графа 2.7 (описание события) в форме без переноса по словам: длинная фабула
     уходит одной строкой за границы ячейки.

  Ф-1.2 выпускается без потерь — ей нужен только п. 2 (и то по желанию).

Что делает скрипт:
  · сверяет выпуск с формой-источником по карте ячеек и ВОЗВРАЩАЕТ потерянные значения;
  · снимает `<sheetProtection>` из XML выпуска (это удаление тега в собственном файле,
    НЕ подбор пароля — хеш не трогается);
  · включает в 2.7 перенос по словам, выравнивание по верху и кегль 8, восстанавливая
    объединение D40:I40, если оно разрушено;
  · ставит в 2.19.а короткий формат даты `ДД.ММ.ГГ`, чтобы ушли решётки;
  · кладёт рядом `.bak_<timestamp>` перед первой правкой.

Использование:
    python post_release_gvp.py \"<…\\дело\\02 Статкарточки>\"
    python post_release_gvp.py \"<…>\" --no-unprotect     # защиту не снимать

Требуется Windows + Excel + pywin32. Осиротевший EXCEL.EXE скрипт прибирает сам.
См. references/09-genres-stat-cards.md §09.18.5."""
import argparse
import glob
import os
import re
import shutil
import subprocess
import sys
import time
import zipfile

sys.stdout.reconfigure(encoding="utf-8")

# ── карта ячеек: что сверять между формой и выпуском ────────────────────────────
CELLS_F1 = [
    "I5", "F7", "F9", "H10", "I11", "F13", "H13", "H14", "I16", "I19", "I24",
    "F26", "H26", "I27", "I30", "I31", "I33", "H35", "H37", "G39", "H39",
    "D40", "D43", "D44", "I45", "J49", "J50", "J51", "J52", "J53", "J54",
    "H69", "I69", "J71", "J74", "J77", "F79", "J80", "J83", "J85",
    "D89", "E89", "F89", "J89", "D91", "E91", "J91", "D93", "E93", "J93",
]
CELLS_F12 = [
    "I5", "F7", "F9", "I10", "H11", "B14", "D14", "G14", "B17", "E17", "G19",
    "G20", "I21", "E22", "E23", "E24", "E25", "E26", "E27", "I28", "I29",
    "I34", "E36", "E37", "I39", "I40", "E45", "F45", "I45", "D51",
    "D66", "E66", "F66", "J66", "D68", "E68", "J68", "D70", "E70", "J70",
]
SHEET_F1, SHEET_F12 = "Ф1стр1", "Ф1.2стр1"
FAB_CELL, FAB_MERGE, FAB_SIZE = "D40", "D40:I40", 8
DATE_CELLS_F1 = ("J89", "J91")

XL_TOP, XL_JUSTIFY = -4160, -4130


def is_f12(name):
    return "Ф-1.2" in name


def match_form(release_name, forms):
    """Сопоставить выпуск с формой-источником по номеру в имени файла."""
    if is_f12(release_name):
        m = re.search(r"_лицо-(\d+)_", release_name)
        if not m:
            return None
        n = int(m.group(1))
        for f in forms:
            b = os.path.basename(f)
            if is_f12(b) and re.search(r"лицо\s*0*%d\b" % n, b):
                return f
    else:
        m = re.search(r"_(\d+)_Ф-1\.xlsx$", release_name)
        if not m:
            return None
        n = int(m.group(1))
        for f in forms:
            b = os.path.basename(f)
            if not is_f12(b) and re.search(r"\(0*%d\)" % n, b):
                return f
    return None


def strip_protection(path):
    """Удалить <sheetProtection …/> из всех листов книги. Возвращает число снятых блоков."""
    z = zipfile.ZipFile(path)
    items = [(i, z.read(i.filename)) for i in z.infolist()]
    z.close()
    n, out = 0, []
    for info, data in items:
        if info.filename.startswith("xl/worksheets/sheet"):
            txt = data.decode("utf-8")
            txt, k = re.subn(r"<sheetProtection[^>]*/>", "", txt)
            n += k
            data = txt.encode("utf-8")
        out.append((info, data))
    if n:
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as w:
            for info, data in out:
                w.writestr(info, data)
    return n


def backup(path):
    dst = "%s.bak_%s" % (path, time.strftime("%Y%m%d_%H%M%S"))
    shutil.copyfile(path, dst)
    return os.path.basename(dst)


def _own_pid(app):
    """PID НАШЕГО экземпляра Excel — по его окну (Application.Hwnd).

    None, если узнать не удалось (нет pywin32/Hwnd) — тогда ничего не убиваем."""
    try:
        import win32process
        pid = win32process.GetWindowThreadProcessId(int(app.Hwnd))[1]
        return pid if pid > 0 else None
    except Exception:
        return None


def kill_excel(pid):
    """Снять зависший СВОЙ EXCEL.EXE — только по PID, и никогда чужой.

    ⛔ Прежняя версия делала `taskkill /F /IM EXCEL.EXE /T` — убивала ВСЕ Excel владельца
    с несохранёнными книгами, хотя сам экземпляр поднимался через DispatchEx именно чтобы
    рабочий Office не трогать. Теперь: PID неизвестен — не убиваем ничего; процесс уже
    вышел после Quit() или PID занят не Excel'ем — тоже не трогаем.
    ⚠ БЕЗ text=True: tasklist пишет в консольной кодировке (cp866)."""
    if not pid:
        return False

    def alive():
        try:
            out = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/FO", "CSV", "/NH"],
                                 capture_output=True, timeout=30).stdout.decode("cp866", "replace")
        except Exception:
            return False                  # не смогли проверить — в сомнении не убиваем
        return bool(re.search(r'"EXCEL\.EXE","%d"' % pid, out, re.I))

    for _ in range(5):                    # дать Excel до ~5 с завершиться самому после Quit()
        if not alive():
            return False                  # вышел сам (или PID уже не Excel) — не трогаем
        time.sleep(1)
    try:
        r = subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
        return r.returncode == 0
    except Exception:
        return False


def main():
    ap = argparse.ArgumentParser(description="Постобработка выпущенных статкарт ГВП")
    ap.add_argument("folder", help="папка «02 Статкарточки» дела")
    ap.add_argument("--no-unprotect", action="store_true", help="не снимать защиту листа")
    args = ap.parse_args()

    base = os.path.abspath(args.folder)
    rel_dir = os.path.join(base, "готовые статкарты")
    if not os.path.isdir(rel_dir):
        sys.exit("нет подпапки «готовые статкарты» в %s" % base)

    forms = sorted(glob.glob(os.path.join(base, "*.xlsm")))
    releases = sorted(f for f in glob.glob(os.path.join(rel_dir, "*.xlsx"))
                      if not os.path.basename(f).startswith("~$"))
    if not releases:
        sys.exit("выпущенных карт не найдено")

    import pythoncom
    import win32com.client.gencache as _gc
    _gc.GetClassForCLSID = lambda clsid: None          # обойти битый gen_py
    import win32com.client as win32

    pythoncom.CoInitialize()
    # DispatchEx — ОТДЕЛЬНЫЙ процесс Excel, чтобы не цепляться к рабочему Office владельца
    xl = win32.DispatchEx("Excel.Application")
    xl_pid = _own_pid(xl)            # PID СВОЕГО процесса — после Quit() Hwnd уже не спросить
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 3        # значения пишем — макросы не нужны
    xl.EnableEvents = False
    xl.Interactive = False
    xl.ScreenUpdating = False
    problems = 0
    try:
        for rel in releases:
            name = os.path.basename(rel)
            form = match_form(name, forms)
            print("==", name)
            if not form:
                print("   ⚠ форма-источник не найдена — пропуск")
                problems += 1
                continue
            print("   источник:", os.path.basename(form))
            print("   бэкап:", backup(rel))

            if not args.no_unprotect:
                print("   снято блоков защиты:", strip_protection(rel))

            f12 = is_f12(name)
            cells = CELLS_F12 if f12 else CELLS_F1
            sheet = SHEET_F12 if f12 else SHEET_F1

            src = xl.Workbooks.Open(form, UpdateLinks=0)
            ss = src.Worksheets(sheet)
            wb = xl.Workbooks.Open(rel, UpdateLinks=0)
            ws = wb.Worksheets(1)

            blank = lambda v: v is None or str(v).strip() == ""
            restored = []
            for a in cells:
                v1, v2 = ss.Range(a).Value, ws.Range(a).Value
                if blank(v1) and blank(v2):
                    continue                      # оба пусты — не расхождение
                if str(v1) != str(v2) and not blank(v1):
                    try:
                        ws.Range(a).Value = v1
                        restored.append("%s=%r" % (a, str(v1)[:22]))
                    except Exception as e:
                        print("   ⚠ %s не записалось: %s" % (a, str(e)[:50]))
                        problems += 1
            print("   возвращено потерянных значений: %d %s"
                  % (len(restored), ("| " + "; ".join(restored)) if restored else ""))

            # разбивка на страницы: Ф-1.2 приходит с «вписать в 1 страницу по высоте»
            # и печатается одним длинным листом — снимаем, как сделано у Ф-1
            ps = ws.PageSetup
            if ps.FitToPagesTall not in (False, 0):
                ps.FitToPagesWide = 1
                ps.FitToPagesTall = False
                print("   печать: снято «вписать в 1 страницу по высоте» — карта делится на листы")

            if not f12:
                c = ws.Range(FAB_CELL)
                if c.MergeArea.Address.replace("$", "") != FAB_MERGE:
                    txt = c.Value
                    ws.Range(FAB_MERGE).Merge()
                    ws.Range(FAB_CELL).Value = txt
                    print("   восстановлено объединение", FAB_MERGE)
                c = ws.Range(FAB_CELL)
                c.WrapText = True
                c.VerticalAlignment = XL_TOP
                c.HorizontalAlignment = XL_JUSTIFY
                c.Font.Size = FAB_SIZE
                c.Font.Bold = False
                for a in DATE_CELLS_F1:
                    if "#" in str(ws.Range(a).Text):
                        for fmt in ("ДД.ММ.ГГ", "dd.mm.yy"):
                            try:
                                ws.Range(a).NumberFormat = fmt
                                break
                            except Exception:
                                continue
                        if "#" in str(ws.Range(a).Text):
                            print("   ⚠ %s всё ещё «решётки»" % a)
                            problems += 1
                print("   2.7: перенос, кегль %d; даты 2.19: %s"
                      % (FAB_SIZE, ", ".join("%s=%s" % (a, ws.Range(a).Text) for a in DATE_CELLS_F1)))

            wb.Save()
            wb.Close(SaveChanges=False)
            src.Close(SaveChanges=False)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass
        del xl
        kill_excel(xl_pid)           # только свой зависший процесс; чужие Excel не трогаем

    print("\nитог: обработано %d карт(ы), проблем: %d" % (len(releases), problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
