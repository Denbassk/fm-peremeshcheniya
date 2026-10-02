# -*- coding: utf-8 -*-
"""
Перемещение сигарет по торговым точкам: раскладка готового заказа в файлы ТСД.

Вход : xlsx с колонками "Название товара", "Штрих-код", "Центр учета", "ЗАКАЗ"
Выход: по папке на каждый адрес, внутри xlsx (для человека) и txt (для ТСД)
Формат txt: штрих-код;количество  (cp1251, CRLF)

Коэффициентов и пересчёта нет: количество берётся из файла как есть.
Пути не прописаны - файл и папка выбираются через проводник.
"""

import os
import re
import sys
import math
from datetime import date, datetime

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

APP_TITLE = "Перемещение сигарет -> ТСД"

TXT_ENCODING = "cp1251"
TXT_NEWLINE = "\r\n"

BC_STD_LEN = (8, 12, 13, 14)
BC_MANUAL = {
    # "54881005906": "054881005906",
}

ALIASES = {
    "name": ["название товара", "назва товару", "наименование", "товар"],
    "bc":   ["штрих-код", "штрих код", "штрихкод", "штрих-код товару"],
    "shop": ["центр учета", "центр обліку", "магазин", "адрес", "адреса", "точка"],
    "qty":  ["заказ", "замовлення", "количество", "кількість", "кол-во", "к-во"],
}

LOG = []


def say(msg):
    print(msg)
    LOG.append(str(msg))

# ------------------------- утилиты -------------------------

def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def safe_name(s):
    s = str(s).strip().replace("/", "-").replace("\\", "-")
    s = re.sub(r'[:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s or "БЕЗ_АДРЕСА"


def fmt_barcode(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if math.isnan(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return str(v).strip()
    if isinstance(v, int):
        return str(v)
    s = str(v).strip().replace(" ", "").replace("\u00a0", "")
    if s.lower().endswith(".0"):
        s = s[:-2]
    return s


def ean_valid(bc):
    if not bc.isdigit() or len(bc) not in BC_STD_LEN:
        return False
    d = [int(ch) for ch in bc]
    body, control = d[:-1][::-1], d[-1]
    total = sum(x * (3 if i % 2 == 0 else 1) for i, x in enumerate(body))
    return (10 - total % 10) % 10 == control


def restore_barcode(bc):
    """Восстанавливает ведущие нули, потерянные Excel. -> (код, пояснение)"""
    if not bc or not bc.isdigit():
        return bc, ""
    if bc in BC_MANUAL:
        return BC_MANUAL[bc], "%s -> %s (вручную)" % (bc, BC_MANUAL[bc])
    if len(bc) in BC_STD_LEN and ean_valid(bc):
        return bc, ""
    for L in BC_STD_LEN:
        if L <= len(bc):
            continue
        cand = bc.zfill(L)
        if ean_valid(cand):
            return cand, "%s -> %s" % (bc, cand)
    return bc, ""


def to_qty(v):
    """Количество -> целое число штук. None, если не число или <= 0."""
    if v is None:
        return None
    s = str(v).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    if s == "":
        return None
    try:
        x = float(s)
    except ValueError:
        return None
    if x <= 0:
        return 0
    return int(round(x))


def extract_date(path):
    m = re.search(r"(\d{1,2})[.\-_ ](\d{1,2})[.\-_ ](\d{2,4})", os.path.basename(path))
    if m:
        d, mo, y = (int(g) for g in m.groups())
        y = y + 2000 if y < 100 else y
        try:
            return date(y, mo, d)
        except ValueError:
            pass
    return date.today()

# ---------------------- чтение источника ----------------------

def detect_header(ws, max_scan=15):
    """-> (номер строки заголовков, {ключ: номер столбца})"""
    best = None
    for r in range(1, min(max_scan, ws.max_row) + 1):
        vals = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                vals[c] = norm(v)
        if not vals:
            continue
        cols = {}
        for key, variants in ALIASES.items():
            hit = None
            for c, txt in vals.items():
                if txt in variants:
                    hit = c
                    break
            if hit is None:
                for c, txt in vals.items():
                    if any(v in txt for v in variants):
                        hit = c
                        break
            if hit:
                cols[key] = hit
        score = len(cols)
        if best is None or score > best[0]:
            best = (score, r, cols)
        if score == 4:
            break
    if best is None:
        return None, {}
    return best[1], best[2]


def read_source(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    say("Лист: %s   (строк в файле: %s)" % (ws.title, ws.max_row))

    hr, cols = detect_header(ws)
    lack = [k for k in ("bc", "shop", "qty") if k not in cols]
    if lack:
        wb.close()
        raise ValueError(
            "В файле не найдены обязательные столбцы: %s.\n"
            "Ожидаются заголовки: Название товара, Штрих-код, Центр учета, ЗАКАЗ."
            % ", ".join(lack))

    say("Строка заголовков: %d" % hr)
    say("Столбцы: штрих-код=%s, адрес=%s, заказ=%s, название=%s"
        % (get_column_letter(cols["bc"]), get_column_letter(cols["shop"]),
           get_column_letter(cols["qty"]),
           get_column_letter(cols["name"]) if "name" in cols else "нет"))

    rows, problems, fixes = [], [], []
    skipped_empty = skipped_zero = 0
    for r, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), start=hr + 1):
        get = lambda k: (row[cols[k] - 1] if k in cols and cols[k] - 1 < len(row) else None)
        bc = fmt_barcode(get("bc"))
        shop = str(get("shop") or "").strip()
        name = str(get("name") or "").strip()
        raw_q = get("qty")

        if not bc and not shop and raw_q in (None, ""):
            skipped_empty += 1
            continue
        if not bc or not shop or shop.lower() == "none":
            problems.append(("Нет штрих-кода или адреса", shop, bc, name,
                             "строка %d" % r))
            continue
        if not bc.isdigit():
            problems.append(("Штрих-код не только из цифр", shop, bc, name,
                             "строка %d" % r))
            continue

        nb, note = restore_barcode(bc)
        if note:
            fixes.append(note)
            bc = nb

        q = to_qty(raw_q)
        if q is None:
            problems.append(("Количество не число", shop, bc, name,
                             "строка %d, значение '%s'" % (r, raw_q)))
            continue
        if q <= 0:
            skipped_zero += 1
            continue

        rows.append({"shop": shop, "bc": bc, "name": name, "qty": q, "src_row": r})

    wb.close()
    say("Прочитано позиций к перемещению: %d" % len(rows))
    if skipped_empty:
        say("Пропущено пустых строк: %d" % skipped_empty)
    if skipped_zero:
        say("Пропущено строк с нулевым количеством: %d" % skipped_zero)
    if fixes:
        uniq = sorted(set(fixes))
        say("ВНИМАНИЕ: восстановлены ведущие нули (%d): %s"
            % (len(uniq), "; ".join(uniq[:8])))
    else:
        say("Ведущие нули штрих-кодов: потерь не обнаружено")
    for p in problems:
        say("ПРОБЛЕМА: %s | %s | %s | %s" % (p[0], p[1], p[2], p[4]))
    return rows, problems, sorted(set(fixes))


def aggregate(rows):
    """Свод дублей одного кода внутри адреса."""
    acc, order, dups = {}, [], 0
    for it in rows:
        k = (it["shop"], it["bc"])
        if k in acc:
            acc[k]["qty"] += it["qty"]
            dups += 1
        else:
            acc[k] = dict(it)
            order.append(k)
    if dups:
        say("Сведено дублей (один код дважды в одном адресе): %d" % dups)
    else:
        say("Дублей внутри адреса нет")
    by_shop = {}
    for k in order:
        by_shop.setdefault(k[0], []).append(acc[k])
    for shop in by_shop:
        by_shop[shop].sort(key=lambda x: norm(x["name"]))
    return by_shop

# --------------------------- запись ---------------------------

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
ERR_FILL = PatternFill("solid", fgColor="FCE4E4")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")

HDRS = ["№", "Название товара", "Штрих-код", "Количество"]
WIDTHS = [5, 46, 18, 13]


def write_shop_xlsx(path, shop, day, items):
    wb = Workbook()
    ws = wb.active
    ws.title = "Перемещение"
    ws["A1"] = "Перемещение сигарет: %s" % shop
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("Дата: %s     Позиций: %d     Всего пачек: %d"
                % (day.strftime("%d.%m.%Y"), len(items),
                   sum(i["qty"] for i in items)))
    ws["A2"].font = Font(italic=True, size=9)

    hr = 4
    for j, h in enumerate(HDRS, start=1):
        c = ws.cell(row=hr, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r = hr
    for n, it in enumerate(items, start=1):
        r += 1
        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=2, value=it["name"])
        bc = ws.cell(row=r, column=3, value=it["bc"])
        bc.number_format = "@"
        bc.alignment = Alignment(horizontal="center")
        q = ws.cell(row=r, column=4, value=it["qty"])
        q.font = Font(bold=True)
        q.alignment = Alignment(horizontal="center")
        for j in range(1, len(HDRS) + 1):
            ws.cell(row=r, column=j).border = BORDER

    last = r
    r += 1
    ws.cell(row=r, column=2, value="ИТОГО:").font = Font(bold=True)
    ws.cell(row=r, column=2).alignment = Alignment(horizontal="right")
    tot = ws.cell(row=r, column=4, value=sum(i["qty"] for i in items))
    tot.font = Font(bold=True)
    tot.fill = WARN_FILL

    for j, wd in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.freeze_panes = "A%d" % (hr + 1)
    if last > hr:
        ws.auto_filter.ref = "A%d:%s%d" % (hr, get_column_letter(len(HDRS)), last)
    wb.save(path)


def write_shop_txt(path, items):
    lines = ["%s;%d" % (i["bc"], i["qty"]) for i in items]
    with open(path, "w", encoding=TXT_ENCODING, errors="replace", newline="") as f:
        f.write(TXT_NEWLINE.join(lines) + TXT_NEWLINE)


TXT_LINE_RE = re.compile(r"^\d{6,14};\d+$")


def verify_written(shop_dir, nm, items):
    problems = []
    xlsx_path = os.path.join(shop_dir, nm + ".xlsx")
    txt_path = os.path.join(shop_dir, nm + ".txt")

    expected = {}
    for i in items:
        expected[i["bc"]] = expected.get(i["bc"], 0) + i["qty"]

    wb = load_workbook(xlsx_path, data_only=False)
    ws = wb.active
    got = {}
    r = 5
    while True:
        v3 = ws.cell(row=r, column=3).value
        if v3 in (None, ""):
            break
        b = fmt_barcode(v3)
        if not b.isdigit():          # строка ИТОГО или иной служебный текст
            break
        v = ws.cell(row=r, column=4).value
        got[b] = got.get(b, 0) + (int(v) if isinstance(v, (int, float)) else 0)
        r += 1
    wb.close()
    if got != expected:
        problems.append("xlsx: позиций %d вместо %d или количества не совпали"
                        % (len(got), len(expected)))

    with open(txt_path, "r", encoding=TXT_ENCODING) as f:
        lines = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    if len(lines) != len(expected):
        problems.append("txt: строк %d вместо %d" % (len(lines), len(expected)))
    tgot = {}
    for ln in lines:
        if not TXT_LINE_RE.match(ln):
            problems.append("txt: неверный формат строки '%s'" % ln)
            continue
        b, q = ln.split(";")
        tgot[b] = tgot.get(b, 0) + int(q)
    if tgot != expected:
        problems.append("txt: содержимое не совпало с расчётом")
    if os.path.getsize(txt_path) == 0:
        problems.append("txt: файл пустой")
    return problems


def write_summary(path, day, src, summary, problems, fixes):
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = "Перемещение сигарет на %s" % day.strftime("%d.%m.%Y")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Источник: %s   |   создано %s" % (src, datetime.now().strftime("%d.%m.%Y %H:%M"))
    ws["A2"].font = Font(italic=True, size=9)

    for j, h in enumerate(["№", "Адрес (Центр учета)", "Позиций", "Пачек",
                           "Проверка файлов", "Папка"], start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
    for n, it in enumerate(summary, start=1):
        ws.cell(row=4 + n, column=1, value=n)
        ws.cell(row=4 + n, column=2, value=it["shop"])
        ws.cell(row=4 + n, column=3, value=it["positions"])
        ws.cell(row=4 + n, column=4, value=it["units"])
        st = ws.cell(row=4 + n, column=5, value=it["verify"])
        st.fill = OK_FILL if it["verify"] == "OK" else ERR_FILL
        ws.cell(row=4 + n, column=6, value=it["folder"])
    r = 5 + len(summary)
    ws.cell(row=r, column=3, value=sum(i["positions"] for i in summary)).font = Font(bold=True)
    ws.cell(row=r, column=4, value=sum(i["units"] for i in summary)).font = Font(bold=True)
    for j, wd in enumerate([5, 34, 10, 10, 40, 64], start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.freeze_panes = "A5"

    ws2 = wb.create_sheet("Проблемы")
    for j, h in enumerate(["Тип", "Адрес", "Штрих-код", "Товар", "Детали"], start=1):
        c = ws2.cell(row=1, column=j, value=h)
        c.font = Font(bold=True)
        c.fill = HEAD_FILL
    for n, p in enumerate(problems, start=1):
        for j, v in enumerate(p, start=1):
            ws2.cell(row=1 + n, column=j, value=v)
    for j, wd in enumerate([36, 26, 18, 42, 40], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = wd

    ws3 = wb.create_sheet("Ведущие нули")
    ws3.cell(row=1, column=1, value="Исправление штрих-кода").font = Font(bold=True)
    for n, f in enumerate(fixes, start=1):
        ws3.cell(row=1 + n, column=1, value=f)
    ws3.column_dimensions["A"].width = 46
    wb.save(path)

# ---------------------------- диалоги ----------------------------

def ask_paths():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    src = filedialog.askopenfilename(
        title="Выберите файл заказа сигарет (xlsx)",
        filetypes=[("Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")])
    if not src:
        root.destroy()
        return None, None

    out = filedialog.askdirectory(
        title="Куда выгрузить файлы для ТСД",
        initialdir=os.path.dirname(src) or os.path.expanduser("~"))
    if not out:
        root.destroy()
        return None, None

    root.destroy()
    return src, out

# ----------------------------- main -----------------------------

def run(src, out_root):
    say("=" * 70)
    say("ПЕРЕМЕЩЕНИЕ СИГАРЕТ -> ТСД  |  %s" % datetime.now().strftime("%d.%m.%Y %H:%M:%S"))
    say("=" * 70)
    say("Источник: %s" % src)
    say("Выгрузка: %s" % out_root)

    day = extract_date(src)
    say("Дата перемещения: %s" % day.strftime("%d.%m.%Y"))

    rows, problems, fixes = read_source(src)
    if not rows:
        raise ValueError("В файле нет ни одной позиции с количеством больше нуля.")

    by_shop = aggregate(rows)
    total_src = sum(i["qty"] for i in rows)

    day_dir = os.path.join(out_root, "СИГАРЕТЫ_%s" % day.strftime("%Y-%m-%d"))
    os.makedirs(day_dir, exist_ok=True)

    say("-" * 70)
    summary, used, fails = [], {}, 0
    for shop in sorted(by_shop, key=norm):
        items = by_shop[shop]
        base = safe_name(shop)
        nm, k = base, 2
        while nm.lower() in used:
            nm = "%s (%d)" % (base, k)
            k += 1
        used[nm.lower()] = shop

        shop_dir = os.path.join(day_dir, nm)
        os.makedirs(shop_dir, exist_ok=True)
        write_shop_xlsx(os.path.join(shop_dir, nm + ".xlsx"), shop, day, items)
        write_shop_txt(os.path.join(shop_dir, nm + ".txt"), items)

        probs = verify_written(shop_dir, nm, items)
        if probs:
            fails += 1
            for p in probs:
                problems.append(("Проверка записанных файлов", shop, "", "", p))
        units = sum(i["qty"] for i in items)
        summary.append({"shop": shop, "positions": len(items), "units": units,
                        "folder": shop_dir,
                        "verify": "OK" if not probs else "СБОЙ: " + "; ".join(probs)})
        say("%-28s позиций: %4d   пачек: %6d   %s"
            % (shop, len(items), units, "OK" if not probs else "СБОЙ"))

    say("-" * 70)
    tot_units = sum(s["units"] for s in summary)
    if tot_units != total_src:
        problems.append(("Сверка итогов", "", "", "",
                         "в файлах %d пачек, в источнике %d" % (tot_units, total_src)))
        say("ОШИБКА: сумма по файлам (%d) не равна сумме в источнике (%d)"
            % (tot_units, total_src))
    else:
        say("Сверка итогов: OK, всего пачек %d" % tot_units)
    if fails:
        say("ОШИБКА: обратная сверка не прошла в %d адресах" % fails)
    else:
        say("Обратная сверка xlsx и txt: OK (%d адресов)" % len(summary))

    sum_path = os.path.join(day_dir, "_СВОДКА.xlsx")
    write_summary(sum_path, day, src, summary, problems, fixes)

    say("Адресов        : %d" % len(summary))
    say("Позиций        : %d" % sum(s["positions"] for s in summary))
    say("Пачек          : %d" % tot_units)
    say("Проблем        : %d" % len(problems))
    say("Результат      : %s" % day_dir)
    say("Сводка         : %s" % sum_path)

    log_path = os.path.join(day_dir, "_ЛОГ_%s.txt" % datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(LOG))

    ok = (fails == 0 and tot_units == total_src)
    return ok, day_dir, len(summary), tot_units, len(problems)


def main():
    src, out_root = ask_paths()
    if not src:
        say("Отменено пользователем.")
        return 1
    try:
        ok, day_dir, shops, units, nprob = run(src, out_root)
    except Exception as e:
        import traceback
        traceback.print_exc()
        say("СБОЙ: %s: %s" % (type(e).__name__, e))
        _msg("Ошибка", "Не удалось обработать файл.\n\n%s: %s" % (type(e).__name__, e), True)
        return 2

    text = ("Готово.\n\nАдресов: %d\nВсего пачек: %d\nПроблем: %d\n\nПапка:\n%s"
            % (shops, units, nprob, day_dir))
    _msg("Готово" if ok else "Готово с замечаниями", text, not ok)
    try:
        os.startfile(day_dir)
    except Exception:
        pass
    return 0 if ok else 2


def _msg(title, text, warn=False):
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    (messagebox.showwarning if warn else messagebox.showinfo)("%s - %s" % (APP_TITLE, title), text)
    root.destroy()


if __name__ == "__main__":
    code = main()
    sys.exit(code)
