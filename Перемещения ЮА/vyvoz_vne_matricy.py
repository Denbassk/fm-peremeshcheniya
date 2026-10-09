# -*- coding: utf-8 -*-
"""
Вывоз из магазинов товара вне действующей ассортиментной матрицы (переход на ЮА Маркет).

Задача: магазины переходят на новое юрлицо (ЮА Маркет), через него заказывается только
действующая матрица. Товар вне матрицы остаётся на полках и занимает место. Скрипт берёт
остатки магазина, сверяет с матрицей и с тем, что заведено на ЮА, и делает списки:

  ВЫВЕЗТИ     нет в матрице, нет на ЮА, не в исключениях -> xlsx + txt для ТСД,
              везём на склад Полевая-Склад (Family);
  ПРОВЕРИТЬ 1 нет в матрице, но заведён на ЮА (по ШК или по точному названию),
              либо в матрице под другим ШК (по названию) -> не вывозим;
  ПРОВЕРИТЬ 2 в матрице, но на ЮА не заведён (ещё не завезли / придёт прямой поставкой);
  ИСКЛЮЧЕНО   розлив, овощи на развес, кеги, кофе из аппарата, пакеты, стаканы, крышки,
              сырьё кофеаппарата -> с полки не снимаем;
  РАЗНЫЕ ШК   один товар под разными ШК (по точному названию) - для разбора.
Обратно в магазины скрипт ничего не раздаёт.

ВХОДЫ (папка ВХОД_ВЫВОЗ рядом со скриптом, выгрузки из Торгсофта):
  МАГАЗИНЫ\\       «Состояние склада» каждого магазина (только в наличии), база Family
  СКЛАД_ЮА\\       «Состояние склада Полевая-Склад ЮА» (все строки, в том числе с нулём)
  МАГАЗИНЫ_ЮА\\    не обязательно: «Состояние склада» магазина в базе ЮА (после перехода)
  ПРИХОДЫ_ЮА\\     не обязательно: «Приходы» ЮА (Движение товара). Берутся ВСЕ приходы с
                  количеством > 0, в том числе от инвентаризации: для вопроса «заведён ли
                  товар на ЮА» это нужно (загрузчик BigQuery их отбрасывает намеренно).
Из BigQuery (с кэшем в Справочник\\): матрица assortment_matrix_full, barcode_recode_map,
поставщик «каваапарат» (сырьё кофеаппарата), история поставок за 120 дней.
РЕЗУЛЬТАТ: ВЫВОЗ\\<дата>\\<магазин>\\<магазин>.xlsx/.txt, _СВОДКА_ВЫВОЗ.xlsx.
Запуск: окно (по умолчанию) или  vyvoz_vne_matricy.py --shops "А;Б" | --auto | --selftest
"""

import os
import re
import sys
import json
import time
import shutil
import difflib
import tempfile
import threading
os.environ.setdefault("GRPC_VERBOSITY", "NONE")
os.environ.setdefault("GLOG_minloglevel", "3")
from datetime import date, datetime, timezone

RZ_DIR_NAME = u"РАСЧЕТ ЗАКАЗА"   # соседняя папка: raschet_zakaza.py и credentials (скрипт лежит в «Перемещения ЮА»)
if not getattr(sys, "frozen", False):
    try:
        _here = os.path.dirname(os.path.abspath(__file__))
        sys.path.insert(0, _here)
        sys.path.insert(1, os.path.join(os.path.dirname(_here), RZ_DIR_NAME))
    except NameError:
        pass

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

import raschet_zakaza as rz
R = rz.R
# ключ BigQuery остаётся в «РАСЧЕТ ЗАКАЗА\credentials»: ищем и там (exe лежит в «Перемещения ЮА»)
_rz_cred = os.path.join(os.path.dirname(rz.BASE_DIR), RZ_DIR_NAME, "credentials")
if _rz_cred not in rz.CRED_DIRS:
    rz.CRED_DIRS.append(_rz_cred)

# ========================= НАСТРОЙКИ =========================

USE_GUI    = True    # False -> без окон (как --auto)
BQ_OFFLINE = os.environ.get("FM_OFFLINE", "").strip() == "1"   # 1 -> в BigQuery не ходить, только кэш

BQ_DATASET     = rz.BQ_PROJECT + ".family_market"
MATRIX_TABLE   = BQ_DATASET + ".assortment_matrix_full"
RECODE_TABLE   = BQ_DATASET + ".barcode_recode_map"
INCOMING_TABLE = BQ_DATASET + ".incoming_transactions"
COFFEE_SUPPLIER_LIKE = u"%каваапарат%"     # поставщик сырья кофеаппарата (СТВ Дистрибюшн)
HIST_DAYS = 120                             # окно истории поставок для списка «В матрице, на ЮА не было»

CACHE_TTL_H         = 1      # кэш свежее этого - в BigQuery не ходим
CACHE_WARN_H        = 24     # работа по кэшу старше суток -> предупреждение
MATRIX_MAX_AGE_DAYS = 3      # сама таблица матрицы в BigQuery старше -> предупреждение
MATRIX_MIN_ROWS     = 1500   # меньше ШК в матрице = неполная загрузка, работа останавливается
OUT_SHARE_ERROR     = 0.40   # вне матрицы больше этой доли позиций = ошибка сопоставления
GUARD_MIN_POS       = 30     # на очень малом остатке порог не применяем
STATE_WARN_H        = 24     # выгрузка Торгсофта старше -> предупреждение
UA_MIN_ROWS         = 300    # в «Состоянии склада ЮА» меньше строк -> подозрительно
# Приход «от инвентаризации» на склад ЮА - ПОД ВОПРОСОМ (источник товара уточняется: нельзя
# заводить на склад ЮА позицию Family инвентаризацией, нужен настоящий приход). Пока вопрос
# открыт, такой товар считается заведённым (чтобы не снять с полки то, что, возможно, законно),
# но везде помечается «только инвентаризация». False -> такой товар НЕ считается заведённым.
UA_COUNT_INVENTORY  = True

# Магазины при переходе на ЮА переименовываются, отличительных черт ЮА в имени нет:
# новое имя -> прежнее (Family). Имя из параметра/окна и из выгрузки считаются одним магазином.
STORE_RENAMES = {u"Болградська 38": u"Грозненська 38"}

DEST_NAME    = u"Полевая-Склад"          # куда везём (Family), только подпись в файлах
SUMMARY_NAME = u"_СВОДКА_ВЫВОЗ.xlsx"
INV_TXT_SUFFIX = u"_инвентаризация_ЮА.txt"
REPLACED_DIR = u"_замененные"

# Что с полки не снимаем. Основа - списки из raschet_zakaza (розлив, овощи, кеги, кофе
# из аппарата, крышки). Пакеты и стаканы там живут отдельно (is_bag / CUP_PACKS).
SKIP_BC_EXTRA = ()                                          # ШК вручную
CUP_BCS       = set(rz.CUP_PACKS) | set(v[0] for v in rz.CUP_PACKS.values())
ACCESSORY_BCS = set([rz.BC_KRISHKA, rz.BC_RUCHKA]) | set(rz.BC_PLYASHKI_ALL)   # комплектация розлива
RE_CUP        = re.compile(r"(^|\s)стакан")
RE_BC_OK      = re.compile(r"^\d{6,14}$")                   # как TXT_LINE_RE у расчёта заказа

STATE_ALIASES = {
    "name":  ("название товара", "назва товару", "наименование"),
    "bc":    ("штрих-код", "штрих код", "штрихкод"),
    "price": ("цена розничная", "ціна роздрібна", "цена"),
    "qty":   ("количество", "кількість"),
    "unit":  ("ед. изм", "од. вим", "ед.изм"),
    "cost":  ("себестоимость", "собівартість"),
    "store": ("склад",),
}
VERDICT_LABEL = {"OK": u"OK", "INFO": u"ИНФО", "WARN": u"ВНИМ", "ERROR": u"ОШИБ"}


class Dirs(object):
    """Папки входов и результата. Все пути от рабочей папки (os.path, без абсолютных)."""
    def __init__(self, base):
        self.base = base
        self.inp = os.path.join(base, u"ВХОД_ВЫВОЗ")
        self.stores = os.path.join(self.inp, u"МАГАЗИНЫ")
        self.ua_wh = os.path.join(self.inp, u"СКЛАД_ЮА")
        self.ua_stores = os.path.join(self.inp, u"МАГАЗИНЫ_ЮА")
        self.ua_receipts = os.path.join(self.inp, u"ПРИХОДЫ_ЮА")
        self.out = os.path.join(base, u"ВЫВОЗ")
        self.cache = os.path.join(base, u"Справочник")
        self.logs = os.path.join(base, u"ЛОГИ")

    def ensure(self):
        for p in (self.stores, self.ua_wh, self.ua_stores, self.ua_receipts,
                  self.out, self.cache, self.logs):
            os.makedirs(p, exist_ok=True)


def data_base(base, env=u""):
    """Где лежат ВХОД_ВЫВОЗ, ВЫВОЗ, кэш и логи: FM_VYVOZ_DIR, иначе папка «Перемещения ЮА»
    рядом с папкой программы, иначе сама папка программы. Ключ BigQuery ищется как и раньше
    (credentials рядом с raschet_zakaza)."""
    if env and env.strip():
        return env.strip()
    cand = os.path.join(os.path.dirname(base), u"Перемещения ЮА")
    return cand if os.path.isdir(cand) else base


DIRS = Dirs(data_base(rz.BASE_DIR, os.environ.get("FM_VYVOZ_DIR", "")))


def reset_report():
    R.checks[:] = []
    R.problems[:] = []
    R.anomalies[:] = []
    R.log[:] = []


# ========================= УТИЛИТЫ ===========================

def bc_key(x):
    """Ключ сравнения ШК: без ведущих нулей. Так сходятся BigQuery (19230010208)
    и Торгсофт (019230010208), у которого нули восстанавливает restore_barcode."""
    b = rz.fmt_barcode(x)
    return b.lstrip("0") if b else ""


def nname(s):
    """Название для точного сравнения: регистр, пробелы, * и апострофы не мешают."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    t = re.sub(u"[’`ʼ'*\"]", "", str(s).lower())
    return re.sub(r"\s+", " ", t).strip()


_TR = {u"а": "a", u"б": "b", u"в": "v", u"г": "g", u"ґ": "g", u"д": "d", u"е": "e", u"є": "ye", u"ж": "zh",
       u"з": "z", u"и": "y", u"і": "i", u"ї": "i", u"й": "y", u"к": "k", u"л": "l", u"м": "m", u"н": "n",
       u"о": "o", u"п": "p", u"р": "r", u"с": "s", u"т": "t", u"у": "u", u"ф": "f", u"х": "h", u"ц": "ts",
       u"ч": "ch", u"ш": "sh", u"щ": "sch", u"ы": "y", u"ь": "", u"ъ": "", u"э": "e", u"ю": "yu", u"я": "ya",
       u"ё": "e"}
_WORD_MAP = {u"єлоу": "yellow", u"еллоу": "yellow", u"єллоу": "yellow", u"йелоу": "yellow", u"пурбл": "purple",
             u"пурпл": "purple", u"пурпле": "purple"}
CIG_RE = re.compile(u"^(сигарет|твен)")             # названия Family: «Сигарети ...», «ТВЕН ...»
CIG_ANY_RE = re.compile(u"сигарет|твен")
KUL_RE = re.compile(u"^(випічка|кулінарія|выпечка|кулинария)\\b")      # своя кулинария/выпечка: не вывозим, на ЮА не переходит
CIG_AMBIG_GAP = 0.02                                # второй кандидат ближе этого - пара неоднозначна, не объединяем
CIG_PAIR_MIN = 0.9                                  # похожесть «скелетов» названий, с которой предлагаем пару


def name_skeleton(s):
    """Каркас названия для сравнения кириллицы и латиницы: «Собраніе Голд» ~ «Sobranie Gold».
    Транслит, слова-паразиты (сигарети, шт, 20), гласные и h убраны, повторы схлопнуты."""
    t = re.sub(u"(\\d)\\s*(шт|мг)(?![а-яіїєґ])", u"\\1 ", nname(s))        # «20шт» -> «20», «10мг» -> «10»
    t = re.sub(u"\\b(сигарети|сигарет|твен|шт|мг)\\b", " ", t)             # слова-паразиты - до транслита
    t = u" ".join(_WORD_MAP.get(w, w) for w in t.split())
    t = u"".join(_TR.get(ch, ch) for ch in t)
    t = re.sub(r"\b20\b", " ", t)
    for a, b in (("w", "v"), ("ph", "f"), ("ck", "k"), ("c", "k"), ("y", "i"), ("x", "ks"), ("q", "k")):
        t = t.replace(a, b)
    t = re.sub(r"[aeiouj]+", "", t).replace("h", "")
    t = re.sub(r"[^a-z0-9]+", " ", t).strip()
    return re.sub(r"(.)\1+", r"\1", t)


_PACK_H = str.maketrans({u"і": u"i", u"о": u"o", u"а": u"a", u"е": u"e", u"с": u"c", u"р": u"p", u"х": u"x", u"к": u"k",
                         u"т": u"t", u"м": u"m", u"н": u"h", u"в": u"b", u"у": u"y", u"ї": u"i", u"є": u"e"})
PACK_RE = re.compile(r"(?:^|\s)1\s*шт\.?$")            # штучная версия: название кончается на «1 шт» / «1шт»


def pack_base(name):
    """Название без «1 шт», «(24)», «Упаковка 24 шт»: штучный товар и его упаковка сводятся к одному ключу
    (латиница/кириллица-двойники и знаки не мешают)."""
    t = nname(name)
    t = re.sub(u"\\([^)]*\\)", u" ", t)
    t = re.sub(u"упаковка\\s*\\d+\\s*шт", u" ", t)
    t = re.sub(u"\\b\\d+\\s*шт\\b", u" ", t)
    t = re.sub(u"\\b1\\s*кг\\b", u" ", t)
    return re.sub(u"[^a-z0-9а-я]+", u"", t.translate(_PACK_H))


KEEP_FILE = u"vyvoz_keep.csv"
DEFAULT_KEEP = [(u"4820116280075", u"Сірники (1)", u"добавлены в матрицу 05.10.2026, на ЮА ещё не приходили")]


def load_keep_shk(dirs):
    """Список ШК «не вывозить» (Справочник vyvoz_keep.csv, правится в Excel): позиции, которые нужно оставить на полке,
    даже если они в матрице, а на ЮА не приходили. -> множество ШК без ведущих нулей. Файла нет - создаётся с Сірники (1)."""
    path = os.path.join(dirs.cache, KEEP_FILE)
    try:
        if not os.path.isfile(path):
            os.makedirs(dirs.cache, exist_ok=True)
            pd.DataFrame(DEFAULT_KEEP, columns=[u"ШК", u"Название", u"Комментарий"]).to_csv(path, index=False, encoding="utf-8-sig")
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")

        def first_bc(x):
            m = re.match(r"^\W*(\d{6,14})(?!\d)", str(x))       # Excel мог положить всю строку в первую ячейку: «ШК,Название,...»
            return bc_key(m.group(1)) if m else bc_key(x)
        got = set(k for k in (first_bc(x) for x in df.iloc[:, 0]) if k and re.match(r"^\d{6,14}$", k))
        if len(df) and not got:
            R.check("WARN", u"Список ШК «не вывозить»", u"в файле %d строк, но ни один ШК не прочитан: проверьте файл %s" % (len(df), KEEP_FILE))
        return got
    except Exception as e:
        R.check("WARN", u"Список ШК «не вывозить»", u"не прочитан (%s): работаю без него" % e)
        return set()


SHK_PAIRS_FILE = u"vyvoz_shk_pairs.csv"
DEFAULT_SHK_PAIRS = [(u"2978950018438", u"4000512992622", u"Бойчак Троллі Глаз 1 шт = Глотзер (упаковка); по вопросу 05.10.2026, подтвердить")]


def load_shk_pairs(dirs):
    """Таблица ручных пар «штучный ШК Family -> ШК упаковки на ЮА» (Справочник\\vyvoz_shk_pairs.csv, правится в Excel).
    -> {ШК Family без нулей: ШК ЮА без нулей}. Файла нет - создаётся с известными парами."""
    path = os.path.join(dirs.cache, SHK_PAIRS_FILE)
    try:
        if not os.path.isfile(path):
            os.makedirs(dirs.cache, exist_ok=True)
            pd.DataFrame(DEFAULT_SHK_PAIRS, columns=[u"ШК Family", u"ШК ЮА", u"Комментарий"]).to_csv(path, index=False, encoding="utf-8-sig")
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        return {bc_key(a): bc_key(b) for a, b in zip(df.iloc[:, 0], df.iloc[:, 1]) if bc_key(a) and bc_key(b)}
    except Exception as e:
        R.check("WARN", u"Таблица пар ШК Family-ЮА", u"не прочитана (%s): работаю без неё" % e)
        return {}


def group_of(name):
    """Группы в выгрузках нет: первое слово названия (подпись в файле это говорит)."""
    t = str(name).strip().split(" ")[0].strip(".,;:()") if str(name).strip() else ""
    return t


def list_xlsx(folder):
    if not os.path.isdir(folder):
        return []
    return [os.path.join(folder, n) for n in sorted(os.listdir(folder))
            if not n.startswith("~$") and n.lower().endswith((".xlsx", ".xlsm"))]


def _find_header(raw, groups, limit=15):
    """Строка заголовков: в ней должна быть подстрока из КАЖДОЙ группы."""
    for i in range(min(len(raw), limit)):
        vals = [rz.norm(v) for v in raw.iloc[i].tolist() if pd.notna(v)]
        if all(any(m in v for v in vals for m in grp) for grp in groups):
            return i
    return None


def _col(cols, aliases):
    low = {rz.norm(c): c for c in cols}
    for a in aliases:
        if a in low:
            return low[a]
    for a in aliases:
        for n, orig in low.items():
            if a in n:
                return orig
    return None


def _first_notna(s):
    s = s.dropna()
    return s.iloc[0] if len(s) else float("nan")


def _skey(s):
    k = rz.route_key(s)
    alias = {rz.route_key(a): rz.route_key(b) for a, b in list(rz.ROUTE_ALIASES.items()) + list(STORE_RENAMES.items())}
    return alias.get(k, k)


def _is_polevaya(s):
    return rz.route_key(s).startswith(u"полевая")


def guess_day(path):
    """Дата выгрузки по имени файла: ДД.ММ.ГГГГ (год из 4 цифр, номер дома из названия магазина не путается с датой:
    «Качанівська 19 8.10.2026» у raschet_zakaza читалось как 19.08.2010 и выбиралась старая выгрузка); иначе как в raschet_zakaza."""
    b = os.path.basename(path)
    for pat in (r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](20\d{2})(?!\d)", r"(?<!\d)(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2})(?!\d)"):
        m = re.search(pat, b)
        if m:
            d_, mo_, y_ = (int(g) for g in m.groups())
            try:
                return date(y_ + 2000 if y_ < 100 else y_, mo_, d_), u"из имени файла"
            except ValueError:
                pass
    return rz.guess_day(path)


def _pick_latest(paths):
    return sorted(paths, key=lambda p: (guess_day(p)[0], os.path.getmtime(p)))[-1]


def check_fresh(label, path):
    st = rz.read_export_stamp(path)
    how = u"по отметке выгрузки"
    if st is None:
        st, how = datetime.fromtimestamp(os.path.getmtime(path)), u"по дате файла"
    age = (datetime.now() - st).total_seconds() / 3600.0
    R.check("WARN" if age > STATE_WARN_H else "OK", u"Свежесть: %s" % label,
            u"%s, %s %s, %.0f ч назад" % (os.path.basename(path), how, st.strftime("%d.%m.%Y %H:%M"), age))
    return age


# ==================== ЧТЕНИЕ ВЫГРУЗОК ТОРГСОФТА ====================

def read_state_file(path):
    """«Состояние склада» -> (DataFrame shop, raw, bc, key, name, qty, unit, cost, price; инфо)."""
    raw = pd.read_excel(path, sheet_name=0, header=None)
    hi = _find_header(raw, [("штрих",), ("назв", "наимен")])
    if hi is None:
        raise ValueError(u"не найдена строка заголовков (Штрих-код / Название)")
    d = raw.iloc[hi + 1:].copy()
    d.columns = [str(c).strip() for c in raw.iloc[hi].tolist()]
    c = {k: _col(d.columns, v) for k, v in STATE_ALIASES.items()}
    miss = [k for k in ("bc", "qty", "name") if not c.get(k)]
    if miss:
        raise ValueError(u"нет столбцов %s; в файле: %s" % (miss, list(d.columns)))
    qty = pd.to_numeric(d[c["qty"]], errors="coerce")
    ok = d[c["bc"]].notna() & qty.notna()          # подвал «Исполнитель: ...» отпадает здесь
    skipped = int((d[c["bc"]].notna() & qty.isna()).sum())
    d, qty = d[ok], qty[ok]
    title = u""
    m = re.search(u'"([^"]+)"', str(raw.iloc[0, 0]))
    if m:
        title = m.group(1).strip()
    df = pd.DataFrame({
        "shop": (d[c["store"]].astype(str).str.strip() if c.get("store") else title),
        "raw": d[c["bc"]],
        "name": d[c["name"]].astype(str).str.strip().replace({"nan": ""}),
        "qty": qty.astype(float),
        "unit": (d[c["unit"]].astype(str).str.strip().replace({"nan": ""}) if c.get("unit") else ""),
        "cost": (pd.to_numeric(d[c["cost"]], errors="coerce") if c.get("cost") else float("nan")),
        "price": (pd.to_numeric(d[c["price"]], errors="coerce") if c.get("price") else float("nan")),
    })
    df["bc"] = [rz.restore_barcode(rz.fmt_barcode(v))[0] for v in df["raw"]]
    df["key"] = [bc_key(v) for v in df["raw"]]
    return df.reset_index(drop=True), {"rows": len(df), "skipped": skipped}


def read_receipts_file(path):
    """«Приходы» (Движение товара) -> строки с количеством > 0: key, name, sender."""
    raw = pd.read_excel(path, sheet_name=0, header=None)
    hi = _find_header(raw, [("штрих",), ("получател", "отримувач")])
    if hi is None:
        raise ValueError(u"не найдена строка заголовков (Штрих-код / Получатель)")
    d = raw.iloc[hi + 1:].copy()
    d.columns = [str(c).strip() for c in raw.iloc[hi].tolist()]
    c = {"name": _col(d.columns, STATE_ALIASES["name"]), "bc": _col(d.columns, STATE_ALIASES["bc"]),
         "qty": _col(d.columns, STATE_ALIASES["qty"]), "sender": _col(d.columns, (u"отправитель", u"відправник"))}
    if not (c["bc"] and c["qty"]):
        raise ValueError(u"нет столбцов Штрих-код / Количество")
    qty = pd.to_numeric(d[c["qty"]], errors="coerce")
    ok = d[c["bc"]].notna() & (qty > 0)           # нулевые приходы приходом не считаются
    d = d[ok]
    sender = (d[c["sender"]].astype(str).str.strip() if c["sender"] else pd.Series(u"", index=d.index))

    def kind(s):
        t = s.lower()
        if u"инвентар" in t or u"інвентар" in t:
            return "inv"
        if u"комплект" in t:                      # и «разукомплектация»
            return "comp"
        return "sup"
    return pd.DataFrame({
        "key": [bc_key(v) for v in d[c["bc"]]],
        "name": (d[c["name"]].astype(str).str.strip() if c["name"] else u""),
        "sender": sender.values,
        "kind": [kind(s) for s in sender],
    }).reset_index(drop=True)


UA_SRC_RANK = ((u"поставщик", 0), (u"магазин ЮА", 1), (u"приход в файле не найден", 2), (u"склад ЮА", 2),
               (u"комплектация", 3), (u"инвентаризац", 4))


def _src_rank(s):
    for pat, r in UA_SRC_RANK:
        if pat in s:
            return r
    return 5


def load_ua(dirs, use_ua=True):
    """Реестр «заведено на ЮА»: склад ЮА + склад магазина в базе ЮА + приходы ЮА.
    У каждого ШК пометка источника: приход поставщика / только инвентаризация / комплектация /
    приход в файле не найден. -> (DataFrame key, name, src; инфо)"""
    info = {"wh_file": u"", "wh_rows": 0, "shops_rows": 0, "receipts_rows": 0, "inv_only": 0}
    parts = []
    # приходы читаем первыми: по ним видно, откуда на складе ЮА каждый товар
    rc = []
    for p in list_xlsx(dirs.ua_receipts):
        try:
            rc.append(read_receipts_file(p))
        except Exception as e:
            R.check("WARN", u"Приходы ЮА: чтение", u"%s: %s" % (os.path.basename(p), e))
    rec = pd.concat(rc, ignore_index=True) if rc else pd.DataFrame(columns=["key", "name", "sender", "kind"])
    sup, inv, com = (set(rec.loc[rec["kind"] == k, "key"]) for k in ("sup", "inv", "comp"))
    info["receipts_rows"] = len(rec)
    if len(rec):
        R.check("OK", u"Приходы ЮА",
                u"%d строк с количеством > 0: поставщики %d ШК, инвентаризация %d ШК, комплектация %d ШК"
                % (len(rec), len(sup), len(inv), len(com)))
    else:
        R.check("INFO", u"Приходы ЮА", u"файлов нет: откуда на складе ЮА товар (приход или инвентаризация) определить нельзя")
    label = {"sup": u"приходы ЮА (поставщик)", "inv": u"приходы ЮА (инвентаризация)", "comp": u"приходы ЮА (комплектация)"}
    for kd, g in rec.groupby("kind"):
        parts.append(pd.DataFrame({"key": g["key"], "name": g["name"], "src": label[kd]}))

    def wh_src(k):
        if not len(rec):
            return u"склад ЮА"
        if k in sup:
            return u"склад ЮА (приход поставщика)"
        if k in inv:
            return u"склад ЮА (только инвентаризация)"
        if k in com:
            return u"склад ЮА (комплектация)"
        return u"склад ЮА (приход в файле не найден)"

    files = list_xlsx(dirs.ua_wh)
    if not files:
        R.check("ERROR" if use_ua else "INFO", u"Склад ЮА",
                u"нет файла «Состояние склада Полевая-Склад ЮА» в %s" % dirs.ua_wh)
    else:
        p = _pick_latest(files)
        if len(files) > 1:
            R.check("WARN", u"Склад ЮА: несколько файлов",
                    u"взят самый свежий: %s (всего %d)" % (os.path.basename(p), len(files)))
        try:
            df, meta = read_state_file(p)
        except Exception as e:
            R.check("ERROR", u"Склад ЮА: чтение", u"%s: %s" % (os.path.basename(p), e))
        else:
            info["wh_file"], info["wh_rows"] = os.path.basename(p), len(df)
            check_fresh(u"склад ЮА", p)
            shops = sorted(set(df["shop"]))
            if not any(u"юа" in rz.norm(s) for s in shops):
                R.check("WARN", u"Склад ЮА: имя склада",
                        u"в файле склад называется %s - это точно выгрузка из базы ЮА?" % shops)
            if len(df) < UA_MIN_ROWS:
                R.check("WARN", u"Склад ЮА: мало строк", u"%d (ожидается от %d)" % (len(df), UA_MIN_ROWS))
            R.check("OK", u"Склад ЮА: прочитан",
                    u"%s: %d строк, из них с остатком > 0: %d, нулевых %d, минусовых %d"
                    % (os.path.basename(p), len(df), int((df["qty"] > 0).sum()),
                       int((df["qty"] == 0).sum()), int((df["qty"] < 0).sum())))
            srcs = [wh_src(k) for k in df["key"]]
            parts.append(pd.DataFrame({"key": df["key"], "name": df["name"], "src": srcs, "bc": df["bc"]}))
            pos = df["qty"] > 0
            info["inv_only"] = int(sum(1 for s_, q in zip(srcs, pos) if q and u"только инвентаризация" in s_))
            if len(rec):
                tags = {}
                for s_, q in zip(srcs, pos):
                    if q:
                        tags[s_] = tags.get(s_, 0) + 1
                R.check("INFO", u"Склад ЮА: откуда товар с остатком",
                        u"; ".join(u"%s: %d" % kv for kv in sorted(tags.items())))
                R.check("WARN", u"Инвентаризация ЮА - под вопросом",
                        u"%d ШК склада ЮА с остатком заведены ТОЛЬКО инвентаризацией, а не приходом поставщика "
                        u"(источник товара уточняется). Сейчас они %s в реестре «заведено на ЮА» (UA_COUNT_INVENTORY=%s); "
                        u"в списках помечены «только инвентаризация»"
                        % (info["inv_only"], u"СЧИТАЮТСЯ" if UA_COUNT_INVENTORY else u"НЕ считаются", UA_COUNT_INVENTORY))
    for p in list_xlsx(dirs.ua_stores):
        try:
            df, meta = read_state_file(p)
        except Exception as e:
            R.check("WARN", u"Магазин в базе ЮА: чтение", u"%s: %s" % (os.path.basename(p), e))
            continue
        info["shops_rows"] += len(df)
        parts.append(pd.DataFrame({"key": df["key"], "name": df["name"], "src": u"магазин ЮА", "bc": df["bc"]}))
    if info["shops_rows"]:
        R.check("OK", u"Магазины в базе ЮА", u"%d строк" % info["shops_rows"])
    else:
        R.check("INFO", u"Магазины в базе ЮА", u"выгрузок нет (магазин ещё не перешёл на ЮА)")
    ua = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=["key", "name", "src", "bc"])
    ua["bc"] = ua["bc"].fillna(u"") if "bc" in ua.columns else u""
    ua = ua[ua["key"] != ""]
    if not UA_COUNT_INVENTORY:
        ua = ua[~ua["src"].str.contains(u"инвентаризац")]
    ua = ua.assign(_r=ua["src"].map(_src_rank)).sort_values("_r", kind="stable").drop(columns="_r")
    return ua.reset_index(drop=True), info


# ====================== BIGQUERY + КЭШ =======================

def _meta_read(dirs):
    try:
        with open(os.path.join(dirs.cache, "vyvoz_cache_meta.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _meta_write(dirs, meta):
    try:
        with open(os.path.join(dirs.cache, "vyvoz_cache_meta.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def _bq_query(sql, table_id=None):
    """-> (успех, DataFrame | сообщение, время изменения таблицы)"""
    if BQ_OFFLINE:
        return False, u"режим офлайн (FM_OFFLINE=1)", None
    cred = rz.find_credentials()
    if not cred:
        return False, u"не найден ключ: задай %s или положи json в %s" % (rz.CRED_ENV, os.path.join(rz._HERE, "credentials")), None
    try:
        from google.cloud import bigquery
    except ImportError:
        return False, u"нет пакета: pip install google-cloud-bigquery db-dtypes", None
    try:
        client = bigquery.Client(project=rz.BQ_PROJECT)
        df = client.query(sql).to_dataframe()
        modified = client.get_table(table_id).modified if table_id else None
        return True, df, modified
    except Exception as e:
        return False, u"%s: %s" % (type(e).__name__, str(e).replace("\n", " ")[:200]), None


def _table_age_check(label, modified, max_days):
    if not modified or not max_days:
        return
    if modified.tzinfo is None:
        modified = modified.replace(tzinfo=timezone.utc)
    days = (datetime.now(timezone.utc) - modified).total_seconds() / 86400.0
    R.check("WARN" if days > max_days else "OK", u"%s: свежесть таблицы в BigQuery" % label,
            u"обновлена %s (%.1f дн. назад)" % (modified.astimezone().strftime("%d.%m.%Y %H:%M"), days))


def bq_cached(dirs, name, sql, label, table_id=None, required=True, max_days=None):
    """BigQuery с локальным CSV-кэшем и офлайн-фолбэком (образец - справочник упаковок)."""
    os.makedirs(dirs.cache, exist_ok=True)
    path = os.path.join(dirs.cache, "vyvoz_%s.csv" % name)
    meta = _meta_read(dirs)
    have = os.path.exists(path)
    age = (time.time() - os.path.getmtime(path)) / 3600.0 if have else None

    def _read():
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")

    def _mod(s):
        try:
            return datetime.fromisoformat(s) if s else None
        except Exception:
            return None

    if have and age < CACHE_TTL_H:
        R.check("INFO", u"%s: источник" % label, u"кэш %.1f ч, запрос в BigQuery не нужен" % age)
        _table_age_check(label, _mod(meta.get(name, {}).get("modified")), max_days)
        return _read()
    ok, res, modified = _bq_query(sql, table_id)
    if ok:
        res = res.fillna("").astype(str)
        res.to_csv(path, index=False, encoding="utf-8-sig")
        meta[name] = {"fetched": datetime.now().isoformat(timespec="seconds"),
                      "modified": modified.isoformat() if modified else ""}
        _meta_write(dirs, meta)
        R.check("OK", u"%s: загружено из BigQuery" % label, u"%d строк" % len(res))
        _table_age_check(label, modified, max_days)
        return res
    if have:
        R.check("WARN" if age > CACHE_WARN_H else "INFO", u"%s: BigQuery недоступен" % label,
                u"%s -> работаю по кэшу возрастом %.1f ч" % (res, age))
        _table_age_check(label, _mod(meta.get(name, {}).get("modified")), max_days)
        return _read()
    R.check("ERROR" if required else "WARN", u"%s: нет данных" % label, u"%s; кэша нет" % res)
    return None


def validate_matrix(df):
    """Матрица должна быть настоящей: пустая или урезанная сделала бы «вне матрицы» всё подряд."""
    if df is None or len(df) == 0:
        return False, u"матрица пуста"
    if "barcode" not in df.columns:
        return False, u"в матрице нет столбца barcode"
    n = len(set(k for k in df["barcode"].map(bc_key) if k))
    if n < MATRIX_MIN_ROWS:
        return False, u"в матрице %d ШК, меньше %d: похоже на неполную загрузку" % (n, MATRIX_MIN_ROWS)
    return True, u"%d ШК" % n


MATRIX_SHEET_ID = u"1sbITHxuTGt7yIRR2PmLvvzIxD-4n_O4eYUDzUNDtBBA"        # Google-таблица «Ассортиментная матрица»
MATRIX_SHEET_NAME = u"Ассортиментная матрица (полная)"                    # лист, который заливается в assortment_matrix_full


def parse_matrix_sheet(values):
    """Строки листа Google -> DataFrame barcode, product_name, status, supplier. Строки ИТОГО и без ШК отбрасываются."""
    cols = ["barcode", "product_name", "status", "supplier"]
    if not values:
        return pd.DataFrame(columns=cols)
    head = [str(h).strip().lower() for h in values[0]]

    def ix(*names):
        for n in names:
            if n in head:
                return head.index(n)
        return None
    ib, inm, ist, isp = ix(u"штрихкод", u"штрих-код"), ix(u"товар"), ix(u"статус"), ix(u"поставщик")
    if ib is None:
        return pd.DataFrame(columns=cols)

    def cell(r, i):
        return str(r[i]).strip() if (i is not None and i < len(r)) else u""
    out = []
    for r in values[1:]:
        bc = re.sub(r"\D", "", cell(r, ib))
        name = cell(r, inm)
        if len(bc) < 6 or u"итого" in name.lower():
            continue
        out.append((bc, name, cell(r, ist), cell(r, isp)))
    return pd.DataFrame(out, columns=cols)


def fetch_matrix_sheet():
    """Лист матрицы из Google Sheets (тот же ключ сервис-аккаунта, только чтение). -> (успех, DataFrame | сообщение)"""
    if BQ_OFFLINE:
        return False, u"режим офлайн (FM_OFFLINE=1)"
    cred = rz.find_credentials()
    if not cred:
        return False, u"не найден ключ сервис-аккаунта"
    try:
        from urllib.parse import quote
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
        creds = service_account.Credentials.from_service_account_file(
            cred, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
        sess = AuthorizedSession(creds)
        url = "https://sheets.googleapis.com/v4/spreadsheets/%s/values/%s" % (MATRIX_SHEET_ID, quote(u"'%s'!A:P" % MATRIX_SHEET_NAME))
        r = sess.get(url, timeout=90)
        r.raise_for_status()
        return True, parse_matrix_sheet(r.json().get("values", []))
    except Exception as e:
        return False, u"%s: %s" % (type(e).__name__, str(e).replace("\n", " ")[:200])


def matrix_from_sheet(dirs):
    """Матрица из листа Google с локальным кэшем (как у остальных справочников). None - листа нет."""
    os.makedirs(dirs.cache, exist_ok=True)
    path = os.path.join(dirs.cache, "vyvoz_matrix_sheet.csv")
    have = os.path.exists(path)
    age = (time.time() - os.path.getmtime(path)) / 3600.0 if have else None

    def _read():
        return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    if have and age < CACHE_TTL_H:
        return _read()
    ok, res = fetch_matrix_sheet()
    if ok and len(res):
        res.to_csv(path, index=False, encoding="utf-8-sig")
        return res
    if have:
        R.check("WARN" if age > CACHE_WARN_H else "INFO", u"Матрица (лист Google): недоступна",
                u"%s -> по кэшу листа возрастом %.1f ч" % (res if not ok else u"лист пуст", age))
        return _read()
    R.check("INFO", u"Матрица (лист Google): недоступна", u"%s; беру BigQuery" % (res if not ok else u"лист пуст"))
    return None


def pick_matrix(m_sheet, m_bq):
    """Какую матрицу брать: лист Google свежее (в BigQuery её заливают раз в день). -> (DataFrame | None, примечание)"""
    ok_s = m_sheet is not None and validate_matrix(m_sheet)[0]
    ok_b = m_bq is not None and validate_matrix(m_bq)[0]
    if ok_s:
        note = u"лист Google «%s»: %d ШК" % (MATRIX_SHEET_NAME, len(set(m_sheet["barcode"].map(bc_key))))
        if ok_b:
            ks, kb = set(m_sheet["barcode"].map(bc_key)), set(m_bq["barcode"].map(bc_key))
            if ks != kb:
                note += u"; BigQuery отстаёт от листа: в листе новых +%d, убрано -%d (эти изменения учтены)" % (len(ks - kb), len(kb - ks))
        return m_sheet, note
    return (m_bq, u"BigQuery") if m_bq is not None else (None, u"нет")


def load_reference(dirs):
    """Матрица, перекодировка, сырьё кофеаппарата, история поставок -> (frames, ok)"""
    sql_m = u"SELECT barcode, product_name, status, supplier FROM `%s`" % MATRIX_TABLE
    m = bq_cached(dirs, "matrix", sql_m, u"Матрица", MATRIX_TABLE, True, MATRIX_MAX_AGE_DAYS)
    if m is not None and "supplier" not in m.columns:        # кэш старого формата (без поставщика): взять заново
        try:
            os.remove(os.path.join(dirs.cache, "vyvoz_matrix.csv"))
        except OSError:
            pass
        m = bq_cached(dirs, "matrix", sql_m, u"Матрица", MATRIX_TABLE, True, MATRIX_MAX_AGE_DAYS)
    m, m_note = pick_matrix(matrix_from_sheet(dirs), m)
    R.check("INFO", u"Матрица: источник", m_note)
    ok, msg = validate_matrix(m)
    if not ok:
        R.check("ERROR", u"Матрица", msg)
        return None
    R.check("OK", u"Матрица", msg)
    rec = bq_cached(dirs, "recode", u"SELECT old_barcode, new_barcode FROM `%s`" % RECODE_TABLE,
                    u"Перекодировка ШК", None, False)
    cof = bq_cached(dirs, "coffee",
                    u"SELECT DISTINCT barcode, product_name FROM `%s` WHERE LOWER(supplier) LIKE '%s'"
                    % (INCOMING_TABLE, COFFEE_SUPPLIER_LIKE), u"Сырьё кофеаппарата", None, False)
    if cof is None:
        R.check("WARN", u"Сырьё кофеаппарата", u"правило не применено: сырьё может попасть в вывоз")
    hist = bq_cached(dirs, "hist",
                     u"SELECT LTRIM(barcode,'0') AS barcode, STRING_AGG(DISTINCT delivery_type ORDER BY delivery_type) AS types, "
                     u"CAST(MAX(doc_date) AS STRING) AS last_date, "
                     u"ARRAY_AGG(supplier ORDER BY doc_date DESC LIMIT 1)[OFFSET(0)] AS last_supplier FROM `%s` "
                     u"WHERE doc_date >= DATE_SUB(CURRENT_DATE(), INTERVAL %d DAY) AND quantity > 0 GROUP BY 1"
                     % (INCOMING_TABLE, HIST_DAYS), u"История поставок", None, False)
    if hist is not None and "last_supplier" not in hist.columns:      # кэш старого формата: взять заново
        try:
            os.remove(os.path.join(dirs.cache, "vyvoz_hist.csv"))
        except OSError:
            pass
        hist = bq_cached(dirs, "hist",
                         u"SELECT LTRIM(barcode,'0') AS barcode, STRING_AGG(DISTINCT delivery_type ORDER BY delivery_type) AS types, "
                         u"CAST(MAX(doc_date) AS STRING) AS last_date, "
                         u"ARRAY_AGG(supplier ORDER BY doc_date DESC LIMIT 1)[OFFSET(0)] AS last_supplier FROM `%s` "
                         u"WHERE doc_date >= DATE_SUB(CURRENT_DATE(), INTERVAL %d DAY) AND quantity > 0 GROUP BY 1"
                         % (INCOMING_TABLE, HIST_DAYS), u"История поставок", None, False)
    return {"matrix": m, "recode": rec, "coffee": cof, "hist": hist}


# ===================== СПРАВОЧНИК СВЕРКИ =====================

class Ref(object):
    pass


def make_ref(matrix_df, recode_df=None, coffee_df=None, ua_df=None, hist_df=None, pairs=None, ua_map=None):
    r = Ref()
    r.matrix, r.matrix_names = {}, {}
    st = matrix_df["status"] if "status" in matrix_df.columns else [u""] * len(matrix_df)
    sp = matrix_df["supplier"] if "supplier" in matrix_df.columns else [u""] * len(matrix_df)
    for bc, nm, s, sup in zip(matrix_df["barcode"], matrix_df["product_name"], st, sp):
        k = bc_key(bc)
        if k:
            r.matrix[k] = (str(nm), str(s), str(sup).strip() if str(sup).lower() not in ("nan", "none") else u"")
            r.matrix_names.setdefault(nname(nm), []).append(k)
    r.recode = {}
    if recode_df is not None:
        for o, n in zip(recode_df["old_barcode"], recode_df["new_barcode"]):
            ko, kn = bc_key(o), bc_key(n)
            if ko and kn and ko != kn:
                r.recode[ko] = kn
    r.coffee_keys, r.coffee_names = set(), set()
    if coffee_df is not None:
        for bc, nm in zip(coffee_df["barcode"], coffee_df["product_name"]):
            if bc_key(bc):
                r.coffee_keys.add(bc_key(bc))
            if nname(nm):
                r.coffee_names.add(nname(nm))
    r.ua_keys, r.ua_names, r.ua_bc = {}, {}, {}
    if ua_df is not None:
        ubc = ua_df["bc"] if "bc" in ua_df.columns else [u""] * len(ua_df)
        for k, nm, src, b0 in zip(ua_df["key"], ua_df["name"], ua_df["src"], ubc):
            if k:
                r.ua_keys.setdefault(k, src)
                r.ua_names.setdefault(nname(nm), []).append((k, src, str(nm)))
                if str(b0).strip() and str(b0).lower() not in ("nan", "none"):
                    r.ua_bc.setdefault(k, str(b0).strip())
    r.pairs = {bc_key(a): bc_key(b) for a, b in (pairs or {}).items() if bc_key(a) and bc_key(b)}
    r.ua_map = {bc_key(a): bc_key(b) for a, b in (ua_map or {}).items() if bc_key(a) and bc_key(b)}     # таблица соответствия Family -> ЮА
    r.ua_recoded = {}             # склейка barcode_recode_map «ШК ЮА -> ШК матрицы» в обратную сторону: ШК Family/матрицы -> ШК ЮА
    for k in r.ua_keys:
        t = _resolve(r, k)
        if t != k and t not in r.ua_keys:
            r.ua_recoded.setdefault(t, k)
    r.ua_base, r.ua_name_by_key = {}, {}              # ЮА: название без «1 шт»/упаковки -> (ШК, название); ШК -> название
    for lst in r.ua_names.values():
        for k2, _src, nm2 in lst:
            r.ua_name_by_key.setdefault(k2, nm2)
            b2 = pack_base(nm2)
            if b2:
                r.ua_base.setdefault(b2, (k2, nm2))
    r.hist, r.last_sup = {}, {}
    if hist_df is not None:
        lsup = hist_df["last_supplier"] if "last_supplier" in hist_df.columns else [u""] * len(hist_df)
        for bc, types, last, sup in zip(hist_df["barcode"], hist_df["types"], hist_df["last_date"], lsup):
            if bc_key(bc):
                r.hist[bc_key(bc)] = (str(types), str(last))
                if str(sup).strip() and str(sup).lower() not in ("nan", "none"):
                    r.last_sup[bc_key(bc)] = str(sup).strip()
    r.ua_cig = []                                 # сигареты ЮА: каркасы названий для сопоставления с кириллицей Family
    seen = set()
    for lst in r.ua_names.values():
        for k2, _src, nm2 in lst:
            if k2 not in seen and CIG_ANY_RE.search(nname(nm2)):
                seen.add(k2)
                r.ua_cig.append((name_skeleton(nm2), k2, nm2))
    return r


def cig_pair(ref, key, name):
    """Сигарета Family, которой на ЮА нет по ШК: есть ли на ЮА то же по названию (латиница, другой ШК).
    -> (ШК на ЮА, название на ЮА, сходство, неоднозначно) либо ("", "", 0, False)."""
    if not CIG_RE.match(nname(name)) or not getattr(ref, "ua_cig", None):
        return u"", u"", 0.0, False
    sk, best, second = name_skeleton(name), (0.0, u"", u""), 0.0
    for sk2, k2, n2 in ref.ua_cig:
        if k2 == key:
            continue
        rt = difflib.SequenceMatcher(None, sk, sk2).ratio()
        if rt > best[0]:
            best, second = (rt, k2, n2), best[0]
        elif rt > second:
            second = rt
    if best[0] < CIG_PAIR_MIN:
        return u"", u"", 0.0, False
    return best[1], best[2], round(best[0], 2), (best[0] - second) < CIG_AMBIG_GAP


def merge_recoded_duplicates(agg, ref):
    """Старый и новый ШК одного товара (barcode_recode_map) в одном магазине - одна строка под НОВЫМ ШК, количества
    складываются. Одиночные перекодированные позиции не трогаем (их ШК остаётся тем, что на полке)."""
    if len(agg) == 0:
        return agg
    res_key = agg["key"].map(lambda k: _resolve(ref, k))
    dup = res_key.duplicated(keep=False)
    if not dup.any():
        return agg
    rows = []
    for rk, idx in res_key[dup].groupby(res_key[dup]).groups.items():
        g = agg.loc[list(idx)]
        base = g[g["key"] == rk]
        first = (base if len(base) else g).iloc[0]
        qty = float(g["qty"].sum())
        if not pd.isna(first["cost"]):
            cost = first["cost"]
        else:
            cost = g["cost"].dropna().iloc[0] if g["cost"].notna().any() else float("nan")
        row = first.to_dict()
        row.update({"qty": round(qty, 3), "cost": cost, "key": rk})
        if len(base) == 0:
            row["bc"] = rz.restore_barcode(rz.fmt_barcode(rk))[0]
        rows.append(row)
    keep = agg.loc[~dup]
    out = pd.concat([keep, pd.DataFrame(rows, columns=agg.columns)], ignore_index=True)
    return out.sort_values("name", kind="stable").reset_index(drop=True)


def _inv_mark(src):
    """Пометка для товара, заведённого на ЮА только инвентаризацией (вопрос открыт)."""
    return u" - только инвентаризация, под вопросом" if (src and u"инвентаризац" in src) else u""


SUPPLIERS_FILE = u"vyvoz_suppliers.csv"
# По умолчанию (просьба 05.10.2026) из вывоза исключены; названия бывают и как в матрице, и как в приходах
DEFAULT_EXCLUDED_SUPPLIERS = [u"УДК", u"Хладік", u"Хладопром", u"Прем*єр Фуд", u"Фаст Фуд",
                              u"Овочі-Фрукти (Протопопов)", u"ФОП Протопопов"]
_SUP_TR = {ord(u"і"): u"и", ord(u"ї"): u"и", ord(u"є"): u"е", ord(u"ё"): u"е", ord(u"ь"): None, ord(u"ъ"): None,
           ord(u"'"): None, ord(u"’"): None, ord(u"`"): None, ord(u"*"): None, ord(u'"'): None}


def sup_norm(s):
    """Имя поставщика для сравнения: регистр, і/и, ь, «*», кавычки и пробелы у скобок не мешают
    («Прем*єр Фуд» = «Премьер Фуд», «Арсенал ПК( Шейк)» = «Арсенал ПК (Шейк)»)."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return u""
    t = re.sub(r"\s+", " ", str(s).lower().translate(_SUP_TR)).strip()
    return re.sub(r"\s*([()])\s*", r"\1", t)


def save_excluded_suppliers(dirs, rules):
    """rules: {sup_norm: имя}. Файл vyvoz_suppliers.csv в папке Справочник: кто исключён из вывоза (остальные вывозятся)."""
    os.makedirs(dirs.cache, exist_ok=True)
    df = pd.DataFrame({u"поставщик": [rules[k] for k in sorted(rules, key=lambda x: rules[x].lower())],
                       u"статус": u"не вывозить"})
    df.to_csv(os.path.join(dirs.cache, SUPPLIERS_FILE), index=False, encoding="utf-8-sig")


def load_excluded_suppliers(dirs):
    """-> {sup_norm: имя}. Файла нет - создаётся со списком по умолчанию; файл есть - берётся как есть (можно править в Excel)."""
    path = os.path.join(dirs.cache, SUPPLIERS_FILE)
    if not os.path.isfile(path):
        save_excluded_suppliers(dirs, {sup_norm(n): n for n in DEFAULT_EXCLUDED_SUPPLIERS})
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        col = u"поставщик" if u"поставщик" in df.columns else df.columns[0]
        st = df[u"статус"].str.lower() if u"статус" in df.columns else pd.Series([u"не вывозить"] * len(df))
        return {sup_norm(n): n.strip() for n, x in zip(df[col], st) if n.strip() and not x.startswith(u"вывоз")}
    except Exception as e:
        R.check("WARN", u"Поставщики исключённые из вывоза", u"файл %s не прочитан (%s): беру список по умолчанию" % (path, e))
        return {sup_norm(n): n for n in DEFAULT_EXCLUDED_SUPPLIERS}


def suppliers_file_time(dirs):
    """Когда сохранён файл исключённых поставщиков («дд.мм чч:мм») или пусто, если файла нет."""
    try:
        return datetime.fromtimestamp(os.path.getmtime(os.path.join(dirs.cache, SUPPLIERS_FILE))).strftime("%d.%m %H:%M")
    except OSError:
        return u""


def suppliers_file_info(dirs):
    """(сколько поставщиков в списке «не вывозить», когда сохранён файл, путь): список действует на каждый расчёт, пока его не изменят."""
    n = len(load_excluded_suppliers(dirs))
    return n, suppliers_file_time(dirs), os.path.join(dirs.cache, SUPPLIERS_FILE)


def known_suppliers(dirs, info=None):
    """Все поставщики, которых можно отметить в окне: последний расчёт, матрица и история поставок (из кэша), файл правил.
    -> {sup_norm: {"name", "pos", "sum", "excluded"}}"""
    rules = load_excluded_suppliers(dirs)
    out = {}

    def add(name, pos=0, sm=0.0):
        n = sup_norm(name)
        if n:
            d = out.setdefault(n, {"name": str(name).strip(), "pos": 0, "sum": 0.0, "excluded": n in rules})
            d["pos"] += pos
            d["sum"] += sm
    for e in (info or {}).get("suppliers", []):
        add(e["supplier"], e["pos"], e["sum"])
    for fname, col in (("vyvoz_matrix.csv", "supplier"), ("vyvoz_hist.csv", "last_supplier")):
        try:
            df = pd.read_csv(os.path.join(dirs.cache, fname), dtype=str, keep_default_na=False, encoding="utf-8-sig")
            for nm in sorted(set(df[col])) if col in df.columns else []:
                add(nm)
        except Exception:
            pass
    for n, nm in rules.items():
        add(nm)
    return out


def _ua_bc(ref, k):
    """ШК товара так, как его знает база ЮА (для файла инвентаризации)."""
    return getattr(ref, "ua_bc", {}).get(k) or rz.restore_barcode(rz.fmt_barcode(k))[0]


def _resolve(ref, k):
    n = 0
    while k in ref.recode and n < 5:          # цепочки old -> new -> newer
        k, n = ref.recode[k], n + 1
    return k


def excl_reason(bc, key, name, ref):
    """Причина не снимать с полки, либо '' (только для позиций вне матрицы)."""
    why = rz._dist_skip_row(bc, name, u"")      # ШК-список, розлив, овощи, кеги, кофе готовый, крышка стакана
    if why:
        return why
    if bc in SKIP_BC_EXTRA:
        return u"ШК в списке исключений (вручную)"
    if rz.is_bag(name, bc):
        return u"пакет"
    if bc in CUP_BCS or RE_CUP.search(nname(name)[:30]):
        return u"стакан"
    if bc in ACCESSORY_BCS:
        return u"комплектация розлива (крышка/ручка/пляшка)"
    if key in ref.coffee_keys:
        return u"сырьё кофеаппарата (поставщик каваапарат, по ШК)"
    if nname(name) in ref.coffee_names:
        return u"сырьё кофеаппарата (поставщик каваапарат, по названию)"
    return u""


RES_COLS = ["shop", "bc", "key", "name", "qty", "unit", "cost", "sum", "no_cost", "group", "verdict",
            "reason", "other_key", "other_name", "other_src", "out_m", "recoded", "m_status", "hist",
            "supplier", "not_on_ua", "supplier_src", "pair_key", "pair_name", "pair_sim", "inv_bc", "pair_amb"]


def prepare_stock(df):
    """Строки выгрузки -> одна строка на (магазин, ШК); остаток > 0. -> (DataFrame, инфо)"""
    if len(df) == 0:
        return df.copy(), {"merged": 0, "zero": 0, "neg": 0}
    g = df.groupby(["shop", "key"], sort=False)
    agg = g.agg(bc=("bc", "first"), name=("name", "first"), unit=("unit", "first"),
                cost=("cost", _first_notna), qty=("qty", "sum")).reset_index()
    agg["qty"] = agg["qty"].round(3)
    info = {"merged": len(df) - g.ngroups, "zero": int((agg["qty"] == 0).sum()), "neg": int((agg["qty"] < 0).sum())}
    return agg[agg["qty"] > 0].reset_index(drop=True), info


def classify_stock(agg, ref, use_ua=True, remove_not_on_ua=False, excl_suppliers=None, keep_keys=None):
    """Каждой позиции с остатком - вердикт:
       IN / IN_UA_NAME (остаётся), CHECK1, CHECK2, EXCL, VYVOZ."""
    rows = []
    for r in agg.itertuples(index=False):
        key, bc, name, nn = r.key, r.bc, str(r.name), nname(r.name)
        alias = _resolve(ref, key)
        mkey = key if key in ref.matrix else (alias if alias in ref.matrix else None)
        by_name = False
        if mkey is None and nn:                    # того же названия в матрице нет по ШК, но есть под другим ШК: ОДИН товар
            mc0 = [k0 for k0 in ref.matrix_names.get(nn, []) if k0 != key]
            if mc0 and RE_BC_OK.match(bc) and not excl_reason(bc, key, name, ref):
                mkey, by_name = mc0[0], True
        out_m = mkey is None
        v, reason, ok_, on_, os_, m_status, hist = u"", u"", u"", u"", u"", u"", u""
        supplier, n_ua, sup_src, inv_bc = u"", False, u"", bc
        if mkey is not None:
            m_status = ref.matrix[mkey][1]
            supplier = ref.matrix[mkey][2]
            sup_src = u"матрица" if supplier else u""
            src = ref.ua_keys.get(key) or ref.ua_keys.get(mkey)
            mp = getattr(ref, "ua_map", {}).get(key) or getattr(ref, "ua_map", {}).get(mkey)
            rc = getattr(ref, "ua_recoded", {}).get(key) or getattr(ref, "ua_recoded", {}).get(mkey)
            if not use_ua:
                v, reason = "IN", u"в матрице"
            elif src:
                v, reason = "IN", u"в матрице, заведён на ЮА (%s)%s" % (src, _inv_mark(src))
            elif mp and mp in ref.ua_keys:
                v, reason = "IN_UA_NAME", u"в матрице; на ЮА под другим ШК (таблица соответствия)"
                ok_, on_, os_ = mp, ref.ua_name_by_key.get(mp, u""), ref.ua_keys[mp]
                inv_bc = _ua_bc(ref, mp)                       # в файл инвентаризации - ШК, который знает ЮА
            elif rc and rc in ref.ua_keys:
                v, reason = "IN_UA_NAME", u"в матрице; на ЮА под другим ШК (склейка barcode_recode_map)"
                ok_, on_, os_ = rc, ref.ua_name_by_key.get(rc, u""), ref.ua_keys[rc]
                inv_bc = _ua_bc(ref, rc)
            else:
                cand = [c for c in ref.ua_names.get(nn, []) if c[0] not in (key, mkey)]
                if cand:
                    v, reason = "IN_UA_NAME", u"в матрице; на ЮА под другим ШК (по названию)"
                    ok_, on_, os_ = ", ".join(sorted(set(c[0] for c in cand))[:3]), cand[0][2], cand[0][1]
                    inv_bc = _ua_bc(ref, cand[0][0])           # в файл инвентаризации - ШК, который знает ЮА
                else:
                    v, reason, n_ua = "CHECK2", u"в матрице, на ЮА не заведён", True
                    h = ref.hist.get(key) or ref.hist.get(mkey)
                    hist = (u"%s; посл. приход %s" % h) if h else u"нет приходов за %d дн." % HIST_DAYS
                    if remove_not_on_ua:
                        v, reason = "VYVOZ", u"в матрице, но на ЮА не завозилось (убираем по настройке)"
        else:
            why = excl_reason(bc, key, name, ref)
            mc = [k for k in ref.matrix_names.get(nn, []) if k != key] if nn else []
            src = ref.ua_keys.get(key) if use_ua else None
            cand = [c for c in ref.ua_names.get(nn, []) if c[0] != key] if (use_ua and nn) else []
            if why:
                v, reason = "EXCL", why
            elif not RE_BC_OK.match(bc):
                v, reason = "CHECK1", u"некорректный ШК (не 6-14 цифр) - в ТСД не загрузить"
            elif mc:
                v, reason = "CHECK1", u"в матрице под другим ШК (по названию)"
                ok_, on_, os_ = ", ".join(mc[:3]), ref.matrix[mc[0]][0], u"матрица"
            elif src:
                v, reason = "CHECK1", u"заведён на ЮА (по ШК)%s" % _inv_mark(src)
                ok_, on_, os_ = u"", u"", src      # тот же ШК: «другого» нет, где найден - в «Источнике»
            elif use_ua and getattr(ref, "ua_map", {}).get(key) in ref.ua_keys:
                mp = ref.ua_map[key]
                v, reason = "CHECK1", u"заведён на ЮА под другим ШК (таблица соответствия)"
                ok_, on_, os_ = mp, ref.ua_name_by_key.get(mp, u""), ref.ua_keys[mp]
                inv_bc = _ua_bc(ref, mp)
            elif use_ua and getattr(ref, "ua_recoded", {}).get(key) in ref.ua_keys:
                rc = ref.ua_recoded[key]
                v, reason = "CHECK1", u"заведён на ЮА под другим ШК (склейка barcode_recode_map)"
                ok_, on_, os_ = rc, ref.ua_name_by_key.get(rc, u""), ref.ua_keys[rc]
                inv_bc = _ua_bc(ref, rc)
            elif cand:
                v, reason = "CHECK1", u"заведён на ЮА под другим ШК (по названию)%s" % _inv_mark(cand[0][1])
                ok_, on_, os_ = ", ".join(sorted(set(c[0] for c in cand))[:3]), cand[0][2], cand[0][1]
            else:
                v, reason = "VYVOZ", u"нет в матрице, на ЮА не заведён"
        if mkey is not None:
            if by_name:
                reason += u" [то же название, в матрице под ШК %s]" % mkey
                if not ok_:
                    ok_, on_, os_ = mkey, ref.matrix[mkey][0], u"матрица"
            if v == "IN" and mkey != key and not ref.ua_keys.get(key) and ref.ua_keys.get(mkey):
                inv_bc = _ua_bc(ref, mkey)         # ЮА знает его под ШК матрицы (перекодировка / то же название)
        if not supplier:                      # вне матрицы (или в матрице без поставщика): последний приход из базы
            ls = ref.last_sup.get(key) or ref.last_sup.get(mkey or "") or ref.last_sup.get(_resolve(ref, key))
            if ls:
                supplier, sup_src = ls, u"последний приход"
        pk, pn, ps, am = (cig_pair(ref, key, name) if (use_ua and v in ("VYVOZ", "CHECK2")) else (u"", u"", 0.0, False))
        pack_hit = None                                # штучный товар, чья упаковка уже есть на ЮА под другим ШК
        if v == "VYVOZ" and use_ua:
            pm = getattr(ref, "pairs", {}).get(key)
            if pm and pm in ref.ua_keys:
                pack_hit = (pm, ref.ua_name_by_key.get(pm, u""), u"таблица пар ШК")
            elif PACK_RE.search(nn):
                h0 = getattr(ref, "ua_base", {}).get(pack_base(name))
                if h0 and h0[0] != key:
                    pack_hit = (h0[0], h0[1], u"название без «1 шт» / упаковки")
        why_x = excl_reason(bc, key, name, ref) if (v == "VYVOZ" and mkey is not None) else u""
        if why_x:                                      # расходники и пр. исключения не вывозим и когда они В МАТРИЦЕ, а на ЮА не было
            v, reason = "EXCL", why_x
        elif v == "VYVOZ" and CIG_RE.match(nn):        # сигареты не вывозим вовсе: остаются на полке
            v, reason = "EXCL", u"сигареты (не вывозим)"
            if pk and not am:                          # то же название на ЮА латиницей: объединяем, в файл - ШК ЮА
                inv_bc = _ua_bc(ref, pk)
        elif v == "VYVOZ" and KUL_RE.match(nn):        # кулинария/выпечка: не вывозим и в инвентаризацию ЮА не включаем
            v, reason = "KUL", u"кулинария (не вывозим, на ЮА не переходит)"
        elif v == "VYVOZ" and pack_hit:                # штучный товар: упаковка на ЮА, не вывозим и в инвентаризацию не включаем
            v, reason = "PACK", u"штучный товар: упаковка уже на ЮА (ШК %s), сопоставлено: %s" % (pack_hit[0], pack_hit[2])
            ok_, on_, os_ = pack_hit[0], pack_hit[1], u"склад ЮА (упаковка)"
        elif v == "VYVOZ" and keep_keys and key in keep_keys:       # ШК из списка «не вывозить»: остаётся на полке
            v, reason = "EXCL", u"в списке «не вывозить» по ШК (vyvoz_keep.csv)"
        elif v == "CHECK1" and keep_keys and key in keep_keys and "(по ШК)" in reason:   # «ждёт решения» -> решено: остаётся, идёт в инвентаризацию
            v, reason = "EXCL", u"в списке «не вывозить» по ШК (vyvoz_keep.csv); заведён на ЮА, допродаём"
        elif v == "VYVOZ" and excl_suppliers and supplier and sup_norm(supplier) in excl_suppliers:
            v, reason = "EXCL", u"поставщик исключён из вывоза: %s" % supplier      # остаётся на полке, идёт в инвентаризацию
        cost = r.cost
        rows.append({"shop": r.shop, "bc": bc, "key": key, "name": name, "qty": float(r.qty), "unit": r.unit,
                     "cost": cost, "sum": round(float(r.qty) * (0.0 if pd.isna(cost) else float(cost)), 2),
                     "no_cost": bool(pd.isna(cost)), "group": group_of(name), "verdict": v, "reason": reason,
                     "other_key": ok_, "other_name": on_, "other_src": os_, "out_m": out_m,
                     "recoded": bool(mkey is not None and mkey != key), "m_status": m_status, "hist": hist,
                     "supplier": supplier, "not_on_ua": n_ua, "supplier_src": sup_src,
                     "pair_key": pk, "pair_name": pn, "pair_sim": ps if pk else u"", "inv_bc": inv_bc, "pair_amb": bool(am)})
    return pd.DataFrame(rows, columns=RES_COLS)


def inventory_rows(res):
    """Что остаётся на полке (в матрице, заведено на ЮА, исключения): это переводится на ЮА инвентаризацией."""
    return res[res["verdict"].isin(INV_VERDICTS)]


def inventory_file_rows(iv):
    """Строки файла инвентаризации ЮА: по ШК, который знает ЮА; одинаковые ШК сливаются, количества складываются."""
    g = iv.groupby("inv_bc", sort=False).agg(qty=("qty", "sum"), unit=("unit", "first")).reset_index()
    g["qty"] = g["qty"].round(3)
    return g


def partition_check(res):
    """Весь остаток магазина делится на части: вывезти / перевести на ЮА / ждёт решения / остаётся без инвентаризации
    (кулинария и штучные, чья упаковка уже на ЮА; ключи kul_*)."""
    parts = {"vyvoz": res[res["verdict"] == "VYVOZ"], "inv": inventory_rows(res),
             "wait": res[res["verdict"].isin(WAIT_VERDICTS)], "kul": res[res["verdict"].isin(("KUL", "PACK"))]}
    out = {}
    for k, d in parts.items():
        out[k + "_pos"], out[k + "_qty"], out[k + "_sum"] = len(d), float(d["qty"].sum()), float(d["sum"].sum())
    out["ok"] = (out["vyvoz_pos"] + out["inv_pos"] + out["wait_pos"] + out["kul_pos"] == len(res))
    return out


def ua_name(shop):
    """Как магазин называется в базе ЮА (после переименования): Грозненська 38 -> Болградська 38."""
    k = _skey(shop)
    for new in STORE_RENAMES:
        if _skey(new) == k:
            return new
    return shop


def shop_stats(res):
    pos = len(res)
    vc = res["verdict"].value_counts()
    v = res[res["verdict"] == "VYVOZ"]
    out_m = int(res["out_m"].sum())
    inv_only = int((res["reason"] + u" " + res["other_src"]).str.contains(u"только инвентаризация").sum())
    pc = partition_check(res)
    return {"inv_pos": pc["inv_pos"], "inv_units": round(pc["inv_qty"], 3), "inv_sum": round(pc["inv_sum"], 2),
            "wait_pos": pc["wait_pos"], "wait_units": round(pc["wait_qty"], 3), "wait_sum": round(pc["wait_sum"], 2),
            "kul_pos": pc["kul_pos"], "kul_units": round(pc["kul_qty"], 3), "kul_sum": round(pc["kul_sum"], 2),
            "partition_ok": pc["ok"], "positions": pos, "units": round(float(res["qty"].sum()), 3), "cost_sum": round(float(res["sum"].sum()), 2),
            "inv_only": inv_only,
            "counts": {k: int(vc.get(k, 0)) for k in ("IN", "IN_UA_NAME", "CHECK1", "CHECK2", "EXCL", "KUL", "PACK", "VYVOZ")},
            "out_m": out_m, "out_share": (out_m / float(pos)) if pos else 0.0,
            "vyvoz_pos": len(v), "vyvoz_units": round(float(v["qty"].sum()), 3), "vyvoz_sum": round(float(v["sum"].sum()), 2)}


def guard_failed(st):
    """Вне матрицы больше OUT_SHARE_ERROR позиций остатка - это ошибка сопоставления, не реальность."""
    if st["positions"] >= GUARD_MIN_POS and st["out_share"] > OUT_SHARE_ERROR:
        return True, (u"вне матрицы %.1f%% позиций (%d из %d), порог %d%% - похоже на ошибку сопоставления; txt не создан"
                      % (100 * st["out_share"], st["out_m"], st["positions"], int(OUT_SHARE_ERROR * 100)))
    return False, u""


def name_collisions(names):
    """Имена, которые после учёта переименований - один магазин: [(новое, прежнее)].
    Данные магазинов под разными именами НЕ склеиваются: это ошибка раскладки выгрузок."""
    seen, out = {}, []
    for n in names:
        k = _skey(n)
        if k in seen:
            pair = tuple(sorted((seen[k], n), key=lambda x: (rz.route_key(x) in
                                                           {rz.route_key(a) for a in STORE_RENAMES}, x), reverse=True))
            out.append(pair)
        else:
            seen[k] = n
    return out


def resolve_shops(requested, available):
    """Имена из параметра/окна -> имена из выгрузок. Полевая не вывозится. -> (магазины, ошибки)"""
    amap = {_skey(n): n for n in available if not _is_polevaya(n)}
    got, errs = [], []
    for q in requested:
        q = str(q).strip()
        if not q:
            continue
        if _is_polevaya(q):
            errs.append(u"«%s» - Полевую не трогаем, вывозить из неё нечего" % q)
            continue
        name = amap.get(_skey(q))
        if name is None:
            sim = difflib.get_close_matches(q, list(amap.values()), n=3, cutoff=0.6)
            errs.append(u"Магазин «%s» не найден среди выгрузок.%s" % (q, (u" Похожие: " + u"; ".join(sim)) if sim else u""))
        elif name not in got:
            got.append(name)
    return got, errs


# ======================= ЗАПИСЬ ФАЙЛОВ =======================

INV_VERDICTS = ("IN", "IN_UA_NAME", "EXCL")      # остаётся на полке -> переводится на ЮА инвентаризацией
WAIT_VERDICTS = ("CHECK1", "CHECK2")             # ждёт решения: убрать или оставить
# Названия листов (до 31 символа): по смыслу, а не «Проверить 1/2»
SHEET_VYVOZ = u"Вывезти на склад"
SHEET_INV = u"Остаётся - инвентаризация ЮА"
SHEET_C1 = u"Нет в матрице, но есть на ЮА"
SHEET_C2 = u"В матрице, на ЮА не было"
SHEET_SUP = u"По поставщикам"
SHEET_EXCL = u"Исключено (не трогаем)"
SHEET_DIFF = u"Один товар, разные ШК"
SHEET_NOM = u"Нет в матрице (весь список)"
SHEET_PACK = u"Штучные - упаковка на ЮА"
SHEET_CIG = u"Сигареты - возможные пары"
OUT_M = "OUT_M"             # отбор листа «Нет в матрице (весь список)»: все позиции остатка, которых нет в матрице
NOT_ON_UA = "NOT_ON_UA"      # отбор листа «В матрице, на ЮА не было»: и «ждёт решения», и ушедшее в вывоз по настройке
ITEM_COLS = [(u"Штрих-код", "bc", 16), (u"Название товара", "name", 50), (u"Количество", "qty", 12),
             (u"Ед. изм.", "unit", 8), (u"Себестоимость", "cost", 13), (u"Сумма", "sum", 13)]
SHEETS = [
    (SHEET_NOM, OUT_M, ITEM_COLS + [(u"Решение", "decision", 34), (u"Причина", "reason", 46),
                                    (u"Поставщик (последний приход)", "supplier", 30)]),
    (SHEET_INV, INV_VERDICTS, ITEM_COLS + [(u"ШК в файле для ЮА", "inv_bc", 18), (u"Основание", "reason", 56)]),
    (SHEET_C1, ("CHECK1",), ITEM_COLS + [(u"Причина", "reason", 46), (u"Другой ШК", "other_key", 18),
                                        (u"Название у другого ШК", "other_name", 46), (u"Источник", "other_src", 16)]),
    (SHEET_C2, NOT_ON_UA, ITEM_COLS + [(u"Поставщик (по матрице)", "supplier", 28), (u"Статус в матрице", "m_status", 30),
                                      (u"Поставки за %d дн. (по базе)" % HIST_DAYS, "hist", 36)]),
    (SHEET_EXCL, ("EXCL", "KUL"), ITEM_COLS + [(u"Причина", "reason", 52)]),
]
SUP_COLS = [(u"Поставщик (по матрице)", "supplier", 32), (u"Позиций", "pos", 11), (u"Сумма", "sum", 13),
            (u"Прямая доставка, поз.", "direct", 14), (u"Через РЦ / опт, поз.", "rc", 14),
            (u"Без поставок за %d дн., поз." % HIST_DAYS, "none", 16)]
PACK_COLS = [(u"Штрих-код Family", "bc", 16), (u"Название (штучный)", "name", 50), (u"Количество", "qty", 12),
             (u"Ед. изм.", "unit", 8), (u"Себестоимость", "cost", 13), (u"Сумма", "sum", 13),
             (u"Поставщик", "supplier", 26), (u"ШК упаковки на ЮА", "other_key", 20), (u"Название на ЮА", "other_name", 50),
             (u"Основание", "reason", 56)]
CIG_COLS = [(u"Штрих-код Family", "bc", 16), (u"Название", "name", 46), (u"Количество", "qty", 12), (u"Сумма", "sum", 13),
            (u"Где сейчас", "decision", 34), (u"Штрих-код на ЮА", "pair_key", 16), (u"Название на ЮА", "pair_name", 42),
            (u"Сходство", "pair_sim", 10), (u"Статус", "pair_status", 32)]
DIFF_COLS = [(u"Куда сопоставлено", "other_src", 18), (u"Штрих-код магазина", "bc", 16), (u"Название", "name", 46),
             (u"Штрих-код там", "other_key", 20), (u"Название там", "other_name", 46), (u"Результат", "reason", 46)]
TEXT_COLS = ("bc", "other_key", "key", "pair_key", "inv_bc")
CENTER_COLS = ("bc", "other_key", "qty", "unit", "cost", "sum", "pos", "direct", "rc", "none", "pair_key", "pair_sim", "inv_bc")   # ШК и числа по центру
HEAD_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
CELL_CENTER = Alignment(horizontal="center", vertical="center")


def qty_format(v):
    """Целое количество - формат «0»: «0.###» рисует хвост «5,» с запятой. Дробное - «0.###» (8,232)."""
    try:
        return "0" if float(v).is_integer() else "0.###"
    except (TypeError, ValueError):
        return "0.###"


def _write_table(ws, df, cols, start_row=1, shop_col=False):
    cols = ([(u"Магазин", "shop", 24)] + cols) if shop_col else cols
    for j, (title, _c, w) in enumerate(cols, 1):
        c0 = ws.cell(row=start_row, column=j, value=title)
        c0.font = Font(bold=True)
        c0.fill = rz.HEAD_FILL
        c0.border = rz.BORDER
        c0.alignment = HEAD_ALIGN
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, rec in enumerate(df.to_dict("records"), 1):
        for j, (_t, col, _w) in enumerate(cols, 1):
            v = rec.get(col)
            if v is None or (isinstance(v, float) and pd.isna(v)):
                v = None
            c0 = ws.cell(row=start_row + i, column=j, value=v)
            if col in CENTER_COLS:
                c0.alignment = CELL_CENTER
            if col in TEXT_COLS:
                c0.number_format = "@"
                c0.value = None if v is None else str(v)
            elif col == "qty":
                c0.number_format = qty_format(v)
            elif col in ("cost", "sum"):
                c0.number_format = "#,##0.00"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    if len(df):
        ws.auto_filter.ref = "A%d:%s%d" % (start_row, get_column_letter(len(cols)), start_row + len(df))


def diff_rows(res):
    """Один товар под разными ШК: сопоставлено по названию, а ШК не совпадают."""
    return res[(res["other_key"] != "") & (res["other_key"] != res["key"])]


def _decision(res):
    """Что с позицией: подпись для листов «Нет в матрице (весь список)» и «Сигареты»."""
    def one(r):
        if r.verdict == "VYVOZ":
            return u"Вывезти (в матрице, на ЮА не было)" if r.not_on_ua else u"Вывезти (нет в матрице)"
        if r.verdict == "CHECK1":
            return u"Нет в матрице, но есть на ЮА"
        if r.verdict in ("EXCL", "KUL"):
            return u"Исключено (не трогаем)"
        if r.verdict == "PACK":
            return u"Штучный товар, упаковка на ЮА: не вывозим"
        if r.verdict == "CHECK2":
            return u"В матрице, на ЮА не было"
        return u"Остаётся (инвентаризация ЮА)"
    return [one(r) for r in res.itertuples()]


def _pick(res, sel):
    """Строки листа: по вердиктам, «в матрице, на ЮА не было» либо «нет в матрице» (по поставщику, затем по названию)."""
    if sel == NOT_ON_UA:
        return res[res["not_on_ua"]].sort_values(["supplier", "name"], kind="stable")
    if sel == OUT_M:
        d = res[res["out_m"]].sort_values(["supplier", "name"], kind="stable")
        return d.assign(decision=_decision(d))
    return res[res["verdict"].isin(sel)]


def cig_rows(res):
    d = res[res["pair_key"] != ""].sort_values(["name"], kind="stable")
    st = [u"неоднозначно - проверить, ШК не менялся" if a else (u"объединено: в файле ШК ЮА" if ib != b else u"пара найдена, ШК не менялся")
          for a, ib, b in zip(d["pair_amb"], d["inv_bc"], d["bc"])]
    return d.assign(decision=_decision(d), pair_status=st)


def supplier_table(res):
    """Для закупщицы: по поставщикам матрицы - сколько позиций «в матрице, на ЮА не было» и как они приходили."""
    d = res[res["not_on_ua"]]
    rows = []
    for sp, g in d.groupby(d["supplier"].replace("", u"(поставщик не указан)"), sort=False):
        types = g["hist"].str.split(";").str[0]
        none = g["hist"].str.startswith(u"нет приходов") | (g["hist"] == "")
        rows.append({"supplier": sp, "pos": len(g), "sum": round(float(g["sum"].sum()), 2),
                     "direct": int(types.str.contains(u"Прямая").sum()),
                     "rc": int((types.str.contains(u"РЦ") | types.str.contains(u"Опт")).sum()),
                     "none": int(none.sum())})
    out = pd.DataFrame(rows, columns=["supplier", "pos", "sum", "direct", "rc", "none"])
    return out.sort_values(["sum", "supplier"], ascending=[False, True], kind="stable").reset_index(drop=True)


def _sheet_lists(wb, res, shop_col=False):
    for title, sel, cols in SHEETS:
        ws = wb.create_sheet(title)
        _write_table(ws, _pick(res, sel), cols, shop_col=shop_col)
        if title == SHEET_C2:
            _write_table(wb.create_sheet(SHEET_SUP), supplier_table(res), SUP_COLS)
        if title == SHEET_EXCL:
            _write_table(wb.create_sheet(SHEET_PACK), res[res["verdict"] == "PACK"], PACK_COLS, shop_col=shop_col)
    ws = wb.create_sheet(SHEET_DIFF)
    _write_table(ws, diff_rows(res), DIFF_COLS, shop_col=shop_col)
    _write_table(wb.create_sheet(SHEET_CIG), cig_rows(res), CIG_COLS, shop_col=shop_col)


MANUAL_MARK = u"обычно не вывозим: "
MANUAL_HEAD = u"Обычно не вывозим (причина)"


def manual_verdicts(res, excl_suppliers=None, keep_keys=None):
    """Ручной выбор (решение пользователя 08.10.2026: «сигареты вывозим, кег вывозим»; галочки «не вывозить» действуют всегда).
    В «Вывезти на склад» возвращаются только сигареты и кеги, которые правила оставляли на полке, - если их поставщик не отмечен
    «не вывозить» и ШК не в списке «не вывозить»; прежняя причина - в отдельной колонке. Поставщики с галочкой, список ШК, пакеты,
    стаканы и расходники, сырьё кофеаппарата, овощи, кулинария и выпечка остаются вне списка, как в обычном расчёте."""
    res = res.copy()
    if len(res) == 0:
        return res
    nn = res["name"].map(nname)
    free = (nn.map(lambda x: bool(CIG_RE.match(x))) | res["name"].map(is_keg)).astype(bool)
    xs = set(excl_suppliers or ())
    ticked = res["supplier"].map(lambda x: bool(sup_norm(x)) and sup_norm(x) in xs).astype(bool)
    kept = res["key"].isin(set(keep_keys or ())).astype(bool)
    m = (res["verdict"] == "EXCL") & free & ~ticked & ~kept
    res.loc[m, "reason"] = MANUAL_MARK + res.loc[m, "reason"].astype(str)
    res.loc[m, "verdict"] = "VYVOZ"
    return res


def write_store_book(path, shop, day, res, st, mode_note=u""):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_VYVOZ
    v = res[res["verdict"] == "VYVOZ"].sort_values(["supplier", "name"], kind="stable")
    ws["A1"] = u"Вывоз вне матрицы: %s -> %s" % (shop, DEST_NAME)
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = (u"Остатки на %s (по учёту Торгсофта)     Позиций: %d     Единиц: %s     Себестоимость: %s"
                % (day.strftime("%d.%m.%Y"), len(v), round(float(v["qty"].sum()), 3), round(float(v["sum"].sum()), 2)) + mode_note)
    ws["A2"].font = Font(italic=True, size=9)
    heads = [u"№", u"Штрих-код", u"Название товара", u"Количество", u"Ед. изм.", u"Себестоимость", u"Сумма",
             u"Поставщик", u"Возможно уже на ЮА под другим ШК (сигареты)"]
    widths = [5, 16, 52, 12, 8, 13, 13, 30, 44]
    manual = bool(v["reason"].astype(str).str.startswith(MANUAL_MARK).any())
    if manual:
        heads, widths = heads + [MANUAL_HEAD], widths + [52]
        ws["A3"] = (u"РУЧНОЙ ВЫБОР: в списке и сигареты, кеги (обычно их не вывозят, причина - в последней колонке). Поставщики с "
                    u"галочкой «не вывозить», список ШК, пакеты, стаканы и расходники, сырьё кофеаппарата, овощи, кулинария в список не "
                    u"попадают. Лишнее удалите, лист сохраните в «Корректировка_ЮА». Файл txt рядом - до вашей правки, для ТСД его не берите.")
        ws["A3"].font = Font(bold=True, color="C00000")
    for j, (h, w) in enumerate(zip(heads, widths), 1):
        c0 = ws.cell(row=4, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = rz.HEAD_FILL
        c0.border = rz.BORDER
        c0.alignment = HEAD_ALIGN
        ws.column_dimensions[get_column_letter(j)].width = w
    r = 4
    for n, rec in enumerate(v.to_dict("records"), 1):
        r += 1
        ws.cell(row=r, column=1, value=n).alignment = CELL_CENTER
        b = ws.cell(row=r, column=2, value=str(rec["bc"]))
        b.number_format = "@"
        b.alignment = CELL_CENTER
        ws.cell(row=r, column=3, value=rec["name"])
        q = ws.cell(row=r, column=4, value=round(float(rec["qty"]), 3))
        q.number_format = qty_format(q.value)
        q.alignment = CELL_CENTER
        ws.cell(row=r, column=5, value=rec["unit"]).alignment = CELL_CENTER
        cc = ws.cell(row=r, column=6, value=None if rec["no_cost"] else round(float(rec["cost"]), 2))
        cc.number_format = "#,##0.00"
        cc.alignment = CELL_CENTER
        sc = ws.cell(row=r, column=7, value=rec["sum"])
        sc.number_format = "#,##0.00"
        sc.alignment = CELL_CENTER
        ws.cell(row=r, column=8, value=rec["supplier"] or None)
        if rec["pair_key"]:
            ws.cell(row=r, column=9, value=u"%s  %s" % (rec["pair_key"], rec["pair_name"]))
        if manual and str(rec["reason"]).startswith(MANUAL_MARK):
            ws.cell(row=r, column=10, value=str(rec["reason"])[len(MANUAL_MARK):])
    r += 1
    ws.cell(row=r, column=3, value=u"ИТОГО:").font = Font(bold=True)
    tq = ws.cell(row=r, column=4, value=round(float(v["qty"].sum()), 3))
    tq.font = Font(bold=True)
    tq.number_format = qty_format(tq.value)
    tq.alignment = CELL_CENTER
    t = ws.cell(row=r, column=7, value=round(float(v["sum"].sum()), 2))
    t.font = Font(bold=True)
    t.fill = rz.WARN_FILL
    t.number_format = "#,##0.00"
    t.alignment = CELL_CENTER
    ws.freeze_panes = "A5"
    _sheet_lists(wb, res)
    wb.save(path)


def write_checks_sheet(ws):
    fills = {"OK": rz.OK_FILL, "WARN": rz.WARN_FILL, "ERROR": rz.ERR_FILL,
             "INFO": PatternFill("solid", fgColor="EDEDED")}
    verdict = u"ЕСТЬ ОШИБКИ" if R.errors else (u"ЕСТЬ ПРЕДУПРЕЖДЕНИЯ" if R.warns else u"ВСЁ ЧИСТО")
    ws["A1"] = u"ИТОГ: %s (ошибок %d, предупреждений %d)" % (verdict, len(R.errors), len(R.warns))
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].fill = rz.ERR_FILL if R.errors else (rz.WARN_FILL if R.warns else rz.OK_FILL)
    for j, h in enumerate([u"Статус", u"Проверка", u"Детали"], 1):
        c0 = ws.cell(row=3, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = rz.HEAD_FILL
        c0.alignment = HEAD_ALIGN
    for n, (st, nm, det) in enumerate(R.checks, 1):
        ws.cell(row=3 + n, column=1, value=VERDICT_LABEL[st]).fill = fills[st]
        ws.cell(row=3 + n, column=2, value=nm)
        ws.cell(row=3 + n, column=3, value=det)
    for j, w in enumerate([10, 52, 110], 1):
        ws.column_dimensions[get_column_letter(j)].width = w


def write_summary_book(path, day, entries, all_res, diag):
    wb = Workbook()
    ws = wb.active
    ws.title = u"На склад"
    v = all_res[all_res["verdict"] == "VYVOZ"]
    ws["A1"] = u"Что едет на %s со всех магазинов (%s)" % (DEST_NAME, day.strftime("%d.%m.%Y"))
    ws["A1"].font = Font(bold=True, size=13)
    shops = sorted(set(v["shop"]))
    heads = [u"Штрих-код", u"Название товара", u"Поставщик", u"Количество", u"Сумма", u"Магазинов"] + shops
    for j, h in enumerate(heads, 1):
        c0 = ws.cell(row=3, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = rz.HEAD_FILL
        c0.alignment = HEAD_ALIGN
        ws.column_dimensions[get_column_letter(j)].width = [16, 50, 30, 12, 13, 10][j - 1] if j <= 6 else 14
    r = 3
    if len(v):
        piv = v.pivot_table(index=["bc", "name", "supplier"], columns="shop", values="qty", aggfunc="sum", fill_value=0)
        sums = v.groupby("bc")["sum"].sum()
        for (bc, name, grp), row in piv.iterrows():
            r += 1
            b = ws.cell(row=r, column=1, value=str(bc))
            b.number_format = "@"
            b.alignment = CELL_CENTER
            ws.cell(row=r, column=2, value=name)
            ws.cell(row=r, column=3, value=grp)
            tq = ws.cell(row=r, column=4, value=round(float(row.sum()), 3))
            tq.number_format = qty_format(tq.value)
            tq.alignment = CELL_CENTER
            ts = ws.cell(row=r, column=5, value=round(float(sums.get(bc, 0)), 2))
            ts.number_format = "#,##0.00"
            ts.alignment = CELL_CENTER
            ws.cell(row=r, column=6, value=int((row > 0).sum())).alignment = CELL_CENTER
            for j, s in enumerate(shops, 7):
                q = float(row.get(s, 0))
                sc = ws.cell(row=r, column=j, value=round(q, 3) if q else None)
                sc.number_format = qty_format(q)
                sc.alignment = CELL_CENTER
        r += 1
        ws.cell(row=r, column=2, value=u"ИТОГО:").font = Font(bold=True)
        tq = ws.cell(row=r, column=4, value=round(float(v["qty"].sum()), 3))
        tq.font = Font(bold=True)
        tq.number_format = qty_format(tq.value)
        tq.alignment = CELL_CENTER
        ts = ws.cell(row=r, column=5, value=round(float(v["sum"].sum()), 2))
        ts.font = Font(bold=True)
        ts.number_format = "#,##0.00"
        ts.alignment = CELL_CENTER
    ws.freeze_panes = "C4"

    w2 = wb.create_sheet(u"По магазинам")
    heads = [u"Магазин", u"Позиций с остатком", u"Единиц", u"Остаток, себестоимость", u"В матрице", u"Вне матрицы",
             u"% вне матрицы", u"Вывезти, позиций", u"Вывезти, единиц", u"Вывезти, сумма", SHEET_C1,
             SHEET_C2, u"Исключено", u"На ЮА только по инвентаризации (под вопросом)",
             u"Перевести на ЮА, позиций", u"Перевести на ЮА, единиц", u"Перевести на ЮА, сумма",
             u"Имя в базе ЮА", u"Статус", u"Папка"]
    for j, (h, w) in enumerate(zip(heads, [26, 11, 11, 14, 10, 10, 10, 11, 11, 13, 10, 10, 10, 18, 12, 12, 14, 24, 52, 60]), 1):
        c0 = w2.cell(row=1, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = rz.HEAD_FILL
        c0.alignment = HEAD_ALIGN
        w2.column_dimensions[get_column_letter(j)].width = w
    for n, e in enumerate(entries, 2):
        s = e["stats"]
        c = s["counts"]
        vals = [e["shop"], s["positions"], s["units"], s["cost_sum"], s["positions"] - s["out_m"], s["out_m"],
                round(100 * s["out_share"], 1), s["vyvoz_pos"], s["vyvoz_units"], s["vyvoz_sum"], c["CHECK1"],
                c["CHECK2"], c["EXCL"] + c["KUL"] + c["PACK"], s["inv_only"], s["inv_pos"], s["inv_units"], s["inv_sum"],
                ua_name(e["shop"]), e["status"], e["folder"]]
        for j, val in enumerate(vals, 1):
            cell = w2.cell(row=n, column=j, value=val)
            if 2 <= j <= 17:                       # числа по центру; имя магазина, статус и папка - слева
                cell.alignment = CELL_CENTER
            if j == heads.index(u"Статус") + 1:
                cell.fill = rz.OK_FILL if val == "OK" else rz.ERR_FILL
    w2.freeze_panes = "B2"

    for title, sel, cols in SHEETS:
        w = wb.create_sheet(title)
        _write_table(w, _pick(all_res, sel), cols, shop_col=True)
        if title == SHEET_C2:
            _write_table(wb.create_sheet(SHEET_SUP), supplier_table(all_res), SUP_COLS)
        if title == SHEET_EXCL:
            _write_table(wb.create_sheet(SHEET_PACK), all_res[all_res["verdict"] == "PACK"], PACK_COLS, shop_col=True)
    w = wb.create_sheet(SHEET_DIFF)
    _write_table(w, diff_rows(all_res), DIFF_COLS, shop_col=True)
    _write_table(wb.create_sheet(SHEET_CIG), cig_rows(all_res), CIG_COLS, shop_col=True)
    if len(diag):
        w = wb.create_sheet(u"Диагностика")
        _write_table(w, diag, ITEM_COLS + [(u"Причина", "reason", 46)], shop_col=True)
    write_checks_sheet(wb.create_sheet(u"Проверки"))
    wb.save(path)


def _read_txt_dict(path, problems, label):
    with open(path, "r", encoding=rz.TXT_ENCODING) as f:
        lines = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    got = {}
    for ln in lines:
        if not rz.TXT_LINE_RE.match(ln):
            problems.append(u"%s: неверный формат строки '%s'" % (label, ln))
            continue
        b, q = ln.split(";")
        got[b] = round(float(q.replace(",", ".")), 3)
    return got


def verify_inventory(shop_dir, fname, inv_rows):
    """Лист «Перевести на ЮА» и txt инвентаризации перечитываются и сверяются с расчётом."""
    problems = []
    expected = {r.bc: round(float(r.qty), 3) for r in inv_rows.itertuples()}                 # лист: по ШК Family
    expected_txt = {r.inv_bc: round(float(r.qty), 3) for r in inventory_file_rows(inv_rows).itertuples()}   # txt: по ШК ЮА
    wb = load_workbook(os.path.join(shop_dir, fname + ".xlsx"), data_only=True)
    ws = wb[SHEET_INV]
    got, r = {}, 2
    while ws.cell(row=r, column=1).value not in (None, ""):
        q = ws.cell(row=r, column=3).value
        got[rz.fmt_barcode(ws.cell(row=r, column=1).value)] = round(float(q), 3) if isinstance(q, (int, float)) else None
        r += 1
    wb.close()
    if got != expected:
        problems.append(u"лист «%s»: расхождений %d" % (SHEET_INV, len(set(expected.items()) ^ set(got.items()))))
    if expected:
        tg = _read_txt_dict(os.path.join(shop_dir, fname + INV_TXT_SUFFIX), problems, u"txt инвентаризации")
        if tg != expected_txt:
            problems.append(u"txt инвентаризации: расхождений %d" % len(set(expected_txt.items()) ^ set(tg.items())))
    return problems


def verify_vyvoz(shop_dir, fname, rows):
    """Обратное чтение xlsx и txt со сверкой с расчётом (как verify_written, но макет другой)."""
    problems = []
    expected = {r.bc: round(float(r.qty), 3) for r in rows.itertuples()}
    xlsx_path, txt_path = os.path.join(shop_dir, fname + ".xlsx"), os.path.join(shop_dir, fname + ".txt")
    wb = load_workbook(xlsx_path, data_only=False)
    ws = wb[SHEET_VYVOZ]
    got, r = {}, 5
    while ws.cell(row=r, column=2).value not in (None, ""):
        b = rz.fmt_barcode(ws.cell(row=r, column=2).value)
        q = ws.cell(row=r, column=4).value
        got[b] = round(float(q), 3) if isinstance(q, (int, float)) else None
        r += 1
    wb.close()
    if got != expected:
        diff = ["%s: расчёт %s / xlsx %s" % (b, expected.get(b), got.get(b)) for b in sorted(set(expected) | set(got))
                if expected.get(b) != got.get(b)]
        problems.append(u"xlsx: расхождений %d -> %s" % (len(diff), "; ".join(diff[:6])))
    if not expected:
        return problems
    with open(txt_path, "r", encoding=rz.TXT_ENCODING) as f:
        lines = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    if len(lines) != len(expected):
        problems.append(u"txt: строк %d вместо %d" % (len(lines), len(expected)))
    tgot = {}
    for ln in lines:
        if not rz.TXT_LINE_RE.match(ln):
            problems.append(u"txt: неверный формат строки '%s'" % ln)
            continue
        b, q = ln.split(";")
        tgot[b] = round(float(q.replace(",", ".")), 3)
    if tgot != expected:
        diff = ["%s: расчёт %s / txt %s" % (b, expected.get(b), tgot.get(b)) for b in sorted(set(expected) | set(tgot))
                if expected.get(b) != tgot.get(b)]
        problems.append(u"txt: расхождений %d -> %s" % (len(diff), "; ".join(diff[:6])))
    return problems


def prepare_day_dir(dirs, day, tag=u""):
    """Папка результата дня. tag - метка режима (зачистка остатка лежит отдельно: ВЫВОЗ\\2026-10-08_зачистка), чтобы
    повторный запуск обычного вывоза за ту же дату не уносил её в _замененные."""
    d = os.path.join(dirs.out, day.strftime("%Y-%m-%d") + tag)
    busy = rz.locked_files(d)
    if busy:
        R.check("ERROR", u"Файлы прошлого вывоза заняты",
                u"закройте в Excel и запустите снова: %s" % "; ".join(os.path.basename(b) for b in busy[:5]))
        return None
    if os.path.isdir(d) and os.listdir(d):
        dst0 = os.path.join(dirs.out, REPLACED_DIR, "%s_%s" % (os.path.basename(d), datetime.now().strftime("%H-%M-%S")))
        dst, n = dst0, 2
        while os.path.exists(dst):                    # два запуска в одну секунду: новая папка с суффиксом, а не «внутрь» первой
            dst, n = "%s_%d" % (dst0, n), n + 1
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.move(d, dst)
        R.check("INFO", u"Прошлый результат этой даты", u"убран в %s" % dst)
    os.makedirs(d, exist_ok=True)
    return d


# ========================= ОСНОВНОЙ ПРОГОН =========================

def scan_stores(dirs):
    """{магазин: (файл, DataFrame)}: у магазина в нескольких файлах берётся свежий (иначе двойной счёт)."""
    out = {}
    for p in list_xlsx(dirs.stores):
        try:
            df, meta = read_state_file(p)
        except Exception as e:
            R.check("ERROR", u"Чтение выгрузки магазина", u"%s: %s" % (os.path.basename(p), e))
            continue
        if meta["skipped"]:
            R.check("WARN", u"Строки без количества", u"%s: пропущено %d" % (os.path.basename(p), meta["skipped"]))
        for shop, g in df.groupby("shop"):
            if shop in out:
                keep = _pick_latest([out[shop][0], p])
                R.check("WARN", u"Магазин в нескольких файлах", u"%s: взят %s" % (shop, os.path.basename(keep)))
                if keep == out[shop][0]:
                    continue
            out[shop] = (p, g.reset_index(drop=True))
    return out


def run_vyvoz(dirs, shops=None, ref=None, use_ua=True, day=None, remove_not_on_ua=False, excl_suppliers=None, keep_keys=None,
              sweep=False, manual=False):
    """Весь расчёт без окон. shops=None -> все магазины из ВХОД_ВЫВОЗ\\МАГАЗИНЫ (кроме Полевой).
    ref - готовый справочник (для самотеста); иначе матрица берётся из BigQuery/кэша."""
    reset_report()
    dirs.ensure()
    info = {"ok": False, "day_dir": u"", "shops": [], "errors": 0, "summary": u"", "log": u"", "excl_n": None, "excl_when": u"", "manual": False}

    stores = scan_stores(dirs)
    if not stores:
        R.check("ERROR", u"Выгрузки магазинов", u"в %s нет файлов «Состояние склада»" % dirs.stores)
    for p in sorted(set(v[0] for v in stores.values())):
        check_fresh(u"магазины", p)
    available = sorted(stores)
    for new, old in name_collisions(available):
        R.check("ERROR", u"Один магазин под двумя именами",
                u"«%s» и «%s» в папке МАГАЗИНЫ: выгрузки базы ЮА кладите в МАГАЗИНЫ_ЮА, в МАГАЗИНЫ - только базу Family" % (new, old))
    if shops is None:
        sel = [s for s in available if not _is_polevaya(s)]
        for s in available:
            if _is_polevaya(s):
                R.check("INFO", u"Полевая", u"%s в выгрузках есть, но не обрабатывается" % s)
    else:
        sel, errs = resolve_shops(shops, available)
        for e in errs:
            R.check("ERROR", u"Выбор магазинов", e)
    if stores and not sel and not R.errors:
        R.check("ERROR", u"Выбор магазинов", u"нечего обрабатывать")
    R.check("INFO", u"Магазинов в работе", u"%d: %s" % (len(sel), u"; ".join(sel)))

    ua_df, uinfo = load_ua(dirs, use_ua)
    if ref is None and not R.errors:
        fr = load_reference(dirs)
        if fr is not None:
            ref = make_ref(fr["matrix"], fr["recode"], fr["coffee"], ua_df, fr["hist"], pairs=load_shk_pairs(dirs),
                           ua_map=load_ua_map(dirs)[0])
    if ref is not None and use_ua:
        inm = sum(1 for k in set(ua_df["key"]) if k in ref.matrix)
        R.check("INFO", u"Склад ЮА и матрица",
                u"в реестре ЮА %d ШК, из них в матрице %d (матрица всего %d ШК)"
                % (len(set(ua_df["key"])), inm, len(ref.matrix)))
    if remove_not_on_ua:
        R.check("WARN", u"Режим: убираем и то, что в матрице, но на ЮА не завозилось",
                u"такие позиции идут в «Вывезти на склад» (в списке «В матрице, на ЮА не было» они остаются для закупщицы; "
                u"настройка включена вручную)")
    if not use_ua:
        R.check("WARN", u"Режим без ЮА", u"проверки по ЮА отключены: списки по ЮА пусты, в вывоз попадёт то, что заведено на ЮА")

    if R.errors or ref is None or not sel:
        R.check("ERROR", u"Остановка", u"критические ошибки, файлы вывоза не записаны")
        _fail_report(dirs, day)
        info["errors"] = len(R.errors)
        return info

    if day is None:
        day = max(guess_day(stores[s][0])[0] for s in sel)
    R.check("INFO", u"Дата вывоза", u"%s (по дате выгрузки магазинов)" % day.strftime("%d.%m.%Y"))
    if use_ua and uinfo["wh_file"]:
        ud = guess_day(_pick_latest(list_xlsx(dirs.ua_wh)))[0]
        if abs((ud - day).days) > 1:
            R.check("WARN", u"Разные даты выгрузок", u"склад ЮА от %s, магазины от %s" % (ud.strftime("%d.%m.%Y"), day.strftime("%d.%m.%Y")))
    day_dir = prepare_day_dir(dirs, day, SWEEP_DAY_TAG if sweep else u"")
    if day_dir is None:
        _fail_report(dirs, day)
        info["errors"] = len(R.errors)
        return info
    info["day_dir"] = day_dir

    excl_set = set(load_excluded_suppliers(dirs)) if excl_suppliers is None else set(excl_suppliers)
    manual = bool(manual and not sweep)
    info["manual"] = manual
    if not sweep:                                         # в зачистке список поставщиков не действует (в ручном выборе - действует)
        info["excl_n"], info["excl_when"] = len(excl_set), suppliers_file_time(dirs)
    R.check("INFO", u"Поставщики, исключённые из вывоза", u"%d: %s" % (
        len(excl_set), u", ".join(sorted(excl_set)) if excl_set else u"никто (вывозятся все)"))
    keep_set = load_keep_shk(dirs) if keep_keys is None else set(keep_keys)
    R.check("INFO", u"Список ШК «не вывозить»", u"%d ШК" % len(keep_set))
    if manual:
        R.check("WARN", u"Режим: РУЧНОЙ ВЫБОР",
                u"сигареты и кеги идут в «Вывезти на склад» (причина, по которой их обычно не вывозят, - в последней колонке), если их "
                u"поставщик не отмечен «не вывозить» и ШК не в списке «не вывозить»; остальные исключения действуют как обычно")
    sweep_rules = None
    if sweep:
        sweep_rules = load_sweep_stay(dirs)
        R.check("WARN", u"Режим: ЗАЧИСТКА ОСТАТКА",
                u"вывозится всё, независимо от матрицы и ЮА; остаются только заморозка, скоропорты, расходники и сырьё кофеаппарата "
                u"(правила: %s; поставщиков %d, фрагментов названий %d). Результат лежит отдельно: %s"
                % (SWEEP_STAY_FILE, len(sweep_rules["sup"]), len(sweep_rules["name"]), os.path.basename(day_dir)))
    entries, used, results, diag = [], {}, [], []
    for shop in sel:
        path, df = stores[shop]
        agg, pinfo = prepare_stock(df)
        n0 = len(agg)
        agg = merge_recoded_duplicates(agg, ref)
        pinfo["merged"] += n0 - len(agg)
        res = classify_stock(agg, ref, use_ua, remove_not_on_ua, excl_set, keep_set)
        if sweep:
            res = sweep_verdicts(res, sweep_rules, ref)
        elif manual:
            res = manual_verdicts(res, excl_set, keep_set)
        st = shop_stats(res)
        xs = res[res["reason"].str.startswith(u"поставщик исключён")]
        if len(xs):
            R.check("INFO", u"%s: поставщики исключены из вывоза" % shop,
                    u"%d поз. / %s грн остаются на полке (%s)" % (len(xs), round(float(xs["sum"].sum()), 2),
                                                                  u", ".join(sorted(set(xs["supplier"]))[:8])))
        c = st["counts"]
        R.check("INFO", u"%s: остаток" % shop,
                u"позиций %d (минусовых %d, нулевых %d, слито дублей %d), единиц %s, себестоимость %s"
                % (st["positions"], pinfo["neg"], pinfo["zero"], pinfo["merged"], st["units"], st["cost_sum"]))
        R.check("INFO" if st["partition_ok"] else "ERROR", u"%s: баланс остатка" % shop,
                u"вывезти %d поз. / %s грн + перевести на ЮА %d / %s + ждёт решения %d / %s + кулинария и штучные (остаются, без инвентаризации) %d / %s"
                u" = остаток %d поз. / %s грн"
                % (st["vyvoz_pos"], st["vyvoz_sum"], st["inv_pos"], st["inv_sum"], st["wait_pos"], st["wait_sum"],
                   st["kul_pos"], st["kul_sum"], st["positions"], st["cost_sum"]))
        R.check("INFO", u"%s: разбор" % shop,
                u"в матрице и на ЮА %d (+%d на ЮА под другим ШК), вывезти %d, проверить 1: %d, проверить 2: %d, исключено %d, штучных (упаковка на ЮА) %d; "
                u"перекодировано %d; на ЮА только по инвентаризации (под вопросом): %d"
                % (c["IN"], c["IN_UA_NAME"], c["VYVOZ"], c["CHECK1"], c["CHECK2"], c["EXCL"] + c["KUL"], c["PACK"],
                   int(res["recoded"].sum()), st["inv_only"]))
        failed, msg = (False, u"") if sweep else guard_failed(st)      # при зачистке «вне матрицы» не ошибка сопоставления
        entry = {"shop": shop, "stats": st, "status": "OK", "folder": u""}
        if failed:
            R.check("ERROR", u"%s: доля вне матрицы" % shop, msg)
            entry["status"] = u"ОШИБ: " + msg
            top = res[res["out_m"]].sort_values("sum", ascending=False).head(50)
            top = top.assign(reason=u"вне матрицы (диагностика ошибки сопоставления)")
            diag.append(top)
            entries.append(entry)
            continue
        R.check("OK", u"%s: доля вне матрицы" % shop, u"%.1f%% (порог %d%%)" % (100 * st["out_share"], int(OUT_SHARE_ERROR * 100)))
        v = res[res["verdict"] == "VYVOZ"]
        if v["no_cost"].any():
            R.check("WARN", u"%s: нет себестоимости" % shop, u"у %d позиций вывоза, сумма считается без них" % int(v["no_cost"].sum()))
        if (v["unit"] == "").any():
            R.check("INFO", u"%s: пустая единица измерения" % shop, u"у %d позиций вывоза" % int((v["unit"] == "").sum()))
        base = rz.safe_name(shop)
        nm, k = base, 2
        while nm.lower() in used:
            nm, k = "%s (%d)" % (base, k), k + 1
        used[nm.lower()] = shop
        shop_dir = os.path.join(day_dir, nm)
        os.makedirs(shop_dir, exist_ok=True)
        write_store_book(os.path.join(shop_dir, nm + ".xlsx"), shop, day, res, st,
                         mode_note=(u"     РЕЖИМ: ЗАЧИСТКА ОСТАТКА" if sweep else u""))
        if sweep:
            try:
                with open(os.path.join(shop_dir, u"ЧИТАЙ_МЕНЯ.txt"), "w", encoding="utf-8") as f_:
                    f_.write(u"ЗАЧИСТКА ОСТАТКА: %s -> всё, кроме заморозки, скоропортов и расходников.\n"
                             u"%s.txt - общий список без деления. В ТСД берите файлы из папки ТСД_по_адресам (после «Распределения»).\n"
                             u"Остаётся на месте: лист «Исключено (не трогаем)» (причина в колонке).\n" % (shop, nm))
            except OSError:
                pass
        if len(v):
            tr = pd.DataFrame({"bc": v["bc"].values, "qty": v["qty"].values, "step": 1.0,
                               "is_weight": [(str(u).lower() in (u"кг", u"kg")) for u in v["unit"]]})
            rz.write_shop_txt(os.path.join(shop_dir, nm + ".txt"), tr)
        iv = inventory_rows(res)
        if len(iv):
            ivg = inventory_file_rows(iv)
            ir = pd.DataFrame({"bc": ivg["inv_bc"].values, "qty": ivg["qty"].values, "step": 1.0,
                               "is_weight": [(str(u).lower() in (u"кг", u"kg")) for u in ivg["unit"]]})
            rz.write_shop_txt(os.path.join(shop_dir, nm + INV_TXT_SUFFIX), ir)
        probs = verify_vyvoz(shop_dir, nm, v) + verify_inventory(shop_dir, nm, iv)
        if probs:
            entry["status"] = u"ОШИБ: сверка файлов: " + "; ".join(probs)
            R.check("ERROR", u"%s: обратная сверка" % shop, "; ".join(probs))
        else:
            R.check("OK", u"%s: обратная сверка" % shop,
                    u"xlsx и txt перечитаны, совпали с расчётом (вывезти %d поз., %s ед.; перевести на ЮА %d поз., %s ед.)"
                    % (len(v), st["vyvoz_units"], st["inv_pos"], st["inv_units"]) if len(v)
                    else u"вывозить нечего, txt вывоза не создан (перевести на ЮА %d поз.)" % st["inv_pos"])
        entry["folder"] = shop_dir
        entries.append(entry)
        results.append(res)

    if not results:
        all_res = pd.DataFrame(columns=RES_COLS)
    else:
        all_res = pd.concat(results, ignore_index=True)
    diag_df = pd.concat(diag, ignore_index=True) if diag else pd.DataFrame(columns=RES_COLS)
    tot = sum(e["stats"]["vyvoz_units"] for e in entries if e["status"] == "OK")
    R.check("INFO", u"Итого к вывозу", u"магазинов %d, позиций %d, единиц %s, себестоимость %s"
            % (sum(1 for e in entries if e["status"] == "OK"),
               sum(e["stats"]["vyvoz_pos"] for e in entries if e["status"] == "OK"), round(tot, 3),
               round(sum(e["stats"]["vyvoz_sum"] for e in entries if e["status"] == "OK"), 2)))
    info["summary"] = os.path.join(day_dir, SUMMARY_NAME)
    write_summary_book(info["summary"], day, entries, all_res, diag_df)
    info["shops"] = [{"shop": e["shop"], "status": e["status"], "vyvoz_pos": e["stats"]["vyvoz_pos"],
                      "vyvoz_units": e["stats"]["vyvoz_units"], "vyvoz_sum": e["stats"]["vyvoz_sum"],
                      "check1": e["stats"]["counts"]["CHECK1"], "check2": e["stats"]["counts"]["CHECK2"],
                      "excl": e["stats"]["counts"]["EXCL"], "inv_pos": e["stats"]["inv_pos"], "inv_sum": e["stats"]["inv_sum"],
                      "wait_pos": e["stats"]["wait_pos"], "kul_pos": e["stats"]["kul_pos"], "folder": e["folder"]} for e in entries]
    cand = all_res[(all_res["verdict"] == "VYVOZ") | all_res["reason"].str.startswith(u"поставщик исключён")]
    info["suppliers"] = [{"supplier": sp, "pos": len(g), "sum": round(float(g["sum"].sum()), 2)}
                         for sp, g in cand.groupby(cand["supplier"].replace("", u"(поставщик не указан)"))]
    info["errors"], info["ok"], info["day"] = len(R.errors), not R.errors, day
    try:
        info["log"] = os.path.join(dirs.logs, "vyvoz_log_%s.txt" % datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
        with open(info["log"], "w", encoding="utf-8") as f:
            f.write("\n".join(R.log))
    except Exception:
        pass
    return info


# ======================= ЗАЧИСТКА ОСТАТКА =======================
# Решение 08.10.2026: остаток магазина вывозится ЦЕЛИКОМ независимо от матрицы и ЮА; остаются на месте только заморозка,
# скоропорты и расходники. Что именно остаётся - Справочник\vyvoz_perishable.csv (правится в Excel).

SWEEP_STAY_FILE = u"vyvoz_perishable.csv"
SWEEP_DAY_TAG = u"_зачистка"
DEFAULT_SWEEP_STAY = [
    (u"поставщик", u"Гама", u"решение 08.10.2026: не вывозим"),
    (u"поставщик", u"Хладік", u"решение 08.10.2026: заморозка"),
    (u"поставщик", u"Хладопром", u"решение 08.10.2026: мороженое Хладік"),
    (u"поставщик", u"Ферма", u"решение 08.10.2026"),
    (u"поставщик", u"Волошкове Поле", u"решение 08.10.2026"),
    (u"поставщик", u"Комо", u"решение 08.10.2026"),
    (u"поставщик", u"Айсберг-Фіш", u"решение 08.10.2026"),
    (u"поставщик", u"Хмельницька Маслосирбаза", u"решение 08.10.2026"),
    (u"поставщик", u"Бащінський МК", u"решение 08.10.2026 (в т.ч. нагетсы: заморозка)"),
    (u"поставщик", u"Кулінарія", u"скоропорт"),
    (u"поставщик", u"Овочі-Фрукти (Протопопов)", u"овощи"),
    (u"поставщик", u"Соління", u"решение 08.10.2026"),
    (u"поставщик", u"Пакети", u"решение 08.10.2026: пакеты никакие не вывозим"),
    (u"поставщик", u"Якобз (Кава апарат)", u"сырьё кофеаппарата"),
    (u"название", u"бащинськ", u"решение 08.10.2026"),
    (u"название", u"бащінськ", u"решение 08.10.2026"),
    (u"название", u"нагетс", u"заморозка"),
    (u"название", u"хладік", u"заморозка"),
    (u"название", u"морозив", u"заморозка"),
    (u"название", u"пельмен", u"заморозка"),
    (u"название", u"п/ф", u"заморозка"),
    (u"название", u"заморож", u"заморозка"),
    (u"название", u"^волошкове", u"решение 08.10.2026"),
    (u"название", u"^комо ", u"решение 08.10.2026"),
    (u"название", u"^гама ", u"решение 08.10.2026"),
    (u"название", u"^ферма ", u"решение 08.10.2026"),
    (u"название", u"соління", u"решение 08.10.2026"),
]


def load_sweep_stay(dirs):
    """Что остаётся при зачистке: {sup: поставщики (sup_norm), name: фрагменты названий (^ = начало названия), bc: ШК}.
    Справочник\\vyvoz_perishable.csv: колонки Тип (поставщик / название / ШК), Значение, Комментарий."""
    df = _csv_default(dirs, SWEEP_STAY_FILE, DEFAULT_SWEEP_STAY, [u"Тип", u"Значение", u"Комментарий"])
    out = {"sup": set(), "name": [], "bc": set()}
    for t, val in zip(df.iloc[:, 0], df.iloc[:, 1]):
        t, val = str(t).strip().lower(), str(val).strip()
        if not val:
            continue
        if t.startswith(u"постав"):
            out["sup"].add(sup_norm(val))
        elif t.startswith(u"назв"):
            out["name"].append(u"^" + nname(val[1:]) if val.startswith(u"^") else nname(val))
        elif t in (u"шк", u"штрих-код", u"штрихкод"):
            k = bc_key(val)
            if k:
                out["bc"].add(k)
    return out


def _sweep_name_hit(nn, rules):
    for w in rules["name"]:
        if w.startswith(u"^"):                     # ^слово = название начинается с этого слова
            t = w[1:]
            if nn == t or nn.startswith(t + u" "):
                return t
        elif w in nn:
            return w
    return u""


def sweep_verdicts(res, rules, ref):
    """Зачистка остатка: вывозим ВСЁ (в матрице, на ЮА, сигареты, алкоголь, кеги, штучные, исключённые поставщики, список
    «не вывозить»), кроме: кулинарии, строк по правилам rules (поставщик / название / ШК), расходников и сырья кофеаппарата,
    овощей и розлива (кеги вывозим), позиций с ШК не 6-14 цифр (в ТСД не загрузить). -> копия res с новыми вердиктами:
    VYVOZ (вывозим), KUL (кулинария) и EXCL (остаётся, причина в колонке «reason»)."""
    out = res.copy()
    verd, reas = list(out["verdict"]), list(out["reason"])
    for i, r in enumerate(out.itertuples(index=False)):
        nn, stay = nname(r.name), u""
        if r.verdict == "KUL":
            verd[i] = "KUL"
            continue
        if not RE_BC_OK.match(str(r.bc)):
            stay = u"некорректный ШК (не 6-14 цифр): в ТСД не загрузить, вывозить вручную"
        elif r.supplier and sup_norm(r.supplier) in rules["sup"]:
            stay = u"поставщик из списка «остаётся»: %s" % r.supplier
        elif _sweep_name_hit(nn, rules):
            stay = u"название из списка «остаётся»: %s" % _sweep_name_hit(nn, rules)
        elif r.key in rules["bc"]:
            stay = u"ШК из списка «остаётся»"
        else:
            why = excl_reason(r.bc, r.key, r.name, ref)
            if why and u"кег" not in why.lower():            # расходники, сырьё кофеаппарата, овощи, розлив; кеги вывозим
                stay = why
        if stay:
            verd[i], reas[i] = "EXCL", u"остаётся (заморозка / скоропорт / расходники): " + stay
        else:
            was = str(r.reason)
            verd[i], reas[i] = "VYVOZ", u"зачистка остатка" + (u" (по обычным правилам: %s)" % was if was else u"")
    out["verdict"], out["reason"] = verd, reas
    return out


def auto_top_n(total, lo=5000.0, hi=7000.0, nmax=None):
    """Сколько точек-получателей, чтобы в среднем на точку пришлось lo..hi грн (ближе к середине диапазона).
    Не меньше 1 и не больше nmax. Распределение по точкам неравномерное (доля по продажам), это именно среднее."""
    nmax = nmax or SPLIT_TOP_N
    if total <= 0:
        return 1
    mid = (lo + hi) / 2.0
    n_lo, n_hi = int(-(-total // hi)), int(total // lo)          # допустимые N: total / N в диапазоне lo..hi
    if n_lo <= n_hi:
        n = min(range(n_lo, n_hi + 1), key=lambda k: abs(total / k - mid))
    else:
        n = max(1, int(round(total / mid)))
    return max(1, min(nmax, n))


def _fail_report(dirs, day):
    """Критическая ошибка: файлы вывоза не пишем, причины - в ЛОГИ\\ВЫВОЗ_СБОЙ_<время>.xlsx"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = u"Проверки"
        write_checks_sheet(ws)
        os.makedirs(dirs.logs, exist_ok=True)
        wb.save(os.path.join(dirs.logs, u"ВЫВОЗ_СБОЙ_%s.xlsx" % datetime.now().strftime("%Y-%m-%d_%H-%M-%S")))
    except Exception:
        pass


# ===================== САМОТЕСТИРОВАНИЕ ======================
# Синтетика: ни BigQuery, ни окон, ни реальных ведомостей не нужно.

def _ean(body):
    """12 цифр -> EAN-13 с верной контрольной цифрой (иначе restore_barcode может дописать ноль)."""
    s = sum(int(ch) * (3 if i % 2 else 1) for i, ch in enumerate(body))
    return body + str((10 - s % 10) % 10)


def self_test(gui=False):
    fails = []

    def ck(cond, msg):
        if not cond:
            fails.append(msg)

    reset_report()

    # ---- ШК и названия приводятся к одному виду ----
    ck(bc_key(19230010208) == "19230010208" and bc_key("019230010208") == "19230010208"
       and bc_key("19230010208.0") == "19230010208", "bc_key: ведущие нули / float")
    ck(nname("Піквік  Фруктово-Трав*яний") == nname("піквік фруктово-трав'яний"),
       "nname: регистр, пробелы, * и апостроф")

    # ---- справочные данные ----
    B_ZERO = "19230010208"                      # в матрице без ведущего нуля
    B_IN, B_IN2, B_NOUA = _ean("482000000001"), _ean("482000000002"), _ean("482000000004")
    B_MNAME = _ean("482000000005")              # в матрице, имя для сопоставления
    B_OLD = _ean("482000000099")                # старый ШК, новый B_IN в матрице
    B_INV_IN, B_INV_OUT = _ean("482000000006"), _ean("600000000009")   # на ЮА только инвентаризацией
    matrix = pd.DataFrame(
        [(B_ZERO, "Роганський МК Сардельки Богатирські", "ok", "Роганський МК"),
         (B_IN, "Пиво Тест 0,5л", "ok", "Пивовар"), (B_IN2, "Сок Тест 1л", "ok", "Союз (Соки)"),
         (B_INV_IN, "Йогурт Тест", "ok", "Молочник"),
         (B_NOUA, "Вода Тест 1,5л", "ok", "Поставщик Тест"), (B_MNAME, "Чай Тест 20п", "Кандидат на вывод", "Чайный дом")]
        + [(_ean("490000000%03d" % i), "Филлер %d" % i, "ok", "Филлер-опт") for i in range(60)],
        columns=["barcode", "product_name", "status", "supplier"])
    recode = pd.DataFrame([(B_OLD, B_IN), ("11111111", "22222222")], columns=["old_barcode", "new_barcode"])
    B_COFFEE, B_COFFEE2 = _ean("871100000001"), _ean("871100000002")
    coffee = pd.DataFrame([(B_COFFEE, "Якобз Кава Тест (1кг)")], columns=["barcode", "product_name"])
    B_UA_ONLY, B_UA_OTHER = _ean("500000000001"), _ean("700000000001")
    ua = pd.DataFrame(
        [(bc_key(B_ZERO), "Роганський МК Сардельки Богатирські", "склад ЮА"),
         (B_IN, "Пиво Тест 0,5л", "склад ЮА"), (B_IN2, "Сок Тест 1л", "склад ЮА"),
         (B_UA_ONLY, "Новинка Тест 100г", "склад ЮА"),
         (B_INV_IN, "Йогурт Тест", "склад ЮА (только инвентаризация)"),
         (B_INV_OUT, "Инв Тест", "склад ЮА (только инвентаризация)"),
         (B_UA_OTHER, "Вафли Тест 100г", "приходы ЮА")]
        + [(_ean("490000000%03d" % i), "Филлер %d" % i, "склад ЮА") for i in range(60)],
        columns=["key", "name", "src"])
    ua["key"] = ua["key"].map(bc_key)
    hist = pd.DataFrame([(B_NOUA, "Прямая", "2026-09-30", "Прямой-поставщик"),
                         (_ean("482000000888"), "РЦ", "2026-09-01", "Сувенир-опт")],
                        columns=["barcode", "types", "last_date", "last_supplier"])
    ref = make_ref(matrix, recode, coffee, ua, hist)

    # ---- проверка матрицы ----
    ok0, _m = validate_matrix(pd.DataFrame(columns=["barcode", "product_name", "status"]))
    ck(not ok0, "пустая матрица должна отвергаться")
    big = pd.DataFrame({"barcode": [str(4820000000000 + i) for i in range(MATRIX_MIN_ROWS + 10)],
                        "product_name": "x", "status": "ok"})
    ck(validate_matrix(big)[0], "нормальная матрица отвергнута")
    ck(not validate_matrix(big.head(50))[0], "урезанная матрица должна отвергаться")
    # матрица из листа Google «Ассортиментная матрица (полная)»: свежее BigQuery (в BigQuery её заливают раз в день)
    sheet_vals = [["Поставщик", "Товар", "Категория", "Решение", "Штрихкод", "Себестоимость", "Прибыль", "Покрытие", "Статус"],
                  ["Бойчак", "Мармелад Тест 1 шт", "Конфеты", "ОСТАВИТЬ", "2978950018438", "26,06", "", "", "🟢 Прибыльный"],
                  ["Бойчак", "ИТОГО Бойчак: 1 SKU", "", "", "", "", "", "", ""],
                  ["Кола", "Швепс Тонік 1л", "Вода", "ОСТАВИТЬ", "5449000044808", "36", "", "", "🟡 Новинка (ещё 90 дн)"],
                  ["Кола", "Без штрихкода", "Вода", "ОСТАВИТЬ", "", "1", "", "", ""],
                  ["", "ОБЩИЙ ИТОГО", "", "", "", "", "", "", ""]]
    ms = parse_matrix_sheet(sheet_vals)
    ck(list(ms.columns) == ["barcode", "product_name", "status", "supplier"] and len(ms) == 2
       and set(ms["barcode"]) == {"2978950018438", "5449000044808"} and ms.loc[ms["barcode"] == "5449000044808", "supplier"].iloc[0] == "Кола",
       "разбор листа матрицы: строки ИТОГО и без ШК отбрасываются: %s" % ms.to_dict("records"))
    big_sheet = pd.concat([big, pd.DataFrame({"barcode": ["5449000044808"], "product_name": ["Швепс"], "status": ["ok"]})], ignore_index=True)
    m_pick, m_note = pick_matrix(big_sheet, big)
    ck(len(m_pick) == len(big) + 1 and "лист" in m_note.lower() and "+1" in m_note, "лист новее BigQuery: берём лист, в примечании разница: %s" % m_note)
    ck(len(pick_matrix(big.head(50), big)[0]) == len(big), "урезанный лист не берём: остаётся BigQuery")
    ck(len(pick_matrix(None, big)[0]) == len(big), "листа нет: BigQuery")
    ck(pick_matrix(None, None)[0] is None, "ни листа, ни BigQuery: матрицы нет")

    # ---- остатки одного магазина ----
    SKU, KG = _ean("482000000888"), _ean("482000000895")
    rows = [
        (19230010208, "Роганський МК Сардельки Богатирські", 3, "кг", 96.9),     # потерянный ноль
        (B_OLD, "Пиво Тест старый ШК", 2, "шт", 20),                             # перекодирован
        (B_IN2, "Сок Тест 1л", 7, "шт", 15),                                      # в матрице и на ЮА
        (_ean("297895009999"), "Овочі/Фрукти Морква 1кг з уцінкою", 1.39, "кг", 10),  # розлив/овощи
        ("2978950020707", "Вода Імператорська 1л", 5, "шт", 8),                  # ШК в списке
        (_ean("297895008888"), "Пакет БОПП 150мм*200мм", 300, "шт", 0.5),        # пакет
        (_ean("297895007777"), "Агропром Стакан Пластик 300 мл 1 шт", 35, "шт", 1),  # стакан
        (rz.BC_KRISHKA, "Кришка біла", 12, "шт", 1),                              # комплектация розлива
        (B_COFFEE, "Якобз Кава Зернова (1кг)", 6.9, "кг", 1400),                  # сырьё по ШК
        (B_COFFEE2, "Якобз  Кава Тест (1кг)", 2.0, "кг", 1400),                   # сырьё по названию
        (_ean("482000000100"), "Нулевой остаток", 0, "шт", 5),
        (_ean("482000000101"), "Минусовой остаток", -2, "шт", 5),
        (B_UA_ONLY, "Новинка Тест 100г", 4, "шт", 12),                            # вне матрицы, на ЮА по ШК
        (B_INV_IN, "Йогурт Тест", 3, "шт", 7),                                    # в матрице, на ЮА только инвентаризацией
        (B_INV_OUT, "Инв Тест", 2, "шт", 5),                                      # вне матрицы, на ЮА только инвентаризацией
        (_ean("700000000099"), "Вафли  Тест 100г", 3, "шт", 12),                  # на ЮА под другим ШК
        (_ean("482000000777"), "Чай Тест 20п", 2, "шт", 9),                       # в матрице под другим ШК
        (B_NOUA, "Вода Тест 1,5л", 5, "шт", 11),                                  # в матрице, на ЮА нет
        (SKU, "Сувенир Тест", 5, "шт", 10),                                        # ВЫВЕЗТИ
        (SKU, "Сувенир Тест", 1, "шт", 10),                                        # дубль строки
        (KG, "Сыр Тест", 0.392, "кг", 100),                                        # ВЫВЕЗТИ, вес
    ] + [(_ean("490000000%03d" % i), "Филлер %d" % i, 5, "шт", 3) for i in range(60)]
    raw = pd.DataFrame(rows, columns=["raw", "name", "qty", "unit", "cost"])
    raw["shop"] = "Тест 1"
    raw["bc"] = [rz.restore_barcode(rz.fmt_barcode(v))[0] for v in raw["raw"]]
    raw["key"] = [bc_key(v) for v in raw["raw"]]
    agg, info = prepare_stock(raw)
    ck(info["merged"] == 1 and info["zero"] == 1 and info["neg"] == 1,
       "prepare_stock: дубли/нули/минусы: %s" % info)
    ck(len(agg) == len(raw) - 3, "prepare_stock: должно остаться позиций %d, а %d" % (len(raw) - 3, len(agg)))
    res = classify_stock(agg, ref)

    def V(bc):
        s = res.loc[res["bc"] == bc, "verdict"]
        return s.iloc[0] if len(s) else None

    def Rn(bc):
        s = res.loc[res["bc"] == bc, "reason"]
        return s.iloc[0] if len(s) else ""

    ck(V("019230010208") == "IN", "потерянный ноль: %s" % V("019230010208"))
    ck(V(B_OLD) == "IN" and bool(res.loc[res["bc"] == B_OLD, "recoded"].iloc[0]),
       "перекодированный ШК не должен быть вывозом")
    ck(not res.loc[res["bc"] == B_OLD, "out_m"].iloc[0], "перекодированный: out_m")
    ck(V(B_IN2) == "IN", "в матрице и на ЮА")
    ck(V(_ean("297895009999")) == "EXCL" and "овочі" in Rn(_ean("297895009999")), "овощи/развеска")
    ck(V("2978950020707") == "EXCL", "ШК из списка исключений")
    ck(V(_ean("297895008888")) == "EXCL" and Rn(_ean("297895008888")) == "пакет", "пакет")
    ck(V(_ean("297895007777")) == "EXCL" and Rn(_ean("297895007777")) == "стакан", "стакан")
    ck(V(rz.BC_KRISHKA) == "EXCL" and "розлив" in Rn(rz.BC_KRISHKA), "комплектация розлива")
    ck(V(B_COFFEE) == "EXCL" and "кофеаппарат" in Rn(B_COFFEE), "сырьё кофеаппарата по ШК")
    ck(V(B_COFFEE2) == "EXCL" and "кофеаппарат" in Rn(B_COFFEE2), "сырьё кофеаппарата по названию")
    ck(V(_ean("482000000100")) is None and V(_ean("482000000101")) is None, "нулевой и минусовой остаток не в списках")
    ck(V(B_UA_ONLY) == "CHECK1" and "ЮА" in Rn(B_UA_ONLY), "вне матрицы, но на ЮА по ШК")
    ck(V(_ean("700000000099")) == "CHECK1", "вне матрицы, на ЮА под другим ШК (по названию)")
    ck(V(B_INV_IN) == "IN" and "инвентаризац" in Rn(B_INV_IN), "в матрице, на ЮА только инвентаризацией: пометка")
    ck(V(B_INV_OUT) == "CHECK1" and "инвентаризац" in Rn(B_INV_OUT), "вне матрицы, на ЮА только инвентаризацией: пометка")
    ck(shop_stats(res)["inv_only"] == 2, "счётчик «только инвентаризация»: %s" % shop_stats(res)["inv_only"])
    ck(V(_ean("482000000777")) == "CHECK2" and not res.loc[res["bc"] == _ean("482000000777"), "out_m"].iloc[0]
       and res.loc[res["bc"] == _ean("482000000777"), "other_key"].iloc[0] == bc_key(B_MNAME)
       and res.loc[res["bc"] == _ean("482000000777"), "supplier"].iloc[0] == "Чайный дом",
       "в матрице под другим ШК (по названию): объединено с матричной позицией, на ЮА нет: %s" % V(_ean("482000000777")))
    ck(V(B_NOUA) == "CHECK2", "в матрице, на ЮА нет: %s" % V(B_NOUA))
    ck("Прямая" in str(res.loc[res["bc"] == B_NOUA, "hist"].iloc[0]), "история поставок в списке «В матрице, на ЮА не было»")
    ck(res.loc[res["bc"] == B_NOUA, "supplier"].iloc[0] == "Поставщик Тест", "поставщик из матрицы в строке")
    ck(bool(res.loc[res["bc"] == B_NOUA, "not_on_ua"].iloc[0]) and int(res["not_on_ua"].sum()) == 2,
       "«В матрице, на ЮА не было»: две позиции (в т.ч. объединённая по названию): %d" % int(res["not_on_ua"].sum()))
    sup_t = supplier_table(res)
    ck(len(sup_t) == 2 and sup_t["supplier"].iloc[0] == "Поставщик Тест" and sup_t["pos"].iloc[0] == 1
       and sup_t["direct"].iloc[0] == 1 and sup_t["rc"].iloc[0] == 0 and sup_t["none"].iloc[0] == 0,
       "таблица по поставщикам: %s" % sup_t.to_dict("records"))
    ck(V(SKU) == "VYVOZ" and abs(res.loc[res["bc"] == SKU, "qty"].iloc[0] - 6) < 1e-9, "вывоз + сумма дублей")
    ck(V(KG) == "VYVOZ", "вывоз весового")
    ck(int((res["verdict"] == "VYVOZ").sum()) == 2, "вывезти должно быть ровно 2 позиции: %d" % int((res["verdict"] == "VYVOZ").sum()))
    ck(res.loc[res["bc"] == SKU, "supplier"].iloc[0] == "Сувенир-опт"
       and res.loc[res["bc"] == SKU, "supplier_src"].iloc[0] == "последний приход",
       "вне матрицы: поставщик последнего прихода из базы: %s" % res.loc[res["bc"] == SKU, ["supplier", "supplier_src"]].to_dict("records"))
    ck(res.loc[res["bc"] == B_NOUA, "supplier_src"].iloc[0] == "матрица", "в матрице: поставщик из матрицы")
    # слияние позиций: старый и новый ШК одного товара (перекодировка) - одна строка, количества складываются
    mg_in = pd.DataFrame({"shop": ["Тест 1"] * 3, "bc": [B_OLD, B_IN, B_IN2],
                          "key": [bc_key(B_OLD), bc_key(B_IN), bc_key(B_IN2)],
                          "name": ["Пиво Тест старый ШК", "Пиво Тест 0,5л", "Сок Тест 1л"], "unit": ["шт"] * 3,
                          "cost": [20.0, 20.0, 15.0], "qty": [2.0, 3.0, 7.0]})
    mg = merge_recoded_duplicates(mg_in, ref)
    ck(len(mg) == 2 and abs(float(mg.loc[mg["key"] == bc_key(B_IN), "qty"].iloc[0]) - 5.0) < 1e-9
       and mg.loc[mg["key"] == bc_key(B_IN), "bc"].iloc[0] == B_IN and bc_key(B_OLD) not in set(mg["key"]),
       "слияние старого и нового ШК: %s" % mg.to_dict("records"))
    ck(len(merge_recoded_duplicates(mg_in.iloc[[0, 2]], ref)) == 2, "слияние: без пары ничего не меняется")
    # сигареты: на ЮА латиницей и другим ШК, у Family кириллицей
    ua_c = pd.DataFrame([("48211260", "Сигарети Sobranie Gold", "склад ЮА"), ("48211253", "Сигарети Sobranie Blue", "склад ЮА"),
                         ("4820270365687", "ТВЕН NEO DEMI RUBY BOOST", "склад ЮА")], columns=["key", "name", "src"])
    ref_c = make_ref(matrix, recode, coffee, ua_c, hist)
    pk, pn, ps, am = cig_pair(ref_c, bc_key("4820000535182"), "Сигарети Собраніе Голд 20шт")
    ck(pk == "48211260" and ps >= 0.9 and not am, "сигареты: Собраніе Голд -> Sobranie Gold: %s %s %s" % (pk, ps, am))
    ua_y = pd.DataFrame([("4820270365656", "ТВЕН NEO DEMI YELLOW BOOST", "склад ЮА"), ("4820270364291", 'ТВЕН "NEO DEMI Royale Boost"', "склад ЮА")],
                        columns=["key", "name", "src"])
    ck(cig_pair(make_ref(matrix, recode, coffee, ua_y, hist), bc_key("4820215629393"), "Сигарети Нео Демі Єлоу Буст 20шт")[0] == "4820270365656",
       "сигареты: Єлоу = Yellow, а не Royale")
    ua_amb = pd.DataFrame([("48211260", "Сигарети Sobranie Gold", "склад ЮА"), ("48211999", "Сигарети Собраніе Голд", "склад ЮА")],
                          columns=["key", "name", "src"])
    ck(cig_pair(make_ref(matrix, recode, coffee, ua_amb, hist), bc_key("4820000535182"), "Сигарети Собраніе Голд 20шт")[3] is True,
       "сигареты: на ЮА два одинаково подходящих ШК - пара неоднозначна")
    ck(cig_pair(ref_c, bc_key("4820000535168"), "Сигарети Собраніе Блю 20шт")[0] == "48211253", "сигареты: Блю -> Blue")
    ck(cig_pair(ref_c, bc_key("4820215626682"), "Сигарети Нео Демі Рубі Буст 20шт")[0] == "4820270365687", "сигареты: Нео Демі Рубі -> NEO DEMI RUBY")
    ck(cig_pair(ref_c, bc_key("482000000999"), "Сувенир Тест")[0] == "", "не сигареты: пары нет")
    ck(cig_pair(ref_c, bc_key("48211260"), "Сигарети Sobranie Gold")[0] == "", "тот же ШК, что на ЮА: пара не нужна")
    # сигареты не вывозятся вообще (даже при настройке «убирать»): остаются на полке и идут в инвентаризацию;
    # если на ЮА та же сигарета под другим ШК и названием латиницей - в файл для ЮА пишется ШК ЮА
    CG_OUT, CG_IN, SKU2 = _ean("482000053518"), _ean("482000053516"), _ean("482000099991")
    K_OUT, K_IN = _ean("298467006001"), _ean("298467006002")      # кулинария: вне матрицы / в матрице, на ЮА не было
    BAG_IN, CUP_IN = "2978950015505", "2978950015550"          # пакет и стакан В МАТРИЦЕ, на ЮА не было: расходники не вывозим
    ua_cg = pd.DataFrame([("48211260", "Сигарети Sobranie Gold", "склад ЮА", "48211260"),
                          ("48211253", "Сигарети Sobranie Blue", "склад ЮА", "48211253"),
                          ("4820000777771", "Чай Тест 20п", "склад ЮА", "4820000777771"),
                          ("4820270931615", "Монжар Драже Тубус Веселка (24) Упаковка 24 шт", "склад ЮА", "4820270931615"),
                          ("4000512992622", "Бойчак  Мармелад Жув. Троллі Глотзер", "склад ЮА", "4000512992622")],
                         columns=["key", "name", "src", "bc"])
    matrix_cg = pd.concat([matrix, pd.DataFrame([(CG_IN, "Сигарети Собраніе Блю 20шт", "ok", "BAT"),
                                                  (K_IN, "Кулінарія Борщ український 0,450гр", "ok", "Кухня"),
                                                  (BAG_IN, "Пакет Майка 100 шт 1 шт", "ok", "Пакети"),
                                                  (CUP_IN, "Агропром Стакан Паперовий Малюнок 175 мл 1 шт", "ok", "Агропром")],
                                                 columns=matrix.columns)], ignore_index=True)
    ref_cg = make_ref(matrix_cg, recode, coffee, ua_cg, hist, pairs={"2978950018438": "4000512992622"})
    PK_AUTO, PK_MAN, PK_NO = "4820212960352", "2978950018438", _ean("298467006009")      # штучные: упаковка на ЮА / ручная пара / без пары
    agg_cg = pd.DataFrame({"shop": ["Тест 1"] * 11, "bc": [CG_OUT, CG_IN, SKU2, B_MNAME, K_OUT, K_IN, BAG_IN, CUP_IN, PK_AUTO, PK_MAN, PK_NO],
                           "key": [bc_key(x) for x in (CG_OUT, CG_IN, SKU2, B_MNAME, K_OUT, K_IN, BAG_IN, CUP_IN, PK_AUTO, PK_MAN, PK_NO)],
                           "name": ["Сигарети Собраніе Голд 20шт", "Сигарети Собраніе Блю 20шт", "Сувенир Тест 2", "Чай Тест 20п",
                                    "Випічка Хліб Козацький з часником 500г", "Кулінарія Борщ український 0,450гр",
                                    "Пакет Майка 100 шт 1 шт", "Агропром Стакан Паперовий Малюнок 175 мл 1 шт",
                                    "Монжар Драже Тубус Веселка (24) 1 шт", "Бойчак  Мармелад Жув. Троллі  Глаз 1 шт", "Сувенир Штучный 1 шт"],
                           "unit": ["пачка", "пачка"] + ["шт"] * 9,
                           "cost": [100.0, 100.0, 10.0, 9.0, 13.0, 24.0, 0.26, 0.6, 9.5, 26.06, 5.0],
                           "qty": [7.0, 8.0, 3.0, 2.0, 3.0, 4.0, 684.0, 45.0, 2.0, 52.0, 6.0]})
    res_cg = classify_stock(agg_cg, ref_cg, remove_not_on_ua=True)

    def G(bc, col):
        return res_cg.loc[res_cg["bc"] == bc, col].iloc[0]
    ck(G(CG_OUT, "verdict") == "EXCL" and "сигарет" in G(CG_OUT, "reason"), "сигарета вне матрицы не вывозится: %s %s" % (G(CG_OUT, "verdict"), G(CG_OUT, "reason")))
    ck(G(CG_IN, "verdict") == "EXCL" and bool(G(CG_IN, "not_on_ua")),
       "сигарета «в матрице, на ЮА не было» при настройке «убирать» не вывозится, но остаётся в списке закупщицы: %s" % G(CG_IN, "verdict"))
    ck(G(SKU2, "verdict") == "VYVOZ", "обычный товар вне матрицы по-прежнему вывозится")
    ck(G(CG_OUT, "inv_bc") == "48211260", "сигарета с парой на ЮА: в файл инвентаризации идёт ШК ЮА: %r" % G(CG_OUT, "inv_bc"))
    ck(G(B_MNAME, "verdict") == "IN_UA_NAME" and G(B_MNAME, "inv_bc") == "4820000777771",
       "в матрице, на ЮА под другим ШК: в файл идёт ШК ЮА: %s %r" % (G(B_MNAME, "verdict"), G(B_MNAME, "inv_bc")))
    ck(G(SKU2, "inv_bc") == SKU2, "обычная позиция: ШК файла = её ШК")
    # расходники (пакеты, стаканы, крышки/пляшки розлива, сырьё кофеаппарата, кеги, овощи) не вывозятся и когда они В МАТРИЦЕ,
    # а на ЮА не было (раньше правило работало только для позиций вне матрицы)
    ck(G(BAG_IN, "verdict") == "EXCL" and G(BAG_IN, "reason") == "пакет" and bool(G(BAG_IN, "not_on_ua")),
       "пакет в матрице, на ЮА не было: не вывозится: %s %s" % (G(BAG_IN, "verdict"), G(BAG_IN, "reason")))
    ck(G(CUP_IN, "verdict") == "EXCL" and G(CUP_IN, "reason") == "стакан", "стакан в матрице, на ЮА не было: не вывозится: %s" % G(CUP_IN, "verdict"))
    # кулинария (выпечка) не вывозится и на ЮА не переходит: ни в вывозе, ни в файле инвентаризации
    ck(G(K_OUT, "verdict") == "KUL" and G(K_IN, "verdict") == "KUL" and bool(G(K_IN, "not_on_ua")),
       "кулинария: не вывозим, даже при настройке «убирать»: %s %s" % (G(K_OUT, "verdict"), G(K_IN, "verdict")))
    ck(K_OUT not in set(inventory_rows(res_cg)["bc"]) and K_IN not in set(inventory_rows(res_cg)["bc"]),
       "кулинария не попадает в файл инвентаризации ЮА")
    pc_cg = partition_check(res_cg)
    ck(pc_cg["ok"] and pc_cg["kul_pos"] == 4 and abs(pc_cg["vyvoz_qty"] + pc_cg["inv_qty"] + pc_cg["wait_qty"] + pc_cg["kul_qty"]
                                                     - float(res_cg["qty"].sum())) < 1e-6,
       "баланс с кулинарией и штучными: вывезти + перевести + ждёт + остаётся без инвентаризации = остаток: %s" % pc_cg)
    # штучный товар («1 шт»), упаковка которого уже есть на ЮА под другим ШК: не вывозим, в инвентаризацию не включаем
    ck(pack_base("Монжар Драже Тубус Веселка (24) 1 шт") == pack_base("Монжар Драже Тубус Веселка (24) Упаковка 24 шт"),
       "штучный и упаковка сводятся к одному названию")
    ck(G(PK_AUTO, "verdict") == "PACK" and G(PK_AUTO, "other_key") == "4820270931615",
       "штучный товар, упаковка на ЮА (по названию): не вывозим: %s %s" % (G(PK_AUTO, "verdict"), G(PK_AUTO, "other_key")))
    ck(G(PK_MAN, "verdict") == "PACK" and G(PK_MAN, "other_key") == "4000512992622",
       "штучный товар по ручной паре ШК (таблица пар): не вывозим: %s" % G(PK_MAN, "verdict"))
    ck(G(PK_NO, "verdict") == "VYVOZ", "штучный товар без упаковки на ЮА вывозится как раньше")
    ck(PK_AUTO not in set(inventory_rows(res_cg)["bc"]) and PK_MAN not in set(inventory_rows(res_cg)["bc"]),
       "штучные не попадают в файл инвентаризации ЮА автоматически (единицы разные: штука и упаковка)")
    # ---- исключение поставщиков из вывоза (файл Справочник\\vyvoz_suppliers.csv, окно с галочками) ----
    ck(sup_norm("Прем*єр Фуд") == sup_norm("Премьер Фуд") and sup_norm("Хладік") == sup_norm("Хладик"),
       "названия поставщиков сравниваются без і/и, ь и «*»")
    tsup = tempfile.mkdtemp(prefix="vyvoz_sup_")
    try:
        dsup = Dirs(tsup)
        dsup.ensure()
        rules = load_excluded_suppliers(dsup)                  # файла нет: создаётся с поставщиками по умолчанию
        ck(os.path.isfile(os.path.join(dsup.cache, SUPPLIERS_FILE)), "файл поставщиков создан")
        for nm_ in (u"УДК", u"Хладік", u"Хладопром", u"Прем*єр Фуд", u"Фаст Фуд", u"Овочі-Фрукти (Протопопов)", u"ФОП Протопопов"):
            ck(sup_norm(nm_) in rules, "по умолчанию исключён: %s" % nm_)
        ck(sup_norm("Монжар") not in rules, "прочие поставщики вывозятся")
        rules2 = dict(rules)
        rules2.pop(sup_norm("УДК"))
        rules2[sup_norm("Тест-пост")] = u"Тест-пост"
        save_excluded_suppliers(dsup, rules2)
        rules3 = load_excluded_suppliers(dsup)                 # файл есть: заново по умолчанию не создаётся
        ck(sup_norm("УДК") not in rules3 and sup_norm("Тест-пост") in rules3, "сохранение: УДК снят, Тест-пост добавлен")
        try:                                                  # окно с галочками собирается и закрывается (если на машине есть Tk)
            import tkinter as tk_
            root_ = tk_.Tk()
            root_.withdraw()
            try:
                rows_, off_ = open_suppliers_dialog(root_, dsup, None, test_mode=True)
                ck(rows_ == 7 and off_ == 7, "окно поставщиков: строк %d, исключено %d (ожидалось 7 и 7)" % (rows_, off_))
                # галочка = НЕ вывозим: поставили галочку на «Шейк», у «Моршин» её нет -> в файле только Шейк
                pd.DataFrame({"supplier": [u"Тест Моршин", u"Тест Шейк"]}).to_csv(
                    os.path.join(dsup.cache, "vyvoz_matrix.csv"), index=False, encoding="utf-8-sig")
                seen = {}

                def hook1(h):
                    seen["free"] = sorted(n for n, v_ in h["on"].items() if not v_.get())
                    h["on"][sup_norm(u"Тест Шейк")].set(True)
                    h["save"]()
                open_suppliers_dialog(root_, dsup, None, test_mode=True, test_hook=hook1)
                rl = load_excluded_suppliers(dsup)
                ck(seen["free"] == sorted([sup_norm(u"Тест Моршин"), sup_norm(u"Тест Шейк")]) and sup_norm(u"Тест Шейк") in rl
                   and sup_norm(u"Тест Моршин") not in rl and len(rl) == 8,
                   "окно: без галочки - вывозим, с галочкой - не вывозим (файл: %d, Шейк в нём, Моршина нет): %s" % (len(rl), sorted(rl)))
                # «Очистить: вывозим всех» + «Сохранить» -> файл реально пуст, новое окно открывается без галочек
                open_suppliers_dialog(root_, dsup, None, test_mode=True, test_hook=lambda h: (h["set_vis"](False), h["save"]()))
                ck(load_excluded_suppliers(dsup) == {}, "«Очистить» сохраняет пустой список: %s" % sorted(load_excluded_suppliers(dsup)))
                seen2 = {}
                open_suppliers_dialog(root_, dsup, None, test_mode=True,
                                      test_hook=lambda h: seen2.update(n=sum(1 for v_ in h["on"].values() if v_.get())))
                ck(seen2.get("n") == 0, "после очистки окно открывается без единой галочки: %s" % seen2)
            finally:
                root_.destroy()
        except Exception as e_:
            R.check("INFO", u"Окно поставщиков в самотесте", u"не проверено: %s" % e_)
        n_f, when_f, path_f = suppliers_file_info(dsup)
        ck(n_f == len(load_excluded_suppliers(dsup)) and when_f and path_f.endswith(SUPPLIERS_FILE), "сведения о файле поставщиков: %s %s" % (n_f, when_f))
        ck(sup_norm("Арсенал ПК (Шейк)") == sup_norm("Арсенал ПК( Шейк)") == sup_norm("арсенал  пк ( шейк )"),
           "пробелы у скобок не создают второго поставщика")
    finally:
        shutil.rmtree(tsup, ignore_errors=True)
    tpair = tempfile.mkdtemp(prefix="vyvoz_pairs_")
    try:
        dpair = Dirs(tpair)
        dpair.ensure()
        pairs_ = load_shk_pairs(dpair)                           # файла нет: создаётся с известными парами
        ck(os.path.isfile(os.path.join(dpair.cache, SHK_PAIRS_FILE)) and pairs_.get("2978950018438") == "4000512992622",
           "таблица пар ШК Family-ЮА создана с парой Глаз = Глотзер: %s" % pairs_)
        pd.DataFrame({u"ШК Family": ["111111111111"], u"ШК ЮА": ["222222222222"], u"Комментарий": ["тест"]}).to_csv(
            os.path.join(dpair.cache, SHK_PAIRS_FILE), index=False, encoding="utf-8-sig")
        ck(load_shk_pairs(dpair) == {"111111111111": "222222222222"}, "таблица пар читается, ведущие нули не нужны")
    finally:
        shutil.rmtree(tpair, ignore_errors=True)
    # ---- список ШК «не вывозить» (Справочник\\vyvoz_keep.csv): позиции, которые пользователь велел оставить ----
    tkeep = tempfile.mkdtemp(prefix="vyvoz_keep_")
    try:
        dkeep = Dirs(tkeep)
        dkeep.ensure()
        keep0 = load_keep_shk(dkeep)                             # файла нет: создаётся с Сірники (1)
        ck(os.path.isfile(os.path.join(dkeep.cache, KEEP_FILE)) and bc_key("4820116280075") in keep0, "список «не вывозить» создан: %s" % keep0)
        pd.DataFrame({u"ШК": ["0123456789012", "bad"], u"Название": ["x", "y"], u"Комментарий": ["", ""]}).to_csv(
            os.path.join(dkeep.cache, KEEP_FILE), index=False, encoding="utf-8-sig")
        ck(load_keep_shk(dkeep) == {bc_key("0123456789012")}, "список «не вывозить»: ведущие нули не нужны, мусор отбрасывается")
        with open(os.path.join(dkeep.cache, KEEP_FILE), "w", encoding="utf-8-sig", newline="") as f_:     # так файл выглядит после сохранения из Excel
            f_.write('ШК,Название,Комментарий\r\n"8711000605561,""Кава, тест"",""Решение"""\r\n"4820206290519,Кава 2,""Решение"""\r\n')
        ck(load_keep_shk(dkeep) == {bc_key("8711000605561"), bc_key("4820206290519")},
           "список «не вывозить»: строка, целиком попавшая в первую ячейку (Excel), читается: %s" % sorted(load_keep_shk(dkeep)))
        with open(os.path.join(dkeep.cache, KEEP_FILE), "w", encoding="utf-8-sig", newline="") as f_:
            f_.write("ШК,Название\r\nмусор,x\r\n")
        reset_report()
        got_ = load_keep_shk(dkeep)
        ck(not got_ and any(nm == u"Список ШК «не вывозить»" and st_ == "WARN" for st_, nm, _d in R.checks),
           "список «не вывозить»: ни один ШК не прочитан - предупреждение")
    finally:
        shutil.rmtree(tkeep, ignore_errors=True)
    res_k = classify_stock(agg, ref, remove_not_on_ua=True, keep_keys={bc_key(SKU), bc_key(B_NOUA)})
    ck(res_k.loc[res_k["bc"] == SKU, "verdict"].iloc[0] == "EXCL" and "не вывозить" in res_k.loc[res_k["bc"] == SKU, "reason"].iloc[0],
       "ШК из списка «не вывозить» не вывозится: %s" % res_k.loc[res_k["bc"] == SKU, "verdict"].iloc[0])
    ck(res_k.loc[res_k["bc"] == B_NOUA, "verdict"].iloc[0] == "EXCL" and bool(res_k.loc[res_k["bc"] == B_NOUA, "not_on_ua"].iloc[0]),
       "позиция «в матрице, на ЮА не было» из списка «не вывозить» остаётся (и в списке закупщицы тоже)")
    ck(res_k.loc[res_k["bc"] == KG, "verdict"].iloc[0] == "VYVOZ", "прочие вывозятся как раньше")
    res_sx = classify_stock(agg, ref, remove_not_on_ua=True,
                            excl_suppliers={sup_norm("Сувенир-опт"), sup_norm("Поставщик Тест")})

    def VX(bc, col="verdict"):
        return res_sx.loc[res_sx["bc"] == bc, col].iloc[0]
    ck(VX(SKU) == "EXCL" and "поставщик" in VX(SKU, "reason"), "исключённый поставщик (последний приход): позиция не вывозится: %s" % VX(SKU))
    ck(VX(B_NOUA) == "EXCL" and bool(VX(B_NOUA, "not_on_ua")), "исключённый поставщик матрицы: не вывозится даже при «убирать»: %s" % VX(B_NOUA))
    ck(VX(KG) == "VYVOZ", "позиция без поставщика вывозится как раньше")
    ck(SKU in set(inventory_rows(res_sx)["bc"]), "исключённая из вывоза позиция остаётся на полке и идёт в инвентаризацию")
    res_kc = classify_stock(agg, ref, keep_keys={bc_key(B_UA_ONLY)})
    ck(res_kc.loc[res_kc["bc"] == B_UA_ONLY, "verdict"].iloc[0] == "EXCL" and B_UA_ONLY in set(inventory_rows(res_kc)["bc"]),
       "«ждёт решения» (заведён на ЮА по ШК) + ШК в vyvoz_keep.csv: остаётся и идёт в инвентаризацию")
    pa = _parse_args(["--exclude-suppliers", "УДК;Тест", "--include-suppliers", "Хладік"])
    ck(pa["excl"] == [u"УДК", u"Тест"] and pa["incl"] == [u"Хладік"], "аргументы поставщиков: %s %s" % (pa["excl"], pa["incl"]))
    # настройка «убирать и то, что в матрице, но на ЮА не завозилось»
    res_rm = classify_stock(agg, ref, remove_not_on_ua=True)
    ck(res_rm.loc[res_rm["bc"] == B_NOUA, "verdict"].iloc[0] == "VYVOZ"
       and "не завозилось" in res_rm.loc[res_rm["bc"] == B_NOUA, "reason"].iloc[0], "флаг: в матрице, на ЮА не завозилось -> вывезти")
    ck(int((res_rm["verdict"] == "VYVOZ").sum()) == 4 and int((res_rm["verdict"] == "CHECK2").sum()) == 0,
       "флаг: вывозов должно стать 4, проверить 2 - 0")
    ck(not res_rm.loc[res_rm["bc"] == B_NOUA, "out_m"].iloc[0], "флаг не меняет «вне матрицы» для порога 40%")
    ck(bool(res_rm.loc[res_rm["bc"] == B_NOUA, "not_on_ua"].iloc[0]) and len(supplier_table(res_rm)) == 2,
       "флаг: позиция остаётся в списке для закупщицы (по поставщикам)")
    # что остаётся на полке -> инвентаризация ЮА; вместе с вывозом и «ждёт решения» это весь остаток
    inv_ = inventory_rows(res)
    ck(set(inv_["verdict"]) <= {"IN", "IN_UA_NAME", "EXCL"} and len(inv_) == int(res["verdict"].isin(["IN", "IN_UA_NAME", "EXCL"]).sum()),
       "inventory_rows: только то, что остаётся на полке")
    ck(V("019230010208") == "IN" and "019230010208" in set(inv_["bc"]) and SKU not in set(inv_["bc"]),
       "в инвентаризацию попадает остающийся товар, вывоз - нет")
    pc = partition_check(res)
    ck(pc["ok"] and abs(pc["vyvoz_qty"] + pc["inv_qty"] + pc["wait_qty"] - float(res["qty"].sum())) < 1e-6
       and abs(pc["vyvoz_sum"] + pc["inv_sum"] + pc["wait_sum"] - float(res["sum"].sum())) < 0.01,
       "баланс остатка: вывезти + перевести + ждёт решения = остаток: %s" % pc)
    ck(ua_name("Грозненська 38") == "Болградська 38" and ua_name("Байрона 156") == "Байрона 156", "имя магазина в базе ЮА")
    # режим без ЮА: проверок по ЮА нет, остальное работает
    res2 = classify_stock(agg, ref, use_ua=False)
    ck(int((res2["verdict"] == "CHECK2").sum()) == 0 and V(B_UA_ONLY) == "CHECK1"
       and res2.loc[res2["bc"] == B_UA_ONLY, "verdict"].iloc[0] == "VYVOZ", "режим без ЮА")

    st = shop_stats(res)
    ck(st["out_m"] > 0 and abs(st["out_share"] - st["out_m"] / float(st["positions"])) < 1e-9, "доля вне матрицы")
    ck(not guard_failed(st)[0], "нормальный магазин не должен падать по порогу 40%%: %.1f%%" % (100 * st["out_share"]))
    # магазин, где вне матрицы больше 40%
    bad = pd.DataFrame([("Тест 2", "x", _ean("5%011d" % i), str(i), "Нет в матрице %d" % i, 1.0, "шт", 1.0)
                        for i in range(60)] +
                       [("Тест 2", "x", _ean("490000000%03d" % i), str(i), "Филлер %d" % i, 1.0, "шт", 1.0)
                        for i in range(20)],
                       columns=["shop", "raw", "bc", "key", "name", "qty", "unit", "cost"])
    bad["key"] = bad["bc"].map(bc_key)
    agg2, _i = prepare_stock(bad)
    st2 = shop_stats(classify_stock(agg2, ref))
    ck(guard_failed(st2)[0] and st2["out_share"] > OUT_SHARE_ERROR, "порог 40%%: %.2f" % st2["out_share"])

    # ---- выбор магазинов ----
    names = ["Грозненська 38", "Байрона 156", "Полевая-Магазин"]
    got, errs = resolve_shops(["Байрона 156", "грозненська 38"], names)
    ck(got == ["Байрона 156", "Грозненська 38"] and not errs, "resolve_shops: %s %s" % (got, errs))
    got, errs = resolve_shops(["Болградська 38"], names)
    ck(got == ["Грозненська 38"] and not errs, "переименование Болградська 38 = Грозненська 38: %s %s" % (got, errs))
    got, errs = resolve_shops(["Грозненська 38"], ["Болградська 38", "Байрона 156"])
    ck(got == ["Болградська 38"] and not errs, "и наоборот, если в выгрузке уже новое имя: %s %s" % (got, errs))
    ck(name_collisions(["Грозненська 38", "Болградська 38", "Байрона 156"]) == [("Болградська 38", "Грозненська 38")],
       "одно имя магазина в двух выгрузках (до и после переименования) должно ловиться: %s"
       % name_collisions(["Грозненська 38", "Болградська 38", "Байрона 156"]))
    ck(name_collisions(["Грозненська 38", "Байрона 156"]) == [], "без переименований коллизий нет")

    # ---- где лежат входы и результаты ----
    t0 = tempfile.mkdtemp(prefix="vyvoz_base_")
    try:
        base = os.path.join(t0, "РАСЧЕТ ЗАКАЗА")
        os.makedirs(base)
        ck(data_base(base, "") == base, "нет папки «Перемещения ЮА» -> рядом с программой")
        os.makedirs(os.path.join(t0, "Перемещения ЮА"))
        ck(data_base(base, "") == os.path.join(t0, "Перемещения ЮА"), "есть папка «Перемещения ЮА» -> данные там")
        ck(data_base(base, os.path.join(t0, "другая")) == os.path.join(t0, "другая"), "FM_VYVOZ_DIR важнее")
    finally:
        shutil.rmtree(t0, ignore_errors=True)
    got, errs = resolve_shops(["Полевая-Магазин"], names)
    ck(not got and errs, "Полевая-Магазин не вывозим")
    got, errs = resolve_shops(["Грозненская 38"], names)
    ck(not got and errs and "Грозненська 38" in errs[0], "неизвестный магазин: подсказка")

    # ---- сквозной прогон: xlsx -> списки -> файлы -> обратная сверка ----
    tmp = tempfile.mkdtemp(prefix="vyvoz_test_")
    try:
        d = Dirs(tmp)
        d.ensure()
        _write_state_xlsx(os.path.join(d.stores, "Состояние склада Тест 1 на 02.10.2026.xlsx"), "Тест 1",
                          [(r[0], r[1], r[2], r[3], r[4]) for r in rows])
        _write_state_xlsx(os.path.join(d.stores, "Состояние склада Тест 2 на 02.10.2026.xlsx"), "Тест 2",
                          [(r["bc"], r["name"], r["qty"], r["unit"], r["cost"]) for _, r in bad.iterrows()])
        _write_state_xlsx(os.path.join(d.ua_wh, "Состояние склада Полевая склад ЮА 02.10.2026.xlsx"),
                          "Полевая-Склад ЮА",
                          [(k, n, 1, "шт", 1.0) for k, n, s in ua[ua["src"] == "склад ЮА"].itertuples(index=False)]
                          + [(B_INV_OUT, "Инв Тест", 1, "шт", 1.0), (B_INV_IN, "Йогурт Тест", 1, "шт", 1.0),
                             (B_UA_OTHER, "Вафли Тест 100г", 1, "шт", 1.0)],
                          shop_name="Полевая-Склад ЮА")
        _write_receipts_xlsx(os.path.join(d.ua_receipts, "Приходы_ЮА_Маркет.xlsx"),
                             [(B_INV_OUT, "Инв Тест", "ПРИХОД ОТ ИНВЕНТАРИЗАЦИИ", 5),
                              (B_INV_IN, "Йогурт Тест", "ПРИХОД ОТ ИНВЕНТАРИЗАЦИИ", 2),
                              (_ean("600000000002"), "Нулевой приход Тест", "Бір Компанія", 0),
                              (B_IN, "Пиво Тест 0,5л", "Бір Компанія", 3)])
        reset_report()
        ua_df, uinfo = load_ua(d)
        srcmap = dict(ua_df.drop_duplicates("key")[["key", "src"]].itertuples(index=False))
        ck(bc_key(B_INV_OUT) in srcmap and "инвентаризац" in srcmap[bc_key(B_INV_OUT)],
           "приход от инвентаризации входит в реестр ЮА с пометкой: %s" % srcmap.get(bc_key(B_INV_OUT)))
        ck("поставщик" in srcmap.get(bc_key(B_IN), ""), "приход поставщика помечен: %s" % srcmap.get(bc_key(B_IN)))
        ck(bc_key(_ean("600000000002")) not in srcmap, "нулевой приход не входит в реестр ЮА")
        ck(uinfo["inv_only"] == 2, "ШК склада ЮА только по инвентаризации: %s" % uinfo["inv_only"])
        ck(any(nm.startswith("Инвентаризация ЮА") for _s, nm, _d in R.checks), "предупреждение «инвентаризация под вопросом»")
        globals()["UA_COUNT_INVENTORY"] = False
        try:
            ua_off, _i = load_ua(d)
        finally:
            globals()["UA_COUNT_INVENTORY"] = True
        ck(bc_key(B_INV_OUT) not in set(ua_off["key"]) and bc_key(B_IN) in set(ua_off["key"]),
           "UA_COUNT_INVENTORY=False: инвентаризация не считается заведённой, поставщик остаётся")
        ref2 = make_ref(matrix, recode, coffee, ua_df, hist)
        out = run_vyvoz(d, ref=ref2)
        shops = {s["shop"]: s for s in out["shops"]}
        ck(shops["Тест 1"]["status"] == "OK", "Тест 1: статус %s" % shops["Тест 1"]["status"])
        ck(shops["Тест 2"]["status"].startswith("ОШИБ"), "Тест 2: должен быть ОШИБ, а %s" % shops["Тест 2"]["status"])
        p1 = os.path.join(out["day_dir"], "Тест 1")
        ck(os.path.isfile(os.path.join(p1, "Тест 1.xlsx")) and os.path.isfile(os.path.join(p1, "Тест 1.txt")),
           "файлы магазина созданы")
        ck(not os.path.exists(os.path.join(out["day_dir"], "Тест 2", "Тест 2.txt")), "по магазину с ОШИБ txt не пишется")
        ck(os.path.isfile(os.path.join(out["day_dir"], SUMMARY_NAME)), "сводка создана")
        with open(os.path.join(p1, "Тест 1.txt"), "rb") as f:
            blob = f.read()
        lines = blob.decode("cp1251").split("\r\n")
        ck(blob.endswith(b"\r\n") and b"\n" not in blob.replace(b"\r\n", b""), "txt: cp1251 и \\r\\n")
        got = dict(ln.split(";") for ln in lines if ln)
        ck(set(got) == {SKU, KG} and float(got[SKU]) == 6.0 and abs(float(got[KG]) - 0.392) < 1e-9,
           "txt содержит ровно 2 вывоза: %s" % got)
        book = load_workbook(os.path.join(p1, "Тест 1.xlsx"))
        ck(book.sheetnames == [SHEET_VYVOZ, SHEET_NOM, SHEET_INV, SHEET_C1, SHEET_C2, SHEET_SUP, SHEET_EXCL, SHEET_PACK, SHEET_DIFF, SHEET_CIG],
           "листы xlsx магазина: %s" % book.sheetnames)
        # оформление: заголовки и числа по центру; целое количество без «хвоста» запятой («5,» от формата 0.###)
        ws_v = book[SHEET_VYVOZ]
        ck(all(ws_v.cell(row=4, column=j).alignment.horizontal == "center" for j in range(1, 9)),
           "«Вывезти на склад»: заголовки по центру")
        fmt_seen = {}
        for r_ in range(5, ws_v.max_row + 1):
            if not ws_v.cell(row=r_, column=2).value:
                continue
            q_ = ws_v.cell(row=r_, column=4)
            fmt_seen[float(q_.value)] = q_.number_format
            ck(all(ws_v.cell(row=r_, column=j).alignment.horizontal == "center" for j in (1, 2, 4, 5, 6, 7)),
               "«Вывезти на склад»: строка %d - № / ШК / количество / ед. / себестоимость / сумма по центру" % r_)
        ck(fmt_seen.get(6.0) == "0", "целое количество без запятой: формат %r" % fmt_seen.get(6.0))
        ck(fmt_seen.get(0.392) == "0.###", "дробное количество: формат %r" % fmt_seen.get(0.392))
        ws_i = book[SHEET_INV]
        ck(all(ws_i.cell(row=1, column=j).alignment.horizontal == "center" for j in range(1, 8)),
           "«Остаётся - инвентаризация ЮА»: заголовки по центру")
        ck(ws_i.cell(row=2, column=1).alignment.horizontal == "center" and ws_i.cell(row=2, column=3).alignment.horizontal == "center",
           "«Остаётся - инвентаризация ЮА»: ШК и количество по центру")
        ck(all((ws_i.cell(row=r_, column=3).number_format == "0") == float(ws_i.cell(row=r_, column=3).value).is_integer()
               for r_ in range(2, ws_i.max_row + 1) if ws_i.cell(row=r_, column=3).value is not None),
           "«Остаётся - инвентаризация ЮА»: целое количество без запятой, дробное с форматом 0.###")
        book.close()
        inv_txt = os.path.join(p1, "Тест 1_инвентаризация_ЮА.txt")
        ck(os.path.isfile(inv_txt), "txt инвентаризации ЮА создан")
        if os.path.isfile(inv_txt):
            ig = dict(ln.split(";") for ln in open(inv_txt, "rb").read().decode("cp1251").split("\r\n") if ln)
            ck("019230010208" in ig and SKU not in ig and KG not in ig and B_NOUA not in ig,
               "инвентаризация: остающееся есть, вывоз и «ждёт решения» нет: %d строк" % len(ig))
            inv_sheet = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_INV, dtype=str)
            ck(set(inv_sheet["ШК в файле для ЮА"]) == set(ig), "лист «Остаётся - инвентаризация ЮА» и txt совпадают по ШК для ЮА")
        diff = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_DIFF, dtype=str)
        ck(len(diff) >= 2 and (diff["Штрих-код магазина"] != diff["Штрих-код там"]).all(),
           "«Один товар, разные ШК»: только строки, где ШК действительно разные (%d строк)" % len(diff))
        c1 = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_C1, dtype=str)
        ck(len(c1) == 3, "«Нет в матрице, но есть на ЮА»: ожидалось 3 строки (ЮА по ШК, ЮА по названию, "
                         "инвентаризация), а %d" % len(c1))
        ck((c1["Другой ШК"].fillna("") != c1["Штрих-код"]).all(),
           "«Нет в матрице, но есть на ЮА»: «Другой ШК» не повторяет ШК самой позиции (заведён на ЮА по тому же ШК -> колонка пустая)")
        nom = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_NOM, dtype=str)
        excl_n = len(pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_EXCL, dtype=str))
        want_nom = shops["Тест 1"]["vyvoz_pos"] + len(c1) + excl_n      # вне матрицы = вывезти + «есть на ЮА» + исключено
        ck(len(nom) == want_nom and nom["Штрих-код"].is_unique and "Решение" in nom.columns
           and "Вывезти (нет в матрице)" in set(nom["Решение"]), "«Нет в матрице (весь список)»: строк %d, ожидалось %d"
           % (len(nom), want_nom))
        ck(ws_v.cell(row=4, column=8).value == "Поставщик", "«Вывезти на склад»: вместо группы - поставщик: %r" % ws_v.cell(row=4, column=8).value)
        c2 = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_C2, dtype=str)
        ck(len(c2) == 2 and c2["Поставщик (по матрице)"].iloc[0] == "Поставщик Тест",
           "«В матрице, на ЮА не было»: поставщик из матрицы в колонке: %s" % c2.to_dict("records"))
        sp2 = pd.read_excel(os.path.join(p1, "Тест 1.xlsx"), sheet_name=SHEET_SUP, dtype=str)
        ck(len(sp2) == 2 and "Поставщик Тест" in set(sp2["Поставщик (по матрице)"]) and sp2["Позиций"].iloc[0] == "1"
           and sp2["Прямая доставка, поз."].iloc[0] == "1", "«По поставщикам»: %s" % sp2.to_dict("records"))
        sbook = load_workbook(os.path.join(out["day_dir"], SUMMARY_NAME))
        for nm in ("На склад", "По магазинам", SHEET_NOM, SHEET_INV, SHEET_C1, SHEET_C2, SHEET_SUP, SHEET_EXCL, SHEET_PACK, SHEET_DIFF, SHEET_CIG, "Проверки"):
            ck(nm in sbook.sheetnames, "в сводке нет листа %s" % nm)
        ws_s = sbook["На склад"]
        ck(all(ws_s.cell(row=3, column=j).alignment.horizontal == "center" for j in range(1, 8)),
           "сводка «На склад»: заголовки по центру")
        ck(ws_s.cell(row=4, column=1).alignment.horizontal == "center" and ws_s.cell(row=4, column=4).alignment.horizontal == "center"
           and ws_s.cell(row=4, column=4).number_format in ("0", "0.###"),
           "сводка «На склад»: ШК и количество по центру, формат без «хвоста» запятой")
        ws_m = sbook["По магазинам"]
        ck(all(ws_m.cell(row=1, column=j).alignment.horizontal == "center" for j in range(1, 21))
           and all(ws_m.cell(row=2, column=j).alignment.horizontal == "center" for j in range(2, 18)),
           "сводка «По магазинам»: заголовки и числа по центру")
        sbook.close()
        # перезапуск той же даты: старый результат уходит в _замененные, а не затирается
        out2 = run_vyvoz(d, ref=ref2)
        ck(os.path.isdir(os.path.join(d.out, REPLACED_DIR)), "повторный запуск: _замененные")
        # с настройкой «убирать и то, что на ЮА не завозилось»: B_NOUA уходит в вывоз, а не ждёт решения
        out3 = run_vyvoz(d, ref=ref2, remove_not_on_ua=True)
        with open(os.path.join(out3["day_dir"], "Тест 1", "Тест 1.txt"), "rb") as f:
            g3 = dict(ln.split(";") for ln in f.read().decode("cp1251").split("\r\n") if ln)
        ck(B_NOUA in g3 and set(g3) == {SKU, KG, B_NOUA, _ean("482000000777")},
           "настройка в прогоне: вывоз вместе с «на ЮА не завозилось»: %s" % sorted(g3))
        out4 = run_vyvoz(d, ref=ref2, excl_suppliers={sup_norm("Сувенир-опт")})
        with open(os.path.join(out4["day_dir"], "Тест 1", "Тест 1.txt"), "rb") as f:
            g4 = dict(ln.split(";") for ln in f.read().decode("cp1251").split("\r\n") if ln)
        ck(SKU not in g4 and KG in g4, "прогон с исключённым поставщиком: txt без его товара: %s" % sorted(g4))
        t4_ = _result_text(out4)                  # итог читает общий отчёт последнего прогона: снимаем его до следующего запуска
        ck(out4.get("excl_n") == 1 and u"Поставщиков в списке «не вывозить»: 1" in t4_,
           "итог прогона показывает число исключённых поставщиков: %s" % out4.get("excl_n"))
        out5 = run_vyvoz(d, ref=ref2, excl_suppliers={sup_norm("Сувенир-опт")}, manual=True, day=date(2026, 9, 30))   # другая дата: не трогает результаты выше
        with open(os.path.join(out5["day_dir"], "Тест 1", "Тест 1.txt"), "rb") as f:
            g5 = dict(ln.split(";") for ln in f.read().decode("cp1251").split("\r\n") if ln)
        ck(SKU not in g5 and g5 == g4, "ручной выбор: поставщик с галочкой «не вывозить» в список не попадает, без сигарет и кегов "
                                       "список как в обычном расчёте: %s" % sorted(g5))
        ws5 = load_workbook(os.path.join(out5["day_dir"], "Тест 1", "Тест 1.xlsx"))[SHEET_VYVOZ]
        ck(ws5.cell(row=4, column=10).value is None and not ws5["A3"].value,
           "ручной выбор: нечего возвращать в список - нет ни колонки причины, ни пометки A3")
        ck(out5.get("manual") and out5.get("excl_n") == 1 and u"РУЧНОЙ ВЫБОР" in _result_text(out5),
           "итог показывает ручной выбор и число поставщиков «не вывозить»: %s" % out5.get("excl_n"))
        ck(u"РУЧНОЙ ВЫБОР" not in t4_ and load_workbook(os.path.join(out4["day_dir"], "Тест 1", "Тест 1.xlsx"))[SHEET_VYVOZ].cell(row=4, column=10).value is None,
           "без ручного выбора лист и итог прежние")
        ck(os.path.isfile(os.path.join(d.cache, SUPPLIERS_FILE)), "прогон создаёт файл поставщиков по умолчанию")
        sp3 = pd.read_excel(os.path.join(out3["day_dir"], "Тест 1", "Тест 1.xlsx"), sheet_name=SHEET_SUP, dtype=str)
        ck(len(sp3) == 2 and "Поставщик Тест" in set(sp3["Поставщик (по матрице)"]),
           "настройка «убирать»: список по поставщикам для закупщицы остаётся: %s" % sp3.to_dict("records"))
        # финал: вывоз проведён, на полке остаётся ВСЁ из выгрузки (в т.ч. то, что в обычном прогоне уходило в вывоз / «ждёт решения»)
        fin_ = run_final_inventory(d, "Тест 1", ref=ref2)
        ck(fin_["ok"], "финальная инвентаризация: %s" % fin_["problems"])
        fin_txt = [f_ for f_ in fin_.get("files", []) if f_.endswith(INV_TXT_SUFFIX)]
        ck(len(fin_txt) == 1, "финал: txt инвентаризации создан: %s" % fin_.get("files"))
        if fin_txt:
            fg_ = dict(ln.split(";") for ln in open(fin_txt[0], "rb").read().decode("cp1251").split("\r\n") if ln)
            ck(SKU in fg_ and KG in fg_ and B_NOUA in fg_ and B_UA_ONLY in fg_ and "019230010208" in fg_,
               "финал: в файле и вывозимое по обычному расчёту, и «ждёт решения», и остающееся: %d строк" % len(fg_))
            ck(float(fg_[SKU]) == 6.0, "финал: количества по факту остатка (дубли слиты): %s" % fg_.get(SKU))
        st_ = fin_.get("stats", {})
        ck(st_.get("stock_pos") == st_.get("inv_pos", 0) + st_.get("out_pos", 0) and st_.get("nocard_pos", 0) >= 2,
           "финал: баланс остатка и список «нет карточки на ЮА»: %s" % st_)
        ck(os.path.isfile(os.path.join(fin_["out_dir"], "Тест 1_инвентаризация_ЮА.xlsx")), "финал: xlsx со сверкой создан")
        ck(not run_final_inventory(d, "Нет такого", ref=ref2)["ok"], "финал: неизвестный магазин -> ошибка, а не пустой файл")
        fr_ = final_inventory_rows(pd.DataFrame({"verdict": ["IN", "VYVOZ", "CHECK1", "CHECK2", "EXCL", "KUL", "PACK"]}))
        ck(len(fr_) == 5, "final_inventory_rows: все, кроме кулинарии и штучных: %d" % len(fr_))
        # зачистка остатка: всё, кроме расходников/розлива/овощей; своя папка дня; порог «вне матрицы» не мешает
        out_sw = run_vyvoz(d, ref=ref2, sweep=True)
        sw_shops = {x["shop"]: x for x in out_sw["shops"]}
        ck(out_sw["day_dir"].endswith(SWEEP_DAY_TAG) and os.path.isdir(out["day_dir"]),
           "зачистка: своя папка дня, обычный вывоз не тронут: %s" % out_sw["day_dir"])
        ck(sw_shops["Тест 1"]["status"] == "OK" and sw_shops["Тест 2"]["status"] == "OK",
           "зачистка: оба магазина OK (порог 40%% не применяется): %s" % [x["status"] for x in out_sw["shops"]])
        sw_txt = os.path.join(out_sw["day_dir"], "Тест 1", "Тест 1.txt")
        if os.path.isfile(sw_txt):
            gs_ = dict(ln.split(";") for ln in open(sw_txt, "rb").read().decode("cp1251").split("\r\n") if ln)
            ck(all(k in gs_ for k in (SKU, KG, B_NOUA, B_IN2, B_UA_ONLY, "019230010208")),
               "зачистка: вывозится и то, что в матрице / на ЮА / ждало решения: %d строк" % len(gs_))
            ck(not any(k in gs_ for k in (_ean("297895008888"), _ean("297895007777"), rz.BC_KRISHKA, B_COFFEE, B_COFFEE2,
                                          _ean("297895009999"), "2978950020707")),
               "зачистка: пакеты, стаканы, розлив, сырьё кофеаппарата и овощи остаются: %s" % sorted(gs_)[:5])
        else:
            ck(False, "зачистка: txt не создан")
        if gui:                                              # окно: все страницы собираются (только при явном --selftest)
            try:
                import tkinter as tk_ui
            except ImportError:
                tk_ui = None
            if tk_ui is not None:
                try:
                    ui_ = run_app2(dirs=d, test_mode=True)
                    ck(set(ui_[0]) == {"vyvoz", "split", "upd", "fin", "serv", "log"},
                       "окно: страницы %s" % (ui_[0],))
                except tk_ui.TclError as e_:
                    R.check("INFO", u"Окно в самотесте", u"не проверено: %s" % e_)
                except Exception as e_:
                    ck(False, "окно run_app2 не собралось: %s: %s" % (type(e_).__name__, e_))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # распределение исправленного списка: маршрут и деление по точкам
    ck(decide_route("", "", "", "")[0] == KIND_RC, "маршрут: нет истории -> РЦ")
    ck(decide_route("2026-09-01", "2026-08-01", "", "")[0] == KIND_RC and decide_route("2026-08-01", "2026-09-01", "", "")[0] == KIND_SHOPS,
       "маршрут: по магазину берётся последнее поступление")
    ck(decide_route("", "", "2026-04-01", "2026-09-01")[0] == KIND_RC and decide_route("", "", "", "2026-09-01")[0] == KIND_SHOPS,
       "маршрут: по сети любой РЦ-приход = РЦ, иначе прямая")
    ck(decide_route("2026-09-01", "", "", "", KIND_SHOPS)[0] == KIND_SHOPS, "маршрут: поставщик из справочника важнее истории")
    sh3 = pd.Series({"A": 0.5, "B": 0.3, "C": 0.2})
    al1 = allocate_qty(10, sh3, 1.0)
    ck(abs(sum(al1.values()) - 10) < 1e-9 and all(float(v).is_integer() for v in al1.values()) and max(al1, key=al1.get) == "A",
       "деление штук: сумма сходится, целые, лидер по доле: %s" % al1)
    al2 = allocate_qty(2, sh3, 1.0)
    ck(al2 == {"A": 2.0}, "меньше минимума на точку: всё лидеру: %s" % al2)
    al3 = allocate_qty(7, sh3, 1.0)
    ck(abs(sum(al3.values()) - 7) < 1e-9 and all(v >= SPLIT_MIN - 1e-9 for v in al3.values()), "минимум %d на точку соблюдён: %s" % (SPLIT_MIN, al3))
    al4 = allocate_qty(2.634, sh3, 0.1)
    ck(abs(sum(al4.values()) - 2.634) < 1e-9, "весовой: сумма сходится до грамма: %s" % al4)
    al5 = allocate_qty(25.5, sh3, 0.1)
    ck(abs(sum(al5.values()) - 25.5) < 1e-9, "весовой 25,5 кг: сумма сходится: %s" % al5)
    pa_ = _parse_args(["--final-inventory", "Тест 1", "--no-rc", "--top", "3", "--final-file", "x.xlsx"])
    ck(pa_["final"] == "Тест 1" and pa_["no_rc"] and pa_["top"] == 3 and pa_["final_file"] == "x.xlsx", "аргументы финала/без РЦ: %s" % pa_)
    gd_ = [guess_day(u"Состояние склада Качанівська 19 8.10.2026.xlsx")[0], guess_day(u"Состояние склада Склад ЮА_07.10.2026_2026-10-07_105335.xlsx")[0],
           guess_day(u"Состояние склада Качановская на  06.10.2026.xlsx")[0], guess_day(u"Состояние склада Грозненська 38 8.10.26.xlsx")[0]]
    ck(gd_ == [date(2026, 10, 8), date(2026, 10, 7), date(2026, 10, 6), date(2026, 10, 8)], "дата по имени файла (номер дома не дата): %s" % gd_)
    tp_ = tempfile.mkdtemp(prefix="vyvoz_days_")
    try:
        fa_, fb_ = os.path.join(tp_, u"Состояние склада Качановская на  06.10.2026.xlsx"), os.path.join(tp_, u"Состояние склада Качанівська 19 8.10.2026.xlsx")
        for f_ in (fb_, fa_):
            open(f_, "wb").close()
        ck(_pick_latest([fa_, fb_]) == fb_, "свежей выгрузкой выбирается 8.10, а не 06.10: %s" % _pick_latest([fa_, fb_]))
    finally:
        shutil.rmtree(tp_, ignore_errors=True)
    _selftest_split_update(ck)
    _selftest_kegs(ck)
    _selftest_split_min(ck)
    _selftest_split_demand(ck)
    _selftest_manual_rules(ck)
    _selftest_sweep(ck)
    _selftest_window(ck)
    _selftest_map(ck)

    reset_report()
    if fails:
        for f in fails:
            print("  FAIL:", f)
        R.check("ERROR", "Самотест вывоза", "; ".join(fails[:6]))
        return False
    R.check("OK", "Самотест вывоза", "все контрольные случаи совпали")
    return True


def _selftest_map(ck):
    """Таблица соответствия ШК: оценка пар, список кандидатов, загрузка решений, применение в расчёте."""
    t1, t2 = _map_tokens(u"ХД Тест Напій Лісовий 0,5л"), _map_tokens(u"Тест Напій Лісовий 0,5л з/б")
    c1 = _map_score(t1, t2, _map_norm(u"ХД Тест Напій Лісовий 0,5л"), _map_norm(u"Тест Напій Лісовий 0,5л з/б"), 9)
    ck(c1[0] == "A" and c1[2], "пара: префикс ХД и «з/б» не мешают: %s" % (c1,))
    c2 = _map_score(_map_tokens(u"Стакан Паперовий 175мл"), _map_tokens(u"Стакан Паперовий 250мл"), u"стакан паперовий 175мл", u"стакан паперовий 250мл", 9)
    ck(c2[0] in ("C", "-") and not c2[2] and u"ЧИСЛА РАЗНЫЕ" in c2[3], "разный объём - не уверенная пара: %s" % (c2,))
    c3 = _map_score(_map_tokens(u"Х.Ц.З. Сіль 1кг"), _map_tokens(u"Сіль ХЦЗ 1кг"), _map_norm(u"Х.Ц.З. Сіль 1кг"), _map_norm(u"Сіль ХЦЗ 1кг"), 9)
    ck(c3[0] == "A", "Х.Ц.З. = ХЦЗ, порядок слов не важен: %s" % (c3,))
    sc = lambda a_, b_, ed=9: _map_score(_map_tokens(a_), _map_tokens(b_), _map_norm(a_), _map_norm(b_), ed)
    for ua_, fm_ in ((u"Грин Дей Вино Ігристе ТМ Villa UA н/сол Біле 0,75л", u"Грин Дей Ігристе Вілла Крим н/с Біле 0,75л"),
                     (u"Сигарети Sobranie Gold", u"Сигарети Собраніе Голд 20шт"),
                     (u"ТВЕН NEO DEMI PURPLE YELLOW BOOST", u"Сигарети Нео Демі Пурбл Єлоу Буст 20шт"),
                     (u"Сигарети Lucky Strike Black Series Amber", u"Сигарети Лакі Страйк Блек Серія Амбер 20шт"),
                     (u"Сигарети Прилуки Класичні 10", u"Сигарети Прилуки Класичні 10мг  20шт"),
                     (u"Вино Ігристе Французький Бульвар Брют Біле 0,75л", u"Ігристе Французький Бульвар брют біле 0,75л (6)"),
                     (u"Комо Сир Пл.Вершковий 35% 75 гр", u"Комо Сир Пл.Вершковий 75 гр")):
        c_ = sc(ua_, fm_)
        ck(c_[0] == "A" and c_[1] >= MAP_AUTO_SIM, "пара ЮА (латиница / ящик / %%) ~ Family: %s ~ %s: %s" % (ua_, fm_, c_))
    c_ = sc(u"ТВЕН NEO DEMI AMBER BOOST", u"Сигарети Нео Демі Пурбл Єлоу Буст 20шт", 2)
    ck(c_[1] < MAP_AUTO_SIM, "другой вкус с похожим ШК заранее не отмечается: %s" % (c_,))
    c_ = sc(u"Комо Сир Пл.Вершковий 35% 75 гр", u"Комо Сир Пл.Вершковий 50% 75 гр")
    ck(c_[0] != "A" and not c_[2], "разная жирность у обеих сторон - разные товары: %s" % (c_,))
    ck(_lev2("4823098203162", "4823098303162") == 1 and _lev2("1234567", "7654321") == 9, "расстояние между ШК")
    FA, FB = _ean("482003000011"), _ean("482003000028")
    UA1, UA2 = _ean("293808000011"), _ean("482003000029")
    tmp = tempfile.mkdtemp(prefix="vyvoz_map_")
    g_ = globals()
    saved_fk = g_["fetch_family_keys"]
    g_["fetch_family_keys"] = lambda keys: set()             # BigQuery в самотесте не трогаем
    try:
        d = Dirs(tmp)
        d.ensure()
        _write_state_xlsx(os.path.join(d.ua_wh, u"Состояние склада Полевая склад ЮА 08.10.2026.xlsx"), u"Полевая-Склад ЮА",
                          [(UA1, u"ХД Тест Напій Лісовий 0,5л", 4, u"шт", 10.0), (UA2, u"Тест Сік Яблуко 1л", 2, u"шт", 20.0)],
                          shop_name=u"Полевая-Склад ЮА")
        _write_state_xlsx(os.path.join(d.stores, u"Состояние склада Тест Магазин на 08.10.2026.xlsx"), u"Тест Магазин",
                          [(FA, u"Тест Напій Лісовий 0,5л з/б", 5, u"шт", 10.0), (FB, u"Тест Сок Яблуко 1л", 3, u"шт", 20.0)])
        matrix = pd.DataFrame([(FA, u"Тест Напій Лісовий 0,5л з/б", "ok", u"Пост"), (FB, u"Тест Сок Яблуко 1л", "ok", u"Пост")],
                              columns=["barcode", "product_name", "status", "supplier"])
        reset_report()
        ua_df, _ui = load_ua(d)
        ref = make_ref(matrix, None, None, ua_df, None)
        info = run_map_list(d, ref=ref)
        ck(info["ok"] and info["stats"]["pool"] == 2 and info["stats"]["A"] == 2, "список соответствия: 2 позиции ЮА, обе с уверенной парой: %s %s"
           % (info.get("problems"), info.get("stats")))
        hist_ = pd.DataFrame({"barcode": [UA2], "types": [u"Прям"], "last_date": ["2026-10-01"], "last_supplier": [u"Пост"]})
        info_h = run_map_list(d, ref=make_ref(matrix, None, None, ua_df, hist_))
        g_["fetch_family_keys"] = lambda keys: set(k for k in keys if k == bc_key(UA1))
        info_b = run_map_list(d, ref=ref)
        g_["fetch_family_keys"] = lambda keys: set()
        ck(info_h["ok"] and info_h["stats"]["pool"] == 2 and info_b["ok"] and info_b["stats"]["pool"] == 2,
           "тот же ШК в базе Family, но есть похожий в остатках - позиция остаётся в списке: %s %s" % (info_h.get("stats"), info_b.get("stats")))
        info_n = run_map_list(d, ref=make_ref(matrix, None, None, ua_df, hist_), per_item=0)     # кандидатов нет ни у кого
        ck(info_n["ok"] and info_n["stats"]["pool"] == 1,
           "тот же ШК есть в базе Family, похожих нет - пара не нужна, в список не идёт: %s" % info_n.get("stats"))
        rec_ = pd.DataFrame({"old_barcode": [UA1], "new_barcode": [FA]})      # склейка проекта матрицы: ШК ЮА -> ШК матрицы
        info_rc = run_map_list(d, ref=make_ref(matrix, rec_, None, ua_df, None))
        ck(info_rc["ok"] and info_rc["stats"]["pool"] == 1,
           "склеенный в barcode_recode_map ШК ЮА (ЮА -> матрица) не попадает в список без пары: %s" % info_rc.get("stats"))
        info = run_map_list(d, ref=ref)                                        # файл списка снова полный (имя файла - до минуты)
        wb = load_workbook(info["path"], data_only=True)
        rows = list(wb[u"Кандидаты"].iter_rows(values_only=True))
        by_ua = {str(r[3]): r for r in rows[1:]}
        ck(by_ua.get(bc_key(UA1)) is not None or by_ua.get(UA1) is not None, "в списке есть позиция ЮА")
        r1 = by_ua.get(UA1) or by_ua.get(bc_key(UA1))
        ck(r1 is not None and str(r1[8]) == FA, "кандидат 1 для ХД Тест Напій = ШК Family %s: %s" % (FA, r1[8] if r1 else None))
        # решения: первая пара ТАК, вторая НЕТ
        wb2 = load_workbook(info["path"])
        ws2 = wb2[u"Кандидаты"]
        for row_ in ws2.iter_rows(min_row=2):
            if str(row_[3].value) == UA1:
                row_[0].value = u"ТАК"
            elif str(row_[3].value) == UA2:
                row_[0].value = u"НЕТ"
        dec = os.path.join(tmp, u"решения.xlsx")
        wb2.save(dec)
        imp = run_map_import(d, dec)
        ck(imp["ok"] and imp["added"] == 1 and imp["rejected"] == 1, "загрузка решений: %s" % imp)
        yes, no = load_ua_map(d)
        ck(yes == {bc_key(FA): bc_key(UA1)} and (bc_key(FB), bc_key(UA2)) in no, "таблица пар прочитана: %s %s" % (yes, no))
        info2 = run_map_list(d, ref=ref)
        ck(info2["ok"] and info2["stats"]["pool"] == 1 and info2["stats"]["A"] == 0, "после решений: ЮА с парой не в списке, отклонённый кандидат не предлагается: %s" % info2.get("stats"))
        # применение: Family-позиция с парой в таблице = заведена на ЮА, в инвентаризацию идёт ШК ЮА
        raw = pd.DataFrame([(FA, u"Тест Напій Лісовий 0,5л з/б", 5, u"шт", 10.0)], columns=["raw", "name", "qty", "unit", "cost"])
        raw["shop"] = "Тест Магазин"
        raw["bc"] = [rz.restore_barcode(rz.fmt_barcode(v_))[0] for v_ in raw["raw"]]
        raw["key"] = [bc_key(v_) for v_ in raw["raw"]]
        agg, _i = prepare_stock(raw)
        r_no = classify_stock(agg, make_ref(matrix, None, None, ua_df, None), True, True, set(), set())
        r_yes = classify_stock(agg, make_ref(matrix, None, None, ua_df, None, ua_map=yes), True, True, set(), set())
        ck(r_no["verdict"].iloc[0] == "VYVOZ" and r_yes["verdict"].iloc[0] == "IN_UA_NAME" and r_yes["inv_bc"].iloc[0] == UA1,
           "таблица соответствия: без неё вывозили бы, с ней остаётся и идёт в инвентаризацию под ШК ЮА: %s -> %s %s"
           % (r_no["verdict"].iloc[0], r_yes["verdict"].iloc[0], r_yes["inv_bc"].iloc[0]))
        r_rc = classify_stock(agg, make_ref(matrix, rec_, None, ua_df, None), True, True, set(), set())
        ck(r_rc["verdict"].iloc[0] == "IN_UA_NAME" and r_rc["inv_bc"].iloc[0] == UA1 and u"barcode_recode_map" in r_rc["reason"].iloc[0],
           "склейка ЮА -> матрица: позиция матрицы остаётся и идёт в инвентаризацию под ШК ЮА: %s %s %s"
           % (r_rc["verdict"].iloc[0], r_rc["inv_bc"].iloc[0], r_rc["reason"].iloc[0]))
    finally:
        g_["fetch_family_keys"] = saved_fk
        shutil.rmtree(tmp, ignore_errors=True)


def _selftest_sweep(ck):
    """Зачистка остатка: что вывозится, что остаётся; число точек по сумме на точку."""
    ck(auto_top_n(29000) == 5 and auto_top_n(12000) == 2 and auto_top_n(3000) == 1 and auto_top_n(100000) == SPLIT_TOP_N
       and auto_top_n(50000) == 8 and auto_top_n(0) == 1,
       "число точек по 5-7 тыс. грн: %s" % [auto_top_n(x) for x in (29000, 12000, 3000, 100000, 50000, 0)])
    ck(all(5000 <= 29000.0 / auto_top_n(29000) <= 7000 for _ in (0,)), "29 тыс. на 5 точек = 5,8 тыс. на точку")
    B = {n: _ean("4821002000%02d" % i) for i, n in enumerate(
        ("inm", "cig", "keg", "bag", "coffee", "gama", "nugget", "kul", "globino", "plain", "veg", "xlad", "sim", "badbc"), 11)}
    matrix = pd.DataFrame([(B["inm"], u"Товар в матрице и на ЮА", "ok", u"Пост А"), (B["gama"], u"Багета с курицей", "ok", u"Гама"),
                           (B["globino"], u"Глобино Колбаса Тест", "ok", u"Глобино")],
                          columns=["barcode", "product_name", "status", "supplier"])
    ua = pd.DataFrame([(bc_key(B["inm"]), u"Товар в матрице и на ЮА", u"склад ЮА")], columns=["key", "name", "src"])
    coffee = pd.DataFrame([(B["coffee"], u"Якобз Кава Тест (1кг)")], columns=["barcode", "product_name"])
    hist = pd.DataFrame([(B["xlad"], u"Прямая", "2026-10-01", u"Хладопром")], columns=["barcode", "types", "last_date", "last_supplier"])
    ref = make_ref(matrix, None, coffee, ua, hist)
    rows = [(B["inm"], u"Товар в матрице и на ЮА", 5, u"шт", 10.0), (B["cig"], u"Сигарети Тест Червоні 20шт", 3, u"шт", 100.0),
            (B["keg"], u"БІР Кег Тест 30л", 2, u"шт", 500.0), (B["bag"], u"Пакет БОПП 150мм*200мм", 100, u"шт", 0.5),
            (B["coffee"], u"Якобз Кава Тест (1кг)", 2, u"кг", 1400.0), (B["gama"], u"Багета с курицей", 1, u"шт", 50.0),
            (B["nugget"], u"Нагетси Курячі 400г", 2, u"шт", 90.0), (B["kul"], u"Кулінарія Борщ Тест", 1, u"кг", 20.0),
            (B["globino"], u"Глобино Колбаса Тест", 1, u"шт", 100.0), (B["plain"], u"Сувенир Тест", 4, u"шт", 10.0),
            (B["veg"], u"Овочі/Фрукти Морква 1кг з уцінкою", 1.4, u"кг", 10.0),
            (B["xlad"], u"Мороженое Тест 75г", 28, u"шт", 37.8),
            (B["sim"], u"Водафон Стартовий Пакет Турбо 150", 3, u"шт", 100.0), ("12345", u"Плохой ШК", 1, u"шт", 5.0)]
    raw = pd.DataFrame(rows, columns=["raw", "name", "qty", "unit", "cost"])
    raw["shop"] = "Тест Зачистка"
    raw["bc"] = [rz.restore_barcode(rz.fmt_barcode(v))[0] for v in raw["raw"]]
    raw["key"] = [bc_key(v) for v in raw["raw"]]
    agg, _i = prepare_stock(raw)
    res0 = classify_stock(agg, ref, True, True, set(), set())
    tmp = tempfile.mkdtemp(prefix="vyvoz_sweep_")
    try:
        d = Dirs(tmp)
        d.ensure()
        rules = load_sweep_stay(d)                                   # файла нет: создаётся с решениями 08.10.2026
        ck(os.path.isfile(os.path.join(d.cache, SWEEP_STAY_FILE)) and sup_norm(u"Гама") in rules["sup"]
           and sup_norm(u"Хладік") in rules["sup"] and "нагетс" in rules["name"], "правила зачистки созданы: %s" % rules["sup"])
        res = sweep_verdicts(res0, rules, ref)
        V = {r.bc: r.verdict for r in res.itertuples()}
        ck(all(V[B[k]] == "VYVOZ" for k in ("inm", "cig", "keg", "globino", "plain", "sim")),
           "зачистка вывозит: в матрице и на ЮА, сигареты, кеги, Глобино, прочее, «Стартовий Пакет» не пакет: %s" % V)
        ck(all(V[B[k]] == "EXCL" for k in ("bag", "coffee", "gama", "nugget", "veg", "xlad")),
           "зачистка оставляет: пакет, кофе-сырьё, Гама, нагетсы, овощи, Хладік: %s" % V)
        ck(V[B["kul"]] == "KUL", "кулинария остаётся кулинарией: %s" % V[B["kul"]])
        bad_ = res.loc[res["name"] == u"Плохой ШК", "verdict"]
        ck(len(bad_) == 1 and bad_.iloc[0] == "EXCL", "некорректный ШК остаётся (в ТСД не загрузить): %s" % list(bad_))
        pc = partition_check(res)
        ck(pc["ok"], "баланс остатка при зачистке сходится: %s" % pc)
        d2 = pd.DataFrame({"u": [1]})
        ck(_sweep_name_hit(u"комо сир", {"name": [u"^комо"]}) == u"комо" and not _sweep_name_hit(u"комодо сир", {"name": [u"^комо"]}),
           "правило «^слово» - начало названия по слову")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _selftest_window(ck):
    """Помощники окна: магазин по умолчанию, убранные магазины, проверка входных папок, лист для распределения."""
    tmp = tempfile.mkdtemp(prefix="vyvoz_win_")
    try:
        d = Dirs(tmp)
        d.ensure()
        t0 = time.time() - 5 * 86400
        fa = os.path.join(d.stores, u"Состояние склада Качанівська 19 8.10.2026.xlsx")
        fb = os.path.join(d.stores, u"Состояние склада Грозненська 38 8.10.2026.xlsx")
        fc = os.path.join(d.stores, u"Состояние склада Старая 6.10.2026.xlsx")
        for f_, dt_ in ((fa, 100), (fb, 200), (fc, 300)):          # у «Старой» файл новее по времени, но выгрузка за 6.10
            open(f_, "wb").close()
            os.utime(f_, (t0 + dt_, t0 + dt_))
        found = {u"Качанівська 19": (fa, None), u"Грозненська 38": (fb, None), u"Старая": (fc, None)}
        ck(pick_default_store(found) == u"Грозненська 38",
           "магазин по умолчанию: свежайшая выгрузка (дата в имени, затем время файла): %s" % pick_default_store(found))
        ck(pick_default_store({}) is None, "магазин по умолчанию: выгрузок нет -> None")
        ck(load_hidden_stores(d) == set(), "убранных магазинов нет")
        save_hidden_stores(d, {u"Качанівська 19", u"Старая"})
        hid = load_hidden_stores(d)
        vis = visible_stores(found, hid)
        ck(hid == {u"Качанівська 19", u"Старая"} and set(vis) == {u"Грозненська 38"} and pick_default_store(vis) == u"Грозненська 38",
           "убранные магазины не показываются: %s" % sorted(vis))
        save_hidden_stores(d, set())
        ck(load_hidden_stores(d) == set() and set(visible_stores(found, load_hidden_stores(d))) == set(found),
           "«вернуть убранные»: все магазины снова в списке")

        _write_state_xlsx(os.path.join(d.ua_wh, u"Состояние склада Полевая склад ЮА 08.10.2026.xlsx"), u"Полевая-Склад ЮА",
                          [(_ean("482100100011"), u"Тест", 1, u"шт", 1.0)])
        rows_ = {x["key"]: x for x in input_status(d)}
        ck(rows_["ua_wh"]["level"] == "OK" and rows_["ua_wh"]["file"].startswith(u"Состояние склада Полевая"),
           "входные папки: свежий склад ЮА - OK: %s" % rows_["ua_wh"])
        ck(rows_["stores"]["level"] == "WARN" and rows_["stores"]["text"].startswith(u"СТАРАЯ") and u"Грозненська" in rows_["stores"]["file"],
           "входные папки: выгрузка магазина пятидневной давности - СТАРАЯ: %s" % rows_["stores"])
        ck(rows_["ua_receipts"]["level"] == "WARN" and rows_["ua_stores"]["level"] == "INFO" and rows_["ref"]["level"] == "INFO",
           "входные папки: нет приходов - ВНИМ, нет магазина в базе ЮА - не обязательно: %s" % {k: v["level"] for k, v in rows_.items()})
        d0 = Dirs(os.path.join(tmp, u"пусто"))
        d0.ensure()
        lv0 = {x["key"]: x["level"] for x in input_status(d0)}
        ck(lv0["stores"] == "ERR" and lv0["ua_wh"] == "ERR", "входные папки пусты: обязательные файлы - ОШИБКА: %s" % lv0)

        shop, day_s = u"Тест Магазин", u"08.10.2026"
        items = [(_ean("482100100011"), u"Товар 1", 5, u"шт", 10.0, u"Пост"), (_ean("482100100012"), u"Товар 2", 3, u"кг", 20.0, u"Пост")]
        auto = os.path.join(d.out, u"2026-10-08", shop, shop + u".xlsx")
        _write_vyvoz_list_xlsx(auto, shop, day_s, items)
        wb_ = load_workbook(auto)
        w_ = wb_[SHEET_VYVOZ]
        w_.cell(row=4, column=9, value=u"Возможно уже на ЮА")
        w_.cell(row=5, column=10, value=u"обычно не вывозим: причина")
        w_["A3"] = MANUAL_NOTE_START + u": тест"
        wb_.create_sheet(u"Исключено (не трогаем)")
        wb_.save(auto)
        wb_.close()
        al_ = auto_lists(d)
        ck(len(al_) == 1 and al_[0][1] == auto and shop in al_[0][0], "автоматические листы находятся в ВЫВОЗ: %s" % al_)
        ck(list_corrected(d) == [] and not is_in_corr(d, auto), "автоматический лист - не из Корректировка_ЮА")
        ck(has_manual_note(auto), "пометка ручного выбора находится в A3")
        dst = copy_list_as_is(d, auto)
        ck(is_in_corr(d, dst) and os.path.basename(dst) == u"Тест_Магазин_2026-10-08.xlsx", "«как есть»: копия в Корректировка_ЮА: %s" % dst)
        wc_ = load_workbook(dst)
        ck(wc_.sheetnames == [SHEET_VYVOZ] and wc_[SHEET_VYVOZ].max_column == 7 and not wc_[SHEET_VYVOZ]["A3"].value and not has_manual_note(dst),
           "«как есть»: один лист, 7 колонок, без пометки: %s, колонок %s" % (wc_.sheetnames, wc_[SHEET_VYVOZ].max_column))
        wc_.close()
        chk_ = corrected_check(dst)
        ck(chk_["ok"] and chk_["shop"] == shop and chk_["n"] == 2 and chk_["day"] == day_s and abs(chk_["sum"] - 110.0) < 1e-6,
           "лист для распределения проверяется: %s" % chk_)
        dst2 = copy_list_as_is(d, auto)
        ck(dst2 != dst and os.path.isfile(dst) and os.path.isfile(dst2) and len(list_corrected(d)) == 2,
           "«как есть» второй раз: новый файл, прежний не затёрт: %s" % os.path.basename(dst2))
        wa_ = load_workbook(auto)
        ck(wa_.sheetnames == [SHEET_VYVOZ, u"Исключено (не трогаем)"] and wa_[SHEET_VYVOZ].max_column >= 10 and has_manual_note(auto),
           "исходный автоматический список не изменён")
        wa_.close()
        bad = os.path.join(corr_dir(d), u"Чужой.xlsx")
        wb_ = Workbook()
        wb_.active.title = u"Лист1"
        wb_.save(bad)
        bad_chk = corrected_check(bad)
        ck(not bad_chk["ok"] and u"нет листа" in bad_chk["text"], "чужой файл не годится для распределения: %s" % bad_chk["text"])
        for nm_ in (u"~$Тест.xlsx", u"Тест_проверка.xlsx"):
            open(os.path.join(corr_dir(d), nm_), "wb").close()
        ck(not any(os.path.basename(p_) in (u"~$Тест.xlsx", u"Тест_проверка.xlsx") for p_ in list_corrected(d)),
           "в списке листов нет временных ~$ и файлов «проверка»")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _write_vyvoz_list_xlsx(path, shop, day_str, rows):
    """Лист «Вывезти на склад» в том виде, как его пишет write_store_book (rows: ШК, название, кол-во, ед., себестоимость,
    поставщик). Итоги в шапке и в «ИТОГО» нарочно неверные, как у списка, из которого строки удалили руками."""
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_VYVOZ
    ws["A1"] = u"Вывоз вне матрицы: %s -> %s" % (shop, DEST_NAME)
    ws["A2"] = u"Остатки на %s (по учёту Торгсофта)     Позиций: 999     Единиц: 9999.5     Себестоимость: 99999.99" % day_str
    for j, h in enumerate([u"№", u"Штрих-код", u"Название товара", u"Количество", u"Ед. изм.", u"Себестоимость", u"Сумма",
                           u"Поставщик"], 1):
        ws.cell(row=4, column=j, value=h)
    r = 4
    for n, (bc, name, qty, unit, cost, sup) in enumerate(rows, 1):
        r += 1
        for j, v in enumerate([n + 100, str(bc), name, qty, unit, cost, round(qty * cost, 2), sup], 1):
            ws.cell(row=r, column=j, value=v)
    ws.cell(row=r + 1, column=3, value=u"ИТОГО:")
    ws.cell(row=r + 1, column=4, value=9999.5)
    ws.cell(row=r + 1, column=7, value=99999.99)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def _selftest_manual_rules(ck):
    """Ручной выбор: в список возвращаются только сигареты и кеги; галочки поставщиков, список ШК, кулинария, пакеты, стаканы,
    сырьё кофе и овощи остаются вне списка (жалоба 08.10.2026: в списке была выпечка и кулинария при галочках «не вывозить»)."""
    rows_ = [  # (ключ, вердикт, причина, название, поставщик)
        ("k01", "EXCL", u"сигареты (не вывозим)", u"Сигарети Тест Червоні 20шт", u"Сигарети_BAT"),
        ("k02", "EXCL", u"название:  кег", u"БІР  Кег Пиво Тест 0,5л (30)", u"БІР_КЕГ"),
        ("k03", "EXCL", u"сигареты (не вывозим)", u"Сигарети Тест Сині 20шт", u"Сигарети Тест"),        # поставщик с галочкой
        ("k04", "EXCL", u"сигареты (не вывозим)", u"Сигарети Тест Білі 20шт", u"Сигарети_BAT"),         # ШК в списке «не вывозить»
        ("k05", "KUL", u"кулинария (не вывозим, на ЮА не переходит)", u"Випічка Хліб Тест 500г", u"Кулінарія"),
        ("k06", "KUL", u"кулинария (не вывозим, на ЮА не переходит)", u"Кулінарія Борщ Тест", u""),
        ("k07", "EXCL", u"поставщик исключён из вывоза: Хладік", u"Хладік Пельмені Тест 0,4кг", u"Хладік"),
        ("k08", "EXCL", u"пакет", u"Пакет БОПП Тест", u"Центр Витратних Матеріалів"),
        ("k09", "EXCL", u"стакан", u"Агропром Стакан Пластик 300 мл 1 шт", u""),
        ("k10", "EXCL", u"сырьё кофеаппарата (поставщик каваапарат, по ШК)", u"Якобз Кава Зернова (1кг)", u"ТОВ СТВ Дистрибюшн (Каваапарат)"),
        ("k11", "EXCL", u"название: овочі/фрукти", u"Овочі/Фрукти Морква 1кг", u"Магазин"),
        ("k12", "EXCL", u"в списке «не вывозить» по ШК (vyvoz_keep.csv); заведён на ЮА, допродаём", u"Кава Тест 3в1 12гр", u"Якобз"),
        ("k13", "VYVOZ", u"нет в матрице, на ЮА не заведён", u"Сувенир Тест", u""),
        ("k14", "CHECK1", u"заведён на ЮА (по ШК)", u"Новинка Тест", u""),
        ("k15", "PACK", u"штучный товар: упаковка уже на ЮА", u"Монжар Драже Тест (24) 1 шт", u"Монжар")]
    res = pd.DataFrame(rows_, columns=["key", "verdict", "reason", "name", "supplier"])
    out = manual_verdicts(res, {sup_norm(u"Сигарети Тест"), sup_norm(u"Хладік"), sup_norm(u"Кулінарія")}, {"k04", "k12"})
    v = dict(zip(out["key"], out["verdict"]))
    why = dict(zip(out["key"], out["reason"]))
    ck(v["k01"] == "VYVOZ" and v["k02"] == "VYVOZ" and why["k01"].startswith(MANUAL_MARK) and why["k02"].startswith(MANUAL_MARK),
       "ручной выбор: сигареты и кеги - в список, прежняя причина сохранена: %s | %s" % (why["k01"], why["k02"]))
    ck(v["k03"] == "EXCL" and v["k04"] == "EXCL",
       "ручной выбор: сигареты поставщика с галочкой и из списка ШК в список не идут: %s %s" % (v["k03"], v["k04"]))
    ck(v["k05"] == "KUL" and v["k06"] == "KUL", "ручной выбор: кулинария и выпечка в список не идут: %s %s" % (v["k05"], v["k06"]))
    stay = [k for k in ("k07", "k08", "k09", "k10", "k11", "k12") if v[k] != "EXCL"]
    ck(not stay, "ручной выбор: поставщик с галочкой, пакет, стакан, сырьё кофе, овощи, список ШК - вне списка: %s" % stay)
    ck(v["k13"] == "VYVOZ" and v["k14"] == "CHECK1" and v["k15"] == "PACK"
       and all(why[k] == r_[2] for r_ in rows_ for k in (r_[0],) if k not in ("k01", "k02")),
       "ручной выбор: остальные вердикты и причины не меняются")
    ck(len(manual_verdicts(res.head(0), None, None)) == 0, "ручной выбор: пустой магазин не ломает расчёт")


def _mock_sales(all_sales, all_checks=None):
    """Подмена fetch_bc_sales для самотестов: продажи и число чеков из словарей."""
    def f(bcs, stores, days=SPLIT_DAYS_BC, checks=None):
        ks, st = set(bc_key(b) for b in bcs), set(stores)
        if checks is not None and all_checks:
            checks.update({kk: v for kk, v in all_checks.items() if kk[0] in ks and kk[1] in st})
        return {kk: v for kk, v in all_sales.items() if kk[0] in ks and kk[1] in st}
    return f


def _selftest_kegs(ck):
    """Кеги (решение пользователя 09.10.2026): по литрам не делятся; объём кега - в скобках названия, порция 0,5 л;
    целые кеги и начатый - каждый целиком на одну выбранную точку с наибольшим спросом на этот ШК; на РЦ и вне выбранных не едут."""
    ck(keg_size_from_name(u"БІР  Кег Квас Тарас 0,5л (50)") == 100.0 and keg_size_from_name(u"Оболонь Кегове Бір Сидр 0,5л(30)") == 60.0
       and keg_size_from_name(u"Кегове Бір Хаус Сидр Груша 0,5л") is None, "кеги: объём по названию: (50) -> 100 порций, (30) -> 60, без скобок - нет")
    ck(keg_pieces(123, 60.0) == [60.0, 60.0, 3.0] and keg_pieces(81, 100.0) == [81.0] and keg_pieces(40, None) == [40.0],
       "кеги: позиция режется на целые кеги и начатый: %s %s" % (keg_pieces(123, 60.0), keg_pieces(81, 100.0)))
    tmp = tempfile.mkdtemp(prefix="vyvoz_keg_")
    g = globals()
    saved = (g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"], g["fetch_keg_sizes"])
    try:
        d = Dirs(tmp)
        d.ensure()
        shop, day_s = u"Тест Магазин", u"08.10.2026"
        K1, K2, K3, N1, K4 = (_ean("4821003000%02d" % n) for n in (11, 12, 13, 14, 15))
        items = [(K1, u"БІР Кег Тест Пиво 0,5л (30)", 135.8, u"шт", 20.0, u"БІР"),     # 60 + 60 + 15,8: продаётся на А и Б
                 (K2, u"БІР Кег Тест Ситро 0,5л (50)", 81, u"шт", 15.0, u"БІР"),       # начатый 50-литровый: продаётся на В (вне выбранных) и Б
                 (K3, u"Кегове Бір Тест Сидр 0,5л", 40, u"шт", 10.0, u"БІР"),          # объём по приходам сети, нигде не продавался
                 (N1, u"Обычный товар РЦ", 6, u"шт", 10.0, u"Пост"),
                 (K4, u"Кегове Бір Тест Квас 0,5л", 130, u"шт", 5.0, u"БІР")]           # объём не найден, продаётся только на Г (вне выбранных)
        corr = os.path.join(tmp, CORR_DIR, u"Тест_Магазин.xlsx")
        _write_vyvoz_list_xlsx(corr, shop, day_s, items)
        g["fetch_route_history"] = lambda bcs, sh: {          # по истории всё приходило через РЦ
            bc_key(b): {"shop_rc": "2026-09-01", "shop_dir": "", "net_rc": "", "net_dir": "", "sup": "Пост"} for b in bcs}
        g["pick_top_shops"] = lambda ua_stores, n=SPLIT_TOP_N, days=SPLIT_DAYS_SHOP: pd.DataFrame(
            {"store": [u"Точка А", u"Точка Б", u"Точка В", u"Точка Г"][:n], "rev": [400.0, 300.0, 200.0, 100.0][:n]})
        g["fetch_bc_sales"] = _mock_sales({(bc_key(K1), u"Точка А"): 100.0, (bc_key(K1), u"Точка Б"): 80.0,
                                           (bc_key(K2), u"Точка В"): 50.0, (bc_key(K2), u"Точка Б"): 10.0,
                                           (bc_key(K4), u"Точка Г"): 20.0})
        g["fetch_keg_sizes"] = lambda bcs, days=SPLIT_KEG_DAYS: {k_: 100.0 for k_ in (bc_key(K3),) if k_ in set(bc_key(b) for b in bcs)}
        ref0 = make_ref(pd.DataFrame(columns=["barcode", "product_name", "status", "supplier"]))
        info = run_split(d, corr, ref=ref0, top_n=2, per_shop=(0.0, 0.0))
        ck(info["ok"], "кеги: распределение прошло: %s" % info.get("problems"))
        al_ = info.get("alloc")
        got = {} if al_ is None else {(a_, b_): round(float(q_), 3) for a_, b_, q_ in zip(al_["addr"], al_["bc"], al_["qty"])}
        ck({k[0]: v for k, v in got.items() if k[1] == K1} == {u"Точка А": 75.8, u"Точка Б": 60.0},
           "кеги: 135,8 = 60 + 60 + 15,8 целыми кусками по спросу (А 60+15,8, Б 60): %s" % {k: v for k, v in got.items() if k[1] == K1})
        ck({k[0]: v for k, v in got.items() if k[1] == K2} == {u"Точка Б": 81.0},
           "кеги: начатый кег целиком на выбранную точку, где ШК продаётся, а не на точку вне выбранных: %s" % {k: v for k, v in got.items() if k[1] == K2})
        ck({k[0]: v for k, v in got.items() if k[1] == K3} == {u"Точка А": 40.0} and {k[0]: v for k, v in got.items() if k[1] == K4} == {u"Точка А": 130.0},
           "кеги: нигде не продавался / продаётся вне выбранных - целиком на точку с наибольшим спросом: %s"
           % {k: v for k, v in got.items() if k[1] in (K3, K4)})
        ck(not any(a_ == SPLIT_RC_NAME and b_ in (K1, K2, K3, K4) for (a_, b_) in got), "кеги: на РЦ не едут")
        ck(got.get((SPLIT_RC_NAME, N1)) == 6.0, "обычный товар по истории РЦ по-прежнему едет на РЦ: %s" % {k: v for k, v in got.items() if k[1] == N1})
        sd = os.path.join(d.out, "2026-10-08", shop, SPLIT_OUT_DIR)
        files = sorted(f for f in os.listdir(sd) if f.endswith(".txt")) if os.path.isdir(sd) else []
        ck(files == sorted([u"Точка А.txt", u"Точка Б.txt", SPLIT_RC_NAME + u".txt"]), "кеги: файлы только на выбранные точки и РЦ: %s" % files)
        wy = {r_.bc: r_.why for r_ in info["routes"].itertuples(index=False)}
        ck(u"продавался" in wy.get(K1, u"") and u"нигде не продавался" in wy.get(K3, u"") and u"РЦ" not in wy.get(K1, u""),
           "кеги: основание маршрута понятно: %s | %s" % (wy.get(K1), wy.get(K3)))
        kn = [n_ for n_ in info.get("notes", []) if n_.startswith(u"Кеги")]
        ck(kn and u"135.8 = 60 + 60 + 15.8" in kn[0] and u"по приходам сети" in kn[0] and u"не найден" in kn[0],
           "кеги: в итоге видно, как порезаны кеги и откуда объём: %s" % kn)
    finally:
        g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"], g["fetch_keg_sizes"] = saved
        shutil.rmtree(tmp, ignore_errors=True)


def _selftest_split_min(ck):
    """Распределение «по сумме»: ни одна точка не получает меньше 5 тыс. грн, кеги только на выбранные точки."""
    tmp = tempfile.mkdtemp(prefix="vyvoz_min_")
    g = globals()
    saved = (g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"])
    try:
        d = Dirs(tmp)
        d.ensure()
        shop, day_s = u"Тест Магазин", u"08.10.2026"
        P1, P2, P3, KG = (_ean("4821004000%02d" % n) for n in (21, 22, 23, 24))
        items = [(P1, u"Альфа товар прямой 1", 40, u"шт", 400.0, u"Прям"),     # 16 000, продаётся на Т02
                 (P2, u"Бета товар прямой 2", 30, u"шт", 200.0, u"Прям"),      # 6 000, продаётся на Т03
                 (P3, u"Гамма товар прямой 3", 3, u"шт", 10.0, u"Прям"),       # 30: мелочь на слабую точку - она выбывает
                 (KG, u"БІР Кег Тест Мин 0,5л (30)", 3, u"шт", 100.0, u"БІР")]  # продаётся только на Т12: та выбывает по порогу
        corr = os.path.join(tmp, CORR_DIR, u"Тест_Магазин.xlsx")
        _write_vyvoz_list_xlsx(corr, shop, day_s, items)
        g["fetch_route_history"] = lambda bcs, sh: {
            bc_key(b): {"shop_rc": "", "shop_dir": "2026-09-01", "net_rc": "", "net_dir": "", "sup": "Прям"} for b in bcs}
        st_ = [u"Т%02d" % i for i in range(1, 13)]
        g["pick_top_shops"] = lambda ua_stores, n=SPLIT_TOP_N, days=SPLIT_DAYS_SHOP: pd.DataFrame(
            {"store": st_[:n], "rev": ([1000.0, 300.0, 200.0, 150.0] + [100.0 - 5 * i for i in range(8)])[:n]})
        g["fetch_bc_sales"] = _mock_sales({(bc_key(P1), u"Т02"): 30.0, (bc_key(P2), u"Т03"): 20.0, (bc_key(P3), u"Т04"): 5.0,
                                           (bc_key(KG), u"Т12"): 9.0},
                                          {(bc_key(P1), u"Т02"): 10, (bc_key(P2), u"Т03"): 8, (bc_key(P3), u"Т04"): 3})
        ref0 = make_ref(pd.DataFrame(columns=["barcode", "product_name", "status", "supplier"]))
        info = run_split(d, corr, ref=ref0, auto_points=True)
        ck(info["ok"], "по сумме: распределение прошло: %s" % info.get("problems"))
        al_ = info.get("alloc")
        sums = {} if al_ is None else al_.groupby("addr")["sum"].sum().to_dict()
        ck(sums and all(v >= 5000.0 - 1e-6 for a_, v in sums.items() if a_ != SPLIT_RC_NAME),
           "по сумме: на точку не меньше 5000 грн: %s" % {k: round(v) for k, v in sums.items()})
        ck(abs(sum(sums.values()) - 22330.0) < 1e-6, "по сумме: весь товар распределён: %.2f" % sum(sums.values()))
        ck(u"Т12" not in sums and al_ is not None and float(al_[(al_["bc"] == KG) & (al_["addr"] == u"Т02")]["qty"].sum()) == 3.0,
           "по сумме: кег не уезжает на выбывшую точку, а целиком на оставшуюся: %s" % sorted(sums))
        ck(sorted(sums) == [u"Т02", u"Т03"], "по сумме: товар делится по спросу, а не сваливается на одну точку: %s" % sorted(sums))
    finally:
        g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"] = saved
        shutil.rmtree(tmp, ignore_errors=True)


def _selftest_split_demand(ck):
    """Точки по спросу на этот список (решение владельца 09.10.2026): слабая по кассе точка-лидер продаж ШК получает товар;
    один чек - не спрос; без своих продаж - по группе, без группы - по спросу точки на список; порог 5000 и в ручном режиме."""
    tmp = tempfile.mkdtemp(prefix="vyvoz_dem_")
    g = globals()
    saved = (g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"])
    try:
        d = Dirs(tmp)
        d.ensure()
        A, B, C, D = (_ean("4821005000%02d" % n) for n in (31, 32, 33, 34))
        items = [(A, u"Альфа тест 1", 50, u"шт", 200.0, u"Прям"),     # 10 000, продаётся только на Т09 (касса мала)
                 (B, u"Бета тест 1", 30, u"шт", 200.0, u"Прям"),      # 6 000, продаётся на Т01; на Т05 один чек на 3 шт
                 (C, u"Гамма тест 1", 10, u"шт", 100.0, u"Прям"),     # 1 000, ни ШК, ни группа не продаются
                 (D, u"Альфа тест 2", 20, u"шт", 100.0, u"Прям")]     # 2 000, сам не продаётся, группа «Альфа» - на Т09
        corr = os.path.join(tmp, CORR_DIR, u"Тест_Спрос.xlsx")
        _write_vyvoz_list_xlsx(corr, u"Тест Спрос", u"08.10.2026", items)
        g["fetch_route_history"] = lambda bcs, sh: {
            bc_key(b): {"shop_rc": "", "shop_dir": "2026-09-01", "net_rc": "", "net_dir": "", "sup": "Прям"} for b in bcs}
        st_ = [u"Т%02d" % i for i in range(1, 13)]
        g["pick_top_shops"] = lambda ua_stores, n=SPLIT_TOP_N, days=SPLIT_DAYS_SHOP: pd.DataFrame(
            {"store": st_[:n], "rev": ([1000.0, 300.0, 200.0, 150.0] + [100.0 - 5 * i for i in range(8)])[:n]})
        g["fetch_bc_sales"] = _mock_sales({(bc_key(A), u"Т09"): 60.0, (bc_key(B), u"Т01"): 40.0, (bc_key(B), u"Т05"): 3.0},
                                          {(bc_key(A), u"Т09"): 30, (bc_key(B), u"Т01"): 20, (bc_key(B), u"Т05"): 1})
        ref0 = make_ref(pd.DataFrame(columns=["barcode", "product_name", "status", "supplier"]))
        for mode in (True, False):
            info = run_split(d, corr, ref=ref0, auto_points=mode, top_n=3)
            al_ = info.get("alloc")
            ck(info["ok"] and al_ is not None, "спрос (%s): распределение прошло: %s" % (mode, info.get("problems")))
            if al_ is None:
                continue
            sums = {k: round(v, 2) for k, v in al_.groupby("addr")["sum"].sum().to_dict().items()}
            q = lambda a_, b_: float(al_[(al_["addr"] == a_) & (al_["bc"] == b_)]["qty"].sum())
            ck(q(u"Т09", A) == 50.0, "спрос (%s): точка с малой кассой, лидер продаж ШК, получила его весь: %s" % (mode, sums))
            ck(sums == {u"Т09": 12600.0, u"Т01": 6400.0}, "спрос (%s): Т09 = 10000 + 2000 (группа) + 600, Т01 = 6000 + 400: %s" % (mode, sums))
            ck(q(u"Т09", C) == 6.0 and q(u"Т01", C) == 4.0, "спрос (%s): без спроса на ШК и группу - по спросу точки на список 6/4" % mode)
            ck(info["demand"].get(u"Т05") == 0.0, "спрос (%s): один чек - не спрос: %s" % (mode, info["demand"].get(u"Т05")))
            ck(info["hows"].get(A) == u"ШК" and info["hows"].get(D) == u"группа" and info["hows"].get(C) == u"спрос на список",
               "спрос (%s): видно, как делился каждый ШК: %s" % (mode, info["hows"]))
            ck(any(u"Точек было 3, едет 2" in n_ for n_ in info.get("notes", [])),
               "спрос (%s): выбывшая по порогу точка показана в итоге: %s" % (mode, info.get("notes")))
    finally:
        g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"] = saved
        shutil.rmtree(tmp, ignore_errors=True)


def _selftest_split_update(ck):
    """Распределение исправленного листа и «Новый день» на синтетике: BigQuery подменяется, окон нет."""
    tmp = tempfile.mkdtemp(prefix="vyvoz_split_")
    g = globals()
    saved = (g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"])
    try:
        d = Dirs(tmp)
        d.ensure()
        shop, day_s = u"Тест Магазин", u"06.10.2026"
        B = {k: _ean("4821001000%02d" % n) for k, n in (("rc1", 11), ("rc2", 12), ("d1", 13), ("d2", 14), ("d3", 15))}
        items = [(B["rc1"], u"РЦ товар 1", 5, u"шт", 10.0, u"Пост"),
                 (B["rc2"], u"РЦ товар 2 (придёт на склад ЮА)", 3, u"шт", 20.0, u"Пост"),
                 (B["d1"], u"Прямой товар 1", 12, u"шт", 5.0, u"Пост"),
                 (B["d2"], u"Прямой весовой", 4.5, u"кг", 100.0, u"Пост"),
                 (B["d3"], u"Прямой (придёт на магазин ЮА)", 6, u"шт", 3.0, u"Пост")]
        corr = os.path.join(tmp, CORR_DIR, u"Тест_Магазин.xlsx")
        _write_vyvoz_list_xlsx(corr, shop, day_s, items)
        lst, shop0, day0, notes = read_corrected_list(corr)
        ck(shop0 == shop and day0 == "2026-10-06" and len(lst) == 5 and not notes,
           "исправленный лист читается: %s %s %d %s" % (shop0, day0, len(lst), notes))
        rc_keys = {bc_key(B["rc1"]), bc_key(B["rc2"])}
        g["fetch_route_history"] = lambda bcs, sh: {
            bc_key(b): ({"shop_rc": "2026-09-01", "shop_dir": "", "net_rc": "", "net_dir": "", "sup": "Пост"} if bc_key(b) in rc_keys
                        else {"shop_rc": "", "shop_dir": "2026-09-01", "net_rc": "", "net_dir": "", "sup": "Пост"}) for b in bcs}
        g["pick_top_shops"] = lambda ua_stores, n=SPLIT_TOP_N, days=SPLIT_DAYS_SHOP: pd.DataFrame(
            {"store": [u"Точка А", u"Точка Б", u"Точка В"][:n], "rev": [300.0, 200.0, 100.0][:n]})
        g["fetch_bc_sales"] = _mock_sales({})
        ref0 = make_ref(pd.DataFrame(columns=["barcode", "product_name", "status", "supplier"]))

        def rd(folder, f):
            if not os.path.isfile(os.path.join(folder, f)):
                return {}                                    # нет файла - пустой словарь: тест сообщит о расхождении, а не упадёт
            with open(os.path.join(folder, f), "rb") as fh:
                return {ln.split(";")[0]: float(ln.split(";")[1]) for ln in fh.read().decode("cp1251").split("\r\n") if ln}

        info = run_split(d, corr, ref=ref0, top_n=2, per_shop=(0.0, 0.0))
        sd = os.path.join(d.out, "2026-10-06", shop, SPLIT_OUT_DIR)
        ck(info["ok"] and os.path.isdir(sd), "распределение: ok и папка результата: %s" % info.get("problems"))
        files = sorted(f for f in os.listdir(sd) if f.endswith(".txt")) if os.path.isdir(sd) else []
        ck(files == sorted([u"Полевая-Склад.txt", u"Точка А.txt", u"Точка Б.txt"]), "по умолчанию файлы: РЦ + 2 точки: %s" % files)
        tot = {}
        for f in files:
            for k, q in rd(sd, f).items():
                tot[k] = round(tot.get(k, 0.0) + q, 3)
        ck(tot == {B["rc1"]: 5, B["rc2"]: 3, B["d1"]: 12, B["d2"]: 4.5, B["d3"]: 6}, "распределение: сумма по адресам = лист: %s" % tot)
        if files:
            ck(set(rd(sd, u"Полевая-Склад.txt")) == {B["rc1"], B["rc2"]}, "файл РЦ: только позиции РЦ")
            ck(rd(sd, u"Точка А.txt").get(B["d2"]) == 4.5 and B["d2"] not in rd(sd, u"Точка Б.txt"),
               "весовой 4,5 кг меньше двух минимумов: целиком лидеру")
            ck(rd(sd, u"Точка А.txt").get(B["d1"]) == 7 and rd(sd, u"Точка Б.txt").get(B["d1"]) == 5, "12 шт делятся 7/5 по доле оборота")
        info_nr = run_split(d, corr, ref=ref0, top_n=2, skip_rc=True, per_shop=(0.0, 0.0))
        ck(info_nr["ok"] and not os.path.exists(os.path.join(sd, u"Полевая-Склад.txt")) and
           os.path.isdir(os.path.join(sd, u"_прошлые_запуски")), "режим «без РЦ»: файла РЦ нет, прошлый запуск в архиве")
        info = run_split(d, corr, ref=ref0, top_n=2, per_shop=(0.0, 0.0))                     # файл РЦ снова на месте для проверки «Нового дня»
        ck(info["ok"] and os.path.isfile(os.path.join(sd, u"Полевая-Склад.txt")), "РЦ-файл по умолчанию возвращается")

        # --- «Новый день»: на складе ЮА появился товар РЦ (с нулевым остатком), на магазине ЮА - прямой ---
        wh = os.path.join(d.ua_wh, u"Состояние склада Полевая склад ЮА 07.10.2026.xlsx")
        _write_state_xlsx(wh, u"Полевая-Склад ЮА", [(B["rc2"], u"РЦ товар 2", 0, u"шт", 20.0), (_ean("999000000011"), u"Чужой", 4, u"шт", 1.0)],
                          shop_name=u"Полевая-Склад ЮА")
        shp = os.path.join(d.ua_stores, u"Состояние склада Тест Магазин ЮА 07.10.2026.xlsx")
        _write_state_xlsx(shp, u"Тест Магазин ЮА", [(B["d3"], u"Прямой", 1, u"шт", 3.0)])
        upd = run_update(d, corr, [wh], shop_paths=[shp], ref=ref0)
        ck(upd["ok"], "Новый день: ошибки %s" % upd["problems"])
        ck({bc_key(x["bc"]) for x in upd["removed"]} == {bc_key(B["rc2"]), bc_key(B["d3"])},
           "Новый день: убраны ровно 2 позиции: %s" % [x["bc"] for x in upd["removed"]])
        out = upd.get("out_dir", u"")
        cl = os.path.join(out, u"Тест_Магазин.xlsx")
        if os.path.isfile(cl):
            ws_ = load_workbook(cl)[SHEET_VYVOZ]
            nums = [ws_.cell(row=r_, column=1).value for r_ in range(5, 8)]
            ck(nums == [1, 2, 3] and ws_.cell(row=8, column=3).value == u"ИТОГО:", "очищенный лист: нумерация подряд, ИТОГО на месте: %s" % nums)
            ck(ws_.cell(row=8, column=4).value == 21.5 and ws_.cell(row=8, column=7).value == 560.0,
               "очищенный лист: ИТОГО пересчитан (21,5 ед. / 560 грн): %s / %s" % (ws_.cell(row=8, column=4).value, ws_.cell(row=8, column=7).value))
            a2 = ws_["A2"].value
            ck(u"Позиций: 3" in a2 and u"Единиц: 21.5" in a2 and u"Себестоимость: 560.0" in a2, "очищенный лист: шапка пересчитана: %s" % a2)
        else:
            ck(False, "очищенный лист не создан: %s" % out)
        td = os.path.join(out, SPLIT_OUT_DIR)
        if os.path.isdir(td):
            ck(set(rd(td, u"Полевая-Склад.txt")) == {B["rc1"]}, "очищенный файл РЦ: без пришедшего на склад ЮА")
            ck(B["d3"] not in rd(td, u"Точка А.txt") and B["d1"] in rd(td, u"Точка А.txt"), "очищенный файл точки: без пришедшего на магазин ЮА")
        else:
            ck(False, "папка очищенных txt не создана")
        ck(os.path.isfile(os.path.join(out, UPD_REMOVED_FILE)) and os.path.isfile(os.path.join(out, u"ЧИТАЙ_МЕНЯ.txt")),
           "Новый день: список удалённого и ЧИТАЙ_МЕНЯ в папке обновления")
        ck(len(rd(sd, u"Полевая-Склад.txt")) == 2, "исходные файлы ТСД не изменены")
        ck(not run_update(d, corr, [])["ok"], "Новый день без файла склада ЮА -> ошибка")
    finally:
        g["fetch_route_history"], g["pick_top_shops"], g["fetch_bc_sales"] = saved
        shutil.rmtree(tmp, ignore_errors=True)


def _write_state_xlsx(path, title, rows, shop_name=None):
    """Синтетическая выгрузка «Состояние склада» в формате Торгсофта (шапка, заголовки, подвал)."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = 'Состояние склада "%s"' % title
    ws["A2"] = "Состояние склада"
    for j, h in enumerate(["№", "Название товара", "Штрих-код", "Цена розничная", "Количество",
                           "Ед. изм.", "Себестоимость", "Склад"], 1):
        ws.cell(row=3, column=j, value=h)
    for i, (bc, name, qty, unit, cost) in enumerate(rows, 1):
        ws.cell(row=3 + i, column=1, value=i)
        ws.cell(row=3 + i, column=2, value=name)
        s = str(bc)
        ws.cell(row=3 + i, column=3, value=int(s) if s.isdigit() else s)
        ws.cell(row=3 + i, column=4, value=1.0)
        ws.cell(row=3 + i, column=5, value=qty)
        ws.cell(row=3 + i, column=6, value=unit)
        ws.cell(row=3 + i, column=7, value=cost)
        ws.cell(row=3 + i, column=8, value=shop_name or title)
    ws.cell(row=3 + len(rows) + 2, column=1, value="Исполнитель:   02.10.2026 12:00")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def _write_receipts_xlsx(path, rows):
    """Синтетические «Приходы» (Движение товара): ШК, название, отправитель, количество."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "Движение товара"
    ws["A2"] = "Перечень товара"
    for j, h in enumerate(["№", "Дата", "№ Док.", "Название товара", "Штрих-код", "Отправитель",
                           "Получатель", "Количество"], 1):
        ws.cell(row=3, column=j, value=h)
    for i, (bc, name, sender, qty) in enumerate(rows, 1):
        for j, v in enumerate([i, "2026-10-02", 100 + i, name, int(bc), sender, "Полевая-Склад ЮА", qty], 1):
            ws.cell(row=3 + i, column=j, value=v)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


# ======================= ОКНО И ЗАПУСК =======================

def open_path(p):
    try:
        if os.name == "nt":
            os.startfile(p)
        else:
            import subprocess
            subprocess.Popen(["xdg-open", p])
    except Exception:
        pass


def _result_text(info):
    lines = []
    if info.get("day_dir"):
        lines.append(u"Результат: %s" % info["day_dir"])
    if info.get("manual"):
        lines.append(u"РЕЖИМ: РУЧНОЙ ВЫБОР - сигареты и кеги в списке вывоза (причина обычного исключения в последней колонке); "
                     u"поставщики с галочкой, список ШК, расходники, кулинария - не вывозим")
    if info.get("excl_n") is not None:
        lines.append(u"Поставщиков в списке «не вывозить»: %d%s (меняется кнопкой «Поставщики: кого не вывозить»)"
                     % (info["excl_n"], (u", список сохранён " + info["excl_when"]) if info.get("excl_when") else u""))
    for s in info.get("shops", []):
        lines.append(u"%-26s вывезти %4d поз., %9s ед., %12s грн   проверить 1: %d, проверить 2: %d, исключено %d   %s"
                     % (s["shop"], s["vyvoz_pos"], s["vyvoz_units"], u"{:,.0f}".format(s["vyvoz_sum"]).replace(",", " "),
                        s["check1"], s["check2"], s["excl"], s["status"]))
    lines.append(u"")
    lines.append(u"Ошибок %d, предупреждений %d" % (len(R.errors), len(R.warns)))
    for st, nm, det in R.checks:
        if st in ("ERROR", "WARN"):
            lines.append(u"[%s] %s: %s" % (VERDICT_LABEL[st], nm, det))
    return u"\n".join(lines)


def open_suppliers_dialog(root, dirs, info=None, test_mode=False, on_save=None, test_hook=None):
    """Окно «Поставщики: кого не вывозить»: ГАЛОЧКА = НЕ вывозим (товар остаётся на полке и идёт в инвентаризацию ЮА),
    без галочки - вывозим. «Сохранить» пишет Справочник\\vyvoz_suppliers.csv: следующий расчёт (и exe, и .py) берёт его сам."""
    import tkinter as tk
    from tkinter import ttk

    known = known_suppliers(dirs, info)
    win = tk.Toplevel(root)
    win.withdraw()
    win.title(u"Поставщики: кого не вывозить")
    k = max(1.0, win.winfo_fpixels("1i") / 96.0)

    def px(v):
        return int(v * k)
    try:
        ttk.Style(win).layout("Accent.TButton")
        acc = "Accent.TButton"
    except tk.TclError:
        acc = "TButton"
    MUT, F = "#6b7280", "Segoe UI"
    on = {n: tk.BooleanVar(master=win, value=bool(d["excluded"])) for n, d in known.items()}
    order = sorted(known, key=lambda n: (-(known[n]["pos"] or 0), known[n]["name"].lower()))

    top = ttk.Frame(win, padding=(px(18), px(14), px(18), px(6)))
    top.pack(fill="x")
    n0_, when0_, path0_ = suppliers_file_info(dirs)
    ttk.Label(top, text=u"Поставщики: кого не вывозить", font=(F, 14, "bold")).pack(anchor="w")
    ttk.Label(top, text=u"ГАЛОЧКА СТОИТ - товар поставщика НЕ вывозим: он остаётся на полке и идёт в инвентаризацию ЮА. "
                        u"Без галочки - вывозим (новые поставщики тоже). Список хранится в файле и действует на все "
                        u"следующие расчёты, пока вы его не измените.",
              foreground=MUT, wraplength=px(600), justify="left").pack(anchor="w", pady=(px(2), px(4)))
    ttk.Label(top, text=u"Сейчас в файле: не вывозим %d (%s)\n%s"
                        % (n0_, (u"сохранён " + when0_) if when0_ else u"ещё не сохранялся", path0_),
              foreground=MUT, wraplength=px(600), justify="left").pack(anchor="w", pady=(0, px(10)))
    r1 = ttk.Frame(top)
    r1.pack(fill="x")
    ttk.Label(r1, text=u"Поиск:").pack(side="left", padx=(0, px(8)))
    sv = tk.StringVar(master=win)
    ent = ttk.Entry(r1, textvariable=sv)
    ent.pack(side="left", fill="x", expand=True)
    mode = tk.StringVar(master=win, value=u"Все")
    ttk.Combobox(r1, textvariable=mode, state="readonly", width=17,
                 values=(u"Все", u"Только не вывозим", u"Только вывозим")).pack(side="left", padx=(px(8), 0))
    r2 = ttk.Frame(top)
    r2.pack(fill="x", pady=(px(8), 0))
    cnt = ttk.Label(r2, text=u"", foreground=MUT)
    cnt.pack(side="right")

    hdr = ttk.Frame(win, padding=(px(18), px(4), px(34), px(2)))
    hdr.pack(fill="x")
    for i, (t, an) in enumerate(((u"Поставщик (галочка - не вывозим)", "w"), (u"Поз.", "e"), (u"Сумма, грн", "e"))):
        ttk.Label(hdr, text=t, font=(F, 9, "bold"), foreground=MUT).grid(row=0, column=i, sticky=an,
                                                                          padx=((px(28), 0) if i == 0 else 0))
    for w_ in (hdr,):
        w_.columnconfigure(0, weight=1)
        w_.columnconfigure(1, minsize=px(60))
        w_.columnconfigure(2, minsize=px(110))

    bt = ttk.Frame(win, padding=(px(18), px(8), px(18), px(14)))
    bt.pack(side="bottom", fill="x")
    lf = ttk.Frame(win, padding=(px(18), 0, px(18), 0))
    lf.pack(fill="both", expand=True)
    bg = ttk.Style(win).lookup("TFrame", "background") or "#fafafa"
    cv = tk.Canvas(lf, highlightthickness=0, borderwidth=0, background=bg)
    sb = ttk.Scrollbar(lf, orient="vertical", command=cv.yview)
    cv.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    cv.pack(side="left", fill="both", expand=True)
    inner = ttk.Frame(cv)
    wid = cv.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
    cv.bind("<Configure>", lambda e: cv.itemconfigure(wid, width=e.width))
    inner.columnconfigure(0, weight=1)
    inner.columnconfigure(1, minsize=px(60))
    inner.columnconfigure(2, minsize=px(110))

    def wheel(e):
        if cv.yview() != (0.0, 1.0):
            cv.yview_scroll(-1 if e.delta > 0 else 1, "units")
    win.bind("<MouseWheel>", wheel)

    def upd(*_a):
        off = sum(1 for v in on.values() if v.get())
        cnt.config(text=u"Не вывозим: %d из %d" % (off, len(on)))

    widgets = {}
    for n in order:
        d = known[n]
        cb = ttk.Checkbutton(inner, text=d["name"], variable=on[n], command=upd)
        l1 = ttk.Label(inner, text=(str(d["pos"]) if d["pos"] else u"-"), anchor="e")
        l2 = ttk.Label(inner, text=(u"{:,.0f}".format(float(d["sum"] or 0)).replace(",", u" ") if d["pos"] else u"-"), anchor="e")
        for w_ in (l1, l2):
            w_.bind("<Button-1>", lambda e, v=on[n]: (v.set(not v.get()), upd()))
        widgets[n] = (cb, l1, l2)

    def visible():
        q, m = sup_norm(sv.get()), mode.get()
        return [n for n in order if q in n and (m == u"Все" or (m == u"Только не вывозим") == bool(on[n].get()))]

    def layout(*_a):
        vis, i = set(visible()), 0
        for n in order:
            cb, l1, l2 = widgets[n]
            if n in vis:
                cb.grid(row=i, column=0, sticky="w", pady=px(1))
                l1.grid(row=i, column=1, sticky="e")
                l2.grid(row=i, column=2, sticky="e", padx=(0, px(4)))
                i += 1
            else:
                for w_ in (cb, l1, l2):
                    w_.grid_remove()
        cv.yview_moveto(0)
        upd()

    def set_vis(val):
        for n in visible():
            on[n].set((not on[n].get()) if val is None else val)
        upd()
    ttk.Button(r2, text=u"Не вывозить всех", command=lambda: set_vis(True)).pack(side="left")
    ttk.Button(r2, text=u"Очистить: вывозим всех", command=lambda: set_vis(False)).pack(side="left", padx=px(8))
    ttk.Button(r2, text=u"Инвертировать", command=lambda: set_vis(None)).pack(side="left")
    ttk.Label(bt, text=u"Кнопки действуют на видимых (с учётом поиска)", foreground=MUT).pack(side="left")

    def save():
        rules = {n: known[n]["name"] for n in on if on[n].get()}
        save_excluded_suppliers(dirs, rules)
        if on_save:
            on_save(len(rules))
        win.destroy()
    ttk.Button(bt, text=u"Отмена", command=win.destroy).pack(side="right")
    ttk.Button(bt, text=u"Сохранить", style=acc, command=save).pack(side="right", padx=(0, px(8)))
    sv.trace_add("write", layout)
    mode.trace_add("write", layout)
    win.bind("<Escape>", lambda e: win.destroy())
    layout()

    W = min(px(660), int(win.winfo_screenwidth() * 0.9))
    H = min(px(620), int(win.winfo_screenheight() * 0.85))
    try:
        root.update_idletasks()
        x = root.winfo_rootx() + max(0, (root.winfo_width() - W) // 2)
        y = root.winfo_rooty() + max(0, (root.winfo_height() - H) // 2)
    except Exception:
        x, y = 100, 100
    win.geometry("%dx%d+%d+%d" % (W, H, x, y))
    res = (len(visible()), sum(1 for v in on.values() if v.get()))
    if test_mode:
        win.update()
        if test_hook:
            test_hook({"on": on, "known": known, "set_vis": set_vis, "save": save})
        try:
            win.destroy()
        except tk.TclError:
            pass
        return res
    win.transient(root)
    win.deiconify()
    try:
        win.grab_set()
    except Exception:
        pass
    ent.focus_set()
    return res


def run_app(test_mode=False):
    """Окно в стиле «Расчёта заказа»: выгрузки, магазины, параметры, поставщики, запуск и итог таблицей."""
    import tkinter as tk
    from tkinter import ttk

    DIRS.ensure()
    app = tk.Tk()
    app.withdraw()
    app.title(u"Вывоз вне матрицы - переход на ЮА Маркет")
    try:
        app.iconbitmap(rz._res("raschet.ico"))
    except Exception:
        pass
    modern = False
    try:
        import sv_ttk                                  # тема Windows 11 (Fluent), как в «Расчёте заказа»
        sv_ttk.set_theme("light")
        modern = True
    except Exception:
        pass
    S = ttk.Style(app)
    if not modern:
        for th in ("vista", "winnative", "clam"):
            try:
                S.theme_use(th)
                break
            except Exception:
                continue
    BG = S.lookup("TFrame", "background") or "#f3f3f3"
    app.configure(bg=BG)
    MUT, RED, GRN, AMB = "#5d6b7a", "#c42b1c", "#0f7b3f", "#9a6700"
    H1, H2 = ("Segoe UI Semibold", 17), ("Segoe UI Semibold", 11)
    FS = ("Segoe UI", 9)
    ACC = "Accent.TButton" if modern else "Go.TButton"
    if not modern:
        S.configure("Go.TButton", background="#2f6fed", foreground="#ffffff", font=("Segoe UI Semibold", 10),
                    borderwidth=0, padding=(22, 9))
        S.map("Go.TButton", background=[("active", "#1f5bd0"), ("disabled", "#a9bde8")])
        S.configure("Card.TFrame", background="#ffffff", relief="solid", borderwidth=1)
    S.configure("Mut.TLabel", foreground=MUT, font=FS)
    S.configure("Red.TLabel", foreground=RED, font=FS)
    S.configure("Grn.TLabel", foreground=GRN, font=FS)
    S.configure("Amb.TLabel", foreground=AMB, font=FS)
    S.configure("H2.TLabel", font=H2)
    S.configure("Treeview", rowheight=24)
    app.geometry("1000x860")
    app.minsize(900, 720)
    st = {"found": {}, "info": None, "thread": None}

    wrap = ttk.Frame(app, padding=(22, 16, 22, 14))
    wrap.pack(fill="both", expand=True)
    ttk.Label(wrap, text=u"Вывоз вне матрицы", font=H1).pack(anchor="w")
    ttk.Label(wrap, text=u"Переход магазинов на ЮА Маркет: что вывезти на склад, что остаётся на полке и переводится на ЮА",
              style="Mut.TLabel").pack(anchor="w", pady=(2, 0))

    def card(title):
        c = ttk.Frame(wrap, style="Card.TFrame", padding=14)
        c.pack(fill="x", pady=(12, 0))
        ttk.Label(c, text=title, style="H2.TLabel").pack(anchor="w")
        return c

    # ---------- 1. выгрузки ----------
    c1 = card(u"1. Выгрузки Торгсофта")
    grid = ttk.Frame(c1)
    grid.pack(fill="x", pady=(8, 0))
    dir_rows = []
    for i, (nm, attr, need) in enumerate(((u"Магазины (Family)", "stores", True), (u"Склад ЮА", "ua_wh", True),
                                          (u"Приходы ЮА", "ua_receipts", True), (u"Магазины в базе ЮА", "ua_stores", False))):
        a = ttk.Label(grid, text=nm, width=20)
        b = ttk.Label(grid, text=u"", width=14)
        cc = ttk.Label(grid, text=getattr(DIRS, attr), style="Mut.TLabel")
        a.grid(row=i, column=0, sticky="w", pady=1)
        b.grid(row=i, column=1, sticky="w")
        cc.grid(row=i, column=2, sticky="w", padx=(6, 0))
        dir_rows.append((attr, need, b))
    row1 = ttk.Frame(c1)
    row1.pack(fill="x", pady=(8, 0))
    ttk.Button(row1, text=u"Открыть папку входов", command=lambda: open_path(DIRS.inp)).pack(side="left")
    ttk.Button(row1, text=u"Обновить", command=lambda: refresh()).pack(side="left", padx=8)

    # ---------- 2. магазины ----------
    c2 = card(u"2. Магазины для вывоза")
    tf = ttk.Frame(c2)
    tf.pack(fill="x", pady=(8, 0))
    tv = ttk.Treeview(tf, columns=("shop", "file", "time"), show="headings", selectmode="extended", height=5)
    for c, t, w, an in (("shop", u"Магазин", 220, "w"), ("file", u"Файл выгрузки", 520, "w"), ("time", u"Выгружено", 150, "center")):
        tv.heading(c, text=t)
        tv.column(c, width=w, anchor=an)
    sb = ttk.Scrollbar(tf, command=tv.yview)
    tv.config(yscrollcommand=sb.set)
    tv.pack(side="left", fill="x", expand=True)
    sb.pack(side="right", fill="y")
    row2 = ttk.Frame(c2)
    row2.pack(fill="x", pady=(8, 0))
    ttk.Button(row2, text=u"Выбрать все", command=lambda: tv.selection_set(tv.get_children())).pack(side="left")
    hint = ttk.Label(row2, text=u"Ctrl / Shift - несколько магазинов. Полевая не вывозится.", style="Mut.TLabel")
    hint.pack(side="left", padx=10)

    # ---------- 3. параметры ----------
    c3 = card(u"3. Параметры")
    rm_var = tk.BooleanVar(value=True)             # правило подтверждено 05.10.2026: в матрице, но на ЮА не было -> вывозим
    ttk.Checkbutton(c3, text=u"Убирать и то, что в матрице, но на ЮА не завозилось", variable=rm_var).pack(anchor="w", pady=(8, 0))
    st["rm"] = rm_var
    row3 = ttk.Frame(c3)
    row3.pack(fill="x", pady=(8, 0))
    sup_lbl = ttk.Label(row3, text=u"", style="Mut.TLabel")

    def sup_refresh(*_a):
        n = len(load_excluded_suppliers(DIRS))
        sup_lbl.config(text=u"Исключено из вывоза поставщиков: %d" % n)
    ttk.Button(row3, text=u"Поставщики для вывоза...",
               command=lambda: open_suppliers_dialog(app, DIRS, st["info"], on_save=sup_refresh)).pack(side="left")
    sup_lbl.pack(side="left", padx=10)

    def open_keep():
        load_keep_shk(DIRS)                            # создаст файл, если его нет
        open_path(os.path.join(DIRS.cache, KEEP_FILE))
    ttk.Button(row3, text=u"Список ШК «не вывозить» (Excel)...", command=open_keep).pack(side="right")
    ttk.Label(c3, text=u"Матрица берётся из листа Google «Ассортиментная матрица (полная)» (свежее BigQuery); "
                       u"если лист недоступен - из BigQuery.", style="Mut.TLabel", wraplength=900).pack(anchor="w", pady=(8, 0))

    # ---------- 4. распределение исправленного списка ----------
    c4 = card(u"4. Распределение исправленного списка (РЦ + точки)")
    row4 = ttk.Frame(c4)
    row4.pack(fill="x", pady=(8, 0))
    split_path = tk.StringVar(value=u"")
    split_top = tk.IntVar(value=SPLIT_TOP_N)
    ttk.Label(row4, text=u"Файл:").pack(side="left")
    ttk.Entry(row4, textvariable=split_path).pack(side="left", padx=6, fill="x", expand=True)

    def split_browse():
        from tkinter import filedialog
        init = os.path.join(DIRS.base, u"Корректировка_ЮА")
        fp = filedialog.askopenfilename(parent=app, title=u"Исправленный лист вывоза",
                                        initialdir=init if os.path.isdir(init) else DIRS.out,
                                        filetypes=[(u"Excel", "*.xlsx *.xlsm"), (u"Все файлы", "*.*")])
        if fp:
            split_path.set(fp)
    ttk.Button(row4, text=u"Обзор...", command=split_browse).pack(side="left")
    row4b = ttk.Frame(c4)
    row4b.pack(fill="x", pady=(8, 0))
    ttk.Label(row4b, text=u"Точек-получателей (1-%d):" % SPLIT_TOP_N).pack(side="left")
    ttk.Spinbox(row4b, from_=1, to=SPLIT_TOP_N, textvariable=split_top, width=5, state="readonly").pack(side="left", padx=6)
    split_status = ttk.Label(row4b, text=u"", style="Mut.TLabel")
    split_status.pack(side="left", padx=10)
    split_btn = ttk.Button(row4b, text=u"Распределить")
    split_btn.pack(side="right")
    split_st = {"res": None, "thread": None}

    def split_job(fp, n):
        try:
            split_st["res"] = run_split(DIRS, fp, top_n=n)
        except Exception as e:
            split_st["res"] = {"ok": False, "problems": [u"%s: %s" % (type(e).__name__, e)]}

    def split_poll():
        if split_st["thread"] is not None and split_st["thread"].is_alive():
            app.after(300, split_poll)
            return
        from tkinter import messagebox
        split_btn.config(state="normal")
        res = split_st["res"] or {}
        txt = _split_text(res) or u"Готово"
        if res.get("ok"):
            split_status.config(text=u"Готово: %d точек, файлов %d" % (res.get("top_n", 0), len(res.get("files", []))))
            messagebox.showinfo(u"Распределение", txt, parent=app)
        else:
            split_status.config(text=u"Есть ошибки - см. сообщение")
            messagebox.showwarning(u"Распределение", txt, parent=app)
        if res.get("out_dir") and os.path.isdir(res["out_dir"]):
            open_path(res["out_dir"])

    def split_start():
        fp = split_path.get().strip()
        if not fp or not os.path.isfile(fp):
            split_status.config(text=u"Выберите исправленный файл вывоза (xlsx).")
            return
        n = _clamp_top(split_top.get())
        split_btn.config(state="disabled")
        split_status.config(text=u"Распределяю на %d точек..." % n)
        split_st["res"] = None
        split_st["thread"] = threading.Thread(target=split_job, args=(fp, n), daemon=True)
        split_st["thread"].start()
        app.after(300, split_poll)
    split_btn.config(command=split_start)

    # ---------- запуск ----------
    act = ttk.Frame(wrap)
    act.pack(fill="x", pady=(14, 0))
    bar = ttk.Progressbar(act, mode="indeterminate", length=260)
    status = ttk.Label(act, text=u"", style="Mut.TLabel")
    status.pack(side="left")
    go = ttk.Button(act, text=u"Сформировать вывоз", style=ACC)
    go.pack(side="right")

    # ---------- итог ----------
    res_card = ttk.Frame(wrap, style="Card.TFrame", padding=14)
    ttk.Label(res_card, text=u"Результат", style="H2.TLabel").pack(anchor="w")
    rt = ttk.Treeview(res_card, columns=("shop", "vz", "vs", "inv", "wait", "st"), show="headings", height=4)
    for c, t, w, an in (("shop", u"Магазин", 190, "w"), ("vz", u"Вывезти, поз.", 100, "center"), ("vs", u"Вывезти, грн", 110, "center"),
                        ("inv", u"Остаётся, поз.", 110, "center"), ("wait", u"Ждёт решения", 110, "center"),
                        ("st", u"Статус", 280, "w")):
        rt.heading(c, text=t)
        rt.column(c, width=w, anchor=an)
    rt.pack(fill="x", pady=(8, 0))
    res_dir = ttk.Label(res_card, text=u"", style="Mut.TLabel")
    res_dir.pack(anchor="w", pady=(6, 0))
    res_warn = ttk.Label(res_card, text=u"", style="Amb.TLabel", wraplength=900, justify="left")
    res_warn.pack(anchor="w")
    rrow = ttk.Frame(res_card)
    rrow.pack(fill="x", pady=(8, 0))
    b_open = ttk.Button(rrow, text=u"Открыть результат")
    b_sum = ttk.Button(rrow, text=u"Сводка и проверки")
    b_open.pack(side="left")
    b_sum.pack(side="left", padx=8)

    def refresh():
        app.config(cursor="watch")
        app.update_idletasks()
        reset_report()
        try:
            found = scan_stores(DIRS)
        finally:
            app.config(cursor="")
        st["found"] = {s: found[s] for s in sorted(found) if not _is_polevaya(s)}
        tv.delete(*tv.get_children())
        for s in st["found"]:
            p = st["found"][s][0] if isinstance(st["found"][s], (tuple, list)) else u""
            try:
                tm = datetime.fromtimestamp(os.path.getmtime(p)).strftime("%d.%m %H:%M")
            except Exception:
                tm = u""
            tv.insert("", "end", iid=s, values=(s, os.path.basename(p), tm))
        tv.selection_set(tv.get_children())
        for attr, need, lab in dir_rows:
            n = len(list_xlsx(getattr(DIRS, attr)))
            lab.config(text=(u"файлов: %d" % n) if n else (u"нет файлов" if need else u"не нужны пока"),
                       foreground=(GRN if n else (RED if need else MUT)))
        sup_refresh()
        status.config(text=(u"Магазинов найдено: %d" % len(st["found"])) if st["found"] else u"В папке МАГАЗИНЫ нет выгрузок «Состояние склада».")

    def show_result(info):
        rt.delete(*rt.get_children())
        for sh in info.get("shops", []):
            rt.insert("", "end", values=(sh["shop"], sh["vyvoz_pos"], u"{:,.0f}".format(sh["vyvoz_sum"]).replace(",", u" "),
                                         sh.get("inv_pos", u""), sh.get("wait_pos", u""), sh["status"]))
        res_dir.config(text=(u"Папка результата: %s" % info["day_dir"]) if info.get("day_dir") else u"")
        errs = [u"%s: %s" % (nm, det) for stt, nm, det in R.checks if stt == "ERROR"]
        warns = [u"%s: %s" % (nm, det) for stt, nm, det in R.checks if stt == "WARN"]
        res_warn.config(text=u"\n".join(([u"ОШИБКИ: " + e for e in errs[:4]]) + ([u"Предупреждений: %d (см. «Сводка и проверки»)" % len(warns)] if warns else [])),
                        foreground=(RED if errs else AMB))
        b_open.config(command=lambda: open_path(info["day_dir"]) if info.get("day_dir") else None)
        b_sum.config(command=lambda: open_path(info["summary"]) if info.get("summary") and os.path.isfile(info["summary"]) else None)
        res_card.pack(fill="x", pady=(12, 0))

    def job(shops, rm):
        try:
            st["info"] = run_vyvoz(DIRS, shops=shops, remove_not_on_ua=rm)
        except Exception as e:                       # окно не должно молча умирать
            st["info"] = {"shops": [], "day_dir": u"", "ok": False}
            R.check("ERROR", u"Сбой программы", u"%s: %s" % (type(e).__name__, e))

    def poll():
        if st["thread"] is not None and st["thread"].is_alive():
            app.after(200, poll)
            return
        bar.stop()
        bar.pack_forget()
        go.config(state="normal")
        status.config(text=u"Готово." if (st["info"] or {}).get("ok") else u"Есть ошибки, см. ниже.")
        show_result(st["info"] or {})

    def start():
        sel = list(tv.selection())
        if not sel:
            status.config(text=u"Выберите хотя бы один магазин.")
            return
        go.config(state="disabled")
        status.config(text=u"Считаю... (выгрузки читаются, матрица берётся из листа Google / BigQuery)")
        bar.pack(side="left", padx=12)
        bar.start(12)
        st["thread"] = threading.Thread(target=job, args=(sel, bool(st["rm"].get())), daemon=True)
        st["thread"].start()
        app.after(200, poll)
    go.config(command=start)
    refresh()
    try:
        rz._win11_chrome(app)                          # скруглённые углы и цвет заголовка как у окон Windows 11
    except Exception:
        pass
    app.deiconify()
    if test_mode:
        app.update()
        app.destroy()
        return 0
    app.mainloop()
    return 0


# ============ РАСПРЕДЕЛЕНИЕ ИСПРАВЛЕННОГО СПИСКА ВЫВОЗА: РЦ -> склад, прямые -> топ-N точек ============
# Лист «Вывезти на склад» правится вручную (строки удаляют) и читается обратно. Каждая позиция уходит либо на РЦ
# (Полевая-Склад), либо раздаётся самым «боевым» точкам, которые на ЮА не переходят. Результат: по txt для ТСД на адрес
# (имя файла = адрес, строка «штрих-код;количество») + xlsx для проверки.

SPLIT_RC_NAME    = u"Полевая-Склад"   # адрес для позиций, что идут на РЦ
SPLIT_TOP_N      = 10                 # на сколько самых «боевых» точек делим прямые поставки
SPLIT_DAYS_SHOP  = 30                 # оборот точки за N дней (turnover_transactions, сумма чеков)
SPLIT_DAYS_BC    = 60                 # продажи ШК на точке за N дней
SPLIT_BC_WEIGHT  = 0.7                # как DIST_BC_WEIGHT в raschet_zakaza: вес продаж именно этого ШК в доле точки
SPLIT_MIN        = 3                  # как DIST_NEW_MIN: меньше этого на точку не везём (шт или кг)
SPLIT_ALIVE_DAYS = 7                  # точка без продаж дольше N дней считается закрытой
SPLIT_HORIZON_DAYS = 30               # спрос точки на список: не больше, чем она продаст ШК за N дней (по продажам за SPLIT_DAYS_BC)
SPLIT_DEMAND_MIN_CHECKS = 2           # штучный ШК: меньше N чеков на точке за SPLIT_DAYS_BC дн. - случайность, не спрос
SPLIT_KEG_DAYS   = 90                 # кег без объёма в названии: объём = самый частый приход одного кега по всей сети за N дней
SPLIT_OUT_DIR    = u"ТСД_по_адресам"
CORR_DIR         = u"Корректировка_ЮА"   # сюда кладётся вручную исправленный лист вывоза
TURNOVER_TABLE   = BQ_DATASET + ".turnover_transactions"
TRANSFER_TABLE   = BQ_DATASET + ".transfer_all_transactions"
UA_STORES_FILE   = u"vyvoz_ua_stores.csv"
ROUTE_SUP_FILE   = u"vyvoz_route_suppliers.csv"
KIND_RC, KIND_SHOPS = u"РЦ", u"Точки"


def _clamp_top(n):
    """Сколько точек-получателей: 1..SPLIT_TOP_N (по умолчанию SPLIT_TOP_N)."""
    try:
        n = int(n) if n not in (None, "") else SPLIT_TOP_N
    except (TypeError, ValueError):
        n = SPLIT_TOP_N
    return max(1, min(SPLIT_TOP_N, n))
DEFAULT_UA_STORES = [
    (u"Грозненська 38", u"переходит на ЮА (в базе ЮА: Болградська 38 ЮА); список от 06.10.2026"),
    (u"Байрона 156", u"список от 06.10.2026"), (u"Зернова 6/5", u"список от 06.10.2026"),
    (u"Ньютона 111", u"список от 06.10.2026"), (u"Байрона 163", u"список от 06.10.2026"),
    (u"Байрона 138/1", u"список от 06.10.2026"), (u"Качанівська 19", u"источник вывоза; список от 06.10.2026"),
    (u"Полевая-Магазин", u"список от 06.10.2026"), (u"Ньютона 102", u"список от 06.10.2026"),
    (u"Полевая-Склад ЮА", u"склад ЮА"),
]
DEFAULT_ROUTE_SUP = [
    (u"Корона", KIND_RC, u"решение 06.10.2026: поставщик сменил доставку с прямой на РЦ"),
    (u"СТВ Схід (Корона)", KIND_RC, u"то же (Корона)"),
    (u"ТОВ СТВ Дистрибюшн (Корона)", KIND_RC, u"то же (Корона)"),
    (u"Баядера", KIND_RC, u"решение 06.10.2026"),
    (u"Монжар", KIND_RC, u"решение 06.10.2026"),
    (u"БІР (Пиво)", KIND_SHOPS, u"решение 06.10.2026: доставка на точки"),
    (u"Авангард Дистрибуції", KIND_SHOPS, u"решение 06.10.2026: доставка на точки"),
    (u"Шейки", KIND_SHOPS, u"решение 06.10.2026: доставка на точки"),
]


def _csv_default(dirs, fname, rows, cols):
    path = os.path.join(dirs.cache, fname)
    if not os.path.isfile(path):
        os.makedirs(dirs.cache, exist_ok=True)
        pd.DataFrame(rows, columns=cols).to_csv(path, index=False, encoding="utf-8-sig")
    return pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")


def load_ua_stores(dirs):
    """Магазины, которые на ЮА перешли или переходят (получателями раздачи не бывают). Справочник\\vyvoz_ua_stores.csv"""
    df = _csv_default(dirs, UA_STORES_FILE, DEFAULT_UA_STORES, [u"Магазин", u"Комментарий"])
    return set(_skey(x) for x in df.iloc[:, 0] if str(x).strip())


def load_route_suppliers(dirs):
    """Поставщик -> «РЦ» | «Точки» (перекрывает историю). Справочник\\vyvoz_route_suppliers.csv -> {sup_norm: вид}"""
    df = _csv_default(dirs, ROUTE_SUP_FILE, DEFAULT_ROUTE_SUP, [u"Поставщик", u"Куда (РЦ или Точки)", u"Комментарий"])
    out = {}
    for s, k in zip(df.iloc[:, 0], df.iloc[:, 1]):
        kk = str(k).strip().lower()
        if sup_norm(s) and kk in (u"рц", u"точки"):
            out[sup_norm(s)] = KIND_RC if kk == u"рц" else KIND_SHOPS
    return out


def read_corrected_list(path):
    """Исправленный вручную лист «Вывезти на склад» -> (DataFrame bc,name,qty,unit,cost,sum; магазин; день ГГГГ-ММ-ДД; замечания)."""
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET_VYVOZ] if SHEET_VYVOZ in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    shop = day = None
    for r in rows[:4]:
        t = str(r[0] or u"")
        m = re.match(u"Вывоз вне матрицы:\\s*(.+?)\\s*->", t)
        if m:
            shop = m.group(1).strip()
        m = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", t)
        if m and day is None:
            day = u"%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
    hi = next((i for i, r in enumerate(rows) if r and any(str(c).strip() == u"Штрих-код" for c in r if c is not None)), None)
    if hi is None:
        raise ValueError(u"в файле нет строки заголовков со «Штрих-код»")
    hdr = [str(c).strip() if c is not None else u"" for c in rows[hi]]

    def ix(name):
        if name not in hdr:
            raise ValueError(u"в файле нет колонки «%s»" % name)
        return hdr.index(name)
    i_bc, i_nm, i_q = ix(u"Штрих-код"), ix(u"Название товара"), ix(u"Количество")
    i_u = hdr.index(u"Ед. изм.") if u"Ед. изм." in hdr else None
    i_c = hdr.index(u"Себестоимость") if u"Себестоимость" in hdr else None
    notes, data = [], {}
    for r in rows[hi + 1:]:
        if not r or r[i_bc] in (None, u"") or u"ИТОГО" in str(r[i_nm] or u"").upper():
            continue
        bc = rz.fmt_barcode(r[i_bc])
        try:
            q = float(r[i_q])
        except (TypeError, ValueError):
            notes.append(u"ШК %s: количество «%s» не число, строка пропущена" % (bc, r[i_q]))
            continue
        if not RE_BC_OK.match(bc):
            notes.append(u"ШК %s: не 6-14 цифр, строка пропущена" % bc)
            continue
        if q <= 0:
            notes.append(u"ШК %s: количество %s, строка пропущена" % (bc, q))
            continue
        cost = r[i_c] if i_c is not None and isinstance(r[i_c], (int, float)) else 0.0
        unit = str(r[i_u]).strip() if i_u is not None and r[i_u] not in (None, u"") else u""
        if bc in data:
            notes.append(u"ШК %s: строка повторяется, количества сложены" % bc)
            data[bc]["qty"] += q
        else:
            data[bc] = {"bc": bc, "name": str(r[i_nm] or u""), "qty": q, "unit": unit, "cost": float(cost)}
    df = pd.DataFrame(list(data.values()), columns=["bc", "name", "qty", "unit", "cost"])
    df["qty"] = df["qty"].round(3)
    df["sum"] = (df["qty"] * df["cost"]).round(2)
    return df, shop, day, notes


def decide_route(shop_rc, shop_dir, net_rc, net_dir, sup_kind=None):
    """Куда везти позицию -> (вид, основание). Решение пользователя 06.10.2026:
    РЦ = отправитель «Полевая-Склад» в перемещениях, прямая = приход от поставщика; нет истории вообще -> РЦ.
    Порядок: поставщик из справочника -> последнее поступление на этот магазин -> сеть (любой РЦ-приход = РЦ) -> РЦ."""
    if sup_kind:
        return sup_kind, u"поставщик в справочнике маршрутов"
    if shop_rc or shop_dir:
        if shop_rc and (not shop_dir or shop_rc >= shop_dir):
            return KIND_RC, u"по истории магазина: последнее поступление с РЦ" + (u" (были и прямые, взято последнее)" if shop_dir else u"")
        return KIND_SHOPS, u"по истории магазина: последнее поступление прямое" + (u" (были и с РЦ, взято последнее)" if shop_rc else u"")
    if net_rc:
        return KIND_RC, u"по истории сети: были перемещения с РЦ"
    if net_dir:
        return KIND_SHOPS, u"по истории сети: только приходы от поставщика"
    return KIND_RC, u"нет истории нигде: РЦ (разкомплектация, решение пользователя)"


def _doc_d(x):
    """Дата документа ГГГГ-ММ-ДД; у документа без даты (NULL в BigQuery) - условная старая, чтобы факт поступления не терялся."""
    x = str(x)
    return x if re.match(r"^\d{4}-\d{2}-\d{2}$", x) else u"0001-01-01"


def fetch_route_history(bcs, shop):
    """BigQuery -> {ключ ШК: dict(shop_rc, shop_dir, net_rc, net_dir, sup)} (даты строками ГГГГ-ММ-ДД или '')"""
    keys = sorted(set(bc_key(b) for b in bcs))
    lst = u",".join(u"'%s'" % k for k in keys)
    ok1, tr, _ = _bq_query(u"SELECT LTRIM(barcode,'0') k, store, CAST(MAX(doc_date) AS STRING) d FROM `%s` "
                           u"WHERE sender='%s' AND LTRIM(barcode,'0') IN (%s) GROUP BY 1,2" % (TRANSFER_TABLE, SPLIT_RC_NAME, lst))
    ok2, inc, _ = _bq_query(u"SELECT LTRIM(barcode,'0') k, store, CAST(MAX(doc_date) AS STRING) d, "
                            u"ARRAY_AGG(supplier ORDER BY doc_date DESC LIMIT 1)[OFFSET(0)] sup FROM `%s` "
                            u"WHERE quantity>0 AND store NOT LIKE '%% ЮА' AND LTRIM(barcode,'0') IN (%s) GROUP BY 1,2" % (INCOMING_TABLE, lst))
    if not (ok1 and ok2):
        raise RuntimeError(tr if not ok1 else inc)
    me = _skey(shop)
    h = {k: {"shop_rc": u"", "shop_dir": u"", "net_rc": u"", "net_dir": u"", "sup": u"", "_sd": u""} for k in keys}
    for k, st, d in zip(tr["k"], tr["store"], tr["d"]):
        x = h.get(str(k))
        if x is None:
            continue
        d = _doc_d(d)
        x["net_rc"] = max(x["net_rc"], d)
        if _skey(st) == me and d > u"0001-01-01":        # по самому магазину документ без даты доказательством не считается
            x["shop_rc"] = max(x["shop_rc"], d)
    for k, st, d, sup in zip(inc["k"], inc["store"], inc["d"], inc["sup"]):
        x = h.get(str(k))
        if x is None:
            continue
        d = _doc_d(d)
        if _skey(st) == _skey(SPLIT_RC_NAME):     # поставщик -> РЦ: путь через РЦ, не прямая поставка («Полевая 83» - это магазин, не РЦ)
            x["net_rc"] = max(x["net_rc"], d)
        else:
            x["net_dir"] = max(x["net_dir"], d)
            if _skey(st) == me and d > u"0001-01-01":
                x["shop_dir"] = max(x["shop_dir"], d)
        if d >= x["_sd"] and str(sup).strip() and str(sup).lower() not in ("nan", "none"):
            x["_sd"], x["sup"] = d, str(sup).strip()
    return h


def pick_top_shops(ua_stores, n=SPLIT_TOP_N, days=SPLIT_DAYS_SHOP):
    """Самые «боевые» точки по сумме чеков за последние N дней данных, без перешедших на ЮА и без закрытых -> DataFrame store, rev."""
    ok, df, _ = _bq_query(u"SELECT store, SUM(check_amount) AS rev, CAST(MAX(DATE(transaction_datetime)) AS STRING) AS last_d "
                          u"FROM `%s` WHERE transaction_datetime >= DATETIME_SUB((SELECT MAX(transaction_datetime) FROM `%s`), INTERVAL %d DAY) GROUP BY 1"
                          % (TURNOVER_TABLE, TURNOVER_TABLE, days))
    if not ok:
        raise RuntimeError(df)
    df["rev"] = pd.to_numeric(df["rev"], errors="coerce").fillna(0.0)
    last = df["last_d"].max()
    lim = (datetime.strptime(last, "%Y-%m-%d") - pd.Timedelta(days=SPLIT_ALIVE_DAYS)).strftime("%Y-%m-%d")
    keep = [(_skey(s) not in ua_stores) and (not _is_polevaya(s)) and (not str(s).lower().endswith(u" юа")) and (d >= lim)
            for s, d in zip(df["store"], df["last_d"])]
    return df[keep].sort_values("rev", ascending=False).head(n).reset_index(drop=True)


def fetch_bc_sales(bcs, stores, days=SPLIT_DAYS_BC, checks=None):
    """Продажи каждого ШК на выбранных точках за N дней -> {(ключ ШК, точка): количество}.
    checks (dict) - заполняется числом чеков {(ключ ШК, точка): чеков}."""
    keys = sorted(set(bc_key(b) for b in bcs))
    ok, df, _ = _bq_query(u"SELECT LTRIM(barcode,'0') k, store, SUM(quantity) q, COUNT(DISTINCT transaction_id) n FROM `%s` WHERE transaction_datetime >= "
                          u"DATETIME_SUB((SELECT MAX(transaction_datetime) FROM `%s`), INTERVAL %d DAY) AND store IN (%s) "
                          u"AND LTRIM(barcode,'0') IN (%s) GROUP BY 1,2"
                          % (TURNOVER_TABLE, TURNOVER_TABLE, days, u",".join(u"'%s'" % s.replace("'", "\\'") for s in stores),
                             u",".join(u"'%s'" % k for k in keys)))
    if not ok:
        raise RuntimeError(df)
    if checks is not None:
        checks.update({(str(k), str(s)): int(n) for k, s, n in zip(df["k"], df["store"], df["n"])})
    return {(str(k), str(s)): max(float(q), 0.0) for k, s, q in zip(df["k"], df["store"], df["q"])}


def fetch_keg_sizes(bcs, days=SPLIT_KEG_DAYS):
    """Кеги без объёма в названии: самый частый приход одного кега по ВСЕЙ сети за N дней -> {ключ ШК: порций в кеге}"""
    keys = sorted(set(bc_key(b) for b in bcs))
    ok, df, _ = _bq_query(u"SELECT LTRIM(barcode,'0') k, quantity q, COUNT(*) n FROM `%s` WHERE quantity >= 10 AND "
                          u"doc_date >= DATE_SUB((SELECT MAX(doc_date) FROM `%s`), INTERVAL %d DAY) AND LTRIM(barcode,'0') IN (%s) "
                          u"GROUP BY 1,2" % (INCOMING_TABLE, INCOMING_TABLE, days, u",".join(u"'%s'" % k for k in keys)))
    if not ok:
        raise RuntimeError(df)
    out = {}
    for k, q, n in sorted(zip(df["k"], df["q"], df["n"]), key=lambda x: (-int(x[2]), float(x[1]))):
        out.setdefault(str(k), float(q))
    return out


def shop_shares(bc, shops, rev, sales):
    """Доля каждой точки для ШК: SPLIT_BC_WEIGHT * продажи ШК + остальное * оборот точки (как DIST_BC_WEIGHT в raschet_zakaza).
    Продаж ШК на точках нет -> только оборот точки. -> Series по точкам, сумма 1, по убыванию."""
    tot_rev = float(sum(rev[s] for s in shops))
    base = pd.Series({s: rev[s] / tot_rev for s in shops})
    sb = pd.Series({s: sales.get((bc_key(bc), s), 0.0) for s in shops})
    if float(sb.sum()) > 0:
        base = SPLIT_BC_WEIGHT * (sb / float(sb.sum())) + (1.0 - SPLIT_BC_WEIGHT) * base
    return (base / float(base.sum())).sort_values(ascending=False, kind="stable")


def allocate_qty(total, shares, step, min_qty=SPLIT_MIN):
    """Делит total по точкам пропорционально shares, кратно step, сумма сходится ровно.
    Меньше min_qty на точку не везём: точек берётся не больше total // min_qty (минимум одна, лидер по доле);
    хвост меньше шага уходит самой крупной доле. -> {точка: количество}"""
    total = round(float(total), 3)
    nmax = int(total // max(min_qty, step) + 1e-9)
    use = shares.head(max(1, min(nmax, len(shares))))
    use = use / float(use.sum())
    while True:
        units = int(total / step + 1e-9)
        raw = use * units
        base = raw.apply(lambda x: int(x + 1e-9))
        rest = units - int(base.sum())
        order = (raw - base).sort_values(ascending=False, kind="stable").index
        for s in order[:rest]:
            base[s] += 1
        qty = (base * step).round(3)
        qty[qty.index[0]] = round(qty.iloc[0] + (total - float(qty.sum())), 3)   # хвост меньше шага -> лидеру
        small = [s for s in qty.index if qty[s] < min_qty - 1e-9]
        if not small or len(use) == 1:
            break
        use = use.drop(small[-1])
        use = use / float(use.sum())
    return {s: float(q) for s, q in qty.items() if q > 1e-9}


def _fit_widths(ws, maxw=60):
    for col in ws.columns:
        w = max(len(str(c.value if c.value is not None else u"")) for c in col) + 2
        ws.column_dimensions[col[0].column_letter].width = min(maxw, max(10, w))
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def is_keg(name):
    """Кег («БІР Кег ...»): на РЦ не везём, делим по точкам, где этот вид продаётся."""
    return bool(re.search(u"(^|\\s)кег", nname(name)))


def keg_size_from_name(name):
    """Порций в кеге по названию: «0,5л (30)» - кег 30 л, порция 0,5 л -> 60 (решение пользователя 09.10.2026: в скобках объём кега).
    Нет скобок или порции -> None."""
    t = str(name or u"").lower()
    v = re.search(r"\(\s*(\d+(?:[.,]\d+)?)\s*(?:л)?\s*\)", t)
    p = re.search(r"(\d+(?:[.,]\d+)?)\s*л", re.sub(r"\([^)]*\)", u" ", t))
    if not v or not p:
        return None
    vol, por = float(v.group(1).replace(",", ".")), float(p.group(1).replace(",", "."))
    return round(vol / por, 3) if vol > 0 and 0 < por < vol else None


def keg_pieces(qty, size):
    """Позиция кега -> куски, которые не делятся: целые кеги по size порций, затем начатый (остаток).
    Объём неизвестен -> вся позиция одним куском."""
    qty = round(float(qty), 3)
    if not size or size <= 0:
        return [qty]
    n = int(qty / size + 1e-9)
    rest = round(qty - n * size, 3)
    return [float(size)] * n + ([rest] if rest > 1e-9 else [])


def keg_alloc(bc, pieces, pts, sales, dem):
    """Куски кега целиком по выбранным точкам (решение пользователя 09.10.2026: кег по литрам не делим): каждый кусок - точке
    с наибольшим неудовлетворённым спросом на ЭТОТ ШК (продажи за SPLIT_HORIZON_DAYS дн. минус уже отданное), целые первыми.
    ШК на выбранных точках не продавался - вся позиция точке с наибольшим спросом на список. -> ({точка: количество}, продавался ли)"""
    k_ = SPLIT_HORIZON_DAYS / float(SPLIT_DAYS_BC)
    kb = bc_key(bc)
    need = {s: sales.get((kb, s), 0.0) * k_ for s in pts if sales.get((kb, s), 0.0) > 0}
    if not need:
        s0 = max(pts, key=lambda s: (dem.get(s, 0.0), -pts.index(s)))
        return {s0: round(sum(pieces), 3)}, False
    out = {}
    for pc in pieces:
        s = max(need, key=lambda s_: (need[s_], -pts.index(s_)))
        need[s] -= pc
        out[s] = round(out.get(s, 0.0) + pc, 3)
    return out, True


def clean_sales(sales, checks, free_keys):
    """Продажи без случайных: штучный ШК меньше SPLIT_DEMAND_MIN_CHECKS чеков на точке спросом не считается.
    Весовой товар и кеги (free_keys) - без фильтра; нет числа чеков - берётся количество."""
    return {kk: v for kk, v in sales.items()
            if v > 0 and (kk[0] in free_keys or checks.get(kk, v) >= SPLIT_DEMAND_MIN_CHECKS)}


def rank_by_demand(lst, stores, rev, sales):
    """Спрос точки на ЭТОТ список: сколько грн товара, который едет на точки, она продаст за SPLIT_HORIZON_DAYS дн.
    (не больше, чем есть в списке). Порядок - по спросу, при равенстве - по обороту. -> (список точек, {точка: грн})"""
    k_ = SPLIT_HORIZON_DAYS / float(SPLIT_DAYS_BC)
    q_by, c_by = {}, {}
    for b, q, c, kd in zip(lst["bc"], lst["qty"], lst["cost"], lst["kind"]):
        if kd == KIND_SHOPS:
            kb = bc_key(b)
            q_by[kb] = q_by.get(kb, 0.0) + float(q)
            c_by[kb] = float(c)
    dem = {s: 0.0 for s in stores}
    for (kb, s), v in sales.items():
        if kb in q_by and s in dem:
            dem[s] += min(v * k_, q_by[kb]) * c_by[kb]
    return sorted(stores, key=lambda s: (-dem[s], -rev.get(s, 0.0))), dem


def group_sales(lst, sales):
    """Продажи групп (group_of: первое слово названия) товаров списка по точкам, грн -> {(группа, точка): грн}"""
    gc = {bc_key(b): (group_of(nm), max(float(c), 0.01)) for b, nm, c in zip(lst["bc"], lst["name"], lst["cost"])}
    out = {}
    for (kb, s), v in sales.items():
        if kb in gc:
            g, c = gc[kb]
            out[(g, s)] = out.get((g, s), 0.0) + v * c
    return out


def demand_shares(r, pts, rev, sales, grp, dem):
    """Доли выбранных точек для ШК: по продажам ЭТОГО ШК; нет - по продажам его группы; нет - по спросу точки на список;
    нет и его - по обороту. Сверх SPLIT_HORIZON_DAYS дн. едет тем же точкам (решение пользователя 09.10.2026).
    -> (Series по точкам, сумма 1, по убыванию; как делили)"""
    kb, g = bc_key(r.bc), group_of(r.name)
    for how, d in ((u"ШК", [(s, sales.get((kb, s), 0.0)) for s in pts]),
                   (u"группа", [(s, grp.get((g, s), 0.0)) for s in pts]),
                   (u"спрос на список", [(s, dem.get(s, 0.0)) for s in pts])):
        d = [(s, v) for s, v in d if v > 0]
        if d:
            sb = pd.Series(dict(d), dtype=float)
            return (sb / float(sb.sum())).sort_values(ascending=False, kind="stable"), how
    return shop_shares(r.bc, pts, rev, {}), u"оборот"


def run_split(dirs, path, shop=None, day=None, ref=None, top_n=None, skip_rc=False, auto_points=False, per_shop=(5000.0, 7000.0)):
    """Исправленный лист вывоза -> txt для ТСД по адресам + xlsx для проверки. -> dict(ok, out_dir, files, problems, summary)
    Точки - по спросу на товары ЭТОГО списка (решение владельца 09.10.2026), не меньше per_shop[0] грн на точку в обоих режимах."""
    reset_report()
    info = {"ok": False, "problems": [], "files": [], "summary": []}
    try:
        lst, shop0, day0, notes = read_corrected_list(path)
    except Exception as e:
        info["problems"].append(u"Файл не прочитан: %s" % e)
        return info
    shop = shop or shop0
    day = day or day0 or date.today().strftime("%Y-%m-%d")
    if not shop:
        info["problems"].append(u"Не определён магазин-источник: он берётся из заголовка файла («Вывоз вне матрицы: <магазин> -> ...») или из --shop")
        return info
    info["notes"] = notes
    if lst.empty:
        info["problems"].append(u"В файле нет позиций")
        return info
    keg_bcs = [b for b, nm_ in zip(lst["bc"], lst["name"]) if is_keg(nm_)]
    keg_size = {bc_key(b): keg_size_from_name(nm_) for b, nm_ in zip(lst["bc"], lst["name"]) if is_keg(nm_)}
    try:
        ua_stores, rsup = load_ua_stores(dirs), load_route_suppliers(dirs)
        if ref is None:
            fr = load_reference(dirs)
            ref = make_ref(fr["matrix"], fr["recode"], fr["coffee"], None, fr["hist"]) if fr else None
        hist = fetch_route_history(lst["bc"], shop)
        top_n = SPLIT_TOP_N if auto_points else _clamp_top(top_n)       # auto_points: берём 10 первых по спросу, нужное число отрежем ниже
        info["top_n"] = top_n
        alive = pick_top_shops(ua_stores, n=10 ** 6)                    # все живые точки не на ЮА: среди них ищем спрос на список
        rev = dict(zip(alive["store"], alive["rev"]))
        checks = {}
        sales_raw = fetch_bc_sales(lst["bc"], list(alive["store"]), checks=checks)
        no_vol = [b for b in keg_bcs if not keg_size[bc_key(b)]]
        keg_net = fetch_keg_sizes(no_vol) if no_vol else {}
    except Exception as e:
        info["problems"].append(u"Не получены данные из BigQuery: %s" % e)
        return info
    free_keys = set(bc_key(b) for b, u_, q_, nm_ in zip(lst["bc"], lst["unit"], lst["qty"], lst["name"])
                    if u_ == u"кг" or abs(float(q_) - round(float(q_))) > 1e-9 or is_keg(nm_))
    sales = clean_sales(sales_raw, checks, free_keys)
    keg_src = {}
    for b in keg_bcs:
        kb = bc_key(b)
        if keg_size[kb]:
            keg_src[kb] = u"объём по названию"
        elif keg_net.get(kb):
            keg_size[kb], keg_src[kb] = keg_net[kb], u"объём по приходам сети за %d дн." % SPLIT_KEG_DAYS
        else:
            keg_src[kb] = u"объём не найден - вся позиция одной точке, проверьте"

    routes = []
    for r in lst.itertuples(index=False):
        k = bc_key(r.bc)
        h = hist.get(k, {})
        sup = u""
        if ref is not None:
            mk = k if k in ref.matrix else _resolve(ref, k)
            sup = (ref.matrix[mk][2] if mk in ref.matrix else u"") or ref.last_sup.get(k) or ref.last_sup.get(_resolve(ref, k), u"")
        sup = sup or h.get("sup", u"")
        kind, why = decide_route(h.get("shop_rc", u""), h.get("shop_dir", u""), h.get("net_rc", u""), h.get("net_dir", u""),
                                 rsup.get(sup_norm(sup)) if sup else None)
        if why.startswith(u"поставщик"):
            why += u": " + sup
        if is_keg(r.name):                                  # кеги не на РЦ: целиком на выбранные точки, где этот ШК продаётся
            sold = any(v_ > 0 for (kk_, _s), v_ in sales_raw.items() if kk_ == k)
            kind = KIND_SHOPS
            why = (u"кег: целиком на выбранные точки, где этот ШК продавался за %d дн." % SPLIT_DAYS_BC if sold else
                   u"кег: этот ШК нигде не продавался за %d дн. - на точку с наибольшим спросом на список, проверьте" % SPLIT_DAYS_BC)
        routes.append((kind, why, sup or u"(поставщик не определён)"))
    lst["kind"], lst["why"], lst["sup"] = [x[0] for x in routes], [x[1] for x in routes], [x[2] for x in routes]
    rank, dem = rank_by_demand(lst, list(alive["store"]), rev, sales)
    grp = group_sales(lst, sales)
    info["demand"] = dem
    shops = rank[:top_n]
    if len(shops) < top_n and not auto_points:
        info["problems"].append(u"Нашлось только %d точек-получателей из %d" % (len(shops), top_n))
    info["notes"] = list(info.get("notes") or []) + [
        u"Точки выбраны по спросу на этот список: продажи за %d дн., не больше чем на %d дн. вперёд, меньше %d чеков - случайность; "
        u"спрос есть у %d точек" % (SPLIT_DAYS_BC, SPLIT_HORIZON_DAYS, SPLIT_DEMAND_MIN_CHECKS, sum(1 for s in rank if dem[s] > 0))]
    if auto_points:                                       # число точек по сумме прямых поставок: per_shop[0]..per_shop[1] грн на точку
        direct_sum = float(lst.loc[lst["kind"] == KIND_SHOPS, "sum"].sum())
        n_auto = min(auto_top_n(direct_sum, per_shop[0], per_shop[1], SPLIT_TOP_N), max(1, len(shops)))
        shops = shops[:n_auto]
        info["top_n"] = len(shops)
        info["notes"] = list(info.get("notes") or []) + [
            u"Точек: %d (прямые поставки %.0f грн / %.0f-%.0f грн на точку = в среднем %.0f грн)"
            % (len(shops), direct_sum, per_shop[0], per_shop[1], direct_sum / max(1, len(shops)))]
    if keg_bcs:
        _f = lambda x: (u"%.3f" % x).rstrip("0").rstrip(".")
        parts = [u"%s %s = %s (%s)" % (b, _f(q_), u" + ".join(_f(p_) for p_ in keg_pieces(q_, keg_size[bc_key(b)])), keg_src[bc_key(b)])
                 for b, nm_, q_ in zip(lst["bc"], lst["name"], lst["qty"]) if is_keg(nm_)]
        info["notes"] = list(info.get("notes") or []) + [
            u"Кеги (%d поз.) на %s не едут и по литрам не делятся: целые кеги и начатый - каждый целиком на одну выбранную точку, "
            u"где этот ШК продаётся больше всего (не продаётся - на точку с наибольшим спросом на список). Порций: %s"
            % (len(keg_bcs), SPLIT_RC_NAME, u"; ".join(parts))]
    n_rc = int((lst["kind"] == KIND_RC).sum())
    info["rc_skipped"] = n_rc if skip_rc else 0
    if skip_rc and n_rc:
        info["notes"] = list(info.get("notes") or []) + [u"На склад (РЦ) %d поз.: файл на %s не делаю (режим «без РЦ»), на точки не делю" % (n_rc, SPLIT_RC_NAME)]
    hows = {}

    def _alloc(pts):
        out_ = []      # (адрес, ШК, количество)
        for r in lst.itertuples(index=False):
            if r.kind == KIND_RC:
                if not skip_rc:
                    out_.append((SPLIT_RC_NAME, r.bc, float(r.qty)))
                continue
            is_w = (r.unit == u"кг") or abs(r.qty - round(r.qty)) > 1e-9
            step = rz.WEIGHT_STEP if is_w else 1.0
            if is_keg(r.name):  # кег целиком: в обоих режимах только на выбранные точки
                got = keg_alloc(r.bc, keg_pieces(r.qty, keg_size.get(bc_key(r.bc))), pts, sales, dem)[0]
                hows[r.bc] = u"кег"
            else:
                sh, hows[r.bc] = demand_shares(r, pts, rev, sales, grp, dem)
                got = allocate_qty(r.qty, sh, step, SPLIT_MIN)
            for s, q in got.items():
                out_.append((s, r.bc, q))
        return out_

    cost_ = {b: float(c) for b, c in zip(lst["bc"], lst["cost"])}
    alloc, dropped, n_want = _alloc(shops), [], len(shops)
    while len(shops) > 1:                           # в обоих режимах: точка с суммой меньше per_shop[0] выбывает, её товар делится между остальными
        sums = {}
        for a_, b_, q_ in alloc:
            sums[a_] = sums.get(a_, 0.0) + q_ * cost_.get(b_, 0.0)
        low = [s_ for s_ in shops if sums.get(s_, 0.0) < per_shop[0] - 1e-6]
        if not low:
            break
        w_ = min(low, key=lambda s_: (sums.get(s_, 0.0), dem.get(s_, 0.0)))
        shops.remove(w_)
        dropped.append(u"%s (%.0f грн)" % (w_, sums.get(w_, 0.0)))
        alloc = _alloc(shops)
    if dropped:
        info["top_n"] = len(shops)
        info["notes"] = list(info.get("notes") or []) + [
            u"Меньше %.0f грн на точку не везём: убраны %s; их товар разделён между остальными. Точек было %d, едет %d"
            % (per_shop[0], u", ".join(dropped), n_want, len(shops))]
    fin = {}
    for a_, b_, q_ in alloc:
        if a_ != SPLIT_RC_NAME:
            fin[a_] = fin.get(a_, 0.0) + q_ * cost_.get(b_, 0.0)
    if len(fin) == 1 and sum(fin.values()) < per_shop[0] - 1e-6:
        info["notes"] = list(info.get("notes") or []) + [
            u"ВНИМАНИЕ: товара на точки всего %.0f грн - меньше %.0f, всё на одну точку (%s)" % (sum(fin.values()), per_shop[0], u", ".join(fin))]
    if hows:
        hs = {}
        for b_, q_, c_ in zip(lst["bc"], lst["qty"], lst["cost"]):
            if b_ in hows:
                hs[hows[b_]] = hs.get(hows[b_], 0.0) + float(q_) * float(c_)
        info["notes"] = list(info.get("notes") or []) + [
            u"Как делились товары на точки, грн: по продажам ШК %.0f, по группе %.0f, по спросу точки на список %.0f, "
            u"по обороту (спроса нет - проверьте) %.0f, кеги %.0f"
            % tuple(hs.get(x_, 0.0) for x_ in (u"ШК", u"группа", u"спрос на список", u"оборот", u"кег"))]
    info["hows"] = hows
    al = pd.DataFrame(alloc, columns=["addr", "bc", "qty"])
    if al.empty:
        info["notes"] = list(info.get("notes") or []) + [u"Позиций с прямой доставкой нет: распределять на точки нечего"]
        info["ok"], info["routes"] = True, lst
        return info
    info_by = lst.set_index("bc")
    al["name"] = al["bc"].map(info_by["name"])
    al["unit"] = al["bc"].map(info_by["unit"])
    al["cost"] = al["bc"].map(info_by["cost"])
    al["sum"] = (al["qty"] * al["cost"]).round(2)
    al["step"] = [1.0 if (u != u"кг" and abs(q - round(q)) < 1e-9) else rz.WEIGHT_STEP for u, q in zip(al["unit"], al["qty"])]

    chk = al.groupby("bc")["qty"].sum().round(3)
    bad = [b for b, _kd in zip(lst["bc"], lst["kind"]) if not (skip_rc and _kd == KIND_RC) and abs(chk.get(b, 0.0) - float(info_by.loc[b, "qty"])) > 1e-6]
    if bad:
        info["problems"].append(u"Не сошлось количество по ШК: %s" % u", ".join(bad[:10]))
        return info

    try:                                                  # список из ВЫВОЗ\\<папка дня>\\<магазин>: результат рядом с ним (в т.ч. ..._зачистка)
        inside = os.path.normcase(os.path.abspath(path)).startswith(os.path.normcase(os.path.abspath(dirs.out)) + os.sep)
    except Exception:
        inside = False
    out_dir = (os.path.join(os.path.dirname(os.path.abspath(path)), SPLIT_OUT_DIR) if inside
               else os.path.join(dirs.out, day, rz.safe_name(shop), SPLIT_OUT_DIR))
    os.makedirs(out_dir, exist_ok=True)
    _old = [f for f in os.listdir(out_dir) if f.lower().endswith(".txt")]
    if _old:
        _arc = os.path.join(out_dir, u"_прошлые_запуски", datetime.now().strftime("%Y-%m-%d_%H%M%S"))
        os.makedirs(_arc, exist_ok=True)
        for _f in _old:
            shutil.move(os.path.join(out_dir, _f), os.path.join(_arc, _f))
    extra = [a_ for a_ in dict.fromkeys(al["addr"]) if a_ != SPLIT_RC_NAME and a_ not in shops]       # точки вне топа: сюда уехали кеги
    extra.sort(key=lambda a_: -rev.get(a_, 0.0))
    order = [SPLIT_RC_NAME] + shops + extra
    files, exp = [], {}
    for addr in order:
        g = al[al["addr"] == addr]
        if g.empty:
            continue
        fp = os.path.join(out_dir, rz.safe_name(addr) + ".txt")
        rz.write_shop_txt(fp, g[["bc", "qty", "step"]])
        exp[fp] = {b: round(float(q), 3) for b, q in zip(g["bc"], g["qty"])}
        files.append(fp)
    for fp, e in exp.items():                       # обратное чтение txt
        with open(fp, "rb") as f:
            raw = f.read().decode(rz.TXT_ENCODING)
        got = {}
        for ln in [x for x in raw.split(rz.TXT_NEWLINE) if x]:
            if not rz.TXT_LINE_RE.match(ln):
                info["problems"].append(u"%s: неверная строка «%s»" % (os.path.basename(fp), ln))
                continue
            b, q = ln.split(";")
            got[b] = round(float(q.replace(",", ".")), 3)
        if got != e or not raw.endswith(rz.TXT_NEWLINE):
            info["problems"].append(u"%s: txt не совпал с расчётом" % os.path.basename(fp))

    xp = os.path.join(out_dir, u"Распределение.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = u"Сводка"
    ws.append([u"Адрес", u"Тип", u"Позиций", u"Единиц", u"Сумма (себестоимость)", u"Оборот точки за %d дн., грн" % SPLIT_DAYS_SHOP,
               u"Спрос на этот список за %d дн., грн" % SPLIT_HORIZON_DAYS])
    for addr in order:
        g = al[al["addr"] == addr]
        if g.empty:
            continue
        ws.append([addr, u"РЦ" if addr == SPLIT_RC_NAME else u"Точка", len(g), round(float(g["qty"].sum()), 3),
                   round(float(g["sum"].sum()), 2), round(float(rev.get(addr, 0.0))),
                   u"" if addr == SPLIT_RC_NAME else round(float(dem.get(addr, 0.0)))])
    ws.append([u"ИТОГО (позиций в списке)", u"", int(lst.shape[0]), round(float(lst["qty"].sum()), 3), round(float(lst["sum"].sum()), 2), u""])
    if skip_rc and info.get("rc_skipped"):
        _rc = lst[lst["kind"] == KIND_RC]
        ws.append([u"  в т.ч. склад (РЦ): файл не делался (режим «без РЦ»)", u"РЦ", int(_rc.shape[0]), round(float(_rc["qty"].sum()), 3), round(float(_rc["sum"].sum()), 2), u""])
    _fit_widths(ws)
    ws = wb.create_sheet(u"Маршрут")
    ws.append([u"Штрих-код", u"Название", u"Кол-во", u"Ед.", u"Сумма", u"Куда", u"Поставщик", u"Основание", u"Как делилось"])
    for r in lst.sort_values(["kind", "sup", "name"]).itertuples(index=False):
        ws.append([r.bc, r.name, r.qty, r.unit, r.sum, (u"Склад (РЦ): файл не делался (режим «без РЦ»)" if skip_rc else u"Склад (РЦ)") if r.kind == KIND_RC else u"Распределение по точкам", r.sup, r.why,
                   hows.get(r.bc, u"")])
    _fit_widths(ws)
    ws = wb.create_sheet(u"Матрица ШК x адрес")
    cols = [a for a in order if a in set(al["addr"])]
    mx = al.pivot_table(index=["bc"], columns="addr", values="qty", aggfunc="sum").reindex(columns=cols)
    ws.append([u"Штрих-код", u"Название"] + cols + [u"Итого"])
    for b in lst["bc"]:
        vals = []
        for c in cols:
            v = mx.loc[b, c] if b in mx.index else None
            vals.append(None if v is None or pd.isna(v) else round(float(v), 3))
        ws.append([b, info_by.loc[b, "name"]] + vals + [round(float(info_by.loc[b, "qty"]), 3)])
    _fit_widths(ws)
    for addr in order:
        g = al[al["addr"] == addr]
        if g.empty:
            continue
        ws = wb.create_sheet(re.sub(r"[\[\]:*?/\\]", u"-", addr)[:31])
        ws.append([u"Штрих-код", u"Название", u"Количество", u"Ед.", u"Себестоимость", u"Сумма"])
        for r in g.itertuples(index=False):
            ws.append([r.bc, r.name, r.qty, r.unit, r.cost, r.sum])
        ws.append([u"", u"ИТОГО", round(float(g["qty"].sum()), 3), u"", u"", round(float(g["sum"].sum()), 2)])
        _fit_widths(ws)
    try:
        wb.save(xp)
        files.append(xp)
    except PermissionError:                               # файл открыт в Excel: сводку пишем под новым именем, а не теряем
        xp2 = os.path.join(out_dir, u"Распределение_%s.xlsx" % datetime.now().strftime("%H%M%S"))
        try:
            wb.save(xp2)
            files.append(xp2)
            info["notes"] = list(info.get("notes") or []) + [u"Распределение.xlsx открыт в Excel: актуальная сводка записана в %s" % os.path.basename(xp2)]
        except PermissionError:
            info["problems"].append(u"Распределение.xlsx занят (закройте в Excel): txt записаны, xlsx нет")
    info["ok"] = not info["problems"]
    info["out_dir"], info["files"] = out_dir, files
    for addr in order:
        g = al[al["addr"] == addr]
        if not g.empty:
            info["summary"].append((addr, len(g), round(float(g["qty"].sum()), 3), round(float(g["sum"].sum()), 2)))
    info["routes"] = lst
    info["alloc"] = al
    return info


def _split_text(info):
    lines = []
    for p in info.get("problems", []):
        lines.append(u"[ОШИБКА] " + p)
    for n in info.get("notes", []):
        lines.append(u"[ВНИМ] " + n)
    if info.get("summary"):
        lines.append(u"Папка: %s" % info["out_dir"])
        for a, n, q, s in info["summary"]:
            lines.append(u"  %-28s %4d поз. %9.3f ед. %10.2f грн" % (a, n, q, s))
    return u"\n".join(lines)


def _split_shops(text):
    return [s.strip() for s in re.split(r"[;|\n]", text) if s.strip()]


def _parse_args(argv):
    a = {"selftest": False, "auto": False, "shops": None, "no_ua": False, "offline": False, "rm": False,
         "excl": [], "incl": [], "split": None, "day": None, "top": None, "with_rc": False, "no_rc": False, "update": None,
         "state": [], "shop_ua": [], "final": None, "final_file": None, "sweep": False, "auto_points": False, "map_list": False, "map_import": None, "manual": False}
    it = iter(argv)
    for x in it:
        t = x.lower().lstrip("-/")
        if t == "selftest":
            a["selftest"] = True
        elif t in ("auto", "a", "no-gui"):
            a["auto"] = True
        elif t == "no-ua":
            a["no_ua"] = True
        elif t == "offline":
            a["offline"] = True
        elif t == "remove-not-on-ua":
            a["rm"] = True
        elif t == "split":                             # исправленный лист вывоза -> txt по адресам (РЦ + топ-точки)
            a["split"] = next(it, "")
        elif t == "update":
            a["update"] = next(it, "")
        elif t in ("shop-ua", "shop_ua"):
            a["shop_ua"] += [x for x in next(it, "").split(";") if x.strip()]
        elif t == "state":
            a["state"] += [x for x in next(it, "").split(";") if x.strip()]
        elif t in ("with-rc", "with_rc"):         # оставлено для совместимости: файл на РЦ делается по умолчанию
            a["with_rc"] = True
        elif t in ("no-rc", "no_rc"):             # не делать файл на Полевая-Склад
            a["no_rc"] = True
        elif t == "sweep":                        # зачистка остатка: вывезти всё, кроме заморозки, скоропортов и расходников
            a["sweep"] = True
        elif t in ("manual", "manual-select"):    # ручной выбор: сигареты и кеги в списке, причина в последней колонке
            a["manual"] = True
        elif t == "map-list":                     # список кандидатов для таблицы соответствия ШК Family <-> ЮА (Excel)
            a["map_list"] = True
        elif t == "map-import":                   # загрузить решения из этого списка в Справочник\vyvoz_ua_map.csv
            a["map_import"] = next(it, "")
        elif t in ("auto-points", "auto_points"):   # число точек по сумме: 5-7 тыс. грн на точку
            a["auto_points"] = True
        elif t in ("final-inventory", "final"):   # финальный файл инвентаризации ЮА из выгрузки Family после вывоза
            a["final"] = next(it, "")
        elif t in ("final-file", "final_file"):   # конкретная выгрузка Family для финала (иначе свежая из МАГАЗИНЫ)
            a["final_file"] = next(it, "")
        elif t in ("top", "points", "tochki"):     # на сколько точек делить (1..SPLIT_TOP_N)
            v = next(it, "")
            try:
                a["top"] = int(v)
            except ValueError:
                print(u"--top: ожидается число 1..%d, получено «%s»" % (SPLIT_TOP_N, v))
        elif t == "day":                               # ГГГГ-ММ-ДД папки результата (по умолчанию из файла)
            a["day"] = next(it, "")
        elif t == "shops":
            a["shops"] = _split_shops(next(it, ""))
        elif t == "exclude-suppliers":                 # на этот прогон добавить к исключённым (в файл не пишется)
            a["excl"] = _split_shops(next(it, ""))
        elif t == "include-suppliers":                 # на этот прогон вернуть в вывоз
            a["incl"] = _split_shops(next(it, ""))
        elif t == "shops-file":
            p = next(it, "")
            try:
                with open(p, encoding="utf-8-sig") as f:
                    a["shops"] = _split_shops(f.read())
            except Exception:
                try:
                    with open(p, encoding="cp1251") as f:
                        a["shops"] = _split_shops(f.read())
                except Exception:
                    print(u"Не прочитан файл со списком магазинов: %s" % p)
                    a["shops"] = []
    return a


UPD_DIR_PREFIX   = u"Обновление_"
UPD_REMOVED_FILE = u"Удалено_по_приходу.xlsx"


APP_ICO = u"vyvoz_fm_v2.ico"
FM_RED, FM_RED_D, FM_GRAPH = "#e31e24", "#b8161b", "#2b2c30"


def _logo_color(u, v):
    """Логотип «Фемелі Маркет»: сумка с буквой М (красный/графит) на сером скруглённом квадрате. u,v в 0..1."""
    GR, RD, BG = (0x2b, 0x2c, 0x30), (0xe3, 0x1e, 0x24), (0xd6, 0xd6, 0xd8)
    x = u - .5
    if .15 <= v < .36:
        d = ((u - .5) ** 2 + (v - .36) ** 2) ** .5
        if .115 <= d <= .165:
            return GR
    if .34 <= v <= .90:
        hw = .25 + .05 * (v - .34) / .56
        if abs(x) <= hw:
            W = hw - .055
            if .40 <= v <= .84 and abs(x) <= W:
                xi, ax = W - .085, abs(x)
                top = v < .62 and ax < xi * .9 * (.62 - v) / .22
                bot = ax < xi and v > .76 - (ax / xi) * .26
                if not (top or bot):
                    return RD
            return GR
    m, r = .03, .2
    if m <= u <= 1 - m and m <= v <= 1 - m:
        cx, cy = min(max(u, m + r), 1 - m - r), min(max(v, m + r), 1 - m - r)
        if (u - cx) ** 2 + (v - cy) ** 2 <= r * r:
            return BG
    return None


def _logo_raster(n, S=4):
    out = []
    for y in range(n):
        row = []
        for x in range(n):
            cs = []
            for i in range(S):
                for j in range(S):
                    c = _logo_color((x + (i + .5) / S) / n, (y + (j + .5) / S) / n)
                    if c:
                        cs.append(c)
            if not cs:
                row.append((0, 0, 0, 0))
                continue
            k = len(cs)
            row.append((sum(c[0] for c in cs) // k, sum(c[1] for c in cs) // k, sum(c[2] for c in cs) // k, 255 * k // (S * S)))
        out.append(row)
    return out


def _make_ico(path):
    """Пишет .ico с логотипом (16..256 px)."""
    import struct
    imgs = []
    for n in (16, 20, 24, 32, 40, 48, 64, 256):
        R = _logo_raster(n, 4 if n <= 64 else 2)
        xor = bytearray()
        for row in reversed(R):
            for r, g, b, a in row:
                xor += bytes((b, g, r, a))
        andm = b"\0" * (((n + 31) // 32) * 4 * n)
        hdr = struct.pack("<IiiHHIIiiII", 40, n, n * 2, 1, 32, 0, len(xor) + len(andm), 0, 0, 0, 0)
        imgs.append((n, hdr + bytes(xor) + andm))
    out, off, body = struct.pack("<HHH", 0, 1, len(imgs)), 6 + 16 * len(imgs), b""
    for n, data in imgs:
        out += struct.pack("<BBBBHHII", n % 256, n % 256, 0, 0, 1, 32, len(data), off)
        off += len(data)
        body += data
    tmp = path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(out + body)
    os.replace(tmp, path)


def _logo_photo(master, n, bg="#ffffff"):
    """Логотип как картинка Tk, края смешаны с цветом фона bg."""
    import tkinter as tk
    br, bgc, bb = int(bg[1:3], 16), int(bg[3:5], 16), int(bg[5:7], 16)
    img = tk.PhotoImage(master=master, width=n, height=n)
    rows = []
    for row in _logo_raster(n, 4):
        cells = []
        for r, g, b, a in row:
            f = a / 255.0
            cells.append("#%02x%02x%02x" % (int(r * f + br * (1 - f)), int(g * f + bgc * (1 - f)), int(b * f + bb * (1 - f))))
        rows.append("{" + " ".join(cells) + "}")
    img.put(" ".join(rows))
    return img


def _set_app_icon(app):
    """Иконка окна и панели задач. Ошибки - в app._fm_icon_err (выводятся в «Журнал»)."""
    import tempfile
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(u"FamilyMarket.VyvozVneMatricy")
    except Exception:
        pass
    errs = []
    for d in (getattr(DIRS, "cache", None), os.path.join(os.environ.get("LOCALAPPDATA") or tempfile.gettempdir(), u"VyvozFM"),
              tempfile.gettempdir()):
        if not d:
            continue
        try:
            os.makedirs(d, exist_ok=True)
            ico = os.path.join(d, APP_ICO)
            if not os.path.isfile(ico) or os.path.getsize(ico) < 1000:
                _make_ico(ico)
            app.iconbitmap(ico)
            app.iconbitmap(default=ico)
            app._fm_ico = ico
            return ico
        except Exception as e:
            errs.append(u"%s: %s: %s" % (d, type(e).__name__, e))
    try:
        img = _logo_photo(app, 32, "#d6d6d8")
        app.iconphoto(True, img)
        app._fm_icon = img
    except Exception as e:
        errs.append(u"iconphoto: %s: %s" % (type(e).__name__, e))
    app._fm_icon_err = u"\n".join(errs)
    return None


def _fm_theme(app, k=1.0):
    """Фирменная тема: светлый фон, графит, красный акцент (ttk clam - красится полностью)."""
    from tkinter import ttk

    def px(v):
        return int(v * k)
    BG, FLD, BRD, TX, MUT, HOV, HD = "#f6f6f7", "#ffffff", "#dcdce1", "#1f2023", "#6b6b70", "#ececee", "#f0f0f2"
    s = ttk.Style(app)
    s.theme_use("clam")
    app.configure(background=BG)
    app.option_add("*Toplevel.background", BG)
    app.option_add("*TCombobox*Listbox.selectBackground", FM_RED)
    app.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
    app.option_add("*TCombobox*Listbox.font", ("Segoe UI", 10))
    s.configure(".", background=BG, foreground=TX, font=("Segoe UI", 10), bordercolor=BRD, lightcolor=BG, darkcolor=BG,
                troughcolor="#e4e4e7", focuscolor=FM_RED, selectbackground=FM_RED, selectforeground="#ffffff",
                fieldbackground=FLD, insertcolor=TX)
    s.map(".", foreground=[("disabled", "#a1a1aa")])
    s.configure("TFrame", background=BG)
    s.configure("Card.TFrame", background=BG, borderwidth=1, relief="solid", bordercolor=BRD)
    s.configure("TLabel", background=BG, foreground=TX)
    s.configure("TButton", background=FLD, foreground=TX, bordercolor=BRD, lightcolor=FLD, darkcolor=FLD,
                padding=(px(12), px(5)), relief="flat", focusthickness=0)
    s.map("TButton", background=[("disabled", "#f1f1f2"), ("pressed", "#e2e2e5"), ("active", HOV)],
          lightcolor=[("pressed", "#e2e2e5"), ("active", HOV)], darkcolor=[("pressed", "#e2e2e5"), ("active", HOV)],
          bordercolor=[("active", "#c4c4c9")])
    s.configure("Accent.TButton", background=FM_RED, foreground="#ffffff", bordercolor=FM_RED, lightcolor=FM_RED,
                darkcolor=FM_RED, font=("Segoe UI", 10, "bold"))
    acc = [("disabled", "#f0b3b5"), ("pressed", "#9c1217"), ("active", FM_RED_D)]
    s.map("Accent.TButton", background=acc, bordercolor=acc, lightcolor=acc, darkcolor=acc, foreground=[("disabled", "#ffffff")])
    s.configure("TRadiobutton", background=BG, foreground=TX, indicatorbackground=FLD, indicatorforeground=FM_RED,
                upperbordercolor="#8e8e94", lowerbordercolor="#8e8e94", padding=px(2))
    s.map("TRadiobutton", background=[("active", BG)], indicatorbackground=[("selected", FLD), ("pressed", HOV)],
          indicatorforeground=[("selected", FM_RED)], upperbordercolor=[("selected", FM_RED)], lowerbordercolor=[("selected", FM_RED)])
    s.configure("TCheckbutton", background=BG, foreground=TX, indicatorbackground=FLD, indicatorforeground="#ffffff",
                upperbordercolor=BRD, lowerbordercolor=BRD, padding=px(2))
    s.map("TCheckbutton", background=[("active", BG)], indicatorbackground=[("selected", FM_RED), ("pressed", HOV)],
          upperbordercolor=[("selected", FM_RED)], lowerbordercolor=[("selected", FM_RED)])
    for w in ("TEntry", "TSpinbox", "TCombobox"):
        s.configure(w, fieldbackground=FLD, background=FLD, bordercolor=BRD, lightcolor=FLD, darkcolor=FLD,
                    padding=px(5), arrowcolor=TX, foreground=TX)
        s.map(w, bordercolor=[("focus", FM_RED)], lightcolor=[("focus", FM_RED)],
              fieldbackground=[("readonly", FLD), ("disabled", "#f1f1f2")], foreground=[("readonly", TX)])
    s.map("TCombobox", selectbackground=[("readonly", FLD)], selectforeground=[("readonly", TX)])
    s.map("TSpinbox", selectbackground=[("readonly", FLD), ("focus", FLD)], selectforeground=[("readonly", TX), ("focus", TX)])
    s.configure("Treeview", background=FLD, fieldbackground=FLD, foreground=TX, bordercolor=BRD, lightcolor=FLD,
                darkcolor=FLD, rowheight=px(26))
    s.map("Treeview", background=[("selected", FM_RED)], foreground=[("selected", "#ffffff")])
    s.configure("Treeview.Heading", background=HD, foreground=TX, bordercolor=BRD, lightcolor=HD, darkcolor=HD,
                relief="flat", font=("Segoe UI", 9, "bold"), padding=(px(6), px(4)))
    s.map("Treeview.Heading", background=[("active", "#e2e2e5")])
    s.configure("TScrollbar", background="#d0d0d4", troughcolor=BG, bordercolor=BG, lightcolor="#d0d0d4",
                darkcolor="#d0d0d4", arrowcolor=MUT, gripcount=0, arrowsize=px(12))
    s.map("TScrollbar", background=[("active", "#b4b4ba")])
    s.configure("TProgressbar", background=FM_RED, troughcolor="#e4e4e7", bordercolor="#e4e4e7",
                lightcolor=FM_RED, darkcolor=FM_RED)
    try:
        import tkinter as tk

        def _rgb(h):
            return int(h[1:3], 16), int(h[3:5], 16), int(h[5:7], 16)

        def _rr(u, v, m, r):
            if u < m or u > 1 - m or v < m or v > 1 - m:
                return False
            cx, cy = min(max(u, m + r), 1 - m - r), min(max(v, m + r), 1 - m - r)
            return (u - cx) ** 2 + (v - cy) ** 2 <= r * r

        def _seg(u, v, a, b):
            (x1, y1), (x2, y2) = a, b
            dx, dy = x2 - x1, y2 - y1
            t = max(0.0, min(1.0, ((u - x1) * dx + (v - y1) * dy) / (dx * dx + dy * dy)))
            return ((u - x1 - t * dx) ** 2 + (v - y1 - t * dy) ** 2) ** .5

        def _chk(n, on, dis=False):
            fill = (("#f0b3b5" if on else "#f1f1f2") if dis else (FM_RED if on else FLD))
            brd = fill if on else ("#c8c8cc" if dis else "#8e8e94")
            cb, cf, cr, cw = _rgb(BG), _rgb(fill), _rgb(brd), (255, 255, 255)
            t = 1.4 / n
            img = tk.PhotoImage(master=app, width=n + px(7), height=n)
            rows, S = [], 4
            for y in range(n):
                cells = []
                for x in range(n):
                    acc = [0, 0, 0]
                    for a in range(S):
                        for b in range(S):
                            u, v = (x + (a + .5) / S) / n, (y + (b + .5) / S) / n
                            if not _rr(u, v, .04, .22):
                                c = cb
                            elif not _rr(u, v, .04 + t, max(.01, .22 - t)):
                                c = cr
                            elif on and min(_seg(u, v, (.25, .52), (.43, .70)), _seg(u, v, (.43, .70), (.76, .33))) <= .075:
                                c = cw
                            else:
                                c = cf
                            acc[0] += c[0]
                            acc[1] += c[1]
                            acc[2] += c[2]
                    cells.append("#%02x%02x%02x" % tuple(int(q / (S * S)) for q in acc))
                rows.append("{" + " ".join(cells) + "}")
            img.put(" ".join(rows), to=(0, 0))
            return img
        n = px(16)
        app._fm_chk = (_chk(n, False), _chk(n, True), _chk(n, False, True), _chk(n, True, True))
        s.element_create("FM.Checkbutton.indicator", "image", app._fm_chk[0],
                         ("disabled", "selected", app._fm_chk[3]), ("disabled", app._fm_chk[2]), ("selected", app._fm_chk[1]))
        s.layout("TCheckbutton", [("Checkbutton.padding", {"sticky": "nswe", "children": [
            ("FM.Checkbutton.indicator", {"side": "left", "sticky": ""}),
            ("Checkbutton.focus", {"side": "left", "sticky": "w", "children": [("Checkbutton.label", {"sticky": "nswe"})]})]})])
    except Exception:
        pass
    return s


def _read_ua_keys(paths):
    """Выгрузки ЮА -> ({ключ ШК: (остаток, файл)}, [названия складов в файлах]).
    Берутся ВСЕ строки: остаток плюс, ноль или минус (приход наперёд) - ШК в выгрузке = товар заведён на ЮА."""
    res, names = {}, []
    for fp in paths:
        df, _inf = read_state_file(fp)
        if df is None or df.empty:
            continue
        if "shop" in df.columns:
            names += [str(x) for x in df["shop"].dropna().unique() if str(x).strip()]
        q = pd.to_numeric(df["qty"], errors="coerce").fillna(0.0)
        for kk, v in zip(df["key"], q):
            kk = str(kk or u"").strip()
            if kk:
                a = res.get(kk)
                res[kk] = (round((a[0] if a else 0.0) + float(v), 3), os.path.basename(fp))
    return res, sorted(set(names))


def _nm_cmp(x):
    return re.sub(u"[^0-9a-zа-яіїєґё]", u"", str(x).lower()).replace(u"юа", u"")


def _txt_lead_bc(t):
    m = re.match(r"\s*(\d{4,})\s*[;\t, ]?\s*(.*)$", t)
    return (m.group(1), m.group(2).strip()) if m else (None, u"")


def _clean_txt(src, dst, removed):
    """Копия txt для ТСД без строк с ШК из removed (остальные байт в байт). -> (было, осталось, [(ШК, кол-во)]).
    Если ничего не осталось - файл не пишется."""
    with open(src, "rb") as f:
        data = f.read()
    nlb = b"\r\n" if b"\r\n" in data else b"\n"
    bom = b"\xef\xbb\xbf" if data.startswith(b"\xef\xbb\xbf") else b""
    keep, drop, was = [], [], 0
    for ln in data[len(bom):].split(nlb):
        t = ln.decode("latin-1").strip()
        if t:
            was += 1
        b, q = _txt_lead_bc(t)
        if b and bc_key(b) in removed:
            drop.append((b, q))
        else:
            keep.append(ln)
    left = sum(1 for ln in keep if ln.strip())
    if left:
        with open(dst, "wb") as f:
            f.write(bom + nlb.join(keep))
    return was, left, drop


def _txt_keys(fp):
    with open(fp, "rb") as f:
        data = f.read().decode("latin-1")
    out = set()
    for t in data.splitlines():
        b, _q = _txt_lead_bc(t)
        if b:
            out.add(bc_key(b))
    return out


def _fix_vyvoz_sheet(ws):
    """Лист «Вывезти на склад» после правки строк: заново нумерует, пересчитывает «ИТОГО» и строку «Позиций / Единиц /
    Себестоимость» в шапке (иначе остаются цифры исходного списка). -> (позиций, единиц, сумма) или None, если макет не тот."""
    hr = None
    for r in range(1, min(ws.max_row, 12) + 1):
        if (str(ws.cell(row=r, column=1).value or u"").strip() == u"№"
                and str(ws.cell(row=r, column=2).value or u"").strip() == u"Штрих-код"):
            hr = r
            break
    if hr is None:
        return None
    n, q_sum, s_sum, tot_row = 0, 0.0, 0.0, None
    for r in range(hr + 1, ws.max_row + 1):
        if str(ws.cell(row=r, column=3).value or u"").strip().upper().startswith(u"ИТОГО"):
            tot_row = r
            break
        if ws.cell(row=r, column=2).value in (None, u""):
            continue
        n += 1
        ws.cell(row=r, column=1).value = n
        q, sm = ws.cell(row=r, column=4).value, ws.cell(row=r, column=7).value
        q_sum += float(q) if isinstance(q, (int, float)) else 0.0
        s_sum += float(sm) if isinstance(sm, (int, float)) else 0.0
    q_sum, s_sum = round(q_sum, 3), round(s_sum, 2)
    if tot_row:
        ws.cell(row=tot_row, column=4).value = q_sum
        ws.cell(row=tot_row, column=7).value = s_sum
    a2 = ws.cell(row=2, column=1)
    if isinstance(a2.value, str) and u"Позиций:" in a2.value:
        t = re.sub(u"Позиций:\\s*\\d+", u"Позиций: %d" % n, a2.value)
        t = re.sub(u"Единиц:\\s*[-\\d.]+", u"Единиц: %s" % q_sum, t)
        a2.value = re.sub(u"Себестоимость:\\s*[-\\d.]+", u"Себестоимость: %s" % s_sum, t)
    return n, q_sum, s_sum


def _clean_xlsx(src, dst, removed):
    """Копия xlsx без строк, где есть ШК из removed; на листе вывоза итоги и нумерация пересчитываются. -> {лист: удалено строк}"""
    wb = load_workbook(src)
    res = {}
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if v is None:
                    continue
                if isinstance(v, float) and v.is_integer():
                    v = int(v)
                t = str(v).strip()
                if t.isdigit() and len(t) >= 6 and bc_key(t) in removed:
                    rows.append(c.row)
                    break
        for ri in sorted(set(rows), reverse=True):
            ws.delete_rows(ri)
        if ws.title == SHEET_VYVOZ:
            _fix_vyvoz_sheet(ws)
        if rows:
            res[ws.title] = len(set(rows))
    wb.save(dst)
    return res


def run_update(dirs, path, wh_paths, shop_paths=None, ref=None):
    """Новый день: из исправленного листа вывоза и готовых txt для ТСД убрать то, что стало ЮА-шным.
    1) ШК есть в «Состоянии склада» Склад ЮА (любой остаток, в т.ч. 0 и минус) -> удаляем при любом маршруте;
    2) позиция прямой доставки (есть в txt точки) и ШК есть в «Состоянии склада» магазина в базе ЮА -> удаляем.
    Исходники не меняются: всё пишется в <магазин>\\Обновление_<дата>. Считается всегда от исходника."""
    reset_report()
    info = {"ok": False, "problems": [], "notes": [], "files": [], "removed": [], "txt_stat": []}
    wh_paths = [x for x in (wh_paths or []) if x]
    shop_paths = [x for x in (shop_paths or []) if x]
    if not wh_paths:
        info["problems"].append(u"Не выбран файл «Состояние склада» Склад ЮА")
        return info
    try:
        lst, shop, day0, _n = read_corrected_list(path)
    except Exception as e:
        info["problems"].append(u"Лист вывоза не прочитан: %s" % e)
        return info
    if lst.empty:
        info["problems"].append(u"В листе вывоза нет позиций")
        return info
    try:
        wh, wh_names = _read_ua_keys(wh_paths)
        shp, shp_names = _read_ua_keys(shop_paths) if shop_paths else ({}, [])
    except Exception as e:
        info["problems"].append(u"Выгрузка ЮА не прочитана: %s" % e)
        return info
    if not wh:
        info["problems"].append(u"В выгрузке Склад ЮА нет ни одного ШК")
        return info
    if wh_names and not any(_is_polevaya(x) for x in wh_names):
        info["notes"].append(u"ВНИМАНИЕ: в поле «Склад ЮА» файл не похож на Полевая-Склад (в файле: %s)" % u", ".join(wh_names[:3]))
    if shop_paths and shp_names:
        if any(_is_polevaya(x) for x in shp_names):
            info["notes"].append(u"ВНИМАНИЕ: в поле «Магазин в базе ЮА» выбран склад, а не магазин (%s)" % u", ".join(shp_names[:3]))
        elif shop and not any(_nm_cmp(shop) in _nm_cmp(x) or _nm_cmp(x) in _nm_cmp(shop) for x in shp_names):
            info["notes"].append(u"ВНИМАНИЕ: магазин в файле ЮА (%s) не совпадает с магазином листа (%s)" % (u", ".join(shp_names[:3]), shop))
    if ref is None:
        try:
            fr = load_reference(dirs)
            ref = make_ref(fr["matrix"], fr["recode"], fr["coffee"], None, fr["hist"], ua_map=load_ua_map(dirs)[0]) if fr else None
        except Exception as e:
            ref = None
            info["notes"].append(u"Справочник перекодировки недоступен (%s): сверка только по ШК" % e)
    pairs = (getattr(ref, "pairs", None) or {}) if ref is not None else {}

    def finder(d):
        rmap = {}
        if ref is not None:
            for x in d:
                rmap.setdefault(_resolve(ref, x), x)

        def hit(kk):
            if kk in d:
                return kk
            if ref is None:
                return None
            for c in (_resolve(ref, kk), pairs.get(kk), getattr(ref, "ua_map", {}).get(kk)):
                if c and c in d:
                    return c
            return rmap.get(_resolve(ref, kk))
        return hit
    hit_wh, hit_shp = finder(wh), finder(shp)

    cands = []
    for sh in [x for x in (shop, rz.safe_name(shop) if shop else None) if x]:
        if day0:
            cands.append(os.path.join(dirs.out, str(day0), sh, SPLIT_OUT_DIR))
    cands.append(os.path.join(os.path.dirname(path), SPLIT_OUT_DIR))
    split_dir = next((c for c in cands if os.path.isdir(c)), None)
    rc_txt = (rz.safe_name(SPLIT_RC_NAME) + ".txt").lower()
    direct = set()
    if split_dir:
        for f in os.listdir(split_dir):
            if f.lower().endswith(".txt") and f.lower() != rc_txt:
                direct |= _txt_keys(os.path.join(split_dir, f))
    if shp and not split_dir:
        info["notes"].append(u"Нет папки %s: маршрут позиций неизвестен, по магазину ЮА ничего не удалено" % SPLIT_OUT_DIR)

    rem, rem_ua = {}, {}
    for r in lst.itertuples(index=False):
        kk = bc_key(r.bc)
        if kk in rem:
            continue
        h = hit_wh(kk)
        if h:
            rem[kk], rem_ua[kk] = (r, u"есть на Склад ЮА", wh[h][0], wh[h][1]), h
            continue
        if shp and kk in direct:
            h = hit_shp(kk)
            if h:
                rem[kk], rem_ua[kk] = (r, u"прямая доставка: есть на магазине ЮА", shp[h][0], shp[h][1]), h
    removed = set(rem)
    info["by_why"] = {}
    for _r, why, _q, _f in rem.values():
        info["by_why"][why] = info["by_why"].get(why, 0) + 1
    if not removed:
        info["ok"] = True
        info["notes"].append(u"Ни одна позиция листа не стала ЮА-шной: обновление не нужно, файлы не создавались")
        return info
    if len(removed) > 0.8 * len(lst):
        info["notes"].append(u"ВНИМАНИЕ: удаляется %d из %d позиций - проверьте, те ли выбраны файлы" % (len(removed), len(lst)))

    base = os.path.dirname(split_dir) if split_dir else os.path.dirname(path)
    out = os.path.join(base, UPD_DIR_PREFIX + date.today().strftime("%Y-%m-%d"))
    if os.path.isdir(out):
        shutil.rmtree(out, ignore_errors=True)
    os.makedirs(out, exist_ok=True)
    info["out_dir"] = out
    try:
        with open(os.path.join(out, u"ЧИТАЙ_МЕНЯ.txt"), "w", encoding="utf-8") as f_:
            f_.write(u"Актуальные файлы после «Нового дня» от %s: из них убрано то, что уже стало ЮА-шным.\n"
                     u"В ТСД берите файлы ИЗ ЭТОЙ ПАПКИ. Исходные файлы (лист вывоза, ТСД_по_адресам) не менялись и устарели.\n"
                     u"Снятые позиции остаются на полке и попадут в инвентаризацию ЮА из финальной выгрузки Family.\n"
                     % date.today().strftime("%d.%m.%Y"))
    except OSError:
        pass
    info["notes"].append(u"Снятые позиции остаются на полке: в инвентаризацию ЮА они попадут из финальной выгрузки (раздел «Инвентаризация ЮА»)")

    dst = os.path.join(out, os.path.basename(path))
    try:
        by_sheet = _clean_xlsx(path, dst, removed)
        info["files"].append(dst)
        info["notes"].append(u"Лист вывоза: удалено строк " + u", ".join(u"%s - %d" % kv for kv in by_sheet.items()))
        lst2 = read_corrected_list(dst)[0]
        k2 = set(bc_key(b) for b in lst2["bc"])
        n_rm = sum(1 for b in lst["bc"] if bc_key(b) in removed)
        if k2 & removed:
            info["problems"].append(u"В очищенном листе остались удаляемые ШК: %s" % u", ".join(sorted(k2 & removed)[:10]))
        if len(lst2) != len(lst) - n_rm:
            info["problems"].append(u"Не сошлось: было %d, удалено %d, осталось %d" % (len(lst), n_rm, len(lst2)))
    except PermissionError:
        info["problems"].append(u"Файл занят (закройте в Excel): %s" % dst)
    except Exception as e:
        info["problems"].append(u"Лист вывоза не очищен: %s" % e)

    where, txts = {}, []
    if split_dir:
        txts += [(os.path.join(split_dir, f), os.path.join(out, SPLIT_OUT_DIR, f), os.path.splitext(f)[0])
                 for f in sorted(os.listdir(split_dir)) if f.lower().endswith(".txt")]
    else:
        info["notes"].append(u"Папка %s не найдена: очищен только лист вывоза" % SPLIT_OUT_DIR)
    own = os.path.splitext(path)[0] + ".txt"
    if os.path.isfile(own):
        txts.append((own, os.path.join(out, os.path.basename(own)), u"основной txt магазина"))
    if txts:
        os.makedirs(os.path.join(out, SPLIT_OUT_DIR), exist_ok=True)
    for srcf, dstf, addr in txts:
        try:
            was, left, drop = _clean_txt(srcf, dstf, removed)
        except Exception as e:
            info["problems"].append(u"%s: не обработан (%s)" % (os.path.basename(srcf), e))
            continue
        if was != left + len(drop):
            info["problems"].append(u"%s: не сошлось строк (было %d, осталось %d, удалено %d)" % (addr, was, left, len(drop)))
        info["txt_stat"].append((addr, was, left, len(drop)))
        if left:
            info["files"].append(dstf)
        else:
            info["notes"].append(u"%s: удалены все позиции, файл не создан" % addr)
        for b, q in drop:
            where.setdefault(bc_key(b), []).append(u"%s %s" % (addr, q))

    # склад (РЦ): файла в ТСД_по_адресам нет (распределяли с галочкой «склад уже распределён») -> собираем из листа (_rc_from_list)
    rc_file = rz.safe_name(SPLIT_RC_NAME) + ".txt"
    if split_dir and not os.path.isfile(os.path.join(split_dir, rc_file)):
        rc_all = [r for r in lst.itertuples(index=False) if bc_key(r.bc) not in direct]
        rc_left = [r for r in rc_all if bc_key(r.bc) not in removed]
        for r in rc_all:
            if bc_key(r.bc) in removed:
                where.setdefault(bc_key(r.bc), []).append(u"%s %s" % (SPLIT_RC_NAME, r.qty))
        if rc_all:
            info["txt_stat"].append((SPLIT_RC_NAME + u" (собран из листа)", len(rc_all), len(rc_left), len(rc_all) - len(rc_left)))
        if rc_left:
            df_rc = pd.DataFrame({"bc": [r.bc for r in rc_left], "qty": [float(r.qty) for r in rc_left]})
            df_rc["step"] = [1.0 if (str(getattr(r, "unit", u"")) != u"кг" and abs(float(r.qty) - round(float(r.qty))) < 1e-9)
                             else rz.WEIGHT_STEP for r in rc_left]
            os.makedirs(os.path.join(out, SPLIT_OUT_DIR), exist_ok=True)
            fp_rc = os.path.join(out, SPLIT_OUT_DIR, rc_file)
            try:
                rz.write_shop_txt(fp_rc, df_rc)
                info["files"].append(fp_rc)
                kw = _txt_keys(fp_rc)
                if kw & removed:
                    info["problems"].append(u"%s: остались удаляемые ШК" % rc_file)
                if len(kw) != len(set(bc_key(r.bc) for r in rc_left)):
                    info["problems"].append(u"%s: не сошлось позиций (в файле %d, должно %d)"
                                            % (rc_file, len(kw), len(set(bc_key(r.bc) for r in rc_left))))
                info["notes"].append(u"%s собран из листа: %d поз. (удалено пришедших %d)"
                                     % (rc_file, len(rc_left), len(rc_all) - len(rc_left)))
            except Exception as e:
                info["problems"].append(u"%s не записан: %s" % (rc_file, e))
        elif rc_all:
            info["notes"].append(u"%s: все позиции склада стали ЮА-шными, файл не создан" % SPLIT_RC_NAME)
    tq = ts = 0.0
    for kk, (r, why, uq, uf) in rem.items():
        q = float(r.qty)
        try:
            cost = float(getattr(r, "cost", 0) or 0)
            cost = 0.0 if cost != cost else cost
        except (TypeError, ValueError):
            cost = 0.0
        tq, ts = tq + q, ts + q * cost
        info["removed"].append({"bc": r.bc, "name": getattr(r, "name", u""), "qty": q, "unit": getattr(r, "unit", u""),
                                "cost": cost, "sum": round(q * cost, 2), "why": why,
                                "where": u"; ".join(where.get(kk, [])) or u"склад (РЦ) / основной список",
                                "ua_qty": uq, "ua_file": uf,
                                "ua_bc": (_ua_bc(ref, rem_ua[kk]) if ref is not None else rem_ua[kk])})
    xp = os.path.join(out, UPD_REMOVED_FILE)
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = u"Удалено"
        ws.append([u"ШК", u"Название", u"Кол-во в списке", u"Ед.", u"Себестоимость", u"Сумма", u"Причина",
                   u"Куда должно было ехать", u"Остаток в выгрузке ЮА", u"Файл ЮА", u"ШК на ЮА"])
        for c in ws[1]:
            c.font = Font(bold=True)
        for d in info["removed"]:
            ws.append([d["bc"], d["name"], d["qty"], d["unit"], d["cost"], d["sum"], d["why"], d["where"], d["ua_qty"], d["ua_file"], d["ua_bc"]])
        ws.append([u"", u"ИТОГО", round(tq, 3), u"", u"", round(ts, 2)])
        _fit_widths(ws)
        w2 = wb.create_sheet(u"Источник")
        w2.append([u"Исходный лист вывоза (не менялся)", path])
        for fp in wh_paths:
            w2.append([u"Склад ЮА", fp])
        for fp in shop_paths:
            w2.append([u"Магазин в базе ЮА", fp])
        w2.append([u"Сформировано", datetime.now().strftime("%d.%m.%Y %H:%M")])
        _fit_widths(w2)
        wb.save(xp)
        info["files"].append(xp)
    except PermissionError:
        info["problems"].append(u"%s занят (закройте в Excel)" % UPD_REMOVED_FILE)
    info["removed_qty"], info["removed_sum"] = round(tq, 3), round(ts, 2)
    info["ok"] = not info["problems"]
    return info


def _update_text(info):
    L = []
    if info.get("out_dir"):
        L.append(u"Папка: %s" % info["out_dir"])
    rm = info.get("removed", [])
    L.append(u"Удалено позиций: %d, ед. %s, сумма %s грн"
             % (len(rm), info.get("removed_qty", 0), u"{:,.2f}".format(info.get("removed_sum", 0)).replace(",", u" ")))
    for why, n in (info.get("by_why") or {}).items():
        L.append(u"  %s: %d" % (why, n))
    for d in rm[:60]:
        L.append(u"  %-14s %-40s %8s %-3s [%s] -> %s" % (d["bc"], str(d["name"])[:40], d["qty"], d["unit"], d["why"], d["where"]))
    if len(rm) > 60:
        L.append(u"  ... ещё %d (см. %s)" % (len(rm) - 60, UPD_REMOVED_FILE))
    if info.get("txt_stat"):
        L.append(u"Файлы для ТСД:")
        for a, was, left, dr in info["txt_stat"]:
            L.append(u"  %-30s было %4d, удалено %4d, осталось %4d" % (a, was, dr, left))
    L += [u"Примечание: " + x for x in info.get("notes", [])]
    L += [u"ОШИБКА: " + x for x in info.get("problems", [])]
    return u"\n".join(L)


# ======================= ТАБЛИЦА СООТВЕТСТВИЯ ШК Family <-> ЮА =======================
# На ЮА один и тот же товар заведён под другим ШК и названием (внутренние коды 2938080..., опечатки ШК, префиксы ХД/МВУ/
# Продбаза/Джамп, «Х.Ц.З.» вместо «ХЦЗ»). Здесь: 1) список кандидатов пар для ручной проверки (Excel), 2) загрузка ваших решений
# в Справочник\vyvoz_ua_map.csv, 3) применение таблицы в расчёте (позиция Family с парой на ЮА считается заведённой на ЮА,
# в инвентаризацию идёт ШК ЮА). В BigQuery (barcode_recode_map) отсюда ничего не пишется: это делает проект матрицы.

UA_MAP_FILE = u"vyvoz_ua_map.csv"
MAP_DIR = u"Соответствие"
MAP_COLS = [u"ШК Family", u"ШК ЮА", u"Название Family", u"Название ЮА", u"Статус", u"Дата", u"Файл"]
_MAP_NOISE = {u"хд", u"мву", u"джамп", u"продбаза", u"корона", u"упаковка", u"шт", u"1шт", u"з/б", u"ж/б",
              u"тм", u"ua", u"юа", u"крим", u"вино", u"сигарети", u"сигарет", u"твен"}     # «Вілла Крим» = «ТМ Villa UA» (ребрендинг)
MAP_AUTO_SIM = 0.75              # ТАК заранее - только A с таким сходством названий (A по одному ШК ±2 знака не отмечаем)
_MAP_OPT_NUM = ("pk", "pc")      # «(6)» - штук в ящике, «35%» - жирность: есть только у одной стороны -> не различие
_MAP_PACK_RE = re.compile(r"(\(\*?\d+\)|упаковка|\(\d+\s*шт\))")


def load_ua_map(dirs):
    """Таблица соответствия: ({ШК Family без нулей: ШК ЮА без нулей} со статусом ТАК, {(ШК Family, ШК ЮА)} со статусом НЕТ)."""
    path = os.path.join(dirs.cache, UA_MAP_FILE)
    yes, no = {}, set()
    try:
        if os.path.isfile(path):
            df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
            for f, u, st in zip(df.iloc[:, 0], df.iloc[:, 1], df.iloc[:, 4]):
                kf, ku = bc_key(f), bc_key(u)
                if not (kf and ku):
                    continue
                if str(st).strip().upper() == u"ТАК":
                    yes[kf] = ku
                elif str(st).strip().upper() == u"НЕТ":
                    no.add((kf, ku))
    except Exception as e:
        R.check("WARN", u"Таблица соответствия ШК Family-ЮА", u"не прочитана (%s): работаю без неё" % e)
    return yes, no


def save_ua_map(dirs, rows):
    """Дописывает решения (список dict по MAP_COLS) в Справочник\\vyvoz_ua_map.csv; одинаковая пара не дублируется, новое решение заменяет старое."""
    path = os.path.join(dirs.cache, UA_MAP_FILE)
    os.makedirs(dirs.cache, exist_ok=True)
    old = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig") if os.path.isfile(path) else pd.DataFrame(columns=MAP_COLS)
    old.columns = MAP_COLS[:len(old.columns)] if len(old.columns) <= len(MAP_COLS) else old.columns
    new = pd.DataFrame(rows, columns=MAP_COLS)
    allr = pd.concat([old, new], ignore_index=True)
    allr["_k"] = [bc_key(a) + u"|" + bc_key(b) for a, b in zip(allr[MAP_COLS[0]], allr[MAP_COLS[1]])]
    allr = allr.drop_duplicates("_k", keep="last").drop(columns="_k")
    allr.to_csv(path, index=False, encoding="utf-8-sig")
    return len(allr)


def _map_norm(s):
    t = str(s).lower().replace(u"’", u"'").replace(u"`", u"'").replace(u"*", u"'").replace(u"ё", u"е")
    t = re.sub(r"(\d),(\d)", r"\1.\2", t)
    t = re.sub(u"н/(сол|сл)(?![а-яіїєґ])", u"н/с", t)
    t = re.sub(r"\(\s*\*?\s*(\d+)\s*(?:шт)?\s*\)", r" pk\1 ", t)                   # «(6)», «(*50)», «(24 шт)» -> pk6
    t = re.sub(r"(\d+(?:\.\d+)?)\s*%", r" pc\1 ", t)                                # «35%» -> pc35
    t = re.sub(r"(\d+(?:\.\d+)?)\s*(л|кг|гр|г|мл|мг|шт|l|ml)\b", lambda m: m.group(1) + {u"гр": u"г", u"мг": u""}.get(m.group(2), m.group(2)), t)
    for _ in range(2):
        t = re.sub(r"(?<=[а-яіїєґa-z])\.(?=[а-яіїєґa-z])", u"", t)          # Х.Ц.З. -> хцз
    return re.sub(r"[^0-9a-zа-яіїєґ.' /]", u" ", t)


def _map_tokens(s):
    out = []
    cig = bool(CIG_ANY_RE.search(nname(s)))
    for t in _map_norm(s).split():
        t = t.strip(u".'")
        if cig and t in (u"20шт", u"20"):                    # «20шт» у сигарет Family - пачка, у ЮА его нет
            continue
        if t and t not in _MAP_NOISE and not re.fullmatch(r"20\d\d", t):
            out.append(t)
    return out


def _lat_skel(t):
    """Слово -> латинский «скелет» (транслит, без гласных и h, повторы схлопнуты): Вілла ~ Villa, Собраніе ~ Sobranie, Грін ~ Green."""
    t = _WORD_MAP.get(t, t)
    t = u"".join(_TR.get(ch, ch) for ch in t)
    for a, b in (("w", "v"), ("ph", "f"), ("ck", "k"), ("c", "k"), ("y", "i"), ("x", "ks"), ("q", "k")):
        t = t.replace(a, b)
    k = re.sub(r"[aeiouj]+", "", t).replace("h", "")
    k = re.sub(r"[^a-z0-9]+", "", k)
    k = re.sub(r"(.)\1+", r"\1", k)
    return (k[:-1] if len(k) > 2 and k.endswith("s") else k) or t        # Series ~ Серія, Slims ~ Слімс


def _word_sim(aw, bw):
    if not (aw and bw):
        return 0.0
    inter = len(aw & bw)
    return max(inter / float(len(aw | bw)), 0.9 * inter / float(min(len(aw), len(bw))) if min(len(aw), len(bw)) >= 2 else 0.0)


def _lev2(a, b):
    """Расстояние Левенштейна для ШК; больше 2 - сразу 9."""
    if abs(len(a) - len(b)) > 2:
        return 9
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1] if prev[-1] <= 2 else 9


def _map_score(ut, ft, ustr, fstr, ed):
    """-> (уверенность A/B/C/-, сходство 0..1, совпали ли числа/объёмы, пояснение)"""
    aw, bw = set(t for t in ut if not re.search(r"\d", t)), set(t for t in ft if not re.search(r"\d", t))
    sim = _word_sim(aw, bw)
    sim = max(sim, 0.97 * _word_sim(set(_lat_skel(t) for t in aw), set(_lat_skel(t) for t in bw)))   # кириллица ~ латиница
    if sim < 0.9:
        qa, qb = ustr.replace(u" ", u""), fstr.replace(u" ", u"")
        sm = difflib.SequenceMatcher(None, qa, qb)
        if sm.quick_ratio() >= 0.75:
            sim = max(sim, 0.9 * sm.ratio())
    na, nb = set(t for t in ut if re.search(r"\d", t)), set(t for t in ft if re.search(r"\d", t))
    for pre in _MAP_OPT_NUM:                                # число этого вида есть только у одной стороны - не различие
        pa, pb = set(x for x in na if x.startswith(pre)), set(x for x in nb if x.startswith(pre))
        if not (pa and pb):
            na, nb = na - pa, nb - pb
    nums_eq = na == nb
    why = []
    if ed <= 2:
        why.append(u"ШК отличается на %d зн." % ed)
    if sim >= 0.9:
        why.append(u"название почти совпало")
    elif sim >= 0.7:
        why.append(u"название похоже")
    if not nums_eq:
        why.append(u"ОБЪЁМ/ВЕС/ЧИСЛА РАЗНЫЕ")
    if (ed <= 2 and sim >= 0.5 and nums_eq) or (sim >= 0.9 and nums_eq):
        c = u"A"
    elif sim >= 0.75 and nums_eq:
        c = u"B"
    elif sim >= 0.55:
        c = u"C"
    else:
        c = u"-"
    return c, sim, nums_eq, u"; ".join(why)


def fetch_family_keys(keys):
    """Какие из ШК есть в базе Family (приходы или продажи на точках не ЮА, за всё время) -> set ключей.
    BigQuery недоступен -> пустое множество (позиции останутся в списке, как раньше)."""
    keys = sorted(set(k for k in keys if k))
    if not keys or BQ_OFFLINE:
        return set()
    lst = u",".join(u"'%s'" % k for k in keys)
    try:
        ok, df, _ = _bq_query(u"SELECT DISTINCT LTRIM(barcode,'0') k FROM `%s` WHERE store NOT LIKE '%% ЮА' AND LTRIM(barcode,'0') IN (%s) "
                              u"UNION DISTINCT SELECT DISTINCT LTRIM(barcode,'0') k FROM `%s` WHERE store NOT LIKE '%% ЮА' AND LTRIM(barcode,'0') IN (%s)"
                              % (INCOMING_TABLE, lst, TURNOVER_TABLE, lst))
    except Exception:
        return set()
    return set(str(k) for k in df["k"]) if ok else set()


def run_map_list(dirs, ref=None, per_item=3):
    """Список кандидатов для таблицы соответствия: позиции ЮА, у которых нет пары в Family (ни по ШК, ни по перекодировке, ни по названию),
    и до 3 похожих позиций Family (матрица + остатки магазинов из ВХОД_ВЫВОЗ\\МАГАЗИНЫ). Решения проставляются в Excel, загружаются --map-import.
    -> dict(ok, path, problems, stats)"""
    reset_report()
    dirs.ensure()
    info = {"ok": False, "problems": [], "path": u"", "stats": {}}
    try:
        yes, no = load_ua_map(dirs)
        ua_df, _ui = load_ua(dirs)
        if ref is None:
            fr = load_reference(dirs)
            if fr is None:
                info["problems"].append(u"Справочник (матрица) не загружен")
                return info
            ref = make_ref(fr["matrix"], fr["recode"], fr["coffee"], ua_df, fr["hist"], pairs=load_shk_pairs(dirs), ua_map=yes)
        stores = scan_stores(dirs)
    except Exception as e:
        info["problems"].append(u"Данные не прочитаны: %s: %s" % (type(e).__name__, e))
        return info
    # остатки на ЮА (склад / магазин) для справки
    qty_wh, qty_shop = {}, {}
    for folder, dst in ((dirs.ua_wh, qty_wh), (dirs.ua_stores, qty_shop)):
        files = list_xlsx(folder)
        if files:
            try:
                dst.update({k: q for k, (q, _f) in _read_ua_keys([_pick_latest(files)])[0].items()})
            except Exception:
                pass
    # Family: матрица + остатки магазинов
    fam = {}
    for k, (nm, _st, sup) in ref.matrix.items():
        fam[k] = {"name": nm, "src": [u"матрица"], "qty": 0.0, "sup": sup}
    for shop, (fp, df) in stores.items():
        lab = u"%s %s" % (shop, guess_day(fp)[0].strftime("%d.%m"))
        for r in df[df["qty"] > 0].itertuples():
            e = fam.setdefault(r.key, {"name": r.name, "src": [], "qty": 0.0, "sup": u""})
            e["src"].append(lab)
            e["qty"] += float(r.qty)
    ua_keys = set(ua_df["key"])
    ua_names = ref.ua_names
    claimed, unmatched = set(), {}
    for k, e in fam.items():
        c = _resolve(ref, k)
        rc = ref.ua_recoded.get(k) or ref.ua_recoded.get(c)
        if k in ua_keys or c in ua_keys:
            claimed.add(k if k in ua_keys else c)
        elif rc in ua_keys:                        # ШК ЮА уже склеен с этим ШК в barcode_recode_map
            claimed.add(rc)
        elif nname(e["name"]) in ua_names:
            claimed.update(x[0] for x in ua_names[nname(e["name"])])
        elif yes.get(k) in ua_keys:
            claimed.add(yes[k])
        elif k in ref.pairs and ref.pairs[k] in ua_keys:
            claimed.add(ref.pairs[k])
        else:
            unmatched[k] = e
    pool = ua_df[~ua_df["key"].isin(claimed)].drop_duplicates("key")
    same = set(k for k in pool["key"] if k in ref.hist) | fetch_family_keys([k for k in pool["key"] if k not in ref.hist])
    n_same = 0                                     # тот же ШК есть в базе Family, похожих в остатках/матрице нет - пара не нужна
    ftok = {k: (_map_tokens(e["name"]), _map_norm(e["name"])) for k, e in unmatched.items()}
    fsk = {k: set(_lat_skel(t) for t in ft if not re.search(r"\d", t)) for k, (ft, _s) in ftok.items()}
    rows = []
    for r in pool.itertuples():
        ut, ustr = _map_tokens(r.name), _map_norm(r.name)
        usk = set(_lat_skel(t) for t in ut if not re.search(r"\d", t))
        nn = nname(r.name)
        cls = u"Сигареты" if CIG_ANY_RE.search(nn) else (u"Упаковка / штучный" if _MAP_PACK_RE.search(nn) else u"Обычный")
        uk = r.key.lstrip(u"0")
        cands = []
        for k, (ft, fstr) in ftok.items():
            if (k, r.key) in no or (k.lstrip(u"0"), uk) in no:
                continue
            ed = _lev2(uk, k.lstrip(u"0"))
            if not (set(ut) & set(ft)) and ed > 2 and not (set(fstr.split()) & set(ustr.split())) and not (usk & fsk[k]):
                continue
            c, sim, neq, why = _map_score(ut, ft, ustr, fstr, ed)
            if c != u"-":
                cands.append((u"ABC".index(c), -sim, k, c, sim, why))
        cands.sort()
        best = cands[:per_item]
        if r.key in same and not best:
            n_same += 1
            continue
        rows.append({"cls": cls, "conf": best[0][3] if best else u"-", "sim": best[0][4] if best else 0.0, "key": r.key, "name": r.name, "src": r.src,
                     "q_wh": qty_wh.get(r.key), "q_shop": qty_shop.get(r.key),
                     "cands": [(b[2], fam[b[2]]["name"], u"; ".join(fam[b[2]]["src"][:3]) + (u" (ост. %g)" % round(fam[b[2]]["qty"], 3) if fam[b[2]]["qty"] else u""),
                                u"%s %.2f%s" % (b[3], b[4], (u": " + b[5]) if b[5] else u"")) for b in best]})
    order = {u"A": 0, u"B": 1, u"C": 2, u"-": 3}
    rows.sort(key=lambda x: ({u"Обычный": 0, u"Упаковка / штучный": 1, u"Сигареты": 2}[x["cls"]], order[x["conf"]], x["name"]))
    out_dir = os.path.join(dirs.base, MAP_DIR)
    os.makedirs(out_dir, exist_ok=True)
    path = os.path.join(out_dir, u"Соответствие_ШК_%s.xlsx" % datetime.now().strftime("%Y-%m-%d_%H%M"))
    wb = Workbook()
    ws = wb.active
    ws.title = u"Инструкция"
    for ln in (u"СПИСОК ДЛЯ ТАБЛИЦЫ СООТВЕТСТВИЯ ШК Family <-> ЮА",
               u"",
               u"Что это. Позиции ЮА, у которых нет пары в Family ни по ШК, ни по перекодировке, ни по названию, и до трёх похожих позиций Family.",
               u"Что делать. На листе «Кандидаты» в первой колонке напишите: ТАК (или 1) - это первый кандидат; 2 или 3 - второй / третий;",
               u"штрих-код Family - если верный товар другой; НЕТ - не пара. Пустые строки пропускаются. Сохраните файл.",
               u"Потом: python vyvoz_vne_matricy.py --map-import \"путь к файлу\" (или Сервис -> «Загрузить решения»). Пары попадут в Справочник\\vyvoz_ua_map.csv.",
               u"Как читать уверенность: A - ШК отличается на 1-2 знака или название почти совпало, числа (объём, вес) те же; B - название похоже, числа те же;",
               u"C - только похоже, ПРОВЕРЬТЕ вкус, объём, граммовку. «ОБЪЁМ/ВЕС/ЧИСЛА РАЗНЫЕ» - скорее всего другой товар.",
               u"ТАК в первой колонке уже стоит у уверенности A с похожим названием (кроме «Упаковка / штучный»): проверьте глазами, неверное сотрите или замените на НЕТ / 2 / 3.",
               u"Пара из этой таблицы работает так: позиция Family считается заведённой на ЮА, в файл инвентаризации ЮА идёт ШК ЮА.",
               u"В BigQuery (barcode_recode_map) отсюда ничего не пишется."):
        ws.append([ln])
    ws.column_dimensions["A"].width = 150
    ws = wb.create_sheet(u"Кандидаты")
    head = [u"Решение", u"Уверенность", u"Класс", u"ШК ЮА", u"Название ЮА", u"Остаток ЮА: склад", u"Остаток ЮА: магазин", u"Где в реестре ЮА"]
    for n in (1, 2, 3):
        head += [u"%d: ШК Family" % n, u"%d: Название Family" % n, u"%d: Где есть в Family" % n, u"%d: Сходство / почему" % n]
    ws.append(head)
    for x in rows:
        auto = x["conf"] == u"A" and x["sim"] >= MAP_AUTO_SIM and x["cls"] != u"Упаковка / штучный"   # упаковка <-> штучный - разные товары
        line = [u"ТАК" if auto else None, x["conf"], x["cls"], str(x["key"]), x["name"], x["q_wh"], x["q_shop"], x["src"]]
        for n in range(3):
            line += list(x["cands"][n]) if n < len(x["cands"]) else [None, None, None, None]
        ws.append(line)
    for col in (4, 9, 13, 17):
        for row_ in ws.iter_rows(min_row=2, min_col=col, max_col=col):
            for c0 in row_:
                c0.number_format = "@"
    for c0 in ws[1]:
        c0.font = Font(bold=True)
    for j, w in enumerate([10, 11, 16, 15, 44, 10, 10, 26] + [15, 44, 34, 30] * 3, 1):
        ws.column_dimensions[get_column_letter(j)].width = w
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = ws.dimensions
    import collections
    stats = collections.Counter((x["cls"], x["conf"]) for x in rows)
    ws = wb.create_sheet(u"Сводка")
    ws.append([u"Показатель", u"Значение"])
    ws.append([u"Позиций Family (матрица + остатки магазинов)", len(fam)])
    ws.append([u"из них уже с парой на ЮА (ШК, перекодировка, название, пары, таблица)", len(fam) - len(unmatched)])
    ws.append([u"Позиций ЮА без пары (в списке)", len(rows)])
    ws.append([u"Не в списке: тот же ШК есть в базе Family, похожих в матрице и остатках нет - пара не нужна", n_same])
    ws.append([u"Подтверждённых пар в таблице / отклонённых", u"%d / %d" % (len(yes), len(no))])
    for (cls, c), n in sorted(stats.items()):
        ws.append([u"%s, уверенность %s" % (cls, c), n])
    ws.column_dimensions["A"].width = 70
    try:
        wb.save(path)
    except PermissionError:
        info["problems"].append(u"Файл занят: %s" % path)
        return info
    info.update({"ok": True, "path": path, "stats": {"fam": len(fam), "matched": len(fam) - len(unmatched), "pool": len(rows),
                                                    "A": sum(1 for x in rows if x["conf"] == u"A"), "B": sum(1 for x in rows if x["conf"] == u"B"),
                                                    "C": sum(1 for x in rows if x["conf"] == u"C"), "yes": len(yes), "no": len(no)}})
    return info


def run_map_import(dirs, path):
    """Решения из списка кандидатов -> Справочник\\vyvoz_ua_map.csv. ТАК/1/2/3 - пара, штрих-код - своя пара, НЕТ - отклонить показанных кандидатов."""
    reset_report()
    info = {"ok": False, "problems": [], "added": 0, "rejected": 0, "skipped": 0}
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb[u"Кандидаты"] if u"Кандидаты" in wb.sheetnames else wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
    except Exception as e:
        info["problems"].append(u"Файл не прочитан: %s" % e)
        return info
    if not rows:
        info["problems"].append(u"Пустой файл")
        return info
    out, today = [], datetime.now().strftime("%Y-%m-%d")
    base = os.path.basename(path)
    for r in rows[1:]:
        dec = str(r[0]).strip() if r[0] is not None else u""
        if not dec:
            continue
        ukey, uname = rz.fmt_barcode(r[3]), str(r[4] or u"")
        cands = [(rz.fmt_barcode(r[8 + 4 * n]), str(r[9 + 4 * n] or u"")) for n in range(3) if len(r) > 9 + 4 * n and r[8 + 4 * n] not in (None, u"")]
        d = dec.upper()
        pick = None
        if d in (u"ТАК", u"ДА", u"1", u"+", u"YES") and len(cands) >= 1:
            pick = cands[0]
        elif d == u"2" and len(cands) >= 2:
            pick = cands[1]
        elif d == u"3" and len(cands) >= 3:
            pick = cands[2]
        elif re.fullmatch(r"\d{6,14}", dec):
            pick = (dec, u"(указано вручную)")
        if pick:
            out.append(dict(zip(MAP_COLS, [pick[0], ukey, pick[1], uname, u"ТАК", today, base])))
            info["added"] += 1
        elif d in (u"НЕТ", u"-", u"NO", u"Н"):
            for fk, fn in cands:
                out.append(dict(zip(MAP_COLS, [fk, ukey, fn, uname, u"НЕТ", today, base])))
            info["rejected"] += 1
        else:
            info["skipped"] += 1
    if out:
        try:
            save_ua_map(dirs, out)
        except Exception as e:
            info["problems"].append(u"Таблица не записана: %s" % e)
            return info
    info["ok"] = True
    return info


def _map_text(info):
    L = [u"[ОШИБКА] " + p for p in info.get("problems", [])]
    if info.get("path"):
        s = info.get("stats", {})
        L += [u"Файл: %s" % info["path"],
              u"Позиций Family %d, с парой на ЮА %d; позиций ЮА без пары %d (уверенность A: %d, B: %d, C: %d); в таблице пар %d, отклонено %d"
              % (s.get("fam", 0), s.get("matched", 0), s.get("pool", 0), s.get("A", 0), s.get("B", 0), s.get("C", 0), s.get("yes", 0), s.get("no", 0))]
    if "added" in info:
        L.append(u"Загружено: пар %d, отклонено строк %d, пропущено (пусто / не понято) %d" % (info["added"], info["rejected"], info["skipped"]))
    return u"\n".join(L)


FINAL_INV_PREFIX = u"Инвентаризация_ЮА_"
SHEET_FIN = u"Инвентаризация ЮА"
SHEET_FIN_OUT = u"Не идёт в инвентаризацию"
SHEET_FIN_NOCARD = u"Нет карточки на ЮА"
SHEET_FIN_CHECK = u"Сверка"


def final_inventory_rows(res):
    """Финал: вывоз уже проведён, поэтому на полке остаётся ВСЁ, что есть в выгрузке; в инвентаризацию ЮА идёт всё, кроме
    кулинарии и штучных товаров (упаковка уже на ЮА). Вердикты «вывезти» и «ждёт решения» здесь значения не имеют:
    что должно было уехать, того в финальной выгрузке уже нет, а что осталось - остаётся на полке."""
    return res[~res["verdict"].isin(("KUL", "PACK"))]


def run_final_inventory(dirs, shop, state_path=None, ref=None, use_ua=True, day=None):
    """Файл инвентаризации ЮА из ФИНАЛЬНОЙ выгрузки Family (снятой после проведения вывоза): весь остаток на полке,
    кроме кулинарии и штучных. ШК в файле - тот, что знает ЮА (перекодировка, то же название, пары), иначе ШК Family.
    Правки списка вывоза вручную значения не имеют: файл строится по факту остатка, а не по списку.
    -> dict(ok, out_dir, files, problems, notes, summary, stats, routes)"""
    reset_report()
    dirs.ensure()
    info = {"ok": False, "problems": [], "notes": [], "files": [], "out_dir": u"", "summary": [], "stats": {}}
    try:
        stores = scan_stores(dirs)
        if state_path:
            df0, _m = read_state_file(state_path)
            for sh0, g0 in df0.groupby("shop"):
                stores[sh0] = (state_path, g0.reset_index(drop=True))      # выбранный вручную файл важнее найденных
        sel, errs = resolve_shops([shop], sorted(stores))
    except Exception as e:
        info["problems"].append(u"Выгрузка не прочитана: %s" % e)
        return info
    if errs or not sel:
        info["problems"] += errs or [u"Магазин «%s» не найден среди выгрузок" % shop]
        return info
    shop_n = sel[0]
    path, df = stores[shop_n]
    age = check_fresh(u"магазин (финал)", path)
    info["notes"].append(u"Выгрузка Family: %s (снята %.0f ч назад). Она должна быть снята ПОСЛЕ проведения вывоза"
                         % (os.path.basename(path), age))
    try:
        ua_df, _uinfo = load_ua(dirs, use_ua)
        if ref is None:
            fr = load_reference(dirs)
            if fr is None:
                info["problems"].append(u"Справочник (матрица) не загружен: " + u"; ".join(d_ for s_, n_, d_ in R.checks if s_ == "ERROR"))
                return info
            ref = make_ref(fr["matrix"], fr["recode"], fr["coffee"], ua_df, fr["hist"], pairs=load_shk_pairs(dirs),
                           ua_map=load_ua_map(dirs)[0])
        agg, _pinfo = prepare_stock(df)
        agg = merge_recoded_duplicates(agg, ref)
        # remove_not_on_ua=True: те же правила кулинарии и штучных, что и в основном прогоне (они применяются к «вывозимым»)
        res = classify_stock(agg, ref, use_ua, True, set(load_excluded_suppliers(dirs)), load_keep_shk(dirs))
    except Exception as e:
        info["problems"].append(u"Расчёт не выполнен: %s: %s" % (type(e).__name__, e))
        return info
    if res.empty:
        info["problems"].append(u"В выгрузке нет позиций с положительным остатком")
        return info
    day = day or guess_day(path)[0]

    fin = final_inventory_rows(res)
    good = fin["inv_bc"].map(lambda b: bool(RE_BC_OK.match(str(b))))
    fin_ok, fin_bad = fin[good], fin[~good]
    outs = pd.concat([res[res["verdict"].isin(("KUL", "PACK"))], fin_bad])
    nocard = fin_ok[fin_ok["verdict"].isin(("VYVOZ", "CHECK2"))]
    ivg = inventory_file_rows(fin_ok)

    nm = rz.safe_name(shop_n)
    out_dir = os.path.join(dirs.out, day.strftime("%Y-%m-%d"), nm, FINAL_INV_PREFIX + datetime.now().strftime("%Y-%m-%d_%H%M%S"))
    os.makedirs(out_dir, exist_ok=True)
    info["out_dir"] = out_dir
    txt = os.path.join(out_dir, nm + INV_TXT_SUFFIX)
    if len(ivg):
        ir = pd.DataFrame({"bc": ivg["inv_bc"].values, "qty": ivg["qty"].values, "step": 1.0,
                           "is_weight": [(str(u).lower() in (u"кг", u"kg")) for u in ivg["unit"]]})
        rz.write_shop_txt(txt, ir)
        info["files"].append(txt)
        expected = {r.inv_bc: round(float(r.qty), 3) for r in ivg.itertuples()}
        got = _read_txt_dict(txt, info["problems"], u"txt инвентаризации")
        if got != expected:
            info["problems"].append(u"txt инвентаризации: расхождений %d" % len(set(expected.items()) ^ set(got.items())))
    tot_q, tot_s = round(float(res["qty"].sum()), 3), round(float(res["sum"].sum()), 2)
    part_q = round(float(fin_ok["qty"].sum()) + float(outs["qty"].sum()), 3)
    if abs(part_q - tot_q) > 0.01 or len(fin_ok) + len(outs) != len(res):
        info["problems"].append(u"Не сошёлся баланс: остаток %s ед. / %d поз., в частях %s ед. / %d поз."
                                % (tot_q, len(res), part_q, len(fin_ok) + len(outs)))

    def _why(r):
        if r.verdict == "VYVOZ":
            return u"карточки на ЮА нет (вне матрицы): создастся новая"
        if r.verdict == "CHECK2":
            return u"карточки на ЮА нет (в матрице, но на ЮА не заводили): создастся новая"
        return str(r.reason)

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_FIN
    ws.append([u"ШК в файле для ЮА", u"ШК Family", u"Название товара", u"Количество", u"Ед. изм.", u"Себестоимость", u"Сумма", u"Основание"])
    for r in fin_ok.sort_values("name", kind="stable").itertuples():
        ws.append([str(r.inv_bc), str(r.bc), r.name, round(float(r.qty), 3), r.unit,
                   (None if r.no_cost else round(float(r.cost), 2)), r.sum, _why(r)])
    ws.append([u"", u"", u"ИТОГО:", round(float(fin_ok["qty"].sum()), 3), u"", u"", round(float(fin_ok["sum"].sum()), 2), u""])
    for row_ in ws.iter_rows(min_row=2, max_col=2):
        for c0 in row_:
            c0.number_format = "@"
    _fit_widths(ws)
    ws = wb.create_sheet(SHEET_FIN_OUT)
    ws.append([u"ШК Family", u"Название товара", u"Количество", u"Ед. изм.", u"Сумма", u"Причина"])
    for r in outs.itertuples():
        ws.append([str(r.bc), r.name, round(float(r.qty), 3), r.unit, r.sum,
                   (u"некорректный ШК (не 6-14 цифр): в файл не попал" if r.verdict not in ("KUL", "PACK") else str(r.reason))])
    for row_ in ws.iter_rows(min_row=2, max_col=1):
        for c0 in row_:
            c0.number_format = "@"
    _fit_widths(ws)
    ws = wb.create_sheet(SHEET_FIN_NOCARD)
    ws.append([u"ШК Family", u"Название товара", u"Количество", u"Ед. изм.", u"Сумма", u"Поставщик", u"Основание"])
    for r in nocard.sort_values("name", kind="stable").itertuples():
        ws.append([str(r.bc), r.name, round(float(r.qty), 3), r.unit, r.sum, r.supplier or None, _why(r)])
    for row_ in ws.iter_rows(min_row=2, max_col=1):
        for c0 in row_:
            c0.number_format = "@"
    _fit_widths(ws)
    ws = wb.create_sheet(SHEET_FIN_CHECK)
    ws.append([u"Показатель", u"Позиций", u"Единиц", u"Сумма (себестоимость)"])
    ws.append([u"Остаток по финальной выгрузке", len(res), tot_q, tot_s])
    ws.append([u"В файле инвентаризации ЮА (строк Family)", len(fin_ok), round(float(fin_ok["qty"].sum()), 3), round(float(fin_ok["sum"].sum()), 2)])
    ws.append([u"  из них ШК в файле заменён на ШК ЮА", int((fin_ok["inv_bc"] != fin_ok["bc"]).sum()), u"", u""])
    ws.append([u"  из них без карточки на ЮА (создастся новая)", len(nocard), round(float(nocard["qty"].sum()), 3), round(float(nocard["sum"].sum()), 2)])
    ws.append([u"Не идёт в инвентаризацию (кулинария, штучные, плохой ШК)", len(outs), round(float(outs["qty"].sum()), 3), round(float(outs["sum"].sum()), 2)])
    ws.append([u"Строк в txt после слияния одинаковых ШК ЮА", len(ivg), round(float(ivg["qty"].sum()), 3) if len(ivg) else 0, u""])
    for n_ in info["notes"]:
        ws.append([n_])
    _fit_widths(ws)
    xp = os.path.join(out_dir, nm + u"_инвентаризация_ЮА.xlsx")
    try:
        wb.save(xp)
        info["files"].append(xp)
    except PermissionError:
        info["problems"].append(u"%s занят (закройте в Excel): txt записан, xlsx нет" % os.path.basename(xp))
    info["stats"] = {"stock_pos": len(res), "stock_units": tot_q, "stock_sum": tot_s,
                     "inv_pos": len(fin_ok), "inv_units": round(float(fin_ok["qty"].sum()), 3), "inv_sum": round(float(fin_ok["sum"].sum()), 2),
                     "file_lines": len(ivg), "remapped": int((fin_ok["inv_bc"] != fin_ok["bc"]).sum()),
                     "nocard_pos": len(nocard), "nocard_sum": round(float(nocard["sum"].sum()), 2),
                     "out_pos": len(outs), "out_sum": round(float(outs["sum"].sum()), 2)}
    s_ = info["stats"]
    info["summary"] = [(u"Остаток по выгрузке", s_["stock_pos"], s_["stock_units"], s_["stock_sum"]),
                       (u"В инвентаризацию ЮА", s_["inv_pos"], s_["inv_units"], s_["inv_sum"]),
                       (u"  из них без карточки на ЮА", s_["nocard_pos"], round(float(nocard["qty"].sum()), 3), s_["nocard_sum"]),
                       (u"Не идёт (кулинария, штучные)", s_["out_pos"], round(float(outs["qty"].sum()), 3), s_["out_sum"])]
    info["ok"] = not info["problems"]
    return info


def _final_text(info):
    L = []
    if info.get("out_dir"):
        L.append(u"Папка: %s" % info["out_dir"])
    for a, n, q, s in info.get("summary", []):
        L.append(u"  %-34s %5d поз. %10.3f ед. %12.2f грн" % (a, n, q, s))
    L += [u"Примечание: " + x for x in info.get("notes", [])]
    L += [u"ОШИБКА: " + x for x in info.get("problems", [])]
    return u"\n".join(L)


# ============ ОКНО: входные папки, выбор магазина, лист для распределения ============

HIDDEN_STORES_FILE = u"vyvoz_hidden_stores.csv"
INPUT_KINDS = (
    # ключ, папка (Dirs.<имя>), что положить, насколько нужно
    ("stores", "stores", u"Склад магазина (Family)", u"обязательно"),
    ("ua_wh", "ua_wh", u"Склад ЮА (все строки)", u"обязательно"),
    ("ua_receipts", "ua_receipts", u"Приходы ЮА, весь период", u"желательно"),
    ("ua_stores", "ua_stores", u"Этот магазин в базе ЮА", u"по желанию"),
)


def _mtime(p):
    try:
        return os.path.getmtime(p)
    except OSError:
        return 0.0


def load_hidden_stores(dirs):
    """Магазины, убранные из списка окна «Списки вывоза» (Справочник\\vyvoz_hidden_stores.csv, колонка «Магазин»).
    Файлы выгрузок не трогаются. -> множество имён."""
    path = os.path.join(dirs.cache, HIDDEN_STORES_FILE)
    try:
        if not os.path.isfile(path):
            return set()
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        return set(x.strip() for x in df.iloc[:, 0] if x.strip()) if len(df.columns) else set()
    except Exception as e:
        R.check("WARN", u"Список убранных магазинов", u"не прочитан (%s): показываю все магазины" % e)
        return set()


def save_hidden_stores(dirs, names):
    os.makedirs(dirs.cache, exist_ok=True)
    path = os.path.join(dirs.cache, HIDDEN_STORES_FILE)
    tmp = path + ".tmp"
    pd.DataFrame({u"Магазин": sorted(set(names))}).to_csv(tmp, index=False, encoding="utf-8-sig")
    os.replace(tmp, path)


def visible_stores(found, hidden):
    hk = set(rz.route_key(x) for x in hidden)
    return {s_: v for s_, v in found.items() if rz.route_key(s_) not in hk}


def pick_default_store(found):
    """Магазин по умолчанию - тот, чью выгрузку сняли последней (дата в имени файла, затем время файла): вывозят обычно
    только что выгруженный магазин. found: {магазин: (файл, ...)} -> имя или None."""
    best, best_key = None, None
    for shop in sorted(found):
        v = found[shop]
        p = v[0] if isinstance(v, (tuple, list)) else v
        try:
            key = (guess_day(p)[0], _mtime(p))
        except Exception:
            continue
        if best_key is None or key > best_key:
            best, best_key = shop, key
    return best


def _file_stamp(path):
    st_ = rz.read_export_stamp(path)
    return st_ if st_ is not None else datetime.fromtimestamp(_mtime(path))


def file_state(path, now=None):
    """Свежесть выгрузки -> (время выгрузки, часов назад, уровень OK/WARN, текст для окна)."""
    now = now or datetime.now()
    stamp = _file_stamp(path)
    age = max(0.0, (now - stamp).total_seconds() / 3600.0)
    ago = u"меньше часа" if age < 1 else (u"%.0f ч" % age if age < 48 else u"%.0f дн." % (age / 24.0))
    ok = age <= STATE_WARN_H
    return stamp, age, ("OK" if ok else "WARN"), (u"свежая, " if ok else u"СТАРАЯ, ") + ago


def input_status(dirs, now=None):
    """Что лежит во входных папках -> список строк для окна: ключ, папка, что положить, файл, когда выгружен, уровень, текст.
    Уровни: OK / WARN / ERR / INFO. Берётся тот же файл, что и в расчёте (склад ЮА и магазин - самый свежий; приходы и магазины ЮА - все)."""
    now = now or datetime.now()
    rows = []
    for key, attr, what, need in INPUT_KINDS:
        folder = getattr(dirs, attr)
        row = {"key": key, "folder": folder, "fold": os.path.basename(folder), "what": what, "need": need,
               "file": u"", "path": u"", "when": u"", "level": "INFO", "text": u""}
        try:
            files = list_xlsx(folder)
            if not files:
                row["level"] = {u"обязательно": "ERR", u"желательно": "WARN"}.get(need, "INFO")
                row["text"] = {u"обязательно": u"НЕТ ФАЙЛА", u"желательно": u"нет файла"}.get(need, u"нет (не обязательно)")
            else:
                pick = _pick_latest(files) if key in ("stores", "ua_wh") else max(files, key=_file_stamp)
                stamp, age, lv, txt = file_state(pick, now)
                row.update(path=pick, file=os.path.basename(pick) + (u"  ×%d" % len(files) if len(files) > 1 else u""),
                           when=stamp.strftime("%d.%m %H:%M"), level=lv, text=txt)
        except Exception as e:
            row.update(level="ERR", text=u"не прочитано: %s" % e)
        rows.append(row)
    rows.append({"key": "ref", "folder": u"", "fold": u"(интернет)", "what": u"Матрица, история поставок",
                 "need": u"само", "file": u"Google-лист / BigQuery", "path": u"", "when": u"",
                 "level": "INFO", "text": u"автоматически"})
    return rows


def corr_dir(dirs):
    return os.path.join(dirs.base, CORR_DIR)


def list_corrected(dirs):
    """Листы в Корректировка_ЮА для распределения: новые сверху; временные ~$ и файлы «проверка» не берём."""
    d_ = corr_dir(dirs)
    try:
        fs = [os.path.join(d_, f_) for f_ in os.listdir(d_)
              if f_.lower().endswith((".xlsx", ".xlsm")) and not f_.startswith("~$") and u"проверка" not in f_.lower()]
    except OSError:
        return []
    return sorted(fs, key=_mtime, reverse=True)


def is_in_corr(dirs, path):
    """Файл лежит прямо в Корректировка_ЮА (вложенные папки не считаются)."""
    try:
        return os.path.normcase(os.path.abspath(os.path.dirname(path))) == os.path.normcase(os.path.abspath(corr_dir(dirs)))
    except Exception:
        return False


def corrected_check(path):
    """Годится ли файл для распределения: лист «Вывезти на склад» в обычном виде (A1 «Вывоз вне матрицы: <магазин> -> ...»,
    колонки Штрих-код / Название товара / Количество). -> dict(ok, shop, day, n, sum, text)"""
    out = {"ok": False, "shop": u"", "day": u"", "n": 0, "sum": 0.0, "text": u""}
    try:
        wb = load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
    except Exception as e:
        out["text"] = u"Файл не открывается: %s" % e
        return out
    if SHEET_VYVOZ not in names:
        out["text"] = u"В файле нет листа «%s» (есть: %s)" % (SHEET_VYVOZ, u"; ".join(names))
        return out
    try:
        lst, shop, day, notes = read_corrected_list(path)
    except Exception as e:
        out["text"] = u"Лист не читается: %s" % e
        return out
    if not shop:
        out["text"] = u"В A1 нет «Вывоз вне матрицы: <магазин> -> ...»: непонятно, чей это вывоз"
        return out
    if lst.empty:
        out["text"] = u"В листе нет ни одной позиции"
        return out
    day_s = u"%s.%s.%s" % (day[8:10], day[5:7], day[0:4]) if day else u"дата не указана"
    out.update(ok=True, shop=shop, day=day_s, n=len(lst), sum=float(lst["sum"].sum()))
    out["text"] = u"%s · остатки на %s · %d поз. · %s грн" % (shop, day_s, len(lst), u"{:,.0f}".format(out["sum"]).replace(",", u" "))
    if notes:
        out["text"] += u" · строк пропущено: %d" % len(notes)
    return out


MANUAL_NOTE_START = u"РУЧНОЙ ВЫБОР"


def has_manual_note(path):
    """Лист собран в режиме ручного выбора (красная пометка в A3): в нём и то, что обычно не вывозят."""
    try:
        wb = load_workbook(path, read_only=True)
        ws = wb[SHEET_VYVOZ] if SHEET_VYVOZ in wb.sheetnames else wb.worksheets[0]
        rows = list(ws.iter_rows(min_row=3, max_row=3, max_col=1, values_only=True))
        wb.close()
        return bool(rows) and str(rows[0][0] or u"").startswith(MANUAL_NOTE_START)
    except Exception:
        return False


def auto_lists(dirs, limit=12):
    """Автоматические листы вывоза ВЫВОЗ\\<дата>[_зачистка]\\<магазин>\\<магазин>.xlsx, новые сверху -> [(подпись, путь)]."""
    found = []
    try:
        for d_ in os.listdir(dirs.out):
            dd = os.path.join(dirs.out, d_)
            if not (re.match(r"^\d{4}-\d{2}-\d{2}", d_) and os.path.isdir(dd)):
                continue
            for sh_ in os.listdir(dd):
                fp = os.path.join(dd, sh_, sh_ + u".xlsx")
                if os.path.isfile(fp):
                    found.append((_mtime(fp), d_, sh_, fp))
    except OSError:
        return []
    found.sort(reverse=True)
    out = []
    for mt, d_, sh_, fp in found[:limit]:
        tag = u" (зачистка)" if d_.endswith(SWEEP_DAY_TAG) else u""
        out.append((u"%s%s · вывоз %s · собран %s" % (sh_, tag, d_[:10], datetime.fromtimestamp(mt).strftime("%d.%m %H:%M")), fp))
    return out


def copy_list_as_is(dirs, src):
    """Автоматический лист вывоза «как есть» -> Корректировка_ЮА в том же виде, что и исправленные листы: один лист
    «Вывезти на склад», колонки № ... Сумма. Существующие файлы не затираются. Исходный список не меняется. -> путь копии"""
    from openpyxl.utils import column_index_from_string
    wb = load_workbook(src)
    try:
        if SHEET_VYVOZ not in wb.sheetnames:
            raise ValueError(u"в файле нет листа «%s»" % SHEET_VYVOZ)
        for n_ in list(wb.sheetnames):
            if n_ != SHEET_VYVOZ:
                del wb[n_]
        ws = wb[SHEET_VYVOZ]
        if ws.max_column > 7:
            ws.delete_cols(8, ws.max_column - 7)
        for letter in list(ws.column_dimensions.keys()):
            if column_index_from_string(letter) > 7:
                del ws.column_dimensions[letter]
        if str(ws["A3"].value or u"").startswith(MANUAL_NOTE_START):
            ws["A3"] = None
        shop = os.path.basename(os.path.dirname(src))
        day = os.path.basename(os.path.dirname(os.path.dirname(src)))[:10]
        base = u"%s_%s" % (re.sub(u'[\\\\/:*?"<>|]+', u"", shop).strip().replace(u" ", u"_"), day)
        dst_dir = corr_dir(dirs)
        os.makedirs(dst_dir, exist_ok=True)
        dst, n = os.path.join(dst_dir, base + u".xlsx"), 2
        while os.path.exists(dst):
            dst, n = os.path.join(dst_dir, u"%s_%d.xlsx" % (base, n)), n + 1
        wb.save(dst)
    finally:
        wb.close()
    return dst


def run_app2(dirs=None, test_mode=False):
    """Окно: боковое меню + разделы (_ui_v2). test_mode: собрать все страницы и закрыть (для самотеста)."""
    dirs = dirs or DIRS
    import io as _io
    import contextlib
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    dirs.ensure()
    app = tk.Tk()
    app.withdraw()
    _set_app_icon(app)
    app.title(u"Фемелі Маркет · Вывоз вне матрицы")
    k = max(1.0, app.winfo_fpixels("1i") / 96.0)
    _fm_theme(app, k)

    def px(v):
        return int(v * k)
    sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
    W, H = min(px(1100), int(sw * 0.85)), min(px(780), int(sh * 0.9))
    app.geometry("%dx%d+%d+%d" % (W, H, (sw - W) // 2, max(0, (sh - H) // 2 - px(40))))
    app.minsize(min(px(860), W), min(px(560), H))

    F = "Segoe UI"
    SIDE, SIDE_HOV, SIDE_ACT, SIDE_FG, ACC = FM_GRAPH, "#34353a", "#3d3e44", "#b4b4ba", FM_RED
    GRN, RED, AMB, MUT = "#16a34a", "#dc2626", "#d97706", "#6b7280"
    sty = ttk.Style(app)
    sty.configure("Hint.TLabel", foreground=MUT)
    sty.configure("H1.TLabel", font=(F, 17, "bold"))
    sty.configure("H2.TLabel", font=(F, 10, "bold"))
    sty.configure("Treeview", rowheight=px(26))
    st = {"found": {}, "info": None, "split": None, "upd": None, "final": None, "busy": False}
    busy = []

    # ---------- каркас ----------
    side = tk.Frame(app, bg=SIDE, width=px(224))
    side.pack(side="left", fill="y")
    side.pack_propagate(False)
    from tkinter import font as tkfont
    _fams = set(tkfont.families(app))
    LF = next((f_ for f_ in (u"Bahnschrift SemiBold Condensed", u"Bahnschrift Condensed", u"Bahnschrift") if f_ in _fams), F)
    hdr_ = tk.Frame(side, bg=SIDE)
    hdr_.pack(fill="x", padx=px(16), pady=(px(20), 0))
    try:
        app._fm_logo = _logo_photo(app, px(46), SIDE)
        tk.Label(hdr_, image=app._fm_logo, bg=SIDE, borderwidth=0).pack(side="left")
    except Exception:
        pass
    tk.Label(hdr_, text=u"ФЕМЕЛІ\nМАРКЕТ", bg=SIDE, fg=FM_RED, font=(LF, 14, "bold"), justify="left").pack(side="left", padx=(px(10), 0))
    tk.Frame(side, bg=FM_RED, height=max(2, px(2))).pack(fill="x", padx=px(16), pady=(px(12), px(6)))
    tk.Label(side, text=u"Вывоз вне матрицы · переход на ЮА", bg=SIDE, fg=SIDE_FG, font=(F, 9), anchor="w").pack(fill="x", padx=px(16), pady=(0, px(14)))
    tk.Label(side, text=u"ПОРЯДОК РАБОТЫ", bg=SIDE, fg="#6f7077", font=(F, 8, "bold"), anchor="w").pack(fill="x", padx=px(20), pady=(0, px(4)))
    main = ttk.Frame(app, padding=(px(26), px(20), px(26), px(10)))
    main.pack(side="left", fill="both", expand=True)
    head = ttk.Frame(main)
    head.pack(fill="x")
    h_title = ttk.Label(head, text=u"", style="H1.TLabel")
    h_title.pack(anchor="w")
    h_sub = ttk.Label(head, text=u"", style="Hint.TLabel", wraplength=W - px(300), justify="left")
    h_sub.pack(anchor="w", pady=(px(2), px(14)))
    sbar = ttk.Frame(main)
    sbar.pack(side="bottom", fill="x", pady=(px(8), 0))
    dot = ttk.Label(sbar, text=u"●", foreground=GRN)
    dot.pack(side="left")
    stat = ttk.Label(sbar, text=u"Готово к работе", style="Hint.TLabel")
    stat.pack(side="left", padx=(px(6), 0))
    prog = ttk.Progressbar(sbar, mode="indeterminate", length=px(160))
    body = ttk.Frame(main)
    body.pack(fill="both", expand=True)

    def status(text, color):
        stat.config(text=text)
        dot.config(foreground=color)

    pages, nav, cur, on_show = {}, {}, {"p": None}, {}
    TITLES = {
        "vyvoz": (u"Списки вывоза", u"Что делает: по остатку магазина составляет список товара, который нужно вывезти на Полевая-Склад (лист для печати и файл для ТСД). Порядок: 1) положите свежие выгрузки; 2) проверьте магазин; 3) нажмите «Сформировать списки вывоза»; 4) откройте список в Excel, удалите лишнее и сохраните в «Корректировка_ЮА»."),
        "split": (u"Распределение по точкам", u"Что делает: делит список вывоза на файлы для ТСД по адресам: товар РЦ - на «Полевая-Склад», прямые поставки - на самые сильные торговые точки. Список берётся ТОЛЬКО из папки «Корректировка_ЮА»: правили его в Excel - сохраните туда; правок нет - нажмите «Взять без правок». На следующий день сюда не возвращайтесь - используйте «Подчистка списка»."),
        "upd": (u"Подчистка списка", u"Что делает: убирает из вашего списка и из готовых файлов ТСД товар, который уже пришёл на ЮА (стал ЮА-шным) - его вывозить не нужно. Берётся ваш исправленный лист и готовые файлы; они не меняются. Очищенные копии лежат в папке Обновление_<дата> - в ТСД берите файлы оттуда."),
        "fin": (u"Инвентаризация ЮА (финал)", u"Выгрузку Family снимайте ПОСЛЕ проведения вывоза. Всё, что осталось на полке, кроме кулинарии и штучных, идёт в файл инвентаризации ЮА с теми ШК, которые знает ЮА. Ваши правки списка вывоза здесь не нужны: файл строится по факту остатка."),
        "serv": (u"Сервис", u"Проверка программы и быстрый доступ к папкам."),
        "log": (u"Журнал", u"Подробный отчёт по всем действиям за сеанс."),
    }

    def show(key):
        if cur["p"]:
            pages[cur["p"]].pack_forget()
        pages[key].pack(fill="both", expand=True)
        cur["p"] = key
        if on_show.get(key):
            on_show[key]()
        for k2, (fr, bar_, lb) in nav.items():
            on = k2 == key
            bg = SIDE_ACT if on else SIDE
            fr.config(bg=bg)
            bar_.config(bg=ACC if on else bg)
            lb.config(bg=bg, fg="#ffffff" if on else SIDE_FG)
        h_title.config(text=TITLES[key][0])
        h_sub.config(text=TITLES[key][1])

    def nav_item(key, text):
        fr = tk.Frame(side, bg=SIDE, cursor="hand2")
        fr.pack(fill="x", padx=px(10), pady=px(2))
        bar_ = tk.Frame(fr, bg=SIDE, width=px(3))
        bar_.pack(side="left", fill="y")
        lb = tk.Label(fr, text=text, bg=SIDE, fg=SIDE_FG, font=(F, 10), anchor="w", padx=px(12), pady=px(8), cursor="hand2")
        lb.pack(side="left", fill="x", expand=True)

        def hover(on):
            if cur["p"] != key:
                c = SIDE_HOV if on else SIDE
                for w in (fr, bar_, lb):
                    w.config(bg=c)
        for w in (fr, lb):
            w.bind("<Button-1>", lambda e: show(key))
            w.bind("<Enter>", lambda e: hover(True))
            w.bind("<Leave>", lambda e: hover(False))
        nav[key] = (fr, bar_, lb)
        pages[key] = ttk.Frame(body)
        return pages[key]

    def card(parent, title=None, expand=False):
        c = ttk.Frame(parent, style="Card.TFrame", padding=(px(16), px(12)))
        c.pack(fill="both" if expand else "x", expand=expand, pady=(0, px(10)))
        if title:
            ttk.Label(c, text=title, style="H2.TLabel").pack(anchor="w", pady=(0, px(6)))
        return c

    def row(parent, pady=0):
        r = ttk.Frame(parent)
        r.pack(fill="x", pady=(pady, 0))
        return r

    def btn(parent, text, cmd, primary=False, lock=True, side_="left"):
        b = ttk.Button(parent, text=text, command=cmd, style=("Accent.TButton" if primary else "TButton"))
        b.pack(side=side_, padx=(0, px(8)) if side_ == "left" else (px(8), 0))
        if lock:
            busy.append(b)
        return b

    def table(parent, cols, height=5, expand=True):
        fr = ttk.Frame(parent)
        fr.pack(fill="both" if expand else "x", expand=expand, pady=(px(6), 0))
        tv = ttk.Treeview(fr, columns=[c[0] for c in cols], show="headings", height=height, selectmode="extended")
        for c, t, w, an in cols:
            tv.heading(c, text=t, anchor=an)
            tv.column(c, width=px(w), anchor=an, stretch=(an == "w"))
        sb = ttk.Scrollbar(fr, command=tv.yview)
        tv.config(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        return tv

    def journal(title, text):
        jt.config(state="normal")
        jt.insert("end", u"%s  %s\n" % (datetime.now().strftime("%H:%M:%S"), title), "h")
        jt.insert("end", (text or u"").rstrip() + u"\n\n")
        jt.see("end")
        jt.config(state="disabled")

    def run_bg(label, fn, done):
        if st["busy"]:
            return
        st["busy"] = True
        for b in busy:
            b.config(state="disabled")
        status(label + u"...", AMB)
        prog.pack(side="right")
        prog.start(12)
        app.config(cursor="watch")
        box = {}

        def work():
            buf = _io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    box["res"] = fn()
            except Exception as e:
                box["err"] = u"%s: %s" % (type(e).__name__, e)
            box["out"] = buf.getvalue()

        th = threading.Thread(target=work, daemon=True)
        th.start()

        def poll():
            if th.is_alive():
                app.after(250, poll)
                return
            st["busy"] = False
            for b in busy:
                b.config(state="normal")
            prog.stop()
            prog.pack_forget()
            app.config(cursor="")
            if "err" in box:
                status(u"Ошибка: %s" % box["err"], RED)
                journal(label + u" - ОШИБКА", box["err"] + u"\n" + box.get("out", u""))
                messagebox.showerror(label, box["err"], parent=app)
            else:
                status(u"Готово: %s" % label, GRN)
                done(box.get("res"), box.get("out", u""))
        app.after(250, poll)

    def nores():
        messagebox.showinfo(u"Нет результата", u"Сначала выполните действие в этом разделе.", parent=app)

    def open_key(d, key):
        v = (st.get(d) or {}).get(key)
        open_path(v) if v and os.path.exists(v) else nores()

    # ================= 1. СПИСКИ ВЫВОЗА =================
    p1h = nav_item("vyvoz", u"1  Списки вывоза")

    def wrap_label(parent, text=u"", style="Hint.TLabel"):
        """Подпись, которая переносится по ширине окна (pack с fill="x")."""
        lb = ttk.Label(parent, text=text, style=style, justify="left", wraplength=px(300))
        lb.bind("<Configure>", lambda e: lb.config(wraplength=max(px(200), e.width - px(4))))
        return lb
    ra = ttk.Frame(p1h)                                     # закреплена внизу страницы: режим и кнопка видны без прокрутки
    ra.pack(side="bottom", fill="x", pady=(px(6), 0))
    ttk.Separator(ra, orient="horizontal").pack(fill="x", pady=(0, px(6)))
    rmode = ttk.Frame(ra)
    rmode.pack(fill="x")
    rb = ttk.Frame(ra)
    rb.pack(fill="x", pady=(px(6), 0))
    btn(rb, u"Сформировать список", lambda: go(), primary=True, side_="right")
    btn(rb, u"Папка результата", lambda: open_key("info", "day_dir"), lock=False)
    btn(rb, u"Сводка (Excel)", lambda: open_key("info", "summary"), lock=False)
    note1 = wrap_label(rb, u"Результат: Excel-лист и файл для ТСД - в папке ВЫВОЗ\\<дата>\\<магазин>.")
    note1.pack(side="left", fill="x", expand=True, padx=px(8))
    cv1 = tk.Canvas(p1h, highlightthickness=0, borderwidth=0, background=ttk.Style(app).lookup("TFrame", "background") or "#f6f6f7", yscrollincrement=px(24))
    sb1 = ttk.Scrollbar(p1h, orient="vertical", command=cv1.yview)
    p1 = ttk.Frame(cv1)
    p1_id = cv1.create_window((0, 0), window=p1, anchor="nw")
    cv1.configure(yscrollcommand=sb1.set)
    p1.bind("<Configure>", lambda e: cv1.configure(scrollregion=cv1.bbox("all")))
    cv1.bind("<Configure>", lambda e: cv1.itemconfigure(p1_id, width=e.width))
    sb1.pack(side="right", fill="y")
    cv1.pack(side="left", fill="both", expand=True)

    def p1_wheel(e):
        """Колесо листает страницу, если она не помещается в окно (над таблицами колесо работает в самих таблицах)."""
        try:
            if cur["p"] != "vyvoz" or cv1.yview() == (0.0, 1.0):
                return
            w_ = app.winfo_containing(e.x_root, e.y_root)
            if w_ is not None and w_.winfo_class() in ("Treeview", "TCombobox", "Text"):
                return
            cv1.yview_scroll(-1 if e.delta > 0 else 1, "units")
        except Exception:
            pass
    app.bind_all("<MouseWheel>", p1_wheel, add="+")

    def drop_prefix(name):
        return name.replace(u"Состояние склада ", u"", 1)

    c = card(p1)
    hr = row(c)
    ttk.Label(hr, text=u"Шаг 1. Проверьте выгрузки в ВХОД_ВЫВОЗ", style="H2.TLabel").pack(side="left")
    btn(hr, u"Открыть папку", lambda: open_input_folder(), lock=False, side_="right")
    btn(hr, u"Перечитать", lambda: refresh(), side_="right")
    tvi = table(c, (("fold", u"Папка", 120, "w"), ("need", u"Нужно", 100, "w"), ("what", u"Что положить", 185, "w"),
                    ("file", u"Файл (самый свежий)", 200, "w"), ("when", u"Выгружено", 85, "center"),
                    ("st", u"Статус", 110, "w")), 5, expand=False)
    for lv_, col_ in (("OK", GRN), ("WARN", AMB), ("ERR", RED), ("INFO", MUT)):
        tvi.tag_configure(lv_, foreground=col_)
    in_folders = {}

    def open_input_folder(*_):
        sel_ = tvi.selection()
        fp_ = in_folders.get(sel_[0]) if sel_ else None
        open_path(fp_ if fp_ and os.path.isdir(fp_) else dirs.inp)
    tvi.bind("<Double-1>", open_input_folder)
    wrap_label(c, u"Зелёное - можно работать; жёлтое или красное - выгрузите заново и нажмите «Перечитать». Склад ЮА - со ВСЕМИ строками, "
                  u"и нулевыми; приходы ЮА - за весь период; ×2 - в папке два файла, берётся самый свежий. Двойной щелчок - открыть папку.").pack(fill="x", pady=(px(4), 0))

    def fill_inputs(rows_):
        tvi.delete(*tvi.get_children())
        in_folders.clear()
        for x in rows_:
            in_folders[x["key"]] = x["folder"]
            tvi.insert("", "end", iid=x["key"], values=(x["fold"], x["need"], x["what"], drop_prefix(x["file"]), x["when"], x["text"]),
                       tags=(x["level"],))

    c = card(p1)
    hr = row(c)
    ttk.Label(hr, text=u"Шаг 2. Выберите магазин", style="H2.TLabel").pack(side="left")
    btn(hr, u"Папка МАГАЗИНЫ", lambda: open_path(dirs.stores), lock=False, side_="right")
    restore_btn = btn(hr, u"Показать скрытые", lambda: restore_hidden(), lock=False, side_="right")
    btn(hr, u"Скрыть", lambda: hide_selected(), lock=False, side_="right")
    btn(hr, u"Выбрать все", lambda: tv1.selection_set(tv1.get_children()), lock=False, side_="right")
    tv1 = table(c, (("shop", u"Магазин", 170, "w"), ("file", u"Файл выгрузки", 330, "w"), ("time", u"Выгружено", 95, "center"),
                    ("age", u"Свежесть", 150, "w")), 2, expand=False)
    for lv_, col_ in (("OK", GRN), ("WARN", AMB), ("ERR", RED), ("INFO", MUT)):
        tv1.tag_configure(lv_, foreground=col_)
    pick_lbl = ttk.Label(c, text=u"", style="H2.TLabel")
    pick_lbl.pack(anchor="w", pady=(px(6), 0))
    hint2 = (u"Выбран магазин с самой свежей выгрузкой; нужен другой - щёлкните по нему (Ctrl / Shift - несколько). «Скрыть» убирает "
             u"магазин только из таблицы: файл выгрузки остаётся в папке.")
    hid_lbl = wrap_label(c, hint2)
    hid_lbl.pack(fill="x", pady=(px(2), 0))

    def upd_store_row(sel_):
        """Строка шага 1 «склад магазина» показывает выгрузку выбранного магазина (из нескольких - самую старую)."""
        if "stores" not in tvi.get_children():
            return
        paths = [st["found"][x][0] for x in sel_ if x in st["found"]]
        if not paths:
            return
        worst = min(paths, key=_file_stamp)
        stamp, age, lv, txt = file_state(worst)
        tvi.set("stores", "file", drop_prefix(os.path.basename(worst)))
        tvi.set("stores", "when", stamp.strftime("%d.%m %H:%M"))
        tvi.set("stores", "st", txt)
        tvi.item("stores", tags=(lv,))

    def on_pick(*_):
        sel_ = list(tv1.selection())
        if not sel_:
            pick_lbl.config(text=u"Магазин не выбран - выделите строку в таблице", foreground=RED)
        elif len(sel_) == 1:
            pick_lbl.config(text=u"Будет сформирован список для: %s" % sel_[0], foreground=GRN)
        else:
            pick_lbl.config(text=u"Будут сформированы списки для %d магазинов: %s - так и задумано?" % (len(sel_), u"; ".join(sel_)), foreground=AMB)
        upd_store_row(sel_)
    tv1.bind("<<TreeviewSelect>>", on_pick)

    def render_stores(keep=None):
        hidden = load_hidden_stores(dirs)
        shown = visible_stores(st["found"], hidden)
        tv1.delete(*tv1.get_children())
        for x, v in shown.items():
            fp = v[0] if isinstance(v, (tuple, list)) else u""
            try:
                stamp, age, lv, txt = file_state(fp)
                tm = stamp.strftime("%d.%m %H:%M")
            except Exception:
                tm, lv, txt = u"", "INFO", u""
            tv1.insert("", "end", iid=x, values=(x, os.path.basename(fp), tm, txt), tags=(lv,))
        pick = [x for x in (keep or []) if x in shown] or ([pick_default_store(shown)] if shown else [])
        pick = [x for x in pick if x]
        if pick:
            tv1.selection_set(pick)
        n_h = len(st["found"]) - len(shown)
        restore_btn.config(state="normal" if n_h else "disabled")
        hid_lbl.config(text=hint2 + ((u" СКРЫТО магазинов: %d (%s) - вернуть кнопкой «Показать скрытые»." % (n_h, u"; ".join(sorted(x for x in st["found"] if x not in shown)))) if n_h else u""))
        on_pick()

    def hide_selected():
        sel_ = list(tv1.selection())
        if not sel_:
            messagebox.showinfo(u"Скрыть магазин", u"Выделите магазин, который нужно скрыть из списка.", parent=app)
            return
        if not messagebox.askyesno(u"Скрыть магазин", u"Скрыть магазин из списка: %s?\n\nФайл выгрузки остаётся в папке, ничего не удаляется. "
                                   u"Вернуть магазин - кнопка «Показать скрытые»." % u"; ".join(sel_), parent=app):
            return
        try:
            save_hidden_stores(dirs, load_hidden_stores(dirs) | set(sel_))
        except Exception as e:
            messagebox.showerror(u"Скрыть магазин", u"Не удалось сохранить: %s" % e, parent=app)
            return
        render_stores()

    def restore_hidden():
        try:
            save_hidden_stores(dirs, set())
        except Exception as e:
            messagebox.showerror(u"Показать скрытые", u"Не удалось сохранить: %s" % e, parent=app)
            return
        render_stores()

    ttk.Label(rmode, text=u"Шаг 3. Режим:", style="H2.TLabel").pack(side="left")
    opts_btn = btn(rmode, u"Настройки ▾", lambda: toggle_opts(), lock=False, side_="right")
    mode_var = tk.StringVar(value="manual")
    manual_var = tk.BooleanVar(value=True)
    sweep_var = tk.BooleanVar(value=False)

    def on_mode(*_):
        m_ = mode_var.get()
        manual_var.set(m_ == "manual")
        sweep_var.set(m_ == "sweep")
    for val_, txt_ in (("normal", u"Обычный"), ("manual", u"Ручной выбор"), ("sweep", u"Зачистка остатка")):
        ttk.Radiobutton(rmode, text=txt_, value=val_, variable=mode_var, command=on_mode).pack(side="left", padx=(px(14), 0))
    mrow = ttk.Frame(ra)
    mrow.pack(fill="x", pady=(px(4), 0), after=rmode)
    mode_lbl = wrap_label(mrow, u"", style="Hint.TLabel")
    mode_lbl.pack(side="left", fill="x", expand=True)
    opts = ttk.Frame(ra)                                    # настройки режима: свёрнуты, пока их не откроют
    rm_var = tk.BooleanVar(value=True)
    rm_cb = ttk.Checkbutton(opts, text=u"Вывозить и товар из матрицы, который на ЮА ещё не завозился", variable=rm_var)
    rm_cb.pack(anchor="w", pady=(px(8), 0))

    def open_keep():
        load_keep_shk(dirs)
        open_path(os.path.join(dirs.cache, KEEP_FILE))
    r = row(opts, px(8))
    btn(r, u"Поставщики: кого не вывозить...", lambda: open_suppliers_dialog(app, dirs, st["info"], on_save=sup_saved), lock=False)
    btn(r, u"Список ШК «не вывозить» (Excel)...", open_keep, lock=False)
    sup_lbl = wrap_label(opts, u"")
    sup_lbl.pack(fill="x", pady=(px(4), 0))

    def sup_refresh():
        n_, when_, _p = suppliers_file_info(dirs)
        if sweep_var.get():
            t_ = u"В зачистке список поставщиков не действует."
        elif manual_var.get():
            t_ = u"Не вывозим поставщиков: %d (список сохранён %s). Галочки действуют и в ручном выборе." % (n_, when_ or u"-")
        else:
            t_ = u"Не вывозим поставщиков: %d (список сохранён %s)." % (n_, when_ or u"-")
        sup_lbl.config(text=t_)

    def mode_refresh(*_):
        n_, when_, _p = suppliers_file_info(dirs)
        if sweep_var.get():
            t_, col_ = u"Вывозим всё, кроме заморозки, скоропортов и расходников. Матрица, ЮА и список поставщиков не учитываются.", AMB
        elif manual_var.get():
            t_, col_ = (u"Сигареты и кеги тоже попадут в список (причина - в последней колонке листа), лишнее удалите в Excel. Не вывозим: %d поставщиков "
                        u"с галочкой, список ШК, пакеты, расходники, сырьё кофе, овощи, кулинарию и выпечку." % n_), GRN
        else:
            t_, col_ = (u"Не вывозим сигареты, кеги, расходники, кулинарию, список ШК и %d поставщиков из списка." % n_), MUT
        if rm_var.get() and not sweep_var.get():
            t_ += u" Вывозим и товар из матрицы, который на ЮА ещё не завозился."
        mode_lbl.config(text=t_, foreground=col_)
        sup_refresh()

    def sup_saved(n):
        mode_refresh()
        status(u"Поставщики сохранены: не вывозим %d" % n, GRN)

    def on_sweep(*_):
        rm_cb.config(state="disabled" if sweep_var.get() else "normal")
        mode_refresh()

    def toggle_opts():
        if opts.winfo_ismapped():
            opts.pack_forget()
            opts_btn.config(text=u"Настройки ▾")
        else:
            opts.pack(fill="x", pady=(px(2), 0), after=mrow)
            opts_btn.config(text=u"Свернуть ▴")
    for v_ in (rm_var, manual_var):
        v_.trace_add("write", mode_refresh)
    sweep_var.trace_add("write", on_sweep)
    mode_refresh()

    tv1r = table(p1, (("shop", u"Магазин", 200, "w"), ("pos", u"Вывезти, поз.", 100, "e"), ("sum", u"Сумма, грн", 110, "e"),
                      ("c1", u"Проверить 1", 95, "e"), ("c2", u"Проверить 2", 95, "e"), ("status", u"Статус", 200, "w")), 3, expand=False)

    def refresh():
        def job():
            reset_report()
            return scan_stores(dirs), input_status(dirs)

        def done(res, out):
            found, inputs = res or ({}, [])
            st["found"] = {x: found[x] for x in sorted(found) if not _is_polevaya(x)}
            fin_cb.config(values=list(st["found"]))
            fill_inputs(inputs)
            render_stores(keep=list(tv1.selection()))
            errs = [u"[%s] %s: %s" % (VERDICT_LABEL[a], b, c_) for a, b, c_ in R.checks if a in ("ERROR", "WARN")]
            status(u"Магазинов найдено: %d" % len(st["found"]) if st["found"] else u"В ВХОД_ВЫВОЗ\\МАГАЗИНЫ нет выгрузок", GRN if st["found"] else AMB)
            if errs:
                journal(u"Поиск выгрузок", u"\n".join(errs))
        run_bg(u"Поиск выгрузок магазинов", job, done)

    def go():
        shops = list(tv1.selection())
        if not shops:
            messagebox.showinfo(u"Списки вывоза", u"Выделите хотя бы один магазин.", parent=app)
            return
        if len(shops) > 1 and not messagebox.askyesno(u"Списки вывоза", u"Выбрано магазинов: %d (%s).\n\nФормировать списки по всем?"
                                                      % (len(shops), u"; ".join(shops)), parent=app, default="no"):
            return
        rm, sw, mn = bool(rm_var.get()), bool(sweep_var.get()), bool(manual_var.get())
        mode_refresh()

        def done(info, out):
            st["info"] = info or {}
            tv1r.delete(*tv1r.get_children())
            for sh in st["info"].get("shops", []):
                tv1r.insert("", "end", values=(sh.get("shop", u""), sh.get("vyvoz_pos", u""),
                                               u"{:,.0f}".format(float(sh.get("vyvoz_sum", 0) or 0)).replace(",", u" "),
                                               sh.get("check1", u""), sh.get("check2", u""), sh.get("status", u"")))
            journal(u"Списки вывоза", _result_text(st["info"]))
            ne = sum(1 for a, _, _ in R.checks if a == "ERROR")
            note1.config(text=(u"Ошибок: %d - см. «Журнал»" % ne) if ne else u"Подробности - в «Журнале»",
                         foreground=RED if ne else MUT)
            if st["info"].get("shops"):
                note1.config(text=note1.cget("text") + u" · дальше: исправьте лист в Excel и сохраните в «%s», либо во вкладке «Распределение» нажмите «Взять без правок»" % CORR_DIR)
            dd = st["info"].get("day_dir")
            names = [x.get("shop", u"") for x in st["info"].get("shops", [])]
            sel_fp = None
            if dd and names:
                for nm in (names[0], rz.safe_name(names[0])):
                    fp = os.path.join(dd, nm, nm + u".xlsx")
                    if os.path.isfile(fp):
                        sel_fp = fp
                        break
            refresh_auto(select=sel_fp)
            app.after(150, lambda: cv1.yview_moveto(1.0))
        run_bg(u"Формирование списков вывоза", lambda: run_vyvoz(dirs, shops=shops, remove_not_on_ua=rm, sweep=sw, manual=mn), done)

    # ================= 2. РАСПРЕДЕЛЕНИЕ =================
    def corrected_default():
        """Самый свежий исправленный лист в Корректировка_ЮА (временные ~$ и файлы «проверка» не берём)."""
        fs = list_corrected(dirs)
        return os.path.normpath(fs[0]) if fs else u""

    def pick_into(var, title):
        init = os.path.join(dirs.base, CORR_DIR)
        if not os.path.isdir(init):
            init = dirs.out if os.path.isdir(dirs.out) else dirs.base
        fp = filedialog.askopenfilename(parent=app, title=title, initialdir=init,
                                        filetypes=[(u"Excel", "*.xlsx *.xlsm"), (u"Все файлы", "*.*")])
        if fp:
            var.set(os.path.normpath(fp))

    def raw_list_ok(fp, what):
        """Автоматический лист из ВЫВОЗ\\<дата>\\<магазин> - не тот, что вы исправляли: без подтверждения по нему не работаем."""
        try:
            inside = os.path.normcase(os.path.abspath(fp)).startswith(os.path.normcase(os.path.abspath(dirs.out)) + os.sep)
        except Exception:
            inside = False
        return (not inside) or messagebox.askyesno(
            what, u"Это автоматический лист из папки ВЫВОЗ, а не исправленный вами (он лежит в «%s»).\n\n"
                  u"Продолжить именно с ним?" % CORR_DIR, parent=app, default="no")

    p2 = nav_item("split", u"2  Распределение")
    c = card(p2, u"Какой список распределяем (берётся только из папки «%s»)" % CORR_DIR)
    r = row(c)
    sp_var = tk.StringVar(value=u"")                        # полный путь выбранного листа (всегда из Корректировка_ЮА)
    sp_cb = ttk.Combobox(r, state="readonly", values=[])
    sp_cb.pack(side="left", fill="x", expand=True, padx=(0, px(8)))
    corr_paths, auto_paths = [], []

    def corr_dir_made():
        d_ = corr_dir(dirs)
        os.makedirs(d_, exist_ok=True)
        return d_
    btn(r, u"Открыть папку «%s»" % CORR_DIR, lambda: open_path(corr_dir_made()), lock=False, side_="right")
    btn(r, u"Перечитать папку", lambda: refresh_corr(), lock=False, side_="right")
    sp_info = ttk.Label(c, text=u"", style="Hint.TLabel", wraplength=W - px(320), justify="left")
    sp_info.pack(anchor="w", pady=(px(4), 0))
    r = row(c, px(12))
    ttk.Label(r, text=u"Не правили список? Возьмите автоматический:").pack(side="left")
    auto_cb = ttk.Combobox(r, state="readonly", values=[], width=50)
    auto_cb.pack(side="left", padx=px(8))
    btn(r, u"Взять без правок", lambda: take_as_is(), lock=False)
    ttk.Label(c, text=u"«Взять без правок» кладёт в «%s» копию листа «Вывезти на склад» (7 колонок, как в исправленных листах); "
                      u"сам автоматический список не меняется." % CORR_DIR,
              style="Hint.TLabel", wraplength=W - px(320), justify="left").pack(anchor="w", pady=(px(4), 0))

    def show_sp_info():
        fp = sp_var.get()
        if not fp:
            sp_info.config(text=u"В «%s» нет листов. Исправьте список вывоза в Excel и сохраните сюда или возьмите автоматический ниже." % CORR_DIR,
                           foreground=AMB)
            return
        chk = corrected_check(fp)
        sp_info.config(text=chk["text"], foreground=GRN if chk["ok"] else RED)

    def on_sp_pick(*_):
        i = sp_cb.current()
        sp_var.set(corr_paths[i] if 0 <= i < len(corr_paths) else u"")
        show_sp_info()
    sp_cb.bind("<<ComboboxSelected>>", on_sp_pick)

    def _ix(paths, want):
        w_ = os.path.normcase(want or u"")
        return next((i for i, p_ in enumerate(paths) if os.path.normcase(p_) == w_), 0 if paths else -1)

    def refresh_corr(select=None):
        corr_paths[:] = list_corrected(dirs)
        sp_cb.config(values=[u"%s    (%s)" % (os.path.basename(p_), datetime.fromtimestamp(_mtime(p_)).strftime("%d.%m %H:%M")) for p_ in corr_paths])
        ix = _ix(corr_paths, select or sp_var.get())
        if ix >= 0:
            sp_cb.current(ix)
        else:
            sp_cb.set(u"")
        on_sp_pick()

    def refresh_auto(select=None):
        items = auto_lists(dirs)
        auto_paths[:] = [fp for _l, fp in items]
        auto_cb.config(values=[l_ for l_, _fp in items])
        ix = _ix(auto_paths, select)
        if ix >= 0:
            auto_cb.current(ix)
        else:
            auto_cb.set(u"")

    def take_as_is():
        i = auto_cb.current()
        if i < 0 or i >= len(auto_paths):
            messagebox.showinfo(u"Взять без правок", u"Автоматических списков пока нет: сформируйте список на вкладке «Списки вывоза».", parent=app)
            return
        src = auto_paths[i]
        if has_manual_note(src) and not messagebox.askyesno(
                u"Взять без правок", u"Это список РУЧНОГО ВЫБОРА: в нём сигареты и кеги, которые обычно не вывозят.\nЕсли вы ничего не "
                                   u"удаляли, они поедут в распределение (кеги - на точки, где они продаются).\n\nВзять список как есть?",
                parent=app, default="no"):
            return
        try:
            dst = copy_list_as_is(dirs, src)
        except Exception as e:
            messagebox.showerror(u"Взять без правок", u"Не получилось: %s" % e, parent=app)
            return
        journal(u"Список вывоза взят без правок", u"%s\n->  %s" % (src, dst))
        refresh_corr(select=dst)
        status(u"Список взят без правок: %s" % os.path.basename(dst), GRN)
    on_show["split"] = lambda: (refresh_corr(), refresh_auto())

    c = card(p2, u"Параметры распределения")
    r = row(c)
    ttk.Label(r, text=u"На сколько торговых точек делить прямые поставки:").pack(side="left")
    top_var = tk.IntVar(value=SPLIT_TOP_N)
    ttk.Spinbox(r, from_=1, to=SPLIT_TOP_N, textvariable=top_var, width=4, state="readonly").pack(side="left", padx=(px(8), px(4)))
    ttk.Label(r, text=u"(из %d самых сильных по обороту)" % SPLIT_TOP_N, style="Hint.TLabel").pack(side="left")
    r = row(c, px(8))
    rc_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(r, text=u"Сделать файл на %s: туда идёт товар РЦ, на точки он не делится" % SPLIT_RC_NAME, variable=rc_var).pack(side="left")
    r = row(c, px(8))
    auto_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(r, text=u"Подобрать число точек по сумме: 5-7 тыс. грн прямых поставок на точку (число выше тогда не учитывается)",
                    variable=auto_var).pack(side="left")

    def open_split_xlsx():
        d = (st["split"] or {}).get("out_dir")
        fp = os.path.join(d, u"Распределение.xlsx") if d else u""
        open_path(fp) if fp and os.path.isfile(fp) else nores()

    r = row(p2)
    btn(r, u"Распределить", lambda: split_go(), primary=True, side_="right")
    btn(r, u"Папка файлов ТСД", lambda: open_key("split", "out_dir"), lock=False)
    btn(r, u"Распределение (Excel)", open_split_xlsx, lock=False)
    note2 = ttk.Label(r, text=u"Файлы ТСД и Распределение.xlsx - в «ТСД_по_адресам» рядом со списком.", style="Hint.TLabel")
    note2.pack(side="left", padx=px(8))
    tv2 = table(p2, (("addr", u"Куда (адрес)", 280, "w"), ("pos", u"Позиций", 90, "e"), ("qty", u"Единиц", 100, "e"),
                     ("sum", u"Сумма, грн", 120, "e")), 6)

    def split_go():
        fp = sp_var.get().strip()
        if not fp or not os.path.isfile(fp):
            messagebox.showinfo(u"Распределение", u"Выберите лист из папки «%s» или возьмите автоматический список («Взять без правок»)." % CORR_DIR, parent=app)
            return
        if not is_in_corr(dirs, fp):
            messagebox.showwarning(u"Распределение", u"Распределение берёт листы только из папки «%s»." % CORR_DIR, parent=app)
            return
        chk = corrected_check(fp)
        if not chk["ok"]:
            messagebox.showwarning(u"Распределение", chk["text"], parent=app)
            return
        n, skip, auto = _clamp_top(top_var.get()), not bool(rc_var.get()), bool(auto_var.get())

        def done(info, out):
            st["split"] = info or {}
            tv2.delete(*tv2.get_children())
            for a, kk, q, sm in st["split"].get("summary", []):
                tv2.insert("", "end", values=(a, kk, u"%.3f" % q, u"{:,.2f}".format(sm).replace(",", u" ")))
            journal(u"Распределение на %d точек" % n, _split_text(st["split"]))
            ok = st["split"].get("ok")
            note2.config(text=u"Подробности - в «Журнале»" if ok else u"Есть ошибки - см. «Журнал»", foreground=MUT if ok else RED)
            if not ok:
                status(u"Распределение: есть ошибки", RED)
        run_bg(u"Распределение на %d точек" % n if not auto else u"Распределение (число точек по сумме)",
               lambda: run_split(dirs, fp, top_n=n, skip_rc=skip, auto_points=auto), done)

    # ================= НОВЫЙ ДЕНЬ =================
    p5 = nav_item("upd", u"3  Подчистка списка")
    c = card(p5, u"Исправленный лист вывоза (не меняется)")
    r = row(c)
    up_var = tk.StringVar(value=corrected_default())
    ttk.Entry(r, textvariable=up_var).pack(side="left", fill="x", expand=True, padx=(0, px(8)))
    btn(r, u"Выбрать лист...", lambda: pick_into(up_var, u"Исправленный лист вывоза"), lock=False, side_="right")

    def file_card(title, hint_text, key, init, clear=False):
        c_ = card(p5, title)
        r_ = row(c_)
        var = tk.StringVar(value=u"")
        ttk.Entry(r_, textvariable=var, state="readonly").pack(side="left", fill="x", expand=True, padx=(0, px(8)))

        def pick_():
            fps = filedialog.askopenfilenames(parent=app, title=title, initialdir=init if os.path.isdir(init) else dirs.base,
                                              filetypes=[(u"Excel", "*.xlsx *.xlsm *.xls"), (u"Все файлы", "*.*")])
            if fps:
                st[key] = [os.path.normpath(x) for x in fps]
                var.set(u"; ".join(os.path.basename(x) for x in st[key]))

        def clr_():
            st[key] = []
            var.set(u"")
        if clear:
            btn(r_, u"Очистить выбор", clr_, lock=False, side_="right")
        btn(r_, u"Выбрать файл(ы)...", pick_, lock=False, side_="right")
        ttk.Label(c_, text=hint_text, style="Hint.TLabel", wraplength=W - px(320), justify="left").pack(anchor="w", pady=(px(4), 0))
    file_card(u"Склад ЮА - обязательно",
              u"Удаляем позиции с любым маршрутом, если ШК есть в выгрузке - с любым остатком: плюс, ноль или минус.",
              "upd_wh", getattr(dirs, "ua_wh", dirs.inp))
    file_card(u"Магазин в базе ЮА - по желанию",
              u"Только для прямой доставки: удаляем, если ШК есть в «Состоянии склада» этого магазина в базе ЮА.",
              "upd_shop", dirs.inp, clear=True)

    def open_removed():
        d = (st.get("upd") or {}).get("out_dir")
        fp = os.path.join(d, UPD_REMOVED_FILE) if d else u""
        open_path(fp) if fp and os.path.isfile(fp) else nores()

    r = row(p5)
    btn(r, u"Подчистить список", lambda: upd_go(), primary=True, side_="right")
    btn(r, u"Папка обновления", lambda: open_key("upd", "out_dir"), lock=False)
    btn(r, u"Список удалённого (Excel)", open_removed, lock=False)
    note3 = ttk.Label(r, text=u"", style="Hint.TLabel")
    note3.pack(side="left", padx=px(8))
    tv3 = table(p5, (("bc", u"ШК", 120, "w"), ("name", u"Название", 260, "w"), ("qty", u"Кол-во", 70, "e"),
                     ("why", u"Причина", 200, "w"), ("where", u"Куда должно было ехать", 200, "w")), 4)

    def upd_go():
        fp = up_var.get().strip()
        if not fp or not os.path.isfile(fp):
            messagebox.showinfo(u"Подчистка списка", u"Выберите исправленный лист вывоза (xlsx).", parent=app)
            return
        if not raw_list_ok(fp, u"Подчистка списка"):
            return
        wh = list(st.get("upd_wh") or [])
        if not wh:
            messagebox.showinfo(u"Подчистка списка", u"Выберите «Состояние склада» Склад ЮА.", parent=app)
            return
        shp = list(st.get("upd_shop") or [])

        def done(info, out):
            st["upd"] = info or {}
            tv3.delete(*tv3.get_children())
            for d in st["upd"].get("removed", []):
                tv3.insert("", "end", values=(d["bc"], d["name"], d["qty"], d["why"], d["where"]))
            journal(u"Подчистка списка", _update_text(st["upd"]))
            ok, n = st["upd"].get("ok"), len(st["upd"].get("removed", []))
            warn = any(x.startswith(u"ВНИМАНИЕ") for x in st["upd"].get("notes", []))
            note3.config(text=(u"Ошибки - см. «Журнал»" if not ok else
                               (u"Удалено: %d" % n if n else u"Ничего не стало ЮА-шным")) + (u" · есть предупреждения" if warn else u""),
                         foreground=RED if (not ok or warn) else MUT)
            if not ok:
                status(u"Подчистка списка: есть ошибки", RED)
        run_bg(u"Подчистка списка", lambda: run_update(dirs, fp, wh, shop_paths=shp), done)

    # ================= ИНВЕНТАРИЗАЦИЯ ЮА (ФИНАЛ) =================
    pf = nav_item("fin", u"4  Инвентаризация ЮА")
    c = card(pf, u"Финальная выгрузка Family")
    r = row(c)
    ttk.Label(r, text=u"Магазин:").pack(side="left")
    fin_shop = tk.StringVar(value=u"")
    fin_cb = ttk.Combobox(r, textvariable=fin_shop, values=[], width=32)
    fin_cb.pack(side="left", padx=(px(8), 0))
    r = row(c, px(8))
    ttk.Label(r, text=u"Файл:").pack(side="left")
    fin_file = tk.StringVar(value=u"")
    ttk.Entry(r, textvariable=fin_file).pack(side="left", fill="x", expand=True, padx=(px(8), px(8)))

    def fin_pick():
        init = dirs.stores if os.path.isdir(dirs.stores) else dirs.base
        fp_ = filedialog.askopenfilename(parent=app, title=u"Финальная выгрузка Family", initialdir=init,
                                         filetypes=[(u"Excel", "*.xlsx *.xlsm *.xls"), (u"Все файлы", "*.*")])
        if fp_:
            fin_file.set(os.path.normpath(fp_))
    btn(r, u"Выбрать выгрузку...", fin_pick, lock=False, side_="right")
    ttk.Label(c, text=u"Файл можно не выбирать: тогда берётся свежая выгрузка магазина из ВХОД_ВЫВОЗ\\МАГАЗИНЫ.",
              style="Hint.TLabel").pack(anchor="w", pady=(px(4), 0))

    def open_final_xlsx():
        d_ = (st["final"] or {}).get("out_dir")
        fs_ = [f_ for f_ in (os.listdir(d_) if d_ and os.path.isdir(d_) else []) if f_.lower().endswith(".xlsx")]
        open_path(os.path.join(d_, fs_[0])) if fs_ else nores()

    r = row(pf)
    btn(r, u"Собрать файл инвентаризации", lambda: fin_go(), primary=True, side_="right")
    btn(r, u"Папка результата", lambda: open_key("final", "out_dir"), lock=False)
    btn(r, u"Сверка (Excel)", open_final_xlsx, lock=False)
    note4 = ttk.Label(r, text=u"", style="Hint.TLabel")
    note4.pack(side="left", padx=px(8))
    tvf = table(pf, (("part", u"Часть", 300, "w"), ("pos", u"Позиций", 90, "e"), ("qty", u"Единиц", 100, "e"),
                     ("sum", u"Сумма, грн", 120, "e")), 6)

    def fin_go():
        shop_ = fin_shop.get().strip()
        if not shop_:
            messagebox.showinfo(u"Инвентаризация ЮА", u"Выберите магазин.", parent=app)
            return
        fp_ = fin_file.get().strip() or None
        if fp_ and not os.path.isfile(fp_):
            messagebox.showinfo(u"Инвентаризация ЮА", u"Файл не найден.", parent=app)
            return
        if not messagebox.askyesno(u"Инвентаризация ЮА", u"Выгрузка Family снята ПОСЛЕ проведения вывоза?\n\n"
                                   u"Если нет, в инвентаризацию попадёт и то, что ещё должно уехать.", parent=app, default="no"):
            return

        def done(info, out):
            st["final"] = info or {}
            tvf.delete(*tvf.get_children())
            for a_, n_, q_, sm_ in st["final"].get("summary", []):
                tvf.insert("", "end", values=(a_, n_, u"%.3f" % q_, u"{:,.2f}".format(sm_).replace(",", u" ")))
            journal(u"Инвентаризация ЮА (финал)", _final_text(st["final"]))
            ok_ = st["final"].get("ok")
            note4.config(text=u"Подробности - в «Журнале»" if ok_ else u"Есть ошибки - см. «Журнал»", foreground=MUT if ok_ else RED)
            if not ok_:
                status(u"Инвентаризация ЮА: есть ошибки", RED)
        run_bg(u"Финальная инвентаризация", lambda: run_final_inventory(dirs, shop_, state_path=fp_), done)

    # ================= 3. СЕРВИС =================
    tk.Frame(side, bg="#3a3b40", height=1).pack(fill="x", padx=px(20), pady=px(10))
    p3 = nav_item("serv", u"Сервис")
    c = card(p3, u"Проверка")
    r = row(c)
    btn(r, u"Запустить самотест", lambda: run_bg(u"Самотест", self_test, lambda ok, out: (
        journal(u"Самотест: %s" % (u"OK" if ok else u"ПРОВАЛ"), out), show("log"))), primary=True)
    ttk.Label(r, text=u"проверяет расчёты на тестовых данных, результат - в «Журнале»", style="Hint.TLabel").pack(side="left")
    c = card(p3, u"Таблица соответствия ШК Family - ЮА")
    r = row(c)

    def map_list_done(res, out):
        journal(u"Список соответствия ШК", _map_text(res or {}))
        if (res or {}).get("path"):
            open_path(res["path"])
        else:
            messagebox.showwarning(u"Список соответствия", _map_text(res or {}) or u"Не получилось", parent=app)

    def map_import_go():
        fp = filedialog.askopenfilename(parent=app, title=u"Список соответствия с вашими решениями",
                                        initialdir=os.path.join(dirs.base, MAP_DIR) if os.path.isdir(os.path.join(dirs.base, MAP_DIR)) else dirs.base,
                                        filetypes=[(u"Excel", "*.xlsx *.xlsm"), (u"Все файлы", "*.*")])
        if not fp:
            return

        def done(res, out):
            journal(u"Загрузка решений по соответствию ШК", _map_text(res or {}))
            messagebox.showinfo(u"Соответствие ШК", _map_text(res or {}), parent=app)
        run_bg(u"Загружаю решения", lambda: run_map_import(dirs, fp), done)
    btn(r, u"Составить список кандидатов (Excel)", lambda: run_bg(u"Список соответствия ШК", lambda: run_map_list(dirs), map_list_done), primary=True)
    btn(r, u"Загрузить мои решения из Excel...", map_import_go)
    btn(r, u"Открыть таблицу пар (Excel)", lambda: (load_ua_map(dirs), open_path(os.path.join(dirs.cache, UA_MAP_FILE))), lock=False)
    c = card(p3, u"Папки")
    r = row(c)
    for text, pth in ((u"Данные", dirs.base), (u"Входные выгрузки", dirs.inp), (u"Результаты", dirs.out),
                      (u"Справочник", dirs.cache), (u"Логи", dirs.logs)):
        btn(r, text, (lambda q=pth: open_path(q)), lock=False)

    # ================= 4. ЖУРНАЛ =================
    p4 = nav_item("log", u"Журнал")
    c = card(p4, expand=True)
    r = row(c)

    def jclear():
        jt.config(state="normal")
        jt.delete("1.0", "end")
        jt.config(state="disabled")

    def jcopy():
        app.clipboard_clear()
        app.clipboard_append(jt.get("1.0", "end"))
        status(u"Журнал скопирован", GRN)
    btn(r, u"Скопировать", jcopy, lock=False)
    btn(r, u"Очистить", jclear, lock=False)
    fr = ttk.Frame(c)
    fr.pack(fill="both", expand=True, pady=(px(8), 0))
    jt = tk.Text(fr, wrap="word", font=("Consolas", 9), relief="flat", borderwidth=0, padx=px(8), pady=px(6),
                 background="#ffffff", foreground="#111827", height=10)
    jt.tag_configure("h", font=(F, 9, "bold"), foreground=ACC)
    jsb = ttk.Scrollbar(fr, command=jt.yview)
    jt.config(yscrollcommand=jsb.set, state="disabled")
    jt.pack(side="left", fill="both", expand=True)
    jsb.pack(side="right", fill="y")

    tk.Label(side, text=u"Данные:\n%s" % dirs.base, bg=SIDE, fg="#6b7280", font=(F, 8), justify="left",
             anchor="w", wraplength=px(180)).pack(side="bottom", fill="x", padx=px(20), pady=px(14))

    if getattr(app, "_fm_icon_err", u""):
        journal(u"Иконка окна не установлена", app._fm_icon_err)
    show("vyvoz")
    refresh_corr()
    refresh_auto()
    if test_mode:
        app.update()
        for k_ in list(pages):                      # каждая страница показывается без ошибок
            show(k_)
        res_ = (sorted(pages), sp_var.get(), up_var.get())
        app.destroy()
        return res_
    app.deiconify()
    app.after(200, refresh)
    app.mainloop()
    return 0


def main(argv=None):
    global BQ_OFFLINE
    args = _parse_args(sys.argv[1:] if argv is None else argv)
    if args["offline"]:
        BQ_OFFLINE = True
    ok = self_test(gui=args["selftest"])
    print(u"САМОТЕСТ: %s" % (u"OK" if ok else u"ПРОВАЛ"))
    if args["selftest"]:
        return 0 if ok else 1
    if not ok:
        print(u"Самотестирование не пройдено, работа прервана.")
        return 2
    if args["update"]:
        info = run_update(DIRS, args["update"], args["state"], shop_paths=args["shop_ua"])
        print(_update_text(info))
        return 0 if info.get("ok") else 2
    if args["final"]:
        info = run_final_inventory(DIRS, args["final"], state_path=args["final_file"])
        print(_final_text(info))
        return 0 if info.get("ok") else 2
    if args["split"]:
        info = run_split(DIRS, args["split"], shop=(args["shops"] or [None])[0], day=args["day"], top_n=args["top"], skip_rc=args["no_rc"],
                         auto_points=args["auto_points"])
        print(_split_text(info))
        return 0 if info.get("ok") else 2
    if args["map_list"]:
        info = run_map_list(DIRS)
        print(_map_text(info))
        return 0 if info.get("ok") else 2
    if args["map_import"]:
        info = run_map_import(DIRS, args["map_import"])
        print(_map_text(info))
        return 0 if info.get("ok") else 2
    if args["sweep"]:
        if not args["shops"]:
            print(u"--sweep требует --shops \"Магазин\"")
            return 2
        info = run_vyvoz(DIRS, shops=args["shops"], use_ua=not args["no_ua"], sweep=True)
        print(_result_text(info))
        ok_all = bool(info.get("ok"))
        if ok_all and (args["auto_points"] or args["top"]):
            for sh in info.get("shops", []):
                if sh.get("status") != "OK" or not sh.get("folder"):
                    continue
                lp = os.path.join(sh["folder"], os.path.basename(sh["folder"]) + ".xlsx")
                sp = run_split(DIRS, lp, top_n=args["top"], skip_rc=args["no_rc"], auto_points=args["auto_points"])
                print(_split_text(sp))
                ok_all = ok_all and bool(sp.get("ok"))
        return 0 if ok_all else 2
    if USE_GUI and not args["auto"] and args["shops"] is None:
        try:
            return run_app2()
        except Exception as e:
            print(u"Окно не открылось (%s: %s). Запустите с --auto или --shops \"А;Б\"" % (type(e).__name__, e))
            return 2
    excl = set(load_excluded_suppliers(DIRS)) | set(sup_norm(x) for x in args["excl"])
    excl -= set(sup_norm(x) for x in args["incl"])
    info = run_vyvoz(DIRS, shops=args["shops"], use_ua=not args["no_ua"], remove_not_on_ua=args["rm"], excl_suppliers=excl,
                     manual=args["manual"])
    print(_result_text(info))
    return 0 if info.get("ok") else 2


if __name__ == "__main__":
    sys.exit(main())
