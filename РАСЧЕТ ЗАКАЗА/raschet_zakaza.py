# -*- coding: utf-8 -*-
"""
Расчёт заказа по магазинам из оборотной ведомости Торгсофт + валидация.

Формула:
    =IF(L<=0; (H+I)*1,2 - L; IF(H>0; MAX((H+I)*1,2 - L; 0); 0))
    Минус на конец (L<0) = потребность, добавляется к заказу (NEG_AS_NEED).
    H = [-] реализация, I = [-] перемещение, L = на конец
Округление: аналог ОКРУГЛВВЕРХ (штучный товар -> целые, весовой -> шаг 0.1).
Для фасовки/шоубоксов действует порог FIX_MIN_FRAC и минимум min_q из справочника.
Нулевые позиции в заказ не попадают.
"""

import os
import re
import io
import math
import sys
import glob
import time
from datetime import date, datetime, timedelta, timezone

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ========================= НАСТРОЙКИ =========================

def _app_dir():
    """Папка, в которой лежит exe (или сам .py). Абсолютных путей в коде нет:
    программа работает из любой папки и на любом компьютере.
    Можно переопределить переменной окружения FM_BASE_DIR."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except NameError:
        return os.path.dirname(os.path.abspath(sys.argv[0]))


# В окне (console=False) стандартных потоков нет: без этой заглушки
# первый же print уронил бы программу с AttributeError.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = sys.stdout

BASE_DIR = os.environ.get("FM_BASE_DIR", "").strip() or _app_dir()
IN_DIR   = os.path.join(BASE_DIR, "ВХОД")
OUT_DIR  = os.path.join(BASE_DIR, "ЗАКАЗЫ")
LOG_DIR  = os.path.join(BASE_DIR, "ЛОГИ")
ARC_DIR  = os.path.join(BASE_DIR, "АРХИВ")

KOEF = 1.2                 # коэффициент запаса из формулы

# PATCH-NEG-END v1: минусовой остаток на конец = продали то, чего ещё нет на приходе.
# True  -> минус добавляется к заказу: (H+I)*1,2 + |L|
# False -> прежнее поведение: минус трактуется как ноль
NEG_AS_NEED = True

# ROUND_MODE:
#   "up"      - аналог ОКРУГЛВВЕРХ: любой остаток -> вверх (рекомендуется)
#   "half_up" - математическое: от 0.5 вверх, ниже - вниз
#   "smart"   - вверх, но остаток <= SMART_FRAC отбрасывается (гасит хвосты от *1,2)
ROUND_MODE  = "up"
SMART_FRAC  = 0.10
ROUND_EPS   = 1e-6         # гашение погрешности float (1.2000000000000002)

WEIGHT_STEP    = 0.1       # шаг округления весового товара, кг

# ---- справочник упаковок: BigQuery + локальный кэш ----
BQ_PROJECT  = "family-market-analytics"
BQ_TABLE    = "family-market-analytics.family_market.transfer_pack_v5"
BQ_MANUAL   = "family-market-analytics.family_market.transfer_pack_manual"
BQ_RECODE   = "family-market-analytics.family_market.barcode_recode_map"   # старый ШК -> новый (ЮА -> матрица)
REF_FILE    = os.path.join(BASE_DIR, "Справочник", "transfer_pack.csv")  # кэш и офлайн-фолбэк
RECODE_FILE = os.path.join(BASE_DIR, "Справочник", "barcode_recode_map.csv")  # кэш перекодировки ШК
REF_TTL_H   = 12          # свежесть кэша, часов; старше - идём в BQ
REF_OFFLINE = False       # True -> в BQ не ходить вообще, только кэш

_HERE = _app_dir()

CRED_ENV  = "GOOGLE_APPLICATION_CREDENTIALS"
CRED_MASK = "family-market-analytics-*.json"
CRED_DIRS = [
    os.path.join(_HERE, "credentials"),
    os.path.join(BASE_DIR, "credentials"),
    os.path.join(os.path.expanduser("~"), ".gcp"),
]

FIX_MIN_FRAC = 0.34        # меньше 0.34 лотка/шоубокса - не заказывать
MIN_Q_MIN_N  = 5           # min_q применяется только если наблюдений >= 5

# ---- сводный заказ поставщику ----
# Перемещение РЦ -> магазин: коробку можно вскрыть (шаг 0.1 / 1 шт).
# Заказ поставщику: только целыми коробками.
MAKE_SUPPLIER_ORDER = True   # писать сводный файл _ЗАКАЗ_ПОСТАВЩИКУ.xlsx

# ---- раздача конкретного списка с Полевой (режим выбирается в окне) ----
DIST_FILE        = ""                   # путь к списку; пусто = обычный расчёт
DIST_SOURCE_SHOP = u"Полевая-Магазин"   # источник, получателем быть не может
DIST_LIMIT_STOCK = True                 # резать потребность под наличие
DIST_DIR_SUFFIX  = u"_РАЗДАЧА"          # своя папка результата
# Позиции из списка, которых на точках никогда не было: истории нет,
# формула их не видит. Раздаём разнарядкой.
DIST_NEW_MODE = "oborot"   # oborot = пропорционально обороту точки
                           # ravno  = поровну всем точкам
                           # net    = не раздавать, оставить на Полевой
DIST_NEW_TOP  = 0          # 0 = всем; N = только N крупнейшим точкам
DIST_NEW_MIN  = 3        # меньше этого на точку не везём (шт или кг)


# ---- что не перемещаем никогда ----
# Розлив, весовая развеска, кеги и напитки из кофеаппарата - это
# комплектация или производство на точке, между магазинами не ездят.
DIST_SKIP_BC = (
    "4820002713199",   # Оболонь Пиво Сеньор Картель 0,33л з/б - не перемещаем
    "2978950020707",   # Вода Імператорська 1л
)
DIST_SKIP_NAME = (
    u"імператорськ",      # вода на розлив
    u"овочі/фрукти",      # весовая развеска
    u"розливн",           # Оболонь Київське Розливне - наливают на точке
    u" кег",              # БІР  Кег ... - комплектация
    u"кегове",            # Кегове Бір Хаус ...
    u"зерн.кава",         # готовый напиток из кофеаппарата
    u"прем.кава",         # он же, вторая линейка
    u"кришка д/стак",     # тех.товар: крышки к стаканам кофеаппарата
)
# Группы из файла вывоза, которые не раздаём (колонка "Группа").
# Сравнение по подстроке, регистр не важен.
DIST_SKIP_GROUP = (
    u"тех.товар",
)
DIST_SKIP_ONLY_IN_DIST = False

# ---- как делить остаток с Полевой (PATCH-DIST-SHARE v1) ----
# Доля точки = DIST_BC_WEIGHT * продажи ЭТОГО ШК + остальное * общий оборот точки.
# Только оборот (прежнее поведение) = DIST_SHARE_BY_BC = False.
DIST_SHARE_BY_BC = True
DIST_BC_WEIGHT   = 0.7    # 1.0 = только продажи ШК; 0.0 = только оборот точки
DIST_SKIP_DEAD   = True   # не везти туда, где ШК лежит (остаток > 0) и не продаётся


def dist_on():
    return bool(DIST_FILE) and os.path.exists(DIST_FILE)


def dist_suffix():
    return DIST_DIR_SUFFIX if dist_on() else ""
SUPPLIER_MIN_FRAC   = ROUND_EPS   # ROUND_EPS = всегда вверх; 0.34 = экономный вариант
INDIVISIBLE = ("ves_fix", "sht_nedelimyy")   # не дробится даже при перемещении
SUP_MAX_PCS_AUTO = 24    # коробка штучного товара крупнее - только через ручной список
# --- Приходная коробка != шаг перемещения ------------------------------
# Для этих групп число в скобках = транспортная коробка поставщика.
# В магазин перемещаем блоками, а не коробками.
BOX_ONLY_PATTERNS = ("пакет", "пляшка", "бутыл", "запальнич", "зажигалк")
# Слово должно стоять В НАЧАЛЕ названия, иначе ловится "Майонез 170 гр пакет",
# "Оливки пакет 250 г", "Водафон Стартовий Пакет" - это не упаковочный товар.
TRANSFER_BLOCK_RULES = (
    ("пакет", 50),
    ("пляшка", 5),
    ("бутыл", 5),
    ("запальнич", 5),
    ("зажигалк", 5),
)
# Пакеты: всегда кратно 50 и всегда вверх (49 -> 50, 55 -> 100), порог не применяется.
BAG_BLOCK = 50
BAG_ALWAYS_UP = True
# Пакеты, у которых слово стоит не первым (бренд впереди) - ведём списком.
BAG_EXTRA_BC = ()   # пакеты, у которых слово стоит не первым словом названия
# Кузя Сервіс - бытовые мусорные пакеты в рулонах, возятся поштучно: под правило 50 не попадают.

# ---- стаканы: в магазин везём упаковками, а не поштучно ----
# поштучный ШК -> (упаковочный ШК, сколько штук в упаковке)
CUP_PACKS = {
    "2978950015550": ("2978950003595", 50),    # Паперовий Малюнок 175 мл -> Упаковка 50 шт
    "2978950015567": ("2978950003601", 50),    # Паперовий Малюнок 250 мл -> Упаковка 50 шт
    "2978950017660": ("2978950003694", 100),   # Пластик 180 мл -> Упаковка 100 шт
    "2978950015574": ("2978950003717", 50),    # Пластик 300 мл -> Упаковка 50 шт
    "2978950015598": ("2978950003700", 50),    # Пластик 500 мл -> Упаковка 50 шт
    "2978950038856": ("2978950003687", 100),   # Пластик 100 мл -> вместимость уточнить
}
# Стакан Паперовий Малюнок 340 мл (2978950003618) упаковки не имеет - остаётся поштучно.

# ---- пакеты, которые возят только упаковками ----
# Поштучная карточка заменяется упаковочной: закупщица заказывает упаковки,
# а не тысячи штук. Слева - карточка «1 шт», справа - «Упаковка N шт».
BAG_PACKS = {
    "2978950015505": ("2978950004523", 100),   # Пакет Майка 100 шт -> Упаковка 100 шт
    "2978950015482": ("2978950004493", 250),   # Пакет Смайл (250шт) -> Упаковка 250 шт
}

# Общая таблица замен «поштучно -> упаковка»
PACK_SWAP = dict(CUP_PACKS)
PACK_SWAP.update(BAG_PACKS)

# ---- комплектность: к пивным бутылкам едут крышки и ручки ----
BC_KRISHKA = "2978950002598"        # Кришка біла - по одной на каждую бутылку
BC_RUCHKA  = "2978950002604"        # Ручка 1-2 л ПЕТ - только к бутылкам 1; 1,5; 2 л
BC_PLYASHKI_ALL = ("2978950002550", "2978950002567", "2978950017301",
                   "2978950002574", "2978950002581")          # 0,5 / 1 / 1,5 / 2 / 3 л
BC_PLYASHKI_RUCHKA = ("2978950002567", "2978950017301", "2978950002574")  # 1 / 1,5 / 2 л
# Доля блока, ниже которой потребность не тянем на целый блок.
# 0.0 = всегда округлять вверх до блока.
BLOCK_MIN_FILL = 0.5

FIX_MIN_FRAC = 0.34        # меньше 0.34 лотка/шоубокса - не заказывать
MIN_Q_MIN_N  = 5           # min_q применяется только если наблюдений >= 5
MIN_Q_MAX    = 100         # min_q крупнее - мусор из единичного прихода
VES_MIN_Q_N  = 10          # столько перемещений нужно, чтобы верить порции весового
VES_MIN_FILL = 0.34        # мельче этой доли порции весовой товар не возим

# ---- весовой товар в ящиках, которые не вскрывают ----
# Ящик мармелада, зефира, печенья на РЦ не подфасовывают: либо целый ящик,
# либо ничего. Дробное количество по такой позиции закупщица исполнить не может.
VES_BOX_RULE      = True   # включить правило «ящик или ноль»
VES_BOX_MIN       = 1.5    # ящиком считаем вес от этого значения, кг.
                           # Ниже - это порция (конфеты Рошен, ХБФ идут по 1 кг),
                           # её вскрывать можно, правило «ящик или ноль» не для неё.
VES_BOX_STOCK_MAX = 1.0    # остаток не больше этого - везём ящик
VES_BOX_QTY       = 1      # сколько ящиков везём за одну волну
# Реализация 0, а на остатке хвост: по одной выгрузке не отличить "продалось
# вчистую" от "лежит мёртвым". Решаем по величине хвоста: крошка = товар
# кончился, везём ящик; заметный остаток = ждём следующей волны, иначе
# доложенный сверху ящик уйдёт в списание.
VES_BOX_NEED_DEMAND = True   # False -> прежнее поведение (ящик без спроса)
VES_BOX_CRUMB       = 0.3    # хвост не больше этого, кг - считаем, что кончился

# Поставщики, у которых есть подфасовка: ящик вскрывают, возим порциями.
# Ключ - начало названия карточки, значение - порция в кг.
# До Бочкового: скобка в названии - это вес ящика прихода, а не шаг перемещения.
# Проверено по приходам на РЦ (incoming_transactions, delivery_type='РЦ'):
# у этих поставщиков приходы НЕ кратны скобке из названия - значит ящика,
# который нельзя вскрыть, там нет.
VES_PORTION_BY_NAME = {
    u"до бочкового":     1.0,   # есть подфасовка, возим по 1 кг
    u"роганський мк":    0.1,   # весовая колбаса, приходы 1,624 / 2,198 кг
    u"безлюдовський мк": 0.1,   # то же самое
    u"глобино мк":       0.1,   # то же самое
}
REF = {}
WEIGHT_CONFLICTS = []      # справочник говорит «шт», ведомость показывает дробь
PACK_MIN_FRAC = ROUND_EPS  # связки/шоубоксы: всегда вверх (60 при связке 50 -> 100)
MAX_KG_STEP  = 10.0        # фасовки весового товара больше 10 кг не бывает
MAX_BOX_PCS  = 200         # больше - это цена в артикуле, а не вложение ящика
REF = {}
WEIGHT_CONFLICTS = []      # справочник говорит «шт», ведомость показывает дробь
WEIGHT_AS_INT  = False     # True -> весовой тоже округлять до целых
TXT_WEIGHT_SEP = "."       # разделитель дробной части в txt для ТСД
DROP_BELOW     = 0.0       # расчёт ниже этого значения считать шумом (0 = выключено)

HEADER_ROW_HINT = 3
TXT_ENCODING = "cp1251"
TXT_NEWLINE  = "\r\n"

BALANCE_TOL     = 0.011    # допуск сверки баланса ведомости
ANOMALY_FACTOR  = 5        # заказ > factor * реализации -> в отчёт аномалий
ANOMALY_ABS     = 3000     # заказ больше этого числа единиц -> в отчёт.
# Ведомость за месяц: 300 давало сотни ложных срабатываний (яйцо, пакеты, крышки).
MAX_REPORT_ROWS = 300      # сколько проблемных строк выводить в отчёт
NEG_END_TOL     = 0.5      # весовой минус мельче этого - шум взвешивания, не аномалия

BC_STD_LEN = (8, 12, 13, 14)   # стандартные длины EAN/UPC
BC_MANUAL = {                  # ручные соответствия, если контрольная не сходится
    # "54881005906": "054881005906",
}

if NEG_AS_NEED:
    FORMULA_TEXT = "IF(L<=0; (H+I)*1,2-L; IF(H>0; MAX((H+I)*1,2-L; 0); 0))"
    PY_EXPR = "((H + I) * K - L) if (L <= 0) else (max((H + I) * K - L, 0.0) if (H > 0) else 0.0)"
else:
    FORMULA_TEXT = "IF(L<=0; (H+I)*1,2; IF(H>0; MAX((H+I)*1,2-L; 0); 0))"
    PY_EXPR = "((H + I) * K) if (L <= 0) else (max((H + I) * K - L, 0.0) if (H > 0) else 0.0)"

# ======================= ОТЧЁТ ПРОВЕРОК ======================

class Report(object):
    def __init__(self):
        self.checks = []
        self.problems = []
        self.anomalies = []
        self.log = []

    def say(self, msg):
        print(msg)
        self.log.append(str(msg))

    def check(self, status, name, detail=""):
        self.checks.append((status, name, str(detail)))
        mark = {"OK": "[ OK ]", "WARN": "[ВНИМ]", "ERROR": "[ОШИБ]", "INFO": "[ИНФО]"}[status]
        self.say("%s %s%s" % (mark, name, ("  -  " + str(detail)) if detail else ""))

    def problem(self, kind, shop, bc, name, detail):
        if len(self.problems) < MAX_REPORT_ROWS:
            self.problems.append((kind, shop, bc, name, str(detail)))

    def anomaly(self, shop, bc, name, sale, out, end, raw, qty, why):
        self.anomalies.append((shop, bc, name, sale, out, end, round(raw, 3), qty, why))

    @property
    def errors(self):
        return [c for c in self.checks if c[0] == "ERROR"]

    @property
    def warns(self):
        return [c for c in self.checks if c[0] == "WARN"]

R = Report()
RUN_INFO = {}   # итоги последнего прогона для окна

# ========================= УТИЛИТЫ ===========================

def ensure_dirs():
    for d in (BASE_DIR, IN_DIR, OUT_DIR, LOG_DIR, ARC_DIR):
        os.makedirs(d, exist_ok=True)


def norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def safe_name(s):
    s = str(s).strip().replace("/", "-").replace("\\", "-")
    s = re.sub(r'[:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip(" .")
    return s or "БЕЗ_АДРЕСА"


def fmt_barcode(v):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, str):
        s = v.strip().replace(" ", "").replace("\u00a0", "")
        return s[:-2] if s.endswith(".0") else s
    try:
        return str(int(round(float(v))))
    except Exception:
        return str(v).strip()


def ean_valid(bc):
    """Контрольная цифра EAN-8/12/13/14. Для внутренних кодов может не сходиться."""
    if not bc.isdigit() or len(bc) not in (8, 12, 13, 14):
        return False
    digits = [int(ch) for ch in bc]
    body, control = digits[:-1], digits[-1]
    body = body[::-1]
    total = 0
    for idx, d in enumerate(body):
        total += d * (3 if idx % 2 == 0 else 1)
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


def frac_threshold():
    if ROUND_MODE == "up":
        return ROUND_EPS
    if ROUND_MODE == "half_up":
        return 0.5 - 1e-9
    return SMART_FRAC


def apply_rounding(x, step):
    """Округление вверх по шагу с учётом выбранного режима."""
    if x is None or x <= 0:
        return 0.0
    u = round(float(x) / step, 9)
    fl = math.floor(u)
    n = fl if (u - fl) <= frac_threshold() else fl + 1
    return round(n * step, 6)


def calc_order(h, i, l):
    base = (h + i) * KOEF
    if l <= 0:
        return (base - l) if NEG_AS_NEED else base
    return max(base - l, 0.0) if h > 0 else 0.0


def calc_order_eval(h, i, l):
    """Независимая реализация: вычисление текста формулы."""
    return float(eval(PY_EXPR, {"max": max, "__builtins__": {}},
                      {"H": float(h), "I": float(i), "L": float(l), "K": KOEF}))


def qty_to_text(q, is_weight):
    if not is_weight or WEIGHT_AS_INT:
        return str(int(round(q)))
    s = ("%.3f" % float(q)).rstrip("0").rstrip(".")
    return (s or "0").replace(".", TXT_WEIGHT_SEP)

# ===================== САМОТЕСТИРОВАНИЕ ======================

def self_test():
    cases = [((1, 0, 6), 0.0), ((1, 0, 0), 1.2), ((3, 0, 7), 0.0), ((0, 0, 5), 0.0),
             ((5, 0, 2), 4.0), ((2, 3, 1), 5.0), ((0, 4, 0), 4.8), ((10, 2, 3), 11.4)]
    if NEG_AS_NEED:
        cases += [((0, 0, -2), 2.0), ((1, 0, -1), 2.2), ((3, 1, -2.5), 7.3)]
    else:
        cases += [((0, 0, -2), 0.0), ((1, 0, -1), 1.2)]
    bad = []
    for (h, i, l), exp in cases:
        got = calc_order(h, i, l)
        got2 = calc_order_eval(h, i, l)
        if abs(got - exp) > 1e-9 or abs(got2 - exp) > 1e-9:
            bad.append("H=%s I=%s L=%s -> %s (ожидалось %s)" % (h, i, l, got, exp))
    if bad:
        R.check("ERROR", "Тест формулы", "; ".join(bad))
        return False
    R.check("OK", "Тест формулы", "%d контрольных случаев совпали" % len(cases))

    rbad = []
    if ROUND_MODE == "up":
        for x, exp in [(1.2, 2), (2.0, 2), (0.2, 1), (4.8, 5), (5.999, 6), (0.0, 0)]:
            if apply_rounding(x, 1) != exp:
                rbad.append("%s -> %s (ожидалось %s)" % (x, apply_rounding(x, 1), exp))
        for x, exp in [(1.046, 1.1), (0.846, 0.9), (3.0, 3.0)]:
            if abs(apply_rounding(x, 0.1) - exp) > 1e-9:
                rbad.append("вес %s -> %s (ожидалось %s)" % (x, apply_rounding(x, 0.1), exp))
    if apply_rounding(3.0000000001, 1) != 3:
        rbad.append("шум float 3.0000000001 -> лишняя единица")

    # min_q из единичного прихода не должен попадать в заказ
    if _sane_min_q({"min_q": 1000.0, "n": 1}) != 0.0:
        rbad.append("min_q=1000 при n=1 не отброшен")
    if _sane_min_q({"min_q": 6.0, "n": 40}) != 6.0:
        rbad.append("достоверный min_q=6 при n=40 потерян")
    if _sane_min_q({"min_q": 2.7, "n": 1}) != 0.0:
        rbad.append("весовой min_q при n=1 не отброшен")

    # заказ поставщику: 5.1 кг потребности при коробке 2.2 -> 6.6 (3 коробки)
    if abs(apply_rounding_thr(5.1, 2.2, ROUND_EPS) - 6.6) > 1e-9:
        rbad.append("заказ поставщику: 5.1 при коробке 2.2 не поднят до 6.6")
    # перемещение того же товара дробится по 0.1
    if abs(apply_rounding_thr(2.73, 0.1, ROUND_EPS) - 2.8) > 1e-9:
        rbad.append("перемещение: 2.73 не округлено до 2.8")

    # порог фасовки: 0.30 лотка -> 0, 0.40 лотка -> 1 лоток
    if apply_rounding_thr(0.30 * 4.5, 4.5, FIX_MIN_FRAC) != 0.0:
        rbad.append("фасовка: 0.30 лотка не отброшена")
    if abs(apply_rounding_thr(0.40 * 4.5, 4.5, FIX_MIN_FRAC) - 4.5) > 1e-9:
        rbad.append("фасовка: 0.40 лотка не поднята до 1 упаковки")
    if rbad:
        R.check("ERROR", "Тест округления", "; ".join(rbad))
        return False
    R.check("OK", "Тест округления",
            "режим '%s', шаг веса %s, порог фасовки %s"
            % (ROUND_MODE, WEIGHT_STEP, FIX_MIN_FRAC))

    pcases = [
        ("АВК Ваф.Трубочки Бам-Бук Згущене Молоко 1кг (4,4)", "4,4", True, None, 4.4),
        ("АВК Цукерки Труфальє 1кг", "6", True, None, 6.0),
        ("Щедрики Печиво 4,5кг", "", True, 4.5, None),
        ("Печиво Вівсяне 2,4кг (9,6)", "9,6", True, 2.4, 9.6),
        ("ДО Бочкового Розлив 1кг", "1000", True, None, None),
        ("АВК Kresko Банановий Смак 140г", "24", False, None, 24.0),
        ("Якобс Монарх 3в1 13г (20)", "20", False, 20.0, 20.0),
        ("МакКофе Оригінал Стик 18г", "26", False, 26.0, 26.0),
        ("Нескафе Класик 250г банка", "6", False, None, 6.0),
        ("Агропром Соломка Для Коктейля (*100) 1 шт", "", False, 1.0, None),
        ("Агропром Стакан Паперовий Малюнок 175 мл Упаковка 50 шт", "", False, 1.0, None),
        ("Агропром Комплект №1 (42)", "", False, None, 42.0),
    ]
    pbad = []
    for nm, ar, wt, exp_s, exp_b in pcases:
        g = parse_pack_from_name(nm, ar, wt)
        if g["step"] != exp_s or g["box"] != exp_b:
            pbad.append("%s -> шаг %s / ящик %s (ожидалось %s / %s)"
                        % (nm[:34], g["step"], g["box"], exp_s, exp_b))
    if pbad:
        R.check("ERROR", "Тест парсера названий", "; ".join(pbad))
        return False
    R.check("OK", "Тест парсера названий",
            "%d случаев: кг вне скобок = шаг, скобка/артикул = ящик, шоубокс кофе"
            % len(pcases))

    if apply_rounding_thr(60, 50, PACK_MIN_FRAC) != 100:
        R.check("ERROR", "Тест связок", "60 при связке 50 не дало 100")
        return False
    R.check("OK", "Тест связок", "60 -> 100, 101 -> 150 (всегда вверх)")

    # База ЮА: те же магазины под именем с «ЮА», одна точка переименована
    ubad = []
    for nm, exp in [(u"Качанівська 19 ЮА", (u"Качанівська 19", True)),
                    (u"качанівська 19 юа", (u"качанівська 19", True)),
                    (u"ЮА Качанівська 19", (u"Качанівська 19", True)),
                    (u"Качанівська 19 ЮA", (u"Качанівська 19", True)),   # латинская A
                    (u"Болградська 38 ЮА", (u"Грозненська 38", True)),
                    (u"Зерновая 6/5 ЮА", (u"Зернова 6/5", True)),
                    (u"Качанівська 19", (u"Качанівська 19", False)),
                    (u"Іскрінський 19", (u"Іскрінський 19", False)),
                    (u"Болградська 38", (u"Болградська 38", False))]:
        if ua_split(nm) != exp:
            ubad.append("%s -> %s (ожидалось %s)" % (nm, ua_split(nm), exp))
    ua2, fm2 = [u"Качанівська 19 ЮА", u"Болградська 38 ЮА"], [u"Байрона 156", u"Зернова 6/5"]
    for shops, exp in [(ua2, "ua"), (fm2, "fm"), (ua2 + fm2, "mixed")]:
        if file_base(shops) != exp:
            ubad.append("база файла %s -> %s (ожидалось %s)" % (shops, file_base(shops), exp))

    tue = date(2026, 10, 6)                      # вторник, куст B
    full_b = list(ROUTES["B"])

    def _rc(shops):
        use, chk = route_compare(tue, shops)
        return use, [c[0] for c in chk], " | ".join(c[2] for c in chk)

    use, st, det = _rc(full_b)
    if use != "B" or "WARN" in st:
        ubad.append("Family, весь куст B: %s %s" % (use, det[:120]))
    # все восемь точек, которые знаем на ЮА (адреса даны владельцем 09.10)
    ua_all = [u"Байрона 138/1 ЮА", u"Байрона 156 ЮА", u"Зерновая 6/5 ЮА", u"Ньютона 111 ЮА",
              u"Байрона 163 ЮА", u"Ньютона 102 ЮА", u"Качанівська 19 ЮА", u"Болградська 38 ЮА"]
    use, st, det = _rc(ua_all)
    if use != "B" or "WARN" in st or u"в файле 8 из 13" not in det:
        ubad.append("файл ЮА со всеми восемью точками: %s %s" % (use, det[:160]))
    moved = [u"Байрона 138/1", u"Байрона 156", u"Зернова 6/5", u"Ньютона 111", u"Байрона 163",
             u"Ньютона 102", u"Качанівська 19", u"Грозненська 38"]
    use, st, det = _rc([s for s in full_b if s not in moved])
    if "WARN" in st:
        ubad.append("Family без перешедших на ЮА точек дал предупреждение: %s" % det[:120])
    if u"все 5 магазинов на месте" not in det:
        ubad.append("перешедшие на ЮА точки посчитаны как «на месте»: %s" % det[:160])
    use, st, det = _rc([s for s in full_b if s != u"Танкопія 16"])
    if "WARN" not in st or u"Танкопія 16" not in det:
        ubad.append("пропажа обычной точки Family перестала замечаться: %s" % det[:120])
    use, st, det = _rc(ua2)
    if use != "B" or "WARN" in st:
        ubad.append("файл ЮА с двумя точками куста B: %s %s" % (use, det[:120]))
    use, st, det = _rc([u"Качанівська 19 ЮА", u"Бучми 32 ЮА"])
    if "WARN" not in st or u"Бучми 32" not in det:
        ubad.append("точка чужого куста в файле ЮА не замечена: %s" % det[:120])

    for shops, exp in [(ua2, "2026-10-06_ЮА"), (fm2, "2026-10-06"), (ua2 + fm2, "2026-10-06")]:
        if out_day_name(tue, shops) != exp:
            ubad.append("папка результата %s -> %s (ожидалось %s)"
                        % (shops, out_day_name(tue, shops), exp))
    # справочник упаковок для ЮА: ШК ЮА берёт запись своего ШК в матрице (barcode_recode_map)
    ref_t = {u"4820017000130": {"unit_type": "sht_nedelimyy", "step": 6.0}}
    pairs_t = {u"4820182062568": u"4820017000130",    # ЮА -> матрица, цель в справочнике
               u"4823127313923": u"4820111111111",    # цели в справочнике нет
               u"4820017000130": u"4820182062568"}    # свой ШК уже в справочнике: не трогаем
    added = apply_recode_to_ref(ref_t, pairs_t)
    if added != 1:
        ubad.append("перекодировка справочника добавила %s записей (ожидалась 1)" % added)
    if ref_t.get(u"4820182062568", {}).get("step") != 6.0:
        ubad.append("ШК ЮА не получил упаковку по перекодировке")
    if u"4823127313923" in ref_t:
        ubad.append("перекодировка на ШК без упаковки создала пустую запись")
    if ref_t[u"4820017000130"].get("step") != 6.0 or \
            ref_t[u"4820182062568"] is ref_t[u"4820017000130"]:
        ubad.append("перекодировка испортила прямую запись или дала общий объект")
    if ubad:
        R.check("ERROR", "Тест базы ЮА", "; ".join(ubad))
        return False
    R.check("OK", "Тест базы ЮА",
            "суффикс ЮА, Болградська=Грозненська, 8 точек, куст, отдельная папка, перекодировка")
    return True

# ==================== ЧТЕНИЕ ИСТОЧНИКА =======================

ALIASES = {
    "name":   ["название товара", "назва товару", "наименование"],
    "art":    ["артикул"],
    "bc":     ["штрих-код", "штрих код", "штрихкод"],
    "start":  ["на начало", "на початок"],
    "inc":    ["приход", "прихід"],
    "move_in": ["[+] перемещение", "[+] переміщення"],
    "sale":   ["[-] реализация", "[-] реалізація", "реализация", "реалізація"],
    "out":    ["[-] перемещение", "[-] переміщення"],
    "woff":   ["списание", "списання"],
    "ret":    ["возврат поставщикам", "повернення постачальникам"],
    "end":    ["на конец", "на кінець"],
    "shop":   ["центр учета", "центр обліку", "магазин"],
}


def find_source_file(quiet=False):
    """Автопоиск: приоритет - дата в названии, затем дата изменения.
    Используется как запасной путь, когда окно выбора файла недоступно,
    и как подсказка по умолчанию для проводника (quiet=True)."""
    found = []
    for folder in (IN_DIR, BASE_DIR):
        if not os.path.isdir(folder):
            continue
        for nm in os.listdir(folder):
            low = nm.lower()
            if nm.startswith("~$") or not low.endswith((".xlsx", ".xlsm")):
                continue
            if not any(k in low for k in ("перемещ", "переміщ", "оборот", "ведом")):
                continue
            full = os.path.join(folder, nm)
            d = extract_date(full)
            found.append((d, os.path.getmtime(full), full))
    if not found:
        return None

    found.sort(key=lambda t: (t[0], t[1]), reverse=True)
    if len(found) > 1 and not quiet:
        R.check("WARN", "В папке несколько ведомостей",
                "найдено %d, взят файл за %s: %s | остальные: %s"
                % (len(found), found[0][0].strftime("%d.%m.%Y"),
                   os.path.basename(found[0][2]),
                   ", ".join(os.path.basename(f[2]) for f in found[1:4])))
    return found[0][2]


def extract_date(path):
    """Дата заказа. Тонкое место: она берётся из имени файла и ни на один
    расчёт не влияет - только на имя папки в ЗАКАЗЫ и подписи в отчётах.
    Ошибиться здесь = записать заказ не в ту папку, а не посчитать не то."""
    return guess_day(path)[0]


def guess_day(path):
    """(дата, откуда взята). Порядок: имя файла -> отметка выгрузки -> сегодня."""
    m = re.search(r"(\d{1,2})[.\-_ ](\d{1,2})[.\-_ ](\d{2,4})", os.path.basename(path))
    if m:
        d, mo, y = (int(g) for g in m.groups())
        y = y + 2000 if y < 100 else y
        try:
            return date(y, mo, d), "из имени файла"
        except ValueError:
            pass
    st = read_export_stamp(path)
    if st:
        return st.date(), "из отметки выгрузки: в имени файла даты нет"
    return date.today(), "сегодняшняя: в имени файла даты нет"


# ==================== МАРШРУТЫ ОТГРУЗКИ ======================
#
# Три куста, каждый отгружается дважды в неделю. Состав снят с фактических
# ведомостей 01-05 и 07.09.2026 (вт, ср, чт, пт, сб, пн) - это ровно одна
# полная неделя, поэтому адреса здесь в том виде, в каком их пишет Торгсофт,
# а не в старом русском написании из графика.
#   A - пн и чт   B - вт и пт   C - ср и сб    Воскресенье отгрузки нет.
# Зачем: если куст в файле не совпал с днём недели - выгрузили не тот период
# или ошиблись датой; если внутри куста магазина нет - он выпал из выгрузки.

ROUTES = {
    "A": [u"Амосова 5А", u"Богдана Хмельницького 8", u"Валентинівська 50А",
          u"Гарібальді 1", u"Зубенко 31В/5", u"Краснодарська 171з",
          u"Переяславська 23", u"Пр-т Героїв Харкова 160", u"Пр-т Ювілейний 67",
          u"Роганська 130/4", u"Роганська 148", u"Салтівське шосе 264В"],
    "B": [u"Іскрінський 19", u"Байрона 138/1", u"Байрона 156", u"Байрона 163",
          u"Грозненська 38", u"Зернова 6/5", u"Качанівська 19",
          u"Небесної Сотні 14/1", u"Ньютона 102", u"Ньютона 111",
          u"Олімпійська 9А", u"Петра Григоренка 37", u"Танкопія 16"],
    "C": [u"Астрономічна 44Г", u"Бучми 32", u"Бучми 52", u"Бучми Джерело",
          u"Валентинівська 24", u"Гвардійців Широнінців 54", u"Зубенко 23",
          u"Михайля Семенка 17", u"Нескорених 33",
          u"Пр-т Тракторобудiвникiв 95", u"Шевченко 341"],
}
ROUTE_BY_DOW = {0: "A", 1: "B", 2: "C", 3: "A", 4: "B", 5: "C"}   # пн..сб

# Точки, отсутствие которых в выгрузке - норма, а не пропажа.
# Убрать из ROUTES совсем, когда закроются окончательно.
# Точки, перешедшие на ЮА Маркет (адреса даны владельцем 09.10.2026): в ведомости
# Family их нет, считаются файлом базы ЮА. Новый переход = строка здесь.
ROUTE_PAUSED = {u"Байрона 138/1": u"переведена на ЮА, считается в файле базы ЮА",
                u"Байрона 156": u"переведена на ЮА, считается в файле базы ЮА",
                u"Байрона 163": u"переведена на ЮА, считается в файле базы ЮА",
                u"Зернова 6/5": u"переведена на ЮА, считается в файле базы ЮА",
                u"Ньютона 102": u"переведена на ЮА, считается в файле базы ЮА",
                u"Ньютона 111": u"переведена на ЮА, считается в файле базы ЮА",
                u"Качанівська 19": u"переведена на ЮА, считается в файле базы ЮА",
                u"Грозненська 38": u"переведена на ЮА (там Болградська 38 ЮА)"}

# Улицы переименованы, в старых выгрузках и графике встречаются прежние
# названия - это те же самые магазины.
ROUTE_ALIASES = {
    u"Героїв Праці 33":         u"Нескорених 33",
    u"Героев труда 33":         u"Нескорених 33",
    u"Героїв Сталінграда 138/1": u"Байрона 138/1",
    u"Героев Сталинграда 138":   u"Байрона 138/1",
    u"Героїв Сталінграда 156":   u"Байрона 156",
    u"Героев Сталинграда 156":   u"Байрона 156",
    u"Героїв Сталінграда 163":   u"Байрона 163",
    u"Героев Сталинграда 163":   u"Байрона 163",
}
DOW_RU = [u"понедельник", u"вторник", u"среда", u"четверг",
          u"пятница", u"суббота", u"воскресенье"]

# латиница, набранная вместо кириллицы, встречается в адресах Торгсофта
_LAT2CYR = {"a": u"а", "c": u"с", "e": u"е", "i": u"і", "o": u"о",
            "p": u"р", "x": u"х", "y": u"у"}


def route_key(s):
    """Ключ сравнения адреса: регистр, пробелы, дроби и латинские двойники
    не должны мешать сопоставлению."""
    t = re.sub(r"\s+", " ", str(s).strip().lower())
    t = "".join(_LAT2CYR.get(ch, ch) for ch in t)
    return re.sub(r"[/\\.-]", "", t)


# База ЮА - те же магазины в другой базе Торгсофта: к имени добавлено «ЮА»,
# одна точка при переходе переименована. route_key, ROUTE_ALIASES и safe_name
# здесь не трогаем: ими пользуется vyvoz_vne_matricy.py, а там «Качанівська 19»
# и «Качанівська 19 ЮА» - разные точки.
UA_RENAMES = {u"Болградська 38": u"Грозненська 38",    # имя в базе ЮА -> имя в графике
              u"Зерновая 6/5": u"Зернова 6/5"}
UA_DIR_SUFFIX = u"_ЮА"    # заказы ЮА - отдельная папка: в Торгсофте это другая программа
_UA_TOKEN_RE = re.compile(r"^\s*ю[аa]\s+|\s+ю[аa]\s*$", re.IGNORECASE)   # a - и латинская


def ua_split(name):
    """(имя точки для графика, взята ли из базы ЮА). Слово «ЮА» в начале или в конце
    отбрасывается, переименованная точка возвращается под именем из графика."""
    s = re.sub(r"\s+", " ", str(name).strip())
    base = _UA_TOKEN_RE.sub("", s, count=1).strip()
    if base == s:
        return s, False
    key = route_key(base)
    for ua_nm, fm_nm in UA_RENAMES.items():
        if route_key(ua_nm) == key:
            return fm_nm, True
    return base, True


def file_base(shops_in_file):
    """"ua" - все точки из базы ЮА, "fm" - ни одной, "mixed" - вперемешку."""
    flags = {ua_split(s)[1] for s in shops_in_file}
    if flags == {True}:
        return "ua"
    return "mixed" if True in flags else "fm"


def out_day_name(day, shops_in_file):
    """Имя папки результата в ЗАКАЗЫ. Файл базы ЮА - своя папка, чтобы ТСД-файлы
    Family и ЮА не смешались и расчёт одной базы не убрал в замену расчёт другой."""
    ua = UA_DIR_SUFFIX if file_base(shops_in_file) == "ua" else ""
    return day.strftime("%Y-%m-%d") + ua + dist_suffix()


def check_route(day, shops_in_file):
    """Сверка состава магазинов с графиком отгрузки."""
    if dist_on():
        R.check("INFO", u"График отгрузки",
                u"раздача с %s: куст не проверяем, точек в файле %d"
                % (DIST_SOURCE_SHOP, len(set(shops_in_file))))
        return ROUTE_BY_DOW.get(day.weekday()) or "A"
    use, checks = route_compare(day, shops_in_file)
    for st, nm, det in checks:
        R.check(st, nm, det)
    return use


def route_compare(day, shops_in_file):
    """(куст, [(статус, имя, детали)]). Отчёт не пишет: check_route печатает,
    самотест читает. В файле базы ЮА нехватка точек куста не замечание: остальные
    точки куста живут в базе Family."""
    out = []
    ck = lambda st, nm, det: out.append((st, nm, det))
    alias = {route_key(k): route_key(v) for k, v in ROUTE_ALIASES.items()}
    have = {}
    for x in shops_in_file:
        k = route_key(ua_split(x)[0])
        have[alias.get(k, k)] = x
    scores = {k: len(have.keys() & {route_key(a) for a in v})
              for k, v in ROUTES.items()}
    best = max(scores, key=lambda k: scores[k])
    exp = ROUTE_BY_DOW.get(day.weekday())

    if exp is None:
        ck("WARN", u"График отгрузки",
           u"%s - отгрузки нет, в файле похоже на куст %s"
           % (DOW_RU[day.weekday()], best))
        use = best
    elif best != exp and scores[best] > scores.get(exp, 0):
        ck("WARN", u"График отгрузки",
           u"дата %s это %s (куст %s), а состав магазинов похож на куст %s"
           u" - проверьте дату или период выгрузки"
           % (day.strftime("%d.%m.%Y"), DOW_RU[day.weekday()], exp, best))
        use = best
    else:
        use = exp

    want = ROUTES[use]
    miss = [a for a in want if route_key(a) not in have]
    extra = [have[k] for k in have
             if k not in {route_key(a) for a in want}]

    if file_base(shops_in_file) == "ua":
        ck("INFO", u"График отгрузки",
           u"куст %s (%s), база ЮА: в файле %d из %d точек куста"
           % (use, DOW_RU[day.weekday()], len(want) - len(miss), len(want)))
    else:
        paused = [a for a in miss if a in ROUTE_PAUSED]
        miss = [a for a in miss if a not in ROUTE_PAUSED]
        for a in paused:
            ck("INFO", u"Точка вне отгрузки", u"%s - %s" % (a, ROUTE_PAUSED[a]))
        if miss:
            ck("WARN", u"Магазина нет в выгрузке",
               u"куст %s, %s, не хватает %d из %d: %s"
               % (use, DOW_RU[day.weekday()], len(miss), len(want),
                  "; ".join(miss)))
        else:
            ck("OK", u"График отгрузки",
               u"куст %s (%s), все %d магазинов на месте"
               % (use, DOW_RU[day.weekday()], len(want) - len(paused)))
    if extra:
        ck("WARN", u"Магазин вне графика этого дня",
           u"%s - в кусте %s его быть не должно" % ("; ".join(extra), use))
    return use, out


# ============ ВЫБОР ФАЙЛА, ДАТЫ И КОНТРОЛЬ ПОВТОРОВ ============
#
# Задача блока - убрать три тихие ошибки:
#   1. посчитали не тот файл (старую выгрузку, вчерашнюю ведомость);
#   2. дата угадана из имени и заказ лёг не в ту папку;
#   3. повторный прогон подмешал свежие файлы к прошлым в той же папке.
# Ни одна из них не портит числа - все три портят то, что уходит в ТСД.

USE_GUI            = True   # False -> без окон, как раньше (автопоиск)
STALE_EXPORT_HOURS = 18     # выгрузка старше -> предупреждение
DATE_SANITY_DAYS   = 10     # дата дальше от сегодня -> предупреждение
HISTORY_FILE       = "история_запусков.csv"
REPLACED_DIR       = "_замененные"


def read_export_stamp(path):
    """Когда выгрузка сделана в Торгсофте. Берётся из служебных свойств xlsx
    (docProps/core.xml), а не из даты файла: копирование и пересохранение
    дату файла меняют, а эту отметку - нет. Возвращает datetime или None."""
    try:
        import zipfile
        with zipfile.ZipFile(path) as z:
            raw = z.read("docProps/core.xml").decode("utf-8", "ignore")
        m = (re.search(r"<dcterms:created[^>]*>([^<]+)<", raw)
             or re.search(r"<dcterms:modified[^>]*>([^<]+)<", raw))
        if not m:
            return None
        t = m.group(1).strip()
        dt = datetime.strptime(t[:19], "%Y-%m-%dT%H:%M:%S")
        if t.endswith("Z") or "+" in t[10:]:
            dt = dt.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)
        return dt
    except Exception:
        return None


def source_fingerprint(path):
    """Отпечаток содержимого: размер + sha1. Один и тот же файл под другим
    именем или после копирования даёт тот же отпечаток."""
    try:
        import hashlib
        h = hashlib.sha1()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return "%d_%s" % (os.path.getsize(path), h.hexdigest()[:16])
    except Exception:
        return ""


def history_read():
    """Список прошлых запусков из ЛОГИ \\ история_запусков.csv."""
    rows, p = [], os.path.join(LOG_DIR, HISTORY_FILE)
    if not os.path.isfile(p):
        return rows
    try:
        with io.open(p, encoding="utf-8") as f:
            for line in f.read().splitlines()[1:]:
                parts = line.split(";")
                if len(parts) >= 5:
                    rows.append(dict(zip(("run", "day", "file", "fp", "shops"),
                                         parts[:5])))
    except Exception:
        pass
    return rows


def history_write(day, src, fp, shops):
    p = os.path.join(LOG_DIR, HISTORY_FILE)
    new = not os.path.isfile(p)
    try:
        with io.open(p, "a", encoding="utf-8", newline="") as f:
            if new:
                f.write(u"запуск;дата заказа;файл;отпечаток;магазинов\n")
            f.write(u"%s;%s;%s;%s;%d\n"
                    % (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       day.strftime("%Y-%m-%d"),
                       os.path.basename(src).replace(";", ","), fp, shops))
    except Exception:
        pass


def history_same_file(fp):
    """Последний запуск с тем же отпечатком или None."""
    if not fp:
        return None
    same = [h for h in history_read() if h.get("fp") == fp]
    return same[-1] if same else None


def day_dir_state(day):
    """(путь, магазинов внутри, время последней записи) для папки этой даты."""
    d = os.path.join(OUT_DIR, day.strftime("%Y-%m-%d"))
    if not os.path.isdir(d):
        return d, 0, None
    try:
        shops = [n for n in os.listdir(d)
                 if os.path.isdir(os.path.join(d, n)) and not n.startswith("_")]
        ts = datetime.fromtimestamp(os.path.getmtime(d))
    except Exception:
        shops, ts = [], None
    return d, len(shops), ts


def locked_files(day_dir):
    """Файлы в папке даты, открытые сейчас в Excel. Ловим их до записи,
    а не на середине: иначе половина магазинов запишется, половина нет."""
    busy = []
    if not os.path.isdir(day_dir):
        return busy
    for root_, _dirs, files in os.walk(day_dir):
        for nm in files:
            if nm.startswith("~$"):
                busy.append(os.path.join(root_, nm[2:]))
                continue
            if not nm.lower().endswith((".xlsx", ".txt")):
                continue
            p = os.path.join(root_, nm)
            try:
                with open(p, "r+b"):
                    pass
            except Exception:
                busy.append(p)
    return sorted(set(busy))


def clear_day_dir(day_dir):
    r"""Прошлый результат этой даты уводится в ЗАКАЗЫ\_замененные\<дата>_<время>.
    Не удаляется: вернуть можно всегда. Смысл - чтобы в папке даты не осталось
    txt магазина, который в новом расчёте заказа не получил."""
    import shutil
    if not os.path.isdir(day_dir) or not os.listdir(day_dir):
        return None
    dst = os.path.join(OUT_DIR, REPLACED_DIR,
                       "%s_%s" % (os.path.basename(day_dir),
                                  datetime.now().strftime("%H-%M-%S")))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.move(day_dir, dst)
    return dst


def archive_source(src):
    """Ведомость после успешного расчёта уезжает в АРХИВ - чтобы в ВХОД
    всегда лежал ровно один файл и выбирать было не из чего."""
    import shutil
    try:
        dst = os.path.join(ARC_DIR, os.path.basename(src))
        k = 2
        while os.path.exists(dst):
            stem, ext = os.path.splitext(os.path.basename(src))
            dst = os.path.join(ARC_DIR, "%s (%d)%s" % (stem, k, ext))
            k += 1
        shutil.move(src, dst)
        return dst
    except Exception as e:
        R.check("WARN", "Перенос ведомости в АРХИВ",
                "не удался (%s), файл остался в ВХОД" % type(e).__name__)
        return None


def _gui_pick(default_src):
    """Проводник + окно подтверждения. Возвращает кортеж настроек,
    None - если отменили, False - если окна недоступны (тогда автопоиск)."""
    import tkinter as tk
    from tkinter import filedialog, ttk

    root = tk.Tk()
    root.withdraw()
    try:
        root.call("tk", "scaling", 1.3)
    except Exception:
        pass
    # без этого диалог открывается ЗА окном консоли и выглядит как зависание
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass

    start_dir = IN_DIR if os.path.isdir(IN_DIR) else BASE_DIR
    src = filedialog.askopenfilename(
        parent=root,
        title="Выберите оборотную ведомость из Торгсофта",
        initialdir=(os.path.dirname(default_src)
                    if default_src else start_dir),
        initialfile=(os.path.basename(default_src) if default_src else ""),
        filetypes=[("Книги Excel", "*.xlsx *.xlsm"), ("Все файлы", "*.*")])
    if not src:
        root.destroy()
        return None
    src = os.path.abspath(src)

    guess, origin = guess_day(src)
    stamp = read_export_stamp(src)
    prev  = history_same_file(source_fingerprint(src))

    state = {"ok": False}

    BG, CARD, FG   = "#f4f6f9", "#ffffff", "#16202c"
    MUTED, BORD    = "#6b7885", "#dfe4ea"
    ACC, ACC_H     = "#2f6fed", "#1f5bd0"
    RED, GRN, AMB  = "#c62828", "#1b7a3d", "#a86400"
    F, FB, FS      = ("Segoe UI", 10), ("Segoe UI Semibold", 10), ("Segoe UI", 9)

    win = tk.Toplevel(root)
    win.title(u"Расчёт заказа перемещений")
    win.configure(bg=BG)
    win.resizable(False, False)

    st = ttk.Style(win)
    try:
        st.theme_use("clam")
    except Exception:
        pass
    st.configure(".", background=CARD, foreground=FG, font=F)
    st.configure("Card.TFrame", background=CARD)
    st.configure("Cap.TLabel", background=CARD, foreground=MUTED, font=FS)
    st.configure("Val.TLabel", background=CARD, foreground=FG,
                 font=("Segoe UI Semibold", 11))
    st.configure("Note.TLabel", background=CARD, foreground=MUTED, font=FS)
    st.configure("Red.TLabel",  background=CARD, foreground=RED, font=FS)
    st.configure("Grn.TLabel",  background=CARD, foreground=GRN, font=FS)
    st.configure("Amb.TLabel",  background=CARD, foreground=AMB, font=FS)
    st.configure("Go.TButton", background=ACC, foreground="#ffffff",
                 font=FB, borderwidth=0, focuscolor=ACC, padding=(20, 9))
    st.map("Go.TButton", background=[("active", ACC_H), ("pressed", ACC_H)])
    st.configure("Ghost.TButton", background=CARD, foreground=MUTED,
                 font=F, borderwidth=0, focuscolor=CARD, padding=(16, 9))
    st.map("Ghost.TButton", background=[("active", "#e9edf3")])
    st.configure("TCheckbutton", background=CARD, foreground=FG, font=FS,
                 focuscolor=CARD)
    st.map("TCheckbutton", background=[("active", CARD)])
    st.configure("D.TEntry", fieldbackground="#ffffff", bordercolor=BORD,
                 lightcolor=BORD, darkcolor=BORD, padding=7)

    # шапка
    head = tk.Frame(win, bg=ACC)
    head.pack(fill="x")
    tk.Label(head, text=u"Расчёт заказа перемещений", bg=ACC, fg="#ffffff",
             font=("Segoe UI Semibold", 14)).pack(anchor="w", padx=22, pady=(16, 2))
    tk.Label(head, text=u"РЦ → магазины · проверьте перед запуском", bg=ACC,
             fg="#cfe0ff", font=FS).pack(anchor="w", padx=22, pady=(0, 16))

    body = tk.Frame(win, bg=CARD)
    body.pack(fill="both", expand=True)
    pad = dict(padx=22, anchor="w")

    def cap(t):
        tk.Label(body, text=t, bg=CARD, fg=MUTED, font=FS).pack(pady=(14, 1), **pad)

    def line(t, style="Note.TLabel"):
        w = ttk.Label(body, text=t, style=style, justify="left")
        w.pack(pady=(2, 0), **pad)
        return w

    def rule():
        tk.Frame(body, bg=BORD, height=1).pack(fill="x", padx=22, pady=(14, 0))

    cap(u"ВЕДОМОСТЬ")
    tk.Label(body, text=os.path.basename(src), bg=CARD, fg=FG,
             font=("Segoe UI Semibold", 11)).pack(**pad)
    line(os.path.dirname(src))

    if stamp:
        age = (datetime.now() - stamp).total_seconds() / 3600.0
        line(u"Выгружено из Торгсофта %s · %.0f ч назад"
             % (stamp.strftime("%d.%m.%Y в %H:%M"), age),
             "Amb.TLabel" if age > STALE_EXPORT_HOURS else "Grn.TLabel")
    else:
        line(u"Отметки времени выгрузки в файле нет")
    if prev:
        line(u"Этот же файл уже считался %s → ЗАКАЗЫ\\%s"
             % (prev["run"], prev["day"]), "Red.TLabel")

    rule()
    cap(u"ДАТА ЗАКАЗА")
    drow = tk.Frame(body, bg=CARD)
    drow.pack(padx=22, anchor="w")
    dv = tk.StringVar(value=guess.strftime("%d.%m.%Y"))
    ttk.Entry(drow, textvariable=dv, width=11, justify="center",
              style="D.TEntry", font=("Segoe UI Semibold", 12)).pack(side="left")
    tk.Label(drow, text="  " + origin, bg=CARD, fg=MUTED,
             font=FS).pack(side="left")
    route_l = line(u"")
    warn_l  = line(u"", "Red.TLabel")

    rule()
    box = tk.Frame(body, bg=CARD)
    box.pack(fill="x", padx=18, pady=(10, 0))
    clean_v, arch_v = tk.IntVar(value=1), tk.IntVar(value=1)
    chk = ttk.Checkbutton(box, variable=clean_v,
                          text=u"прошлый результат этой даты убрать в ЗАКАЗЫ\\%s"
                               % REPLACED_DIR)
    chk.pack(anchor="w", pady=1)
    ttk.Checkbutton(box, variable=arch_v,
                    text=u"после расчёта убрать ведомость в АРХИВ").pack(
                        anchor="w", pady=1)

    def parse_dv():
        t = dv.get().strip().replace("/", ".").replace("-", ".")
        m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", t)
        if not m:
            return None
        d, mo, y = (int(g) for g in m.groups())
        y = y + 2000 if y < 100 else y
        try:
            return date(y, mo, d)
        except ValueError:
            return None

    def refresh(*_a):
        d = parse_dv()
        if d is None:
            route_l.config(text=u"")
            warn_l.config(text=u"Дата не разобрана, нужен вид ДД.ММ.ГГГГ")
            chk.pack_forget()
            return
        rk = ROUTE_BY_DOW.get(d.weekday())
        if rk:
            route_l.config(text=u"%s · куст %s · ожидается %d магазинов"
                                % (DOW_RU[d.weekday()].capitalize(), rk,
                                   len(ROUTES[rk])), style="Note.TLabel")
        else:
            route_l.config(text=u"Воскресенье — отгрузки по графику нет",
                           style="Amb.TLabel")
        msgs = []
        _dir, n, ts = day_dir_state(d)
        if n:
            msgs.append(u"Папка ЗАКАЗЫ\\%s уже есть: %d магазинов, записана %s"
                        % (d.strftime("%Y-%m-%d"), n,
                           ts.strftime("%d.%m в %H:%M") if ts else "?"))
            chk.pack(anchor="w", pady=1, before=box.winfo_children()[-1])
        else:
            chk.pack_forget()
        off = (d - date.today()).days
        if abs(off) > DATE_SANITY_DAYS:
            msgs.append(u"Дата отстоит от сегодняшней на %d дн." % off)
        warn_l.config(text="\n".join(msgs))

    dv.trace_add("write", refresh)

    def on_ok(*_a):
        if parse_dv() is None:
            refresh()
            return
        state["ok"] = True
        win.destroy()

    def on_cancel(*_a):
        state["ok"] = False
        win.destroy()

    foot = tk.Frame(win, bg=BG)
    foot.pack(fill="x")
    ttk.Button(foot, text=u"Рассчитать", style="Go.TButton",
               command=on_ok).pack(side="right", padx=(0, 22), pady=16)
    ttk.Button(foot, text=u"Отмена", style="Ghost.TButton",
               command=on_cancel).pack(side="right", padx=8, pady=16)
    for w_ in foot.winfo_children():
        w_.configure(takefocus=1)
    tk.Frame(foot, bg=BG).pack(side="left")

    refresh()
    win.bind("<Return>", on_ok)
    win.bind("<Escape>", on_cancel)
    win.protocol("WM_DELETE_WINDOW", on_cancel)
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry("%dx%d+%d+%d"
                 % (max(520, win.winfo_reqwidth()), win.winfo_reqheight(),
                    max(0, (sw - max(520, win.winfo_reqwidth())) // 2),
                    max(0, (sh - win.winfo_reqheight()) // 3)))
    win.grab_set()
    win.lift()
    win.focus_force()
    try:
        win.attributes("-topmost", True)
        win.after(400, lambda: win.attributes("-topmost", False))
    except Exception:
        pass
    root.wait_window(win)

    day = parse_dv() if state["ok"] else None
    clean = bool(clean_v.get())
    arch = bool(arch_v.get())
    root.destroy()
    if not state["ok"] or day is None:
        return None
    return src, day, clean, arch, ("указана вручную"
                                   if day != guess else origin)


GUI_CHOICE = None   # заполняется окном перед запуском расчёта


def pick_source():
    """(файл, дата, чистить_папку, архивировать, откуда_дата) или None.
    Порядок: --auto -> автопоиск без окон; путь аргументом -> он;
    иначе проводник. Любая осечка окон = откат на прежнее поведение."""
    if GUI_CHOICE:
        return GUI_CHOICE
    args = [a for a in sys.argv[1:] if a.strip()]
    auto = any(a.lower().lstrip("-/") in ("auto", "a") for a in args)
    paths = [a for a in args if not a.startswith(("-", "/"))]
    given = os.path.abspath(paths[0]) if paths and os.path.isfile(paths[0]) else None

    if not auto and USE_GUI:
        res = False
        try:
            res = _gui_pick(given or find_source_file(quiet=True))
        except Exception as e:
            R.check("WARN", "Окно выбора файла",
                    "не открылось (%s: %s), взят автопоиск"
                    % (type(e).__name__, e))
        if res is None:
            return None
        if res:
            return res

    src = given or find_source_file()
    if not src:
        return None
    day, origin = guess_day(src)
    return src, day, True, False, origin



def detect_header_row(path):
    probe = pd.read_excel(path, sheet_name=0, header=None, nrows=15)
    for idx in range(len(probe)):
        vals = [norm(v) for v in probe.iloc[idx].tolist() if pd.notna(v)]
        if any("штрих" in v for v in vals) and any("центр" in v or "магазин" in v for v in vals):
            return idx + 1
    return HEADER_ROW_HINT


def resolve_columns(df):
    cols = {norm(c): c for c in df.columns}
    res = {}
    for key, variants in ALIASES.items():
        hit = None
        for v in variants:
            if v in cols:
                hit = cols[v]
                break
        if hit is None:
            for v in variants:
                for cn, orig in cols.items():
                    if v in cn:
                        hit = orig
                        break
                if hit:
                    break
        res[key] = hit
    return res

# ======================= ПОДГОТОВКА ==========================

def prepare(df, c):
    num = lambda col: pd.to_numeric(df[col], errors="coerce")

    raw_sale, raw_out, raw_end = num(c["sale"]), num(c["out"]), num(c["end"])
    bad_cells = int(raw_sale.isna().sum() + raw_out.isna().sum() + raw_end.isna().sum())
    if bad_cells:
        R.check("WARN", "Нечисловые значения в расчётных столбцах",
                "%d ячеек трактованы как 0" % bad_cells)
    else:
        R.check("OK", "Числовые данные", "нечисловых ячеек в H/I/L нет")

    w = pd.DataFrame({
        "shop": df[c["shop"]].astype(str).str.strip(),
        "name": (df[c["name"]].astype(str).str.strip() if c.get("name") else ""),
        "art":  (df[c["art"]].astype(str).str.strip() if c.get("art") else ""),
        "bc":   df[c["bc"]].map(fmt_barcode),
        "sale": raw_sale.fillna(0.0),
        "out":  raw_out.fillna(0.0),
        "end":  raw_end.fillna(0.0),
        "start": (num(c["start"]).fillna(0.0) if c.get("start") else 0.0),
    })
    w["name"] = w["name"].replace({"nan": "", "None": ""})
    w["art"] = w["art"].replace({"nan": "", "None": ""})
    _fixed, _pairs = [], []
    for b in w["bc"]:
        nb, note = restore_barcode(b)
        _fixed.append(nb)
        if note:
            _pairs.append(note)
    w["bc"] = _fixed
    if _pairs:
        uniq_pairs = sorted(set(_pairs))
        R.check("WARN", "Восстановлены ведущие нули штрих-кодов",
                "%d кодов: %s" % (len(uniq_pairs), "; ".join(uniq_pairs[:6])))
        for p in uniq_pairs:
            R.problem("Ведущий ноль восстановлен (Excel обрезал)", "", p, "",
                      "проверь выгрузку Торгсофта")
    else:
        R.check("OK", "Ведущие нули штрих-кодов", "потерь не обнаружено")

    w["src_row"] = df.index + 1

    dropped = w[(w["bc"] == "") | (w["shop"] == "") | (w["shop"].str.lower() == "nan")]
    for _, r0 in dropped.iterrows():
        R.problem("Пропущена строка (нет штрих-кода или адреса)", r0["shop"], r0["bc"],
                  r0["name"], "строка источника ~%s" % r0["src_row"])
    if len(dropped):
        R.check("WARN", "Строки без штрих-кода/адреса", "исключено %d" % len(dropped))
    else:
        R.check("OK", "Полнота ключевых полей", "штрих-код и адрес есть во всех строках")

    w = w.drop(dropped.index)
    return w


def balance_check(df, c):
    need = ("start", "inc", "move_in", "sale", "out", "woff", "ret", "end")
    if any(not c.get(k) for k in need):
        R.check("WARN", "Сверка баланса ведомости", "нет части столбцов, проверка пропущена")
        return
    g = lambda k: pd.to_numeric(df[c[k]], errors="coerce").fillna(0.0)
    calc = g("start") + g("inc") + g("move_in") - g("sale") - g("out") - g("woff") - g("ret")
    diff = (calc - g("end")).abs()
    bad = int((diff > BALANCE_TOL).sum())
    if bad:
        R.check("WARN", "Сверка баланса (начало+приход+перемещ-расход = конец)",
                "расхождений: %d строк, макс. %.3f" % (bad, diff.max()))
    else:
        R.check("OK", "Сверка баланса ведомости", "все %d строк сходятся" % len(df))


def detect_weight(w):
    frac = lambda s: (s - s.round()).abs() > 0.001
    flag = frac(w["sale"]) | frac(w["out"]) | frac(w["end"]) | frac(w["start"])
    by_bc = flag.groupby(w["bc"]).any()
    kg_hint = (w.assign(h=w["name"].str.contains("кг", case=False, na=False))
                .groupby("bc")["h"].any())
    art_hint = (w.assign(h=w["art"].str.contains(r"[.,]", na=False))
                 .groupby("bc")["h"].any()) & kg_hint
    weight = (by_bc | art_hint).to_dict()
    R.check("INFO", "Определение весового товара",
            "весовых артикулов: %d из %d" % (sum(1 for v in weight.values() if v), len(weight)))
    return weight


def barcode_check(w):
    uniq = sorted(set(w["bc"]))
    nondigit = [b for b in uniq if not b.isdigit()]
    badlen = [b for b in uniq if b.isdigit() and len(b) not in (8, 12, 13, 14)]
    badsum = [b for b in uniq if b.isdigit() and len(b) in (8, 12, 13, 14) and not ean_valid(b)]
    if nondigit:
        R.check("ERROR", "Штрих-коды: посторонние символы",
                "%d кодов, напр.: %s" % (len(nondigit), ", ".join(nondigit[:5])))
    else:
        R.check("OK", "Штрих-коды: только цифры", "проверено %d кодов" % len(uniq))
    if badlen:
        R.check("WARN", "Штрих-коды: нетипичная длина",
                "%d кодов, напр.: %s" % (len(badlen), ", ".join(badlen[:5])))
    if badsum:
        R.check("WARN", "Штрих-коды: не сходится контрольная цифра",
                "%d кодов (норма для внутренних), напр.: %s" % (len(badsum), ", ".join(badsum[:5])))

    mix = w.groupby("bc")["name"].nunique()
    mix = mix[mix > 1]
    if len(mix):
        R.check("WARN", "Один штрих-код у разных названий", "%d кодов" % len(mix))
        for b in list(mix.index)[:20]:
            names = sorted(set(w.loc[w["bc"] == b, "name"]))
            R.problem("Один код - разные товары", "", b, " | ".join(names[:3]), "")
    return not bool(nondigit)

# ========================== РАСЧЁТ ===========================
# ================== КРАТНОСТЬ ИЗ НАЗВАНИЯ ====================

_MASK_PATTERNS = [
    r"\d+\s*(?:в|b|in)\s*1\b",                    # 3в1, 2в1, 3in1
    r"\d+(?:[.,]\d+)?\s*(?:мл|мм|см|%)",          # 175 мл, 250 мл
    r"№\s*\d+",                                   # Комплект №1
    r"\d+\s*сорт",                                # 1 сорт
]
_RE_STAR = re.compile(r"\(\s*[*xх]\s*(\d+)\s*\)", re.I)
_RE_UPAK = re.compile(r"(?:упаковка|уп\.?)\s*(\d+)\s*шт", re.I)
_RE_BOX  = re.compile(r"\(\s*(\d+(?:[.,]\d+)?)\s*\)")
_RE_KG   = re.compile(r"(\d+(?:[.,]\d+)?)\s*кг\b", re.I)
_RE_G    = re.compile(r"(\d+(?:[.,]\d+)?)\s*г\b", re.I)
_RE_STICK = re.compile(r"стик|\d\s*(?:в|in)\s*1", re.I)
_RE_COFFEE = re.compile(
    r"якобс|jacobs|мак\s*ко|maccoffee|mccoffee|петровськ|петровск|"
    r"нескафе|nescafe|карт\s*нуар|carte\s*noire|чорна\s*карта|галка|жокей", re.I)

STEP_FROM_NAME = []


def _fnum(s):
    try:
        return float(str(s).replace("\u00a0", "").replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return None


def _mask_name(s):
    out = s
    for p in _MASK_PATTERNS:
        out = re.sub(p, lambda m: "\x00" * len(m.group(0)), out, flags=re.I)
    return out


def _no_brackets(s):
    return re.sub(r"\([^)]*\)", lambda m: "\x00" * len(m.group(0)), s)


def _small_gram(masked):
    mg = _RE_G.search(masked)
    v = _fnum(mg.group(1)) if mg else None
    return bool(v and v <= 30)


def guess_box(masked, artikul, is_weight):
    """Вложение ящика. Только для справки, в расчёт не идёт."""
    lim = MAX_KG_STEP if is_weight else MAX_BOX_PCS
    vals = [v for v in (_fnum(x) for x in _RE_BOX.findall(masked)) if v]
    for v in reversed(vals):
        if 0 < v <= lim:
            return v, "название"
    v = _fnum(artikul)
    if v and 0 < v <= lim:
        return v, "артикул"
    return None, ""


def parse_pack_from_name(name, artikul="", is_weight=False):
    """-> dict(step, thr, box, kind, note). step=None -> решает эвристика."""
    s = re.sub(r"\s+", " ", str(name or "").replace("\u00a0", " ")).strip()
    m = _mask_name(s)
    box, box_src = guess_box(m, artikul, is_weight)
    res = {"step": None, "thr": None, "box": box, "kind": "", "note": ""}

    if is_weight:
        mk = _RE_KG.search(_no_brackets(m))
        v = _fnum(mk.group(1)) if mk else None
        if v and abs(v - 1.0) < 1e-9:
            res["note"] = "цена за 1 кг, вес плавающий"
            return res
        if v and 0.1 <= v <= MAX_KG_STEP:
            res.update(step=v, thr=FIX_MIN_FRAC, kind="ves_fix",
                       note="фасовка %g кг из названия" % v)
        return res

    # штучный товар: стиковый кофе - только целым шоубоксом
    if _RE_STICK.search(s) or (_RE_COFFEE.search(s) and _small_gram(m)):
        if box and float(box).is_integer() and 2 <= box <= 100:
            res.update(step=float(int(box)), thr=PACK_MIN_FRAC, kind="sht_nedelimyy",
                       note="шоубокс %d шт (%s)" % (int(box), box_src))
            return res
    mu = _RE_UPAK.search(m)
    if mu:
        res.update(step=1.0, thr=frac_threshold(), kind="pack_unit",
                   note="SKU = упаковка %s шт, шаг 1" % mu.group(1))
        return res
    ms = _RE_STAR.search(m)
    if ms:
        res.update(step=1.0, thr=frac_threshold(), kind="piece",
                   note="вложение %s шт, шаг 1" % ms.group(1))
    return res



def find_credentials():
    """Ключ сервис-аккаунта: env -> папки-кандидаты. -> путь или None"""
    p = os.environ.get(CRED_ENV, "").strip().strip('"')
    if p and os.path.exists(p):
        return p
    for d in CRED_DIRS:
        if not d or not os.path.isdir(d):
            continue
        hits = sorted(glob.glob(os.path.join(d, CRED_MASK)))
        if not hits:
            hits = sorted(glob.glob(os.path.join(d, "*.json")))
        if hits:
            os.environ[CRED_ENV] = hits[0]
            return hits[0]
    return None


_MQ_PAREN = re.compile(r"\((\d{1,3})\)")


def _mq_in_name(mq, pname):
    """Подтверждён ли min_q числом в скобках названия: '... (25) 18г' + min_q=25."""
    if not pname:
        return False
    want = int(round(float(mq)))
    return any(int(g) == want for g in _MQ_PAREN.findall(str(pname)))


def _min_q_ok(info):
    """min_q достоверен только при статистике, вменяемой величине и подтверждении."""
    if is_box_only(info.get("pname", "")):
        return False  # (50)/(200)/(120) - приход на склад, не шаг перемещения
    mq = info.get("min_q")
    if not mq or float(mq) <= 0:
        return False
    if int(info.get("n") or 0) < MIN_Q_MIN_N:
        return False
    if float(mq) > MIN_Q_MAX:
        return False
    if str(info.get("confidence", "")) == "manual":
        return True
    ut = str(info.get("unit_type", ""))
    if ut == "ves_plav":
        return int(info.get("n") or 0) >= VES_MIN_Q_N   # порция подтверждена историей
    if ut.startswith("ves"):
        return False          # ves_fix: шаг и есть упаковка, min_q не нужен
    if ut == "sht_delimyy":
        return _mq_in_name(mq, info.get("pname", ""))
    return True


def _sane_min_q(info):
    return float(info["min_q"]) if _min_q_ok(info) else 0.0


def refresh_reference_from_bq():
    """Тянет справочник в REF_FILE. -> (успех, сообщение)"""
    if REF_OFFLINE:
        return False, "включён режим офлайн (REF_OFFLINE=True)"
    cred = find_credentials()
    if not cred:
        return False, ("не найден ключ: задай %s или положи json в %s"
                       % (CRED_ENV, os.path.join(_HERE, "credentials")))
    try:
        from google.cloud import bigquery
    except ImportError:
        return False, "нет пакета: pip install google-cloud-bigquery db-dtypes"
    sql = """
    SELECT p.barcode, p.pname,
           COALESCE(m.unit_type,  p.unit_type)  AS unit_type,
           COALESCE(m.order_step, p.order_step) AS order_step,
           COALESCE(m.min_q,      p.min_q)      AS min_q,
           COALESCE(m.step_supplier, p.step_supplier, p.order_step) AS step_supplier,
           p.box_kg, p.n, p.review,
           IF(m.barcode IS NULL, IFNULL(p.confidence, ''), 'manual') AS confidence
    FROM `%s` p
    LEFT JOIN `%s` m USING (barcode)
    """ % (BQ_TABLE, BQ_MANUAL)
    try:
        client = bigquery.Client(project=BQ_PROJECT)
        d = client.query(sql).to_dataframe()
    except Exception as e:
        return False, "%s: %s" % (type(e).__name__, str(e).replace("\n", " ")[:200])
    if d is None or d.empty:
        return False, "запрос вернул 0 строк"
    d["barcode"] = d["barcode"].astype(str).str.strip()
    os.makedirs(os.path.dirname(REF_FILE), exist_ok=True)
    d.to_csv(REF_FILE, index=False, encoding="utf-8-sig")
    return True, "%d позиций, ключ %s" % (len(d), os.path.basename(cred))


def load_reference():
    """Справочник упаковок: BigQuery с кэшем в REF_FILE."""
    fresh = (os.path.exists(REF_FILE)
             and (time.time() - os.path.getmtime(REF_FILE)) < REF_TTL_H * 3600)
    if fresh:
        R.check("INFO", "Справочник: источник",
                "кэш свежее %d ч, запрос в BigQuery не нужен" % REF_TTL_H)
    else:
        ok, msg = refresh_reference_from_bq()
        if ok:
            R.check("OK", "Справочник: обновлён из BigQuery", msg)
        elif os.path.exists(REF_FILE):
            age = (time.time() - os.path.getmtime(REF_FILE)) / 3600.0
            R.check("WARN", "Справочник: BigQuery недоступен",
                    "%s -> работаю по кэшу возрастом %.1f ч" % (msg, age))
        else:
            R.check("WARN", "Справочник: BigQuery недоступен",
                    "%s; кэша нет -> работа по эвристике" % msg)

    if not os.path.exists(REF_FILE):
        return {}

    try:
        d = pd.read_csv(REF_FILE, sep=None, engine="python",
                        encoding="utf-8-sig", dtype=str, keep_default_na=False)
    except Exception as e:
        R.check("ERROR", "Справочник упаковок",
                "не удалось прочитать %s: %s" % (REF_FILE, e))
        return {}

    d.columns = [str(cn).strip().lstrip("\ufeff").lower() for cn in d.columns]
    lack = [k for k in ("barcode", "unit_type", "order_step") if k not in d.columns]
    if lack:
        R.check("ERROR", "Справочник упаковок",
                "нет столбцов %s; в файле: %s" % (lack, list(d.columns)))
        return {}

    def _num(v):
        s = str(v).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
        if s == "" or s.lower() in ("nan", "none", "null"):
            return None
        try:
            x = float(s)
        except ValueError:
            return None
        return x if x > 0 else None

    out, broken, dups = {}, [], 0
    for _, r0 in d.iterrows():
        bc = fmt_barcode(r0["barcode"])
        if not bc:
            continue
        bc, _ = restore_barcode(bc)
        st = _num(r0["order_step"])
        if st is None:
            broken.append("%s '%s' -> '%s'" % (bc, str(r0.get("pname", ""))[:26],
                                               r0["order_step"]))
            continue
        if bc in out:
            dups += 1
        out[bc] = {
            "unit_type": str(r0["unit_type"]).strip(),
            "step": st,
            "min_q": _num(r0["min_q"]) if "min_q" in d.columns else None,
            "n": int(_num(r0.get("n", 0)) or 0),
            "pname": str(r0.get("pname", "")).strip(),
            "step_sup": _num(r0.get("step_supplier", "")) or st,
            "box_kg": _num(r0.get("box_kg", "")),
            "review": str(r0.get("review", "")).strip(),
            "confidence": str(r0.get("confidence", "")).strip(),
        }

    if broken:
        R.check("WARN", "Справочник: нечисловой шаг",
                "%d строк пропущено (похоже на порчу Excel - даты вместо чисел), напр.: %s"
                % (len(broken), "; ".join(broken[:3])))
    if dups:
        R.check("WARN", "Справочник: дубли штрих-кодов",
                "%d повторов, взята последняя строка" % dups)

    cnt = lambda t: sum(1 for v in out.values() if v["unit_type"] == t)
    mq_use = sum(1 for v in out.values() if _min_q_ok(v))
    mq_bad = [v for v in out.values() if v.get("min_q") and not _min_q_ok(v)]
    R.check("OK", "Справочник упаковок",
            "%d позиций: недел. %d, фасовка %d, вес %d; min_q применяется у %d"
            % (len(out), cnt("sht_nedelimyy"), cnt("ves_fix"), cnt("ves_plav"), mq_use))
    cand = [v for v in out.values()
            if v.get("min_q") and not _min_q_ok(v)
            and int(v.get("n") or 0) >= 20
            and v.get("unit_type") == "sht_delimyy"
            and float(v["min_q"]) >= 6]
    if cand:
        cand.sort(key=lambda v: -float(v["min_q"]))
        R.check("WARN", "Похоже на шоубокс, но не подтверждено названием",
                "%d позиций - проверь и внеси в transfer_pack_manual: %s"
                % (len(cand), "; ".join("%s min_q=%g n=%s"
                                        % (v.get("pname", "")[:30], v["min_q"], v.get("n"))
                                        for v in cand[:5])))
        for v in cand[:40]:
            R.problem("Возможный шоубокс без подтверждения", "", "", v.get("pname", ""),
                      "min_q=%g при n=%s; если это блок - задать order_step вручную"
                      % (v["min_q"], v.get("n")))

    if mq_bad:
        R.check("WARN", "Справочник: min_q отброшен как недостоверный",
                "%d позиций (наблюдений < %d или значение > %g), напр.: %s"
                % (len(mq_bad), MIN_Q_MIN_N, MIN_Q_MAX,
                   "; ".join("%s min_q=%g n=%s" % (v.get("pname", "")[:22],
                                                   v["min_q"], v.get("n"))
                             for v in mq_bad[:3])))
    rv = {}
    for v in out.values():
        r_ = v.get("review", "")
        if r_ and r_ != "OK":
            rv[r_] = rv.get(r_, 0) + 1
    if rv:
        R.check("INFO", "Справочник: позиции на проверку",
                ", ".join("%s: %d" % kv for kv in sorted(rv.items())))
    return out


def refresh_recode_from_bq():
    """Тянет перекодировку ШК в RECODE_FILE. -> (успех, сообщение)"""
    if REF_OFFLINE:
        return False, "включён режим офлайн (REF_OFFLINE=True)"
    cred = find_credentials()
    if not cred:
        return False, ("не найден ключ: задай %s или положи json в %s"
                       % (CRED_ENV, os.path.join(_HERE, "credentials")))
    try:
        from google.cloud import bigquery
    except ImportError:
        return False, "нет пакета: pip install google-cloud-bigquery db-dtypes"
    try:
        client = bigquery.Client(project=BQ_PROJECT)
        d = client.query("SELECT old_barcode, new_barcode FROM `%s`" % BQ_RECODE).to_dataframe()
    except Exception as e:
        return False, "%s: %s" % (type(e).__name__, str(e).replace("\n", " ")[:200])
    if d is None or d.empty:
        return False, "запрос вернул 0 строк"
    os.makedirs(os.path.dirname(RECODE_FILE), exist_ok=True)
    d.astype(str).to_csv(RECODE_FILE, index=False, encoding="utf-8-sig")
    return True, "%d пар, ключ %s" % (len(d), os.path.basename(cred))


def load_recode():
    """Перекодировка ШК (старый -> новый; ШК ЮА -> ШК матрицы): BigQuery с кэшем в
    RECODE_FILE. Нужна только файлам базы ЮА. -> {старый ШК: новый ШК}"""
    fresh = (os.path.exists(RECODE_FILE)
             and (time.time() - os.path.getmtime(RECODE_FILE)) < REF_TTL_H * 3600)
    if fresh:
        R.check("INFO", "Перекодировка ШК: источник",
                "кэш свежее %d ч, запрос в BigQuery не нужен" % REF_TTL_H)
    else:
        ok, msg = refresh_recode_from_bq()
        if ok:
            R.check("OK", "Перекодировка ШК: обновлена из BigQuery", msg)
        elif os.path.exists(RECODE_FILE):
            age = (time.time() - os.path.getmtime(RECODE_FILE)) / 3600.0
            R.check("WARN", "Перекодировка ШК: BigQuery недоступен",
                    "%s -> работаю по кэшу возрастом %.1f ч" % (msg, age))
        else:
            R.check("WARN", "Перекодировка ШК: BigQuery недоступен",
                    "%s; кэша нет -> упаковки ЮА берутся только по их собственным ШК" % msg)
    if not os.path.exists(RECODE_FILE):
        return {}
    try:
        d = pd.read_csv(RECODE_FILE, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        pairs = list(zip(d["old_barcode"], d["new_barcode"]))
    except Exception as e:
        R.check("WARN", "Перекодировка ШК", "не удалось прочитать %s: %s" % (RECODE_FILE, e))
        return {}
    out = {}
    for old, new in pairs:
        o, n = (restore_barcode(fmt_barcode(x))[0] for x in (old, new))
        if o and n and o != n:
            out[o] = n
    return out


def apply_recode_to_ref(ref, pairs):
    """ШК ЮА, которого нет в справочнике, получает копию записи своего нового ШК
    (ШК матрицы). Собственная запись сильнее перекодировки. -> сколько записей добавлено.
    Только в памяти: txt и xlsx остаются с ШК ЮА, его знает программа ЮА на терминале."""
    added = 0
    for old, new in pairs.items():
        if old not in ref and new in ref:
            ref[old] = dict(ref[new])
            added += 1
    return added

def _norm_name(s: str) -> str:
    s = (s or "").lower()
    for a, b in (("і", "и"), ("ї", "и"), ("є", "е"), ("ы", "и")):
        s = s.replace(a, b)
    return s


def is_box_only(name: str) -> bool:
    """True, если скобки в названии = коробка прихода, а не шаг перемещения."""
    n = _norm_name(name)
    return any(p in n for p in BOX_ONLY_PATTERNS)


def is_pack_card(name: str) -> bool:
    """Карточка, единица которой сама по себе упаковка: «... Упаковка 100 шт».
    Кратность 50 к такой карточке применять нельзя - это было бы 50 упаковок."""
    return bool(re.search(r"упаковка\s*\d+\s*шт", _norm_name(name)))


def is_bag(name: str, bc: str = "") -> bool:
    """Упаковочный пакет, считаемый штуками. Карточка-упаковка сюда не входит."""
    if is_pack_card(name):
        return False
    if str(bc) in BAG_EXTRA_BC:
        return True
    return bool(re.match(r"\s*пакет\b", _norm_name(name)))


def transfer_block(name: str, bc: str = "") -> int:
    if is_bag(name, bc):
        return BAG_BLOCK
    n = _norm_name(name)
    for pat, blk in TRANSFER_BLOCK_RULES:
        if pat == "пакет":
            continue                      # пакеты разобраны выше, по началу названия
        if re.match(r"\s*" + pat, n):     # слово должно быть в начале названия
            return blk
    return 0


def box_from_name(name: str) -> float:
    """Вес ящика весового товара, если он написан в названии:
    «(4кг)», «(1,5кг в ящ)», «1кг (5)», «1 кг(0,8)». Ноль - не нашли.
    Названию верим больше, чем шагу из истории: история показывает,
    сколько отгружали по факту, а ящик физически не делится."""
    n = _norm_name(name).strip()
    for pat in (r"\((\d+[.,]?\d*)\s*кг[^)]*\)",
                r"кг\s*\((\d+[.,]?\d*)\)",
                r"\((\d+[.,]?\d*)\)\s*$"):
        m = re.search(pat, n)
        if m:
            try:
                v = float(m.group(1).replace(",", "."))
            except ValueError:
                continue
            if 0.2 <= v <= MAX_KG_STEP:
                return v
    return 0.0


def resolve_step(bc, heur_weight, name="", art=""):
    """Приоритет: справочник -> кратность из названия -> эвристика.
    -> (шаг, порог дробной части, минимальный заказ, пояснение)"""
    info = REF.get(bc)

    _nm = name or (info or {}).get("pname", "")
    _blk = transfer_block(_nm, bc)
    if _blk:
        _tag = "пакет" if is_bag(_nm, bc) else "блок"
        return float(_blk), ROUND_EPS, float(_blk), \
               "%s: кратно %d шт (коробка - только поставщику)" % (_tag, _blk)

    if info and info.get("step"):
        s = float(info["step"])
        ut = str(info.get("unit_type", ""))
        mq = _sane_min_q(info)                 # min_q при n < MIN_Q_MIN_N отбрасывается
        tail = (", min %g" % mq) if mq > 0 else ""
        rv = str(info.get("review", ""))
        if rv and rv != "OK":
            tail += " [%s]" % rv

        if ut == "sht_delimyy" and heur_weight and not WEIGHT_AS_INT:
            WEIGHT_CONFLICTS.append((bc, name, s))
            return WEIGHT_STEP, frac_threshold(), 0.0, \
                   "вес по ведомости (в справочнике 'шт')"
        if ut == "sht_nedelimyy":
            return s, FIX_MIN_FRAC, mq, "шоубокс %g шт%s" % (s, tail)
        if ut == "ves_fix":
            return s, FIX_MIN_FRAC, mq, "фасовка %g кг%s" % (s, tail)
        if ut == "ves_plav":
            return (s if s > 0 else WEIGHT_STEP), frac_threshold(), mq, \
                   "вес, шаг %g кг%s" % (s, tail)
        return s, frac_threshold(), mq, ((ut or "не указано") + tail)

    g = parse_pack_from_name(name, art, heur_weight)
    if g and g.get("step"):
        STEP_FROM_NAME.append((bc, name, float(g["step"]), g.get("box"), g.get("note", "")))
        return float(g["step"]), float(g["thr"]), 0.0, g.get("note", "из названия")

    s = WEIGHT_STEP if (heur_weight and not WEIGHT_AS_INT) else 1.0
    return s, frac_threshold(), 0.0, "эвристика: " + ("вес" if heur_weight else "шт")

def resolve_supplier_step(bc, transfer_step, unit_type=""):
    """Шаг для заказа поставщику: только целыми коробками.
    -> (шаг, пояснение)"""
    info = REF.get(bc)
    if not info:
        return float(transfer_step), "нет в справочнике"
    ut = str(info.get("unit_type", "") or unit_type)
    sup = float(info.get("step_sup") or info.get("step") or transfer_step)
    if ut in INDIVISIBLE:
        return sup, "неделимо, %g" % sup
    manual = str(info.get("confidence", "")) == "manual"
    if ut == "sht_delimyy" and sup > SUP_MAX_PCS_AUTO and not manual:
        return float(transfer_step), "кратность %g отклонена (транспортная коробка)" % sup
    if sup > float(transfer_step) + 1e-9:
        return sup, "коробка %g (перемещение по %g)" % (sup, transfer_step)
    return sup, "коробка неизвестна, по %g" % sup


def _dist_from_lines(lines):
    """Строки вида 'ШК' или 'ШК;количество' -> {ШК: сумма}."""
    tot = {}
    for ln in lines:
        p = [x.strip() for x in re.split(r"[;\t,]", ln) if x.strip()]
        if not p:
            continue
        q = _fnum(p[1]) if len(p) > 1 else None
        tot[p[0]] = tot.get(p[0], 0.0) + (q or 0.0)
    return tot


def _dist_read_txt(path):
    for enc in ("cp1251", "utf-8-sig"):
        try:
            with io.open(path, encoding=enc) as f:
                return f.read().splitlines()
        except (UnicodeDecodeError, IOError):
            continue
    return []


def _dist_from_txt_dir(folder):
    """Папка ВЫВОЗ: складываем TXT всех складов - это и приехало на Полевую."""
    tot, files = {}, 0
    for nm in sorted(os.listdir(folder)):
        if not nm.lower().endswith(".txt") or nm.upper().startswith(u"МАТРИЦА"):
            continue
        files += 1
        for bc, q in _dist_from_lines(_dist_read_txt(os.path.join(folder, nm))).items():
            tot[bc] = tot.get(bc, 0.0) + q
    return tot, files


def _dist_skip_row(bc, name, group):
    """Причина не перемещать строку вывоза, либо '' если можно."""
    if bc and bc in set(DIST_SKIP_BC):
        return u"ШК в списке исключений"
    low = (u"%s" % (name or u"")).lower()
    for pat in DIST_SKIP_NAME:
        if pat in low:
            return u"название: %s" % pat
    grp = (u"%s" % (group or u"")).strip().lower()
    for pat in DIST_SKIP_GROUP:
        if pat in grp:
            return u"группа: %s" % pat
    return u""


def _dist_pick_col(cols, *keys):
    """Колонка по подстроке имени: заголовки в файле вывоза плавают."""
    for c in cols:
        n = norm(c)
        for k in keys:
            if k in n:
                return c
    return None


DIST_NAMES = {}   # ШК -> название из файла вывоза (в ведомости его нет)


def _dist_skip_row(bc, name, group):
    """Причина не перемещать строку вывоза, либо '' если можно."""
    if bc and bc in set(DIST_SKIP_BC):
        return u"ШК в списке исключений"
    low = (u"%s" % (name or u"")).lower()
    for pat in DIST_SKIP_NAME:
        if pat in low:
            return u"название: %s" % pat
    grp = (u"%s" % (group or u"")).strip().lower()
    for pat in DIST_SKIP_GROUP:
        if pat in grp:
            return u"группа: %s" % pat
    return u""


def _dist_pick_col(cols, *keys):
    """Колонка по подстроке имени: заголовки в файле вывоза плавают."""
    for c in cols:
        n = norm(c)
        for k in keys:
            if k in n:
                return c
    return None


def _dist_from_xlsx(path):
    """Вывоз_<дата>.xlsx: строки с «Вывоз = ДА» после правок закупщицы.
    Розлив, кеги, кофе из аппарата и тех.товар отсекаются здесь же:
    позже нечем - у позиций без истории на точках названия в ведомости
    нет, и фильтр по названию в drop_excluded их не видит."""
    tot, sheets, cut = {}, 0, {}
    for sh in (u"Вне матрицы", u"Матрица на точках"):
        try:
            d = pd.read_excel(path, sheet_name=sh, dtype={u"Штрих-код": str})
        except (ValueError, IOError, KeyError):
            continue
        if u"Штрих-код" not in d.columns:
            continue
        if u"Вывоз" not in d.columns:
            R.check("WARN", u"Лист вывоза без колонки «Вывоз»",
                    u"«%s»: отметки закупщицы на этом листе не прочитаны" % sh)
            continue
        sheets += 1
        d = d[d[u"Вывоз"].astype(str).str.strip().str.upper() == u"ДА"]
        qs = pd.to_numeric(d[u"Количество"], errors="coerce").fillna(0)
        c_nm = _dist_pick_col(d.columns, u"назва", u"название", u"наименов")
        c_gr = _dist_pick_col(d.columns, u"групп")
        nms = d[c_nm].astype(str) if c_nm else pd.Series(u"", index=d.index)
        grs = d[c_gr].astype(str) if c_gr else pd.Series(u"", index=d.index)
        for bc, q, nm, gr in zip(d[u"Штрих-код"], qs, nms, grs):
            b = fmt_barcode(bc)
            nm = u"" if nm in (u"nan", u"None") else nm
            why = _dist_skip_row(b, nm, gr)
            if why:
                k = (b, nm, why)
                cut[k] = cut.get(k, 0.0) + float(q)
                continue
            if nm:
                DIST_NAMES[b] = nm
            tot[b] = tot.get(b, 0.0) + float(q)
    if cut:
        R.check("INFO", u"Из вывоза исключено до раздачи",
                u"%d ШК, %g ед. - розлив/кеги/кофе/тех.товар"
                % (len(cut), round(sum(cut.values()), 3)))
        for (b, nm, why), q in sorted(cut.items(), key=lambda t: -t[1])[:40]:
            R.problem(u"Не перемещаем (из файла вывоза)", "", b, nm,
                      u"%s, %g ед." % (why, round(q, 3)))
    return tot, sheets


def load_dist_list(path, quiet=False):
    """Ассортимент для раздачи. Принимается:
       * TXT/CSV - строки 'ШК' или 'ШК;количество';
       * папка ВЫВОЗ_<дата> (или её TXT) - сумма файлов всех складов;
       * Вывоз_<дата>.xlsx - строки с «Вывоз = ДА».
       -> (множество ШК, {ШК: сколько лежит на Полевой})"""
    raw, src = {}, os.path.basename(path.rstrip("\\/"))
    try:
        if os.path.isdir(path):
            sub = os.path.join(path, "TXT")
            folder = sub if os.path.isdir(sub) else path
            raw, n = _dist_from_txt_dir(folder)
            src = u"%s, файлов складов: %d" % (os.path.basename(folder), n)
        elif path.lower().endswith((".xlsx", ".xlsm")):
            raw, n = _dist_from_xlsx(path)
            src = u"%s, листов: %d" % (src, n)
        else:
            raw = _dist_from_lines(_dist_read_txt(path))
    except (IOError, OSError) as e:
        if not quiet:
            R.check("ERROR", u"Список раздачи", u"%s: %s" % (src, e))
        return set(), {}

    bcs, avail = set(), {}
    for bc0, q in raw.items():
        bc, _n = restore_barcode(fmt_barcode(bc0))
        if not bc.isdigit():
            continue
        bcs.add(bc)
        if q:
            avail[bc] = round(avail.get(bc, 0.0) + float(q), 6)
    if not quiet:
        R.check("OK" if bcs else "ERROR", u"Список раздачи",
                u"%s -> %d ШК, из них с количеством %d"
                % (src, len(bcs), len(avail)))
    return bcs, avail


def dist_count(path):
    """Сколько позиций в выбранном источнике - для подписи в окне."""
    try:
        return len(load_dist_list(path, quiet=True)[0])
    except Exception:
        return 0


def distribute_new(orders, w, weight_map, bcs, avail):
    """Разнарядка по позициям без истории на точках."""
    if DIST_NEW_MODE == "net" or w is None or not avail:
        return orders
    src = route_key(DIST_SOURCE_SHOP)
    shop_rows = w[w["shop"].map(lambda s: route_key(s) != src)]
    if shop_rows.empty:
        return orders
    have = set(orders["bc"]) if len(orders) else set()
# "Нет и не было на остатках" = ШК вообще отсутствует в ведомости либо
# за период по нему нулевое движение. Товар, который на точках есть и
# продаётся, разнарядкой не досылаем: у него просто закрыт спрос.
    _mv = w[["sale", "out", "start", "end"]].abs().sum(axis=1)
    seen_bc = set(w.loc[_mv > 1e-9, "bc"])
    skip_seen = sorted((bcs & seen_bc) - have)
    if skip_seen:
        R.check("INFO", u"Разнарядка: пропущены товары с историей",
                u"%d ШК есть на точках, но потребности нет - остаются на Полевой"
                % len(skip_seen))
        for b in skip_seen[:40]:
            R.problem(u"История есть, потребности нет", "", b, "",
                      u"разнарядкой не раздаётся")
    new_bc = [b for b in sorted(bcs)
              if b not in have and b not in seen_bc
              and float(avail.get(b, 0)) > 0]
    if not new_bc:
        return orders

    tw = shop_rows.groupby("shop")["sale"].sum()
    if DIST_NEW_MODE == "ravno" or float(tw.sum()) <= 0:
        tw = pd.Series(1.0, index=sorted(set(shop_rows["shop"])))
    tw = tw[tw > 0].sort_values(ascending=False)
    if DIST_NEW_TOP:
        tw = tw.head(DIST_NEW_TOP)
    if tw.empty:
        return orders
    share = tw / float(tw.sum())
    names = w.groupby("bc")["name"].first().to_dict()

    rows, spread = [], 0
    for bc in new_bc:
        total = float(avail[bc])
        is_w = bool((weight_map or {}).get(bc))
        step = WEIGHT_STEP if is_w else 1.0
        nmax = int(total // max(DIST_NEW_MIN, step))
        sh = share if nmax >= len(share) else share.head(max(nmax, 1))
        sh = sh / float(sh.sum())
        left = total
        for shop in sh.index:
            q = math.floor(total * float(sh[shop]) / step + 1e-9) * step
            q = round(min(q, left), 6)
            if q < DIST_NEW_MIN - 1e-9:
                continue
            left = round(left - q, 6)
            rows.append({"shop": shop, "bc": bc, "name": names.get(bc) or DIST_NAMES.get(bc, u""),
                         "qty": q, "raw": q, "step": step})
        if left > 1e-9:
            R.problem(u"Разнарядка: остаток на Полевой", "", bc,
                      names.get(bc, ""), u"не разошлось %g из %g" % (left, total))
        spread += 1

    if not rows:
        R.check("INFO", u"Разнарядка", u"делить нечего: доли меньше минимума")
        return orders
    add = pd.DataFrame(rows)
    for c in orders.columns:
        if c not in add.columns:
            add[c] = 0.0 if orders[c].dtype.kind in "fiu" else ""
    add = add[orders.columns]
    R.check("OK", u"Разнарядка по новым позициям",
            u"%d ШК без истории разложены на %d точек (%s), строк %d"
            % (spread, len(share), DIST_NEW_MODE, len(add)))
    return (pd.concat([orders, add], ignore_index=True)
              .sort_values(["shop", "name"], kind="stable"))


def drop_excluded(orders):
    """Убирает из заказа то, что перевозке не подлежит."""
    if orders.empty or (DIST_SKIP_ONLY_IN_DIST and not dist_on()):
        return orders
    bad = orders["bc"].isin(set(DIST_SKIP_BC))
    low = orders["name"].astype(str).str.lower()
    for pat in DIST_SKIP_NAME:
        bad = bad | low.str.contains(pat, regex=False, na=False)
    if bad.any():
        _skip = set(orders.loc[bad, "bc"])
        R.anomalies = [a for a in R.anomalies if a[1] not in _skip]
    if not bad.any():
        R.check("OK", u"Исключения из перемещения", u"под правило ничего не попало")
        return orders
    cut = orders[bad]
    R.check("INFO", u"Исключено из перемещения",
            u"%d строк, %g ед. - розлив/овощи/кеги/кофе не возим"
            % (len(cut), round(float(cut["qty"].sum()), 3)))
    for bc, g in cut.groupby("bc"):
        R.problem(u"Исключено: не перемещаем", "", bc, g.iloc[0]["name"],
                  u"точек %d, %g ед." % (len(g), round(float(g["qty"].sum()), 3)))
    return orders[~bad].copy()



# ============== ВЫВОЗ ПОДЧИСТУЮ (раздача с Полевой) ==============
DIST_SHIP_ALL = True


def _dist_banned(bc, name):
    """Кеги, розлив, овощи, кофе не раздаём даже разнарядкой."""
    if bc in DIST_SKIP_BC:
        return True
    n = norm(name or "")
    return any(k in n for k in DIST_SKIP_NAME) if n else False


def _dist_shares(w):
    """Доли точек по обороту. Источник исключён."""
    if w is None or not len(w):
        return None
    src = route_key(DIST_SOURCE_SHOP)
    rows = w[w["shop"].map(lambda s: route_key(s) != src)]
    if rows.empty:
        return None
    tw = rows.groupby("shop")["sale"].sum()
    if DIST_NEW_MODE == "ravno" or float(tw.sum()) <= 0:
        tw = pd.Series(1.0, index=sorted(set(rows["shop"])))
    tw = tw[tw > 0].sort_values(ascending=False)
    if DIST_NEW_TOP:
        tw = tw.head(DIST_NEW_TOP)
    return None if tw.empty else tw / float(tw.sum())


def _spread_qty(total, share, step, min_q):
    """Разложить total по точкам, сумма выдачи строго равна total."""
    total = round(float(total), 6)
    if total <= 0 or share is None or len(share) == 0:
        return {}
    units = int(math.floor(total / step + 1e-9))
    tail = round(total - units * step, 6)
    if units <= 0:
        return {share.index[0]: total}
    kmin = max(1, int(round(float(min_q) / step)))
    n = max(1, min(len(share), units // kmin))
    sh = share.head(n)
    sh = sh / float(sh.sum())
    raw = [units * float(sh[s]) for s in sh.index]
    base = [int(math.floor(x)) for x in raw]
    order = sorted(range(len(raw)), key=lambda i: raw[i] - base[i], reverse=True)
    for i in order[:units - sum(base)]:
        base[i] += 1
    give = {s: k for s, k in zip(sh.index, base) if k > 0}
    if not give:
        give = {sh.index[0]: units}
    if sum(give.values()) < units:
        top = max(give, key=lambda s: give[s])
        give[top] += units - sum(give.values())
    out = {s: round(k * step, 6) for s, k in give.items()}
    if tail > 1e-6:
        top = max(out, key=lambda s: out[s])
        out[top] = round(out[top] + tail, 6)
    return out


_BC_CACHE = {}


def _bc_shares(w, bc, share, st=None):
    """Доли точек для конкретного ШК: продажи этого ШК + общий оборот,
    без точек, где ШК лежит мёртвым. Нет продаж нигде -> общий оборот."""
    if share is None or w is None or not len(w):
        return share
    if not (DIST_SHARE_BY_BC or DIST_SKIP_DEAD):
        return share
    key = (id(w), len(w))
    if _BC_CACHE.get("key") != key:
        src = route_key(DIST_SOURCE_SHOP)
        rows = w[w["shop"].map(lambda s: route_key(s) != src)]
        _BC_CACHE.clear()
        _BC_CACHE["key"] = key
        _BC_CACHE["g"] = rows.groupby(["bc", "shop"])[["sale", "end"]].sum()
    g = _BC_CACHE["g"]
    try:
        sub = g.xs(bc, level="bc")
    except KeyError:
        if st is not None:
            st["shop"] += 1
        return share
    base = share.copy()
    if DIST_SKIP_DEAD:
        dead = [s for s in sub.index[(sub["sale"] <= 0) & (sub["end"] > 0)]
                if s in base.index]
        if dead and len(dead) < len(base):
            base = base.drop(dead)
            if st is not None:
                st["dead"] += len(dead)
    base = base / float(base.sum())
    if DIST_SHARE_BY_BC:
        s = sub["sale"].reindex(base.index).fillna(0.0).clip(lower=0.0)
        if float(s.sum()) > 0:
            k = float(DIST_BC_WEIGHT)
            mix = k * s / float(s.sum()) + (1.0 - k) * base
            mix = mix[mix > 0]
            if st is not None:
                st["bc"] += 1
            return (mix / float(mix.sum())).sort_values(ascending=False)
    if st is not None:
        st["shop"] += 1
    return base.sort_values(ascending=False)


def apply_dist_list(orders, w=None, weight_map=None):
    """Вывоз списка с Полевой подчистую."""
    if not dist_on():
        return orders
    if not DIST_SHIP_ALL:
        return apply_dist_list_legacy(orders, w, weight_map)
    bcs, avail = load_dist_list(DIST_FILE)
    if not bcs:
        return orders

    src = route_key(DIST_SOURCE_SHOP)
    was = len(orders) if orders is not None else 0
    cols = (list(orders.columns) if orders is not None and len(orders.columns)
            else ["shop", "bc", "name", "qty", "raw", "step", "is_weight"])
    if orders is None or orders.empty:
        o = pd.DataFrame(columns=cols)
    else:
        o = orders[orders["bc"].isin(bcs)].copy()
        o = o[o["shop"].map(lambda s: route_key(s) != src)]
    R.check("OK", u"Раздача по списку",
            u"потребность: из %d строк осталось %d, точек %d"
            % (was, len(o), o["shop"].nunique() if len(o) else 0))

    share = _dist_shares(w)
    names = {}
    if w is not None and len(w):
        names = {k: v for k, v in w.groupby("bc")["name"].first().to_dict().items() if v}
    try:
        for b, nm0 in DIST_NAMES.items():
            if nm0:
                names.setdefault(b, nm0)
    except NameError:
        pass

    keep, add, rest, ban = [], [], 0.0, 0
    st = {"need": 0, "extra": 0, "full": 0, "noqty": 0, "cut": 0}
    _bs = {"bc": 0, "shop": 0, "dead": 0}
    for bc in sorted(bcs):
        g = o[o["bc"] == bc]
        nm = names.get(bc, u"")
        if _dist_banned(bc, nm):
            ban += 1
            continue
        have = avail.get(bc)
        have = float(have) if have is not None else None
        is_w = bool((weight_map or {}).get(bc))
        step = WEIGHT_STEP if is_w else 1.0
        if have is None or have <= 0:
            if len(g):
                keep.append(g)
            else:
                st["noqty"] += 1
            continue

        left = round(have, 6)
        if len(g):
            g = g.sort_values("raw", ascending=False, kind="stable")
            rows = []
            for i, r0 in g.iterrows():
                if left <= 1e-9:
                    break
                q = float(r0["qty"])
                if q > left + 1e-9:
                    q = math.floor(left / step + 1e-9) * step
                    st["cut"] += 1
                q = round(min(q, left), 6)
                if q < step - 1e-9:
                    continue
                rows.append((i, q))
                left = round(left - q, 6)
            if rows:
                gg = g.loc[[i for i, _ in rows]].copy()
                gg["qty"] = [q for _, q in rows]
                keep.append(gg)
                st["need"] += len(rows)

        if left > 1e-9 and share is not None:
            for shop, q in _spread_qty(left, _bc_shares(w, bc, share, _bs), step, DIST_NEW_MIN).items():
                add.append({"shop": shop, "bc": bc, "name": nm, "qty": q,
                            "raw": q, "step": step, "is_weight": is_w})
                left = round(left - q, 6)
            st["extra" if len(g) else "full"] += 1
        if left > 1e-9:
            rest += left
            R.problem(u"Осталось на Полевой", "", bc, nm, u"не разошлось %g" % left)

    res = pd.concat(keep) if keep else o.iloc[0:0]
    if add:
        a = pd.DataFrame(add)
        for c in cols:
            if c not in a.columns:
                a[c] = 0.0 if (len(res) and res[c].dtype.kind in "fiu") else ""
        res = pd.concat([res, a[cols]], ignore_index=True)
    if len(res):
        res = res.reset_index(drop=True)
        if res.duplicated(["shop", "bc"]).any():
            agg = dict((c, ("sum" if c in ("qty", "raw") else "first"))
                       for c in res.columns if c not in ("shop", "bc"))
            res = res.groupby(["shop", "bc"], as_index=False, sort=False).agg(agg)[cols]
        res = res.sort_values(["shop", "name"], kind="stable")

    ship = res.groupby("bc")["qty"].sum() if len(res) else pd.Series(dtype=float)
    over = [b for b in ship.index
            if b in avail and float(ship[b]) > float(avail[b]) + 1e-6]
    R.check("ERROR" if over else "OK", u"ПЕРЕБОР по наличию на Полевой",
            (u"%d ШК больше остатка: %s" % (len(over), ", ".join(over[:5])))
            if over else u"перебора нет")
    R.check("INFO", u"Доли раздачи остатка",
            u"по продажам ШК (вес %g): %d ШК; по обороту точки: %d ШК; "
            u"пропущено точек, где ШК лежит без продаж: %d"
            % (DIST_BC_WEIGHT, _bs["bc"], _bs["shop"], _bs["dead"]))
    R.check("OK" if rest <= 1e-9 else "WARN", u"Вывоз подчистую",
            u"ШК %d, строк %d, единиц %g; потребностью %d строк, урезано %d, "
            u"доложено разнарядкой %d ШК, роздано целиком %d ШК; "
            u"не перемещаем %d ШК; без количества %d ШК; осталось %g ед."
            % (res["bc"].nunique() if len(res) else 0, len(res),
               round(float(res["qty"].sum()), 3) if len(res) else 0,
               st["need"], st["cut"], st["extra"], st["full"],
               ban, st["noqty"], round(rest, 3)))
    return res


def apply_dist_list_legacy(orders, w=None, weight_map=None):
    """Только ШК из списка, источник убираем, режем под наличие на Полевой.
    Дефицит отдаём точкам с наибольшей потребностью целыми строками:
    количества уже кратны упаковке, пересчитывать округление нельзя."""
    if not dist_on() or orders.empty:
        return orders
    bcs, avail = load_dist_list(DIST_FILE)
    if not bcs:
        return orders
    was = len(orders)
    src = route_key(DIST_SOURCE_SHOP)
    o = orders[orders["bc"].isin(bcs)].copy()
    o = o[o["shop"].map(lambda s: route_key(s) != src)]
    R.check("OK", u"Раздача по списку",
            u"из %d строк осталось %d, точек-получателей %d"
            % (was, len(o), o["shop"].nunique() if len(o) else 0))
    dead = sorted(bcs - set(o["bc"]))
    if dead:
        R.check("INFO", u"Из списка никому не нужны",
                u"%d ШК - потребности нет, остаются на Полевой" % len(dead))
        for b in dead[:MAX_REPORT_ROWS]:
            R.problem(u"Потребности нет ни на одной точке", "", b, "",
                      u"остаётся на Полевой")
    if not (DIST_LIMIT_STOCK and avail) or o.empty:
        return distribute_new(o, w, weight_map, bcs, avail)

    keep, cut = [], 0
    for bc, g in o.groupby("bc", sort=False):
        have = avail.get(bc)
        if have is None:
            keep.append(g)
            continue
        g = g.sort_values("raw", ascending=False, kind="stable")
        left, rows = float(have), []
        for i, r0 in g.iterrows():
            if float(r0["qty"]) <= left + ROUND_EPS:
                left -= float(r0["qty"])
                rows.append(i)
        if len(rows) < len(g):
            cut += len(g) - len(rows)
            R.problem(u"На Полевой не хватает", "", bc, g.iloc[0]["name"],
                      u"нужно %g, есть %g - без отгрузки %d точек"
                      % (round(float(g["qty"].sum()), 3), have, len(g) - len(rows)))
        keep.append(g.loc[rows])
    res = pd.concat(keep) if keep else o.iloc[0:0]
    R.check("OK" if not cut else "WARN", u"Ограничение по наличию на Полевой",
            u"отгружаем %d строк, не хватило на %d" % (len(res), cut))
    res = res.sort_values(["shop", "name"], kind="stable")
    return distribute_new(res, w, weight_map, bcs, avail)


def build_supplier_order(orders):
    """Сводит заказы магазинов в закупку у поставщика с округлением до коробки."""
    if orders.empty:
        return orders
    g = (orders.groupby("bc", as_index=False)
               .agg(name=("name", "first"), art=("art", "first"),
                    shops=("shop", "nunique"), need=("qty", "sum"),
                    step_tr=("step", "first"), ut=("unit_note", "first")))
    sup_step, sup_note, sup_qty = [], [], []
    for bc, need, st in zip(g["bc"], g["need"], g["step_tr"]):
        s, note = resolve_supplier_step(bc, st)
        q = apply_rounding_thr(need, s, SUPPLIER_MIN_FRAC)
        sup_step.append(s); sup_note.append(note); sup_qty.append(q)
    g["sup_step"], g["sup_note"], g["order"] = sup_step, sup_note, sup_qty
    g["izlishek"] = [round(o - n, 3) for o, n in zip(g["order"], g["need"])]

    split = int(sum(1 for a, b in zip(g["sup_step"], g["step_tr"]) if a > b + 1e-9))
    R.check("INFO", "Заказ поставщику",
            "%d позиций; из них %d вскрываются при перемещении (коробка > шага)"
            % (len(g), split))
    big = g[g["izlishek"] > 2 * g["sup_step"]]
    if len(big):
        R.check("WARN", "Заказ поставщику: крупный излишек округления",
                "%d позиций, см. лист заказа" % len(big))
    return g[g["order"] > 0].sort_values("name", kind="stable")


SUP_HDRS = ["№", "Название товара", "Артикул", "Штрих-код", "Магазинов",
            "Нужно магазинам", "Шаг перемещения", "Коробка", "ЗАКАЗ У ПОСТАВЩИКА",
            "Излишек", "Комментарий"]
SUP_W = [5, 46, 11, 15, 11, 16, 16, 10, 20, 11, 34]


def write_supplier_xlsx(path, day, g):
    wb = Workbook()
    ws = wb.active
    ws.title = "Заказ поставщику"
    ws["A1"] = "Сводный заказ поставщику на %s" % day.strftime("%d.%m.%Y")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("Потребность магазинов сведена и округлена вверх до целой коробки. "
                "Перемещение на точки допускает дробление.")
    ws["A2"].font = Font(italic=True, size=9)
    for j, h in enumerate(SUP_HDRS, start=1):
        c0 = ws.cell(row=4, column=j, value=h)
        c0.font = Font(bold=True); c0.fill = HEAD_FILL; c0.border = BORDER
        c0.alignment = Alignment(horizontal="center", wrap_text=True)
    r = 4
    for n, (_, row) in enumerate(g.iterrows(), start=1):
        r += 1
        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=2, value=row["name"])
        ws.cell(row=r, column=3, value=row["art"])
        bc = ws.cell(row=r, column=4, value=row["bc"]); bc.number_format = "@"
        ws.cell(row=r, column=5, value=int(row["shops"]))
        ws.cell(row=r, column=6, value=round(float(row["need"]), 3))
        ws.cell(row=r, column=7, value=float(row["step_tr"]))
        ws.cell(row=r, column=8, value=float(row["sup_step"]))
        q = ws.cell(row=r, column=9, value=round(float(row["order"]), 3))
        q.font = Font(bold=True)
        iz = ws.cell(row=r, column=10, value=float(row["izlishek"]))
        if float(row["izlishek"]) > 2 * float(row["sup_step"]):
            iz.fill = WARN_FILL
        ws.cell(row=r, column=11, value=row["sup_note"])
        for j in range(1, len(SUP_HDRS) + 1):
            ws.cell(row=r, column=j).border = BORDER
            if 5 <= j <= 10:
                ws.cell(row=r, column=j).number_format = "0.###"
    r += 1
    ws.cell(row=r, column=8, value="ИТОГО:").font = Font(bold=True)
    t = ws.cell(row=r, column=9, value=round(float(g["order"].sum()), 3))
    t.font = Font(bold=True); t.fill = WARN_FILL
    for j, wd in enumerate(SUP_W, start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.freeze_panes = "A5"
    wb.save(path)


def apply_rounding_thr(x, step, thr):
    if x is None or x <= 0:
        return 0.0
    u = round(float(x) / float(step), 9)
    fl = math.floor(u)
    n = fl if (u - fl) <= thr else fl + 1
    return round(n * float(step), 6)


def build_orders(w, weight_map):
    import numpy as np

    WEIGHT_CONFLICTS.clear()
    del STEP_FROM_NAME[:]

    h, i, l = w["sale"].to_numpy(float), w["out"].to_numpy(float), w["end"].to_numpy(float)
    _neg = -np.minimum(l, 0.0) if NEG_AS_NEED else np.zeros_like(l)
    vec = np.where(l <= 0, (h + i) * KOEF + _neg,
                   np.where(h > 0, np.maximum((h + i) * KOEF - l, 0.0), 0.0))
    if NEG_AS_NEED:
        _nm = l < 0
        if int(_nm.sum()):
            R.check("INFO", "Минусовой остаток включён в потребность",
                    "%d строк, добавлено %.3f ед. к расчёту"
                    % (int(_nm.sum()), float(-l[_nm].sum())))
    loop = np.array([calc_order(a, b, cc) for a, b, cc in zip(h, i, l)])
    ev = np.array([calc_order_eval(a, b, cc) for a, b, cc in zip(h, i, l)])

    d1 = float(np.max(np.abs(vec - loop))) if len(vec) else 0.0
    d2 = float(np.max(np.abs(vec - ev))) if len(vec) else 0.0
    if max(d1, d2) > 1e-9:
        R.check("ERROR", "Сверка трёх реализаций формулы",
                "макс. расхождение %.12f" % max(d1, d2))
    else:
        R.check("OK", "Сверка трёх реализаций формулы",
                "векторная = построчная = разбор текста, %d строк" % len(vec))

    w = w.copy()
    w["raw"] = vec

    grp = (w.groupby(["shop", "bc"], as_index=False)
            .agg(name=("name", "first"), art=("art", "first"),
                 sale=("sale", "sum"), out=("out", "sum"),
                 end=("end", "sum"), raw=("raw", "sum"), dup=("bc", "size")))

    dups = int((grp["dup"] > 1).sum())
    R.check("INFO", "Дубли товара внутри магазина",
            "сведено позиций: %d" % dups if dups else "дублей нет")

    grp["is_weight"] = grp["bc"].map(lambda b: bool(weight_map.get(b, False)))
    if DROP_BELOW > 0:
        grp.loc[grp["raw"] < DROP_BELOW, "raw"] = 0.0

    _s, _t, _m, _nt = [], [], [], []
    for _b, _hw, _nm, _ar in zip(grp["bc"], grp["is_weight"], grp["name"], grp["art"]):
        a, b2, m2, c2 = resolve_step(_b, _hw, _nm, _ar)
        _s.append(a); _t.append(b2); _m.append(m2); _nt.append(c2)
    grp["step"], grp["thr"], grp["min_q"], grp["unit_note"] = _s, _t, _m, _nt

    # весовой ящик: вес из названия важнее шага из истории
    boxes, steps2, n_nm, n_ref = [], [], 0, 0
    for _b, _hw, _nm2, _st in zip(grp["bc"], grp["is_weight"], grp["name"], grp["step"]):
        if not _hw:
            boxes.append(0.0)
            steps2.append(float(_st))
            continue
        _n2 = _norm_name(_nm2).strip()
        _por = next((v for k, v in VES_PORTION_BY_NAME.items() if _n2.startswith(k)), 0.0)
        if _por > 0:
            boxes.append(0.0)          # ящик не навязываем: есть подфасовка
            steps2.append(_por)
            continue
        nb = box_from_name(_nm2)
        if nb > 0:
            steps2.append(nb)          # вес из названия - всегда шаг, даже мелкий
        else:
            steps2.append(float(_st))
        bx = 0.0
        if nb >= VES_BOX_MIN:
            bx = nb
            n_nm += 1
        elif nb <= 0 and str(REF.get(_b, {}).get("unit_type", "")) == "ves_fix" \
                and float(_st) >= VES_BOX_MIN:
            bx = float(_st)
            n_ref += 1
        boxes.append(bx)
    grp["boxkg"] = boxes
    grp["step"] = steps2
    # ПАТЧ 1: порог обязан соответствовать шагу. Если шаг подменён весом ящика,
    # thr от весового режима (почти ноль) поднимает до двух ящиков, VES_BOX_RULE
    # срезает обратно до одного -> ложный ERROR "Округление не вверх".
    _thr2, _n_thr = [], 0
    for _bk2, _th2 in zip(grp["boxkg"], grp["thr"]):
        if float(_bk2 or 0.0) >= VES_BOX_MIN and float(_th2) <= 0.01:
            _thr2.append(FIX_MIN_FRAC)
            _n_thr += 1
        else:
            _thr2.append(float(_th2))
    grp["thr"] = _thr2
    if _n_thr:
        R.check("INFO", "Порог согласован с шагом ящика",
                "%d позиций: шаг = целый ящик, порог %s" % (_n_thr, FIX_MIN_FRAC))
    if n_nm or n_ref:
        R.check("INFO", "Весовой товар в ящиках",
                "ящик из названия: %d позиций, из справочника: %d" % (n_nm, n_ref))

    if WEIGHT_CONFLICTS:
        uniq = {c[0]: c for c in WEIGHT_CONFLICTS}
        R.check("WARN", "Справочник помечает весовой товар как штучный",
                "%d позиций переведены на шаг %g по данным ведомости"
                % (len(uniq), WEIGHT_STEP))
        for bc_, nm_, st_ in list(uniq.values())[:20]:
            R.problem("Тип в справочнике: 'шт', фактически вес", "", bc_, nm_,
                      "было %g -> стало %g; поправить unit_type в BigQuery" % (st_, WEIGHT_STEP))

    in_ref = int(grp["bc"].map(lambda b: b in REF).sum())
    R.check("INFO", "Покрытие справочником",
            "%d из %d позиций найдены в transfer_pack" % (in_ref, len(grp)))

    if STEP_FROM_NAME:
        _u = {}
        for _bx, _nx, _sx, _boxx, _ntx in STEP_FROM_NAME:
            _u[_bx] = (_nx, _sx, _ntx)
        R.check("WARN", "Шаг взят из названия (нет в справочнике)",
                "%d артикулов - проверь лист 'Проблемы данных'" % len(_u))
        for _bx, (_nx, _sx, _ntx) in list(_u.items())[:80]:
            R.problem("Шаг из названия (проверь)", "", _bx, _nx,
                      "шаг %g; %s" % (_sx, _ntx))

    _ex = grp["bc"].map(lambda b: REF.get(b, {}).get("unit_type") == "exclude")
    if int(_ex.sum()):
        R.check("INFO", "Исключено справочником", "%d позиций" % int(_ex.sum()))
        grp = grp[~_ex].copy()

    # 1) базовое округление по шагу и порогу
    grp["base"] = [apply_rounding_thr(x, s, t)
                   for x, s, t in zip(grp["raw"], grp["step"], grp["thr"])]

    # 2) нижняя граница: страховка при нулевом остатке + min_q из справочника
    floors, cuts, n_ins, n_mq, n_cut = [], [], 0, 0, 0
    for base, raw, end, step, thr, mq in zip(grp["base"], grp["raw"], grp["end"],
                                             grp["step"], grp["thr"], grp["min_q"]):
        fl, cut = 0.0, False
        if base <= 0 and raw > 0 and end <= 0 and thr > 0.01:
            fl = float(step)
            n_ins += 1
        if mq and mq > 0 and max(base, fl) > 0:
            mq_c = apply_rounding_thr(mq, step, ROUND_EPS)
            if max(base, fl) + 1e-9 < mq_c:
                if step <= WEIGHT_STEP + 1e-9 and raw < VES_MIN_FILL * mq_c:
                    cut = True          # мельче доли порции - в эту волну не везём
                    n_cut += 1
                else:
                    fl = mq_c
                    n_mq += 1
        floors.append(0.0 if cut else fl)
        cuts.append(cut)
    grp["floor"] = floors
    grp["qty"] = [0.0 if c else max(b, f)
                  for b, f, c in zip(grp["base"], grp["floor"], cuts)]

    # 2а) ящик не вскрывают: либо целый ящик, либо ноль.
    #     Остаток не больше VES_BOX_STOCK_MAX - везём ящик, иначе ждём.
    # ПАТЧ 2: помечаем позиции, количество которых задала политика ящика,
    # и разделяем подъём и срез - раньше срез считался как "поднято".
    grp["ves_box"] = False
    if VES_BOX_RULE:
        qs, box_hit, n_up, n_down, n_zero, n_wait = [], [], 0, 0, 0, 0
        wait_rows = []
        for q, bx, end_, raw_, shop_, bc_, nm_ in zip(
                grp["qty"], grp["boxkg"], grp["end"], grp["raw"],
                grp["shop"], grp["bc"], grp["name"]):
            bx = float(bx or 0.0)
            if bx < VES_BOX_MIN:
                qs.append(q)
                box_hit.append(False)
                continue
            if float(end_) <= VES_BOX_STOCK_MAX:
                if (VES_BOX_NEED_DEMAND and float(raw_) <= 1e-9
                        and float(end_) > VES_BOX_CRUMB):
                    n_wait += 1
                    wait_rows.append((shop_, bc_, nm_, float(end_)))
                    qs.append(0.0)
                    box_hit.append(True)
                    continue
                nq = bx * VES_BOX_QTY
                if nq > float(q) + 1e-9:
                    n_up += 1
                elif nq < float(q) - 1e-9:
                    n_down += 1
                qs.append(nq)
                box_hit.append(True)
            else:
                if float(q) > 0:
                    n_zero += 1
                qs.append(0.0)
                box_hit.append(True)
        grp["ves_box"] = box_hit
        if n_wait:
            R.check("INFO", "Весовой ящик: ждём спроса",
                    "%d позиций (продаж нет, хвост больше %g кг)"
                    % (n_wait, VES_BOX_CRUMB))
            for _s, _b, _n, _e in wait_rows[:40]:
                R.problem("Ящик не повезли: нет продаж", _s, _b, _n,
                          "остаток %.3f кг, продаж 0 - нужен ли товар в точке?" % _e)
        if n_down:
            R.check("WARN", "Весовой ящик: заказ срезан до целого ящика",
                    "%d позиций (больше %d ящика за волну не возим)"
                    % (n_down, VES_BOX_QTY))
        grp["qty"] = qs
        if n_up:
            R.check("INFO", "Весовой ящик: поднято до целого ящика",
                    "%d позиций (остаток не больше %g кг)" % (n_up, VES_BOX_STOCK_MAX))
        if n_zero:
            R.check("INFO", "Весовой ящик: потребность обнулена",
                    "%d позиций (остаток больше %g кг, ящик не вскрываем)"
                    % (n_zero, VES_BOX_STOCK_MAX))
    if n_cut:
        R.check("INFO", "Весовой: мелкая потребность отброшена",
                "%d позиций (меньше %g порции по истории перемещений)"
                % (n_cut, VES_MIN_FILL))

    if n_ins:
        R.check("INFO", "Страховка при нулевом остатке",
                "поднято до 1 упаковки: %d позиций (порог %s срезал бы в ноль)"
                % (n_ins, FIX_MIN_FRAC))
    if n_mq:
        R.check("INFO", "Поднято до min_q справочника", "%d позиций" % n_mq)

    _chk = grp[grp["step"] > 1.0]
    _bad = _chk[((_chk["qty"] / _chk["step"]) -
                 (_chk["qty"] / _chk["step"]).round()).abs() > 1e-6]
    if len(_bad):
        R.check("ERROR", "Кратность нарушена", "%d позиций" % len(_bad))
    else:
        R.check("OK", "Кратность соблюдена",
                "все заказы кратны шагу (%d позиций с шагом > 1)" % len(_chk))

    # контроль направления округления
    # для фасовки/шоубоксов недобор допустим, но не больше порога FIX_MIN_FRAC от упаковки
    # ПАТЧ 3: позиции, заданные правилом ящика, из проверки округления убраны:
    # их количество определяет политика отгрузки, а не формула потребности.
    free = grp[~grp["ves_box"].astype(bool)]
    plain = free[(free["qty"] > 0) & (free["thr"] <= 0.01)]
    wrong = plain[plain["qty"] + 1e-9 < plain["raw"]]

    packs = free[(free["qty"] > 0) & (free["thr"] > 0.01)]
    over = packs[packs["raw"] - packs["qty"] > packs["thr"] * packs["step"] + 1e-9]
    cut = int((packs["raw"] - packs["qty"] > 1e-9).sum())

    if len(wrong) or len(over):
        R.check("ERROR", "Округление не вверх",
                "штучных/весовых занижено %d, фасовок ниже допустимого %d"
                % (len(wrong), len(over)))
        for _, b0 in list(wrong.iterrows())[:10] + list(over.iterrows())[:10]:
            R.problem("Занижено против расчёта", b0["shop"], b0["bc"], b0["name"],
                      "расчёт %.3f -> заказ %.3f, шаг %g" % (b0["raw"], b0["qty"], b0["step"]))
    else:
        R.check("OK", "Округление вверх",
                "занижений нет; неполная упаковка отброшена по порогу %s: %d позиций"
                % (FIX_MIN_FRAC, cut))

    if BLOCK_MIN_FILL > 0:
        _blk = [float(transfer_block(nm, bc)) for nm, bc in zip(grp["name"], grp["bc"])]
        _bag = [bool(is_bag(nm, bc)) for nm, bc in zip(grp["name"], grp["bc"])]
        _blk = pd.Series(_blk, index=grp.index)
        _bag = pd.Series(_bag, index=grp.index)
        # пакеты не отбрасываем: 1 шт потребности -> всё равно везём 50
        _thin = (_blk > 0) & (~_bag if BAG_ALWAYS_UP else True) & (grp["raw"] < _blk * BLOCK_MIN_FILL)
        if int(_thin.sum()):
            R.check("INFO", "Блочный товар: мелкая потребность отброшена",
                    "%d позиций (меньше %g блока)" % (int(_thin.sum()), BLOCK_MIN_FILL))
            grp.loc[_thin, "qty"] = 0.0

    res = grp[grp["qty"] > 0].copy()
    zero = len(grp) - len(res)
    R.check("INFO", "Отброшено нулевых позиций", zero)

    for _, r0 in res.iterrows():
        why = []
        if r0["sale"] <= 0 and r0["qty"] > 0:
            why.append("заказ без реализации (остаток <= 0)")
        if r0["end"] < 0:
            _ut = str(REF.get(r0["bc"], {}).get("unit_type", ""))
            _piece = (float(r0["step"]).is_integer() and float(r0["step"]) >= 1
                      and not _ut.startswith("ves") and not bool(r0.get("is_weight")))
            if _piece or r0["end"] <= -NEG_END_TOL:
                why.append("отрицательный остаток %.3f%s"
                           % (r0["end"], " (штучный товар!)" if _piece else ""))
        if r0["qty"] > ANOMALY_ABS:
            why.append("очень крупный заказ")
        if r0["sale"] > 0 and r0["qty"] > ANOMALY_FACTOR * r0["sale"] and r0["qty"] >= 10:
            why.append("заказ > %dx реализации" % ANOMALY_FACTOR)
        if why:
            R.anomaly(r0["shop"], r0["bc"], r0["name"], r0["sale"], r0["out"],
                      r0["end"], r0["raw"], r0["qty"], "; ".join(why))
    if R.anomalies:
        R.check("WARN", "Аномальные позиции", "%d шт. - см. лист 'Аномалии'" % len(R.anomalies))
    else:
        R.check("OK", "Аномальные позиции", "не обнаружены")

    return res.sort_values(["shop", "name"], kind="stable")

def replace_cups_with_packs(res, w):
    """Стаканы в магазин везут упаковками. Поштучную позицию убираем,
    вместо неё ставим упаковку: количество = потребность в штуках / вместимость,
    округление вверх, кратность 1 упаковка."""
    if res is None or res.empty:
        return res
    import math as _m
    import pandas as _pd
    hit = res["bc"].isin(PACK_SWAP)
    if not bool(hit.any()):
        return res
    src = w.groupby(["shop", "bc"], as_index=False).agg(
        name=("name", "first"), art=("art", "first"),
        sale=("sale", "sum"), out=("out", "sum"), end=("end", "sum"))
    base = {c: 0.0 for c in ("sale", "out", "end", "raw", "base", "floor")}
    new_rows, moved = [], 0
    for _, r0 in res[hit].iterrows():
        pack_bc, per_pack = PACK_SWAP[r0["bc"]]
        need_pcs = float(r0["qty"])
        n_pack = max(1, int(_m.ceil(need_pcs / float(per_pack) - 1e-9)))
        moved += 1
        sel = (res["shop"] == r0["shop"]) & (res["bc"] == pack_bc)
        if bool(sel.any()):
            res.loc[sel, "qty"] = res.loc[sel, "qty"].astype(float) + n_pack
            continue
        info = REF.get(pack_bc) or {}
        row = dict(base)
        row.update(shop=r0["shop"], bc=pack_bc, qty=float(n_pack), step=1.0,
                   thr=ROUND_EPS, min_q=0.0, dup=1, is_weight=False,
                   unit_note="упаковка замість %g шт" % need_pcs)
        f = src[(src["shop"] == r0["shop"]) & (src["bc"] == pack_bc)]
        row["name"] = (f["name"].iloc[0] if len(f) else str(info.get("pname") or pack_bc))
        row["art"] = (f["art"].iloc[0] if len(f) else "")
        if len(f):
            row["sale"], row["out"], row["end"] = (float(f["sale"].iloc[0]),
                                                   float(f["out"].iloc[0]),
                                                   float(f["end"].iloc[0]))
        new_rows.append(row)
    res = res[~hit].copy()
    if new_rows:
        res = _pd.concat([res, _pd.DataFrame(new_rows)], ignore_index=True)
    R.check("INFO", "Поштучные карточки заменены упаковками",
            "%d позиций (стаканы и пакеты: закупщице нужны упаковки, не штуки)"
            % moved)
    return res.sort_values(["shop", "name"], kind="stable")


def add_bottle_accessories(res, w):
    """К пивным бутылкам добавляет крышки (1:1) и ручки (1:1 к бутылкам 1-2 л).
    Проверено по истории мая-августа: крышек на бутылку 1,02; ручек на бутылку 1-2 л 0,90."""
    if res is None or res.empty:
        return res
    import pandas as _pd
    base = {c: 0.0 for c in ("sale", "out", "end", "raw", "base", "floor")}
    src = w.groupby(["shop", "bc"], as_index=False).agg(
        name=("name", "first"), art=("art", "first"),
        sale=("sale", "sum"), out=("out", "sum"), end=("end", "sum"))
    added, raised = 0, 0
    new_rows = []
    for shop, part in res.groupby("shop", sort=False):
        got = dict(zip(part["bc"], part["qty"]))
        need = {
            BC_KRISHKA: sum(v for b, v in got.items() if b in BC_PLYASHKI_ALL),
            BC_RUCHKA:  sum(v for b, v in got.items() if b in BC_PLYASHKI_RUCHKA),
        }
        for bc, n in need.items():
            if n <= 0:
                continue
            info = REF.get(bc) or {}
            step = float(info.get("step") or 1.0) or 1.0
            n_up = apply_rounding_thr(n, step, ROUND_EPS)
            if bc in got:
                if got[bc] + 1e-9 < n_up:
                    res.loc[(res["shop"] == shop) & (res["bc"] == bc), "qty"] = n_up
                    raised += 1
                continue
            row = dict(base)
            row.update(shop=shop, bc=bc, qty=n_up, step=step, thr=ROUND_EPS,
                       min_q=0.0, dup=1, is_weight=False,
                       unit_note="комплект к пивній пляшці")
            f = src[(src["shop"] == shop) & (src["bc"] == bc)]
            row["name"] = (f["name"].iloc[0] if len(f)
                           else str(info.get("pname") or bc))
            row["art"] = (f["art"].iloc[0] if len(f) else "")
            if len(f):
                row["sale"], row["out"], row["end"] = (float(f["sale"].iloc[0]),
                                                       float(f["out"].iloc[0]),
                                                       float(f["end"].iloc[0]))
            new_rows.append(row); added += 1
    if new_rows:
        res = _pd.concat([res, _pd.DataFrame(new_rows)], ignore_index=True)
    if added or raised:
        R.check("INFO", "Комплект к пивным бутылкам",
                "добавлено позиций: %d, поднято до нормы: %d (крышка 1:1, ручка к 1-2 л)"
                % (added, raised))
    return res.sort_values(["shop", "name"], kind="stable")


# ========================== ЗАПИСЬ ===========================

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
OK_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
ERR_FILL = PatternFill("solid", fgColor="FCE4E4")

HDRS = ["№", "Название товара", "Артикул", "Штрих-код", "[-] реализация",
        "[-] перемещение", "на конец", "Расчёт (скрипт)", "Расчёт (формула Excel)",
        "ЗАКАЗ", "Контроль", "Упаковка"]
WIDTHS = [5, 44, 11, 15, 13, 14, 10, 15, 17, 10, 14, 24]
LAST_COL = len(HDRS)          # 12 -> L


def excel_qty_formula(row_i, step, thr, floor=0.0):
    """Excel-формула округления, идентичная логике скрипта (с порогом строки и минимумом)."""
    s = repr(float(step))
    t = repr(float(thr))
    u = "I{r}/{s}".format(r=row_i, s=s)
    f = "IF({u}-INT({u})<={t},INT({u}),INT({u})+1)*{s}".format(u=u, t=t, s=s)
    if floor and float(floor) > 0:
        f = "MAX({f},{fl})".format(f=f, fl=repr(float(floor)))
    return f


def write_shop_xlsx(path, shop, day, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Заказ"
    ws["A1"] = "Заказ: %s" % shop
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("Дата ведомости: %s     Позиций: %d     Единиц: %s     Формула: %s     "
                "Округление: %s, порог фасовки %s"
                % (day.strftime("%d.%m.%Y"), len(rows),
                   round(float(rows["qty"].sum()), 3),
                   FORMULA_TEXT, ROUND_MODE, FIX_MIN_FRAC))
    ws["A2"].font = Font(italic=True, size=9)

    hr = 4
    for j, h in enumerate(HDRS, start=1):
        c0 = ws.cell(row=hr, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = HEAD_FILL
        c0.border = BORDER
        c0.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    r = hr
    for n, (_, row) in enumerate(rows.iterrows(), start=1):
        r += 1
        step = float(row["step"])
        thr = float(row["thr"])
        floor = float(row.get("floor", 0.0) or 0.0)

        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=2, value=row["name"])
        ws.cell(row=r, column=3, value=row["art"])
        bc = ws.cell(row=r, column=4, value=row["bc"])
        bc.number_format = "@"
        ws.cell(row=r, column=5, value=float(row["sale"]))
        ws.cell(row=r, column=6, value=float(row["out"]))
        ws.cell(row=r, column=7, value=float(row["end"]))
        ws.cell(row=r, column=8, value=round(float(row["raw"]), 4))
        ws.cell(row=r, column=9, value=(
            ("=IF(G{r}<=0,(E{r}+F{r})*{k}-G{r},IF(E{r}>0,MAX((E{r}+F{r})*{k}-G{r},0),0))" if NEG_AS_NEED else "=IF(G{r}<=0,(E{r}+F{r})*{k},IF(E{r}>0,MAX((E{r}+F{r})*{k}-G{r},0),0))")
        ).format(r=r, k=repr(KOEF)))
        q = ws.cell(row=r, column=10, value=round(float(row["qty"]), 3))
        q.font = Font(bold=True)
        ws.cell(row=r, column=11, value=(
            '=IF(AND(ABS(H{r}-I{r})<0.0005,ABS({f}-J{r})<0.0005),"OK","РАСХОЖДЕНИЕ")'
        ).format(r=r, f=excel_qty_formula(r, step, thr, floor)))
        ws.cell(row=r, column=12, value=row.get("unit_note", ""))

        for j in range(1, LAST_COL + 1):
            cell = ws.cell(row=r, column=j)
            cell.border = BORDER
            if 5 <= j <= 10:
                cell.number_format = "0.###"
                cell.alignment = Alignment(horizontal="center")

    last = r
    r += 1
    ws.cell(row=r, column=9, value="ИТОГО:").font = Font(bold=True)
    tot = ws.cell(row=r, column=10, value=round(float(rows["qty"].sum()), 3))
    tot.font = Font(bold=True)
    tot.fill = WARN_FILL

    ws["A3"] = ('=IF(COUNTIF(K5:K{n},"РАСХОЖДЕНИЕ")=0,"Excel-контроль: OK, расчёт скрипта '
                'совпал с формулой во всех строках","Excel-контроль: РАСХОЖДЕНИЙ "'
                '&COUNTIF(K5:K{n},"РАСХОЖДЕНИЕ"))').format(n=last)
    ws["A3"].font = Font(bold=True, color="1F6F3F")

    for j, wd in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.freeze_panes = "A%d" % (hr + 1)
    ws.auto_filter.ref = "A%d:%s%d" % (hr, get_column_letter(LAST_COL), last)
    wb.save(path)


def write_shop_txt(path, rows):
    # PATCH-TXT-FRAC v3: формат числа задаёт САМО количество, а не step.
    # Дробное qty при целом step уходило в int(round(q)): 0.175 -> 0, 1.1 -> 1.
    lines = []
    for _, row in rows.iterrows():
        q = float(row["qty"])
        try:
            step = float(row["step"])
        except Exception:
            step = 1.0
        as_w = (abs(q - round(q)) > 1e-9
                or step <= 0 or not float(step).is_integer()
                or bool(row.get("is_weight", False)))
        lines.append("%s;%s" % (row["bc"], qty_to_text(q, as_w)))
    with open(path, "w", encoding=TXT_ENCODING, errors="replace", newline="") as f:
        f.write(TXT_NEWLINE.join(lines) + TXT_NEWLINE)
    return lines


# =================== ПРОВЕРКА ЗАПИСАННОГО =====================

TXT_LINE_RE = re.compile(r"^\d{6,14};\d+(?:[.,]\d+)?$")


def verify_written(shop_dir, fname, rows):
    """Обратное чтение xlsx и txt со сверкой с расчётом."""
    problems = []
    xlsx_path = os.path.join(shop_dir, fname + ".xlsx")
    txt_path = os.path.join(shop_dir, fname + ".txt")

    expected = {}
    for _, row in rows.iterrows():
        expected[row["bc"]] = round(float(row["qty"]), 3)

    wb = load_workbook(xlsx_path, data_only=False)
    ws = wb.active
    got = {}
    r = 5
    while ws.cell(row=r, column=4).value not in (None, ""):
        b = fmt_barcode(ws.cell(row=r, column=4).value)
        v = ws.cell(row=r, column=10).value
        got[b] = round(float(v), 3) if isinstance(v, (int, float)) else None
        r += 1
    wb.close()
    if got != expected:
        _diff = []
        for _b in sorted(set(expected) | set(got)):
            _e, _g = expected.get(_b), got.get(_b)
            if _e != _g:
                _diff.append("%s: расчёт %s / xlsx %s" % (_b, _e, _g))
        problems.append("xlsx: позиций %d из %d, расхождений %d -> %s"
                        % (len(got), len(expected), len(_diff), "; ".join(_diff[:8])))

    with open(txt_path, "r", encoding=TXT_ENCODING) as f:
        raw_lines = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    if len(raw_lines) != len(expected):
        problems.append("txt: строк %d вместо %d" % (len(raw_lines), len(expected)))
    tgot = {}
    for ln in raw_lines:
        if not TXT_LINE_RE.match(ln):
            problems.append("txt: неверный формат строки '%s'" % ln)
            continue
        b, q = ln.split(";")
        tgot[b] = round(float(q.replace(",", ".")), 3)
    if tgot != expected:
        _diff = []
        for _b in sorted(set(expected) | set(tgot)):
            _e, _g = expected.get(_b), tgot.get(_b)
            if _e != _g:
                _diff.append("%s: расчёт %s / txt %s" % (_b, _e, _g))
        problems.append("txt: расхождений %d -> %s"
                        % (len(_diff), "; ".join(_diff[:8])))
    if os.path.getsize(txt_path) == 0:
        problems.append("txt: файл пустой")
    return problems

# ======================= ОТЧЁТЫ ==============================

def write_summary(path, day, summary, src):
    wb = Workbook()
    ws = wb.active
    ws.title = "Сводка"
    ws["A1"] = "Сводка заказов на %s" % day.strftime("%d.%m.%Y")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Источник: %s" % src
    ws["A2"].font = Font(italic=True, size=9)
    for j, h in enumerate(["№", "Адрес (Центр учета)", "Позиций", "Единиц",
                           "Проверка файлов", "Папка"], start=1):
        c0 = ws.cell(row=4, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = HEAD_FILL
    for n, it in enumerate(summary, start=1):
        ws.cell(row=4 + n, column=1, value=n)
        ws.cell(row=4 + n, column=2, value=it["shop"])
        ws.cell(row=4 + n, column=3, value=it["positions"])
        ws.cell(row=4 + n, column=4, value=it["units"])
        st = ws.cell(row=4 + n, column=5, value=it["verify"])
        st.fill = OK_FILL if it["verify"] == "OK" else ERR_FILL
        ws.cell(row=4 + n, column=6, value=it["folder"])
    for j, wd in enumerate([5, 32, 10, 10, 34, 62], start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd
    ws.freeze_panes = "A5"
    wb.save(path)


def write_validation(path, day, src, shops_all, shops_with_order):
    wb = Workbook()
    ws = wb.active
    ws.title = "Проверки"
    ws["A1"] = "Отчёт валидации заказа на %s" % day.strftime("%d.%m.%Y")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Источник: %s   |   запуск %s" % (src, datetime.now().strftime("%d.%m.%Y %H:%M"))
    ws["A2"].font = Font(italic=True, size=9)
    verdict = "ЕСТЬ ОШИБКИ" if R.errors else ("ЕСТЬ ПРЕДУПРЕЖДЕНИЯ" if R.warns else "ВСЁ ЧИСТО")
    v = ws["A3"]
    v.value = "ИТОГ: %s (ошибок %d, предупреждений %d)" % (verdict, len(R.errors), len(R.warns))
    v.font = Font(bold=True, size=12)
    v.fill = ERR_FILL if R.errors else (WARN_FILL if R.warns else OK_FILL)

    for j, h in enumerate(["Статус", "Проверка", "Детали"], start=1):
        c0 = ws.cell(row=5, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = HEAD_FILL
    for n, (st, nm, det) in enumerate(R.checks, start=1):
        ws.cell(row=5 + n, column=1, value=st).fill = {
            "OK": OK_FILL, "WARN": WARN_FILL, "ERROR": ERR_FILL,
            "INFO": PatternFill("solid", fgColor="EDEDED")}[st]
        ws.cell(row=5 + n, column=2, value=nm)
        ws.cell(row=5 + n, column=3, value=det)
    for j, wd in enumerate([10, 52, 90], start=1):
        ws.column_dimensions[get_column_letter(j)].width = wd

    ws2 = wb.create_sheet("Проблемы данных")
    for j, h in enumerate(["Тип", "Адрес", "Штрих-код", "Товар", "Детали"], start=1):
        c0 = ws2.cell(row=1, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = HEAD_FILL
    for n, row in enumerate(R.problems, start=1):
        for j, v2 in enumerate(row, start=1):
            ws2.cell(row=1 + n, column=j, value=v2)
    for j, wd in enumerate([38, 26, 16, 42, 40], start=1):
        ws2.column_dimensions[get_column_letter(j)].width = wd

    ws3 = wb.create_sheet("Аномалии")
    for j, h in enumerate(["Адрес", "Штрих-код", "Товар", "реализация", "перемещение",
                           "на конец", "Расчёт", "ЗАКАЗ", "Причина"], start=1):
        c0 = ws3.cell(row=1, column=j, value=h)
        c0.font = Font(bold=True)
        c0.fill = HEAD_FILL
    for n, row in enumerate(R.anomalies, start=1):
        for j, v2 in enumerate(row, start=1):
            ws3.cell(row=1 + n, column=j, value=v2)
    for j, wd in enumerate([26, 16, 42, 12, 13, 11, 11, 10, 44], start=1):
        ws3.column_dimensions[get_column_letter(j)].width = wd

    ws4 = wb.create_sheet("Адреса")
    ws4.cell(row=1, column=1, value="Адрес (Центр учета)").font = Font(bold=True)
    ws4.cell(row=1, column=2, value="Есть заказ").font = Font(bold=True)
    for n, s in enumerate(shops_all, start=1):
        ws4.cell(row=1 + n, column=1, value=s)
        ws4.cell(row=1 + n, column=2, value="да" if s in shops_with_order else "нет")
    ws4.column_dimensions["A"].width = 34
    wb.save(path)

# =========================== MAIN ============================

# ---- итог раздачи для окна (PATCH-DIST-SUM v1) ----
_PRICE_KEYS = ((u"себест", u"собівар"),      # себестоимость - приоритет
               (u"закуп",),
               (u"цена", u"ціна"),
               (u"сумма", u"сума"))          # сумма строки / количество


def _num_col(s):
    return pd.to_numeric(s.astype(str).str.replace(u"\u00a0", u"")
                          .str.replace(u" ", u"").str.replace(u",", u"."),
                         errors="coerce")


def _dist_price_map():
    """ШК -> цена за единицу из файла вывоза. -> (dict, откуда взято)"""
    res, used = {}, u""
    try:
        book = pd.read_excel(DIST_FILE, sheet_name=None, dtype=str)
    except Exception as e:
        return res, u"файл вывоза не прочитан (%s)" % type(e).__name__
    for sh, d in book.items():
        cols = {norm(c): c for c in d.columns}
        c_bc = next((cols[k] for k in cols if u"штрих" in k), None)
        if c_bc is None:
            continue
        c_pr, kind = None, None
        for n, keys in enumerate(_PRICE_KEYS):
            c_pr = next((cols[k] for k in cols if any(x in k for x in keys)), None)
            if c_pr is not None:
                kind = n
                break
        if c_pr is None:
            continue
        pr = _num_col(d[c_pr])
        if kind == len(_PRICE_KEYS) - 1:
            c_q = next((cols[k] for k in cols if u"колич" in k or u"кільк" in k), None)
            if c_q is None:
                continue
            qq = _num_col(d[c_q])
            pr = pr / qq.where(qq > 0)
        for b, p in zip(d[c_bc], pr):
            if pd.isna(p) or p <= 0:
                continue
            b = restore_barcode(fmt_barcode(b))[0]
            if b and b not in res:
                res[b] = float(p)
        if not used:
            used = u"лист '%s', колонка '%s'" % (sh, c_pr)
    return res, used


def _money(x):
    return (u"{:,.2f}".format(float(x))).replace(u",", u" ")


def dist_summary_text(orders):
    """Строка итога раздачи для окна + разбивка по точкам в лог."""
    try:
        if not dist_on() or orders is None or not len(orders):
            return u""
        o = orders[orders["qty"] > 0]
        units = round(float(o["qty"].sum()), 3)
        txt = (u"\n\nРаздача с %s: %s ед. · точек %d · строк %d · ШК %d"
               % (DIST_SOURCE_SHOP, units, o["shop"].nunique(), len(o),
                  o["bc"].nunique()))
        prices, used = {}, u""   # цен нет - сумму не считаем
        R.say(u"-" * 70)
        R.say(u"Раздача с %s      : %s ед." % (DIST_SOURCE_SHOP, units))
        if prices:
            pv = o["bc"].map(prices)
            o = o.assign(_m=o["qty"] * pv)
            money = float(o["_m"].sum())
            miss = int(pv.isna().sum())
            txt += u"\nСумма раздачи: %s грн" % _money(money)
            if miss:
                txt += u" (без цены %d строк)" % miss
            R.say(u"Сумма раздачи       : %s грн, цены: %s, без цены строк: %d"
                  % (_money(money), used, miss))
            per = o.groupby("shop")[["qty", "_m"]].sum().sort_values("_m", ascending=False)
            for s, r in per.iterrows():
                R.say(u"  %-26s %10s ед. %14s грн"
                      % (s, round(float(r["qty"]), 3), _money(r["_m"])))
        else:
            per = o.groupby("shop")["qty"].sum().sort_values(ascending=False)
            for s, q in per.items():
                R.say(u"  %-26s %10s ед." % (s, round(float(q), 3)))
        return txt
    except Exception as e:
        R.say(u"Итог раздачи не посчитан: %s: %s" % (type(e).__name__, e))
        return u""


def main():
    ensure_dirs()
    R.say("=" * 70)
    R.say("РАСЧЁТ ЗАКАЗА С ВАЛИДАЦИЕЙ  |  %s" % datetime.now().strftime("%d.%m.%Y %H:%M:%S"))
    R.say("=" * 70)

    if not self_test():
        R.say("Самотестирование не пройдено, работа прервана.")
        return 2

    picked = pick_source()
    if not picked:
        R.check("ERROR", "Файл-источник",
                "не выбран (отмена) либо не найден - положите ведомость в %s" % IN_DIR)
        return 2
    src, day, do_clean, do_archive, date_origin = picked
    R.check("OK", "Файл-источник", src)
    R.check("INFO", "Дата заказа", "%s (%s)" % (day.strftime("%d.%m.%Y"), date_origin))

    stamp = read_export_stamp(src)
    if stamp:
        age = (datetime.now() - stamp).total_seconds() / 3600.0
        R.check("WARN" if age > STALE_EXPORT_HOURS else "OK",
                "Выгрузка из Торгсофта",
                "%s, это %.0f ч назад" % (stamp.strftime("%d.%m.%Y %H:%M"), age))
    else:
        R.check("INFO", "Выгрузка из Торгсофта", "отметки времени в файле нет")

    fp = source_fingerprint(src)
    prev = history_same_file(fp)
    if prev:
        R.check("WARN", "Повторный расчёт",
                "тот же файл уже считался %s в папку %s" % (prev["run"], prev["day"]))

    off = (day - date.today()).days
    if abs(off) > DATE_SANITY_DAYS:
        R.check("WARN", "Дата заказа",
                "отстоит от сегодняшней на %d дн., проверьте имя файла" % off)

    hr = detect_header_row(src)
    R.check("OK" if hr == HEADER_ROW_HINT else "WARN", "Строка заголовков", hr)

    df = pd.read_excel(src, sheet_name=0, header=hr - 1).dropna(how="all")
    R.check("INFO", "Прочитано строк", len(df))

    c = resolve_columns(df)
    missing = [k for k in ("bc", "sale", "out", "end", "shop") if not c.get(k)]
    if missing:
        R.check("ERROR", "Обязательные столбцы",
                "не найдены: %s; заголовки файла: %s" % (missing, list(df.columns)))
        return 2
    R.check("OK", "Сопоставление столбцов",
            "H='%s', I='%s', L='%s', адрес='%s'" % (c["sale"], c["out"], c["end"], c["shop"]))

    balance_check(df, c)
    w = prepare(df, c)
    if not barcode_check(w):
        return 2
    weight_map = detect_weight(w)

    shops_all = sorted(set(w["shop"]))
    R.check("INFO", "Уникальных адресов", "%d: %s" % (len(shops_all), "; ".join(shops_all)))
    db_kind = file_base(shops_all)
    if db_kind == "ua":
        R.check("INFO", "База", "ЮА: результат в ЗАКАЗЫ\\%s" % out_day_name(day, shops_all))
    elif db_kind == "mixed":
        R.check("WARN", "База",
                "в файле точки и ЮА, и Family: результат в общую папку, ТСД-файлы "
                "баз не разделены; выгружайте ведомости баз отдельно")
    check_route(day, shops_all)

    busy = locked_files(os.path.join(OUT_DIR, out_day_name(day, shops_all)))
    if busy:
        R.check("ERROR", "Файлы прошлого расчёта заняты",
                "закройте в Excel и запустите снова: %s"
                % "; ".join(os.path.basename(b) for b in busy[:5]))
        return 2

    names, collide = {}, []
    for s in shops_all:
        names.setdefault(safe_name(s).lower(), []).append(s)
    for k, v in names.items():
        if len(v) > 1:
            collide.append(" / ".join(v))
    if collide:
        R.check("WARN", "Совпадение имён папок после очистки",
                "; ".join(collide) + " - будут суффиксы (2), (3)")
    else:
        R.check("OK", "Имена файлов/папок", "коллизий нет")

    REF.clear()
    REF.update(load_reference())
    if db_kind == "ua":
        R.check("INFO", "Справочник ЮА",
                "по перекодировке ШК добавлено упаковок: %d"
                % apply_recode_to_ref(REF, load_recode()))

    orders = build_orders(w, weight_map)
    orders = replace_cups_with_packs(orders, w)
    orders = add_bottle_accessories(orders, w)
    orders = drop_excluded(orders)
    if R.errors:
        R.say("Обнаружены критические ошибки, файлы не записаны.")
        write_validation(os.path.join(LOG_DIR, "ВАЛИДАЦИЯ_СБОЙ_%s.xlsx"
                                      % datetime.now().strftime("%Y-%m-%d_%H-%M-%S")),
                         day, src, shops_all, set())
        return 2
    if orders.empty:
        R.check("WARN", "Результат", "по формуле ни одна позиция не дала количество > 0")
        return 0

    orders = apply_dist_list(orders, w, weight_map)
    if dist_on():
        R.check("OK", u"Режим", u"раздача с %s по списку %s"
                % (DIST_SOURCE_SHOP, os.path.basename(DIST_FILE)))
        if orders.empty:
            R.check("WARN", u"Раздача", u"ни одной точке ничего не нужно")
    day_dir = os.path.join(OUT_DIR, out_day_name(day, shops_all))
    if do_clean:
        moved = clear_day_dir(day_dir)
        if moved:
            R.check("INFO", "Прошлый результат этой даты",
                    "убран в %s" % moved)
    os.makedirs(day_dir, exist_ok=True)

    R.say("-" * 70)
    summary, used, verify_fail = [], {}, 0
    for shop, rows in orders.groupby("shop", sort=True):
        base_nm = safe_name(shop)
        nm, k = base_nm, 2
        while nm.lower() in used:
            nm = "%s (%d)" % (base_nm, k)
            k += 1
        used[nm.lower()] = shop

        shop_dir = os.path.join(day_dir, nm)
        os.makedirs(shop_dir, exist_ok=True)
        write_shop_xlsx(os.path.join(shop_dir, nm + ".xlsx"), shop, day, rows)
        write_shop_txt(os.path.join(shop_dir, nm + ".txt"), rows)

        probs = verify_written(shop_dir, nm, rows)
        status = "OK" if not probs else "СБОЙ: " + "; ".join(probs)
        if probs:
            verify_fail += 1
            for p in probs:
                R.problem("Проверка записанных файлов", shop, "", "", p)
        summary.append({"shop": shop, "positions": len(rows),
                        "units": round(float(rows["qty"].sum()), 3),
                        "folder": shop_dir, "verify": status})
        R.say("%-24s позиций: %4d   единиц: %9s   %s"
              % (shop, len(rows), round(float(rows["qty"].sum()), 3),
                 "OK" if not probs else "СБОЙ"))

    if verify_fail:
        R.check("ERROR", "Обратная сверка записанных файлов",
                "расхождения в %d магазинах" % verify_fail)
    else:
        R.check("OK", "Обратная сверка записанных файлов",
                "xlsx и txt перечитаны, совпали с расчётом (%d магазинов)" % len(summary))

    tot_units = round(sum(s["units"] for s in summary), 3)
    if abs(tot_units - round(float(orders["qty"].sum()), 3)) > 0.001:
        R.check("ERROR", "Сверка итогов", "сумма по файлам не равна сумме расчёта")
    else:
        R.check("OK", "Сверка итогов", "всего единиц %s" % tot_units)

    if MAKE_SUPPLIER_ORDER and not dist_on():
        sup = build_supplier_order(orders)
        if not sup.empty:
            sp = os.path.join(day_dir, "_ЗАКАЗ_ПОСТАВЩИКУ.xlsx")
            write_supplier_xlsx(sp, day, sup)
            R.say("Заказ поставщику    : %d позиций -> %s" % (len(sup), sp))

    write_summary(os.path.join(day_dir, "_СВОДКА.xlsx"), day, summary, src)
    write_validation(os.path.join(day_dir, "_ВАЛИДАЦИЯ.xlsx"), day, src,
                     shops_all, set(s["shop"] for s in summary))

    R.say("-" * 70)
    R.say("Магазинов с заказом : %d из %d" % (len(summary), len(shops_all)))
    R.say("Всего позиций       : %d" % sum(s["positions"] for s in summary))
    R.say("Всего единиц        : %s" % tot_units)
    R.say("Ошибок / предупр.   : %d / %d" % (len(R.errors), len(R.warns)))
    R.say("Результат           : %s" % day_dir)
    R.say("Отчёт валидации     : %s" % os.path.join(day_dir, "_ВАЛИДАЦИЯ.xlsx"))
    RUN_INFO.update({"day_dir": day_dir, "day": day,
                     "validation": os.path.join(day_dir, "_ВАЛИДАЦИЯ.xlsx"),
                     "shops": len(summary), "shops_all": len(shops_all),
                     "positions": sum(s["positions"] for s in summary),
                     "units": tot_units,
                     "dist_txt": dist_summary_text(orders)})

    history_write(day, src, fp, len(summary))
    if do_archive and not R.errors:
        moved = archive_source(src)
        if moved:
            R.say("Ведомость в АРХИВ   : %s" % os.path.basename(moved))

    lp = os.path.join(LOG_DIR, "log_%s.txt" % datetime.now().strftime("%Y-%m-%d_%H-%M-%S"))
    with open(lp, "w", encoding="utf-8") as f:
        f.write("\n".join(R.log))
    print("Лог: %s" % lp)
    return 2 if R.errors else 0


# ======================= ОКНО ПРОГРАММЫ ======================

MONTHS_RU = [u"января", u"февраля", u"марта", u"апреля", u"мая", u"июня",
             u"июля", u"августа", u"сентября", u"октября", u"ноября", u"декабря"]


def _res(name):
    """Файл, вшитый в exe (иконка), либо лежащий рядом со скриптом."""
    return os.path.join(getattr(sys, "_MEIPASS", _app_dir()), name)


def _win11_chrome(win, dark=False):
    """Скруглённые углы и цвет заголовка как у обычных окон Windows 11."""
    try:
        import ctypes
        win.update_idletasks()
        h = ctypes.windll.user32.GetParent(win.winfo_id())
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            h, 33, ctypes.byref(ctypes.c_int(2)), 4)          # скругление
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            h, 20, ctypes.byref(ctypes.c_int(1 if dark else 0)), 4)
    except Exception:
        pass


def run_app():
    """Окно программы: выбор файла, ход расчёта и результат внутри окна."""
    global GUI_CHOICE
    import threading, queue
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    app = tk.Tk()
    app.withdraw()
    app.title(u"Расчёт заказа перемещений")
    try:
        app.iconbitmap(_res("raschet.ico"))
    except Exception:
        pass

    modern = False
    try:
        import sv_ttk                     # тема Windows 11 (Fluent)
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
    H1 = ("Segoe UI Semibold", 17)
    H2 = ("Segoe UI Semibold", 11)
    BOLD = ("Segoe UI Semibold", 11)
    FS = ("Segoe UI", 9)
    ACC = "Accent.TButton" if modern else "Go.TButton"
    if not modern:
        S.configure("Go.TButton", background="#2f6fed", foreground="#ffffff",
                    font=("Segoe UI Semibold", 10), borderwidth=0, padding=(20, 9))
        S.map("Go.TButton", background=[("active", "#1f5bd0"),
                                        ("disabled", "#a9bde8")])
        S.configure("Card.TFrame", background="#ffffff", relief="solid",
                    borderwidth=1)
    S.configure("Mut.TLabel", foreground=MUT, font=FS)
    S.configure("Red.TLabel", foreground=RED, font=FS)
    S.configure("Grn.TLabel", foreground=GRN, font=FS)
    S.configure("Amb.TLabel", foreground=AMB, font=FS)
    S.configure("H2.TLabel", font=H2)
    S.configure("Big.TLabel", font=BOLD)

    root = ttk.Frame(app, padding=(22, 18, 22, 0))
    root.pack(fill="both", expand=True)
    ttk.Label(root, text=u"Расчёт заказа перемещений", font=H1).pack(anchor="w")
    ttk.Label(root, text=u"Сколько товара везти с РЦ в магазины",
              style="Mut.TLabel").pack(anchor="w", pady=(2, 0))

    setup = ttk.Frame(root)
    work  = ttk.Frame(root)
    setup.pack(fill="both", expand=True, pady=(16, 0))

    def card(parent):
        c = ttk.Frame(parent, style="Card.TFrame", padding=16)
        c.pack(fill="x", pady=(0, 12))
        return c

    # ---------------- откуда берём данные ----------------
    c1 = card(setup)
    top1 = ttk.Frame(c1)
    top1.pack(fill="x")
    ttk.Label(top1, text=u"Откуда берём данные", style="H2.TLabel").pack(side="left")
    ttk.Button(top1, text=u"Выбрать файл",
               command=lambda: choose()).pack(side="right")
    fname = ttk.Label(c1, text="", style="Big.TLabel")
    fname.pack(anchor="w", pady=(10, 0))
    fdir = ttk.Label(c1, text="", style="Mut.TLabel")
    fdir.pack(anchor="w")
    fstamp = ttk.Label(c1, text="", style="Mut.TLabel")
    fstamp.pack(anchor="w", pady=(8, 0))
    fprev = ttk.Label(c1, text="", style="Red.TLabel", wraplength=560,
                      justify="left")
    fprev.pack(anchor="w")

    # ---------------- на какой день везём ----------------
    c2 = card(setup)
    ttk.Label(c2, text=u"На какой день везём", style="H2.TLabel").pack(anchor="w")
    drow = ttk.Frame(c2)
    drow.pack(anchor="w", pady=(10, 0))
    dv = tk.StringVar()
    ttk.Entry(drow, textvariable=dv, width=12, justify="center",
              font=("Segoe UI Semibold", 13)).pack(side="left", ipady=4)
    dorig = ttk.Label(drow, text="", style="Mut.TLabel")
    dorig.pack(side="left", padx=10)
    route_l = ttk.Label(c2, text="", style="Mut.TLabel", wraplength=560,
                        justify="left")
    route_l.pack(anchor="w", pady=(8, 0))
    warn_l = ttk.Label(c2, text="", style="Amb.TLabel", wraplength=560,
                       justify="left")
    warn_l.pack(anchor="w")

    # ---------------- что сделать после ----------------
    # ---------------- режим расчёта ----------------
    mode = card(setup)
    ttk.Label(mode, text=u"Что считаем", style="H2.TLabel").pack(anchor="w")
    mode_v, distv = tk.IntVar(value=0), tk.StringVar()

    def dist_pick():
        f = filedialog.askopenfilename(
            parent=app, title=u"Список позиций для раздачи с Полевой",
            initialdir=(os.path.dirname(distv.get()) if distv.get() else BASE_DIR),
            filetypes=[(u"Вывоз и списки", "*.xlsx *.txt *.csv"),
                       (u"Все файлы", "*.*")])
        if f:
            distv.set(os.path.abspath(f))
        mode_sync()

    def dist_pick_dir():
        d = filedialog.askdirectory(
            parent=app, title=u"Папка ВЫВОЗ_<дата> со скрипта вывоза",
            initialdir=(distv.get() if os.path.isdir(distv.get() or "")
                        else BASE_DIR))
        if d:
            distv.set(os.path.abspath(d))
        mode_sync()

    def mode_sync():
        if not mode_v.get():
            dlbl.config(text=u"получатели и куст определяются по ведомости",
                        style="Mut.TLabel")
            dbtn.pack_forget()
            return
        dbtn.pack(anchor="w", padx=(24, 0), pady=(6, 0))
        p = distv.get()
        if not p:
            dlbl.config(text=u"укажите, что раздаём: файл Вывоз_<дата>.xlsx, "
                             u"папку ВЫВОЗ_<дата> или список ШК;количество",
                        style="Mut.TLabel")
            return
        dlbl.config(text=u"%s · %d позиций · %s исключена из получателей\n"
                         u"результат ляжет в ЗАКАЗЫ\\<дата>%s, "
                         u"обычный заказ не затрагивается"
                         % (os.path.basename(p.rstrip("\\/")), dist_count(p),
                            DIST_SOURCE_SHOP, DIST_DIR_SUFFIX),
                    style="Grn.TLabel")

    ttk.Radiobutton(mode, variable=mode_v, value=0, command=mode_sync,
                    text=u"Обычный заказ · РЦ → магазины"
                    ).pack(anchor="w", pady=(10, 2))
    ttk.Radiobutton(mode, variable=mode_v, value=1, command=mode_sync,
                    text=u"Раздача с Полевой по списку"
                    ).pack(anchor="w", pady=(0, 2))
    dlbl = ttk.Label(mode, text=u"", style="Mut.TLabel", justify="left",
                     wraplength=560)
    dlbl.pack(anchor="w", padx=(24, 0))
    dbtn = ttk.Frame(mode)
    ttk.Button(dbtn, text=u"Файл вывоза / список", padding=(10, 4),
               command=lambda: dist_pick()).pack(side="left")
    ttk.Button(dbtn, text=u"Папка ВЫВОЗ", padding=(10, 4),
               command=lambda: dist_pick_dir()).pack(side="left", padx=(8, 0))
    mode_sync()

    opt = ttk.Frame(setup)
    opt.pack(fill="x", pady=(2, 0))
    clean_v, arch_v = tk.IntVar(value=1), tk.IntVar(value=1)
    ttk.Checkbutton(opt, variable=clean_v,
                    text=u"прежний расчёт за этот день убрать в папку замен"
                    ).pack(anchor="w", pady=2)
    ttk.Checkbutton(opt, variable=arch_v,
                    text=u"ведомость после расчёта убрать в АРХИВ"
                    ).pack(anchor="w", pady=2)

    foot = ttk.Frame(root, padding=(0, 16, 0, 18))
    foot.pack(fill="x", side="bottom")
    go_b = ttk.Button(foot, text=u"Рассчитать заказ", style=ACC, padding=(18, 8))
    go_b.pack(side="right")
    ttk.Button(foot, text=u"Закрыть", command=app.destroy,
               padding=(14, 8)).pack(side="right", padx=(0, 10))

    srcv = tk.StringVar()
    origin = [""]

    def parse_dv():
        t = dv.get().strip().replace("/", ".").replace("-", ".")
        m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$", t)
        if not m:
            return None
        d, mo, y = (int(g) for g in m.groups())
        y = y + 2000 if y < 100 else y
        try:
            return date(y, mo, d)
        except ValueError:
            return None

    def human(d):
        return u"%d %s" % (d.day, MONTHS_RU[d.month - 1])

    def refresh(*_a):
        d = parse_dv()
        go_b.state(["!disabled"] if (srcv.get() and d) else ["disabled"])
        if d is None:
            route_l.config(text=u"")
            warn_l.config(text=u"Дата не понята. Нужен вид 07.09.2026",
                          style="Red.TLabel")
            return
        rk = ROUTE_BY_DOW.get(d.weekday())
        if rk:
            route_l.config(
                text=u"%s, %s. В этот день возим %d магазинов (группа %s)"
                     % (DOW_RU[d.weekday()].capitalize(), human(d),
                        len(ROUTES[rk]), rk))
        else:
            route_l.config(text=u"Воскресенье — по графику отгрузки нет")
        msgs = []
        _d, n, ts = day_dir_state(d)
        if n:
            msgs.append(u"За %s расчёт уже делали %s — %d магазинов. "
                        u"Прежний уберём в папку замен."
                        % (human(d), ts.strftime("%d.%m в %H:%M") if ts else "?", n))
        off = (d - date.today()).days
        if abs(off) > DATE_SANITY_DAYS:
            msgs.append(u"Это %s от сегодняшнего дня — проверьте, тот ли файл."
                        % (u"на %d дн. вперёд" % off if off > 0
                           else u"на %d дн. назад" % -off))
        warn_l.config(text=u"\n".join(msgs), style="Amb.TLabel")

    def load(path):
        srcv.set(path or "")
        if not path:
            fname.config(text=u"файл не выбран")
            fdir.config(text=u"")
            fstamp.config(text=u"Нажмите «Выбрать файл»", style="Mut.TLabel")
            fprev.config(text=u"")
            dv.set("")
            refresh()
            return
        fname.config(text=os.path.basename(path))
        fdir.config(text=os.path.dirname(path))
        g, o = guess_day(path)
        origin[0] = o
        dv.set(g.strftime("%d.%m.%Y"))
        dorig.config(text={u"из имени файла": u"взято из названия файла"}.get(o, o))
        stamp = read_export_stamp(path)
        if stamp:
            age = (datetime.now() - stamp).total_seconds() / 3600.0
            when = (u"%.0f ч назад" % age if age < 48
                    else u"%.0f дн. назад" % (age / 24.0))
            fstamp.config(
                text=u"Выгружено из Торгсофта %s, это %s"
                     % (stamp.strftime("%d.%m.%Y в %H:%M"), when),
                style="Amb.TLabel" if age > STALE_EXPORT_HOURS else "Grn.TLabel")
        else:
            fstamp.config(text=u"В файле не записано, когда его выгрузили",
                          style="Mut.TLabel")
        pv = history_same_file(source_fingerprint(path))
        fprev.config(text=(u"Этот же файл уже считали %s — заказ лёг в папку %s"
                           % (pv["run"], pv["day"])) if pv else u"")
        refresh()

    def choose():
        cur = srcv.get()
        f = filedialog.askopenfilename(
            parent=app, title=u"Выберите оборотную ведомость из Торгсофта",
            initialdir=(os.path.dirname(cur) if cur
                        else (IN_DIR if os.path.isdir(IN_DIR) else BASE_DIR)),
            filetypes=[(u"Книги Excel", "*.xlsx *.xlsm"), (u"Все файлы", "*.*")])
        if f:
            load(os.path.abspath(f))

    dv.trace_add("write", refresh)

    # ---------------- ход расчёта ----------------
    wcap = ttk.Label(work, text=u"", font=("Segoe UI Semibold", 14))
    wcap.pack(anchor="w")
    wsub = ttk.Label(work, text=u"", style="Mut.TLabel", justify="left")
    wsub.pack(anchor="w", pady=(4, 0))
    bar = ttk.Progressbar(work, mode="indeterminate")
    bar.pack(fill="x", pady=14)
    tw = ttk.Frame(work)
    tw.pack(fill="both", expand=True, pady=(0, 4))
    sb = ttk.Scrollbar(tw)
    sb.pack(side="right", fill="y")
    txt = tk.Text(tw, bg="#1b1b1b", fg="#d6d6d6", font=("Cascadia Mono", 9),
                  bd=0, padx=12, pady=10, wrap="none", yscrollcommand=sb.set,
                  height=12, highlightthickness=0)
    txt.pack(fill="both", expand=True)
    sb.config(command=txt.yview)
    txt.configure(state="disabled")
    for tag, col in (("ERROR", "#ff6b6b"), ("WARN", "#ffc266"),
                     ("OK", "#7fd18c"), ("INFO", "#8fb6e8")):
        txt.tag_configure(tag, foreground=col)

    def put(line):
        tag = ("ERROR" if line.startswith("[ОШИБ") else
               "WARN" if line.startswith("[ВНИМ") else
               "OK" if line.startswith("[ OK ") else
               "INFO" if line.startswith("[ИНФО") else "")
        txt.configure(state="normal")
        txt.insert("end", line + "\n", tag)
        txt.see("end")
        txt.configure(state="disabled")

    def open_path(pth):
        try:
            os.startfile(pth)
        except Exception:
            messagebox.showinfo(u"Путь", pth, parent=app)

    running = [False]

    def start():
        global GUI_CHOICE, DIST_FILE
        d = parse_dv()
        if running[0] or not srcv.get() or d is None:
            return
        running[0] = True
        DIST_FILE = distv.get() if mode_v.get() else ""
        GUI_CHOICE = (srcv.get(), d, bool(clean_v.get()), bool(arch_v.get()),
                      origin[0])
        setup.pack_forget()
        for w_ in foot.winfo_children():
            w_.destroy()
        work.pack(fill="both", expand=True, pady=(16, 0))
        wcap.config(text=u"Считаю…")
        wsub.config(text=u"%s · везём на %s"
                         % (os.path.basename(srcv.get()), human(d)))
        bar.start(12)
        txt.configure(state="normal")
        txt.delete("1.0", "end")
        txt.configure(state="disabled")
        app.geometry("720x680")

        R.checks, R.problems, R.anomalies, R.log = [], [], [], []
        RUN_INFO.clear()
        q = queue.Queue()

        def say(m):
            R.log.append(str(m))
            q.put(str(m))
        R.say = say

        def worker():
            try:
                code = main()
            except Exception as e:
                import traceback
                q.put(u"СБОЙ: %s: %s" % (type(e).__name__, e))
                for ln in traceback.format_exc().splitlines():
                    q.put(ln)
                code = 2
            q.put(("done", code))

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            done = None
            while True:
                try:
                    it = q.get_nowait()
                except queue.Empty:
                    break
                if isinstance(it, tuple):
                    done = it[1]
                else:
                    put(it)
            if done is None:
                app.after(70, poll)
            else:
                finish(done)
        app.after(70, poll)

    def finish(code):
        bar.stop()
        bar.pack_forget()
        ne, nw = len(R.errors), len(R.warns)
        i = RUN_INFO
        if code == 0 and not ne and i:
            wcap.config(text=u"Готово", foreground=GRN)
            wsub.config(text=u"%d магазинов из %d · %d позиций · %s единиц\n"
                             u"Предупреждений %d, ошибок нет\n%s"
                             % (i["shops"], i["shops_all"], i["positions"],
                                i["units"], nw, i["day_dir"]) + i.get("dist_txt", u""))
        else:
            wcap.config(text=u"Файлы не записаны", foreground=RED)
            wsub.config(text=u"Ошибок %d, предупреждений %d. Что именно не так — "
                             u"в списке выше и в отчёте." % (ne, nw))
        f2 = ttk.Frame(root, padding=(0, 12, 0, 18))
        f2.pack(fill="x", side="bottom")
        if i:
            ttk.Button(f2, text=u"Открыть папку с заказами", style=ACC,
                       padding=(18, 8),
                       command=lambda: open_path(i["day_dir"])).pack(side="right")
            ttk.Button(f2, text=u"Отчёт проверки", padding=(14, 8),
                       command=lambda: open_path(i["validation"])).pack(
                           side="right", padx=(0, 10))
        else:
            ttk.Button(f2, text=u"Открыть папку ЛОГИ", style=ACC,
                       padding=(18, 8),
                       command=lambda: open_path(LOG_DIR)).pack(side="right")
        ttk.Button(f2, text=u"Закрыть", command=app.destroy,
                   padding=(14, 8)).pack(side="right", padx=(0, 10))

    go_b.config(command=start)
    app.bind("<Return>", lambda e: start())
    ensure_dirs()
    load(find_source_file(quiet=True))
    app.update_idletasks()
    w_, h_ = 720, max(600, app.winfo_reqheight())
    sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
    app.geometry("%dx%d+%d+%d" % (w_, h_, max(0, (sw - w_) // 2),
                                  max(0, (sh - h_) // 3)))
    app.minsize(660, 560)
    app.deiconify()
    _win11_chrome(app)
    app.lift()
    app.focus_force()
    app.mainloop()
    return 0


if __name__ == "__main__":
    _args = [a.lower().lstrip("-/") for a in sys.argv[1:]]
    if USE_GUI and "auto" not in _args and "a" not in _args:
        try:
            sys.exit(run_app())
        except Exception as _e:
            try:
                import tkinter.messagebox as _mb
                _mb.showerror("Расчёт заказа",
                              "Окно не открылось:\n%s: %s\n\n"
                              "Запустите с ключом --auto для работы без окна."
                              % (type(_e).__name__, _e))
            except Exception:
                print("СБОЙ: %s: %s" % (type(_e).__name__, _e))
            sys.exit(2)
    try:
        code = main()
    except Exception as e:
        print("СБОЙ: %s: %s" % (type(e).__name__, e))
        import traceback
        traceback.print_exc()
        code = 2
    try:
        input("\nГотово. Нажмите Enter для выхода...")
    except (EOFError, RuntimeError):
        pass
    sys.exit(code)
