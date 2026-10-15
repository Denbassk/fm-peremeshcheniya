"""Токены дизайна Фемелі Маркет (.claude/skills/pyqt6-ui-designer/references/design_tokens.md).

Красный #E31E24 - только кнопка главного действия и полоска активного пункта меню. Выделение, галочки, фокус
нейтральные; ошибка - свой красный и всегда со значком. Шрифт - Inter (у Hanken Grotesk нет кириллицы).
Размеры - логические px: Qt 6 сам умножает их на масштаб Windows (125 %, 150 %).
"""

FONT = "Inter"
FONT_FALLBACK = "Segoe UI"         # если Inter не загрузился или выбран в «Сервисе»
FONT_MODES = (("inter", u"Inter"), ("segoe", u"Segoe UI"))
FONT_ICONS = "Material Symbols Outlined"
# размер шрифта в «Сервисе» (решение 15.10.2026: 14 px при масштабе Windows 125-150 % слишком крупно): код, подпись, основной px
FONT_SIZES = (("s", u"Мелкий", 12), ("m", u"Средний", 13), ("l", u"Крупный", 14))
FONT_SIZE_DEFAULT = "m"

SPACING_XS, SPACING_SM, SPACING_MD, SPACING_LG, SPACING_XL = 4, 8, 16, 24, 32
RADIUS_SM, RADIUS_DEFAULT, RADIUS_LG, RADIUS_FULL = 2, 4, 8, 9999

# (size_px, weight)
TYPO_HEADLINE_LG = (24, 600)    # заголовок экрана
TYPO_HEADLINE_SM = (16, 600)    # заголовок карточки / шага
TYPO_BODY_MD = (14, 400)        # основной текст, таблицы, кнопки
TYPO_BODY_SM = (13, 400)        # вторичное: подписи, время
TYPO_LABEL_MD = (12, 600)       # заголовки колонок
TYPO_LABEL_SM = (11, 600)       # заголовки групп меню, чипы

SIDEBAR_W = 240
ROW_H = 32                      # компактная строка таблицы
CONTROL_H = 36                  # высота кнопок и полей

BRAND_RED = "#e31e24"

LIGHT = {
    "bg": "#f6f6f7", "surface": "#ffffff", "surface_low": "#f1f1f3", "surface_mid": "#ececee", "surface_high": "#e4e4e7",
    "text": "#1f2023", "text_muted": "#5f5f66", "outline": "#8e8e94", "outline_variant": "#dcdce1",
    "primary": BRAND_RED, "on_primary": "#ffffff", "primary_hover": "#b8161b", "primary_pressed": "#9c1217",
    "primary_disabled_bg": "#dcdce1", "primary_disabled_text": "#5f5f66",
    "selection_bg": "#e4e4e7", "selection_text": "#1f2023",
    "checked": "#3f3f46", "on_checked": "#ffffff", "focus": "#3f3f46",
    "error": "#b42318", "error_bg": "#fee4e2", "on_error_bg": "#912018",
    "success": "#15803d", "success_bg": "#dcfce7",
    "warning": "#a15c07", "warning_bg": "#fef3c7",
    "info": "#5f5f66", "info_bg": "#ececee",
    "tooltip_bg": "#2b2c30", "tooltip_text": "#ededef",
    # боковая панель - тёмная
    "side_bg": "#2b2c30", "side_hover": "#34353a", "side_active": "#3d3e44", "side_text": "#b4b4ba",
    "side_text_active": "#ffffff", "side_group": "#a1a1aa", "side_divider": "#3a3b40", "side_indicator": BRAND_RED,
}

DARK = dict(LIGHT)
DARK.update({
    "bg": "#18181b", "surface": "#202024", "surface_low": "#27272a", "surface_mid": "#2f2f34", "surface_high": "#3a3a40",
    "text": "#ededef", "text_muted": "#a1a1aa", "outline": "#71717a", "outline_variant": "#3a3a40",
    "primary_disabled_bg": "#3a3a40", "primary_disabled_text": "#a1a1aa",
    "selection_bg": "#3a3a40", "selection_text": "#ededef",
    "checked": "#d4d4d8", "on_checked": "#18181b", "focus": "#d4d4d8",
    "error": "#ff8a80", "error_bg": "#3b1d1b", "on_error_bg": "#ffb4ab",
    "success": "#4ade80", "success_bg": "#16301f",
    "warning": "#fbbf24", "warning_bg": "#33270f",
    "info": "#a1a1aa", "info_bg": "#2f2f34",
    "tooltip_bg": "#ededef", "tooltip_text": "#18181b",
    "side_bg": "#111113", "side_hover": "#1c1c1f", "side_active": "#27272a",
})

# уровень строки ядра (OK / WARN / ERR / ERROR / INFO) -> (значок, токен цвета, слово для чтения без цвета)
LEVELS = {
    "OK": ("check_circle", "success", u"норма"),
    "WARN": ("warning", "warning", u"внимание"),
    "ERR": ("error", "error", u"ошибка"),
    "ERROR": ("error", "error", u"ошибка"),
    "INFO": ("info", "info", u"инфо"),
}
