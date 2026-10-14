# Design Tokens — Фемелі Маркет (Modern Enterprise Dashboard)

Основа — Modern Enterprise Design System из DESIGN.md. Цвета и шрифты заменены на фирменные
Фемелі Маркет (проект fm-peremeshcheniya, окно «Вывоз вне матрицы»). Use these exact values
in all Python token constants. Имена токенов сохранены, поэтому шаблоны из qss_patterns.md
и component_library.md работают без переименований.

## Правила проекта (важнее общих правил скилла)

1. **Фирменный красный `#E31E24` только в двух местах:** кнопка главного действия
   (`QPushButton[class="primary"]`) и полоска 3 px у активного пункта бокового меню
   (у активной вкладки — подчёркивание). Красным не красим: выделение строк, галочки,
   радиокнопки, рамку фокуса, прогресс, текст.
2. **Выделение нейтральное:** `COLOR_SELECTION_BG` + `COLOR_SELECTION_TEXT`. Красная строка
   выглядит как ошибка.
3. **Ошибка — свой красный** (`COLOR_ERROR`, темнее фирменного) и всегда со значком или словом.
   Цвет никогда не единственный признак состояния.
4. **Боковая панель тёмная в обеих темах.** На ней красным пишем только логотип (крупный текст);
   мелкий красный текст на панели не проходит контраст (2.97:1).
5. **Шрифт один — Inter** (и заголовки, и текст). Hanken Grotesk 3.013 проверен через fontTools:
   кириллицы в нём нет совсем (ни А–я, ни і ї є ґ І Ї Є Ґ). Смешивать шрифты в одном слове нельзя.
6. **Масштаб Windows 125–150 %.** Размеры ниже — логические пиксели. Qt 6 сам умножает их
   на масштаб системы (high-DPI всегда включён, округление PassThrough по умолчанию).
   Не вызывать `SetProcessDpiAwareness`, не пересчитывать размеры вручную, не задавать
   `setFixedSize` у элементов с текстом. Проверка в облаке: `QT_QPA_PLATFORM=offscreen`
   и `QT_SCALE_FACTOR=1.25` / `1.5`, затем `widget.grab()` в PNG.
7. **Тема по умолчанию светлая.** Переключатель в «Сервисе»: Светлая / Тёмная / Как в Windows,
   выбор сохраняется (QSettings).

Контраст всех пар текст/фон ниже проверен по формуле WCAG 2.x: не ниже 4.5:1.

---

## Color Palette

### Light Mode (Primary)
```python
# Surfaces
COLOR_BACKGROUND           = "#f6f6f7"   # Холст окна (как в текущем окне tkinter)
COLOR_SURFACE_LOWEST       = "#ffffff"   # Карточки, строки таблиц, поля ввода
COLOR_SURFACE_LOW          = "#f1f1f3"   # Шапка таблицы, чередование строк, disabled-фон
COLOR_SURFACE              = "#ececee"   # Hover вторичной кнопки
COLOR_SURFACE_HIGH         = "#e4e4e7"   # Hover пунктов списка
COLOR_SURFACE_HIGHEST      = "#dcdce1"
COLOR_SURFACE_BRIGHT       = "#f6f6f7"
COLOR_SURFACE_DIM          = "#dcdce1"
COLOR_SURFACE_VARIANT      = "#e4e4e7"

# Content
COLOR_ON_SURFACE           = "#1f2023"   # Основной текст            15.1:1 на фоне
COLOR_ON_SURFACE_VARIANT   = "#5f5f66"   # Вторичный текст, подписи   5.9:1 на фоне
COLOR_OUTLINE              = "#8e8e94"   # Рамки галочек и полей      3.3:1 (UI, норма 3:1)
COLOR_OUTLINE_VARIANT      = "#dcdce1"   # Тонкие разделители, рамки карточек

# Primary — фирменный красный. ТОЛЬКО кнопка действия и полоска активного пункта
COLOR_PRIMARY              = "#e31e24"   # белый текст на нём 4.69:1
COLOR_ON_PRIMARY           = "#ffffff"
COLOR_PRIMARY_CONTAINER    = "#b8161b"   # hover кнопки действия (6.6:1 с белым)
COLOR_PRIMARY_PRESSED      = "#9c1217"   # pressed кнопки действия (8.4:1 с белым)
COLOR_ON_PRIMARY_CONTAINER = "#ffffff"
COLOR_INVERSE_PRIMARY      = "#ff6b70"
COLOR_SURFACE_TINT         = "#e31e24"
# В шаблонах скилла PRIMARY_FIXED — фон выделения и подложка значков: делаем нейтральным
COLOR_PRIMARY_FIXED        = "#e4e4e7"
COLOR_PRIMARY_FIXED_DIM    = "#dcdce1"

# Нейтральные состояния (вместо красного)
COLOR_SELECTION_BG         = "#e4e4e7"   # выделенная строка / пункт
COLOR_SELECTION_TEXT       = "#1f2023"   # 12.8:1 на COLOR_SELECTION_BG
COLOR_CONTROL_CHECKED      = "#3f3f46"   # отмеченные галочка и радио, ползунок, прогресс
COLOR_ON_CONTROL_CHECKED   = "#ffffff"
COLOR_FOCUS                = "#3f3f46"   # рамка фокуса 2px (клавиатура)

# Secondary
COLOR_SECONDARY            = "#52525b"
COLOR_ON_SECONDARY         = "#ffffff"
COLOR_SECONDARY_CONTAINER  = "#e4e4e7"
COLOR_ON_SECONDARY_CONTAINER = "#3f3f46"

# Tertiary (не используется; оставлено для совместимости шаблонов)
COLOR_TERTIARY             = "#52525b"
COLOR_ON_TERTIARY          = "#ffffff"
COLOR_TERTIARY_CONTAINER   = "#e4e4e7"
COLOR_ON_TERTIARY_CONTAINER = "#3f3f46"

# Error — НЕ фирменный красный, всегда со значком (error e000 / warning e002) или словом
COLOR_ERROR                = "#b42318"   # 6.1:1 на фоне
COLOR_ON_ERROR             = "#ffffff"
COLOR_ERROR_CONTAINER      = "#fee4e2"
COLOR_ON_ERROR_CONTAINER   = "#912018"   # 7.2:1

# Inverse (тултипы)
COLOR_INVERSE_SURFACE      = "#2b2c30"
COLOR_INVERSE_ON_SURFACE   = "#ededef"
```

### Sidebar (тёмная в обеих темах)
```python
SIDEBAR_BG                 = "#2b2c30"   # FM_GRAPH из кода
SIDEBAR_HOVER              = "#34353a"
SIDEBAR_ACTIVE             = "#3d3e44"
SIDEBAR_TEXT               = "#b4b4ba"   # 6.8:1
SIDEBAR_TEXT_ACTIVE        = "#ffffff"   # 10.7:1 на SIDEBAR_ACTIVE
SIDEBAR_GROUP_LABEL        = "#a1a1aa"   # заголовки групп «Сегодня», ... 5.4:1 (было #6f7077 - 2.8:1)
SIDEBAR_INDICATOR          = "#e31e24"   # полоска 3px у активного пункта; не единственный признак:
                                         # у активного пункта ещё фон SIDEBAR_ACTIVE и белый текст
SIDEBAR_DIVIDER            = "#3a3b40"
DARK_SIDEBAR_BG            = "#111113"   # в тёмной теме панель темнее холста
DARK_SIDEBAR_HOVER         = "#1c1c1f"
DARK_SIDEBAR_ACTIVE        = "#27272a"
```

### Dark Mode Overrides
```python
DARK_BACKGROUND            = "#18181b"
DARK_SURFACE               = "#202024"   # карточки, строки таблиц
DARK_SURFACE_CONTAINER     = "#27272a"
DARK_SURFACE_HIGH          = "#2f2f34"
DARK_BORDER                = "#3a3a40"
DARK_ON_SURFACE            = "#ededef"   # 13.9:1 на DARK_SURFACE
DARK_ON_SURFACE_VARIANT    = "#a1a1aa"   # 6.3:1
DARK_OUTLINE               = "#71717a"
DARK_PRIMARY               = "#e31e24"   # кнопка действия та же: белый текст 4.69:1
DARK_PRIMARY_CONTAINER     = "#b8161b"
DARK_PRIMARY_TEXT          = "#ff6b70"   # если красный нужен линией/текстом на тёмном: 5.9:1
DARK_SELECTION_BG          = "#3a3a40"
DARK_SELECTION_TEXT        = "#ededef"   # 9.7:1
DARK_CONTROL_CHECKED       = "#d4d4d8"
DARK_ON_CONTROL_CHECKED    = "#18181b"
DARK_FOCUS                 = "#d4d4d8"
DARK_ERROR                 = "#ff8a80"   # 7.1:1
DARK_SUCCESS               = "#4ade80"   # 9.3:1
DARK_WARNING               = "#fbbf24"   # 9.7:1

# Semantic (light). Всегда вместе со значком или словом
COLOR_SUCCESS              = "#15803d"   # 4.6:1 на фоне
COLOR_SUCCESS_BG           = "#dcfce7"
COLOR_WARNING              = "#a15c07"   # 4.8:1 на фоне
COLOR_WARNING_BG           = "#fef3c7"
COLOR_DANGER               = "#b42318"
COLOR_DANGER_BG            = "#fee4e2"
```

---

## Typography Scale

```python
# Font families. Hanken Grotesk без кириллицы (проверено fontTools) -> везде Inter
FONT_HEADING  = "Inter"            # Заголовки (SemiBold 600)
FONT_BODY     = "Inter"            # Текст, подписи, таблицы
FONT_ICONS    = "Material Symbols Outlined"
# Файлы: Перемещения ЮА/assets/fonts/ (Inter[opsz,wght].ttf, MaterialSymbolsOutlined[...].ttf, OFL.txt)
# Загрузка: QFontDatabase.addApplicationFont; папка = getattr(sys, "_MEIPASS", папка скрипта)
# В exe: datas += [("assets/fonts", "assets/fonts")] в vyvoz_vne_matricy.spec
# Не загрузился -> "Segoe UI" и запись в журнал

# Scale  (size_px, weight, line_height_px, letter_spacing_em) — логические px, Qt масштабирует сам
TYPO_HEADLINE_LG = (24, 600, 32, -0.01)   # Заголовок экрана
TYPO_HEADLINE_MD = (20, 600, 28, -0.01)   # Заголовок карточки / диалога
TYPO_HEADLINE_SM = (16, 600, 24,  0.00)   # Подзаголовок, шаг
TYPO_BODY_LG     = (15, 400, 22,  0.00)
TYPO_BODY_MD     = (14, 400, 20,  0.00)   # ОСНОВНОЙ текст: формы, кнопки, меню, таблицы
TYPO_BODY_SM     = (13, 400, 18,  0.00)   # Только вторичное: подписи под полями, время
TYPO_LABEL_MD    = (12, 600, 16, +0.05)   # Заголовки колонок
TYPO_LABEL_SM    = (11, 500, 14,  0.00)   # Чипы статуса, заголовки групп меню
```

**Rules:**
- Основной текст 14 px (решение владельца: настольная программа с плотными таблицами;
  крупнее — через масштаб Windows, окно обязано масштабироваться без обрезки)
- На одном экране не больше четырёх размеров
- Ячейки таблиц: `TYPO_BODY_MD` (14 px), заголовки колонок: `TYPO_LABEL_MD`
- Пункты меню: `TYPO_BODY_MD` Medium; заголовки групп меню: `TYPO_LABEL_SM` заглавными
- Текст кнопок: `TYPO_BODY_MD` SemiBold (заглавные в кнопках не используем: русский текст длинный)
- Длинные пояснения не выводим серым текстом: tooltip или кнопка «?»

---

## Spacing (4px base grid)

```python
SPACING_UNIT = 4    # px — never deviate from multiples of 4
SPACING_XS   = 4    # px  — tight gaps, icon padding
SPACING_SM   = 8    # px  — button padding (vertical), item gaps
SPACING_MD   = 16   # px  — button padding (horizontal), card padding
SPACING_LG   = 24   # px  — section gaps, container padding
SPACING_XL   = 32   # px  — large section separators
SPACING_CONTAINER_MARGIN = 24  # Main content area padding
SPACING_GUTTER           = 16  # Column gap in multi-column layouts
```

---

## Border Radius

```python
RADIUS_SM      = 2    # px  — inputs, chips, feedback toasts
RADIUS_DEFAULT = 4    # px  — buttons, most small components
RADIUS_MD      = 6    # px
RADIUS_LG      = 8    # px  — cards, modals, containers
RADIUS_XL      = 12   # px
RADIUS_FULL    = 9999 # px  — pills, avatars
```

---

## Elevation / Depth

| Level | Surface | Border | Shadow | Usage |
|-------|---------|--------|--------|-------|
| 0 | `COLOR_BACKGROUND` | none | none | Canvas / page background |
| 1 | `COLOR_SURFACE_LOWEST` | 1px `COLOR_OUTLINE_VARIANT` | none | Cards, table rows, inputs |
| 2 | `COLOR_SURFACE_LOWEST` | 1px `COLOR_OUTLINE_VARIANT` | `blur=12, rgba(0,0,0,0.08)` | Dropdowns, modals, floating panels |
| Active | any | 2px `COLOR_PRIMARY` | none | Focused inputs, active sidebar items |

**Rule:** Prefer tonal backgrounds (`surface-container-*` variants) + 1px borders for
hierarchy. Reserve `QGraphicsDropShadowEffect` for Level 2 only.

---

## Layout Structure

```
┌──────────────────────────────────────────────────────────┐
│ Sidebar (240px fixed, collapsible to 64px icon-only)     │
│ TopBar  (64px fixed height, full remaining width)         │
│ Content Area (fluid, 24px padding all sides, max 1440px) │
└──────────────────────────────────────────────────────────┘
```

- **Sidebar width:** 240px expanded, 64px collapsed
- **Top bar height:** 64px
- **Content padding:** `SPACING_LG` (24px) all sides
- **Table row heights:** Compact = 32px, Standard = 40px
- **Sidebar item padding:** 8px vertical, 16px horizontal

---

## Component Shape Rules

| Component | Radius | Notes |
|-----------|--------|-------|
| Primary / secondary buttons | `RADIUS_DEFAULT` (4px) | |
| Ghost buttons | `RADIUS_DEFAULT` (4px) | |
| Text inputs | `RADIUS_SM` (2px) | Hair-line 1px border |
| Cards | `RADIUS_LG` (8px) | 1px border, no shadow |
| Modals / dialogs | `RADIUS_LG` (8px) | Level-2 shadow |
| Chips / badges | `RADIUS_FULL` (9999px) | |
| Toasts | `RADIUS_SM` (2px) | |
| Active sidebar indicator | 0px left edge, 3px right curve | Left-anchored bar |
| Avatars | `RADIUS_FULL` (9999px) | |