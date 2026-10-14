# Design Tokens — Modern Enterprise Dashboard

Extracted from DESIGN.md. Use these exact values in all Python token constants.

---

## Color Palette

### Light Mode (Primary)
```python
# Surfaces
COLOR_SURFACE              = "#faf9ff"
COLOR_SURFACE_DIM          = "#ccdaff"
COLOR_SURFACE_BRIGHT       = "#faf9ff"
COLOR_SURFACE_LOWEST       = "#ffffff"   # Cards, table rows — max contrast
COLOR_SURFACE_LOW          = "#f1f3ff"
COLOR_SURFACE              = "#e9edff"
COLOR_SURFACE_HIGH         = "#e1e8ff"
COLOR_SURFACE_HIGHEST      = "#d8e2ff"
COLOR_BACKGROUND           = "#faf9ff"   # Base canvas (soft grey feel)
COLOR_SURFACE_VARIANT      = "#d8e2ff"

# Content
COLOR_ON_SURFACE           = "#051a3e"   # Primary text
COLOR_ON_SURFACE_VARIANT   = "#434654"   # Secondary / muted text
COLOR_OUTLINE              = "#737685"   # Structural borders
COLOR_OUTLINE_VARIANT      = "#c3c6d6"   # Subtle borders (preferred for most dividers)

# Primary brand (Corporate Blue)
COLOR_PRIMARY              = "#003d9b"
COLOR_ON_PRIMARY           = "#ffffff"
COLOR_PRIMARY_CONTAINER    = "#0052cc"
COLOR_ON_PRIMARY_CONTAINER = "#c4d2ff"
COLOR_INVERSE_PRIMARY      = "#b2c5ff"
COLOR_SURFACE_TINT         = "#0c56d0"
COLOR_PRIMARY_FIXED        = "#dae2ff"
COLOR_PRIMARY_FIXED_DIM    = "#b2c5ff"

# Secondary
COLOR_SECONDARY            = "#535f73"
COLOR_ON_SECONDARY         = "#ffffff"
COLOR_SECONDARY_CONTAINER  = "#d4e0f8"
COLOR_ON_SECONDARY_CONTAINER = "#576377"

# Tertiary (warm accent)
COLOR_TERTIARY             = "#7b2600"
COLOR_ON_TERTIARY          = "#ffffff"
COLOR_TERTIARY_CONTAINER   = "#a33500"
COLOR_ON_TERTIARY_CONTAINER = "#ffc6b2"

# Error / Danger
COLOR_ERROR                = "#ba1a1a"
COLOR_ON_ERROR             = "#ffffff"
COLOR_ERROR_CONTAINER      = "#ffdad6"
COLOR_ON_ERROR_CONTAINER   = "#93000a"

# Inverse (used in dark surfaces / tooltips in light mode)
COLOR_INVERSE_SURFACE      = "#1d3054"
COLOR_INVERSE_ON_SURFACE   = "#edf0ff"
```

### Dark Mode Overrides
```python
# Dark mode — these replace the light tokens above when dark theme is active
DARK_BACKGROUND            = "#0B121F"   # Deep navy-black canvas
DARK_SURFACE               = "#161C27"   # Slightly lighter slate for cards
DARK_SURFACE_CONTAINER     = "#1E2738"
DARK_SURFACE_HIGH          = "#252D3D"   # Also used as dark border color
DARK_BORDER                = "#252D3D"
DARK_ON_SURFACE            = "#edf0ff"   # Light text
DARK_ON_SURFACE_VARIANT    = "#9ca3b8"   # Muted text in dark
DARK_OUTLINE               = "#3d4560"
DARK_PRIMARY               = "#b2c5ff"   # Accent in dark mode (inverse-primary)
DARK_PRIMARY_CONTAINER     = "#0040a2"

# Semantic colors (same in both modes — adjust opacity if needed)
COLOR_SUCCESS              = "#1e7d4a"   # Emerald green
COLOR_SUCCESS_BG           = "#d1fae5"
COLOR_WARNING              = "#b45309"   # Amber
COLOR_WARNING_BG           = "#fef3c7"
COLOR_DANGER               = "#ba1a1a"   # Ruby red
COLOR_DANGER_BG            = "#ffdad6"
```

---

## Typography Scale

```python
# Font families
FONT_HEADING  = "Hanken Grotesk"   # Headlines, titles, brand marks
FONT_BODY     = "Inter"            # Body text, labels, data, UI text

# Scale  (size_px, weight, line_height_px, letter_spacing_em)
TYPO_HEADLINE_LG = (24, 600, 32, -0.01)   # Page titles
TYPO_HEADLINE_MD = (20, 600, 28, -0.01)   # Section headers, card titles
TYPO_HEADLINE_SM = (16, 600, 24,  0.00)   # Sub-section headers, sidebar group labels
TYPO_BODY_LG     = (15, 400, 22,  0.00)   # Lead / emphasis body text
TYPO_BODY_MD     = (14, 400, 20,  0.00)   # Standard body / nav items
TYPO_BODY_SM     = (13, 400, 18,  0.00)   # Dense tables, secondary content
TYPO_LABEL_MD    = (12, 600, 16, +0.05)   # ALL-CAPS column headers, button text
TYPO_LABEL_SM    = (11, 500, 14,  0.00)   # Micro labels, status chips, timestamps
```

**Rules:**
- Use `FONT_HEADING` for `headline-*` sizes only
- Use `FONT_BODY` for everything else
- Data table cell content: `TYPO_BODY_SM`
- Data table column headers: `TYPO_LABEL_MD` (uppercase, tracked)
- Navigation items: `TYPO_BODY_MD` medium weight
- Button text: `TYPO_LABEL_MD`

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