# PyQt6 UI Designer Skill

A Claude skill for designing and refining **modern PyQt6 desktop interfaces** with a consistent visual system, light/dark themes, and production-ready QSS.

This repository is built around a **`SKILL.md`** file and supporting design references:

- `references/design_tokens.md` — canonical colors, typography, spacing, radius, and elevation
- `references/qss_patterns.md` — reusable QSS patterns for PyQt6 widgets and theme states
- `references/component_library.md` — ready-made component patterns for common UI blocks

## What this skill is for

Use this skill whenever you need to:

- build a new PyQt6 UI
- improve an existing PyQt6 interface
- modernize a window, sidebar, table, form, card, modal, or settings panel
- add light/dark theme support
- apply a clean enterprise look with rounded shapes, subtle depth, and modern icons
- write or refine QSS stylesheets for PyQt6 widgets

## What it does

The skill acts as a **PyQt6 UI design assistant**. It does not just generate code; it guides the design process so interfaces stay consistent and professional.

It focuses on:

- token-based design decisions
- 4px spacing discipline
- light and dark theme parity
- consistent component styling
- modern icon usage
- clean layout hierarchy
- reusable PyQt6 patterns
- Context7 MCP lookups for PyQt6 API accuracy when needed

## Included structure

```text
SKILL.md
references/
├── component_library.md
├── design_tokens.md
└── qss_patterns.md
```

## Core design rules

- Use the token values from `design_tokens.md`
- Keep spacing on a 4px grid
- Provide both light and dark variants
- Prefer subtle borders and soft depth over heavy shadows
- Keep typography consistent
- Style all interactive states: default, hover, pressed, focus, disabled
- Use modern icons where appropriate
- Keep layouts clean, readable, and responsive

## How the skill should be used

When Claude receives a PyQt6 UI request, the skill should:

1. Read the design tokens first
2. Apply the QSS patterns
3. Reuse the component library when relevant
4. Use Context7 MCP for PyQt6 details if needed
5. Return clean, modular, ready-to-use Python code

## Good fits

- dashboard UIs
- admin panels
- settings screens
- forms and data entry views
- data tables and filters
- sidebar-based app shells
- theme refreshes and UI polish

## Example use cases

- “Make this PyQt6 window look modern.”
- “Add a dark mode sidebar.”
- “Refine this table and toolbar.”
- “Style my buttons and inputs consistently.”
- “Generate a professional app shell with cards and navigation.”

## Notes

This skill is intended to improve the look and usability of PyQt6 applications without changing their core purpose. It helps keep design choices consistent across the whole app.

---

## License

Add your preferred license here.
