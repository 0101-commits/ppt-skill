# -*- coding: utf-8 -*-
"""Pure design tokens for HCG-Slide-Design-System v1.0 (the absolute standard).

No third-party imports — consumed by BOTH core/designer.py (PPTX) and
core/html_renderer.py (HTML). Mirrors skill_ppt_design.json.

Canvas: 16:9, 1280x720 px (top-left origin). px -> in = px/96; pt = px*0.75.
"""

PALETTE = {
    "primary": "#356CB5", "primary-deep": "#1F4E8C", "primary-steel": "#336699",
    "primary-navy": "#0D1467", "primary-soft": "#A1D1F1",
    "tint-blue-1": "#D4E1F2", "tint-blue-2": "#B0CEF2", "tint-blue-3": "#A5C4E1",
    "tint-blue-4": "#EAF2FB",
    "accent-coral": "#F16249", "accent-gold": "#FFCC66", "accent-teal": "#50B8B6",
    "emphasis-red": "#C00000", "positive-green": "#00B050",
    "ink": "#000000", "ink-slate": "#263238", "text-muted": "#919191",
    "text-faint": "#B7B7B7", "divider": "#DDDDDD", "divider-soft": "#EAEAEA",
    "surface": "#FFFFFF", "surface-gray": "#F5F6F8", "hcg-maroon": "#921F0B",
}

# role -> (px, weight, lineHeight)
TYPO = {
    "cover-title": (38, 800, 1.22), "cover-subtitle": (20, 500, 1.40),
    "cover-date": (15, 400, 1.40), "slide-title": (22, 700, 1.28),
    "kicker": (18, 700, 1.20), "chapter-indicator": (13, 600, 1.10),
    "section-header": (18, 700, 1.25), "block-header": (16, 700, 1.25),
    "body": (15, 400, 1.45), "body-strong": (15, 700, 1.45),
    "body-sm": (13, 400, 1.40), "body-sm-strong": (13, 700, 1.40),
    "caption": (12, 400, 1.35), "source-note": (11, 400, 1.30),
    "page-number": (13, 600, 1.10), "chart-axis": (12, 500, 1.20),
    "chart-value": (13, 700, 1.10), "kpi-number": (44, 800, 1.05),
    "kpi-unit": (16, 600, 1.10),
}

SAFE = {"left": 56, "right": 56, "top": 28, "bottom": 28}
CONTENT = {"x": 56, "y": 152, "w": 1168, "h": 512}
GUTTER, BLOCK_GAP = 24, 12
ROUNDED = {"sm": 3, "md": 6, "lg": 10, "card": 8, "chip": 13, "circle": 9999}
FONT = "Pretendard"
SERIES_PALETTE = ["primary", "accent-coral", "accent-teal",
                  "accent-gold", "primary-soft", "primary-steel"]
ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ"]


def apply_design_tokens(tokens, overrides=None):
    """Inject palette from parsed skill_ppt_design.json (mutates PALETTE in place
    so every importer sees the update). Reads tokens["colors"] (v2.0) or
    tokens["theme_colors"] (legacy: hex or {"rgb": hex}); `overrides` (client
    identity.colors, {name: "#hex"}) applied last."""
    src = tokens.get("colors") or tokens.get("theme_colors") or {}
    for name, val in src.items():
        hexv = val.get("rgb") if isinstance(val, dict) else val
        if hexv:
            PALETTE[name] = "#" + str(hexv).lstrip("#")
    for name, hexv in (overrides or {}).items():
        if hexv:
            PALETTE[name] = "#" + str(hexv).lstrip("#")
    return PALETTE


def hex6(name_or_hex):
    """Token name or hex -> '#RRGGBB'."""
    return "#" + str(PALETTE.get(name_or_hex, name_or_hex)).lstrip("#").upper()


def segments(text):
    """Parse '**emphasis**' markup -> [(chunk, is_emph), ...]."""
    out = []
    for i, part in enumerate(str(text).split("**")):
        if part:
            out.append((part, i % 2 == 1))
    return out or [("", False)]
