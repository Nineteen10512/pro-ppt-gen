"""Design tokens: centralized typography, spacing, color, and shape constants.

All layout modules MUST reference tokens instead of hard-coding numeric values.
"""
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ---------------------------------------------------------------------------
# Base tokens (academic theme defaults)
# ---------------------------------------------------------------------------
TOKENS = {
    "slide": {"w": SLIDE_W, "h": SLIDE_H, "aspect": "16:9"},
    "margin": {
        "top": Inches(0.4),
        "bottom": Inches(0.4),
        "left": Inches(0.5),
        "right": Inches(0.5),
    },
    "grid": {"cols": 12, "gutter": Inches(0.15)},
    "title_bar": {"height": Inches(0.8), "gap_after": Inches(0.2)},
    "footer": {"height": Inches(0.3)},
    "chart": {
        "margin": Inches(0.1),
        "title_font_size": Pt(14),
        "label_font_size": Pt(9),
        "legend_font_size": Pt(10),
        "data_label_font_size": Pt(9),
        "axis_title_font_size": Pt(10),
        "gridline_width_pt": 0.5,
        "marker_size": 5,
    },
    "font": {
        "family": {
            "heading": "Arial",
            "body": "Arial",
            "cn_heading": "微软雅黑",
            "cn_body": "微软雅黑",
            "accent": "Georgia",
        },
        "size": {
            "cover_title": Pt(40),
            "cover_subtitle": Pt(20),
            "cover_meta": Pt(14),
            "toc_num": Pt(28),
            "toc_text": Pt(18),
            "section_num": Pt(72),
            "section_title": Pt(32),
            "slide_title": Pt(28),
            "body_l1": Pt(18),
            "body_l2": Pt(15),
            "body_l3": Pt(13),
            "caption": Pt(11),
            "kpi_value": Pt(44),
            "kpi_label": Pt(14),
            "quote": Pt(24),
            "quote_attr": Pt(14),
            "table_header": Pt(13),
            "table_body": Pt(12),
            "page_num": Pt(10),
        },
    },
    "color": {
        "primary": RGBColor(0x1F, 0x38, 0x64),
        "secondary": RGBColor(0x2E, 0x75, 0xB6),
        "accent": RGBColor(0xC0, 0x50, 0x4D),
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
        "text": RGBColor(0x33, 0x33, 0x33),
        "text_light": RGBColor(0x66, 0x66, 0x66),
        "text_on_primary": RGBColor(0xFF, 0xFF, 0xFF),
        "title_bar_bg": RGBColor(0x1F, 0x38, 0x64),
        "table_header_bg": RGBColor(0x1F, 0x38, 0x64),
        "table_header_text": RGBColor(0xFF, 0xFF, 0xFF),
        "table_alt_row": RGBColor(0xF2, 0xF6, 0xFC),
        "section_bg": RGBColor(0x1F, 0x38, 0x64),
        "section_num_color": RGBColor(0x2E, 0x75, 0xB6),
        "kpi_bg": RGBColor(0xF2, 0xF6, 0xFC),
        "quote_border": RGBColor(0x2E, 0x75, 0xB6),
        "divider": RGBColor(0x2E, 0x75, 0xB6),
        "chart_gridline": RGBColor(0xDD, 0xDD, 0xDD),
        "chart_palette": [
            RGBColor(0x1F, 0x38, 0x64),  # deep navy
            RGBColor(0xC0, 0x50, 0x4D),  # red-brown accent
            RGBColor(0x2E, 0x75, 0xB6),  # steel blue
            RGBColor(0x7F, 0x60, 0x4F),  # taupe/gray-brown
        ],
    },
    "bullet": {
        "char": {"l1": "●", "l2": "○", "l3": "▪"},
        "indent_l1": Inches(0.25),
        "indent_l2": Inches(0.5),
        "indent_l3": Inches(0.75),
        "space_after": Pt(6),
    },
    "shape": {"corner_radius": Pt(4), "shadow": False},
}


def deep_merge(base: dict, override: dict) -> dict:
    """Deep-merge override into base (returns new dict; base untouched)."""
    out = {k: (dict(v) if isinstance(v, dict) else v) for k, v in base.items()}
    for k, v in override.items():
        if k in out and isinstance(out[k], dict) and isinstance(v, dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def get_tokens(theme="academic") -> dict:
    """Return a fully-resolved token dict for the given theme.

    ``theme`` may be:
    - a string (built-in or extended theme name, e.g. ``"tech"``)
    - a dict (partial overrides merged onto base tokens)
    - ``{"name": "<theme_name>", "overrides": {...}}`` for named + overrides
    """
    from .themes import ThemeFactory
    overrides, _name = ThemeFactory.resolve(theme)
    return deep_merge(TOKENS, overrides)
