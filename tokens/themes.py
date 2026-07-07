"""Theme presets: academic / business / teaching (built-in) plus extended
styles via ThemeFactory (tech / dark / minimal / nature / sunset / ocean /
forest / warm / premium / chinese_red / gradient / light / high_contrast).

Each theme is a partial dict that deep-merges over the base TOKENS.

@since v1.4.0 颜色 HEX 值统一从 ``skills.shared.themes`` 读取（ARCH-3），
             保证 PPT / DOCX / SVG 迷你图配色一致；字体 / 间距仍本地定义。
"""
from __future__ import annotations

from pptx.util import Pt
from pptx.dml.color import RGBColor


def _c(rgb_hex: str) -> RGBColor:
    """Parse '#RRGGBB' → RGBColor."""
    h = rgb_hex.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def _colors_from_shared(theme_name: str, chart_palette_override: list[str] | None = None) -> dict:
    """从 shared.themes 读取颜色 HEX 并转成 RGBColor dict（保持 v1.3 颜色值不变）。"""
    # 用 try/except 兼容独立运行场景（不依赖 skills 包路径）
    try:
        from ...shared.themes import get_palette, get_chart_palette
    except ImportError:  # pragma: no cover - 备用路径
        try:
            from skills.shared.themes import get_palette, get_chart_palette
        except ImportError:
            from shared.themes import get_palette, get_chart_palette

    pal = get_palette(theme_name)
    color_dict: dict = {}
    # 核心颜色键
    _key_map = [
        ("primary", "primary"),
        ("secondary", "secondary"),
        ("accent", "accent"),
        ("bg", "background"),       # shared 用 background，PPT 用 bg
        ("text", "text"),
        ("text_light", "muted"),    # PPT 的 text_light 即 muted
        ("text_on_primary", "text_on_primary"),
        ("title_bar_bg", "title_bar_bg"),
        ("table_header_bg", "table_header_bg"),
        ("table_header_text", "table_header_text"),
        ("table_alt_row", "table_alt_row"),
        ("section_bg", "section_bg"),
        ("section_num_color", "section_num_color"),
        ("kpi_bg", "kpi_bg"),
        ("quote_border", "quote_border"),
        ("divider", "divider"),
        ("chart_gridline", "chart_gridline"),
    ]
    for ppt_key, shared_key in _key_map:
        hex_val = pal.get(shared_key)
        if hex_val:
            color_dict[ppt_key] = _c(hex_val)

    # chart_palette
    chart_hex = chart_palette_override or get_chart_palette(theme_name)
    color_dict["chart_palette"] = [_c(hx) for hx in chart_hex]
    return color_dict


# -----------------------------------------------------------------
# Base themes (preserve v1.1 appearance exactly)
# -----------------------------------------------------------------
THEMES: dict[str, dict] = {
    # Academic – deep navy + accent red-brown (base defaults – empty override)
    "academic": {
        "color": _colors_from_shared("academic"),
    },

    # Business – slate blue-gray
    "business": {
        "color": _colors_from_shared("business"),
    },

    # Teaching – green primary, larger fonts
    "teaching": {
        "font": {
            "size": {
                "cover_title": Pt(44),
                "cover_subtitle": Pt(22),
                "cover_meta": Pt(16),
                "toc_num": Pt(30),
                "toc_text": Pt(20),
                "section_num": Pt(76),
                "section_title": Pt(36),
                "slide_title": Pt(32),
                "body_l1": Pt(20),
                "body_l2": Pt(17),
                "body_l3": Pt(15),
                "caption": Pt(12),
                "kpi_value": Pt(48),
                "kpi_label": Pt(16),
                "quote": Pt(26),
                "quote_attr": Pt(16),
                "table_header": Pt(15),
                "table_body": Pt(14),
                "page_num": Pt(11),
            },
        },
        "color": _colors_from_shared("teaching"),
    },
}


# -----------------------------------------------------------------
# Extended themes (v1.2+)
# -----------------------------------------------------------------
EXTENDED_THEMES: dict[str, dict] = {
    # Tech – deep navy + electric blue accents, dark-light contrast
    "tech": {
        "color": _colors_from_shared(
            "tech",
            chart_palette_override=["#00D4FF", "#64FFDA", "#7C3AED", "#F472B6"],
        ),
    },

    # Dark – near-black bg with light text, neon accents
    "dark": {
        "color": _colors_from_shared(
            "dark",
            chart_palette_override=["#60A5FA", "#F59E0B", "#34D399", "#F472B6"],
        ),
    },

    # Light / Minimal – pure white + black text, one accent only
    "light": {
        "color": _colors_from_shared(
            "minimal",
            chart_palette_override=["#2563EB", "#10B981", "#F59E0B", "#EF4444"],
        ),
    },
    "minimal": "light",  # alias

    # Nature – earthy green/brown
    "nature": {
        "color": _colors_from_shared(
            "nature",
            chart_palette_override=["#65A30D", "#A16207", "#CA8A04", "#16A34A"],
        ),
    },

    # Sunset – warm orange/pink/purple gradient vibe
    "sunset": {
        "color": _colors_from_shared(
            "sunset",
            chart_palette_override=["#EA580C", "#DB2777", "#F59E0B", "#A855F7"],
        ),
    },

    # Ocean – deep blue + teal
    "ocean": {
        "color": _colors_from_shared(
            "ocean",
            chart_palette_override=["#0284C7", "#06B6D4", "#0EA5E9", "#0891B2"],
        ),
    },

    # Forest – deep green + emerald accents
    "forest": {
        "color": _colors_from_shared(
            "forest",
            chart_palette_override=["#059669", "#D97706", "#10B981", "#84CC16"],
        ),
    },

    # Warm – cream bg + burgundy
    "warm": {
        "color": _colors_from_shared(
            "warm",
            chart_palette_override=["#B45309", "#D97706", "#DC2626", "#92400E"],
        ),
    },

    # Premium – gold/black luxury
    "premium": {
        "font": {
            "family": {
                "heading": "Georgia",
                "body": "Arial",
            },
        },
        "color": _colors_from_shared(
            "premium",
            chart_palette_override=["#D4AF37", "#1C1917", "#B08D57", "#78716C"],
        ),
    },

    # Chinese Red – red/gold traditional
    "chinese_red": {
        "color": _colors_from_shared(
            "chinese_red",
            chart_palette_override=["#C41E3A", "#D4AF37", "#8B0000", "#E67E22"],
        ),
    },

    # High-contrast – pure black / white for accessibility/print
    "high_contrast": {
        "color": _colors_from_shared(
            "high_contrast",
            chart_palette_override=["#000000", "#444444", "#888888", "#BBBBBB"],
        ),
    },
}

# Register extended themes
THEMES.update(EXTENDED_THEMES)


# -----------------------------------------------------------------
# ThemeFactory – resolve a theme from string, dict, or natural language hint
# -----------------------------------------------------------------

# Keyword → theme mapping for natural-language style queries
_STYLE_KEYWORDS: list[tuple[tuple[str, ...], str]] = [
    (("科技", "tech", "未来", "赛博", "cyber", "neon", "深色科技", "互联网", "ai", "AI"), "tech"),
    (("深色", "dark", "暗黑", "夜间", "dark mode", "黑色"), "dark"),
    (("极简", "minimal", "简约", "简洁", "白色", "light", "纯白", "苹果"), "minimal"),
    (("自然", "nature", "绿色", "环保", "生态", "植物"), "nature"),
    (("日落", "sunset", "橙粉", "暖色", "活力", "橘色"), "sunset"),
    (("海洋", "ocean", "蓝色", "天蓝", "海水", "海蓝"), "ocean"),
    (("森林", "forest", "深绿", "翠绿"), "forest"),
    (("温暖", "warm", "米色", "咖啡", "温馨", "棕黄"), "warm"),
    (("高端", "premium", "奢华", "黑金", "金色", "奢侈", "质感"), "premium"),
    (("中国红", "chinese_red", "红色", "国风", "传统", "喜庆", "春节"), "chinese_red"),
    (("学术", "答辩", "论文", "深蓝", "academic"), "academic"),
    (("商务", "企业", "汇报", "business", "路演"), "business"),
    (("教学", "课件", "教育", "绿色", "teaching"), "teaching"),
    (("高对比", "黑白", "print", "contrast", "打印"), "high_contrast"),
]


class ThemeFactory:
    """Resolve a theme specification into a theme override dict."""

    @staticmethod
    def from_style(query: str) -> dict:
        """Given a natural-language style hint, return a matching theme override.

        Falls back to ``business`` if no keyword matches.
        """
        q = (query or "").lower()
        for keywords, theme_name in _STYLE_KEYWORDS:
            for kw in keywords:
                if kw.lower() in q:
                    return ThemeFactory.get_theme(theme_name)
        # Default: business
        return ThemeFactory.get_theme("business")

    @staticmethod
    def get_theme(name: str) -> dict:
        """Return theme override dict by name (resolves aliases)."""
        if name not in THEMES:
            raise ValueError(
                f"Unknown theme: {name!r}. Available: {sorted(THEMES.keys())}"
            )
        t = THEMES[name]
        # Resolve string aliases (e.g. "minimal" → "light")
        if isinstance(t, str):
            return ThemeFactory.get_theme(t)
        # Return a deep copy so callers can mutate safely
        return _deep_copy_dict(t)

    @staticmethod
    def merge_overrides(base: dict, overrides: dict) -> dict:
        """Deep-merge a partial overrides dict onto base tokens."""
        out = _deep_copy_dict(base)
        for k, v in overrides.items():
            if k in out and isinstance(out[k], dict) and isinstance(v, dict):
                out[k] = ThemeFactory.merge_overrides(out[k], v)
            else:
                out[k] = v
        return out

    @staticmethod
    def resolve(theme_spec):
        """Resolve a theme specification (str | dict) to (overrides, resolved_name).

        Accepts:
        - A known theme name (str) e.g. ``"academic"``, ``"tech"``
        - A natural-language style hint (str) e.g. ``"科技感深色背景"``, ``"中国风红色"``
          – looked up via keyword matching; falls back to ``"business"``
        - A dict with partial overrides (merged onto business defaults)
        - A dict ``{"name": "tech", "overrides": {...}}`` for name + partial
        """
        if isinstance(theme_spec, str):
            if theme_spec in THEMES and not isinstance(THEMES.get(theme_spec), str):
                return ThemeFactory.get_theme(theme_spec), theme_spec
            # Try natural-language lookup
            matched = ThemeFactory.from_style(theme_spec)
            # from_style returns overrides dict; we need to know matched name
            name = "business"
            for keywords, theme_name in _STYLE_KEYWORDS:
                q = theme_spec.lower()
                for kw in keywords:
                    if kw.lower() in q:
                        name = theme_name
                        break
                if name != "business":
                    break
            return matched, name

        if isinstance(theme_spec, dict):
            if "name" in theme_spec:
                base = ThemeFactory.get_theme(theme_spec["name"])
                overrides = theme_spec.get("overrides", {})
                return ThemeFactory.merge_overrides(base, overrides), theme_spec["name"]
            return theme_spec, "custom"

        raise TypeError(f"theme must be str or dict, got {type(theme_spec)}")


def _deep_copy_dict(d: dict) -> dict:
    out = {}
    for k, v in d.items():
        if isinstance(v, dict):
            out[k] = _deep_copy_dict(v)
        elif isinstance(v, list):
            out[k] = list(v)
        else:
            out[k] = v
    return out


def list_themes() -> list[str]:
    """Return sorted list of available theme names (no aliases)."""
    return sorted(
        name for name, v in THEMES.items() if not isinstance(v, str)
    )
