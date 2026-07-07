"""Taste and craft preflight checks for PRO-PPTX."""

from __future__ import annotations

import copy
from typing import Any

from shared.import_helper import import_shared
from shared.taste.adapters import ppt_units
from shared.taste.core import build_preflight, finalize_report, infer_design_read as _infer_design_read
from shared.taste.rules import check_copy_and_density, check_layout_rhythm, check_visual_intent, issue

from .tokens.design_tokens import get_tokens
from .tokens.themes import ThemeFactory
from .visual_redesign import visual_redesign_check


_HARD_FAIL_CODES = {"contrast_iron_law", "chart_color_variety", "placeholder_text", "word_doc_on_slide"}


def _resolved_theme_name(theme: Any) -> str:
    if isinstance(theme, str):
        return theme
    try:
        _overrides, name = ThemeFactory.resolve(theme)
        return str(name or "custom")
    except Exception:
        return "custom"


def _color_hex(value: Any) -> str | None:
    try:
        from pptx.dml.color import RGBColor as _RGBColor

        if isinstance(value, _RGBColor):
            return "#{:02X}{:02X}{:02X}".format(value[0], value[1], value[2])
    except Exception:
        pass
    if isinstance(value, str) and value.startswith("#") and len(value) == 7:
        return value.upper()
    if hasattr(value, "__iter__") and not isinstance(value, str):
        try:
            r, g, b = [int(v) for v in list(value)[:3]]
            return "#{:02X}{:02X}{:02X}".format(r, g, b)
        except Exception:
            return None
    return None


def _check_contrast_iron_law(theme: Any, issues: list[dict]) -> None:
    (hex_to_rgb, relative_luminance) = import_shared(
        "quality",
        attrs=["hex_to_rgb", "relative_luminance"],
    )
    tokens = get_tokens(theme)
    colors = tokens.get("color", {})

    pairs = [
        ("body text", colors.get("text"), colors.get("bg")),
        (
            "title-bar text",
            colors.get("text_on_primary") or colors.get("text"),
            colors.get("title_bar_bg") or colors.get("primary"),
        ),
    ]

    for label, fg, bg in pairs:
        fg_hex = _color_hex(fg)
        bg_hex = _color_hex(bg)
        if not fg_hex or not bg_hex:
            continue

        fg_lum = relative_luminance(*hex_to_rgb(fg_hex))
        bg_lum = relative_luminance(*hex_to_rgb(bg_hex))
        bg_is_dark = bg_lum <= 0.35
        bg_is_light = bg_lum >= 0.65
        fg_is_light = fg_lum >= 0.55
        fg_is_dark = fg_lum <= 0.45

        violated = (bg_is_dark and not fg_is_light) or (bg_is_light and not fg_is_dark)
        if violated:
            expected = "light text on dark background" if bg_is_dark else "dark text on light background"
            issue(
                issues,
                "error",
                "contrast_iron_law",
                f"{label} violates the contrast iron law: expected {expected}.",
                suggestion="Use light text on dark fills and dark text on light fills before rendering.",
            )


def _check_chart_color_variety(units: list[dict], theme: Any, issues: list[dict]) -> None:
    tokens = get_tokens(theme)
    palette = tokens.get("color", {}).get("chart_palette") or []
    unique_colors = {hx for hx in (_color_hex(color) for color in palette) if hx}

    for unit in units:
        if unit.get("kind") != "chart":
            continue
        series_count = int(unit.get("series_count") or 0)
        category_count = int(unit.get("category_count") or 0)
        if max(series_count, category_count) >= 2 and len(unique_colors) < 2:
            issue(
                issues,
                "error",
                "chart_color_variety",
                "Chart palette collapses to a single color, which is disallowed.",
                page=unit.get("index"),
                suggestion="Provide at least two distinct chart colors before generating the deck.",
            )


def infer_design_read(content: dict, theme: str = "academic", lang: str = "cn") -> dict:
    units = ppt_units(content)
    return _infer_design_read(
        units,
        theme=_resolved_theme_name(theme),
        meta=(content or {}).get("meta") or {},
        lang=lang,
        medium="presentation",
    )


def _enrich_chart_metadata(content: dict, units: list[dict]) -> None:
    slides = (content or {}).get("slides") or []
    for unit, slide in zip(units, (item for item in slides if isinstance(item, dict)), strict=False):
        if unit.get("kind") != "chart":
            continue
        chart = slide.get("chart") or {}
        unit["series_count"] = len(chart.get("series") or [])
        unit["category_count"] = len(chart.get("categories") or [])


def taste_check(
    content: dict,
    theme: str = "academic",
    lang: str = "cn",
    strict: bool = False,
) -> dict:
    from . import ppt_jsx

    working = copy.deepcopy(content)
    if isinstance(working, dict) and "slides" in working:
        try:
            working = ppt_jsx.auto_layout(working)
        except Exception:
            pass

    units = ppt_units(working)
    _enrich_chart_metadata(working, units)
    issues: list[dict] = []

    theme_name = _resolved_theme_name(theme)
    design_read = _infer_design_read(
        units,
        theme=theme_name,
        meta=(working or {}).get("meta") or {},
        lang=lang,
        medium="presentation",
    )
    _check_contrast_iron_law(theme, issues)
    check_copy_and_density(units, issues)
    check_visual_intent(units, issues)
    _check_chart_color_variety(units, theme, issues)
    check_layout_rhythm(units, issues)
    visual_report = visual_redesign_check(working, strict=strict)
    issues.extend(visual_report.get("issues") or [])

    base_quality = None
    try:
        base_quality = ppt_jsx.quality_check(working, theme=theme, lang=lang)
    except Exception as exc:
        issue(
            issues,
            "warning",
            "base_quality_unavailable",
            f"Base quality_check could not run: {exc}",
            suggestion="Fix schema or theme inputs, then rerun taste_check.",
        )

    story = None
    try:
        story = ppt_jsx.story_check(working, method="pyramid")
    except Exception:
        story = None

    preflight = build_preflight(
        issues,
        base_quality=base_quality,
        story=story,
        include_visual_intent=True,
    )
    return finalize_report(
        version="1.5",
        design_read=design_read,
        issues=issues,
        preflight=preflight,
        strict=strict,
        hard_fail_codes=_HARD_FAIL_CODES,
        base_quality=base_quality,
        story=story,
        visual_redesign=visual_report,
    )


__all__ = ["taste_check", "infer_design_read"]
