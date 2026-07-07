from __future__ import annotations

from typing import Any


def _text_items(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, dict):
        out = []
        for key in ("title", "text", "label", "value", "subtext", "desc", "caption"):
            if value.get(key):
                out.append(str(value[key]))
        for key in ("bullets", "items"):
            out.extend(_text_items(value.get(key)))
        return out
    if isinstance(value, list):
        out = []
        for item in value:
            out.extend(_text_items(item))
        return out
    return []


def ppt_units(content: dict) -> list[dict]:
    """Convert PRO-PPTX semantic slides to generic taste-check units."""
    slides = (content or {}).get("slides") or []
    units: list[dict] = []
    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        layout = slide.get("layout") or "content"
        texts = []
        for key in ("title", "subtitle", "text", "caption", "overlay", "notes"):
            texts.extend(_text_items(slide.get(key)))
        for key in ("bullets", "items", "events", "headers", "rows", "left", "right"):
            texts.extend(_text_items(slide.get(key)))
        chart = slide.get("chart") or {}
        kind = "chart" if chart else layout
        units.append(
            {
                "index": index,
                "kind": kind,
                "layout": layout,
                "title": slide.get("title") or slide.get("text") or "",
                "texts": [text for text in texts if text],
                "has_visual": bool(slide.get("image_path") or slide.get("image_prompt") or chart or slide.get("free_shape") or slide.get("background_shape")),
                "has_chart": bool(chart),
                "bullet_count": len(slide.get("bullets") or []),
                "raw": slide,
            }
        )
    return units
