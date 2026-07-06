"""Semantic JSON → slide layout compiler.

``compile_slides`` validates input, resolves a layout handler per slide,
and invokes each handler with a (renderer, grid, tokens, slide_dict) tuple.
Handlers are discovered automatically from ``skills.pro_ppt_gen.layouts``.

@since v1.3: renders background_shape before layout body, decorations after.
"""
from __future__ import annotations

import importlib
from typing import Callable

from .validators import (
    validate_slides,
    validate_meta,
    META_DEFAULTS,
    ValidationError,
    ValidationWarning,
)


# Layout-name → module path within skills.pro_ppt_gen.layouts
_LAYOUT_MODULES = {
    "cover": "title",
    "toc": "toc",
    "section": "section",
    "content": "content",
    "two_column": "two_column",
    "image_text": "image_text",
    "full_image": "full_image",
    "table": "table",
    "chart": "chart",
    "kpi": "kpi",
    "quote": "quote",
    "summary": "summary",
    "thanks": "thanks",
    "timeline": "timeline",
    "references": "references",
}

_HANDLER_CACHE: dict[str, Callable] = {}


def _get_handler(layout: str) -> Callable:
    if layout in _HANDLER_CACHE:
        return _HANDLER_CACHE[layout]
    mod_name = _LAYOUT_MODULES.get(layout)
    if mod_name is None:
        # English: f"No handler for layout: {layout!r}"
        raise ValidationError(
            f"没有找到版式 {layout!r} 的处理器。修复建议：请从支持的版式中选择：{sorted(_LAYOUT_MODULES.keys())}"
        )
    from .. import layouts as _layouts_pkg
    mod = importlib.import_module(f".{mod_name}", _layouts_pkg.__name__)
    if not hasattr(mod, "render"):
        # English: f"Layout module {mod_name} missing render() function"
        raise ValidationError(
            f"版式模块 {mod_name} 缺少 render() 函数。修复建议：请检查该版式文件是否被意外修改"
        )
    handler = mod.render
    _HANDLER_CACHE[layout] = handler
    return handler


def _render_background_svg(slide_dict, renderer, slide, grid, tokens):
    """v1.3: Inject background_shape SVG as bottom layer."""
    svg = slide_dict.get("background_shape")
    if not svg:
        return
    try:
        renderer.free_shape(
            slide, svg,
            0, 0, int(grid.slide_w), int(grid.slide_h),
            shape_id=1, name="BgShape",
        )
    except Exception as e:
        print(f"[background_shape] SVG render failed: {e}")


def _render_decorations(slide_dict, renderer, slide, grid, tokens):
    """v1.3: Inject page-level decoration SVGs."""
    from ..layouts._decorations import render_decorations
    try:
        render_decorations(renderer, slide, slide_dict, grid, tokens)
    except Exception as e:
        print(f"[decorations] render failed: {e}")


def normalize_meta(content: dict):
    """Normalize meta fields (logo, position, page_number_style, transition).

    规范化 slide 元信息：裁剪 logo 路径空白、logo_position 枚举大小写归一化（tl/tr/bl/br/center）、
    page_number_style 枚举校验（plain/slash/chinese/roman）、transition 字段透传。缺省字段不填默认值
    （以保持 v1.3 行为）。"""
    meta = content.get("meta")
    if not isinstance(meta, dict):
        meta = {}
        content["meta"] = meta
    for k, v in META_DEFAULTS.items():
        meta.setdefault(k, v)
    # Normalize case for enum fields
    if isinstance(meta.get("logo_position"), str):
        meta["logo_position"] = meta["logo_position"].lower()
    if isinstance(meta.get("page_number_style"), str):
        meta["page_number_style"] = meta["page_number_style"].lower()
    # Trim whitespace on logo path
    if isinstance(meta.get("logo"), str):
        meta["logo"] = meta["logo"].strip() or None


def compile_slides(content: dict, renderer, grid, tokens) -> list[ValidationWarning]:
    """Compile semantic slides dict onto the renderer via layout handlers.

    将语义 slides 列表编译到 renderer 上：按 slide.layout 分派到 layouts/<layout>.py 的 render 函数，
    每个 layout 通过 GridLayout 计算 EMU 坐标、调用 Renderer 原语构造形状。v1.4 起在编译前先调用
    normalize_meta 规范化 meta.logo/logo_position/page_number_style/transition 并 validate_meta 校验。"""
    normalize_meta(content)
    warnings = validate_meta(content)
    warnings.extend(validate_slides(content))
    slides = content["slides"]
    for slide_dict in slides:
        layout = slide_dict["layout"]
        handler = _get_handler(layout)
        # Track metadata (e.g., for page-number skipping, notes)
        if hasattr(renderer, "begin_slide"):
            renderer.begin_slide(layout)

        # Handler creates new_slide (we hook into begin_slide/new_slide? no —
        # handlers call renderer.new_slide() internally. To support bg shapes
        # we patch: instead we use a wrapper new_slide that records the slide
        # and draws the bg immediately after creation. We provide a slide
        # post-creation hook by monkey-patching begin_slide to set a callback.
        # Simpler: hook into renderer.new_slide via wrapping.
        _orig_new_slide = renderer.new_slide
        bg_drawn = {"v": False}
        dec_drawn = {"v": False}

        def _new_slide_wrap(bg_color=None):
            sl = _orig_new_slide(bg_color=bg_color)
            if not bg_drawn["v"]:
                _render_background_svg(slide_dict, renderer, sl, grid, tokens)
                bg_drawn["v"] = True
            return sl

        renderer.new_slide = _new_slide_wrap
        try:
            handler(slide_dict, renderer, grid, tokens)
        finally:
            renderer.new_slide = _orig_new_slide

        # Render decorations after main content (on the most recent slide)
        if renderer.slides:
            _render_decorations(slide_dict, renderer, renderer.slides[-1], grid, tokens)

        # Attach speaker notes if provided (centralized so layouts don't need to)
        # v1.4 P1-1: prefer new `speaker_notes` field; fall back to legacy `notes`
        notes = slide_dict.get("speaker_notes")
        if not notes:
            notes = slide_dict.get("notes")
        if notes and hasattr(renderer, "add_speaker_notes"):
            try:
                renderer.add_speaker_notes(renderer.slides[-1], notes)
            except Exception:
                pass

        # v1.4 P2-2: slide transition
        transition = slide_dict.get("transition")
        if transition is not None and hasattr(renderer, "apply_transition"):
            try:
                renderer.apply_transition(renderer.slides[-1], transition)
            except Exception:
                pass

        # v1.5.3 B-5: element animation
        animation = slide_dict.get("animation")
        if animation is not None and hasattr(renderer, "set_slide_animation"):
            try:
                renderer.set_slide_animation(animation)
            except Exception:
                pass

    # v1.5.3: build <p:timing> for all animated slides after all shapes placed
    if hasattr(renderer, "finalize_slide_animations"):
        try:
            renderer.finalize_slide_animations()
        except Exception as e:
            warnings.append(f"animation finalize failed: {e}")

    return warnings
