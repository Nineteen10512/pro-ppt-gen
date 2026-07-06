"""Standard content page: title bar + bullet list body (and optional chart).

Supports three modes based on provided fields:

    1. bullets only (classic)   – bullets fill the body.
    2. chart only               – chart fills the body.
    3. bullets + chart          – auto split: left bullets, right chart.
        * 1–3 bullets → left 4 cols / right 8 cols
        * 4+ bullets  → left 6 cols / right 6 cols
"""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import emu
from ._titlebar import draw_title_bar
from ._decorations import render_decorations
from ..engine.chart_renderer import render_chart


def _render_chart_in_box(slide, renderer, grid, tokens, chart_spec, x, y, w, h):
    """Helper: render a chart with a small top inset so it aligns visually."""
    inset = emu(Inches(0.05))
    render_chart(
        slide, chart_spec,
        x, y + inset, w, h - inset,
        tokens, cn_font=renderer.cn_body_font,
    )


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a standard content slide with title + bullets/content blocks.

    渲染正文（content）版式：顶部标题栏，下方 12 列网格内容区。支持 bullets 列表、chart 内嵌图表、多段 text 文本块自动分配高度。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    bullets = slide_dict.get("bullets") or []
    chart_spec = slide_dict.get("chart")

    body_l, body_t, body_w, body_h = grid.content_area()
    pad = emu(Inches(0.1))
    body_l = emu(body_l)
    body_t = emu(body_t) + pad
    body_w = emu(body_w)
    body_h = emu(body_h) - pad

    if bullets and chart_spec:
        # Split: bullets left, chart right
        if len(bullets) <= 3:
            bullets_span, chart_span = 4, 8
        else:
            bullets_span, chart_span = 6, 6
        # bullets occupy cols [0 .. bullets_span-1]
        bx, by, bw, bh = grid.col_box(0, bullets_span, body_t, body_h)
        # chart starts at col (bullets_span+1) to leave a gutter
        chart_col_start = bullets_span + 1
        chart_span_eff = 12 - chart_col_start
        # If we can't fit (span too small), fall back to 6/6 split
        if chart_span_eff < 5:
            chart_col_start = 6
            chart_span_eff = 6
        cx_, cy_, cw_, ch_ = grid.col_box(chart_col_start, chart_span_eff, body_t, body_h)
        renderer.bullets(
            slide, emu(bx), emu(by), emu(bw), emu(bh),
            bullets, color=tokens["color"]["text"],
        )
        _render_chart_in_box(slide, renderer, grid, tokens, chart_spec,
                             emu(cx_), emu(cy_), emu(cw_), emu(ch_))
    elif chart_spec:
        _render_chart_in_box(slide, renderer, grid, tokens, chart_spec,
                             body_l, body_t, body_w, body_h)
    else:
        renderer.bullets(
            slide,
            body_l, body_t, body_w, body_h,
            bullets,
            color=tokens["color"]["text"],
        )

    render_decorations(renderer, slide, slide_dict, grid, tokens)
