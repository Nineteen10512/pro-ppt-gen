"""Two-column layout.

Each column (``left``/``right``) accepts either ``bullets`` (classic) or a
``chart`` semantic spec, optionally with a column ``title``.
"""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu
from ._titlebar import draw_title_bar
from ..engine.chart_renderer import render_chart


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a two-column layout comparing left/right content.

    渲染双栏（two_column）版式：标题栏下方等宽左右两栏，每栏可独立放 bullets/chart/image。支持 left_title/right_title 子标题。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    body_l, body_t, body_w, body_h = grid.content_area()
    body_t = emu(body_t) + emu(Inches(0.1))
    body_h = emu(body_h) - emu(Inches(0.1))
    body_l = emu(body_l)

    # Left: col 0..4 (5 cols), right: col 6..11 (6 cols) – gap at col 5
    lx, ly, lw, lh = grid.col_box(0, 5, body_t, body_h)
    rx, ry, rw, rh = grid.col_box(6, 6, body_t, body_h)

    left = slide_dict.get("left") or {}
    right = slide_dict.get("right") or {}

    label_size = pt(tokens["font"]["size"]["body_l1"])

    def _render_col(x, y, w, h, col_dict):
        cur_y = y
        col_title = col_dict.get("title")
        avail_h = h
        if col_title:
            th_ = emu(Inches(0.45))
            renderer.textbox(
                slide, x, cur_y, w, th_,
                col_title,
                font_size_pt=label_size,
                color=tokens["color"]["primary"],
                bold=True,
                align="left",
                anchor="middle",
                font_family=renderer.en_heading_font,
                cn_font_family=renderer.cn_heading_font,
            )
            renderer.rect(
                slide,
                x,
                cur_y + th_ - emu(Inches(0.04)),
                emu(grid.col_w(1)),
                emu(Inches(0.04)),
                fill_color=tokens["color"]["accent"],
            )
            cur_y += th_ + emu(Inches(0.1))
            avail_h -= (th_ + emu(Inches(0.1)))

        chart_spec = col_dict.get("chart")
        bullets = col_dict.get("bullets", [])
        if chart_spec:
            inset = emu(Inches(0.05))
            render_chart(
                slide, chart_spec,
                x, cur_y + inset, w, avail_h - inset,
                tokens, cn_font=renderer.cn_body_font,
            )
            # Optional small bullets under chart
            extra_bullets = col_dict.get("notes") or []
            if extra_bullets:
                note_h = emu(Inches(0.9))
                renderer.bullets(
                    slide, x, cur_y + avail_h - note_h, w, note_h,
                    extra_bullets, color=tokens["color"]["text"],
                )
        elif bullets:
            renderer.bullets(slide, x, cur_y, w, avail_h, bullets,
                             color=tokens["color"]["text"])

    _render_col(emu(lx), emu(ly), emu(lw), emu(lh), left)
    _render_col(emu(rx), emu(ry), emu(rw), emu(rh), right)
