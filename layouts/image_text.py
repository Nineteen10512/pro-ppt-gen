"""Image + text layout (image left or right, bullets on the other side).

@since v1.3.0: supports image_prompt placeholder and decorations.
"""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import emu
from ._titlebar import draw_title_bar
from ._decorations import (
    render_image_or_placeholder, render_decorations,
)


def render(slide_dict: dict, renderer, grid, tokens):
    """Render an image + text side-by-side layout.

    渲染图文混排（image_text）版式：标题栏下左图右文或左文右图（image_side=left/right）。图片按比例 fit 进 6 列网格，文本区支持 bullets 与多段。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    side = slide_dict.get("side", "left")
    bullets = slide_dict.get("bullets", [])

    body_l, body_t, body_w, body_h = grid.content_area()
    body_t = emu(body_t) + emu(Inches(0.1))
    body_h = emu(body_h) - emu(Inches(0.1))

    if side == "left":
        img_box = grid.col_box(0, 6, body_t, body_h)
        txt_box = grid.col_box(6, 6, body_t, body_h)
    else:
        img_box = grid.col_box(6, 6, body_t, body_h)
        txt_box = grid.col_box(0, 6, body_t, body_h)

    il, it_, iw, ih = [emu(v) for v in img_box]
    tl, tt, tw, th = [emu(v) for v in txt_box]

    render_image_or_placeholder(renderer, slide, slide_dict, tokens,
                                il, it_, iw, ih, fit=True)

    pad = emu(grid.gutter)
    if side == "left":
        tl += pad
        tw -= pad
    else:
        tw -= pad

    renderer.bullets(slide, tl, tt, tw, th, bullets, color=tokens["color"]["text"])
    render_decorations(renderer, slide, slide_dict, grid, tokens)
