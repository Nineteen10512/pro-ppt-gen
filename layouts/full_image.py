"""Full-image layout – image fills the content area; optional title with overlay.

@since v1.3.0: supports image_prompt placeholder, background_shape SVG.
"""
from __future__ import annotations

from pptx.util import Inches
from pptx.dml.color import RGBColor

from ._helpers import pt, emu
from ._decorations import (
    render_image_or_placeholder, render_decorations, render_background_shape,
)


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a full-bleed image slide with optional overlay title.

    渲染全图（full_image）版式：图片铺满幻灯片（12 列 × 全高），可选 overlay 在底部半透明条上叠加标题/说明。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])

    # Background decorative SVG (v1.3) — drawn before image so image can cover
    render_background_shape(renderer, slide, slide_dict, grid, tokens)

    title = slide_dict.get("title")
    overlay = slide_dict.get("overlay", False) or bool(title)

    il = emu(grid.margin_left)
    it_ = emu(grid.margin_top)
    iw = emu(grid.content_w)
    ih = emu(grid.content_h)

    render_image_or_placeholder(renderer, slide, slide_dict, tokens,
                                il, it_, iw, ih, fit=True)

    if title and overlay:
        overlay_h = emu(Inches(1.4))
        overlay_top = emu(grid.slide_h - grid.margin_bottom) - overlay_h
        renderer.rect(
            slide, il, overlay_top, iw, overlay_h,
            fill_color=RGBColor(0x00, 0x00, 0x00),
        )
        renderer.textbox(
            slide,
            il + emu(Inches(0.4)),
            overlay_top + emu(Inches(0.3)),
            iw - emu(Inches(0.8)),
            overlay_h - emu(Inches(0.4)),
            title,
            font_size_pt=pt(tokens["font"]["size"]["slide_title"]),
            color=tokens["color"]["text_on_primary"],
            bold=True,
            align="left",
            anchor="middle",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )
    elif title:
        renderer.textbox(
            slide,
            il + emu(Inches(0.2)),
            it_ + emu(Inches(0.2)),
            iw - emu(Inches(0.4)),
            emu(Inches(0.6)),
            title,
            font_size_pt=pt(tokens["font"]["size"]["slide_title"]),
            color=tokens["color"]["primary"],
            bold=True,
            align="left",
            anchor="top",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )
    render_decorations(renderer, slide, slide_dict, grid, tokens)
