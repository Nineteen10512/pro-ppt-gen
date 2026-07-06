"""Cover / title page layout."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu


def render(slide_dict: dict, renderer, grid, tokens):
    """Render the cover (title) slide with title/subtitle/author/date.

    渲染封面（cover）版式：主标题居中大字，副标题在主标题下方，作者/机构/日期位于页脚区域。背景使用主题 primary 色矩形装饰条；支持 background_image 全屏底图。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])

    # Decorative left accent bar
    bar_left = emu(grid.margin_left)
    bar_top = emu(Inches(1.8))
    bar_h = emu(Inches(2.0))
    renderer.vertical_line(slide, bar_left, bar_top, bar_h, tokens["color"]["accent"], width_pt=6)

    # Title
    title_size = pt(tokens["font"]["size"]["cover_title"])
    title_left = bar_left + emu(Inches(0.3))
    title_top = emu(Inches(2.0))
    title_w = emu(grid.content_w - Inches(0.3))
    title_h = emu(Inches(1.2))
    renderer.textbox(
        slide, title_left, title_top, title_w, title_h,
        slide_dict["title"],
        font_size_pt=title_size,
        color=tokens["color"]["primary"],
        bold=True,
        align="left",
        anchor="middle",
        font_family=renderer.en_heading_font,
        cn_font_family=renderer.cn_heading_font,
    )

    # Subtitle
    subtitle = slide_dict.get("subtitle")
    if subtitle:
        sub_size = pt(tokens["font"]["size"]["cover_subtitle"])
        renderer.textbox(
            slide,
            title_left,
            title_top + title_h + emu(Inches(0.1)),
            title_w,
            emu(Inches(0.8)),
            subtitle,
            font_size_pt=sub_size,
            color=tokens["color"]["text_light"],
            align="left",
            anchor="top",
            font_family=renderer.en_body_font,
            cn_font_family=renderer.cn_body_font,
        )

    # Meta line (author / institution / date) – bottom area
    meta_parts = [p for p in (slide_dict.get("author"), slide_dict.get("institution"), slide_dict.get("date")) if p]
    if meta_parts:
        meta_text = "    ".join(meta_parts)
        meta_size = pt(tokens["font"]["size"]["cover_meta"])
        meta_top = emu(grid.slide_h - grid.margin_bottom - Inches(0.8))
        renderer.textbox(
            slide,
            title_left,
            meta_top,
            title_w,
            emu(Inches(0.4)),
            meta_text,
            font_size_pt=meta_size,
            color=tokens["color"]["text_light"],
            align="left",
            anchor="top",
        )

    # Accent bottom stripe
    stripe_h = Inches(0.08)
    renderer.rect(
        slide,
        emu(grid.margin_left),
        emu(grid.slide_h - grid.margin_bottom - stripe_h),
        emu(grid.content_w),
        emu(stripe_h),
        fill_color=tokens["color"]["secondary"],
    )
