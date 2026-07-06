"""Section divider page – full primary-color background with big number + title."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a section divider slide with large section number and title.

    渲染章节（section）分隔页：大号章节编号+章节名居中，使用主题 secondary 色做装饰，用于章节切换过渡。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["section_bg"])

    number = str(slide_dict["number"])
    title = slide_dict["title"]

    # Large watermark number (col 2-5, tall, light)
    num_size = pt(tokens["font"]["size"]["section_num"])
    nx, ny, nw, nh = grid.col_box(2, 4, emu(Inches(1.0)), emu(Inches(2.5)))
    renderer.textbox(
        slide, nx, ny, nw, nh,
        number,
        font_size_pt=num_size,
        color=tokens["color"]["section_num_color"],
        bold=True,
        align="left",
        anchor="top",
        font_family=renderer.en_heading_font,
        cn_font_family=renderer.cn_heading_font,
    )
    # Make number semi-transparent via XML (lighter feel)
    # (python-pptx has no direct alpha setter; leave as solid secondary color – acceptable.)

    # Title below number (col 2-11)
    title_size = pt(tokens["font"]["size"]["section_title"])
    tx, ty, tw, th = grid.col_box(2, 10, emu(Inches(3.6)), emu(Inches(1.2)))
    renderer.textbox(
        slide, tx, ty, tw, th,
        title,
        font_size_pt=title_size,
        color=tokens["color"]["text_on_primary"],
        bold=True,
        align="left",
        anchor="top",
        font_family=renderer.en_heading_font,
        cn_font_family=renderer.cn_heading_font,
    )

    # Accent bar left of title
    bar_h = Inches(0.06)
    renderer.rect(
        slide,
        grid.col_x(2),
        ty - emu(Inches(0.3)),
        emu(grid.col_w(2)),
        emu(bar_h),
        fill_color=tokens["color"]["accent"],
    )
