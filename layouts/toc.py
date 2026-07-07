"""Table of contents layout."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu
from ._titlebar import draw_title_bar


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a table-of-contents slide listing all sections.

    渲染目录（toc）版式：自动从 content.slides 中收集 layout=section 的条目，按编号+标题分行列出。支持 numbered(bool) 控制是否显示序号，columns 控制分栏。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    title = slide_dict.get("title", "目录")
    draw_title_bar(renderer, slide, grid, tokens, title)

    items = slide_dict["items"]
    n = len(items)

    # Start below title bar
    top0 = grid.body_top() + emu(Inches(0.2))
    row_h = emu(Inches(0.6))
    gap = emu(Inches(0.1))
    total_h = n * row_h + (n - 1) * gap

    # Center vertically if short
    avail_h = grid.body_height() - emu(Inches(0.4))
    if total_h < avail_h:
        top0 = grid.body_top() + (avail_h - total_h) // 2

    num_col_start = 1
    num_col_span = 2
    txt_col_start = 3
    txt_col_span = 8

    num_size = pt(tokens["font"]["size"]["toc_num"])
    txt_size = pt(tokens["font"]["size"]["toc_text"])

    for i, it in enumerate(items):
        if isinstance(it, dict):
            num = str(it.get("num", f"{i + 1:02d}"))
            text = it.get("text", "")
        else:
            num = f"{i + 1:02d}"
            text = str(it)

        y = top0 + i * (row_h + gap)

        # Number
        nx, ny, nw, nh = grid.col_box(num_col_start, num_col_span, y, row_h)
        renderer.textbox(
            slide, nx, ny, nw, nh,
            num,
            font_size_pt=num_size,
            color=tokens["color"]["accent"],
            bold=True,
            align="right",
            anchor="middle",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )
        # Divider dot (use a small square bar)
        renderer.rect(
            slide,
            grid.col_x(3) - emu(grid.gutter) // 2 - emu(Inches(0.04)),
            y + row_h // 2 - emu(Inches(0.02)),
            emu(Inches(0.04)),
            emu(Inches(0.04)),
            fill_color=tokens["color"]["divider"],
        )
        # Text
        tx, ty, tw, th = grid.col_box(txt_col_start, txt_col_span, y, row_h)
        renderer.textbox(
            slide, tx, ty, tw, th,
            text,
            font_size_pt=txt_size,
            color=tokens["color"]["text"],
            bold=False,
            align="left",
            anchor="middle",
            font_family=renderer.en_body_font,
            cn_font_family=renderer.cn_body_font,
        )
