"""Table page: title bar + auto-sized table + optional caption."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu
from ._titlebar import draw_title_bar


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a table slide with title and a styled table.

    渲染表格（table）版式：标题栏下居中放置格式化表格，自动按列数分配宽度，表头加粗底色，隔行变色；支持 header_rows、col_widths、align。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    headers = slide_dict["headers"]
    rows = slide_dict["rows"]
    caption = slide_dict.get("caption")

    body_l, body_t, body_w, body_h = grid.content_area()
    body_l = emu(body_l)
    body_t = emu(body_t) + emu(Inches(0.1))
    body_w = emu(body_w)
    body_h = emu(body_h) - emu(Inches(0.1))

    if caption:
        cap_h = emu(Inches(0.35))
        renderer.textbox(
            slide, body_l, body_t, body_w, cap_h,
            caption,
            font_size_pt=pt(tokens["font"]["size"]["caption"]),
            color=tokens["color"]["text_light"],
            align="left",
            anchor="top",
        )
        body_t += cap_h
        body_h -= cap_h

    # Reserve space for header + body rows with a comfortable row height.
    n_rows = len(rows) + 1
    # Auto row height: divide body_h evenly, minimum 0.32in
    min_row = emu(Inches(0.32))
    row_h = max(min_row, body_h // n_rows)
    table_h = row_h * n_rows
    # If table is shorter than available, top-align (no extra centering)

    renderer.table(
        slide,
        body_l,
        body_t,
        body_w,
        table_h,
        headers,
        rows,
    )
