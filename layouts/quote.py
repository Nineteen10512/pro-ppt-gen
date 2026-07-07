"""Quote / pull-out page: big decorative text with optional attribution."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a quote slide with large centered quotation and attribution.

    渲染金句（quote）版式：大引号装饰 + 居中引文 + 右下角 attribution（作者/出处），使用主题 accent 色。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])

    text = slide_dict["text"]
    attr = slide_dict.get("attribution")

    # Decorative left border
    border_l = grid.col_x(1)
    border_t = emu(Inches(1.8))
    border_h = emu(Inches(3.5))
    renderer.vertical_line(
        slide, border_l, border_t, border_h,
        tokens["color"]["quote_border"], width_pt=5,
    )

    # Giant opening quotation mark
    q_size = pt(tokens["font"]["size"]["quote"]) * 3
    renderer.textbox(
        slide,
        border_l - emu(Inches(0.1)),
        border_t - emu(Inches(0.5)),
        emu(Inches(1.2)),
        emu(Inches(1.2)),
        "\u201C",
        font_size_pt=q_size,
        color=tokens["color"]["quote_border"],
        bold=False,
        align="left",
        anchor="top",
        font_family=renderer.accent_font,
        cn_font_family=renderer.cn_heading_font,
    )

    # Body quote (col 2-10)
    ql, qt, qw, qh = grid.col_box(1, 10, border_t + emu(Inches(0.2)), emu(Inches(2.8)))
    ql += emu(Inches(0.4))
    qw -= emu(Inches(0.4))
    renderer.textbox(
        slide, emu(ql), emu(qt), emu(qw), emu(qh),
        text,
        font_size_pt=pt(tokens["font"]["size"]["quote"]),
        color=tokens["color"]["text"],
        italic=True,
        align="left",
        anchor="top",
        font_family=renderer.accent_font,
        cn_font_family=renderer.cn_body_font,
    )

    if attr:
        _, at_, aw, ah = grid.col_box(1, 10, emu(Inches(5.2)), emu(Inches(0.5)))
        renderer.textbox(
            slide,
            emu(ql),
            at_,
            emu(aw) - emu(Inches(0.4)),
            emu(ah),
            "— " + attr,
            font_size_pt=pt(tokens["font"]["size"]["quote_attr"]),
            color=tokens["color"]["text_light"],
            align="right",
            anchor="top",
            font_family=renderer.en_body_font,
            cn_font_family=renderer.cn_body_font,
        )
