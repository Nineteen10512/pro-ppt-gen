"""Timeline layout: horizontal timeline with nodes + title + description.

Semantic fields::

    {
      "layout": "timeline",
      "title": "页面标题",
      "events": [
        {"date": "2020", "title": "里程碑标题", "desc": "简要描述（可选）"},
        ...
      ]
    }

Events are placed horizontally in order, up to 6 recommended. Each node
is a filled circle on a horizontal line (secondary color), with date above
and title+description below.
"""
from __future__ import annotations

from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE

from ._helpers import pt, emu
from ._titlebar import draw_title_bar


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a horizontal timeline slide with milestones on a line.

    渲染时间轴（timeline）版式：水平中轴线+圆点里程碑，milestones 列表按年份/时间点均匀分布在 12 列网格上，每个点上下交替显示标题+描述。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    events = slide_dict.get("events", [])
    n = len(events)

    body_l, body_t, body_w, body_h = grid.content_area()
    body_l = emu(body_l)
    body_t = emu(body_t) + emu(Inches(0.3))
    body_w = emu(body_w)
    body_h = emu(body_h) - emu(Inches(0.4))

    # Horizontal line at vertical middle (between date above and title below)
    line_y = body_t + body_h // 2
    line_color = tokens["color"]["secondary"]
    renderer.rect(
        slide,
        body_l + emu(Inches(0.3)),
        line_y - Pt(2).emu,
        body_w - emu(Inches(0.6)),
        Pt(4).emu,
        fill_color=line_color,
    )

    # Node positions: evenly distributed
    date_size = pt(tokens["font"]["size"]["body_l2"])
    title_size = pt(tokens["font"]["size"]["body_l1"])
    desc_size = pt(tokens["font"]["size"]["caption"])

    pad = emu(Inches(0.3))
    avail_w = body_w - 2 * pad
    if n == 1:
        xs = [body_l + body_w // 2]
    else:
        step = avail_w // (n - 1)
        xs = [body_l + pad + i * step for i in range(n)]

    node_d = emu(Inches(0.28))
    node_r = node_d // 2

    for i, ev in enumerate(events):
        cx = xs[i]
        # Node circle
        node = slide.shapes.add_shape(
            MSO_SHAPE.OVAL,
            cx - node_r, line_y - node_r, node_d, node_d,
        )
        node.fill.solid()
        node.fill.fore_color.rgb = tokens["color"]["accent"]
        node.line.color.rgb = tokens["color"]["bg"]
        node.line.width = Pt(2)
        node.shadow.inherit = False

        # Date label above the line
        date_text = str(ev.get("date", ""))
        date_h = emu(Inches(0.4))
        date_w = emu(Inches(1.6))
        renderer.textbox(
            slide,
            cx - date_w // 2,
            line_y - date_h - emu(Inches(0.1)),
            date_w, date_h,
            date_text,
            font_size_pt=date_size,
            color=tokens["color"]["accent"],
            bold=True,
            align="center",
            anchor="bottom",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )

        # Event title below the line
        title_str = str(ev.get("title", ""))
        title_h = emu(Inches(0.4))
        title_w = emu(Inches(1.8))
        renderer.textbox(
            slide,
            cx - title_w // 2,
            line_y + emu(Inches(0.12)),
            title_w, title_h,
            title_str,
            font_size_pt=title_size,
            color=tokens["color"]["primary"],
            bold=True,
            align="center",
            anchor="top",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )

        # Description (if any) below title
        desc = ev.get("desc") or ev.get("description") or ""
        if desc:
            desc_h = emu(Inches(0.8))
            desc_w = emu(Inches(1.9))
            renderer.textbox(
                slide,
                cx - desc_w // 2,
                line_y + emu(Inches(0.55)),
                desc_w, desc_h,
                str(desc),
                font_size_pt=desc_size,
                color=tokens["color"]["text_light"],
                bold=False,
                align="center",
                anchor="top",
                font_family=renderer.en_body_font,
                cn_font_family=renderer.cn_body_font,
            )
