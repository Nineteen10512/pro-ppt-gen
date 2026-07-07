"""Full-page chart layout: title bar + native OOXML chart body.

Semantic fields (LLM writes only these – no coordinates/colors):

    {
      "layout": "chart",
      "title": "页面标题（必填，显示在标题栏）",
      "chart": { ...chart semantic spec... },       # required
      "caption": "可选图注（显示在图表下方）",
      "bullets":  [ ...optional bullets shown as insight under chart... ]
    }
"""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu
from ._titlebar import draw_title_bar

# Chart renderer lives in engine; we import lazily to keep layouts decoupled.
from ..engine.chart_renderer import render_chart


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a slide dominated by a single chart with title and caption.

    渲染图表（chart）版式：标题栏下居中放置图表（8 列宽 × 主体高度），下方可选 caption 描述。图表类型由 chart.kind 指定（bar/line/pie/area/scatter/radar）。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    body_l, body_t, body_w, body_h = grid.content_area()
    body_l = emu(body_l)
    body_t = emu(body_t) + emu(Inches(0.1))
    body_w = emu(body_w)
    body_h = emu(body_h) - emu(Inches(0.1))

    # Optional caption above the chart
    caption = slide_dict.get("caption")
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

    # Optional insight bullets below the chart
    bullets = slide_dict.get("bullets") or []
    chart_top_pad = emu(Inches(0.05))
    body_t += chart_top_pad
    body_h -= chart_top_pad

    if bullets:
        # Reserve bottom strip for up to 3 bullets
        fs = tokens["font"]["size"]["body_l2"]
        fs_pt = pt(fs)
        # Estimate bullets height
        from ..engine.layout import bullets_height_estimate
        bullets_h = bullets_height_estimate(bullets, fs_pt, body_w, space_after_pt=4.0)
        # Cap bullets to 35% of body
        max_bullets_h = int(body_h * 0.35)
        bullets_h = min(bullets_h + emu(Inches(0.1)), max_bullets_h)
        chart_h = body_h - bullets_h - emu(Inches(0.1))
        render_chart(
            slide, slide_dict["chart"],
            body_l, body_t, body_w, chart_h,
            tokens, cn_font=renderer.cn_body_font,
        )
        # Insight label
        insight_y = body_t + chart_h + emu(Inches(0.05))
        renderer.bullets(
            slide, body_l, insight_y, body_w, bullets_h,
            bullets, color=tokens["color"]["text"],
        )
    else:
        render_chart(
            slide, slide_dict["chart"],
            body_l, body_t, body_w, body_h,
            tokens, cn_font=renderer.cn_body_font,
        )
