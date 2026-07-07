"""Summary page – content layout but with a customizable title (default "总结")."""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import emu
from ._titlebar import draw_title_bar


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a summary slide with key takeaways as bullet highlights.

    渲染总结（summary）版式：顶部标题「核心总结/Key Takeaways」，下方高亮要点列表，每条要点前带主题色编号圆点，适合演示收尾。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    title = slide_dict.get("title", "总结")
    draw_title_bar(renderer, slide, grid, tokens, title)

    bullets = slide_dict.get("bullets", [])

    # Accent check-mark prefix replaces bullet glyphs for summary feel –
    # we achieve this by prepending "✓ " to each bullet text at level 1.
    rendered = []
    for b in bullets:
        if isinstance(b, str):
            rendered.append({"text": "✓  " + b, "level": 1})
        else:
            lvl = b.get("level", 1)
            new_b = dict(b)
            if lvl == 1:
                new_b["text"] = "✓  " + b.get("text", "")
            rendered.append(new_b)

    body_l, body_t, body_w, body_h = grid.content_area()
    pad = emu(Inches(0.1))
    renderer.bullets(
        slide,
        emu(body_l),
        emu(body_t) + pad,
        emu(body_w),
        emu(body_h) - pad,
        rendered,
        color=tokens["color"]["text"],
    )
