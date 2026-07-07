"""Shared helper: render the standard title bar (primary rectangle + title text)."""
from __future__ import annotations

from ._helpers import pt, emu


def draw_title_bar(renderer, slide, grid, tokens, title: str):
    """Render a full-width primary-colored title bar with white title text."""
    tl, tt, tw, th = grid.title_bar_rect()
    renderer.rect(slide, emu(tl), emu(tt), emu(tw), emu(th), fill_color=tokens["color"]["title_bar_bg"])
    pl, pt_, pw, ph = grid.title_bar_text_box()
    renderer.textbox(
        slide, emu(pl), emu(pt_), emu(pw), emu(ph),
        title,
        font_size_pt=pt(tokens["font"]["size"]["slide_title"]),
        color=tokens["color"]["text_on_primary"],
        bold=True,
        align="left",
        anchor="middle",
        font_family=renderer.en_heading_font,
        cn_font_family=renderer.cn_heading_font,
    )
