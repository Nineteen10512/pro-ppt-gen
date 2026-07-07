"""Deterministic layout calculator.

Provides a 12-column grid system, text-height estimation, auto-fit scaling,
and content-area geometry helpers.  All coordinates are derived from tokens;
layout handlers MUST NOT hard-code left/top values.
"""
from __future__ import annotations

from pptx.util import Emu, Pt


class GridLayout:
    """12-column grid over the slide safe area."""

    def __init__(self, tokens: dict):
        m = tokens["margin"]
        g = tokens["grid"]
        tb = tokens["title_bar"]
        ft = tokens["footer"]

        self.margin_top = m["top"]
        self.margin_bottom = m["bottom"]
        self.margin_left = m["left"]
        self.margin_right = m["right"]
        self.cols: int = g["cols"]
        self.gutter = g["gutter"]
        self.title_bar_h = tb["height"]
        self.title_bar_gap = tb["gap_after"]
        self.footer_h = ft["height"]

        self.slide_w = tokens["slide"]["w"]
        self.slide_h = tokens["slide"]["h"]

        # Content region (inside margins)
        self.content_x = self.margin_left
        self.content_y = self.margin_top
        self.content_w = self.slide_w - self.margin_left - self.margin_right
        self.content_h = self.slide_h - self.margin_top - self.margin_bottom

        # Column width: (content_w - (cols-1)*gutter) / cols
        total_gutter = self.gutter * (self.cols - 1)
        self.col_w_emu = (self.content_w - total_gutter) // self.cols

    # ------------------------------------------------------------------
    # Grid coordinate helpers
    # ------------------------------------------------------------------
    def col_x(self, col_index: int) -> int:
        """Return left X (EMU) for column ``col_index`` (0-based)."""
        if not 0 <= col_index < self.cols:
            # English: f"col_index {col_index} out of range [0, {self.cols})"
            raise ValueError(
                f"列索引 {col_index} 超出范围 [0, {self.cols})。"
                f"修复建议：col_index 必须在 0 到 {self.cols - 1} 之间"
            )
        return self.content_x + col_index * (self.col_w_emu + self.gutter)

    def col_w(self, span: int) -> int:
        """Return width (EMU) spanning ``span`` columns (includes inner gutters)."""
        if not 1 <= span <= self.cols:
            # English: f"span {span} out of range [1, {self.cols}]"
            raise ValueError(
                f"列跨度 span={span} 超出范围 [1, {self.cols}]。"
                f"修复建议：span 必须在 1 到 {self.cols} 之间"
            )
        return span * self.col_w_emu + (span - 1) * self.gutter

    def col_box(self, col_start: int, col_span: int, top: int, height: int) -> tuple[int, int, int, int]:
        """Return (left, top, width, height) EMU tuple for a column-anchored box."""
        return (self.col_x(col_start), top, self.col_w(col_span), height)

    # ------------------------------------------------------------------
    # Title bar & content regions
    # ------------------------------------------------------------------
    def title_bar_rect(self) -> tuple[int, int, int, int]:
        """Full-width title bar rectangle (spans margin-left to margin-right)."""
        return (
            self.margin_left,
            self.margin_top,
            self.content_w,
            self.title_bar_h,
        )

    def title_bar_text_box(self) -> tuple[int, int, int, int]:
        """Text box inside the title bar (with horizontal padding)."""
        pad = self.gutter
        left, top, w, h = self.title_bar_rect()
        return (left + pad, top + Emu(int(Pt(4).emu)), w - 2 * pad, h - Emu(int(Pt(8).emu)))

    def body_top(self) -> int:
        """Y coordinate where body content begins (below title bar + gap)."""
        return self.margin_top + self.title_bar_h + self.title_bar_gap

    def body_height(self) -> int:
        """Usable body height (from body_top to footer top)."""
        footer_top = self.slide_h - self.margin_bottom - self.footer_h
        return footer_top - self.body_top()

    def content_area(self) -> tuple[int, int, int, int]:
        """Full body area (left, top, width, height)."""
        return (self.content_x, self.body_top(), self.content_w, self.body_height())

    def footer_rect(self) -> tuple[int, int, int, int]:
        """Footer (page-number) area."""
        h = self.footer_h
        return (
            self.margin_left,
            self.slide_h - self.margin_bottom - h,
            self.content_w,
            h,
        )


# ----------------------------------------------------------------------
# Text helpers (heuristic – used for overflow detection only; final
# wrapping is still performed by PowerPoint at render time).
# ----------------------------------------------------------------------

# Average rendered width per CJK char ≈ font_size, per ASCII char ≈ 0.55 * font_size
# (pt is 12700 EMU).
def _is_cjk(ch: str) -> bool:
    cp = ord(ch)
    return (
        0x3000 <= cp <= 0x9FFF
        or 0xFF00 <= cp <= 0xFFEF
        or 0x4E00 <= cp <= 0x9FFF
        or 0x3400 <= cp <= 0x4DBF
    )


def text_height_estimate(text: str, font_size_pt: float, width_emu: int, line_spacing: float = 1.25) -> int:
    """Rough estimate of rendered text block height in EMU.

    This is intentionally conservative; it is used for pre-flight overflow
    warnings and auto-fit size selection, not for pixel-perfect layout.
    """
    if not text:
        return 0
    pt_to_emu = 12700
    fs_emu = int(font_size_pt * pt_to_emu)
    width_pt = width_emu / pt_to_emu

    # Split by explicit newlines first
    lines = text.split("\n")
    total_lines = 0
    for line in lines:
        if not line:
            total_lines += 1
            continue
        # Estimate chars per line
        # char widths in pt
        width_used_pt = 0.0
        line_breaks = 1
        for ch in line:
            chw = font_size_pt if _is_cjk(ch) else font_size_pt * 0.55
            if width_used_pt + chw > width_pt and width_used_pt > 0:
                line_breaks += 1
                width_used_pt = chw
            else:
                width_used_pt += chw
        total_lines += line_breaks

    line_h = int(fs_emu * line_spacing)
    return total_lines * line_h


def fit_text(
    text: str,
    max_width_emu: int,
    max_height_emu: int,
    start_size_pt: float,
    min_size_pt: float = 10.0,
    step: float = 2.0,
    line_spacing: float = 1.25,
) -> float:
    """Return the largest font size ≤ start_size_pt that fits the box.

    Shrinks by ``step`` pt until the estimated height ≤ max_height, bottoming
    out at ``min_size_pt``.
    """
    size = start_size_pt
    while size > min_size_pt:
        h = text_height_estimate(text, size, max_width_emu, line_spacing)
        if h <= max_height_emu:
            return size
        size -= step
    return min_size_pt


# ----------------------------------------------------------------------
# Convenience: bullets height estimate
# ----------------------------------------------------------------------
def bullets_height_estimate(
    bullets: list,
    font_size_pt: float,
    width_emu: int,
    space_after_pt: float = 6.0,
    line_spacing: float = 1.25,
) -> int:
    """Estimate total height for a (possibly nested) bullet list."""
    pt_to_emu = 12700
    total = 0
    for b in bullets:
        if isinstance(b, str):
            text, level = b, 1
        else:
            text = b.get("text", "")
            level = b.get("level", 1)
        # indent slightly reduces usable width
        indent_emu = int({1: 0.25, 2: 0.5, 3: 0.75}.get(level, 0.25) * 914400)
        w = max(width_emu - indent_emu, int(0.5 * 914400))
        total += text_height_estimate(text, font_size_pt - (level - 1) * 2, w, line_spacing)
        total += int(space_after_pt * pt_to_emu)
    return total
