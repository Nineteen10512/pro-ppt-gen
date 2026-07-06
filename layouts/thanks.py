"""Thanks / Q&A closing page.

v1.4 P3-3: Optional ``signature`` (signature image below author line),
``qr_code`` (QR image on the right side), and ``contact`` (multi-line
contact text rendered below the QR or signature) are now supported.
All three are optional; when omitted, layout is unchanged from v1.3.
"""
from __future__ import annotations

from pptx.util import Inches, Pt

from ._helpers import pt, emu


def _safe_image(renderer, slide, path, left, top, width=None, height=None, fit=False):
    """Best-effort image insert; returns shape or None on failure."""
    try:
        if not path or not isinstance(path, str):
            return None
        return renderer.image(slide, path, left, top, width=width, height=height, fit=fit)
    except Exception as e:
        print(f"[thanks] image insert failed ({path}): {e}")
        return None


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a thank-you / Q&A slide with optional signature, QR code, contact.

    渲染致谢页（thanks）：居中大字谢谢/Q&A；v1.4 P3-3 起支持 signature（签名区）、qr_code（二维码图片，触发左右分栏布局）、contact（联系方式列表）。缺失图片静默跳过。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])

    title = slide_dict.get("title", "感谢聆听")
    subtitle = slide_dict.get("subtitle", "Q & A")
    signature = slide_dict.get("signature")
    qr_code = slide_dict.get("qr_code")
    contact = slide_dict.get("contact")

    # Layout decision: qr_code present → split screen (text left, QR right);
    # otherwise keep the original centered layout.
    qr_size = emu(Inches(1.5))
    pad = emu(Inches(0.3))

    if qr_code:
        # Left content area width = content_w - qr_size - gap
        gap = emu(Inches(0.4))
        left_w = emu(grid.content_w) - qr_size - gap
        text_x = emu(grid.margin_left)
        text_w = left_w
        qr_x = emu(grid.margin_left) + left_w + gap
        qr_y = emu(Inches(2.6))
    else:
        text_x = emu(grid.margin_left)
        text_w = emu(grid.content_w)
        qr_x = qr_y = None

    # Central title (left-aligned when QR is on right; centered otherwise)
    title_align = "left" if qr_code else "center"
    title_size = pt(tokens["font"]["size"]["cover_title"])
    title_top = emu(Inches(2.2))
    title_h = emu(Inches(1.3))
    renderer.textbox(
        slide,
        text_x,
        title_top,
        text_w,
        title_h,
        title,
        font_size_pt=title_size,
        color=tokens["color"]["primary"],
        bold=True,
        align=title_align,
        anchor="middle",
        font_family=renderer.en_heading_font,
        cn_font_family=renderer.cn_heading_font,
    )

    # Accent divider
    div_w = emu(Inches(1.2))
    div_h = emu(Inches(0.05))
    if qr_code:
        div_x = text_x
    else:
        div_x = emu(grid.margin_left) + (emu(grid.content_w) - div_w) // 2
    divider_y = emu(Inches(3.7))
    renderer.rect(
        slide,
        div_x,
        divider_y,
        div_w,
        div_h,
        fill_color=tokens["color"]["accent"],
    )

    # Subtitle
    sub_size = pt(tokens["font"]["size"]["cover_subtitle"])
    sub_top = emu(Inches(3.9))
    sub_h = emu(Inches(0.7))
    renderer.textbox(
        slide,
        text_x,
        sub_top,
        text_w,
        sub_h,
        subtitle,
        font_size_pt=sub_size,
        color=tokens["color"]["text_light"],
        align=title_align,
        anchor="top",
        font_family=renderer.en_body_font,
        cn_font_family=renderer.cn_body_font,
    )

    # Signature (just below subtitle in the text column)
    sig_h = emu(Inches(0.6))
    sig_top = emu(Inches(4.75))
    if signature:
        # Place signature below subtitle, aligned like the text
        if qr_code:
            sig_left = text_x
            sig_w = min(emu(Inches(2.5)), text_w)
        else:
            sig_w = min(emu(Inches(2.5)), emu(grid.content_w))
            sig_left = emu(grid.margin_left) + (emu(grid.content_w) - sig_w) // 2
        _safe_image(renderer, slide, signature, sig_left, sig_top,
                    width=sig_w, height=sig_h, fit=True)

    # QR code (right side)
    if qr_code:
        _safe_image(renderer, slide, qr_code, qr_x, qr_y,
                    width=qr_size, height=qr_size, fit=True)

    # Contact text:
    # - If QR present: render below QR (right side), left-aligned with QR.
    # - Else if signature present: render below signature, centered.
    # - Else: render centered below subtitle, centered.
    if contact:
        contact_size = pt(Pt(11))
        contact_color = tokens["color"]["text_light"]
        contact_h = emu(Inches(1.2))
        if qr_code:
            contact_x = qr_x
            contact_y = qr_y + qr_size + emu(Inches(0.15))
            contact_w = qr_size + emu(Inches(0.6))
            contact_align = "center"
        elif signature:
            contact_w = emu(grid.content_w)
            contact_x = emu(grid.margin_left)
            contact_y = sig_top + sig_h + emu(Inches(0.15))
            contact_align = "center"
        else:
            contact_w = emu(grid.content_w)
            contact_x = emu(grid.margin_left)
            contact_y = sig_top
            contact_align = "center"
        renderer.textbox(
            slide,
            contact_x,
            contact_y,
            contact_w,
            contact_h,
            str(contact),
            font_size_pt=contact_size,
            color=contact_color,
            align=contact_align,
            anchor="top",
            font_family=renderer.en_body_font,
            cn_font_family=renderer.cn_body_font,
        )
