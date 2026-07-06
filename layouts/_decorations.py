"""PaperJSX PPT v1.3 layout helpers — image_prompt resolution, decorations, etc.

@since v1.3.0
"""
from __future__ import annotations

from typing import Optional

from pptx.dml.color import RGBColor

from ._helpers import emu


def _gray_placeholder(renderer, slide, x: int, y: int, w: int, h: int,
                      label: str, tokens: dict):
    """Draw a gray placeholder rectangle with a prompt label (for image_prompt-only pages)."""
    # Light gray fill + dashed border
    shp = renderer.rect(
        slide, int(x), int(y), int(w), int(h),
        fill_color=RGBColor(0xEE, 0xEE, 0xEE),
        line_color=RGBColor(0xBB, 0xBB, 0xBB),
    )
    # Add dashed outline via XML
    from pptx.oxml.ns import qn
    from lxml import etree
    ln = shp.line._get_or_add_ln()
    # remove existing prstDash
    for pd in ln.findall(qn("a:prstDash")):
        ln.remove(pd)
    prstDash = etree.SubElement(ln, qn("a:prstDash"))
    prstDash.set("val", "dash")
    # label text
    renderer.textbox(
        slide,
        int(x) + emu(50000), int(y), int(w) - emu(100000), int(h),
        f"[AI生成图]\n{label[:60]}",
        font_size_pt=11,
        color=RGBColor(0x99, 0x99, 0x99),
        align="center", anchor="middle",
        font_family=renderer.en_body_font,
        cn_font_family=renderer.cn_body_font,
    )
    return shp


def resolve_image(slide_dict: dict) -> Optional[str]:
    """Return the usable image_path for rendering, or None if only image_prompt given."""
    if slide_dict.get("image_path"):
        return slide_dict["image_path"]
    return None


def image_prompt_of(slide_dict: dict) -> Optional[str]:
    return slide_dict.get("image_prompt")


def render_image_or_placeholder(renderer, slide, slide_dict, tokens,
                                x: int, y: int, w: int, h: int,
                                fit: bool = True):
    """Render image_path if present, or a gray placeholder for image_prompt."""
    path = resolve_image(slide_dict)
    if path:
        renderer.image(slide, path, int(x), int(y), width=int(w), height=int(h), fit=fit)
    else:
        prompt = image_prompt_of(slide_dict) or ""
        _gray_placeholder(renderer, slide, int(x), int(y), int(w), int(h), prompt, tokens)


def render_decorations(renderer, slide, slide_dict, grid, tokens):
    """Render page-level decorations: free-shape SVGs positioned by grid cols."""
    decos = slide_dict.get("decorations") or []
    if not decos:
        return
    body_l = int(grid.margin_left)
    body_t = int(grid.margin_top)
    body_w = int(grid.content_w)
    body_h = int(grid.content_h)
    sw = int(grid.slide_w)
    sh = int(grid.slide_h)
    cnt = 0
    for dec in decos:
        if not isinstance(dec, dict):
            continue
        svg = dec.get("svg")
        if not svg:
            continue
        # positioning
        x = dec.get("x"); y_ = dec.get("y"); w = dec.get("w"); h = dec.get("h")
        # Support col/row based grid alignment
        col = dec.get("col"); span = dec.get("span")
        align = dec.get("align")
        if col is not None and span is not None:
            bx, by, bw, bh = grid.col_box(int(col), int(span),
                                          body_t + (dec.get("top_offset", 0) or 0),
                                          int(dec.get("height_emu", body_h)))
            x_emu, y_emu, w_emu, h_emu = emu(bx), emu(by), emu(bw), emu(bh)
        else:
            # percentage based
            def _pct(v, base, default):
                if v is None: return default
                if isinstance(v, str) and v.endswith("%"): return int(float(v[:-1])/100*base)
                return int(v)
            x_emu = _pct(x, sw, body_l)
            y_emu = _pct(y_, sh, body_t)
            w_emu = _pct(w, sw, body_w)
            h_emu = _pct(h, sh, body_h)
        if align == "center":
            x_emu = (sw - w_emu) // 2
        elif align == "right":
            x_emu = sw - w_emu - int(grid.margin_right)
        try:
            renderer.free_shape(slide, svg, x_emu, y_emu, w_emu, h_emu,
                                shape_id=100 + cnt, name=f"Deco{cnt}")
        except Exception as e:
            # don't fail the whole deck on bad SVG
            print(f"[decoration] SVG render failed: {e}")
        cnt += 1


def render_background_shape(renderer, slide, slide_dict, grid, tokens):
    """Render background_shape SVG behind the slide content (z-order 0, full-bleed)."""
    svg = slide_dict.get("background_shape")
    if not svg:
        return
    try:
        renderer.free_shape(
            slide, svg,
            0, 0, int(grid.slide_w), int(grid.slide_h),
            shape_id=1, name="BgShape",
        )
    except Exception as e:
        print(f"[background_shape] SVG render failed: {e}")


__all__ = [
    "resolve_image", "image_prompt_of",
    "render_image_or_placeholder", "render_decorations", "render_background_shape",
]
