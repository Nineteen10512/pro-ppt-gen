"""KPI cards page: 2-4 large numeric cards evenly distributed.

v1.4 P2-1: each KPI item may carry a ``mini_chart`` spec
``{kind: "bar"|"line"|"pie", data: [...], w: inches?, h: inches?}`` which is
rendered as an SVG sparkline below the value/label block.
"""
from __future__ import annotations

from pptx.util import Inches

from ._helpers import pt, emu
from ._titlebar import draw_title_bar


def _build_mini_chart_svg(spec: dict, tokens) -> str:
    """Build SVG string from a mini_chart spec using theme palette."""
    _mc = None
    # sys.path typically includes skills/; prefer direct import, fall back to relative.
    try:
        from shared.svg_engine import mini_chart as _mc
    except ImportError:
        try:
            from ..shared.svg_engine import mini_chart as _mc
        except ImportError:
            pass
    if _mc is None:
        return ""
    # Build a minimal theme_colors dict from tokens (avoids needing theme name)
    color = tokens.get("color", {})
    # Build a chart palette from theme primary+accent+palette if available
    chart_palette = None
    cp = color.get("_chart_palette")
    if cp and isinstance(cp, (list, tuple)):
        chart_palette = list(cp)
    else:
        # Derive a 4-color palette from primary/accent etc.
        chart_palette = [
            _rgb_hex(color.get("primary")),
            _rgb_hex(color.get("accent")),
            _rgb_hex(color.get("secondary")) or _rgb_hex(color.get("text")),
            "#7F604F",
        ]
        chart_palette = [c for c in chart_palette if c]
    theme_colors = {
        "primary": _rgb_hex(color.get("primary")) or "#1F3864",
        "accent": _rgb_hex(color.get("accent")) or "#C0504D",
        "secondary": _rgb_hex(color.get("secondary")) or "#2E75B6",
        "text": _rgb_hex(color.get("text")) or "#333333",
        "_chart_palette": chart_palette,
    }
    kind = spec.get("kind", "bar")
    data = spec.get("data", [])
    kwargs = {"theme_colors": theme_colors}
    if "color" in spec:
        kwargs["color"] = spec["color"]
    if "fill" in spec and kind in ("line", "mini_line"):
        kwargs["fill"] = bool(spec["fill"])
    if kind in ("pie", "mini_pie") and "colors" in spec:
        kwargs["colors"] = spec["colors"]
    if kind in ("pie", "mini_pie"):
        kwargs["size"] = int(spec.get("size", 80))
    else:
        kwargs["width"] = int(spec.get("width", 140))
        kwargs["height"] = int(spec.get("height", 42))
    return _mc(kind, data, **kwargs)


def _rgb_hex(v) -> str:
    """Convert a token color value (RGBColor or str) to #rrggbb hex string."""
    if v is None:
        return ""
    try:
        # python-pptx RGBColor
        return str(v)
    except Exception:
        s = str(v).strip()
        if not s:
            return ""
        if s.startswith("#"):
            return s
        try:
            int(s, 16)
            return f"#{s}" if len(s) in (3, 6) else ""
        except Exception:
            return ""


def render(slide_dict: dict, renderer, grid, tokens):
    """Render a KPI dashboard slide with 2/3/4 metric cards and mini charts.

    渲染 KPI 指标看板版式：顶部 KPI 标题，下方 2/3/4 个指标卡片横向排列，每个卡片包含value 大号数字/label/delta 环比，可选 mini_bar/mini_line/mini_pie SVG 迷你图。"""

    slide = renderer.new_slide(bg_color=tokens["color"]["bg"])
    draw_title_bar(renderer, slide, grid, tokens, slide_dict["title"])

    items = slide_dict["items"]
    n = len(items)

    body_l, body_t, body_w, body_h = grid.content_area()
    body_t = emu(body_t) + emu(Inches(0.3))
    body_h = emu(body_h) - emu(Inches(0.6))

    span_map = {2: 5, 3: 3, 4: 2}
    span = span_map[n]
    gap_cols = 1
    total_cols_used = n * span + (n - 1) * gap_cols
    start_col = (12 - total_cols_used) // 2

    val_size = pt(tokens["font"]["size"]["kpi_value"])
    lbl_size = pt(tokens["font"]["size"]["kpi_label"])
    cap_size = pt(tokens["font"]["size"]["caption"])

    # Reserve room for optional sparkline under the label
    has_any_spark = any(isinstance(it, dict) and it.get("mini_chart") for it in items)
    card_h = emu(Inches(2.4 if has_any_spark else 2.2))
    card_top = body_t + max(0, (body_h - card_h) // 3)

    for i, it in enumerate(items):
        col = start_col + i * (span + gap_cols)
        cl, ct_, cw, ch = grid.col_box(col, span, card_top, card_h)
        cl, ct_, cw, ch = emu(cl), emu(ct_), emu(cw), emu(ch)

        # Rounded card background
        renderer.rounded_rect(slide, cl, ct_, cw, ch, fill_color=tokens["color"]["kpi_bg"])

        pad = emu(Inches(0.2))
        inner_l = cl + pad
        inner_t = ct_ + pad
        inner_w = cw - 2 * pad

        value = str(it.get("value", ""))
        label = str(it.get("label", ""))
        subtext = it.get("subtext")
        spark_spec = it.get("mini_chart") if isinstance(it, dict) else None

        val_h = emu(Inches(0.8))
        renderer.textbox(
            slide, inner_l, inner_t, inner_w, val_h,
            value,
            font_size_pt=val_size,
            color=tokens["color"]["primary"],
            bold=True,
            align="center",
            anchor="middle",
            font_family=renderer.en_heading_font,
            cn_font_family=renderer.cn_heading_font,
        )
        # Divider
        div_w = emu(Inches(0.6))
        div_h = emu(Inches(0.04))
        div_y = inner_t + val_h + emu(Inches(0.05))
        renderer.rect(
            slide,
            cl + (cw - div_w) // 2,
            div_y,
            div_w,
            div_h,
            fill_color=tokens["color"]["accent"],
        )
        # Label
        lbl_h = emu(Inches(0.35))
        lbl_y = inner_t + val_h + emu(Inches(0.15))
        renderer.textbox(
            slide, inner_l, lbl_y, inner_w, lbl_h,
            label,
            font_size_pt=lbl_size,
            color=tokens["color"]["text"],
            bold=True,
            align="center",
            anchor="middle",
            font_family=renderer.en_body_font,
            cn_font_family=renderer.cn_body_font,
        )
        cur_y = lbl_y + lbl_h
        if subtext:
            sub_h = emu(Inches(0.3))
            renderer.textbox(
                slide, inner_l, cur_y, inner_w, sub_h,
                str(subtext),
                font_size_pt=cap_size,
                color=tokens["color"]["text_light"],
                align="center",
                anchor="top",
                font_family=renderer.en_body_font,
                cn_font_family=renderer.cn_body_font,
            )
            cur_y += sub_h
        # v1.4 P2-1: optional mini sparkline chart
        if spark_spec:
            try:
                svg_str = _build_mini_chart_svg(spark_spec, tokens)
                chart_w_in = float(spark_spec.get("w", 1.3))
                chart_h_in = float(spark_spec.get("h", 0.4))
                chart_w = emu(Inches(chart_w_in))
                chart_h = emu(Inches(chart_h_in))
                chart_x = cl + (cw - chart_w) // 2
                chart_y = cur_y + emu(Inches(0.05))
                # clamp so the chart stays inside card
                if chart_y + chart_h > ct_ + ch - pad:
                    chart_y = ct_ + ch - pad - chart_h
                if chart_y < cur_y:
                    chart_y = cur_y
                renderer._render_mini_chart(slide, chart_x, chart_y, chart_w, chart_h, svg_str)
            except Exception as e:
                print(f"[kpi.mini_chart] render failed: {e}")
