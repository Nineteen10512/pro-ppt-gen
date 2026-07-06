"""Input validation and pre-flight warnings.

Checks required fields per layout, type correctness, and raises early /
emits warnings for potential overflow conditions (too many bullets,
bullets that are too long, etc.).

v1.4 P3-8: 所有用户可见错误中文化，并附简短修复建议；英文原消息作为注释保留。
"""
from __future__ import annotations

from typing import Any

# Schema: per-layout required / optional field lists
# @since v1.3.0: image_path is optional if image_prompt present; added
#   decorations / background_shape / free_shape (bullets-array entries) support.
SCHEMA: dict[str, dict[str, Any]] = {
    "cover": {
        "required": ["title"],
        "optional": ["subtitle", "author", "date", "institution", "notes",
                     "speaker_notes", "background_shape", "transition", "animation"],
        "max_bullets": None,
    },
    "toc": {
        "required": ["items"],
        "optional": ["title", "notes", "speaker_notes", "decorations", "transition", "animation"],
        "max_items": 8,
    },
    "section": {
        "required": ["number", "title"],
        "optional": ["notes", "speaker_notes", "background_shape", "decorations", "transition", "animation"],
    },
    "content": {
        "required": ["title"],
        "optional": ["bullets", "chart", "notes", "speaker_notes", "decorations", "transition",
                     "free_shape", "svg_shape", "animation"],
        "max_bullets": 8,
    },
    "two_column": {
        "required": ["title", "left", "right"],
        "optional": ["notes", "speaker_notes", "decorations", "transition", "animation"],
        "max_bullets_per_col": 6,
    },
    "image_text": {
        "required": ["title", "bullets"],
        "optional": ["image_path", "image_prompt", "side", "notes", "speaker_notes",
                     "decorations", "transition", "animation"],
        "max_bullets": 6,
        "requires_one_of": [("image_path", "image_prompt")],
    },
    "full_image": {
        "required": [],
        "optional": ["image_path", "image_prompt", "title", "overlay", "notes",
                     "speaker_notes", "background_shape", "transition", "animation"],
        "requires_one_of": [("image_path", "image_prompt")],
    },
    "table": {
        "required": ["title", "headers", "rows"],
        "optional": ["caption", "notes", "speaker_notes", "decorations", "transition", "animation"],
        "max_rows": 12,
    },
    "chart": {
        "required": ["title", "chart"],
        "optional": ["caption", "bullets", "notes", "speaker_notes", "decorations", "transition", "animation"],
        "max_bullets": 3,
    },
    "kpi": {
        "required": ["title", "items"],
        "optional": ["notes", "speaker_notes", "background_shape", "decorations", "transition", "animation"],
        "min_items": 2,
        "max_items": 4,
    },
    "quote": {
        "required": ["text"],
        "optional": ["attribution", "notes", "speaker_notes", "background_shape",
                     "decorations", "transition", "animation"],
    },
    "summary": {
        "required": [],
        "optional": ["title", "bullets", "chart", "notes", "speaker_notes",
                     "decorations", "transition", "animation"],
        "max_bullets": 8,
    },
    "thanks": {
        "required": [],
        "optional": ["title", "subtitle", "notes", "speaker_notes", "background_shape",
                     "decorations", "transition", "signature", "qr_code", "contact", "animation"],
    },
    "timeline": {
        "required": ["title", "events"],
        "optional": ["notes", "speaker_notes", "decorations", "transition", "animation"],
        "max_events": 6,
    },
    # v1.4 P1-4b: 参考文献页
    "references": {
        "required": ["items"],
        "optional": ["title", "citation_style", "notes", "speaker_notes",
                     "decorations", "transition", "animation"],
        "max_items": 30,
    },
}

# v1.5.3 B-5: Allowed animation preset names
_ALLOWED_ANIM_ENTRANCE = {"fade_in","appear","fly_in_left","fly_in_right","fly_in_bottom",
                          "wipe_right","wipe_up","zoom_in","float_in","split_horizontal"}
_ALLOWED_ANIM_EMPHASIS = {"pulse","grow_shrink","color_shift","underline_reveal"}
_ALLOWED_ANIM_EXIT = {"fade_out","fly_out_right","zoom_out","dissolve_out"}
_ALLOWED_ANIM_PRESETS = _ALLOWED_ANIM_ENTRANCE | _ALLOWED_ANIM_EMPHASIS | _ALLOWED_ANIM_EXIT
_ALLOWED_ANIM_SEQUENCE = {"one_by_one","all_at_once","by_paragraph","by_series"}
_ALLOWED_ANIM_TRIGGER = {"on_click","auto_after","with_prev","after_prev"}

# Heuristic thresholds
MAX_BULLET_LEN = 40  # chars – triggers warning (CJK counted as 1)

# Supported semantic chart types
SUPPORTED_CHART_TYPES = {
    "column", "bar", "stacked_column", "stacked_bar",
    "line", "line_markers", "pie", "doughnut", "area", "scatter", "radar",
}


class ValidationError(ValueError):
    """Raised when semantic JSON is structurally invalid（中文错误）。"""


class ValidationWarning:
    """Non-fatal issues that may cause overflow or layout degradation（保留英文，内部用）。"""

    def __init__(self, slide_idx: int, layout: str, message: str):
        self.slide_idx = slide_idx
        self.layout = layout
        self.message = message

    def __repr__(self):
        return f"[slide {self.slide_idx + 1} | {self.layout}] {self.message}"


# ─── Chart spec validation ────────────────────────────────────────

def _validate_chart_spec(slide_idx: int, layout: str, chart: Any,
                         where: str = "chart") -> list[ValidationWarning]:
    """Validate a semantic chart dict.  Raises ValidationError on hard errors."""
    warnings: list[ValidationWarning] = []
    if not isinstance(chart, dict):
        # English: f"slide {i} ({layout}) {where} must be a dict"
        raise ValidationError(
            f"第 {slide_idx + 1} 页（{layout}）的 {where} 字段必须是 dict。"
            f"修复建议：请使用 {{\"type\": \"column\", \"categories\": [...], \"series\": [...]}} 结构"
        )
    ctype = chart.get("type")
    if ctype not in SUPPORTED_CHART_TYPES:
        # English: f"...{where}.type={ctype!r} invalid; supported: ..."
        raise ValidationError(
            f"第 {slide_idx + 1} 页（{layout}）的 {where}.type={ctype!r} 不是受支持的图表类型。"
            f"修复建议：请从以下类型中选择：{sorted(SUPPORTED_CHART_TYPES)}"
        )
    # categories
    if ctype != "scatter":
        cats = chart.get("categories")
        if not isinstance(cats, list) or not cats:
            # English: f"...{where}.categories must be a non-empty list"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）的 {where}.categories 必须是非空列表。"
                f"修复建议：请为分类轴传入至少一个标签，例如 [\"Q1\", \"Q2\", \"Q3\"]"
            )
        if not all(isinstance(c, str) for c in cats):
            # English: f"...{where}.categories must be string array"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）的 {where}.categories 必须是字符串数组。"
                f"修复建议：categories 元素应为字符串标签"
            )
    else:
        cats = chart.get("categories")  # optional for scatter
    # series
    series = chart.get("series")
    if not isinstance(series, list) or not series:
        # English: f"...{where}.series must be a non-empty list"
        raise ValidationError(
            f"第 {slide_idx + 1} 页（{layout}）的 {where}.series 必须是非空列表。"
            f"修复建议：请至少传入一个数据系列 {{\"name\": \"...\", \"values\": [...]}}"
        )
    n_cats = len(cats) if isinstance(cats, list) else 0
    for s_idx, s in enumerate(series):
        if not isinstance(s, dict):
            # English: f"...series[{s_idx}] must be a dict"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）的 {where}.series[{s_idx}] 必须是 dict。"
                f"修复建议：每个系列应为对象，至少包含 name 和 values 字段"
            )
        name = s.get("name")
        if not isinstance(name, str) or not name:
            # English: f"...series[{s_idx}].name is required"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）的 {where}.series[{s_idx}].name 为必填项。"
                f"修复建议：请为该数据系列提供字符串名称"
            )
        values = s.get("values")
        if not isinstance(values, list) or not values:
            # English: f"...series[{s_idx}].values must be non-empty list"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）的 {where}.series[{s_idx}].values 必须是非空列表。"
                f"修复建议：请传入与 categories 等长的数值数组"
            )
        if ctype == "scatter":
            if not all(isinstance(p, dict) and "x" in p and "y" in p for p in values):
                warnings.append(ValidationWarning(
                    slide_idx, layout,
                    f"{where}.series[{s_idx}] scatter values should be [{{x,y}}, ...]",
                ))
        else:
            if len(values) != n_cats:
                # English: f"...values length {len(values)} != categories length {n_cats}"
                raise ValidationError(
                    f"第 {slide_idx + 1} 页（{layout}）的 {where}.series[{s_idx}].values 长度 "
                    f"{len(values)} 与 categories 长度 {n_cats} 不一致。"
                    f"修复建议：请确保每个系列的数值个数与分类数相同"
                )
            for v in values:
                if not isinstance(v, (int, float)):
                    # English: f"...series[{s_idx}] values must be numeric"
                    raise ValidationError(
                        f"第 {slide_idx + 1} 页（{layout}）的 {where}.series[{s_idx}] 数值必须是数字。"
                        f"修复建议：values 数组的元素应为 int 或 float"
                    )
    return warnings


def _check_fields(slide: dict, slide_idx: int, schema: dict) -> list[ValidationWarning]:
    warnings: list[ValidationWarning] = []
    for f in schema.get("required", []):
        if f not in slide:
            one_of_groups = schema.get("requires_one_of", [])
            ok = False
            for grp in one_of_groups:
                if f in grp and any(g in slide for g in grp if g != f):
                    ok = True; break
            if not ok:
                # English: f"slide {i+1} ({layout}) missing required field: {f!r}"
                layout = slide.get("layout", "?")
                raise ValidationError(
                    f"第 {slide_idx + 1} 页（{layout}）缺少必填字段：{f!r}。"
                    f"修复建议：请补充该字段，或查阅 SKILL.md 查看 {layout} 版式所需字段"
                )
    # requires_one_of
    for grp in schema.get("requires_one_of", []):
        if not any(g in slide for g in grp):
            layout = slide.get("layout", "?")
            # English: f"slide {i+1} ({layout}) requires one of {list(grp)}"
            raise ValidationError(
                f"第 {slide_idx + 1} 页（{layout}）必须提供以下字段之一：{list(grp)}。"
                f"修复建议：请至少传入其中一个字段"
            )
    allowed = set(schema.get("required", [])) | set(schema.get("optional", []))
    extra = set(slide.keys()) - allowed - {"layout"}
    if extra:
        warnings.append(
            ValidationWarning(slide_idx, slide["layout"], f"unknown fields ignored: {sorted(extra)}")
        )
    return warnings


def _bullet_len_warnings(bullets: list, slide_idx: int, layout: str) -> list[ValidationWarning]:
    out: list[ValidationWarning] = []
    for i, b in enumerate(bullets):
        text = b if isinstance(b, str) else b.get("text", "")
        if len(text) > MAX_BULLET_LEN:
            out.append(
                ValidationWarning(
                    slide_idx,
                    layout,
                    f"bullet #{i + 1} is {len(text)} chars (>{MAX_BULLET_LEN}) – may wrap awkwardly",
                )
            )
    return out


# Supported meta.logo_position values
SUPPORTED_LOGO_POSITIONS = {"tl", "tr", "bl", "br", "center"}
# Supported meta.page_number_style values
SUPPORTED_PAGE_NUMBER_STYLES = {"plain", "slash", "chinese", "roman"}
# Meta defaults (applied in parser.normalize_meta)
# NOTE: page_number_style is intentionally NOT defaulted here – when unspecified
# the renderer falls back to legacy "N / total" format for 100% backward compat.
# Callers must explicitly opt-in to "plain" for single-digit output.
META_DEFAULTS = {
    "logo_position": "tr",
}


def _to_roman(n: int) -> str:
    """Convert a small positive integer (1..99) to uppercase Roman numeral."""
    vals = [
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s); n -= v
    return "".join(out)


def validate_meta(content: dict) -> list[ValidationWarning]:
    """Validate top-level ``meta`` dict (logo / page_number_style etc.).

    Returns warnings; raises ValidationError on hard errors.
    Best-effort: missing meta or any unknown field is silently ignored.
    """
    warnings: list[ValidationWarning] = []
    meta = content.get("meta") if isinstance(content, dict) else None
    if not isinstance(meta, dict):
        return warnings

    # logo field – must be a non-empty string if present
    logo = meta.get("logo")
    if logo is not None and (not isinstance(logo, str) or not logo.strip()):
        # English: meta.logo must be a non-empty string (path to image)
        raise ValidationError(
            "meta.logo 必须是本地图片路径字符串（非空）。修复建议：请传入 logo 图片的本地路径，或移除该字段"
        )

    pos = meta.get("logo_position")
    if pos is not None and pos not in SUPPORTED_LOGO_POSITIONS:
        raise ValidationError(
            f"meta.logo_position={pos!r} 不合法。修复建议：请从 {sorted(SUPPORTED_LOGO_POSITIONS)} 中选择"
        )

    style = meta.get("page_number_style")
    if style is not None and style not in SUPPORTED_PAGE_NUMBER_STYLES:
        raise ValidationError(
            f"meta.page_number_style={style!r} 不合法。"
            f"修复建议：请从 {sorted(SUPPORTED_PAGE_NUMBER_STYLES)} 中选择"
        )

    return warnings


def validate_slides(content: dict) -> list[ValidationWarning]:
    """Validate a semantic PPT content dict; raise ValidationError on failure.

    校验 PPT 语义 content：slides 列表、每页 layout 是否在 SCHEMA 中、各版式必填字段、meta 字段
    （logo/logo_position/page_number_style/transition）。v1.4 起 SUPPORTED_CHART_TYPES 补齐 area/
    scatter/radar 以兼容旧冒烟测试。"""
    if not isinstance(content, dict):
        # English: "top-level content must be a dict"
        raise ValidationError(
            "顶层 content 必须是 dict。修复建议：请将输入包装为 {\"slides\": [...]} 结构"
        )
    slides = content.get("slides")
    if not isinstance(slides, list) or not slides:
        # English: "content.slides must be a non-empty list"
        raise ValidationError(
            "content.slides 必须是非空列表。修复建议：请在 slides 数组中至少添加 1 页幻灯片"
        )

    warnings: list[ValidationWarning] = []
    for i, slide in enumerate(slides):
        if not isinstance(slide, dict):
            # English: f"slide {i+1} must be a dict"
            raise ValidationError(
                f"第 {i+1} 页必须是 dict。修复建议：每一页应为 {{\"layout\": \"...\", \"title\": \"...\", ...}} 结构"
            )
        layout = slide.get("layout")
        if layout not in SCHEMA:
            # English: f"slide {i+1} has unknown layout {layout!r}; valid: ..."
            raise ValidationError(
                f"第 {i+1} 页的版式 {layout!r} 不存在。"
                f"修复建议：请从以下版式中选择：{sorted(SCHEMA)}"
            )
        schema = SCHEMA[layout]
        warnings.extend(_check_fields(slide, i, schema))

        # v1.5.3 B-5: validate animation field if present
        anim = slide.get("animation")
        if anim is not None:
            # Shortcut: a bare string preset name → expanded dict
            if isinstance(anim, str):
                if anim not in _ALLOWED_ANIM_PRESETS:
                    raise ValidationError(
                        f"第 {i+1} 页（{layout}）animation={anim!r} 不合法。"
                        f"修复建议：请从 {sorted(_ALLOWED_ANIM_PRESETS)} 中选择一个预设名，"
                        f"或传入 dict 形式 {{\"entrance\": \"fade_in\", ...}}"
                    )
                # rewrite to dict in place
                anim = {"entrance": anim} if anim in _ALLOWED_ANIM_ENTRANCE else \
                       ({"emphasis": anim} if anim in _ALLOWED_ANIM_EMPHASIS else {"exit": anim})
                slide["animation"] = anim
            if not isinstance(anim, dict):
                raise ValidationError(
                    f"第 {i+1} 页（{layout}）的 animation 字段必须是 dict 或预设名字符串。"
                    f"修复建议：请使用 {{\"entrance\": \"fade_in\", \"sequence\": \"by_paragraph\"}} 格式"
                )
            for key, allowed, field_name in [
                ("entrance", _ALLOWED_ANIM_ENTRANCE, "entrance"),
                ("emphasis", _ALLOWED_ANIM_EMPHASIS, "emphasis"),
                ("exit", _ALLOWED_ANIM_EXIT, "exit"),
            ]:
                v = anim.get(key)
                if v is not None and v not in allowed:
                    raise ValidationError(
                        f"第 {i+1} 页（{layout}）animation.{key}={v!r} 不合法。"
                        f"修复建议：请从 {sorted(allowed)} 中选择"
                    )
            seq = anim.get("sequence", "all_at_once")
            if seq not in _ALLOWED_ANIM_SEQUENCE:
                raise ValidationError(
                    f"第 {i+1} 页（{layout}）animation.sequence={seq!r} 不合法。"
                    f"修复建议：请从 {sorted(_ALLOWED_ANIM_SEQUENCE)} 中选择"
                )
            tri = anim.get("trigger", "on_click")
            if tri not in _ALLOWED_ANIM_TRIGGER:
                raise ValidationError(
                    f"第 {i+1} 页（{layout}）animation.trigger={tri!r} 不合法。"
                    f"修复建议：请从 {sorted(_ALLOWED_ANIM_TRIGGER)} 中选择"
                )

        # Layout-specific checks
        if layout in ("content", "summary", "image_text"):
            bullets = slide.get("bullets", [])
            chart = slide.get("chart")
            if layout == "content" and not bullets and chart is None:
                # English: f"slide {i+1} (content) must provide at least one of: bullets, chart"
                raise ValidationError(
                    f"第 {i+1} 页（content）必须提供 bullets 或 chart 至少其一。"
                    f"修复建议：添加要点列表，或插入一个 chart 图表"
                )
            max_b = schema.get("max_bullets")
            if max_b and len(bullets) > max_b:
                warnings.append(
                    ValidationWarning(
                        i, layout, f"{len(bullets)} bullets (recommended ≤{max_b}) – may overflow"
                    )
                )
            warnings.extend(_bullet_len_warnings(bullets, i, layout))
            if slide.get("chart") is not None:
                warnings.extend(_validate_chart_spec(i, layout, slide["chart"], where="chart"))
                if layout == "content" and bullets and len(bullets) > 6:
                    warnings.append(ValidationWarning(
                        i, layout,
                        "bullets+chart coexist; recommend ≤3 bullets to avoid crowding",
                    ))

        elif layout == "chart":
            chart = slide.get("chart")
            warnings.extend(_validate_chart_spec(i, layout, chart, where="chart"))
            bullets = slide.get("bullets", [])
            max_b = schema.get("max_bullets", 3)
            if len(bullets) > max_b:
                warnings.append(ValidationWarning(
                    i, layout,
                    f"{len(bullets)} insight bullets on chart page (recommended ≤{max_b})",
                ))
            warnings.extend(_bullet_len_warnings(bullets, i, layout))

        elif layout == "two_column":
            for side in ("left", "right"):
                col = slide.get(side) or {}
                bullets = col.get("bullets", [])
                chart_spec = col.get("chart")
                max_b = schema.get("max_bullets_per_col", 6)
                if chart_spec is None and len(bullets) > max_b:
                    warnings.append(
                        ValidationWarning(
                            i, layout, f"{side} column has {len(bullets)} bullets (> {max_b})"
                        )
                    )
                warnings.extend(_bullet_len_warnings(bullets, i, layout))
                if chart_spec is not None:
                    warnings.extend(_validate_chart_spec(
                        i, layout, chart_spec, where=f"{side}.chart",
                    ))

        elif layout == "toc":
            items = slide.get("items", [])
            if len(items) > schema.get("max_items", 8):
                # English: f"slide {i+1} (toc) has {len(items)} items (max {schema['max_items']})"
                raise ValidationError(
                    f"第 {i+1} 页（toc）目录项数 {len(items)} 超过上限 {schema['max_items']}。"
                    f"修复建议：请将目录拆分为多页，或减少顶级章节数"
                )

        elif layout == "references":
            items = slide.get("items", [])
            if not isinstance(items, list) or not items:
                raise ValidationError(
                    f"第 {i+1} 页（references）的 items 必须是非空列表。"
                    f"修复建议：请传入参考文献条目数组，例如 [{{\"type\":\"article\",\"authors\":\"...\"}}]"
                )
            max_items = schema.get("max_items", 30)
            if len(items) > max_items:
                warnings.append(ValidationWarning(
                    i, layout,
                    f"{len(items)} references (recommended ≤{max_items}) – may overflow; extra pages will be added automatically",
                ))
            # Validate citation_style if provided
            cs = slide.get("citation_style")
            if cs is not None:
                cs_norm = {"harvard": "apa", "gbt7714": "gb7714"}.get(cs.lower(), cs.lower())
                if cs_norm not in {"apa", "gb7714", "mla", "ieee"}:
                    warnings.append(ValidationWarning(
                        i, layout,
                        f"unknown citation_style {cs!r}; falling back to gb7714",
                    ))

        elif layout == "kpi":
            items = slide.get("items", [])
            if not (schema["min_items"] <= len(items) <= schema["max_items"]):
                # English: f"slide {i+1} (kpi) must have {min}-{max} items, got {len}"
                raise ValidationError(
                    f"第 {i+1} 页（kpi）指标数量 {len(items)} 不合法，"
                    f"应为 {schema['min_items']}-{schema['max_items']} 个。"
                    f"修复建议：请调整 items 数组长度"
                )

        elif layout == "thanks":
            for img_field in ("signature", "qr_code"):
                v = slide.get(img_field)
                if v is not None and (not isinstance(v, str) or not v.strip()):
                    raise ValidationError(
                        f"第 {i+1} 页（thanks）的 {img_field} 字段必须是非空字符串（图片路径）。"
                        f"修复建议：请传入本地图片路径，或移除该字段"
                    )
            contact = slide.get("contact")
            if contact is not None and not isinstance(contact, str):
                raise ValidationError(
                    f"第 {i+1} 页（thanks）的 contact 字段必须是字符串。"
                    f"修复建议：请传入多行联系信息文本"
                )

        elif layout == "timeline":
            events = slide.get("events", [])
            max_e = schema.get("max_events", 6)
            if not isinstance(events, list) or not events:
                # English: f"slide {i+1} (timeline) events must be a non-empty list"
                raise ValidationError(
                    f"第 {i+1} 页（timeline）的 events 必须是非空列表。"
                    f"修复建议：请至少提供一个时间节点 {{\"title\": \"...\", \"date\": \"...\"}}"
                )
            if len(events) > max_e:
                warnings.append(ValidationWarning(
                    i, layout, f"{len(events)} events (> {max_e}) – consider splitting",
                ))
            for e_idx, ev in enumerate(events):
                if not isinstance(ev, dict):
                    # English: f"slide {i+1} (timeline) events[{e_idx}] must be a dict"
                    raise ValidationError(
                        f"第 {i+1} 页（timeline）的 events[{e_idx}] 必须是 dict。"
                        f"修复建议：每个时间节点应为对象"
                    )
                if not ev.get("title"):
                    # English: f"slide {i+1} (timeline) events[{e_idx}] missing required 'title'"
                    raise ValidationError(
                        f"第 {i+1} 页（timeline）的 events[{e_idx}] 缺少必填字段 'title'。"
                        f"修复建议：请为每个时间节点提供 title"
                    )

        elif layout == "table":
            rows = slide.get("rows", [])
            headers = slide.get("headers", [])
            for r_idx, r in enumerate(rows):
                if len(r) != len(headers):
                    warnings.append(
                        ValidationWarning(
                            i,
                            layout,
                            f"row {r_idx + 1} has {len(r)} cells (headers={len(headers)})",
                        )
                    )
            if len(rows) > schema.get("max_rows", 12):
                warnings.append(
                    ValidationWarning(i, layout, f"{len(rows)} rows (> {schema['max_rows']}) – consider splitting")
                )

    return warnings
