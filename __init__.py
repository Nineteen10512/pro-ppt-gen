"""PRO-PPTX v1.5.3 — PaperJSX semantic compilation architecture for professional PPT generation.

Pro Ppt Gen (internal package name ``pro_ppt_gen``, product name **PRO-PPTX**)
produces native editable 16:9 ``.pptx`` files from high-level semantic JSON or
Markdown. LLMs never write coordinates / font sizes / color values — the
deterministic 12-column grid layout engine and design token system compute
those automatically.

v1.5.3 highlights (2026-07):
- 12 套内置场景模板（template_name 参数）：高考复习/答辩/路演/年会/培训等
- 本地 WPS/Office 模板主题自动提取（template_path 参数）
- 18 种元素入场/强调/退出动画（animation 字段）
- 6 种扩展切换效果（cube/zoom/reveal/ferris/gallery/conveyor）
- 模板-内容调性自动匹配提示（auto_taste_match）

v1.5 highlights:
- taste_check: Impeccable/Taste-inspired preflight for deck craft.

本文件仅维护对外导入符号与版本号，**不修改包名/import 路径/函数签名**以保证 100% 向后兼容。
"""
from __future__ import annotations

import importlib
import sys


def _ensure_shared_alias() -> None:
    """Support absolute ``shared`` imports when installed under ``skills.*``."""
    if "shared" in sys.modules:
        return
    try:
        importlib.import_module("shared")
        return
    except ImportError:
        pass
    try:
        sys.modules["shared"] = importlib.import_module("skills.shared")
    except ImportError:
        pass


_ensure_shared_alias()

__version__ = "1.5.3"

from .ppt_jsx import (
    generate,
    generate_from_markdown,
    auto_layout,
    outline,
    estimate_length,
    update_slide,
    add_slide,
    delete_slide,
    collect_image_prompts,
    quality_check,
    taste_check,
    assertive_title,
    story_check,
    mini_bar,
    mini_line,
    mini_pie,
    mini_chart_svg,
    list_themes,
    check_template_match,
)


def list_templates() -> list:
    """v1.5.3: 返回所有内置 PPT 模板的简要信息列表。"""
    from .templates import list_templates as _lt
    return _lt()


def get_template(name):
    """v1.5.3: 按名称返回 PPTTemplate 数据对象（只读）。"""
    from .templates import get_template as _gt
    return _gt(name)


def scan_local_templates(dirs=None, recursive: bool = True) -> list:
    """v1.5.3: 扫描本地 WPS/Office 模板目录，返回 PPT 类 TemplateInfo 列表。"""
    from shared.import_helper import import_shared

    (_slt,) = import_shared("template_scanner", attrs=["scan_local_templates"])
    return [t for t in _slt(dirs=dirs, recursive=recursive) if getattr(t, "endpoint", None) == "ppt"]


def extract_template_theme(path: str):
    """v1.5.3: 从本地 .pptx/.potx 文件提取主题色/字体/背景 → ExtractedTheme。"""
    from shared.import_helper import import_shared

    (_ett,) = import_shared("template_scanner", attrs=["extract_template_theme"])
    return _ett(path)


__all__ = [
    "generate",
    "generate_from_markdown",
    "auto_layout",
    "outline",
    "estimate_length",
    "update_slide",
    "add_slide",
    "delete_slide",
    "collect_image_prompts",
    "quality_check",
    "taste_check",
    "assertive_title",
    "story_check",
    "mini_bar",
    "mini_line",
    "mini_pie",
    "mini_chart_svg",
    "list_themes",
    "list_templates",
    "get_template",
    "scan_local_templates",
    "extract_template_theme",
    "check_template_match",
    "__version__",
]
