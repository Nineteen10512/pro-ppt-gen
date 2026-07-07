from __future__ import annotations

import copy
from collections import Counter
from typing import Any


STRUCTURE_ELEMENTS = {
    "connector": ("connector", "arrow", "line", "flow", "step", "timeline", "连接", "箭头", "流程", "递进"),
    "node": ("node", "circle", "milestone", "status", "节点", "圆点", "里程碑", "状态"),
    "spatial_contrast": ("compare", "vs", "split", "contrast", "mirror", "对比", "分屏", "差异"),
    "ring_cycle": ("cycle", "loop", "ring", "flywheel", "环形", "循环", "闭环", "飞轮"),
    "progress_state": ("progress", "percent", "stage", "status", "进度", "阶段", "完成", "%"),
    "asymmetry": ("asymmetric", "hero", "focus", "magazine", "非对称", "主视觉", "焦点", "大数字"),
    "hierarchy": ("tree", "layer", "architecture", "matrix", "层级", "架构", "矩阵", "树形"),
    "data_visual": ("chart", "kpi", "dashboard", "trend", "图表", "数据", "看板", "趋势"),
}

CORE_LAYOUTS = {"content", "two_column", "image_text", "table", "chart", "kpi", "timeline", "summary"}
STRUCTURAL_LAYOUT_HINTS = {
    "timeline": {"connector", "node", "progress_state"},
    "chart": {"data_visual"},
    "kpi": {"data_visual", "asymmetry"},
    "table": {"hierarchy", "spatial_contrast"},
    "two_column": {"spatial_contrast"},
    "image_text": {"asymmetry"},
}


def _texts(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    if isinstance(value, dict):
        out = []
        for item in value.values():
            out.extend(_texts(item))
        return out
    if isinstance(value, list):
        out = []
        for item in value:
            out.extend(_texts(item))
        return out
    return [str(value)]


def infer_structure_elements(slide: dict) -> list[str]:
    layout = slide.get("layout") or "content"
    elements = set(STRUCTURAL_LAYOUT_HINTS.get(layout, set()))
    text = " ".join(_texts(slide)).lower()
    for name, keywords in STRUCTURE_ELEMENTS.items():
        if any(keyword.lower() in text for keyword in keywords):
            elements.add(name)
    if slide.get("chart"):
        elements.add("data_visual")
    if slide.get("events"):
        elements.update({"connector", "node", "progress_state"})
    if slide.get("items") and layout == "kpi":
        elements.update({"data_visual", "asymmetry"})
    if slide.get("left") and slide.get("right"):
        elements.add("spatial_contrast")
    if slide.get("free_shape") or slide.get("background_shape"):
        elements.add("hierarchy")
    return sorted(elements)


def _looks_like_flat_text_block(slide: dict) -> bool:
    layout = slide.get("layout") or "content"
    if layout not in {"content", "two_column", "summary"}:
        return False
    has_visual = any(slide.get(key) for key in ("chart", "image_path", "image_prompt", "events", "free_shape", "background_shape"))
    bullets = slide.get("bullets") or []
    if layout == "two_column":
        left = slide.get("left") or {}
        right = slide.get("right") or {}
        bullets = list(left.get("bullets") or []) + list(right.get("bullets") or [])
        has_visual = has_visual or bool(left.get("chart") or right.get("chart") or left.get("image_path") or right.get("image_path"))
    return bool(bullets) and not has_visual


def visual_redesign_check(content: dict, *, strict: bool = False) -> dict:
    slides = (content or {}).get("slides") or []
    issues: list[dict] = []
    slide_reports: list[dict] = []
    core_signatures: list[str] = []

    for index, slide in enumerate(slides, start=1):
        if not isinstance(slide, dict):
            continue
        layout = slide.get("layout") or "content"
        elements = infer_structure_elements(slide)
        is_core = layout in CORE_LAYOUTS
        if is_core:
            core_signatures.append("|".join([layout] + elements[:3]))
        if is_core and len(elements) < 2:
            issues.append(
                {
                    "level": "error" if strict else "warning",
                    "code": "flat_layout_risk",
                    "page": index,
                    "message": "Core slide has fewer than 2 infographic structure elements.",
                    "suggestion": "Add connectors/arrows, nodes, progress/status markers, hierarchy, asymmetry, or data visualization.",
                }
            )
        if _looks_like_flat_text_block(slide):
            issues.append(
                {
                    "level": "error",
                    "code": "word_doc_on_slide",
                    "page": index,
                    "message": "Slide looks like colored text blocks instead of an infographic layout.",
                    "suggestion": "Convert bullets into a flow, comparison, hierarchy, cycle, dashboard, or asymmetric hero layout.",
                }
            )
        title = str(slide.get("title") or "")
        if is_core and title and slide.get("title_alignment") == "center":
            issues.append(
                {
                    "level": "warning",
                    "code": "core_title_centered",
                    "page": index,
                    "message": "Core narrative slide title is centered; default should be left-aligned with top breathing room.",
                }
            )
        slide_reports.append({"page": index, "layout": layout, "structure_elements": elements, "core": is_core})

    repeated = [sig for sig, count in Counter(core_signatures).items() if count >= 3 and sig]
    if repeated:
        issues.append(
            {
                "level": "warning",
                "code": "repeated_core_layout_structure",
                "message": "Three or more core slides share the same layout structure.",
                "suggestion": "Vary core pages by relationship: flow, comparison, hierarchy, cycle, dashboard, or editorial hero.",
            }
        )

    score = max(0, 100 - sum(18 if item["level"] == "error" else 8 for item in issues))
    return {
        "version": "1.6.4",
        "score": score,
        "passed": score >= (85 if strict else 75) and not any(item["level"] == "error" for item in issues),
        "issues": issues,
        "slides": slide_reports,
    }


def apply_visual_redesign_guidance(content: dict) -> dict:
    working = copy.deepcopy(content)
    for slide in working.get("slides", []) or []:
        if not isinstance(slide, dict):
            continue
        elements = infer_structure_elements(slide)
        if len(elements) < 2 and (slide.get("layout") or "content") in CORE_LAYOUTS:
            elements = sorted(set(elements) | {"connector", "node"})
        slide["visual_plan"] = {
            "structure_elements": elements,
            "anti_flattening": "Use at least two structure elements; avoid pure rectangles plus bullet lists.",
        }
    return working


def build_visual_redesign_prompt(content: dict, *, page_count: int | None = None) -> str:
    slides = (content or {}).get("slides") or []
    target = f"{page_count} core narrative pages" if page_count else "single-page or multi-page based on content complexity"
    titles = [str(slide.get("title") or slide.get("text") or f"Slide {i}") for i, slide in enumerate(slides, start=1) if isinstance(slide, dict)]
    title_block = "\n".join(f"- {title}" for title in titles[:12]) or "- User-provided material"
    return (
        "Create a Pro PPT visual redesign plan before rendering.\n"
        f"Target: {target}.\n\n"
        "Process:\n"
        "1. Decompose content into information blocks and assign each block a logical role.\n"
        "2. Identify relationships: sequence, cause-effect, contrast, hierarchy, cycle, dashboard/data story.\n"
        "3. Plan pages: user-requested page count means core narrative pages; cover/summary may be added as structure pages.\n"
        "4. Layer each page into core, support, and detail information.\n"
        "5. Choose a visual style that fits content; do not force every page into cards.\n"
        "6. For every core page, specify at least two structure elements: connectors/arrows, nodes, spatial contrast, ring/cycle, progress/status, asymmetry, hierarchy, or data visualization.\n"
        "7. End each page prompt with an anti-flat instruction: no plain rectangle stacks or side-by-side text blocks.\n\n"
        "Current slide/title hints:\n"
        f"{title_block}\n"
    )


__all__ = [
    "STRUCTURE_ELEMENTS",
    "infer_structure_elements",
    "visual_redesign_check",
    "apply_visual_redesign_guidance",
    "build_visual_redesign_prompt",
]
