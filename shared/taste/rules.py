from __future__ import annotations

import re


def issue(issues: list[dict], level: str, code: str, message: str, **extra) -> None:
    item = {"level": level, "code": code, "message": message}
    item.update({k: v for k, v in extra.items() if v is not None})
    issues.append(item)


_PLACEHOLDER_RE = re.compile(r"\b(TODO|TBD|placeholder|lorem ipsum|\[image\]|\[chart\])\b|占位|示例文本", re.I)
_AI_COPY_RE = re.compile(r"值得注意的是|总的来说|综上所述|在当今.*背景下|赋能|抓手|闭环|底座")


def check_copy_and_density(units: list[dict], issues: list[dict]) -> None:
    for unit in units:
        texts = unit.get("texts") or []
        joined = " ".join(texts)
        if _PLACEHOLDER_RE.search(joined):
            issue(issues, "error", "placeholder_text", "Slide contains placeholder text.", page=unit.get("index"))
        if _AI_COPY_RE.search(joined):
            issue(issues, "warning", "generic_ai_copy", "Slide copy has generic AI-flavored wording.", page=unit.get("index"))
        if unit.get("bullet_count", 0) > 6:
            issue(issues, "warning", "dense_bullets", "Slide has too many bullets.", page=unit.get("index"))
        if len(joined) > 900:
            issue(issues, "warning", "text_wall", "Slide text is too dense.", page=unit.get("index"))


def check_visual_intent(units: list[dict], issues: list[dict]) -> None:
    for unit in units:
        layout = unit.get("layout")
        if layout in {"content", "two_column", "summary"} and not unit.get("has_visual"):
            issue(
                issues,
                "warning",
                "missing_visual_intent",
                "Text-heavy slide lacks an image, chart, shape, or explicit visual structure.",
                page=unit.get("index"),
            )
        if unit.get("kind") == "chart" and not unit.get("raw", {}).get("bullets"):
            issue(issues, "warning", "missing_chart_insight", "Chart slide should include 1-3 insight bullets.", page=unit.get("index"))


def check_layout_rhythm(units: list[dict], issues: list[dict]) -> None:
    layouts = [unit.get("layout") for unit in units if unit.get("layout")]
    core = [layout for layout in layouts if layout not in {"cover", "toc", "section", "thanks"}]
    if len(set(core)) <= 1 and len(core) >= 4:
        issue(issues, "warning", "low_layout_variety", "Core slides repeat the same layout too often.")
