from __future__ import annotations


def infer_design_read(units: list[dict], *, theme: str, meta: dict, lang: str, medium: str) -> dict:
    layouts = [unit.get("layout") for unit in units if unit.get("layout")]
    visual_count = sum(1 for unit in units if unit.get("has_visual"))
    density = "high" if sum(len(" ".join(unit.get("texts") or [])) for unit in units) / max(len(units), 1) > 420 else "medium"
    return {
        "summary": f"{medium} uses {theme} theme with {len(set(layouts))} layout types and {visual_count} visualized slides.",
        "design_variance": min(100, 30 + len(set(layouts)) * 12),
        "motion_intensity": 25,
        "visual_density": density,
        "lang": lang,
        "meta_title": (meta or {}).get("title"),
    }


def build_preflight(issues: list[dict], *, base_quality=None, story=None, include_visual_intent: bool = True) -> dict:
    errors = [item for item in issues if item.get("level") == "error"]
    warnings = [item for item in issues if item.get("level") == "warning"]
    return {
        "error_count": len(errors),
        "warning_count": len(warnings),
        "base_quality_passed": None if base_quality is None else bool(base_quality.get("passed")),
        "story_compliant": None if story is None else bool(story.get("compliant", story.get("passed", False))),
        "visual_intent_checked": include_visual_intent,
    }


def finalize_report(
    *,
    version: str,
    design_read: dict,
    issues: list[dict],
    preflight: dict,
    strict: bool,
    hard_fail_codes: set[str],
    base_quality=None,
    story=None,
    **extra,
) -> dict:
    threshold = 85 if strict else 75
    score = 100
    for item in issues:
        score -= 18 if item.get("level") == "error" else 7
    score = max(0, min(100, score))
    hard_failed = any(item.get("code") in hard_fail_codes for item in issues)
    passed = score >= threshold and not hard_failed
    report = {
        "version": version,
        "score": score,
        "passed": passed,
        "threshold": threshold,
        "design_read": design_read,
        "preflight": preflight,
        "issues": issues,
        "base_quality_score": None if base_quality is None else base_quality.get("total_score"),
        "story_compliant": None if story is None else story.get("compliant", story.get("passed")),
    }
    report.update(extra)
    return report
