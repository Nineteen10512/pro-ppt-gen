"""PPT template registry for PRO-PPTX v1.5.3.

Holds TEMPLATE_REGISTRY: dict[str, PPTTemplate] with all 12 scene presets.
Templates are pure data (< 30 lines each in separate files).

@since v1.5.3
"""
from __future__ import annotations

from shared.import_helper import import_shared

(PPTTemplate,) = import_shared("template_registry", attrs=["PPTTemplate"])


def _register() -> dict[str, PPTTemplate]:
    """Lazily import all template modules and build registry dict."""
    from . import (
        academic, business, teaching,
        thesis_defense, startup_pitch, project_report,
        gaokao_review, academic_conference, product_launch,
        year_end_review, training_workshop, minimal_clean,
    )
    mods = [
        academic, business, teaching,
        thesis_defense, startup_pitch, project_report,
        gaokao_review, academic_conference, product_launch,
        year_end_review, training_workshop, minimal_clean,
    ]
    reg = {}
    for m in mods:
        t: PPTTemplate = m.TEMPLATE
        reg[t.name] = t
    return reg


TEMPLATE_REGISTRY: dict[str, PPTTemplate] = _register()


def list_templates() -> list[dict]:
    """Return lightweight list of all templates for user-facing listing."""
    out = []
    for name, t in TEMPLATE_REGISTRY.items():
        out.append({
            "name": t.name,
            "display_name": t.display_name,
            "endpoint": "ppt",
            "scene": t.scene,
            "description": t.description,
            "default_transition": t.default_transition,
            "cover_style": t.cover_style,
        })
    return out


def get_template(name: str) -> PPTTemplate | None:
    return TEMPLATE_REGISTRY.get(name)
