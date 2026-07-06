"""year_end_review — 年终总结。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="year_end_review",
    display_name="年终总结",
    scene="business",
    base_theme="business",
    theme_overrides={"color": {"primary": "#B8860B", "secondary": "#8B4513", "bg": "#FFF8F0"}},
    default_transition="fade",
    default_animation={"entrance": "fade_in", "sequence": "all_at_once", "trigger": "on_click"},
    cover_style="年份大字 + 主题词 + 汇报人",
    typical_slides=["cover", "toc", "section", "kpi", "chart", "content", "summary", "thanks"],
    description="暖金棕商务风，适用于年终总结、年度汇报、新年规划。",
)
