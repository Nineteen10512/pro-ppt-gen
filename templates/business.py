"""business template — migrated from v1.5 base theme."""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="business",
    display_name="通用商务路演",
    scene="business",
    base_theme="business",
    theme_overrides={},
    default_transition="push",
    default_animation=None,
    cover_style="公司 logo + 标语/项目名 + 汇报人 + 日期",
    typical_slides=["cover", "toc", "section", "content", "kpi", "chart", "table", "thanks"],
    description="石板蓝灰商务风，适用于商业汇报、项目路演、行业分析等企业场景。",
)
