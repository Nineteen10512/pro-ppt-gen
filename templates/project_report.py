"""project_report — 项目汇报（高频主推）。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="project_report",
    display_name="项目汇报",
    scene="business",
    base_theme="business",
    theme_overrides={"color": {"primary": "#2E5AAC"}},
    default_transition="push",
    default_animation=None,
    cover_style="项目名 + 汇报人 + 部门 + 日期",
    typical_slides=["cover", "toc", "section", "content", "kpi", "chart", "table", "summary", "thanks"],
    description="冷静蓝商务项目汇报风格，无多余动画，适用于项目进度汇报、工作总结、周报月报。",
)
