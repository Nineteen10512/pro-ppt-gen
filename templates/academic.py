"""academic template — migrated from v1.5 base theme."""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="academic",
    display_name="通用学术答辩",
    scene="academic",
    base_theme="academic",
    theme_overrides={},
    default_transition="fade",
    default_animation=None,
    cover_style="校名/logo + 论文标题 + 答辩人 + 导师 + 日期",
    typical_slides=["cover", "toc", "section", "content", "chart", "references", "thanks"],
    description="深蓝+棕红强调的正式学术风格，适用于课程汇报、开题、答辩等学术场景。",
)
