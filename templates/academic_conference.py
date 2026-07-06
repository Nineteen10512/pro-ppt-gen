"""academic_conference — 学术会议报告。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="academic_conference",
    display_name="学术会议报告",
    scene="academic",
    base_theme="academic",
    theme_overrides={"color": {"primary": "#0F2C4D"}},
    default_transition="fade",
    default_animation=None,
    cover_style="会议 logo + 讲题 + 讲者姓名 + 单位",
    typical_slides=["cover", "section", "content", "chart", "image_text", "summary", "thanks"],
    description="极简学术风格，无动画无多余装饰，适用于国际会议、研讨会口头报告。",
)
