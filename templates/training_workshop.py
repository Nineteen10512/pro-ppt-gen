"""training_workshop — 培训讲座。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="training_workshop",
    display_name="培训讲座",
    scene="teaching",
    base_theme="teaching",
    theme_overrides={"color": {"primary": "#F4B400", "accent": "#F4B400"}},
    default_transition="reveal",
    default_animation={"entrance": "float_in", "sequence": "by_paragraph", "trigger": "on_click"},
    cover_style="培训主题 + 讲师 + 日期",
    typical_slides=["cover", "toc", "section", "content", "two_column", "image_text", "summary", "thanks"],
    description="教学活泼黄风格，reveal切换+浮动入场，适用于企业培训、工作坊、公开课。",
)
