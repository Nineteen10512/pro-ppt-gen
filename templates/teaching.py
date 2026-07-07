"""teaching template — migrated from v1.5 base theme (default on animations for classroom)."""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="teaching",
    display_name="教学课件",
    scene="teaching",
    base_theme="teaching",
    theme_overrides={},
    default_transition="fade",
    default_animation={"entrance": "fade_in", "sequence": "by_paragraph", "trigger": "on_click"},
    cover_style="课程名 + 章节名 + 授课教师 + 日期",
    typical_slides=["cover", "section", "content", "two_column", "chart", "image_text", "summary", "thanks"],
    description="绿色大字号教学风格，默认开启段落级入场动画，适用于日常课堂教学课件。",
)
