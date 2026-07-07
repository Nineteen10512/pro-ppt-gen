"""gaokao_review — 高考复习课件（高频主推，面向主人的核心场景）。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="gaokao_review",
    display_name="高考复习课件",
    scene="teaching",
    base_theme="teaching",
    theme_overrides={"color": {"accent": "#CC0000", "title_bar_bg": "#1B5E20"}},
    default_transition="wipe_right",
    default_animation={"entrance": "wipe_right", "sequence": "by_paragraph", "trigger": "on_click"},
    cover_style="学科 + 章节名 + 考纲要求 + 授课教师 + 日期",
    typical_slides=["cover", "section", "content", "two_column", "chart", "table", "image_text", "summary"],
    description="教学模板变体，重点红标注+右擦动画，按段落逐项出现，专为高考一轮/二轮复习课件设计。",
)
