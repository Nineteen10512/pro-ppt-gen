"""thesis_defense — 学位论文答辩（高频主推）。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="thesis_defense",
    display_name="学位论文答辩",
    scene="academic",
    base_theme="academic",
    theme_overrides={"color": {"title_bar_bg": "#1F3864"}},
    default_transition="fade",
    default_animation={"entrance": "fade_in", "sequence": "all_at_once", "trigger": "on_click"},
    cover_style="校徽 + 论文题目 + 答辩人 + 导师 + 学院 + 日期",
    typical_slides=["cover", "toc", "section", "content", "chart", "two_column", "references", "thanks"],
    description="学术答辩专用，深蓝灰细节配色，严肃正式，适用于本科/硕士/博士学位答辩。",
)
