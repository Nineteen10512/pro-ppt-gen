"""minimal_clean — 极简风。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="minimal_clean",
    display_name="极简风",
    scene="general",
    base_theme="academic",
    theme_overrides={"color": {"bg": "#FFFFFF", "text": "#1A1A1A", "primary": "#1A1A1A", "accent": "#666666", "title_bar_bg": "#1A1A1A"}},
    default_transition="none",
    default_animation=None,
    cover_style="大标题 + 几何线框装饰 + 署名",
    typical_slides=["cover", "section", "content", "two_column", "quote", "thanks"],
    description="黑白灰极简风格，无切换无动画，适用于极简设计、作品集、设计提案。",
)
