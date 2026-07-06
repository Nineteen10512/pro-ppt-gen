"""product_launch — 产品发布会。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="product_launch",
    display_name="产品发布会",
    scene="business",
    base_theme="business",
    theme_overrides={"color": {"bg": "#1A1A2E", "primary": "#7B2FBE", "text": "#FFFFFF", "text_secondary": "#CCCCCC"}},
    default_transition="zoom",
    default_animation={"entrance": "zoom_in", "sequence": "all_at_once", "trigger": "on_click"},
    cover_style="深色背景 + 产品大图 + 产品名 + 发布日期",
    typical_slides=["cover", "full_image", "section", "content", "kpi", "image_text", "thanks"],
    description="深色紫蓝渐变发布会风格，zoom切换+放大动画，适用于产品发布会、新品介绍。",
)
