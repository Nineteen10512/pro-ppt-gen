"""startup_pitch — 创业路演（高频主推）。"""
from shared.template_registry import PPTTemplate

TEMPLATE = PPTTemplate(
    name="startup_pitch",
    display_name="创业路演",
    scene="business",
    base_theme="business",
    theme_overrides={"color": {"primary": "#FF6B35", "accent": "#FF6B35"}},
    default_transition="conveyor",
    default_animation={"entrance": "fly_in_bottom", "sequence": "all_at_once", "trigger": "on_click"},
    cover_style="产品大图 + 项目名 + Slogan + 日期",
    typical_slides=["cover", "content", "kpi", "chart", "image_text", "table", "thanks"],
    description="商务活力橙风格，conveyor切换+底部飞入动画，适用于创业路演、融资BP、产品发布。",
)
