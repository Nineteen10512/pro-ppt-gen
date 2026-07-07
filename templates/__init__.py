"""PPT templates package for PRO-PPTX v1.6.3.

Usage:
    from pro_ppt_gen.templates import list_templates, get_template
"""
from .registry import TEMPLATE_REGISTRY, list_templates, get_template

__all__ = ["TEMPLATE_REGISTRY", "list_templates", "get_template"]
